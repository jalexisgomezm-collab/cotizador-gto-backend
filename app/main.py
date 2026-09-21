#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cotizador GTO — API backend.

Recibe los datos de una cotización desde el frontend, genera el .docx con
cotizacion_core.generar_cotizacion() (el mismo motor ya aprobado), lo
convierte a PDF con LibreOffice, sube ambos archivos a Supabase Storage y
guarda el registro en la base de datos.

Variables de entorno requeridas (ver .env.example):
    SUPABASE_URL
    SUPABASE_SERVICE_ROLE_KEY
    CORS_ORIGINS            (opcional, por defecto "*")
    SUPABASE_STORAGE_BUCKET (opcional, por defecto "cotizaciones")
"""

import os
import subprocess
import tempfile
from datetime import date, timedelta
from typing import List, Optional

import requests
from fastapi import FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from pydantic import BaseModel, Field
from supabase import create_client, Client

from cotizacion_core import generar_cotizacion, generar_cotizacion_en, moneda_letras_en, Item

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_SERVICE_ROLE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
BUCKET = os.environ.get("SUPABASE_STORAGE_BUCKET", "cotizaciones")
SUNAT_API_TOKEN = os.environ.get("SUNAT_API_TOKEN")

_supabase: Optional[Client] = None


def get_supabase() -> Client:
    """Cliente Supabase perezoso: solo falla si de verdad intentas usarlo
    sin las variables de entorno configuradas (así el backend puede
    levantar y responder /api/health aunque falte configurar Supabase)."""
    global _supabase
    if _supabase is None:
        if not SUPABASE_URL or not SUPABASE_SERVICE_ROLE_KEY:
            raise HTTPException(
                status_code=500,
                detail="Faltan las variables de entorno SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY.",
            )
        _supabase = create_client(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)
    return _supabase


app = FastAPI(title="Cotizador GTO API", version="1.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=os.environ.get("CORS_ORIGINS", "*").split(","),
    allow_methods=["*"],
    allow_headers=["*"],
)


# ---------------------------------------------------------------------------
# Esquemas de entrada
# ---------------------------------------------------------------------------
class ItemIn(BaseModel):
    descripcion: str
    cantidad: float = 1
    valor_unitario: float


class ClienteIn(BaseModel):
    ruc: str = "-"
    razon_social: str
    direccion: str = "-"
    contacto: str = "-"
    telefono: str = "-"
    correo: str = "-"


class CotizacionIn(BaseModel):
    cliente: ClienteIn
    items: List[ItemIn]
    referencia: str = ""
    operacion_gravada: bool = True
    moneda_simbolo: str = "US$"
    moneda_letras: str = "DÓLARES AMERICANOS"
    igv_pct: int = 18
    activities: Optional[List[str]] = None
    condiciones: Optional[dict] = None
    asesor_id: Optional[str] = None       # uuid del perfil elegido (o del usuario autenticado)
    asesor_nombre: Optional[str] = None   # si viene escrito a mano, tiene prioridad sobre asesor_id
    asesor_celular: Optional[str] = None
    asesor_correo: Optional[str] = None
    dias_validez: int = Field(default=30, description="Días hasta la fecha de vencimiento")
    fecha_emision: Optional[str] = None  # "YYYY-MM-DD"; solo se usa al editar, para corregirla sin
                                          # afectar el número (crear_cotizacion siempre usa hoy)


class CotizacionOut(BaseModel):
    id: Optional[str] = None
    numero: str
    docx_url: str
    pdf_url: str
    subtotal: float
    igv: float
    total: float


def resolver_asesor(sb: Client, payload: "CotizacionIn") -> Optional[dict]:
    """Datos de asesor que van impresos en el documento.
    Prioridad: si viene un nombre escrito a mano (asesor_nombre), se usa tal
    cual. Si no, se busca el perfil correspondiente a asesor_id (asesor
    elegido de la lista, o el usuario autenticado).

    Además, si el payload trae un asesor_id, se actualiza ese perfil con lo
    que se haya escrito/elegido aquí, para que la próxima vez ya aparezca
    correcto en el desplegable (reemplaza a la antigua página "Mi perfil")."""
    asesor_dict = None

    if payload.asesor_nombre and payload.asesor_nombre.strip():
        asesor_dict = {
            "nombre": payload.asesor_nombre.strip(),
            "celular": (payload.asesor_celular or "").strip() or "-",
            "correo": (payload.asesor_correo or "").strip() or "-",
        }
    elif payload.asesor_id:
        try:
            perfil_res = (
                sb.table("profiles")
                .select("nombre, celular, correo")
                .eq("id", payload.asesor_id)
                .maybe_single()
                .execute()
            )
            if perfil_res.data:
                asesor_dict = {
                    "nombre": perfil_res.data.get("nombre") or "-",
                    "celular": perfil_res.data.get("celular") or "-",
                    "correo": perfil_res.data.get("correo") or "-",
                }
        except Exception:
            asesor_dict = None

    if payload.asesor_id and asesor_dict:
        try:
            sb.table("profiles").update({
                "nombre": asesor_dict["nombre"],
                "celular": None if asesor_dict["celular"] == "-" else asesor_dict["celular"],
                "correo": None if asesor_dict["correo"] == "-" else asesor_dict["correo"],
            }).eq("id", payload.asesor_id).execute()
        except Exception:
            pass  # no bloquear la generación del documento si esto falla

    return asesor_dict


def guardar_contacto_cliente(sb: Client, cliente_id: str, nombre: str, telefono: str, correo: str) -> None:
    """Guarda (o actualiza) un contacto de compras de este cliente. Un mismo
    cliente puede tener varios contactos —uno por área o comprador— así que
    esto NO sobrescribe el contacto principal del cliente: solo agrega/actualiza
    una entrada en su lista de contactos, identificada por el nombre."""
    nombre = (nombre or "").strip()
    if not nombre or nombre == "-":
        return
    try:
        sb.table("cliente_contactos").upsert(
            {
                "cliente_id": cliente_id,
                "nombre": nombre,
                "telefono": (telefono or "").strip() or None,
                "correo": (correo or "").strip() or None,
            },
            on_conflict="cliente_id,nombre",
        ).execute()
    except Exception:
        pass  # no bloquear la generación del documento si esto falla


def usuario_actual(sb: Client, authorization: Optional[str] = None) -> Optional[dict]:
    """Identifica al usuario autenticado a partir del token que manda el
    frontend (header 'Authorization: Bearer <token>') y si es administrador
    (ve todas las cotizaciones) o vendedor normal (solo las suyas).
    Devuelve None si no viene token o no es válido — en ese caso el llamador
    trata la petición como si no supiera quién es (modo abierto, para no
    romper nada si el frontend todavía no manda el token)."""
    if not authorization or not authorization.lower().startswith("bearer "):
        return None
    token = authorization.split(" ", 1)[1].strip()
    if not token:
        return None
    try:
        user_res = sb.auth.get_user(token)
        uid = user_res.user.id
    except Exception:
        return None
    es_admin = False
    try:
        perfil = sb.table("profiles").select("es_admin").eq("id", uid).maybe_single().execute()
        es_admin = bool(perfil.data and perfil.data.get("es_admin"))
    except Exception:
        pass
    return {"uid": uid, "es_admin": es_admin}


def _liberar_numero(sb: Client, numero: str) -> None:
    """Si algo falla después de sacar un número de cotización (antes de
    guardarla), devuelve ese número al contador para que no quede
    "quemado" sin usarse y la numeración no se salte."""
    try:
        sb.rpc("revertir_numero_si_no_se_uso", {"numero_int": int(numero)}).execute()
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------
@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.get("/api/sunat/ruc/{numero}")
def consultar_ruc(numero: str):
    """Consulta un RUC en SUNAT vía Decolecta/apis.net.pe y devuelve los
    datos ya listos para llenar el formulario (razón social, dirección,
    estado). Requiere la variable de entorno SUNAT_API_TOKEN."""
    numero = numero.strip()
    if len(numero) != 11 or not numero.isdigit():
        raise HTTPException(status_code=400, detail="El RUC debe tener 11 dígitos.")
    if not SUNAT_API_TOKEN:
        raise HTTPException(
            status_code=500,
            detail="No está configurada la variable SUNAT_API_TOKEN en el backend.",
        )

    try:
        resp = requests.get(
            "https://api.decolecta.com/v1/sunat/ruc",
            params={"numero": numero, "token": SUNAT_API_TOKEN},
            headers={
                "Accept": "application/json",
                "Content-Type": "application/json",
                "Authorization": f"Bearer {SUNAT_API_TOKEN}",
                "User-Agent": "cotizador-gto/1.0",
            },
            timeout=10,
        )
    except requests.RequestException as exc:
        raise HTTPException(status_code=502, detail=f"No se pudo contactar a SUNAT: {exc}")

    if resp.status_code == 404:
        raise HTTPException(status_code=404, detail="RUC no encontrado en SUNAT.")
    if not resp.ok:
        raise HTTPException(
            status_code=502,
            detail=f"SUNAT/Decolecta respondió con error ({resp.status_code}): {resp.text[:300]}",
        )

    data = resp.json()
    return {
        "ruc": data.get("numero_documento") or numero,
        "razon_social": data.get("razon_social") or data.get("nombre_o_razon_social") or "",
        "direccion": data.get("direccion") or data.get("direccion_completa") or "",
        "estado": data.get("estado") or "",
        "condicion": data.get("condicion") or "",
    }


@app.post("/api/cotizaciones", response_model=CotizacionOut)
def crear_cotizacion(payload: CotizacionIn, authorization: Optional[str] = Header(None)):
    sb = get_supabase()

    # Si el que llama es un vendedor normal (no administrador), la cotización
    # queda siempre a su propio nombre, sin importar lo que haya llegado en
    # el payload — así su plantilla siempre muestra sus propios datos y
    # nadie puede emitir "a nombre de" otro vendedor por error.
    usuario = usuario_actual(sb, authorization)
    if usuario and not usuario["es_admin"]:
        payload.asesor_id = usuario["uid"]
        payload.asesor_nombre = None

    # 1) número correlativo (función SQL definida en supabase/schema.sql)
    numero_res = sb.rpc("siguiente_numero_cotizacion", {}).execute()
    numero = numero_res.data
    if not numero:
        raise HTTPException(status_code=500, detail="No se pudo obtener el número de cotización.")

    try:
        hoy = date.today()
        vencimiento = hoy + timedelta(days=payload.dias_validez)

        # 2) upsert del cliente (por RUC, si lo tiene)
        cliente_dict = payload.cliente.dict()
        if cliente_dict.get("ruc") and cliente_dict["ruc"] != "-":
            cliente_row = sb.table("clientes").upsert(cliente_dict, on_conflict="ruc").execute()
        else:
            cliente_row = sb.table("clientes").insert(cliente_dict).execute()
        cliente_id = cliente_row.data[0]["id"]
        guardar_contacto_cliente(
            sb, cliente_id, cliente_dict.get("contacto", ""),
            cliente_dict.get("telefono", ""), cliente_dict.get("correo", ""),
        )

        # 3) calcular totales
        subtotal = sum(i.cantidad * i.valor_unitario for i in payload.items)
        igv = round(subtotal * payload.igv_pct / 100, 2) if payload.operacion_gravada else 0.0
        total = subtotal + igv

        # 3.5) datos del asesor que va impreso en el documento
        asesor_dict = resolver_asesor(sb, payload)

        # 4) generar el .docx y convertirlo a PDF en una carpeta temporal
        with tempfile.TemporaryDirectory() as tmp:
            docx_path = os.path.join(tmp, f"Cotizacion_{numero}.docx")
            try:
                generar_cotizacion(
                    salida_path=docx_path,
                    numero_cotizacion=numero,
                    fecha_emision=hoy.strftime("%d/%m/%Y"),
                    fecha_vencimiento=vencimiento.strftime("%d/%m/%Y"),
                    referencia=payload.referencia,
                    cliente=cliente_dict,
                    items=[Item(i.descripcion, i.cantidad, i.valor_unitario) for i in payload.items],
                    asesor=asesor_dict,
                    operacion_gravada=payload.operacion_gravada,
                    moneda_simbolo=payload.moneda_simbolo,
                    moneda_letras=payload.moneda_letras,
                    igv_pct=payload.igv_pct,
                    activities=payload.activities,
                    condiciones=payload.condiciones,
                )
            except Exception as exc:
                raise HTTPException(status_code=500, detail=f"Error generando el Word: {exc}")

            try:
                subprocess.run(
                    ["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, docx_path],
                    check=True, timeout=60, capture_output=True,
                )
            except Exception as exc:
                raise HTTPException(status_code=500, detail=f"Error convirtiendo a PDF: {exc}")
            pdf_path = docx_path[:-5] + ".pdf"

            # 5) subir ambos archivos a Supabase Storage
            docx_key = f"{numero}/Cotizacion_{numero}.docx"
            pdf_key = f"{numero}/Cotizacion_{numero}.pdf"
            with open(docx_path, "rb") as f:
                sb.storage.from_(BUCKET).upload(
                    docx_key, f.read(),
                    {"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
                )
            with open(pdf_path, "rb") as f:
                sb.storage.from_(BUCKET).upload(pdf_key, f.read(), {"content-type": "application/pdf"})

        docx_url = sb.storage.from_(BUCKET).get_public_url(docx_key)
        pdf_url = sb.storage.from_(BUCKET).get_public_url(pdf_key)

        # 6) guardar cotización + ítems en la base de datos
        cot_row = sb.table("cotizaciones").insert({
            "numero": numero,
            "cliente_id": cliente_id,
            "asesor_id": payload.asesor_id,
            "referencia": payload.referencia,
            "fecha_emision": hoy.isoformat(),
            "fecha_vencimiento": vencimiento.isoformat(),
            "moneda_simbolo": payload.moneda_simbolo,
            "moneda_letras": payload.moneda_letras,
            "operacion_gravada": payload.operacion_gravada,
            "igv_pct": payload.igv_pct,
            "subtotal": subtotal,
            "igv": igv,
            "total": total,
            "activities": payload.activities,
            "condiciones": payload.condiciones,
            "docx_url": docx_url,
            "pdf_url": pdf_url,
        }).execute()
        cotizacion_id = cot_row.data[0]["id"]

        sb.table("cotizacion_items").insert([
            {
                "cotizacion_id": cotizacion_id,
                "descripcion": i.descripcion,
                "cantidad": i.cantidad,
                "valor_unitario": i.valor_unitario,
            }
            for i in payload.items
        ]).execute()

        return CotizacionOut(
            id=cotizacion_id, numero=numero, docx_url=docx_url, pdf_url=pdf_url,
            subtotal=subtotal, igv=igv, total=total,
        )
    except HTTPException:
        _liberar_numero(sb, numero)
        raise
    except Exception as exc:
        _liberar_numero(sb, numero)
        raise HTTPException(status_code=500, detail=f"Error inesperado generando la cotización: {exc}")


@app.get("/api/cotizaciones")
def listar_cotizaciones(limit: int = 50, authorization: Optional[str] = Header(None)):
    sb = get_supabase()
    q = (
        sb.table("cotizaciones")
        .select("*, clientes(razon_social, ruc)")
        .order("created_at", desc=True)
        .limit(limit)
    )
    # Un vendedor normal (no administrador) solo ve sus propias cotizaciones.
    # Si no llega token (frontend viejo, o falla la verificación), se
    # devuelve todo — igual que antes — para no romper nada.
    usuario = usuario_actual(sb, authorization)
    if usuario and not usuario["es_admin"]:
        q = q.eq("asesor_id", usuario["uid"])
    res = q.execute()
    return res.data


class ClienteUpdateIn(BaseModel):
    razon_social: str
    ruc: Optional[str] = None
    direccion: Optional[str] = None
    contacto: Optional[str] = None
    telefono: Optional[str] = None
    correo: Optional[str] = None


@app.get("/api/clientes")
def listar_clientes():
    sb = get_supabase()
    res = sb.table("clientes").select("*").order("razon_social").execute()
    return res.data


@app.put("/api/clientes/{cliente_id}")
def editar_cliente(cliente_id: str, payload: ClienteUpdateIn):
    """Corrige un cliente del catálogo (p. ej. uno que quedó mal guardado
    por un error de tipeo o un RUC duplicado en el documento original)."""
    sb = get_supabase()
    if not payload.razon_social.strip():
        raise HTTPException(status_code=400, detail="La razón social es obligatoria.")

    def limpio(v: Optional[str]) -> Optional[str]:
        v = (v or "").strip()
        return v or None

    update_dict = {
        "razon_social": payload.razon_social.strip(),
        "ruc": limpio(payload.ruc),
        "direccion": limpio(payload.direccion),
        "contacto": limpio(payload.contacto),
        "telefono": limpio(payload.telefono),
        "correo": limpio(payload.correo),
    }
    try:
        res = sb.table("clientes").update(update_dict).eq("id", cliente_id).execute()
    except Exception as exc:
        msg = str(exc).lower()
        if "duplicate" in msg or "unique" in msg:
            raise HTTPException(status_code=409, detail="Ya existe otro cliente con ese RUC.")
        raise HTTPException(status_code=500, detail=f"No se pudo actualizar el cliente: {exc}")
    if not res.data:
        raise HTTPException(status_code=404, detail="Cliente no encontrado.")
    return res.data[0]


@app.delete("/api/clientes/{cliente_id}")
def eliminar_cliente(cliente_id: str):
    """Elimina un cliente del catálogo. Se bloquea si tiene cotizaciones
    asociadas, para no romper el historial."""
    sb = get_supabase()
    asociadas = (
        sb.table("cotizaciones")
        .select("id", count="exact")
        .eq("cliente_id", cliente_id)
        .limit(1)
        .execute()
    )
    if asociadas.count and asociadas.count > 0:
        raise HTTPException(
            status_code=409,
            detail="No se puede eliminar: este cliente tiene cotizaciones asociadas en el historial.",
        )
    try:
        sb.table("clientes").delete().eq("id", cliente_id).execute()
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"No se pudo eliminar el cliente: {exc}")
    return {"ok": True}


@app.get("/api/clientes/{cliente_id}/contactos")
def listar_contactos_cliente(cliente_id: str):
    """Lista los contactos (compradores) guardados para este cliente, para
    poder elegir entre ellos al cotizar en vez de escribirlos de nuevo."""
    sb = get_supabase()
    res = (
        sb.table("cliente_contactos")
        .select("id, nombre, telefono, correo")
        .eq("cliente_id", cliente_id)
        .order("nombre")
        .execute()
    )
    return res.data


@app.delete("/api/clientes/{cliente_id}/contactos/{contacto_id}")
def eliminar_contacto_cliente(cliente_id: str, contacto_id: str):
    sb = get_supabase()
    try:
        sb.table("cliente_contactos").delete().eq("id", contacto_id).eq(
            "cliente_id", cliente_id
        ).execute()
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"No se pudo eliminar el contacto: {exc}")
    return {"ok": True}


@app.post("/api/cotizaciones/importar", response_model=CotizacionOut)
async def importar_cotizacion(
    numero: str = Form(...),
    ruc: str = Form("-"),
    razon_social: str = Form(...),
    direccion: str = Form("-"),
    contacto: str = Form("-"),
    telefono: str = Form("-"),
    correo: str = Form("-"),
    referencia: str = Form(""),
    fecha_emision: str = Form(...),  # YYYY-MM-DD
    moneda_simbolo: str = Form("US$"),
    moneda_letras: str = Form("DÓLARES AMERICANOS"),
    subtotal: float = Form(0),
    igv: float = Form(0),
    total: float = Form(0),
    docx: UploadFile = File(...),
    pdf: Optional[UploadFile] = File(None),
):
    """Registra una cotización hecha ANTES de este sistema: guarda sus
    archivos Word/PDF originales tal cual (no se regeneran) y crea la fila
    en la base de datos para que aparezca en el historial, con estado
    'importada' (no se puede editar/regenerar, porque no viene de nuestra
    plantilla)."""
    sb = get_supabase()

    numero = numero.strip()
    if not numero:
        raise HTTPException(status_code=400, detail="El número de cotización es obligatorio.")

    existe = sb.table("cotizaciones").select("id").eq("numero", numero).maybe_single().execute()
    if existe.data:
        raise HTTPException(
            status_code=400, detail=f"Ya existe una cotización con el número {numero}."
        )

    cliente_dict = {
        "ruc": ruc.strip() or "-",
        "razon_social": razon_social.strip(),
        "direccion": direccion.strip() or "-",
        "contacto": contacto.strip() or "-",
        "telefono": telefono.strip() or "-",
        "correo": correo.strip() or "-",
    }
    if cliente_dict["ruc"] != "-":
        cliente_row = sb.table("clientes").upsert(cliente_dict, on_conflict="ruc").execute()
    else:
        cliente_row = sb.table("clientes").insert(cliente_dict).execute()
    cliente_id = cliente_row.data[0]["id"]

    try:
        docx_bytes = await docx.read()
        docx_key = f"{numero}/Cotizacion_{numero}.docx"
        sb.storage.from_(BUCKET).upload(
            docx_key, docx_bytes,
            {
                "content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "upsert": "true",
            },
        )
        docx_url = sb.storage.from_(BUCKET).get_public_url(docx_key)

        pdf_url = ""
        if pdf is not None:
            pdf_bytes = await pdf.read()
            pdf_key = f"{numero}/Cotizacion_{numero}.pdf"
            sb.storage.from_(BUCKET).upload(
                pdf_key, pdf_bytes, {"content-type": "application/pdf", "upsert": "true"}
            )
            pdf_url = sb.storage.from_(BUCKET).get_public_url(pdf_key)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Error subiendo los archivos: {exc}")

    cot_row = sb.table("cotizaciones").insert({
        "numero": numero,
        "cliente_id": cliente_id,
        "asesor_id": None,
        "referencia": referencia.strip(),
        "fecha_emision": fecha_emision,
        "fecha_vencimiento": fecha_emision,
        "moneda_simbolo": moneda_simbolo,
        "moneda_letras": moneda_letras,
        "operacion_gravada": igv > 0,
        "igv_pct": 18,
        "subtotal": subtotal,
        "igv": igv,
        "total": total,
        "docx_url": docx_url,
        "pdf_url": pdf_url or None,
        "estado": "importada",
    }).execute()
    cotizacion_id = cot_row.data[0]["id"]

    return CotizacionOut(
        id=cotizacion_id, numero=numero, docx_url=docx_url, pdf_url=pdf_url,
        subtotal=subtotal, igv=igv, total=total,
    )


@app.put("/api/cotizaciones/{cotizacion_id}", response_model=CotizacionOut)
def editar_cotizacion(
    cotizacion_id: str, payload: CotizacionIn, authorization: Optional[str] = Header(None)
):
    """Corrige una cotización ya existente y regenera el Word/PDF.
    Mantiene el mismo número correlativo y la misma fecha de emisión;
    todo lo demás (cliente, ítems, condiciones, moneda, etc.) se
    sobreescribe con lo que llegue en el payload."""
    sb = get_supabase()

    existente = (
        sb.table("cotizaciones")
        .select("numero, fecha_emision, asesor_id")
        .eq("id", cotizacion_id)
        .maybe_single()
        .execute()
    )
    if not existente.data:
        raise HTTPException(status_code=404, detail="Cotización no encontrada.")

    # Un vendedor normal (no administrador) no puede editar cotizaciones
    # ajenas, y las suyas quedan siempre a su propio nombre.
    usuario = usuario_actual(sb, authorization)
    if usuario and not usuario["es_admin"]:
        if existente.data.get("asesor_id") != usuario["uid"]:
            raise HTTPException(
                status_code=403, detail="No puedes editar una cotización de otro vendedor."
            )
        payload.asesor_id = usuario["uid"]
        payload.asesor_nombre = None

    numero = existente.data["numero"]
    if payload.fecha_emision:
        try:
            fecha_emision = date.fromisoformat(payload.fecha_emision)
        except ValueError:
            raise HTTPException(status_code=400, detail="Fecha de emisión inválida.")
    else:
        fecha_emision = date.fromisoformat(existente.data["fecha_emision"])
    vencimiento = fecha_emision + timedelta(days=payload.dias_validez)

    # 1) upsert del cliente (por RUC, si lo tiene) — igual que en creación
    cliente_dict = payload.cliente.dict()
    if cliente_dict.get("ruc") and cliente_dict["ruc"] != "-":
        cliente_row = sb.table("clientes").upsert(cliente_dict, on_conflict="ruc").execute()
    else:
        cliente_row = sb.table("clientes").insert(cliente_dict).execute()
    cliente_id = cliente_row.data[0]["id"]
    guardar_contacto_cliente(
        sb, cliente_id, cliente_dict.get("contacto", ""),
        cliente_dict.get("telefono", ""), cliente_dict.get("correo", ""),
    )

    # 2) recalcular totales
    subtotal = sum(i.cantidad * i.valor_unitario for i in payload.items)
    igv = round(subtotal * payload.igv_pct / 100, 2) if payload.operacion_gravada else 0.0
    total = subtotal + igv

    # 3) datos del asesor que va impreso en el documento
    asesor_dict = resolver_asesor(sb, payload)

    # 4) regenerar el .docx y el PDF con el mismo número
    with tempfile.TemporaryDirectory() as tmp:
        docx_path = os.path.join(tmp, f"Cotizacion_{numero}.docx")
        try:
            generar_cotizacion(
                salida_path=docx_path,
                numero_cotizacion=numero,
                fecha_emision=fecha_emision.strftime("%d/%m/%Y"),
                fecha_vencimiento=vencimiento.strftime("%d/%m/%Y"),
                referencia=payload.referencia,
                cliente=cliente_dict,
                items=[Item(i.descripcion, i.cantidad, i.valor_unitario) for i in payload.items],
                asesor=asesor_dict,
                operacion_gravada=payload.operacion_gravada,
                moneda_simbolo=payload.moneda_simbolo,
                moneda_letras=payload.moneda_letras,
                igv_pct=payload.igv_pct,
                activities=payload.activities,
                condiciones=payload.condiciones,
            )
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"Error generando el Word: {exc}")

        try:
            subprocess.run(
                ["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, docx_path],
                check=True, timeout=60, capture_output=True,
            )
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"Error convirtiendo a PDF: {exc}")
        pdf_path = docx_path[:-5] + ".pdf"

        # 5) volver a subir ambos archivos a Supabase Storage, sobreescribiendo los anteriores
        docx_key = f"{numero}/Cotizacion_{numero}.docx"
        pdf_key = f"{numero}/Cotizacion_{numero}.pdf"
        with open(docx_path, "rb") as f:
            sb.storage.from_(BUCKET).upload(
                docx_key, f.read(),
                {
                    "content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "upsert": "true",
                },
            )
        with open(pdf_path, "rb") as f:
            sb.storage.from_(BUCKET).upload(
                pdf_key, f.read(),
                {"content-type": "application/pdf", "upsert": "true"},
            )

    docx_url = sb.storage.from_(BUCKET).get_public_url(docx_key)
    pdf_url = sb.storage.from_(BUCKET).get_public_url(pdf_key)

    # 6) actualizar la fila de la cotización
    sb.table("cotizaciones").update({
        "cliente_id": cliente_id,
        "asesor_id": payload.asesor_id,
        "referencia": payload.referencia,
        "fecha_emision": fecha_emision.isoformat(),
        "fecha_vencimiento": vencimiento.isoformat(),
        "moneda_simbolo": payload.moneda_simbolo,
        "moneda_letras": payload.moneda_letras,
        "operacion_gravada": payload.operacion_gravada,
        "igv_pct": payload.igv_pct,
        "subtotal": subtotal,
        "igv": igv,
        "total": total,
        "activities": payload.activities,
        "condiciones": payload.condiciones,
        "docx_url": docx_url,
        "pdf_url": pdf_url,
        "estado": "generada",
    }).eq("id", cotizacion_id).execute()

    # 7) reemplazar los ítems (borrar los anteriores e insertar los nuevos)
    sb.table("cotizacion_items").delete().eq("cotizacion_id", cotizacion_id).execute()
    sb.table("cotizacion_items").insert([
        {
            "cotizacion_id": cotizacion_id,
            "descripcion": i.descripcion,
            "cantidad": i.cantidad,
            "valor_unitario": i.valor_unitario,
        }
        for i in payload.items
    ]).execute()

    return CotizacionOut(
        id=cotizacion_id, numero=numero, docx_url=docx_url, pdf_url=pdf_url,
        subtotal=subtotal, igv=igv, total=total,
    )


@app.get("/api/cotizaciones/{cotizacion_id}/ingles")
def descargar_cotizacion_ingles(
    cotizacion_id: str, formato: str = "pdf", authorization: Optional[str] = Header(None)
):
    """Genera al vuelo la versión en inglés de una cotización ya existente:
    arma el documento con la plantilla en inglés (títulos, secciones y
    textos fijos traducidos, mismo diseño/marca). Los campos de texto libre
    (descripción de ítems, alcance del servicio, condiciones personalizadas,
    referencia) se dejan tal como fueron escritos — no se traducen
    automáticamente, para no depender de un servicio de traducción de pago.
    No guarda nada ni modifica la cotización original en español."""
    if formato not in ("pdf", "docx"):
        raise HTTPException(status_code=400, detail="Formato inválido (usa pdf o docx).")

    sb = get_supabase()
    cot = (
        sb.table("cotizaciones")
        .select("*, clientes(ruc, razon_social, direccion, contacto, telefono, correo)")
        .eq("id", cotizacion_id)
        .maybe_single()
        .execute()
    )
    if not cot.data:
        raise HTTPException(status_code=404, detail="Cotización no encontrada.")
    row = cot.data

    usuario = usuario_actual(sb, authorization)
    if usuario and not usuario["es_admin"] and row.get("asesor_id") != usuario["uid"]:
        raise HTTPException(
            status_code=403, detail="No puedes descargar una cotización de otro vendedor."
        )

    if row.get("estado") == "importada":
        raise HTTPException(
            status_code=400,
            detail="Esta cotización fue importada (archivo original, sin datos estructurados) "
                   "y no se puede traducir automáticamente.",
        )

    items_res = (
        sb.table("cotizacion_items")
        .select("descripcion, cantidad, valor_unitario")
        .eq("cotizacion_id", cotizacion_id)
        .order("id")
        .execute()
    )
    items_data = items_res.data or []
    if not items_data:
        raise HTTPException(status_code=400, detail="Esta cotización no tiene ítems.")

    asesor_dict = None
    if row.get("asesor_id"):
        try:
            perfil_res = (
                sb.table("profiles")
                .select("nombre, celular, correo")
                .eq("id", row["asesor_id"])
                .maybe_single()
                .execute()
            )
            if perfil_res.data:
                asesor_dict = {
                    "nombre": perfil_res.data.get("nombre") or "-",
                    "celular": perfil_res.data.get("celular") or "-",
                    "correo": perfil_res.data.get("correo") or "-",
                }
        except Exception:
            asesor_dict = None

    cliente_dict = row.get("clientes") or {}
    activities = row.get("activities") or None
    condiciones_es = row.get("condiciones") or None
    referencia = row.get("referencia") or ""

    items_en = [
        Item(it.get("descripcion", ""), it.get("cantidad", 1), it.get("valor_unitario", 0))
        for it in items_data
    ]

    try:
        fecha_emision_dt = date.fromisoformat(row["fecha_emision"])
        fecha_vencimiento_dt = date.fromisoformat(row["fecha_vencimiento"])
    except Exception:
        raise HTTPException(status_code=500, detail="Fechas inválidas en la cotización.")

    numero = row["numero"]
    with tempfile.TemporaryDirectory() as tmp:
        docx_path = os.path.join(tmp, f"Quotation_{numero}_EN.docx")
        try:
            generar_cotizacion_en(
                salida_path=docx_path,
                numero_cotizacion=numero,
                fecha_emision=fecha_emision_dt.strftime("%m/%d/%Y"),
                fecha_vencimiento=fecha_vencimiento_dt.strftime("%m/%d/%Y"),
                referencia=referencia,
                cliente=cliente_dict,
                items=items_en,
                asesor=asesor_dict,
                operacion_gravada=row.get("operacion_gravada", True),
                moneda_simbolo=row.get("moneda_simbolo", "US$"),
                moneda_letras=moneda_letras_en(row.get("moneda_letras", "DÓLARES AMERICANOS")),
                igv_pct=row.get("igv_pct", 18),
                activities=activities,
                condiciones=condiciones_es,
            )
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"Error generando el Word en inglés: {exc}")

        if formato == "docx":
            salida_path = docx_path
            media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            filename = f"Quotation_{numero}_EN.docx"
        else:
            try:
                subprocess.run(
                    ["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, docx_path],
                    check=True, timeout=60, capture_output=True,
                )
            except Exception as exc:
                raise HTTPException(status_code=500, detail=f"Error convirtiendo a PDF: {exc}")
            salida_path = docx_path[:-5] + ".pdf"
            media_type = "application/pdf"
            filename = f"Quotation_{numero}_EN.pdf"

        with open(salida_path, "rb") as f:
            contenido = f.read()

    return Response(
        content=contenido,
        media_type=media_type,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
