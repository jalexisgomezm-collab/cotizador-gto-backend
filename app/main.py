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

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from supabase import create_client, Client

from cotizacion_core import generar_cotizacion, Item

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_SERVICE_ROLE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
BUCKET = os.environ.get("SUPABASE_STORAGE_BUCKET", "cotizaciones")

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
    asesor_id: Optional[str] = None       # uuid del usuario autenticado (auth.uid())
    dias_validez: int = Field(default=30, description="Días hasta la fecha de vencimiento")


class CotizacionOut(BaseModel):
    numero: str
    docx_url: str
    pdf_url: str
    subtotal: float
    igv: float
    total: float


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------
@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.post("/api/cotizaciones", response_model=CotizacionOut)
def crear_cotizacion(payload: CotizacionIn):
    sb = get_supabase()

    # 1) número correlativo (función SQL definida en supabase/schema.sql)
    numero_res = sb.rpc("siguiente_numero_cotizacion", {}).execute()
    numero = numero_res.data
    if not numero:
        raise HTTPException(status_code=500, detail="No se pudo obtener el número de cotización.")

    hoy = date.today()
    vencimiento = hoy + timedelta(days=payload.dias_validez)

    # 2) upsert del cliente (por RUC, si lo tiene)
    cliente_dict = payload.cliente.dict()
    if cliente_dict.get("ruc") and cliente_dict["ruc"] != "-":
        cliente_row = sb.table("clientes").upsert(cliente_dict, on_conflict="ruc").execute()
    else:
        cliente_row = sb.table("clientes").insert(cliente_dict).execute()
    cliente_id = cliente_row.data[0]["id"]

    # 3) calcular totales
    subtotal = sum(i.cantidad * i.valor_unitario for i in payload.items)
    igv = round(subtotal * payload.igv_pct / 100, 2) if payload.operacion_gravada else 0.0
    total = subtotal + igv

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
        numero=numero, docx_url=docx_url, pdf_url=pdf_url,
        subtotal=subtotal, igv=igv, total=total,
    )


@app.get("/api/cotizaciones")
def listar_cotizaciones(limit: int = 50):
    sb = get_supabase()
    res = (
        sb.table("cotizaciones")
        .select("*, clientes(razon_social, ruc)")
        .order("created_at", desc=True)
        .limit(limit)
        .execute()
    )
    return res.data
