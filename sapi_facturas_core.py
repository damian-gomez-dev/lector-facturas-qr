"""
Módulo Core de Carga Masiva de Facturas Digitales (SAPI Web -> Lector Facturas QR).

Incluye:
- Mapeo oficial de tipos de comprobante AFIP/ARCA a SSSalud (comprobantes_sssalud)
- Codificación oficial del Nomenclador de Discapacidad SSSalud 1..99 y Sin Recupero (codificacion_facturas)
- Servicio multicapa de extracción de QR AFIP/ARCA y heurísticas de texto PDF (QrService)
- Validación opcional de CUIT Institucional Receptor (validar_cuit_receptor_institucion)
- Generación de ZIP de facturas adaptadas a 1 página en formato SSSalud (3 carpetas)
- Validación interactiva de filas para Revisión y Autorización (sin auditoría vs autorizaciones)
- Exportación a CSV, Excel (.xlsx) y TXT formato SSSalud
"""

from __future__ import annotations

import base64
import csv
import io
import json
import os
import re
import unicodedata
import urllib.parse
import zipfile
from datetime import date, datetime
from functools import lru_cache
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import fitz  # PyMuPDF
import numpy as np
import pandas as pd

try:
    import cv2
    HAS_CV2 = True
except ImportError:
    cv2 = None  # type: ignore
    HAS_CV2 = False

try:
    import pytesseract
    HAS_TESSERACT = True
except ImportError:
    HAS_TESSERACT = False

try:
    from pyzbar.pyzbar import decode as pyzbar_decode
    HAS_PYZBAR = True
except ImportError:
    HAS_PYZBAR = False

# Límite de versión Demo de Carga Masiva SAPI Web
MAX_PDFS_POR_TANDA_DEMO = 15
MAX_PDF_BYTES = 15 * 1024 * 1024  # 15 MB por archivo PDF


# ==============================================================================
# 1. COMPROBANTES SSSALUD Y AFIP/ARCA (de backend/app/core/comprobantes_sssalud.py)
# ==============================================================================

MAPA_TIPO_COMPROBANTE_AFIP: Dict[int, str] = {
    # Facturas
    1: "FA_A",
    6: "FA_B",
    11: "FA_C",
    51: "FA_M",
    201: "FA_A",  # Factura de Crédito Electrónica MiPyMEs A
    206: "FA_B",  # Factura de Crédito Electrónica MiPyMEs B
    211: "FA_C",  # Factura de Crédito Electrónica MiPyMEs C
    # Recibos
    4: "REC_A",
    9: "REC_B",
    15: "REC_C",
    54: "REC_M",
    # Notas de Crédito (no permitidas en carga de facturas)
    3: "NC_A",
    8: "NC_B",
    13: "NC_C",
    53: "NC_M",
    203: "NC_A",
    208: "NC_B",
    213: "NC_C",
    # Notas de Débito
    2: "ND_A",
    7: "ND_B",
    12: "ND_C",
    52: "ND_M",
    202: "ND_A",
    207: "ND_B",
    212: "ND_C",
}

CODIGOS_AFIP_NOTA_CREDITO = {3, 8, 13, 53, 203, 208, 213}
CODIGOS_AFIP_RECIBO = {4, 9, 15, 54}
CODIGOS_AFIP_FACTURA = {1, 6, 11, 51, 201, 206, 211}

# Equivalencia directa Código AFIP -> Código SSSalud (1 al 8)
MAPA_AFIP_A_SSSALUD: Dict[int, int] = {
    1: 1,    # Factura A -> 1
    4: 2,    # Recibo A -> 2
    6: 3,    # Factura B -> 3
    9: 4,    # Recibo B -> 4
    11: 5,   # Factura C -> 5
    15: 6,   # Recibo C -> 6
    51: 7,   # Factura M -> 7
    54: 8,   # Recibo M -> 8
    201: 1,  # FCE A -> 1
    206: 3,  # FCE B -> 3
    211: 5,  # FCE C -> 5
}

# Equivalencia Código Interno SAPI -> Código SSSalud (1 al 8)
MAPA_TIPO_INTERNO_A_SSSALUD: Dict[str, int] = {
    "FA_A": 1,
    "FACTURA A": 1,
    "A": 1,
    "1": 1,
    "REC_A": 2,
    "RECIBO A": 2,
    "RA": 2,
    "2": 2,
    "FA_B": 3,
    "FACTURA B": 3,
    "B": 3,
    "3": 3,
    "REC_B": 4,
    "RECIBO B": 4,
    "RB": 4,
    "4": 4,
    "FA_C": 5,
    "FACTURA C": 5,
    "C": 5,
    "5": 5,
    "REC_C": 6,
    "RECIBO C": 6,
    "RC": 6,
    "6": 6,
    "FA_M": 7,
    "FACTURA M": 7,
    "M": 7,
    "7": 7,
    "REC_M": 8,
    "RECIBO M": 8,
    "RM": 8,
    "8": 8,
}

MAPA_SSSALUD_A_ETIQUETA: Dict[int, str] = {
    1: "Factura A",
    2: "Recibo A",
    3: "Factura B",
    4: "Recibo B",
    5: "Factura C",
    6: "Recibo C",
    7: "Factura M",
    8: "Recibo M",
}


def normalizar_tipo_comprobante_afip(tipo_cmp_afip: Any) -> str:
    try:
        codigo = int(tipo_cmp_afip)
    except (TypeError, ValueError):
        return "FA_C"
    return MAPA_TIPO_COMPROBANTE_AFIP.get(codigo, "FA_C")


def es_nota_credito_comprobante(tipo_comprobante: Any, tipo_cmp_afip: Any = None) -> bool:
    if tipo_cmp_afip is not None:
        try:
            if int(tipo_cmp_afip) in CODIGOS_AFIP_NOTA_CREDITO:
                return True
        except (TypeError, ValueError):
            pass

    texto = str(tipo_comprobante or "").strip().upper()
    if not texto:
        return False

    return (
        texto.startswith("NC_")
        or texto.startswith("NC ")
        or texto in {"NC", "NCA", "NCB", "NCC", "NCM"}
        or "NOTA DE CREDITO" in texto
        or "NOTA DE CRÉDITO" in texto
    )


def es_recibo_comprobante(tipo_comprobante: Any, tipo_cmp_afip: Any = None) -> bool:
    if tipo_cmp_afip is not None:
        try:
            if int(tipo_cmp_afip) in CODIGOS_AFIP_RECIBO:
                return True
        except (TypeError, ValueError):
            pass

    texto = str(tipo_comprobante or "").strip().upper()
    if not texto:
        return False

    return (
        texto.startswith("REC_")
        or texto.startswith("REC ")
        or texto in {"REC", "RA", "RB", "RC", "RM", "2", "4", "6", "8"}
        or "RECIBO" in texto
    )


def obtener_codigo_sssalud_desde_tipo_comprobante(
    tipo_comprobante: Any,
    tipo_cmp_afip: Any = None,
) -> int:
    if tipo_cmp_afip is not None:
        try:
            codigo_afip = int(tipo_cmp_afip)
            if codigo_afip in MAPA_AFIP_A_SSSALUD:
                return MAPA_AFIP_A_SSSALUD[codigo_afip]
        except (TypeError, ValueError):
            pass

    texto = str(tipo_comprobante or "").strip().upper()
    if not texto:
        return 5

    if texto in MAPA_TIPO_INTERNO_A_SSSALUD:
        return MAPA_TIPO_INTERNO_A_SSSALUD[texto]

    es_recibo = "REC" in texto
    letra = "C"
    if texto.endswith("_A") or texto.endswith(" A") or texto == "A":
        letra = "A"
    elif texto.endswith("_B") or texto.endswith(" B") or texto == "B":
        letra = "B"
    elif texto.endswith("_M") or texto.endswith(" M") or texto == "M":
        letra = "M"
    elif texto.endswith("_C") or texto.endswith(" C") or texto == "C":
        letra = "C"

    if es_recibo:
        return {"A": 2, "B": 4, "C": 6, "M": 8}[letra]
    return {"A": 1, "B": 3, "C": 5, "M": 7}[letra]


def obtener_etiqueta_tipo_comprobante(
    tipo_comprobante: Any,
    tipo_cmp_afip: Any = None,
) -> str:
    if es_nota_credito_comprobante(tipo_comprobante, tipo_cmp_afip):
        texto = str(tipo_comprobante or "").strip().upper()
        letra = texto.split("_")[-1] if "_" in texto else "C"
        return f"Nota de Crédito {letra}"

    codigo_sssalud = obtener_codigo_sssalud_desde_tipo_comprobante(
        tipo_comprobante=tipo_comprobante,
        tipo_cmp_afip=tipo_cmp_afip,
    )
    return MAPA_SSSALUD_A_ETIQUETA.get(codigo_sssalud, "Factura C")


def generar_identificador_factura(
    cuit_prestador: Any,
    tipo_comprobante: Any,
    punto_venta: Any,
    numero_comprobante: Any,
    tipo_cmp_afip: Any = None,
) -> str:
    """
    Genera el identificador estructurado oficial de factura SAPI / SSSalud:
    '{CUIT}_{TIPO_SSSALUD}_{PUNTO_VENTA}_{NUMERO_COMPROBANTE}'
    Ejemplo: '20105808039_5_4_733'
    """
    cuit_clean = re.sub(r"\D", "", str(cuit_prestador or "")) or "00000000000"
    tipo_sssalud = obtener_codigo_sssalud_desde_tipo_comprobante(
        tipo_comprobante=tipo_comprobante,
        tipo_cmp_afip=tipo_cmp_afip,
    )

    pto_digits = re.sub(r"\D", "", str(punto_venta or ""))
    nro_digits = re.sub(r"\D", "", str(numero_comprobante or ""))

    pto_int = int(pto_digits) if pto_digits else 0
    nro_int = int(nro_digits) if nro_digits else 0

    return f"{cuit_clean}_{tipo_sssalud}_{pto_int}_{nro_int}"


def formatear_comprobante_visible(punto_venta: Any, numero_comprobante: Any) -> str:
    pto_digits = re.sub(r"\D", "", str(punto_venta or ""))
    nro_digits = re.sub(r"\D", "", str(numero_comprobante or ""))
    pto_int = int(pto_digits) if pto_digits else 0
    nro_int = int(nro_digits) if nro_digits else 0
    return f"{pto_int:04d}-{nro_int:08d}"


def formatear_cuit_guiones(cuit: Any) -> str:
    digits = re.sub(r"\D", "", str(cuit or ""))
    if len(digits) == 11:
        return f"{digits[:2]}-{digits[2:10]}-{digits[10]}"
    return digits or "S/D"


# ==============================================================================
# 2. CODIFICACIÓN DE PRESTACIONES Y NOMENCLADOR SSSALUD
#    (de backend/app/services/codificacion_facturas_service.py)
# ==============================================================================

CODIGOS_CON_CANTIDAD_CUATRO = {"90", "91"}
CODIGOS_CON_CANTIDAD_MANUAL = {"96"}

CODIGOS_DEPENDENCIA_VALIDOS = {
    *(str(i) for i in range(1, 13)),
    *(str(i) for i in range(37, 49)),
    *(str(i) for i in range(58, 73)),
    *(str(i) for i in range(76, 79)),
    "96",
}

PRESTACIONES_SIN_RECUPERO_KEYWORDS = (
    "sin recupero",
    "hidroterapia",
    "equinoterapia",
    "musicoterapia",
    "acompanante terapeutico",
    "acompanamiento terapeutico",
    "cuidador domiciliario",
    "asistente domiciliario",
    "asistencia domiciliaria",
    "internacion domiciliaria",
    "enfermeria domiciliaria",
    "psicopedagogia sin recupero",
    "kinesiologia domiciliaria",
    "terapia recreativa",
    "taller protegido",
    "deporte adaptado",
    "colonia de vacaciones",
    "ortesis",
    "protesis",
    "audifonos",
    "silla de ruedas",
    "insumos descartables",
    "panales",
    "medicacion",
    "cobertura especial sin recupero",
)

REGLAS_CODIFICACION_POR_NOMBRE: List[Tuple[str, str]] = [
    ("transporte", "96"),
    ("traslado", "96"),
    ("alimentacion", "95"),
    ("modulo de apoyo a la integracion escolar", "89"),
    ("apoyo a la integracion escolar", "89"),
    ("saie", "89"),
    ("maestro de apoyo", "87"),
    ("modulo maestro de apoyo", "87"),
    ("estimulacion temprana", "85"),
    ("prestacion de apoyo", "86"),
    ("terapias de apoyo", "86"),
    ("fonoaudiolog", "86"),
    ("psicolog", "86"),
    ("psicopedagog", "86"),
    ("kinesiolog", "86"),
    ("terapia ocupacional", "86"),
    ("psicomotric", "86"),
    ("modulo integral intensivo", "90"),
    ("rehabilitacion integral intensivo", "90"),
    ("modulo integral simple", "91"),
    ("rehabilitacion integral simple", "91"),
    ("hosp de dia jornada simple", "92"),
    ("hospital de dia jornada simple", "92"),
    ("hosp de dia jornada doble", "93"),
    ("hospital de dia jornada doble", "93"),
    ("rehabilitacion internacion", "94"),
    ("escuela comun", "97"),
    ("escuela especial pre primaria", "98"),
    ("escuela especial primaria", "99"),
    ("centro de dia jornada doble cat a", "1"),
    ("centro de dia jornada doble cat b", "2"),
    ("centro de dia jornada doble cat c", "3"),
    ("centro de dia jornada simple cat a", "4"),
    ("centro de dia jornada simple cat b", "5"),
    ("centro de dia jornada simple cat c", "6"),
    ("centro de dia", "1"),
    ("centro educativo terapeutico jornada doble cat a", "7"),
    ("centro educativo terapeutico jornada doble cat b", "8"),
    ("centro educativo terapeutico jornada doble cat c", "9"),
    ("centro educativo terapeutico jornada simple cat a", "10"),
    ("centro educativo terapeutico jornada simple cat b", "11"),
    ("centro educativo terapeutico jornada simple cat c", "12"),
    ("centro educativo terapeutico", "7"),
    ("cet jornada simple", "10"),
    ("cet jornada doble", "7"),
    ("formacion laboral jornada doble cat a", "13"),
    ("formacion laboral jornada doble cat b", "14"),
    ("formacion laboral jornada doble cat c", "15"),
    ("formacion laboral jornada simple cat a", "16"),
    ("formacion laboral jornada simple cat b", "17"),
    ("formacion laboral jornada simple cat c", "18"),
    ("formacion laboral", "13"),
    ("aprestamiento laboral jornada doble cat a", "19"),
    ("aprestamiento laboral jornada doble cat b", "20"),
    ("aprestamiento laboral jornada doble cat c", "21"),
    ("aprestamiento laboral jornada simple cat a", "22"),
    ("aprestamiento laboral jornada simple cat b", "23"),
    ("aprestamiento laboral jornada simple cat c", "24"),
    ("aprestamiento laboral", "19"),
    ("escolaridad pre primaria jornada doble cat a", "25"),
    ("escolaridad pre primaria jornada doble cat b", "26"),
    ("escolaridad pre primaria jornada doble cat c", "27"),
    ("escolaridad pre primaria jornada simple cat a", "28"),
    ("escolaridad pre primaria jornada simple cat b", "29"),
    ("escolaridad pre primaria jornada simple cat c", "30"),
    ("escolaridad pre primaria", "28"),
    ("escolaridad primaria jornada doble cat a", "31"),
    ("escolaridad primaria jornada doble cat b", "32"),
    ("escolaridad primaria jornada doble cat c", "33"),
    ("escolaridad primaria jornada simple cat a", "34"),
    ("escolaridad primaria jornada simple cat b", "35"),
    ("escolaridad primaria jornada simple cat c", "36"),
    ("escolaridad primaria", "34"),
    ("hogar permanente cat a", "37"),
    ("hogar permanente cat b", "38"),
    ("hogar permanente cat c", "39"),
    ("hogar lunes a viernes cat a", "40"),
    ("hogar lunes a viernes cat b", "41"),
    ("hogar lunes a viernes cat c", "42"),
    ("pequeno hogar permanente", "76"),
    ("pequeno hogar", "73"),
    ("residencia permanente", "82"),
    ("residencia lunes a viernes", "79"),
    ("hogar", "37"),
]


def normalizar_texto(texto: Optional[str]) -> str:
    if not texto:
        return ""
    normalizado = unicodedata.normalize("NFD", str(texto).strip().lower())
    sin_acentos = "".join(ch for ch in normalizado if unicodedata.category(ch) != "Mn")
    limpio = re.sub(r"[^a-z0-9\s]", " ", sin_acentos)
    return re.sub(r"\s+", " ", limpio).strip()


@lru_cache(maxsize=1)
def _cargar_nomenclador_raw() -> List[Dict[str, Any]]:
    ruta_json = Path(__file__).resolve().parent / "nomenclador_discapacidad.json"
    if not ruta_json.exists():
        return []
    with ruta_json.open("r", encoding="utf-8") as f:
        return json.load(f)


@lru_cache(maxsize=1)
def _cargar_nomenclador() -> Dict[str, Dict[str, Any]]:
    datos = _cargar_nomenclador_raw()
    catalogo: Dict[str, Dict[str, Any]] = {}
    for item in datos:
        codigo = str(item.get("codigo", "")).strip()
        if codigo:
            catalogo[codigo] = item
    return catalogo


def obtener_catalogo_nomenclador() -> List[Dict[str, Any]]:
    """Retorna la lista completa de prestaciones del nomenclador oficial SSSalud (códigos 1..99) + opción Sin Recupero."""
    items = list(_cargar_nomenclador_raw())
    return items


def obtener_opciones_selector_prestacion() -> List[Dict[str, Any]]:
    """
    Retorna opciones listas para el selector interactivo de Revisión y Autorización:
    Cada elemento tiene: 'key', 'label', 'codigo', 'descripcion', 'sin_recupero', 'cantidad_default'
    """
    opciones: List[Dict[str, Any]] = []
    for item in obtener_catalogo_nomenclador():
        cod = str(item.get("codigo", "")).strip()
        desc = str(item.get("descripcion", "")).strip()
        cant_raw = str(item.get("cantidad", "1")).strip()
        cant_def = 4 if cod in CODIGOS_CON_CANTIDAD_CUATRO else (1 if cant_raw == "-" else int(cant_raw or 1))
        opciones.append({
            "key": f"COD_{cod}",
            "label": f"[{cod}] {desc}",
            "codigo": cod,
            "descripcion": desc,
            "sin_recupero": False,
            "cantidad_default": cant_def,
        })

    # Prestaciones frecuentes de apoyo (Código 86 con nombre específico)
    terapias_apoyo_especificas = [
        ("86", "Prestación de Apoyo - Psicología"),
        ("86", "Prestación de Apoyo - Fonoaudiología"),
        ("86", "Prestación de Apoyo - Psicopedagogía"),
        ("86", "Prestación de Apoyo - Kinesiología"),
        ("86", "Prestación de Apoyo - Terapia Ocupacional"),
        ("86", "Prestación de Apoyo - Psicomotricidad"),
    ]
    for cod, desc in terapias_apoyo_especificas:
        opciones.append({
            "key": f"APOYO_{normalizar_texto(desc).replace(' ', '_')}",
            "label": f"[{cod}] {desc}",
            "codigo": cod,
            "descripcion": desc,
            "sin_recupero": False,
            "cantidad_default": 1,
        })

    # Prestaciones Sin Recupero frecuentes
    sin_recupero_items = [
        "Hidroterapia (Sin Recupero)",
        "Equinoterapia (Sin Recupero)",
        "Musicoterapia (Sin Recupero)",
        "Acompañante Terapéutico (Sin Recupero)",
        "Cuidador / Asistente Domiciliario (Sin Recupero)",
        "Internación / Enfermería Domiciliaria (Sin Recupero)",
        "Ortesis / Prótesis / Ayudas Técnicas (Sin Recupero)",
        "Insumos / Pañales / Medicación (Sin Recupero)",
        "Cobertura Especial Sin Recupero",
    ]
    for desc in sin_recupero_items:
        opciones.append({
            "key": f"SR_{normalizar_texto(desc).replace(' ', '_')}",
            "label": f"[S/R] {desc}",
            "codigo": "",
            "descripcion": desc,
            "sin_recupero": True,
            "cantidad_default": 1,
        })

    return opciones


def limpiar_codigo_nomenclador(codigo: Optional[str]) -> Optional[str]:
    if codigo is None:
        return None
    limpio = str(codigo).strip()
    if not limpio or limpio in {"-", "0", "00", "S/C", "N/A", "NONE", "NULL"}:
        return None
    if limpio.isdigit():
        return str(int(limpio))
    match = re.search(r"\b(\d{1,2})\b", limpio)
    if match:
        return str(int(match.group(1)))
    return None


def es_prestacion_sin_recupero(
    prestacion_nombre: Optional[str],
    sin_recupero_explicito: Optional[bool] = None,
) -> bool:
    if sin_recupero_explicito is True:
        return True
    texto = normalizar_texto(prestacion_nombre)
    if not texto:
        return False
    return any(keyword in texto for keyword in PRESTACIONES_SIN_RECUPERO_KEYWORDS)


def inferir_codigo_por_nombre_prestacion(prestacion_nombre: Optional[str]) -> Optional[str]:
    texto = normalizar_texto(prestacion_nombre)
    if not texto:
        return None

    catalogo = _cargar_nomenclador()
    for codigo, item in catalogo.items():
        if normalizar_texto(item.get("descripcion")) == texto:
            return codigo

    for patron, codigo in REGLAS_CODIFICACION_POR_NOMBRE:
        if patron in texto:
            return codigo

    return None


def resolver_codificacion_prestacion(
    codigo_nomenclador: Optional[str] = None,
    prestacion_nombre: Optional[str] = None,
    cantidad_solicitada: Optional[int] = None,
    sin_recupero_solicitado: Optional[bool] = None,
    dependencia_solicitada: bool = False,
) -> Dict[str, Any]:
    """
    Resuelve la codificación oficial de una factura según las reglas SSSalud de SAPI Web:
    - Si es una prestación Sin Recupero -> sin_recupero=True, codigo_nomenclador=None, cantidadeditable.
    - Si tiene código del nomenclador -> valida contra nomenclador_discapacidad.json:
      * códigos 90 y 91 -> cantidad = 4
      * código 96 (Transporte - Km) -> cantidad manual (entero >= 1)
      * resto de códigos -> cantidad = 1 (o manual si usuario edita sesiones de apoyo)
    """
    catalogo = _cargar_nomenclador()
    codigo_limpio = limpiar_codigo_nomenclador(codigo_nomenclador)
    sin_recupero_por_nombre = es_prestacion_sin_recupero(prestacion_nombre)

    if sin_recupero_solicitado is True or (sin_recupero_por_nombre and not codigo_limpio):
        cantidad_sr = cantidad_solicitada if (cantidad_solicitada and cantidad_solicitada >= 1) else 1
        return {
            "sin_recupero": True,
            "codigo_nomenclador": None,
            "descripcion_nomenclador": prestacion_nombre or "Prestación Sin Recupero",
            "cantidad": int(cantidad_sr),
            "cantidad_editable": True,
            "admite_dependencia": False,
            "dependencia": False,
        }

    if not codigo_limpio or codigo_limpio not in catalogo:
        codigo_inferido = inferir_codigo_por_nombre_prestacion(prestacion_nombre)
        if codigo_inferido and codigo_inferido in catalogo:
            codigo_limpio = codigo_inferido

    if not codigo_limpio or codigo_limpio not in catalogo:
        if sin_recupero_solicitado is False:
            return {
                "sin_recupero": False,
                "codigo_nomenclador": codigo_limpio,
                "descripcion_nomenclador": prestacion_nombre or "Sin código asignado",
                "cantidad": int(cantidad_solicitada) if (cantidad_solicitada and cantidad_solicitada >= 1) else 1,
                "cantidad_editable": True,
                "admite_dependencia": False,
                "dependencia": False,
            }

        cantidad_fallback = cantidad_solicitada if (cantidad_solicitada and cantidad_solicitada >= 1) else 1
        return {
            "sin_recupero": True,
            "codigo_nomenclador": None,
            "descripcion_nomenclador": prestacion_nombre or "Prestación Sin Recupero",
            "cantidad": int(cantidad_fallback),
            "cantidad_editable": True,
            "admite_dependencia": False,
            "dependencia": False,
        }

    item = catalogo[codigo_limpio]
    admite_dep = bool(item.get("admite_dependencia") == 1 or codigo_limpio in CODIGOS_DEPENDENCIA_VALIDOS)

    if codigo_limpio in CODIGOS_CON_CANTIDAD_CUATRO:
        cantidad_final = 4
        cantidad_editable = False
    elif codigo_limpio in CODIGOS_CON_CANTIDAD_MANUAL:
        cantidad_final = int(cantidad_solicitada) if (cantidad_solicitada and cantidad_solicitada >= 1) else 1
        cantidad_editable = True
    else:
        cantidad_cfg = str(item.get("cantidad", "1")).strip()
        if cantidad_solicitada and cantidad_solicitada >= 1 and codigo_limpio == "86":
            cantidad_final = int(cantidad_solicitada)
        else:
            cantidad_final = int(cantidad_cfg) if cantidad_cfg.isdigit() else 1
        cantidad_editable = codigo_limpio == "86"

    return {
        "sin_recupero": False,
        "codigo_nomenclador": codigo_limpio,
        "descripcion_nomenclador": item.get("descripcion", prestacion_nombre or ""),
        "cantidad": cantidad_final,
        "cantidad_editable": cantidad_editable,
        "admite_dependencia": admite_dep,
        "dependencia": bool(dependencia_solicitada and admite_dep),
    }


# ==============================================================================
# 3. CRONOGRAMA DE PRESENTACIÓN SSSALUD (de factura_service.py / cronogramaPresentacion.ts)
# ==============================================================================

NOMBRES_MESES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}


def formatear_periodo_legible(periodo_aaaamm: str) -> str:
    limpio = str(periodo_aaaamm or "").strip()
    if len(limpio) != 6 or not limpio.isdigit():
        return limpio or "S/D"
    anio = int(limpio[:4])
    mes = int(limpio[4:6])
    nombre_mes = NOMBRES_MESES.get(mes, f"Mes {mes}")
    return f"{nombre_mes} {anio}"


def obtener_cronograma_presentacion(fecha_ref: Optional[date] = None) -> Dict[str, Any]:
    """
    Calcula la ventana móvil de 3 períodos prestacionales habilitados por SSSalud.
    Regla: En el mes actual M se presentan los 3 meses cerrados inmediatos anteriores:
    (M-3, M-2, M-1).
    """
    hoy = fecha_ref or date.today()
    mes_actual = hoy.month
    anio_actual = hoy.year

    periodos_habilitados: List[str] = []
    periodos_detalle: List[Dict[str, Any]] = []

    for offset in (3, 2, 1):
        m = mes_actual - offset
        y = anio_actual
        while m <= 0:
            m += 12
            y -= 1
        codigo = f"{y:04d}{m:02d}"
        periodos_habilitados.append(codigo)
        periodos_detalle.append({
            "periodo": codigo,
            "etiqueta": f"{NOMBRES_MESES[m]} {y}",
            "mes": m,
            "anio": y,
        })

    return {
        "fecha_referencia": hoy.isoformat(),
        "mes_calendario_actual": f"{NOMBRES_MESES[mes_actual]} {anio_actual}",
        "periodos_habilitados": periodos_habilitados,
        "periodos_detalle": periodos_detalle,
        "periodo_minimo": periodos_habilitados[0],
        "periodo_maximo": periodos_habilitados[-1],
    }


# ==============================================================================
# 4. SERVICIO MULTICAPA DE EXTRACCIÓN DE QR Y TEXTO PDF (QrService de SAPI Web)
# ==============================================================================

class QrService:
    """
    Servicio para leer e interpretar códigos QR de facturas de AFIP/ARCA
    desde archivos PDF y extraer texto complementario (DNI Afiliado, Período,
    CUIT Receptor, Razón Social Emisor y Prestación sugerida).
    """

    @staticmethod
    def extraer_url_de_pdf_bytes(pdf_bytes: bytes) -> Optional[str]:
        """
        Extrae la URL del QR de AFIP/ARCA desde los bytes de un PDF con arquitectura multicapa:
        1. Enlaces directos embebidos en el PDF (instantáneo)
        2. Imágenes incrustadas en el PDF (QR aislado o página escaneada completa)
        3. Renderizado de alta resolución focalizado en la mitad inferior del PDF
        """
        if not pdf_bytes or len(pdf_bytes) > MAX_PDF_BYTES:
            return None

        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            if doc.page_count == 0:
                doc.close()
                return None

            # PASO 1: Enlaces directos en el PDF
            for page in doc:
                for link in page.get_links():
                    uri = link.get("uri", "")
                    if "afip.gob.ar" in uri or "arca.gob.ar" in uri:
                        doc.close()
                        return uri
                # También buscar si la URL AFIP/ARCA figura en el texto vectorial de la página
                texto_p = page.get_text() or ""
                m_url = re.search(r"https?://[^\s\"'<>]+?(?:afip|arca)\.gob\.ar/fe/qr/\?p=[A-Za-z0-9+/=_-]+", texto_p)
                if m_url:
                    doc.close()
                    return m_url.group(0)

            if not HAS_CV2:
                doc.close()
                return None

            paginas_a_inspeccionar = [doc[0]]
            if doc.page_count > 1:
                paginas_a_inspeccionar.append(doc[-1])

            # PASO 2: Imágenes incrustadas dentro del PDF
            xrefs_vistos = set()
            for p in paginas_a_inspeccionar:
                img_list = p.get_images(full=True)
                for img_info in img_list:
                    xref = img_info[0]
                    if xref in xrefs_vistos:
                        continue
                    xrefs_vistos.add(xref)

                    try:
                        base_image = doc.extract_image(xref)
                        if not base_image or "image" not in base_image:
                            continue
                        image_bytes = base_image["image"]
                        nparr = np.frombuffer(image_bytes, np.uint8)
                        cv_img = cv2.imdecode(nparr, cv2.IMREAD_GRAYSCALE)
                        if cv_img is None:
                            continue

                        h, w = cv_img.shape[:2]
                        if h < 35 or w < 35:
                            continue

                        ratio = w / float(h)

                        # CASO 2A: Imagen cuadrada/aislada (típico QR de AFIP/ARCA)
                        if 0.65 <= ratio <= 1.55 and min(w, h) >= 50:
                            url = QrService._decodificar_matriz_qr(cv_img)
                            if url:
                                doc.close()
                                return url

                        # CASO 2B: Página completa escaneada incrustada como imagen
                        elif (h >= 750 and w >= 550) and (ratio < 0.85 or ratio > 1.25):
                            rois = [
                                cv_img[int(h * 0.60):h, 0:int(w * 0.45)],
                                cv_img[int(h * 0.50):h, 0:int(w * 0.55)],
                                cv_img[int(h * 0.55):h, int(w * 0.45):w],
                                cv_img[int(h * 0.50):h, :],
                                cv_img,
                            ]
                            for roi in rois:
                                if roi.size == 0:
                                    continue
                                url = QrService._decodificar_matriz_qr(roi)
                                if url:
                                    doc.close()
                                    return url
                    except Exception:
                        continue

            # PASO 3: Renderizado de página (mitad inferior y página completa)
            for p in paginas_a_inspeccionar:
                rect = p.rect
                clip_bottom = fitz.Rect(0, rect.height * 0.45, rect.width, rect.height)

                for zoom in (3.0, 2.0):
                    mat = fitz.Matrix(zoom, zoom)
                    pix = p.get_pixmap(matrix=mat, clip=clip_bottom, colorspace=fitz.csGRAY, alpha=False)
                    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width)

                    h_img, w_img = img.shape[:2]
                    sub_rois = [
                        img[:, 0:int(w_img * 0.55)],
                        img[:, int(w_img * 0.45):w_img],
                        img,
                    ]
                    for s_roi in sub_rois:
                        if s_roi.size == 0:
                            continue
                        url = QrService._decodificar_matriz_qr(s_roi)
                        if url:
                            doc.close()
                            return url

                # Fallback: página completa a zoom 2.0 si el QR estuviese en la mitad superior
                mat_full = fitz.Matrix(2.0, 2.0)
                pix_full = p.get_pixmap(matrix=mat_full, colorspace=fitz.csGRAY, alpha=False)
                img_full = np.frombuffer(pix_full.samples, dtype=np.uint8).reshape(pix_full.height, pix_full.width)
                url = QrService._decodificar_matriz_qr(img_full)
                if url:
                    doc.close()
                    return url

            doc.close()
            return None
        except Exception:
            return None

    @staticmethod
    def _decodificar_matriz_qr(img_gray: np.ndarray) -> Optional[str]:
        """
        Intenta decodificar un QR desde una matriz numpy en escala de grises
        aplicando bordes blancos (quiet zone), escalados y filtros morfológicos/Otsu.
        """
        if img_gray is None or img_gray.size == 0:
            return None

        detector = cv2.QRCodeDetector()

        def _es_url_qr_valida(texto: str) -> bool:
            if not texto:
                return False
            return (
                "afip.gob.ar" in texto
                or "arca.gob.ar" in texto
                or "?p=" in texto
                or texto.startswith(("http://", "https://"))
            )

        def _intentar(imagen: np.ndarray) -> Optional[str]:
            if HAS_PYZBAR:
                try:
                    decoded_objs = pyzbar_decode(imagen)
                    for obj in decoded_objs:
                        texto = obj.data.decode("utf-8", errors="ignore")
                        if _es_url_qr_valida(texto):
                            return texto
                except Exception:
                    pass

            try:
                data, _, _ = detector.detectAndDecode(imagen)
                if _es_url_qr_valida(data):
                    return data
            except Exception:
                pass

            try:
                ok, decoded_info, _, _ = detector.detectAndDecodeMulti(imagen)
                if ok and decoded_info:
                    for item in decoded_info:
                        if _es_url_qr_valida(item):
                            return item
            except Exception:
                pass

            return None

        # 1. Con borde blanco (quiet zone)
        for border_px in (40, 15):
            padded = cv2.copyMakeBorder(
                img_gray, border_px, border_px, border_px, border_px,
                cv2.BORDER_CONSTANT, value=255
            )
            res = _intentar(padded)
            if res:
                return res

        # 2. Imagen directa
        res = _intentar(img_gray)
        if res:
            return res

        h, w = img_gray.shape[:2]
        max_dim = max(h, w)

        # 3. QRs pequeños (< 350px)
        if max_dim < 350:
            for scale in (3.0, 2.0, 3.5, 4.0, 5.0):
                for interp in (cv2.INTER_CUBIC, cv2.INTER_LANCZOS4, cv2.INTER_LINEAR, cv2.INTER_NEAREST):
                    resized = cv2.resize(img_gray, None, fx=scale, fy=scale, interpolation=interp)
                    for b_px in (40, 20):
                        padded_r = cv2.copyMakeBorder(
                            resized, b_px, b_px, b_px, b_px, cv2.BORDER_CONSTANT, value=255
                        )
                        res = _intentar(padded_r)
                        if res:
                            return res
                        _, thresh_r = cv2.threshold(padded_r, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                        res = _intentar(thresh_r)
                        if res:
                            return res

        # 4. Imágenes muy grandes (> 1400px)
        if max_dim > 1400:
            for target_max in (1000.0, 800.0):
                scale = target_max / max_dim
                resized = cv2.resize(img_gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
                padded_r = cv2.copyMakeBorder(resized, 30, 30, 30, 30, cv2.BORDER_CONSTANT, value=255)
                res = _intentar(padded_r)
                if res:
                    return res

        # 5. Imágenes medianas (350px a 2200px)
        if 350 <= max_dim <= 2200:
            for scale in (1.25, 1.5):
                resized = cv2.resize(img_gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_LINEAR)
                padded_r = cv2.copyMakeBorder(resized, 30, 30, 30, 30, cv2.BORDER_CONSTANT, value=255)
                res = _intentar(padded_r)
                if res:
                    return res

        # 6. Preprocesamiento con CLAHE y binarización Otsu
        try:
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
            enhanced = clahe.apply(img_gray)
            _, thresh = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            for candidate in (enhanced, thresh):
                for b_px in (40, 20):
                    padded_c = cv2.copyMakeBorder(
                        candidate, b_px, b_px, b_px, b_px, cv2.BORDER_CONSTANT, value=255
                    )
                    res = _intentar(padded_c)
                    if res:
                        return res
                    if max_dim < 350:
                        for interp in (cv2.INTER_CUBIC, cv2.INTER_NEAREST):
                            up_c = cv2.resize(padded_c, None, fx=3.0, fy=3.0, interpolation=interp)
                            res = _intentar(up_c)
                            if res:
                                return res
        except Exception:
            pass

        return None

    @staticmethod
    def decodificar_url_afip(url: str) -> Optional[Dict[str, Any]]:
        """
        Decodifica el parámetro 'p' en Base64 de la URL del QR de AFIP/ARCA
        y retorna los datos fiscales estructurados.
        """
        try:
            parsed_url = urllib.parse.urlparse(url)
            query_params = urllib.parse.parse_qs(parsed_url.query)

            if "p" not in query_params:
                if "?p=" in url:
                    base64_str = url.split("?p=")[1].split("&")[0]
                else:
                    return None
            else:
                base64_str = query_params["p"][0]

            padding = 4 - (len(base64_str) % 4)
            if padding != 4:
                base64_str += "=" * padding

            json_bytes = base64.b64decode(base64_str)
            json_str = json_bytes.decode("utf-8")
            datos_afip = json.loads(json_str)

            tipo_cmp_raw = datos_afip.get("tipoCmp")
            try:
                tipo_cmp_afip = int(tipo_cmp_raw) if tipo_cmp_raw is not None else None
            except (TypeError, ValueError):
                tipo_cmp_afip = None

            tipo_comprobante = normalizar_tipo_comprobante_afip(tipo_cmp_afip)
            es_nota_credito = es_nota_credito_comprobante(tipo_comprobante, tipo_cmp_afip)
            es_recibo = es_recibo_comprobante(tipo_comprobante, tipo_cmp_afip)
            codigo_sssalud = obtener_codigo_sssalud_desde_tipo_comprobante(
                tipo_comprobante=tipo_comprobante,
                tipo_cmp_afip=tipo_cmp_afip,
            )

            nro_doc_rec_raw = datos_afip.get("nroDocRec")
            cuit_receptor_qr = re.sub(r"\D", "", str(nro_doc_rec_raw or "")) if nro_doc_rec_raw is not None else ""
            if len(cuit_receptor_qr) != 11:
                cuit_receptor_qr = cuit_receptor_qr if cuit_receptor_qr else ""

            return {
                "fecha_emision": datos_afip.get("fecha"),
                "cuit_emisor": str(datos_afip.get("cuit", "")).strip(),
                "punto_venta": datos_afip.get("ptoVta"),
                "tipo_comprobante_afip": tipo_cmp_afip,
                "tipo_comprobante": tipo_comprobante,
                "tipo_comprobante_etiqueta": obtener_etiqueta_tipo_comprobante(tipo_comprobante, tipo_cmp_afip),
                "codigo_comprobante_sssalud": codigo_sssalud,
                "es_nota_credito": es_nota_credito,
                "es_recibo": es_recibo,
                "numero_comprobante": datos_afip.get("nroCmp"),
                "importe_total": float(datos_afip.get("importe", 0.0)),
                "moneda": datos_afip.get("moneda", "PES"),
                "tipo_doc_receptor": datos_afip.get("tipoDocRec"),
                "nro_doc_receptor": datos_afip.get("nroDocRec"),
                "cuit_receptor": cuit_receptor_qr,
                "tipo_cod_autorizacion": datos_afip.get("codAut"),
                "cae": str(datos_afip.get("codAut", "")),
                "raw_data": datos_afip,
            }
        except Exception:
            return None

    @staticmethod
    def extraer_texto_de_pdf(pdf_bytes: bytes) -> str:
        """
        Extrae texto vectorial de todas las páginas del PDF.
        Si tiene menos de 30 caracteres y Tesseract está disponible, aplica OCR sobre la primera página.
        """
        texto_completo = ""
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            for page in doc:
                texto_completo += page.get_text() + "\n"

            if len(texto_completo.strip()) < 30 and HAS_TESSERACT and doc.page_count > 0:
                try:
                    page = doc[0]
                    mat = fitz.Matrix(1.5, 1.5)
                    pix = page.get_pixmap(matrix=mat, colorspace=fitz.csGRAY, alpha=False)
                    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width)
                    texto_ocr = pytesseract.image_to_string(img, lang="spa")
                    texto_completo += "\n" + texto_ocr
                except Exception:
                    pass

            doc.close()
        except Exception:
            pass
        return texto_completo

    @staticmethod
    def extraer_cuit_receptor_de_pdf(texto: str, cuit_emisor: Optional[str] = None) -> Optional[str]:
        """
        Extrae el CUIT del receptor (cliente / obra social / institución) desde el texto de la factura
        cuando el QR no incluye nroDocRec o viene vacío.
        """
        if not texto:
            return None

        cuit_emisor_clean = re.sub(r"\D", "", str(cuit_emisor or ""))

        patrones_receptor = [
            r"(?:cliente|receptor|destinatario|obra\s*social|señor(?:es)?|sr\.?\s*\/?\s*sra\.?|raz[oó]n\s*social)[^\n]{0,120}?c\.?u\.?i\.?t\.?\s*[:#º°]?\s*(\d{2}[-\s.]?\d{8}[-\s.]?\d)",
            r"c\.?u\.?i\.?t\.?\s*(?:cliente|receptor|destinatario|obra\s*social)?\s*[:#º°]?\s*(\d{2}[-\s.]?\d{8}[-\s.]?\d)",
            r"\b(30[-\s.]?\d{8}[-\s.]?\d)\b",
            r"\b((?:20|23|24|27|30|33|34)[-\s.]?\d{8}[-\s.]?\d)\b",
        ]

        for patron in patrones_receptor:
            for match in re.finditer(patron, texto, re.IGNORECASE):
                candidato = re.sub(r"\D", "", match.group(1))
                if len(candidato) == 11 and candidato != cuit_emisor_clean:
                    return candidato

        return None

    @staticmethod
    def extraer_dni_de_pdf(
        texto: str,
        nro_cmp: Optional[int] = None,
        cuit_emisor: Optional[str] = None,
        cuit_receptor: Optional[str] = None,
    ) -> Optional[str]:
        """
        Busca posibles DNIs de afiliados dentro del texto de la factura.
        Garantiza que el DNI retornado tenga estrictamente 7 u 8 dígitos numéricos,
        excluyendo el CUIT/DNI del prestador emisor, el CUIT receptor y el número de comprobante.
        """
        if not texto:
            return None

        excluir = set()
        if nro_cmp is not None:
            excluir.add(str(nro_cmp))

        if cuit_emisor:
            cuit_clean = re.sub(r"\D", "", str(cuit_emisor))
            excluir.add(cuit_clean)
            if len(cuit_clean) == 11:
                excluir.add(cuit_clean[2:10])
                excluir.add(cuit_clean[2:10].lstrip("0"))
                excluir.add(cuit_clean[:8])
                excluir.add(cuit_clean[3:])

        if cuit_receptor:
            cuit_rec_clean = re.sub(r"\D", "", str(cuit_receptor))
            excluir.add(cuit_rec_clean)
            if len(cuit_rec_clean) == 11:
                excluir.add(cuit_rec_clean[2:10])
                excluir.add(cuit_rec_clean[2:10].lstrip("0"))
                excluir.add(cuit_rec_clean[:8])
                excluir.add(cuit_rec_clean[3:])

        def _limpiar_y_validar(candidato: str) -> Optional[str]:
            solo_digitos = re.sub(r"\D", "", candidato)
            if len(solo_digitos) not in (7, 8):
                return None
            if solo_digitos in excluir or solo_digitos.lstrip("0") in excluir:
                return None
            val = int(solo_digitos)
            if 5000000 <= val <= 99999999:
                return solo_digitos
            return None

        # 1. Precedido por palabras clave (DNI, Documento, Afiliado, Beneficiario, Paciente)
        patrones_dni = [
            r"(?:dni|d\.n\.i\.|documento|doc\.?|nro\.?\s*doc\.?|n[uú]mero\s*de\s*documento|afiliado|beneficiario|paciente)[\s:º°#-]*([0-9]{1,2}\.?[0-9]{3}\.?[0-9]{3})",
            r"(?:dni|d\.n\.i\.|documento|doc\.?|afiliado|beneficiario|paciente)[\s:º°#-]*([0-9]{7,8})\b",
        ]
        for patron in patrones_dni:
            for match in re.finditer(patron, texto, re.IGNORECASE):
                dni_valido = _limpiar_y_validar(match.group(1))
                if dni_valido:
                    return dni_valido

        # 2. Formato con puntos (ej: 45.123.456 o 8.123.456)
        patron_puntos = r"\b([0-9]{1,2}\.[0-9]{3}\.[0-9]{3})\b"
        for m in re.findall(patron_puntos, texto):
            dni_valido = _limpiar_y_validar(m)
            if dni_valido:
                return dni_valido

        # 3. Formato con comas
        patron_comas = r"\b([0-9]{1,2},[0-9]{3},[0-9]{3})\b"
        for m in re.findall(patron_comas, texto):
            dni_valido = _limpiar_y_validar(m)
            if dni_valido:
                return dni_valido

        # 4. Números aislados de 7 u 8 dígitos que no sean parte de CUIT ni CAE
        texto_sin_cuits = re.sub(r"\b\d{2}-\d{8}-\d\b", " ", texto)
        for m in re.findall(r"(?<![0-9-])([0-9]{7,8})(?![0-9-])", texto_sin_cuits):
            dni_valido = _limpiar_y_validar(m)
            if dni_valido:
                return dni_valido

        return None

    @staticmethod
    def extraer_periodo_de_pdf(texto: str) -> Optional[str]:
        """
        Extrae el período prestacional de la factura y lo devuelve en formato AAAAMM.
        """
        if not texto:
            return None

        texto_lower = texto.lower()

        # Priorizar bloques cercanos a "período", "mes", "correspondiente a", "prestación"
        meses_map = {
            "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
            "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
            "septiembre": "09", "setiembre": "09", "octubre": "10",
            "noviembre": "11", "diciembre": "12",
            "ene": "01", "feb": "02", "mar": "03", "abr": "04",
            "may": "05", "jun": "06", "jul": "07", "ago": "08",
            "sep": "09", "set": "09", "oct": "10", "nov": "11", "dic": "12",
        }

        # 1. Buscar nombre de mes + año (ej: "Octubre 2025", "Mes de Marzo de 2026")
        patron_mes_nombre = (
            r"\b(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre|"
            r"ene|feb|mar|abr|may|jun|jul|ago|sep|set|oct|nov|dic)"
            r"[\s./-]*(?:de\s*|del\s*)?(20[2-3][0-9])\b"
        )
        match_mes = re.search(patron_mes_nombre, texto_lower)
        if match_mes:
            mes_str = meses_map[match_mes.group(1)]
            anio_str = match_mes.group(2)
            return f"{anio_str}{mes_str}"

        # 1B. "Período Facturado Desde: DD/MM/AAAA Hasta: DD/MM/AAAA" (clásico de factura electrónica AFIP)
        patron_periodo_afip = (
            r"per[ií]odo\s*facturado\s*desde\s*[:#]?\s*\d{2}[/-](0[1-9]|1[0-2])[/-](20[2-3][0-9])"
        )
        match_afip = re.search(patron_periodo_afip, texto_lower)
        if match_afip:
            return f"{match_afip.group(2)}{match_afip.group(1)}"

        # 2. Formato MM/AAAA o MM-AAAA (sin estar precedido por día DD/)
        patron_mm_aaaa = r"(?<![0-9/.-])\b(0?[1-9]|1[0-2])\s*[/-]\s*(20[2-3][0-9])\b(?![/-][0-9])"
        match_mm_aaaa = re.search(patron_mm_aaaa, texto_lower)
        if match_mm_aaaa:
            mes_num = int(match_mm_aaaa.group(1))
            anio_str = match_mm_aaaa.group(2)
            return f"{anio_str}{mes_num:02d}"

        # 3. Formato AAAA-MM o AAAA/MM (sin día después)
        patron_aaaa_mm = r"(?<![0-9/.-])\b(20[2-3][0-9])\s*[/-]\s*(0?[1-9]|1[0-2])\b(?![/-][0-9])"
        match_aaaa_mm = re.search(patron_aaaa_mm, texto_lower)
        if match_aaaa_mm:
            anio_str = match_aaaa_mm.group(1)
            mes_num = int(match_aaaa_mm.group(2))
            return f"{anio_str}{mes_num:02d}"

        # 4. Formato directo AAAAMM (ej: "Periodo: 202603")
        patron_aaaamm = r"(?:per[ií]odo|mes)\s*[:#-]?\s*(20[2-3][0-9])(0[1-9]|1[0-2])\b"
        match_aaaamm = re.search(patron_aaaamm, texto_lower)
        if match_aaaamm:
            return f"{match_aaaamm.group(1)}{match_aaaamm.group(2)}"

        return None

    @staticmethod
    def extraer_razon_social_emisor_de_pdf(texto: str, cuit_emisor: Optional[str] = None) -> str:
        """
        Extrae la Razón Social / Apellido y Nombre del prestador emisor desde el encabezado del PDF AFIP.
        Como en esta versión demo no existe base de datos previa de prestadores, permite mostrar el nombre real.
        """
        if not texto:
            return f"Prestador CUIT {formatear_cuit_guiones(cuit_emisor)}"

        lineas = [ln.strip() for ln in texto.splitlines() if ln.strip()]

        # En facturas electrónicas AFIP, el primer "Razón Social:" corresponde al emisor,
        # y el segundo "Apellido y Nombre / Razón Social:" corresponde al receptor.
        patrones_emisor = [
            r"^raz[oó]n\s*social\s*[:#]?\s*(.+)$",
            r"^emisor\s*[:#]?\s*(.+)$",
            r"^titular\s*[:#]?\s*(.+)$",
        ]
        for ln in lineas[:25]:
            for pat in patrones_emisor:
                m = re.match(pat, ln, re.IGNORECASE)
                if m:
                    valor = m.group(1).strip()
                    # Evitar tomar etiquetas vacías o datos del receptor
                    if len(valor) >= 3 and not re.search(r"^(fecha|cuit|domicilio|condici[oó]n|iva)\b", valor, re.IGNORECASE):
                        return valor[:80]

        # Buscar la línea inmediatamente posterior a "Razón Social:" si viene en línea separada
        for idx, ln in enumerate(lineas[:20]):
            if re.match(r"^raz[oó]n\s*social\s*:?$", ln, re.IGNORECASE) and idx + 1 < len(lineas):
                cand = lineas[idx + 1].strip()
                if len(cand) >= 3 and not re.search(r"^(domicilio|fecha|cuit|condici[oó]n|iva|punto)\b", cand, re.IGNORECASE):
                    return cand[:80]

        # Heurística adicional: primeras líneas en mayúsculas antes de "Domicilio Comercial" o "Condición frente al IVA"
        ignorar = {
            "ORIGINAL", "DUPLICADO", "TRIPLICADO", "FACTURA", "RECIBO",
            "A", "B", "C", "M", "COD. 011", "COD. 006", "COD. 001", "COD. 015",
            "COD. 01", "COD. 06", "COD. 11", "COD. 15",
        }
        for ln in lineas[:12]:
            ln_up = ln.upper().strip()
            if ln_up in ignorar or ln_up.startswith("COD.") or ln_up.startswith("CÓD."):
                continue
            if any(k in ln_up for k in ("DOMICILIO", "CONDICIÓN", "CONDICION", "PUNTO DE VENTA", "COMP. NRO", "FECHA DE EMISIÓN", "CUIT:", "INGRESOS BRUTOS")):
                continue
            if re.match(r"^[A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s.,&-]{4,65}$", ln_up) and not re.search(r"\d{4,}", ln_up):
                return ln.strip()[:80]

        return f"Prestador ({formatear_cuit_guiones(cuit_emisor)})"

    @staticmethod
    def extraer_nombre_afiliado_de_pdf(texto: str) -> Optional[str]:
        """
        Intenta extraer el nombre del afiliado / paciente / beneficiario desde el detalle de la factura.
        """
        if not texto:
            return None

        patrones = [
            r"(?:afiliado|beneficiario|paciente|alumno|concurrente)\s*[:#-]?\s*([A-ZÁÉÍÓÚÑa-záéíóúñ][A-ZÁÉÍÓÚÑa-záéíóúñ\s,.'-]{3,50}?)(?=\s+(?:dni|d\.n\.i|doc|cuit|mes|per[ií]odo|prestaci[oó]n|c[oó]digo|\d|$)|\n)",
        ]
        for pat in patrones:
            m = re.search(pat, texto, re.IGNORECASE)
            if m:
                nombre = re.sub(r"\s+", " ", m.group(1)).strip(" ,.-:")
                if len(nombre) >= 4 and not any(w in nombre.lower() for w in ("obra social", "factura", "total", "importe", "servicio")):
                    return nombre[:60]
        return None

    @staticmethod
    def detectar_prestacion_en_texto_pdf(texto: str) -> Dict[str, Any]:
        """
        Analiza el cuerpo descriptivo del PDF para sugerir automáticamente el código del Nomenclador SSSalud
        o detectar si es una prestación Sin Recupero.
        """
        if not texto:
            # Por defecto sugerimos código 86 (Prestación de Apoyo) que es el más habitual en facturación individual
            return resolver_codificacion_prestacion(codigo_nomenclador="86", prestacion_nombre="Prestacion de Apoyo")

        # 1. ¿Menciona alguna prestación Sin Recupero explícita?
        if es_prestacion_sin_recupero(texto):
            texto_norm = normalizar_texto(texto)
            for kw in PRESTACIONES_SIN_RECUPERO_KEYWORDS:
                if kw in texto_norm:
                    titulo = kw.title() + " (Sin Recupero)"
                    return resolver_codificacion_prestacion(
                        codigo_nomenclador=None,
                        prestacion_nombre=titulo,
                        sin_recupero_solicitado=True,
                    )

        # 2. ¿Menciona alguna prestación del nomenclador por nombre?
        codigo_inf = inferir_codigo_por_nombre_prestacion(texto)
        if codigo_inf:
            # Si es transporte (96), intentar detectar cantidad de Km en el texto
            cant_sugerida = 1
            if codigo_inf == "96":
                m_km = re.search(r"\b(\d{1,5})\s*(?:km|kil[oó]metros?)\b", texto, re.IGNORECASE)
                if m_km:
                    cant_sugerida = max(1, int(m_km.group(1)))
            return resolver_codificacion_prestacion(
                codigo_nomenclador=codigo_inf,
                cantidad_solicitada=cant_sugerida,
            )

        # 3. Default editable en Paso 3: Código 86 - Prestación de Apoyo
        return resolver_codificacion_prestacion(codigo_nomenclador="86", prestacion_nombre="Prestacion de Apoyo")

    @staticmethod
    def procesar_pdf_bytes(pdf_bytes: bytes) -> Dict[str, Any]:
        """
        Flujo integral de SAPI Web para procesar los bytes de un PDF de factura:
        1. Extrae URL del QR AFIP/ARCA
        2. Decodifica el JSON Base64 de AFIP
        3. Extrae el texto del PDF para detectar:
           - CUIT Receptor (si no vino en el QR)
           - DNI del afiliado (7-8 dígitos)
           - Período prestacional (AAAAMM)
           - Razón Social del prestador emisor
           - Nombre del afiliado (si figura en el concepto)
           - Prestación / Código Nomenclador SSSalud sugerido
           - Detección de Nota de Crédito o Recibo por texto
        """
        if not pdf_bytes:
            return {
                "exito": False,
                "error": "El archivo PDF está vacío o dañado.",
                "datos": None,
            }

        if len(pdf_bytes) > MAX_PDF_BYTES:
            return {
                "exito": False,
                "error": "El archivo PDF supera el tamaño máximo permitido de 15 MB.",
                "datos": None,
            }

        url_qr = QrService.extraer_url_de_pdf_bytes(pdf_bytes)
        if not url_qr:
            return {
                "exito": False,
                "error": "No se detectó un código QR válido de AFIP/ARCA en el documento PDF.",
                "datos": None,
            }

        datos_qr = QrService.decodificar_url_afip(url_qr)
        if not datos_qr:
            return {
                "exito": False,
                "error": "Se encontró un QR pero no pertenece al formato de facturación electrónica de AFIP/ARCA.",
                "datos": None,
            }

        texto_pdf = QrService.extraer_texto_de_pdf(pdf_bytes)

        cuit_receptor_qr = str(datos_qr.get("cuit_receptor") or "").strip()
        cuit_receptor_pdf = QrService.extraer_cuit_receptor_de_pdf(
            texto_pdf,
            cuit_emisor=datos_qr.get("cuit_emisor"),
        )
        cuit_receptor_final = cuit_receptor_qr or cuit_receptor_pdf or ""
        datos_qr["cuit_receptor"] = cuit_receptor_final
        datos_qr["cuit_receptor_detectado_por"] = (
            "qr" if cuit_receptor_qr else ("pdf" if cuit_receptor_pdf else "no_detectado")
        )

        # Extraer DNI del afiliado excluyendo el nro de comprobante y los CUITs
        dni_detectado = QrService.extraer_dni_de_pdf(
            texto_pdf,
            nro_cmp=datos_qr.get("numero_comprobante"),
            cuit_emisor=datos_qr.get("cuit_emisor"),
            cuit_receptor=cuit_receptor_final,
        )
        datos_qr["dni_afiliado"] = dni_detectado or ""

        # Extraer período AAAAMM (fallback a mes de emisión si no se encuentra en concepto)
        periodo_detectado = QrService.extraer_periodo_de_pdf(texto_pdf)
        if not periodo_detectado:
            fecha_em = str(datos_qr.get("fecha_emision") or "")
            m_fecha = re.match(r"^(20\d{2})[-/](\d{2})", fecha_em)
            if m_fecha:
                periodo_detectado = f"{m_fecha.group(1)}{m_fecha.group(2)}"
        datos_qr["periodo_detectado"] = periodo_detectado or ""

        # Extraer Razón Social del emisor y nombre del afiliado
        datos_qr["razon_social_emisor"] = QrService.extraer_razon_social_emisor_de_pdf(
            texto_pdf,
            cuit_emisor=datos_qr.get("cuit_emisor"),
        )
        datos_qr["nombre_afiliado_detectado"] = QrService.extraer_nombre_afiliado_de_pdf(texto_pdf) or ""

        # Sugerir codificación SSSalud desde el texto de la factura
        codificacion_sugerida = QrService.detectar_prestacion_en_texto_pdf(texto_pdf)
        datos_qr["codificacion_sugerida"] = codificacion_sugerida

        # Verificación adicional por texto si es Nota de Crédito o Recibo
        texto_norm = (texto_pdf or "").lower()
        if not datos_qr.get("es_nota_credito"):
            if re.search(r"\bnota\s+de\s+cr[eé]dito\b", texto_norm):
                datos_qr["es_nota_credito"] = True
                letra_actual = str(datos_qr.get("tipo_comprobante", "FA_C")).split("_")[-1]
                datos_qr["tipo_comprobante"] = f"NC_{letra_actual}"
                datos_qr["tipo_comprobante_etiqueta"] = f"Nota de Crédito {letra_actual}"

        if not datos_qr.get("es_nota_credito") and not datos_qr.get("es_recibo"):
            if re.search(r"\brecibo\s+[abcm]\b", texto_norm) or re.search(r"\bcod\.?\s*0?15\b", texto_norm):
                letra_actual = str(datos_qr.get("tipo_comprobante", "FA_C")).split("_")[-1]
                datos_qr["es_recibo"] = True
                datos_qr["tipo_comprobante"] = f"REC_{letra_actual}"
                datos_qr["codigo_comprobante_sssalud"] = obtener_codigo_sssalud_desde_tipo_comprobante(
                    datos_qr["tipo_comprobante"],
                    datos_qr.get("tipo_comprobante_afip"),
                )
                datos_qr["tipo_comprobante_etiqueta"] = obtener_etiqueta_tipo_comprobante(
                    datos_qr["tipo_comprobante"],
                    datos_qr.get("tipo_comprobante_afip"),
                )

        datos_qr["registro_identificador"] = generar_identificador_factura(
            cuit_prestador=datos_qr.get("cuit_emisor"),
            tipo_comprobante=datos_qr.get("tipo_comprobante"),
            punto_venta=datos_qr.get("punto_venta"),
            numero_comprobante=datos_qr.get("numero_comprobante"),
            tipo_cmp_afip=datos_qr.get("tipo_comprobante_afip"),
        )
        datos_qr["url_qr"] = url_qr

        return {
            "exito": True,
            "error": None,
            "datos": datos_qr,
        }


# ==============================================================================
# 5. VALIDACIÓN DE CUIT INSTITUCIONAL RECEPTOR (Opcional en versión Demo)
# ==============================================================================

def limpiar_cuit_institucional(cuit_input: Optional[str]) -> str:
    """Limpia un CUIT ingresado manualmente dejando sólo dígitos."""
    return re.sub(r"\D", "", str(cuit_input or ""))


def validar_formato_cuit_opcional(cuit_input: Optional[str]) -> Tuple[bool, str, str]:
    """
    Valida el CUIT Institucional opcional ingresado por el usuario.
    Retorna: (es_valido, cuit_limpio_11_digitos, mensaje_estado)
    - Si está vacío: (True, "", "Sin filtro de CUIT institucional (se aceptarán facturas de cualquier receptor)")
    - Si tiene 11 dígitos: (True, digits, "Verificando que las facturas correspondan al CUIT XX-XXXXXXXX-X")
    - Si tiene otra longitud: (False, digits, "El CUIT institucional debe tener exactamente 11 dígitos numéricos.")
    """
    limpio = limpiar_cuit_institucional(cuit_input)
    if not limpio:
        return (
            True,
            "",
            "Opcional: Sin CUIT institucional configurado (no se filtrará por CUIT receptor).",
        )
    if len(limpio) != 11:
        return (
            False,
            limpio,
            f"El CUIT ingresado tiene {len(limpio)} dígitos. Debe contener exactamente 11 dígitos numéricos (ej: 30-12345678-9).",
        )
    return (
        True,
        limpio,
        f"Verificación activa para CUIT {formatear_cuit_guiones(limpio)}.",
    )


def validar_cuit_receptor_institucion(
    datos_qr: Dict[str, Any],
    cuit_institucion_esperado: Optional[str],
) -> Tuple[bool, Optional[str]]:
    """
    Verifica que la factura escaneada corresponda al CUIT de la institución configurado por el usuario.
    Si el usuario dejó el CUIT institucional en blanco (opcional), siempre retorna (True, None).
    """
    cuit_esperado = limpiar_cuit_institucional(cuit_institucion_esperado)
    if not cuit_esperado or len(cuit_esperado) != 11:
        return True, None

    cuit_receptor = limpiar_cuit_institucional(datos_qr.get("cuit_receptor"))
    if not cuit_receptor:
        return (
            False,
            f"La factura no registra el CUIT de la institución receptora (CUIT esperado: {formatear_cuit_guiones(cuit_esperado)}).",
        )

    if cuit_receptor != cuit_esperado:
        return (
            False,
            f"La factura no corresponde a la institución (CUIT Receptor detectado: {formatear_cuit_guiones(cuit_receptor)} · CUIT Institucional esperado: {formatear_cuit_guiones(cuit_esperado)}).",
        )

    return True, None


# ==============================================================================
# 6. ESCANEO MASIVO DE LOTE Y ADAPTACIÓN SSSALUD (Paso 1 -> Paso 2)
# ==============================================================================

def formatear_fecha_ddmmaaaa(fecha_iso: Optional[str]) -> str:
    if not fecha_iso:
        return "S/D"
    partes = str(fecha_iso).strip().split("T")[0].split("-")
    if len(partes) == 3:
        return f"{partes[2]}/{partes[1]}/{partes[0]}"
    return str(fecha_iso)


def escanear_lote_pdfs(
    archivos_pdf: List[Tuple[str, bytes]],
    cuit_institucional: Optional[str] = None,
    progreso_callback: Optional[Callable[[int, int, str], None]] = None,
) -> Dict[str, Any]:
    """
    Procesa un lote de hasta 15 archivos PDF (nombre_archivo, bytes) replicando
    `escanear_pdf_masivo` de SAPI Web, sin auditoría contra base de datos de autorizaciones.

    Retorna:
    - 'procesadas': lista de facturas válidas listas para el Paso 2 y Paso 3
    - 'rechazadas': lista de facturas rechazadas (Nota de Crédito, CUIT Receptor no coincide, Sin QR)
    """
    procesadas: List[Dict[str, Any]] = []
    rechazadas: List[Dict[str, Any]] = []

    lote_limitado = archivos_pdf[:MAX_PDFS_POR_TANDA_DEMO]
    total = len(lote_limitado)

    for idx, (nombre_archivo, pdf_bytes) in enumerate(lote_limitado):
        if progreso_callback:
            progreso_callback(idx + 1, total, nombre_archivo)

        if not nombre_archivo.lower().endswith(".pdf"):
            rechazadas.append({
                "id": f"rech_{idx}",
                "archivo": nombre_archivo,
                "motivo": "El archivo no tiene extensión .pdf válida.",
                "tipo_error": "sin_qr",
                "pdf_bytes": pdf_bytes,
            })
            continue

        resultado = QrService.procesar_pdf_bytes(pdf_bytes)
        if not resultado["exito"]:
            rechazadas.append({
                "id": f"rech_{idx}",
                "archivo": nombre_archivo,
                "motivo": resultado["error"] or "No se pudo leer el código QR del PDF.",
                "tipo_error": "sin_qr",
                "pdf_bytes": pdf_bytes,
            })
            continue

        datos_qr = resultado["datos"]

        # 1. Rechazar Notas de Crédito (no permitidas en carga de facturas SSSalud)
        if datos_qr.get("es_nota_credito"):
            etiqueta_nc = datos_qr.get("tipo_comprobante_etiqueta", "Nota de Crédito")
            rechazadas.append({
                "id": f"rech_{idx}",
                "archivo": nombre_archivo,
                "motivo": (
                    f"Comprobante rechazado: se detectó una {etiqueta_nc}. "
                    "En el módulo de Carga de Facturas sólo se admiten Facturas y Recibos."
                ),
                "tipo_error": "nota_credito",
                "cuit_emisor": datos_qr.get("cuit_emisor"),
                "cuit_receptor": datos_qr.get("cuit_receptor"),
                "pdf_bytes": pdf_bytes,
            })
            continue

        # 2. Validar CUIT Receptor Institucional (si el usuario ingresó su CUIT opcional)
        cuit_ok, motivo_cuit = validar_cuit_receptor_institucion(datos_qr, cuit_institucional)
        if not cuit_ok:
            rechazadas.append({
                "id": f"rech_{idx}",
                "archivo": nombre_archivo,
                "motivo": motivo_cuit or "La factura no corresponde al CUIT de la institución.",
                "tipo_error": "cuit_invalido",
                "cuit_emisor": datos_qr.get("cuit_emisor"),
                "cuit_receptor": datos_qr.get("cuit_receptor"),
                "pdf_bytes": pdf_bytes,
            })
            continue

        # 3. Factura aceptada: preparar estructura para Informe (Paso 2) y Revisión (Paso 3)
        codif = datos_qr.get("codificacion_sugerida") or resolver_codificacion_prestacion(
            codigo_nomenclador="86",
            prestacion_nombre="Prestacion de Apoyo",
        )
        sin_rec = bool(codif.get("sin_recupero", False))
        cod_nom = "" if sin_rec else str(codif.get("codigo_nomenclador") or "86")
        desc_nom = str(codif.get("descripcion_nomenclador") or "Prestacion de Apoyo")
        cantidad_inicial = int(codif.get("cantidad") or 1)

        procesadas.append({
            "id": f"fac_{idx}_{datos_qr.get('registro_identificador', idx)}",
            "archivo": nombre_archivo,
            "pdf_bytes": pdf_bytes,
            # Datos Emisor / Prestador
            "prestador_nombre": datos_qr.get("razon_social_emisor") or f"Prestador ({formatear_cuit_guiones(datos_qr.get('cuit_emisor'))})",
            "cuit_emisor": str(datos_qr.get("cuit_emisor") or ""),
            "cuit_receptor": str(datos_qr.get("cuit_receptor") or ""),
            "cuit_receptor_origen": str(datos_qr.get("cuit_receptor_detectado_por") or ""),
            # Datos Comprobante Fiscal
            "punto_venta": int(datos_qr.get("punto_venta") or 0),
            "numero_comprobante": int(datos_qr.get("numero_comprobante") or 0),
            "comprobante_visible": formatear_comprobante_visible(
                datos_qr.get("punto_venta"),
                datos_qr.get("numero_comprobante"),
            ),
            "registro_identificador": datos_qr.get("registro_identificador") or "",
            "tipo_comprobante": datos_qr.get("tipo_comprobante") or "FA_C",
            "tipo_comprobante_afip": datos_qr.get("tipo_comprobante_afip"),
            "tipo_comprobante_etiqueta": datos_qr.get("tipo_comprobante_etiqueta") or "Factura C",
            "codigo_comprobante_sssalud": int(datos_qr.get("codigo_comprobante_sssalud") or 5),
            "es_nota_credito": False,
            "es_recibo": bool(datos_qr.get("es_recibo", False)),
            "cae": str(datos_qr.get("cae") or ""),
            "fecha_emision": str(datos_qr.get("fecha_emision") or ""),
            "fecha_emision_ddmmaaaa": formatear_fecha_ddmmaaaa(datos_qr.get("fecha_emision")),
            "importe_total": float(datos_qr.get("importe_total") or 0.0),
            "moneda": str(datos_qr.get("moneda") or "PES"),
            # Datos editables en Paso 3 (Revisión y Autorización)
            "periodo": str(datos_qr.get("periodo_detectado") or ""),
            "afiliado_dni": str(datos_qr.get("dni_afiliado") or ""),
            "afiliado_nombre": str(datos_qr.get("nombre_afiliado_detectado") or ""),
            "prestacion": desc_nom,
            "codigo_nomenclador": cod_nom,
            "cantidad": cantidad_inicial,
            "sin_recupero": sin_rec,
            "dependencia": bool(codif.get("dependencia", False)),
            "seleccionada": True,
            "url_qr": datos_qr.get("url_qr") or "",
        })

    # Recalcular estados iniciales de revisión para todas las procesadas
    recalcular_estados_lote(procesadas)

    return {
        "procesadas": procesadas,
        "rechazadas": rechazadas,
        "total_recibidas": len(archivos_pdf),
        "total_escaneadas": len(lote_limitado),
        "excedidas_demo": max(0, len(archivos_pdf) - MAX_PDFS_POR_TANDA_DEMO),
    }


def adaptar_pdf_a_primera_pagina(pdf_bytes: bytes) -> bytes:
    """
    Extrae únicamente la primera página del PDF (original fiscal) según requerimiento de SSSalud.
    """
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if doc.page_count > 0:
            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=0, to_page=0)
            out_bytes = new_doc.tobytes(garbage=4, deflate=True)
            new_doc.close()
            doc.close()
            return out_bytes
        doc.close()
    except Exception:
        pass
    return pdf_bytes


def generar_zip_sssalud_adaptado(
    procesadas: List[Dict[str, Any]],
    rechazadas: List[Dict[str, Any]],
) -> bytes:
    """
    Replica exactamente `adaptar_pdf_sssalud` de SAPI Web (`backend/app/api/facturas.py`):
    Recorta cada PDF a 1 sola página y construye un archivo ZIP organizado en 3 carpetas:
    1. 'Facturas Procesadas/{CUIT}_{TIPO}_{PTO}_{NRO}.pdf'
    2. 'No Procesadas - CUIT No Corresponde a la Entidad/{archivo_original}.pdf'
    3. 'No Procesadas - Sin QR/{archivo_original}.pdf'
    """
    zip_buffer = io.BytesIO()
    nombres_usados: Dict[str, int] = {}

    def _ruta_unica(carpeta: str, nombre_base: str) -> str:
        nombre_limpio = re.sub(r'[<>:"/\\|?*]', "_", nombre_base.strip())
        if not nombre_limpio.lower().endswith(".pdf"):
            nombre_limpio += ".pdf"
        ruta = f"{carpeta}/{nombre_limpio}"
        contador = nombres_usados.get(ruta, 0)
        nombres_usados[ruta] = contador + 1
        if contador > 0:
            stem = nombre_limpio[:-4]
            ruta = f"{carpeta}/{stem}_{contador + 1}.pdf"
        return ruta

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        # 1. Facturas Procesadas con nombre oficial SSSalud
        for fac in procesadas:
            pdf_bytes = fac.get("pdf_bytes")
            if not pdf_bytes:
                continue
            pdf_1p = adaptar_pdf_a_primera_pagina(pdf_bytes)
            identificador = fac.get("registro_identificador") or generar_identificador_factura(
                cuit_prestador=fac.get("cuit_emisor"),
                tipo_comprobante=fac.get("tipo_comprobante"),
                punto_venta=fac.get("punto_venta"),
                numero_comprobante=fac.get("numero_comprobante"),
                tipo_cmp_afip=fac.get("tipo_comprobante_afip"),
            )
            ruta_zip = _ruta_unica("Facturas Procesadas", f"{identificador}.pdf")
            zip_file.writestr(ruta_zip, pdf_1p)

        # 2. Facturas Rechazadas clasificadas por causal
        for rech in rechazadas:
            pdf_bytes = rech.get("pdf_bytes")
            if not pdf_bytes:
                continue
            pdf_1p = adaptar_pdf_a_primera_pagina(pdf_bytes)
            tipo_err = rech.get("tipo_error", "sin_qr")
            if tipo_err == "cuit_invalido":
                carpeta = "No Procesadas - CUIT No Corresponde a la Entidad"
            elif tipo_err == "nota_credito":
                carpeta = "No Procesadas - Notas de Credito"
            else:
                carpeta = "No Procesadas - Sin QR"
            ruta_zip = _ruta_unica(carpeta, rech.get("archivo", "documento.pdf"))
            zip_file.writestr(ruta_zip, pdf_1p)

    zip_buffer.seek(0)
    return zip_buffer.getvalue()


# ==============================================================================
# 7. VALIDACIÓN DE FILAS EN PASO 3: REVISIÓN Y AUTORIZACIÓN
#    (Adaptado de `calcularEstadoFila` en CargaMasivaFacturas.tsx sin auditoría vs autorizaciones)
# ==============================================================================

def dni_coincide_con_cuit(dni: str, cuit: Optional[str]) -> bool:
    dni_clean = re.sub(r"\D", "", str(dni or ""))
    cuit_clean = re.sub(r"\D", "", str(cuit or ""))
    if not dni_clean or not cuit_clean:
        return False
    if dni_clean == cuit_clean:
        return True
    if len(cuit_clean) == 11:
        dni_del_cuit = cuit_clean[2:10]
        if dni_clean == dni_del_cuit or dni_clean.lstrip("0") == dni_del_cuit.lstrip("0"):
            return True
    return False


def calcular_estado_fila_revision(
    fila: Dict[str, Any],
    todas_las_filas: List[Dict[str, Any]],
) -> Tuple[str, str, List[str]]:
    """
    Calcula el estado ('Listo' o 'Error'), el motivo principal y la lista de advertencias no bloqueantes
    de una factura dentro del Paso 3 (Revisión y Autorización).

    Nota: Por requerimiento del usuario, NO realiza validación de autorización vs facturación
    (ya que no hay listado de autorizaciones previas en esta versión demo).
    """
    advertencias: List[str] = []

    # 1. No permitir Notas de Crédito
    if fila.get("es_nota_credito") or es_nota_credito_comprobante(
        fila.get("tipo_comprobante"),
        fila.get("tipo_comprobante_afip"),
    ):
        return ("Error", "Nota de Crédito no permitida en carga de facturas", advertencias)

    # 2. Verificar duplicado exacto de comprobante dentro del lote (CUIT + Punto Venta + Nro Comprobante)
    cuit_em = re.sub(r"\D", "", str(fila.get("cuit_emisor") or ""))
    pto = int(fila.get("punto_venta") or 0)
    nro = int(fila.get("numero_comprobante") or 0)
    fila_id = fila.get("id")

    for otra in todas_las_filas:
        if otra.get("id") == fila_id:
            continue
        if (
            re.sub(r"\D", "", str(otra.get("cuit_emisor") or "")) == cuit_em
            and int(otra.get("punto_venta") or 0) == pto
            and int(otra.get("numero_comprobante") or 0) == nro
        ):
            return ("Error", "Factura duplicada en este lote (mismo CUIT, Pto. Venta y Nro.)", advertencias)

    # 3. Validar DNI del Afiliado (7 u 8 dígitos numéricos, sin coincidir con CUIT emisor ni receptor)
    dni_limpio = re.sub(r"\D", "", str(fila.get("afiliado_dni") or ""))
    if not dni_limpio:
        return ("Error", "Falta ingresar el DNI del afiliado (7 u 8 dígitos)", advertencias)
    if len(dni_limpio) not in (7, 8) or not (5000000 <= int(dni_limpio) <= 99999999):
        return ("Error", "El DNI del afiliado debe tener 7 u 8 dígitos numéricos válidos", advertencias)
    if dni_coincide_con_cuit(dni_limpio, fila.get("cuit_emisor")):
        return ("Error", "El DNI ingresado pertenece al CUIT del Prestador emisor", advertencias)
    if fila.get("cuit_receptor") and dni_coincide_con_cuit(dni_limpio, fila.get("cuit_receptor")):
        return ("Error", "El DNI ingresado pertenece al CUIT Receptor de la institución", advertencias)

    # 4. Validar Período Prestacional (formato AAAAMM)
    periodo_limpio = str(fila.get("periodo") or "").strip()
    if not re.match(r"^20[2-3]\d(0[1-9]|1[0-2])$", periodo_limpio):
        return ("Error", "Período inválido (use formato AAAAMM, ej: 202603)", advertencias)

    # Verificar ventana cronograma SSSalud (informativo/advertencia para no bloquear pruebas con facturas históricas)
    cronograma = obtener_cronograma_presentacion()
    if periodo_limpio not in cronograma["periodos_habilitados"]:
        advertencias.append(
            f"Período {formatear_periodo_legible(periodo_limpio)} fuera del trimestre SSSalud actual "
            f"({', '.join( cronograma['periodos_habilitados'] )})."
        )

    # 5. Validar Codificación Nomenclador SSSalud / Sin Recupero
    sin_rec = bool(fila.get("sin_recupero", False))
    cod_nom = limpiar_codigo_nomenclador(fila.get("codigo_nomenclador"))
    catalogo = _cargar_nomenclador()

    if not sin_rec:
        if not cod_nom or cod_nom not in catalogo:
            return ("Error", "Seleccione un Código Nomenclador SSSalud (1-99) o marque Sin Recupero", advertencias)

    # 6. Validar Cantidad >= 1
    try:
        cant = int(fila.get("cantidad") or 0)
    except (TypeError, ValueError):
        cant = 0
    if cant < 1:
        return ("Error", "La cantidad debe ser un número entero mayor o igual a 1", advertencias)

    # 7. Validar duplicado de Afiliado + Prestador + Período en el lote (excepto Transporte código 96)
    es_transporte = (not sin_rec and cod_nom == "96") or ("transporte" in normalizar_texto(fila.get("prestacion")))
    if not es_transporte:
        for otra in todas_las_filas:
            if otra.get("id") == fila_id:
                continue
            otra_dni = re.sub(r"\D", "", str(otra.get("afiliado_dni") or ""))
            otra_cuit = re.sub(r"\D", "", str(otra.get("cuit_emisor") or ""))
            otra_per = str(otra.get("periodo") or "").strip()
            otra_cod = limpiar_codigo_nomenclador(otra.get("codigo_nomenclador"))
            otra_transporte = (not otra.get("sin_recupero") and otra_cod == "96") or (
                "transporte" in normalizar_texto(otra.get("prestacion"))
            )
            if not otra_transporte and otra_dni == dni_limpio and otra_cuit == cuit_em and otra_per == periodo_limpio:
                return (
                    "Error",
                    "Mismo Afiliado + Prestador + Período repetido en el lote (sólo permitido en Cód. 96 Transporte)",
                    advertencias,
                )

    # 8. Validar Importe > 0
    if float(fila.get("importe_total") or 0.0) <= 0:
        return ("Error", "El importe de la factura debe ser mayor a $0,00", advertencias)

    return ("Listo", "Listo para autorizar y registrar", advertencias)


def recalcular_estados_lote(filas: List[Dict[str, Any]]) -> None:
    """Actualiza in-place `estado`, `motivo_estado` y `advertencias` para cada factura del lote."""
    for fila in filas:
        # Sincronizar el identificador estructurado por si cambió algún dato
        fila["registro_identificador"] = generar_identificador_factura(
            cuit_prestador=fila.get("cuit_emisor"),
            tipo_comprobante=fila.get("tipo_comprobante"),
            punto_venta=fila.get("punto_venta"),
            numero_comprobante=fila.get("numero_comprobante"),
            tipo_cmp_afip=fila.get("tipo_comprobante_afip"),
        )
        estado, motivo, advs = calcular_estado_fila_revision(fila, filas)
        fila["estado"] = estado
        fila["motivo_estado"] = motivo
        fila["advertencias"] = advs
        if estado == "Error":
            fila["seleccionada"] = False


# ==============================================================================
# 8. EXPORTADORES (CSV, EXCEL .XLSX Y TXT SSSALUD)
# ==============================================================================

def generar_csv_revision(filas: List[Dict[str, Any]]) -> bytes:
    """
    Replica la exportación CSV de `handleExportarCsv` en `CargaMasivaFacturas.tsx` de SAPI Web
    con BOM UTF-8 para apertura directa en Microsoft Excel.
    """
    output = io.StringIO()
    output.write("\ufeff")  # BOM UTF-8
    writer = csv.writer(output, delimiter=";", quoting=csv.QUOTE_ALL)

    writer.writerow([
        "Archivo PDF",
        "Prestador",
        "CUIT Prestador",
        "CUIT Receptor",
        "Concepto / Identificador SSSalud",
        "Tipo Comprobante",
        "Comprobante",
        "CAE",
        "Fecha Emisión",
        "Importe ($)",
        "Período (AAAAMM)",
        "DNI Afiliado",
        "Nombre Afiliado",
        "Código Nomenclador",
        "Prestación",
        "Cantidad",
        "Sin Recupero",
        "Estado Validación",
        "Observación Validación",
        "Seleccionada",
    ])

    for f in filas:
        importe_str = f"{float(f.get('importe_total') or 0.0):.2f}".replace(".", ",")
        writer.writerow([
            f.get("archivo", ""),
            f.get("prestador_nombre", ""),
            f.get("cuit_emisor", ""),
            f.get("cuit_receptor", "") or "No especificado",
            f.get("registro_identificador", ""),
            f.get("tipo_comprobante_etiqueta", ""),
            f.get("comprobante_visible", ""),
            f.get("cae", ""),
            f.get("fecha_emision_ddmmaaaa", ""),
            importe_str,
            f.get("periodo", ""),
            f.get("afiliado_dni", ""),
            f.get("afiliado_nombre", "") or "S/D",
            "S/R" if f.get("sin_recupero") else (f.get("codigo_nomenclador") or ""),
            f.get("prestacion", ""),
            f.get("cantidad", 1),
            "SÍ" if f.get("sin_recupero") else "NO",
            f.get("estado", ""),
            f.get("motivo_estado", ""),
            "SÍ" if f.get("seleccionada") else "NO",
        ])

    return output.getvalue().encode("utf-8")


def generar_excel_revision(
    procesadas: List[Dict[str, Any]],
    rechazadas: List[Dict[str, Any]],
    cuit_institucional: Optional[str] = None,
) -> bytes:
    """
    Genera un libro Excel (.xlsx) profesional con 3 hojas:
    1. 'Revision_y_Autorizacion' (todas las facturas procesadas y su codificación SSSalud)
    2. 'Rechazadas_Informe' (detalle de facturas no procesadas y su motivo)
    3. 'Resumen_Lote' (métricas generales de la tanda)
    """
    buffer = io.BytesIO()

    filas_proc = []
    for idx, f in enumerate(procesadas, start=1):
        filas_proc.append({
            "#": idx,
            "Archivo PDF": f.get("archivo", ""),
            "Prestador (Razón Social)": f.get("prestador_nombre", ""),
            "CUIT Prestador": f.get("cuit_emisor", ""),
            "CUIT Receptor": f.get("cuit_receptor", "") or "S/D",
            "Identificador SSSalud": f.get("registro_identificador", ""),
            "Tipo Comprobante": f.get("tipo_comprobante_etiqueta", ""),
            "Cód. Comprobante SSSalud": f.get("codigo_comprobante_sssalud", 5),
            "Punto de Venta": f.get("punto_venta", 0),
            "Nro. Comprobante": f.get("numero_comprobante", 0),
            "CAE": f.get("cae", ""),
            "Fecha Emisión": f.get("fecha_emision_ddmmaaaa", ""),
            "Importe Total ($)": round(float(f.get("importe_total") or 0.0), 2),
            "Período (AAAAMM)": f.get("periodo", ""),
            "DNI Afiliado": f.get("afiliado_dni", ""),
            "Nombre Afiliado": f.get("afiliado_nombre", "") or "",
            "Código Nomenclador": "S/R" if f.get("sin_recupero") else (f.get("codigo_nomenclador") or ""),
            "Prestación": f.get("prestacion", ""),
            "Cantidad": int(f.get("cantidad") or 1),
            "Sin Recupero": "SÍ" if f.get("sin_recupero") else "NO",
            "Dependencia": "SÍ" if f.get("dependencia") else "NO",
            "Estado": f.get("estado", ""),
            "Detalle Estado": f.get("motivo_estado", ""),
            "Autorizada / Seleccionada": "SÍ" if f.get("seleccionada") else "NO",
        })

    filas_rech = []
    for idx, r in enumerate(rechazadas, start=1):
        filas_rech.append({
            "#": idx,
            "Archivo PDF": r.get("archivo", ""),
            "Categoría Rechazo": {
                "nota_credito": "Nota de Crédito",
                "cuit_invalido": "CUIT No Corresponde a la Entidad",
                "sin_qr": "Sin QR / Formato Inválido",
            }.get(r.get("tipo_error", "sin_qr"), "Rechazada"),
            "Motivo Detallado": r.get("motivo", ""),
            "CUIT Emisor Detectado": r.get("cuit_emisor", "") or "S/D",
            "CUIT Receptor Detectado": r.get("cuit_receptor", "") or "S/D",
        })

    seleccionadas = [f for f in procesadas if f.get("seleccionada") and f.get("estado") == "Listo"]
    importe_total_lote = sum(float(f.get("importe_total") or 0.0) for f in procesadas)
    importe_autorizado = sum(float(f.get("importe_total") or 0.0) for f in seleccionadas)

    resumen_rows = [
        {"Indicador": "Fecha y Hora de Procesamiento", "Valor": datetime.now().strftime("%d/%m/%Y %H:%M:%S")},
        {"Indicador": "CUIT Institucional Verificado", "Valor": formatear_cuit_guiones(cuit_institucional) if cuit_institucional else "No especificado (Opcional)"},
        {"Indicador": "Total Facturas Procesadas con Éxito", "Valor": len(procesadas)},
        {"Indicador": "Total Facturas No Procesadas / Rechazadas", "Valor": len(rechazadas)},
        {"Indicador": "Total Facturas Listas y Seleccionadas", "Valor": len(seleccionadas)},
        {"Indicador": "Importe Total Procesado ($)", "Valor": round(importe_total_lote, 2)},
        {"Indicador": "Importe Total Autorizado ($)", "Valor": round(importe_autorizado, 2)},
    ]

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_proc = pd.DataFrame(filas_proc) if filas_proc else pd.DataFrame(columns=["Sin facturas procesadas"])
        df_proc.to_excel(writer, index=False, sheet_name="Revision_y_Autorizacion")

        df_rech = pd.DataFrame(filas_rech) if filas_rech else pd.DataFrame(columns=["Sin facturas rechazadas"])
        df_rech.to_excel(writer, index=False, sheet_name="Rechazadas_Informe")

        df_res = pd.DataFrame(resumen_rows)
        df_res.to_excel(writer, index=False, sheet_name="Resumen_Lote")

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_len = max((len(str(cell.value or "")) for cell in col), default=12)
                col_letter = col[0].column_letter
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 48)

    buffer.seek(0)
    return buffer.getvalue()


def generar_txt_sssalud(filas: List[Dict[str, Any]]) -> bytes:
    """
    Exporta las facturas seleccionadas y listas (con recupero) al formato TXT oficial de SSSalud
    (replicando `exportar_facturas_txt` de `backend/app/services/factura_service.py`).
    """
    lineas: List[str] = []
    for f in filas:
        if not f.get("seleccionada") or f.get("estado") != "Listo":
            continue
        if f.get("sin_recupero"):
            continue
        cuit_prest = re.sub(r"\D", "", str(f.get("cuit_emisor") or "")).zfill(11)
        cod_cmp = str(f.get("codigo_comprobante_sssalud") or 5)
        pto_vta = str(int(f.get("punto_venta") or 0)).zfill(4)
        nro_cmp = str(int(f.get("numero_comprobante") or 0)).zfill(8)
        fecha_em = f.get("fecha_emision_ddmmaaaa", "")
        cae = re.sub(r"\D", "", str(f.get("cae") or "")).zfill(14)
        periodo = str(f.get("periodo") or "")
        dni = re.sub(r"\D", "", str(f.get("afiliado_dni") or "")).zfill(8)
        cod_nom = str(f.get("codigo_nomenclador") or "86").zfill(2)
        cant = str(int(f.get("cantidad") or 1))
        importe_fmt = f"{float(f.get('importe_total') or 0.0):.2f}".replace(".", ",")
        dep = "1" if f.get("dependencia") else "0"

        linea = "|".join([
            f"{cuit_prest}_{cod_cmp}_{int(pto_vta)}_{int(nro_cmp)}",
            cuit_prest,
            cod_cmp,
            pto_vta,
            nro_cmp,
            fecha_em,
            cae,
            periodo,
            dni,
            cod_nom,
            cant,
            importe_fmt,
            dep,
        ])
        lineas.append(linea)

    contenido = "\r\n".join(lineas) + ("\r\n" if lineas else "")
    return contenido.encode("utf-8")


def renderizar_preview_pdf_png(pdf_bytes: bytes, zoom: float = 1.5) -> Optional[bytes]:
    """Renderiza la primera página de un PDF a PNG en memoria para previsualizar en el Paso 3."""
    if not pdf_bytes:
        return None
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if doc.page_count == 0:
            doc.close()
            return None
        page = doc[0]
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        png_bytes = pix.tobytes("png")
        doc.close()
        return png_bytes
    except Exception:
        return None
