import html
import streamlit as st
import pandas as pd
from datetime import datetime
from typing import Any, Dict, List

from sapi_facturas_core import (
    MAX_PDFS_POR_TANDA_DEMO,
    CODIGOS_CON_CANTIDAD_CUATRO,
    CODIGOS_DEPENDENCIA_VALIDOS,
    escanear_lote_pdfs,
    formatear_cuit_guiones,
    generar_excel_revision,
    generar_txt_sssalud,
    generar_zip_sssalud_adaptado,
    obtener_opciones_selector_prestacion,
    recalcular_estados_lote,
    validar_formato_cuit_opcional,
)

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA Y ESTILOS VISUALES (ESTÉTICA SAPI WEB)
# ==============================================================================
st.set_page_config(
    page_title="SAPI Web · Carga Masiva de Facturas Digitales (QR)",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3rem;
        max-width: 1120px;
    }

    /* Cabecera principal estilo SAPI Web (100% ancho y simétrica) */
    .sapi-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 60%, #1d4ed8 100%);
        border-radius: 16px;
        padding: 1.5rem 1.75rem;
        color: #ffffff;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.18);
        margin-bottom: 1.15rem;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        text-align: center;
        gap: 0.65rem;
        width: 100%;
        box-sizing: border-box;
    }
    .sapi-header-title {
        font-size: 1.6rem;
        font-weight: 800;
        letter-spacing: -0.02em;
        margin: 0;
        color: #ffffff;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 0.6rem;
    }
    .sapi-header-sub {
        font-size: 0.92rem;
        color: #cbd5e1;
        margin: 0.2rem 0 0 0;
        text-align: center;
    }
    .sapi-badge-demo {
        background: rgba(255, 255, 255, 0.14);
        border: 1px solid rgba(255, 255, 255, 0.28);
        color: #f8fafc;
        padding: 0.4rem 1rem;
        border-radius: 999px;
        font-size: 0.78rem;
        font-weight: 700;
        letter-spacing: 0.03em;
        backdrop-filter: blur(6px);
        display: inline-block;
    }

    /* Stepper de 3 pasos (100% ancho y tarjetas de igual alto y simetría) */
    .sapi-stepper {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 0.9rem;
        margin-top: 0.4rem;
        margin-bottom: 1.15rem;
        width: 100%;
        box-sizing: border-box;
    }
    .sapi-step {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 0.95rem 1.15rem;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 0.85rem;
        min-height: 82px;
        height: 100%;
        box-sizing: border-box;
        transition: all 0.2s ease;
    }
    .sapi-step.active {
        background: #eff6ff;
        border: 2px solid #2563eb;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.12);
    }
    .sapi-step.done {
        background: #f0fdf4;
        border: 1px solid #86efac;
    }
    .sapi-step-num {
        width: 32px;
        height: 32px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 800;
        font-size: 0.88rem;
        flex-shrink: 0;
        background: #f1f5f9;
        color: #64748b;
    }
    .sapi-step.active .sapi-step-num {
        background: #2563eb;
        color: #ffffff;
    }
    .sapi-step.done .sapi-step-num {
        background: #16a34a;
        color: #ffffff;
    }
    .sapi-step-label {
        font-weight: 700;
        font-size: 0.92rem;
        color: #334155;
        margin: 0 0 0.15rem 0;
    }
    .sapi-step.active .sapi-step-label {
        color: #1e40af;
    }
    .sapi-step.done .sapi-step-label {
        color: #166534;
    }
    .sapi-step-desc {
        font-size: 0.78rem;
        color: #64748b;
        margin: 0;
        line-height: 1.3;
    }

    /* Tarjetas contenedoras */
    .sapi-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 1.25rem 1.4rem;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
        margin-bottom: 1rem;
    }
    .sapi-card-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #0f172a;
        margin: 0 0 0.35rem 0;
    }
    .sapi-card-sub {
        font-size: 0.85rem;
        color: #64748b;
        margin: 0 0 0.9rem 0;
    }

    /* Tarjetas KPI Informe de Validación */
    .kpi-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 1rem;
        margin-bottom: 1.25rem;
    }
    .kpi-box-ok {
        background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%);
        border: 1px solid #86efac;
        border-radius: 14px;
        padding: 1.15rem 1.4rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .kpi-box-err {
        background: linear-gradient(135deg, #fef2f2 0%, #fee2e2 100%);
        border: 1px solid #fca5a5;
        border-radius: 14px;
        padding: 1.15rem 1.4rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .kpi-num-ok {
        font-size: 2.2rem;
        font-weight: 800;
        color: #15803d;
        line-height: 1;
    }
    .kpi-num-err {
        font-size: 2.2rem;
        font-weight: 800;
        color: #b91c1c;
        line-height: 1;
    }

    /* Badges de estado */
    .badge-ok {
        display: inline-flex;
        align-items: center;
        gap: 0.35rem;
        padding: 0.28rem 0.7rem;
        border-radius: 999px;
        font-size: 0.78rem;
        font-weight: 700;
        background: #dcfce7;
        color: #15803d;
        border: 1px solid #86efac;
    }
    .badge-err {
        display: inline-flex;
        align-items: center;
        gap: 0.35rem;
        padding: 0.28rem 0.7rem;
        border-radius: 999px;
        font-size: 0.78rem;
        font-weight: 700;
        background: #fee2e2;
        color: #b91c1c;
        border: 1px solid #fca5a5;
    }
    .badge-warn {
        display: inline-flex;
        align-items: center;
        gap: 0.35rem;
        padding: 0.28rem 0.7rem;
        border-radius: 999px;
        font-size: 0.76rem;
        font-weight: 600;
        background: #fef3c7;
        color: #92400e;
        border: 1px solid #fde68a;
    }
    .badge-info {
        display: inline-flex;
        align-items: center;
        gap: 0.35rem;
        padding: 0.25rem 0.65rem;
        border-radius: 6px;
        font-size: 0.78rem;
        font-weight: 700;
        background: #eff6ff;
        color: #1d4ed8;
        border: 1px solid #bfdbfe;
        font-family: monospace;
    }

    /* Centrado y unificación del cuadro de carga en Paso 1 */
    [data-testid="stFileUploader"] {
        width: 100% !important;
    }
    [data-testid="stFileUploaderDropzone"] {
        display: flex !important;
        flex-direction: column !important;
        align-items: center !important;
        justify-content: center !important;
        text-align: center !important;
        height: auto !important;
        min-height: 155px !important;
        padding: 1.5rem 1.5rem !important;
        gap: 0.45rem !important;
        background: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 14px !important;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04) !important;
        margin: 0 auto !important;
        cursor: pointer !important;
        transition: border-color 0.2s ease, background-color 0.2s ease, box-shadow 0.2s ease !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        border-color: #2563eb !important;
        background: #f8fafc !important;
        box-shadow: 0 4px 14px rgba(37, 99, 235, 0.08) !important;
    }
    [data-testid="stFileUploaderDropzone"] > span,
    [data-testid="stFileUploaderDropzone"] button {
        display: none !important;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] {
        display: flex !important;
        flex-direction: column !important;
        align-items: center !important;
        justify-content: center !important;
        text-align: center !important;
        margin-right: 0 !important;
        width: 100% !important;
        gap: 0.4rem !important;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] > span,
    [data-testid="stFileUploaderDropzoneInstructions"] > div {
        display: none !important;
    }
    [data-testid="stFileUploaderDropzoneInstructions"]::before {
        content: "📂 Selección de Facturas Digitales en PDF";
        display: block;
        font-size: 1.05rem;
        font-weight: 700;
        color: #0f172a;
        line-height: 1.35;
        text-align: center;
    }
    [data-testid="stFileUploaderDropzoneInstructions"]::after {
        content: "Arrastre y suelte los archivos aquí o haga clic para seleccionar las facturas electrónicas AFIP/ARCA.\\ALímite de versión Demo: hasta __MAX_PDFS__ archivos PDF por tanda (máx. 15 MB por archivo).";
        display: block;
        white-space: pre-line;
        font-size: 0.85rem;
        color: #64748b;
        line-height: 1.5;
        text-align: center;
    }
    [data-testid="stFileUploader"] label,
    [data-testid="stTextInput"] label {
        display: flex !important;
        justify-content: center !important;
        width: 100% !important;
        text-align: center !important;
        font-weight: 600 !important;
    }
    [data-testid="stTextInput"] input {
        text-align: center !important;
    }

    /* Notificaciones emergentes (toast) en esquina superior derecha */
    [data-testid="stToastContainer"] {
        z-index: 999999 !important;
    }
    [data-testid="stToast"] {
        background-color: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 12px !important;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.14) !important;
    }

    /* Modal emergente centrado para progreso de escaneo */
    .sapi-scan-overlay {
        position: fixed;
        inset: 0;
        background: rgba(15, 23, 42, 0.45);
        backdrop-filter: blur(4px);
        z-index: 999998;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 1rem;
        box-sizing: border-box;
    }
    .sapi-scan-modal {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 1.75rem 2rem;
        width: 100%;
        max-width: 480px;
        box-shadow: 0 20px 45px -10px rgba(15, 23, 42, 0.25);
        text-align: center;
        box-sizing: border-box;
    }
    .sapi-scan-icon {
        font-size: 2rem;
        line-height: 1;
        margin-bottom: 0.6rem;
    }
    .sapi-scan-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #0f172a;
        margin: 0 0 0.35rem 0;
    }
    .sapi-scan-sub {
        font-size: 0.88rem;
        color: #475569;
        margin: 0;
        word-break: break-word;
        line-height: 1.4;
    }
    .sapi-progress-track {
        width: 100%;
        height: 10px;
        background: #e2e8f0;
        border-radius: 999px;
        overflow: hidden;
        margin: 1rem 0 0.5rem 0;
    }
    .sapi-progress-fill {
        height: 100%;
        background: linear-gradient(90deg, #1d4ed8 0%, #3b82f6 100%);
        border-radius: 999px;
        transition: width 0.25s ease;
    }
    .sapi-scan-pct {
        font-size: 0.8rem;
        font-weight: 700;
        color: #1d4ed8;
        margin: 0;
    }

    /* Tarjetas métricas y barra compacta superior en Paso 3 */
    .rev-kpi-bar {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 0.7rem 1.25rem;
        margin-bottom: 0.85rem;
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        align-items: center;
        box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
    }
    .rev-kpi-item {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 0.15rem 1.25rem;
    }
    .rev-kpi-item:first-child {
        padding-left: 0.25rem;
    }
    .rev-kpi-item:last-child {
        padding-right: 0.25rem;
    }
    .rev-kpi-item + .rev-kpi-item {
        border-left: 1px solid #e2e8f0;
    }
    .rev-kpi-label {
        font-size: 0.84rem;
        font-weight: 600;
        color: #475569;
        margin: 0;
    }
    .rev-kpi-val {
        font-size: 1.4rem;
        font-weight: 800;
        line-height: 1;
        margin: 0;
    }
    .st-key-btn_sel_all_paso3 button,
    .st-key-btn_desel_all_paso3 button {
        font-size: 0.78rem !important;
        padding: 0.25rem 0.65rem !important;
        min-height: 32px !important;
        height: 32px !important;
        border-radius: 8px !important;
        white-space: nowrap !important;
    }
    .st-key-btn_excel_icon button,
    .st-key-btn_info_icon button {
        width: 28px !important;
        min-width: 28px !important;
        height: 28px !important;
        min-height: 28px !important;
        padding: 0 !important;
        border: none !important;
        background-color: transparent !important;
        box-shadow: none !important;
        background-repeat: no-repeat !important;
        background-position: center !important;
        background-size: 22px 22px !important;
        transition: transform 0.15s ease, opacity 0.15s ease !important;
    }
    .st-key-btn_excel_icon button {
        background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24'%3E%3Cpath fill='%23185C37' d='M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z'/%3E%3Cpath fill='%2321A366' d='M14 2v6h6z'/%3E%3Cpath fill='%23107C41' d='M14 8h6v12a2 2 0 0 1-2 2h-4z'/%3E%3Crect x='2' y='6.5' width='11' height='11' rx='2' fill='%23107C41'/%3E%3Cpath fill='%23FFFFFF' d='M5.1 9.2h1.55l1.15 2.05 1.15-2.05h1.5l-1.8 2.8 1.9 2.9H8.95l-1.2-2.1-1.2 2.1H5.0l1.9-2.9z'/%3E%3C/svg%3E") !important;
    }
    .st-key-btn_info_icon button {
        background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24'%3E%3Ccircle cx='12' cy='12' r='10' fill='%232563EB'/%3E%3Crect x='11' y='10.5' width='2' height='6.5' rx='1' fill='%23FFFFFF'/%3E%3Ccircle cx='12' cy='7.5' r='1.2' fill='%23FFFFFF'/%3E%3C/svg%3E") !important;
    }
    .st-key-btn_excel_icon button:hover,
    .st-key-btn_excel_icon button:focus,
    .st-key-btn_excel_icon button:active,
    .st-key-btn_info_icon button:hover,
    .st-key-btn_info_icon button:focus,
    .st-key-btn_info_icon button:active {
        border: none !important;
        background-color: transparent !important;
        box-shadow: none !important;
        transform: scale(1.1);
    }
    .st-key-btn_excel_icon button > div,
    .st-key-btn_info_icon button > div {
        display: none !important;
    }
</style>
""".replace("__MAX_PDFS__", str(MAX_PDFS_POR_TANDA_DEMO)),
    unsafe_allow_html=True,
)


# ==============================================================================
# ESTADO DE SESIÓN (SESSION STATE)
# ==============================================================================
def inicializar_estado() -> None:
    defaults: Dict[str, Any] = {
        "paso_actual": "seleccion",  # 'seleccion' | 'informe' | 'revision'
        "cuit_institucional_input": "",
        "cuit_institucional_validado": "",
        "ultimo_cuit_notificado": "",
        "archivos_seleccionados": [],
        "toast_limite_pendiente": None,
        "toast_informe_pendiente": False,
        "facturas_procesadas": [],
        "facturas_rechazadas": [],
        "descargar_formato_sssalud": False,
        "zip_sssalud_bytes": None,
        "lote_confirmado": False,
        "facturas_registradas_final": [],
        "uploader_key": 0,
        "revision_editor_version": 0,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


inicializar_estado()


def reiniciar_proceso_completo() -> None:
    st.session_state["paso_actual"] = "seleccion"
    st.session_state["archivos_seleccionados"] = []
    st.session_state["toast_limite_pendiente"] = None
    st.session_state["toast_informe_pendiente"] = False
    st.session_state["facturas_procesadas"] = []
    st.session_state["facturas_rechazadas"] = []
    st.session_state["zip_sssalud_bytes"] = None
    st.session_state["lote_confirmado"] = False
    st.session_state["facturas_registradas_final"] = []
    st.session_state["uploader_key"] += 1
    st.session_state["revision_editor_version"] += 1


def formatear_moneda_ars(valor: float) -> str:
    entero, dec = f"{float(valor or 0.0):,.2f}".split(".")
    entero_ar = entero.replace(",", ".")
    return f"$ {entero_ar},{dec}"


# ==============================================================================
# CABECERA PRINCIPAL, CUIT INSTITUCIONAL Y STEPPER DE 3 PASOS
# ==============================================================================
def render_cabecera_y_stepper() -> None:
    st.markdown(
        f"""
        <div class="sapi-header">
            <div>
                <h1 class="sapi-header-title">🧾 Carga Masiva de Facturas Digitales (QR AFIP/ARCA)</h1>
                <p class="sapi-header-sub">
                    Módulo Contable SAPI Web · Decodificación fiscal QR, adaptación SSSalud (1 página) y codificación Nomenclador de Discapacidad
                </p>
            </div>
            <div class="sapi-badge-demo">
                ⚡ VERSIÓN DEMO · MÁX. {MAX_PDFS_POR_TANDA_DEMO} PDFs POR TANDA
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    paso = st.session_state["paso_actual"]

    # Campo de CUIT del mismo ancho completo (100%), ubicado entre el título y los 3 pasos
    cuit_raw = st.text_input(
        "🏛️ CUIT Institucional (Opcional)",
        value=st.session_state["cuit_institucional_input"],
        placeholder="Ej: 30-12345678-9 (dejar vacío para no verificar CUIT receptor)",
        disabled=(paso != "seleccion"),
    )
    st.session_state["cuit_institucional_input"] = cuit_raw
    cuit_limpio_input = cuit_raw.strip()
    cuit_valido, _, msg_cuit = validar_formato_cuit_opcional(cuit_raw)
    if cuit_limpio_input != st.session_state.get("ultimo_cuit_notificado", ""):
        st.session_state["ultimo_cuit_notificado"] = cuit_limpio_input
        if cuit_limpio_input:
            if cuit_valido:
                st.toast(msg_cuit, icon="✅")
            else:
                st.toast(msg_cuit, icon="⚠️", duration="long")

    clase_1 = "active" if paso == "seleccion" else "done"
    clase_2 = "active" if paso == "informe" else ("done" if paso == "revision" else "")
    clase_3 = "active" if paso == "revision" else ""

    icono_1 = "✓" if clase_1 == "done" else "1"
    icono_2 = "✓" if clase_2 == "done" else "2"
    icono_3 = "3"

    st.markdown(
        f"""
        <div class="sapi-stepper">
            <div class="sapi-step {clase_1}">
                <div class="sapi-step-num">{icono_1}</div>
                <div>
                    <p class="sapi-step-label">1. Selección de PDFs</p>
                    <p class="sapi-step-desc">Carga de hasta {MAX_PDFS_POR_TANDA_DEMO} facturas PDF por tanda</p>
                </div>
            </div>
            <div class="sapi-step {clase_2}">
                <div class="sapi-step-num">{icono_2}</div>
                <div>
                    <p class="sapi-step-label">2. Informe de Validación</p>
                    <p class="sapi-step-desc">Escaneo QR y adaptación SSSalud (1 página)</p>
                </div>
            </div>
            <div class="sapi-step {clase_3}">
                <div class="sapi-step-num">{icono_3}</div>
                <div>
                    <p class="sapi-step-label">3. Revisión y Autorización</p>
                    <p class="sapi-step-desc">Codificación SSSalud, DNI, Período y exportación</p>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ==============================================================================
# PASO 1: SELECCIÓN DE PDFs (MISMO ANCHO COMPLETO Y SIMÉTRICO)
# ==============================================================================
def render_paso_1_seleccion() -> None:
    cuit_raw = st.session_state.get("cuit_institucional_input", "")
    cuit_valido, cuit_limpio, _ = validar_formato_cuit_opcional(cuit_raw)

    toast_limite = st.session_state.pop("toast_limite_pendiente", None)
    if toast_limite:
        st.toast(toast_limite, icon="⚠️", duration="long")

    modal_progreso = st.empty()
    archivos_guardados: List[tuple[str, bytes]] = st.session_state.get("archivos_seleccionados", [])

    if not archivos_guardados:
        uploaded_files = st.file_uploader(
            f"Seleccione hasta {MAX_PDFS_POR_TANDA_DEMO} archivos PDF",
            type=["pdf"],
            accept_multiple_files=True,
            key=f"pdf_uploader_{st.session_state['uploader_key']}",
        )

        if uploaded_files:
            total_subidos = len(uploaded_files)
            archivos_activos = uploaded_files[:MAX_PDFS_POR_TANDA_DEMO]
            st.session_state["archivos_seleccionados"] = [
                (f.name, f.getvalue()) for f in archivos_activos
            ]
            if total_subidos > MAX_PDFS_POR_TANDA_DEMO:
                st.session_state["toast_limite_pendiente"] = (
                    f"Límite de Versión Demo: se aceptan hasta {MAX_PDFS_POR_TANDA_DEMO} archivos. "
                    f"Se tomaron los primeros {MAX_PDFS_POR_TANDA_DEMO} de {total_subidos}."
                )
            st.session_state["uploader_key"] += 1
            st.rerun()
        else:
            st.markdown(
                f"""
                <div style="width: 100%; box-sizing: border-box; text-align: center; background: #eff6ff; border: 1px solid #bfdbfe; color: #1e40af; padding: 0.9rem 1rem; border-radius: 12px; font-size: 0.9rem; margin-top: 0.5rem;">
                    👆 Arrastre o seleccione entre <b>1 y {MAX_PDFS_POR_TANDA_DEMO} archivos PDF</b> para iniciar el escaneo QR y validación.
                </div>
                """,
                unsafe_allow_html=True,
            )
    else:
        # Tabla resumen de archivos seleccionados (ocupa el espacio completo del bloque de carga)
        df_files = pd.DataFrame([
            {
                "#": idx + 1,
                "Archivo PDF": nombre,
                "Tamaño (KB)": f"{len(contenido) / 1024:.1f} KB",
                "Estado en Tanda": f"Incluido en la tanda (1-{MAX_PDFS_POR_TANDA_DEMO})",
            }
            for idx, (nombre, contenido) in enumerate(archivos_guardados)
        ])
        st.dataframe(df_files, use_container_width=True, hide_index=True, height=270)

        col_btn_clear, col_btn_scan = st.columns([1, 2])
        with col_btn_clear:
            if st.button("🗑️ Vaciar lista de PDFs", use_container_width=True):
                st.session_state["archivos_seleccionados"] = []
                st.session_state["toast_limite_pendiente"] = None
                st.session_state["uploader_key"] += 1
                st.rerun()

        with col_btn_scan:
            boton_deshabilitado = not cuit_valido
            if st.button(
                f"🔍 Escanear y Validar {len(archivos_guardados)} Factura(s) Digital(es) →",
                type="primary",
                use_container_width=True,
                disabled=boton_deshabilitado,
                help=(
                    "Corrija el CUIT Institucional (11 dígitos) o déjelo vacío para continuar."
                    if boton_deshabilitado
                    else None
                ),
            ):
                def _render_modal_progreso(texto: str, porc: int) -> None:
                    modal_progreso.markdown(
                        f"""
<div class="sapi-scan-overlay">
    <div class="sapi-scan-modal">
        <div class="sapi-scan-icon">🔍</div>
        <p class="sapi-scan-title">Escaneando Facturas Digitales (QR)</p>
        <p class="sapi-scan-sub">{html.escape(texto)}</p>
        <div class="sapi-progress-track">
            <div class="sapi-progress-fill" style="width: {porc}%;"></div>
        </div>
        <p class="sapi-scan-pct">{porc}% completado</p>
    </div>
</div>
""",
                        unsafe_allow_html=True,
                    )

                _render_modal_progreso("Iniciando lectura multicapa de códigos QR AFIP/ARCA...", 0)
                lote_tuplas = list(archivos_guardados)

                def _actualizar_progreso(actual: int, total: int, nombre: str) -> None:
                    porc = int((actual / max(total, 1)) * 100)
                    _render_modal_progreso(f"Escaneando ({actual}/{total}): {nombre}...", porc)

                resultado_lote = escanear_lote_pdfs(
                    archivos_pdf=lote_tuplas,
                    cuit_institucional=cuit_limpio if cuit_valido else "",
                    progreso_callback=_actualizar_progreso,
                )
                modal_progreso.empty()

                st.session_state["cuit_institucional_validado"] = cuit_limpio if cuit_valido else ""
                st.session_state["facturas_procesadas"] = resultado_lote["procesadas"]
                st.session_state["facturas_rechazadas"] = resultado_lote["rechazadas"]
                st.session_state["zip_sssalud_bytes"] = generar_zip_sssalud_adaptado(
                    resultado_lote["procesadas"],
                    resultado_lote["rechazadas"],
                )
                st.session_state["lote_confirmado"] = False
                st.session_state["toast_informe_pendiente"] = True
                st.session_state["paso_actual"] = "informe"
                st.rerun()


# ==============================================================================
# PASO 2: INFORME DE VALIDACIÓN
# ==============================================================================
def _indice_original_lote(registro_id: Any) -> int:
    partes = str(registro_id or "").split("_")
    if len(partes) >= 2 and partes[1].isdigit():
        return int(partes[1])
    return 9999


def render_paso_2_informe() -> None:
    procesadas: List[Dict[str, Any]] = st.session_state["facturas_procesadas"]
    rechazadas: List[Dict[str, Any]] = st.session_state["facturas_rechazadas"]
    cuit_inst = st.session_state.get("cuit_institucional_validado", "")

    if st.session_state.pop("toast_informe_pendiente", False):
        st.toast(
            f"Facturas Procesadas con Éxito: {len(procesadas)} · QR AFIP/ARCA decodificado",
            icon="✅",
            duration="long",
        )
        msg_rech = f"Facturas No Procesadas / Rechazadas: {len(rechazadas)}"
        if cuit_inst:
            msg_rech += f" · CUIT {formatear_cuit_guiones(cuit_inst)}"
        st.toast(msg_rech, icon="❌", duration="long")

    etiquetas_err = {
        "nota_credito": "🚫 Nota de Crédito",
        "cuit_invalido": "🏛️ Entidad No Coincide",
        "sin_qr": "🔍 Sin QR / Ilegible",
    }

    registros_unificados: List[tuple[int, Dict[str, Any]]] = []
    for f in procesadas:
        pto_vta = int(f.get("punto_venta") or 0)
        nro_comp = int(f.get("numero_comprobante") or 0)
        registros_unificados.append((
            _indice_original_lote(f.get("id")),
            {
                "Archivo": f["archivo"],
                "Estado": "Procesado",
                "Prestador": f"{f['prestador_nombre']} ({formatear_cuit_guiones(f['cuit_emisor'])})",
                "Comp": f"{pto_vta}-{nro_comp}",
                "Tipo": f["tipo_comprobante_etiqueta"],
                "Importe": formatear_moneda_ars(f["importe_total"]),
            },
        ))

    for r in rechazadas:
        cuit_em = r.get("cuit_emisor")
        registros_unificados.append((
            _indice_original_lote(r.get("id")),
            {
                "Archivo": r["archivo"],
                "Estado": "No procesado",
                "Prestador": formatear_cuit_guiones(cuit_em) if cuit_em else "—",
                "Comp": "—",
                "Tipo": etiquetas_err.get(r.get("tipo_error", "sin_qr"), "Error"),
                "Importe": "—",
            },
        ))

    registros_unificados.sort(key=lambda item: item[0])
    filas_tabla = [fila for _, fila in registros_unificados]

    if filas_tabla:
        df_informe = pd.DataFrame(filas_tabla)
        alto_tabla = min(400, max(220, 42 + 35 * len(filas_tabla)))
        st.dataframe(df_informe, use_container_width=True, hide_index=True, height=alto_tabla)
    else:
        st.warning("No se encontraron comprobantes para informar en este lote.")

    zip_bytes = st.session_state.get("zip_sssalud_bytes")
    if not zip_bytes:
        zip_bytes = generar_zip_sssalud_adaptado(procesadas, rechazadas)
        st.session_state["zip_sssalud_bytes"] = zip_bytes

    st.markdown("---")
    col_back, col_zip, col_next = st.columns([1.0, 1.15, 1.25], gap="medium")
    with col_back:
        if st.button("← Seleccionar otros archivos", use_container_width=True):
            reiniciar_proceso_completo()
            st.rerun()

    with col_zip:
        st.download_button(
            label="📦 Descargar ZIP SSSalud (1 pág.)",
            data=zip_bytes,
            file_name=f"Facturas_SSSalud_Adaptadas_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
            mime="application/zip",
            use_container_width=True,
            type="secondary",
            help=(
                "Adapta cada PDF dejando únicamente la primera página y los clasifica en 3 carpetas SSSalud."
            ),
        )

    with col_next:
        if st.button(
            f"Confirmar y Procesar Facturas Aceptadas ({len(procesadas)}) →",
            type="primary",
            use_container_width=True,
            disabled=(len(procesadas) == 0),
        ):
            st.session_state["paso_actual"] = "revision"
            st.rerun()


# ==============================================================================
# PASO 3: REVISIÓN Y AUTORIZACIÓN (SIN AUDITORÍA VS AUTORIZACIONES)
# ==============================================================================
def render_paso_3_revision() -> None:
    procesadas: List[Dict[str, Any]] = st.session_state["facturas_procesadas"]
    rechazadas: List[Dict[str, Any]] = st.session_state["facturas_rechazadas"]
    cuit_inst = st.session_state.get("cuit_institucional_validado", "")

    if not procesadas:
        st.warning("No hay facturas procesadas para revisar.")
        if st.button("← Volver a Selección de PDFs"):
            reiniciar_proceso_completo()
            st.rerun()
        return

    recalcular_estados_lote(procesadas)

    # Catálogo de opciones del Nomenclador SSSalud (1..99 + Apoyo + Sin Recupero)
    opciones_prestacion = obtener_opciones_selector_prestacion()
    labels_opciones = [op["label"] for op in opciones_prestacion]
    mapa_label_a_op = {op["label"]: op for op in opciones_prestacion}

    def _obtener_label_actual(fila: Dict[str, Any]) -> str:
        sin_rec = bool(fila.get("sin_recupero", False))
        cod = str(fila.get("codigo_nomenclador") or "").strip()
        desc = str(fila.get("prestacion") or "").strip()
        if sin_rec:
            for op in opciones_prestacion:
                if op["sin_recupero"] and op["descripcion"].lower() == desc.lower():
                    return op["label"]
            for op in opciones_prestacion:
                if op["sin_recupero"]:
                    return op["label"]
        else:
            for op in opciones_prestacion:
                if not op["sin_recupero"] and op["codigo"] == cod and op["descripcion"].lower() == desc.lower():
                    return op["label"]
            for op in opciones_prestacion:
                if not op["sin_recupero"] and op["codigo"] == cod:
                    return op["label"]
        return labels_opciones[0]

    # Espacio superior para métricas (Total en Revisión, Listas, Con Observaciones) y barra de acciones
    metrics_ph = st.empty()
    toolbar_ph = st.empty()

    # Construcción de la tabla interactiva (filas y columnas) para Revisión y Autorización
    def _texto_observacion(fila: Dict[str, Any]) -> str:
        if fila.get("estado") != "Listo":
            return str(fila.get("motivo_estado") or "Error")
        advs = fila.get("advertencias") or []
        if advs:
            return "⚠️ " + " | ".join(str(a) for a in advs)
        return "Listo para autorizar"

    filas_editor = []
    snapshots_previos = []
    for fila in procesadas:
        es_listo = fila.get("estado") == "Listo"
        pto_vta = int(fila.get("punto_venta") or 0)
        nro_comp = int(fila.get("numero_comprobante") or 0)
        label_actual = _obtener_label_actual(fila)
        obs_actual = _texto_observacion(fila)

        filas_editor.append({
            "Aut.": bool(fila.get("seleccionada", False)) and es_listo,
            "Estado": "✓ Listo" if es_listo else "✕ Error",
            "Archivo": fila["archivo"],
            "Comp": f"{pto_vta}-{nro_comp}",
            "Tipo": fila["tipo_comprobante_etiqueta"],
            "Importe": formatear_moneda_ars(fila["importe_total"]),
            "Prestador": str(fila.get("prestador_nombre") or ""),
            "Período": str(fila.get("periodo") or ""),
            "DNI Afiliado": str(fila.get("afiliado_dni") or ""),
            "Nombre Afiliado": str(fila.get("afiliado_nombre") or ""),
            "Nomenclador SSSalud": label_actual,
            "Cant.": int(fila.get("cantidad") or 1),
            "Dep.": bool(fila.get("dependencia", False)),
            "Observación": obs_actual,
        })
        snapshots_previos.append({
            "estado": fila.get("estado"),
            "motivo": fila.get("motivo_estado"),
            "obs": obs_actual,
            "seleccionada": bool(fila.get("seleccionada", False)) and es_listo,
            "cantidad": int(fila.get("cantidad") or 1),
            "dependencia": bool(fila.get("dependencia", False)),
            "dni": str(fila.get("afiliado_dni") or ""),
        })

    df_editor = pd.DataFrame(filas_editor)
    alto_editor = min(480, max(240, 42 + 36 * len(filas_editor)))

    edited_df = st.data_editor(
        df_editor,
        use_container_width=True,
        hide_index=True,
        height=alto_editor,
        num_rows="fixed",
        key=f"tabla_revision_{st.session_state['revision_editor_version']}",
        disabled=["Estado", "Archivo", "Comp", "Tipo", "Importe", "Observación"],
        column_config={
            "Aut.": st.column_config.CheckboxColumn(
                "Aut.",
                width="small",
                pinned=True,
                help="Incluir esta factura en la autorización y exportación final.",
            ),
            "Estado": st.column_config.TextColumn("Estado", width="small", pinned=True),
            "Archivo": st.column_config.TextColumn("Archivo", width="medium"),
            "Comp": st.column_config.TextColumn("Comp", width="small"),
            "Tipo": st.column_config.TextColumn("Tipo", width="small"),
            "Importe": st.column_config.TextColumn("Importe", width="small"),
            "Prestador": st.column_config.TextColumn("Prestador", width="medium"),
            "Período": st.column_config.TextColumn(
                "Período",
                width="small",
                max_chars=6,
                help="Período prestacional en formato AAAAMM (ej: 202603).",
            ),
            "DNI Afiliado": st.column_config.TextColumn(
                "DNI Afiliado",
                width="small",
                max_chars=10,
                help="DNI del afiliado/beneficiario (7 u 8 dígitos numéricos).",
            ),
            "Nombre Afiliado": st.column_config.TextColumn(
                "Nombre Afiliado",
                width="medium",
                help="Opcional: Apellido y Nombre del afiliado.",
            ),
            "Nomenclador SSSalud": st.column_config.SelectboxColumn(
                "Nomenclador SSSalud",
                width="large",
                options=labels_opciones,
                required=True,
            ),
            "Cant.": st.column_config.NumberColumn(
                "Cant.",
                width="small",
                min_value=1,
                max_value=99999,
                step=1,
                format="%d",
                help="Para códigos 90 y 91 es fijo (4). Para Transporte (96), Apoyo (86) o Sin Recupero puede editarse.",
            ),
            "Dep.": st.column_config.CheckboxColumn(
                "Dep.",
                width="small",
                help="Adicional por Dependencia (sólo códigos SSSalud habilitados).",
            ),
            "Observación": st.column_config.TextColumn("Observación", width="large"),
        },
    )

    # Sincronizar ediciones de la tabla con el estado interno
    def _str_celda(val: Any) -> str:
        return "" if pd.isna(val) else str(val)

    def _bool_celda(val: Any) -> bool:
        return False if pd.isna(val) else bool(val)

    def _int_celda(val: Any, default: int = 1) -> int:
        if pd.isna(val):
            return default
        try:
            return max(1, int(val))
        except (TypeError, ValueError):
            return default

    requiere_reset_editor = False

    for idx, fila in enumerate(procesadas):
        if idx >= len(edited_df):
            break
        row = edited_df.iloc[idx]

        prestador_edit = _str_celda(row["Prestador"]).strip()
        fila["prestador_nombre"] = prestador_edit or f"Prestador ({formatear_cuit_guiones(fila['cuit_emisor'])})"
        fila["periodo"] = _str_celda(row["Período"]).strip()
        dni_raw = _str_celda(row["DNI Afiliado"])
        dni_digits = "".join(ch for ch in dni_raw if ch.isdigit())
        if dni_digits != dni_raw:
            requiere_reset_editor = True
        fila["afiliado_dni"] = dni_digits
        fila["afiliado_nombre"] = _str_celda(row["Nombre Afiliado"]).strip()

        label_elegido = _str_celda(row["Nomenclador SSSalud"]) or labels_opciones[0]
        op_elegida = mapa_label_a_op.get(label_elegido, opciones_prestacion[0])
        cambio_prestacion = (
            op_elegida["codigo"] != fila.get("codigo_nomenclador")
            or op_elegida["sin_recupero"] != fila.get("sin_recupero")
            or op_elegida["descripcion"] != fila.get("prestacion")
        )
        fila["codigo_nomenclador"] = op_elegida["codigo"]
        fila["prestacion"] = op_elegida["descripcion"]
        fila["sin_recupero"] = op_elegida["sin_recupero"]

        cant_edit = _int_celda(row["Cant."], 1)
        if cambio_prestacion:
            if op_elegida["codigo"] in CODIGOS_CON_CANTIDAD_CUATRO:
                fila["cantidad"] = 4
                requiere_reset_editor = True
            elif not op_elegida["sin_recupero"] and op_elegida["codigo"] != "96":
                fila["cantidad"] = int(op_elegida.get("cantidad_default") or 1)
                requiere_reset_editor = True
            else:
                fila["cantidad"] = cant_edit
        else:
            cod_actual = str(fila.get("codigo_nomenclador") or "")
            es_fijo_4 = (not fila.get("sin_recupero")) and (cod_actual in CODIGOS_CON_CANTIDAD_CUATRO)
            if es_fijo_4:
                if cant_edit != 4:
                    requiere_reset_editor = True
                fila["cantidad"] = 4
            else:
                fila["cantidad"] = cant_edit

        admite_dep = (not fila.get("sin_recupero")) and (
            str(fila.get("codigo_nomenclador") or "") in CODIGOS_DEPENDENCIA_VALIDOS
        )
        dep_edit = _bool_celda(row["Dep."])
        if not admite_dep and dep_edit:
            requiere_reset_editor = True
        fila["dependencia"] = dep_edit if admite_dep else False

        fila["seleccionada"] = _bool_celda(row["Aut."])

    # Recalcular validaciones después de sincronizar
    recalcular_estados_lote(procesadas)

    cambio_estado_o_obs = False
    for idx, fila in enumerate(procesadas):
        if idx >= len(edited_df):
            break
        prev = snapshots_previos[idx]
        obs_nueva = _texto_observacion(fila)
        sel_post = bool(fila.get("seleccionada", False)) and (fila.get("estado") == "Listo")
        if sel_post != _bool_celda(edited_df.iloc[idx]["Aut."]):
            requiere_reset_editor = True
        if (
            fila.get("estado") != prev["estado"]
            or obs_nueva != prev["obs"]
            or int(fila.get("cantidad") or 1) != prev["cantidad"]
            or bool(fila.get("dependencia", False)) != prev["dependencia"]
        ):
            cambio_estado_o_obs = True

    if requiere_reset_editor:
        st.session_state["revision_editor_version"] += 1
        st.rerun()
    elif cambio_estado_o_obs:
        st.rerun()

    # Poblar métricas superiores y barra compacta de acciones sobre la tabla
    listas_post = [f for f in procesadas if f.get("estado") == "Listo"]
    con_error_post = [f for f in procesadas if f.get("estado") != "Listo"]
    seleccionadas_post = [f for f in procesadas if f.get("seleccionada") and f.get("estado") == "Listo"]

    color_err = "#dc2626" if len(con_error_post) > 0 else "#64748b"
    metrics_ph.markdown(
        f"""
        <div class="rev-kpi-bar">
            <div class="rev-kpi-item">
                <p class="rev-kpi-label">📋 Total en Revisión</p>
                <p class="rev-kpi-val" style="color: #0f172a;">{len(procesadas)}</p>
            </div>
            <div class="rev-kpi-item">
                <p class="rev-kpi-label" style="color: #166534;">✅ Listas (Válidas)</p>
                <p class="rev-kpi-val" style="color: #15803d;">{len(listas_post)}</p>
            </div>
            <div class="rev-kpi-item">
                <p class="rev-kpi-label" style="color: {color_err};">⚠️ Con Observaciones</p>
                <p class="rev-kpi-val" style="color: {color_err};">{len(con_error_post)}</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    excel_bytes = generar_excel_revision(procesadas, rechazadas, cuit_inst)
    with toolbar_ph.container(horizontal=True, horizontal_alignment="right", vertical_alignment="center", gap="small"):
        if st.button("☑️ Seleccionar todas", key="btn_sel_all_paso3"):
            for f in procesadas:
                f["seleccionada"] = (f.get("estado") == "Listo")
            st.session_state["revision_editor_version"] += 1
            st.rerun()
        if st.button("⬜ Desmarcar todas", key="btn_desel_all_paso3"):
            for f in procesadas:
                f["seleccionada"] = False
            st.session_state["revision_editor_version"] += 1
            st.rerun()
        st.download_button(
            label="",
            data=excel_bytes,
            file_name=f"SAPI_Carga_Masiva_Facturas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_excel_icon",
            type="tertiary",
            help="Exportar a Excel (.xlsx)",
        )
        st.button(
            label="",
            key="btn_info_icon",
            type="tertiary",
            help=(
                "Verifique o complete el Período (AAAAMM), el DNI del Afiliado y la "
                "Codificación del Nomenclador SSSalud (Códigos 1 al 99 o Sin Recupero)."
            ),
        )

    # Barra inferior de acciones
    st.markdown("---")
    col_nav_back, _, col_confirm = st.columns([1.0, 1.0, 1.4], gap="small")

    with col_nav_back:
        if st.button("← Volver a Informe", use_container_width=True):
            st.session_state["paso_actual"] = "informe"
            st.rerun()

    with col_confirm:
        if st.button(
            f"✅ Confirmar y Registrar {len(seleccionadas_post)} Factura(s)",
            type="primary",
            use_container_width=True,
            disabled=(len(seleccionadas_post) == 0),
        ):
            st.session_state["lote_confirmado"] = True
            st.session_state["facturas_registradas_final"] = [dict(f) for f in seleccionadas_post]

    # Panel de Confirmación Final cuando el usuario registra el lote
    if st.session_state.get("lote_confirmado") and st.session_state.get("facturas_registradas_final"):
        registradas = st.session_state["facturas_registradas_final"]
        con_rec = [f for f in registradas if not f.get("sin_recupero")]
        sin_rec = [f for f in registradas if f.get("sin_recupero")]
        total_imp = sum(float(f.get("importe_total") or 0.0) for f in registradas)

        st.markdown("---")
        st.success(
            f"🎉 **¡Lote confirmado y autorizado con éxito!** Se registraron **{len(registradas)} factura(s)** "
            f"por un total de **{formatear_moneda_ars(total_imp)}** "
            f"({len(con_rec)} Con Recupero SSSalud · {len(sin_rec)} Sin Recupero)."
        )

        df_conf = pd.DataFrame([
            {
                "#": i + 1,
                "Identificador SSSalud": f["registro_identificador"],
                "Prestador": f["prestador_nombre"],
                "CUIT Emisor": formatear_cuit_guiones(f["cuit_emisor"]),
                "Comprobante": f"{f['tipo_comprobante_etiqueta']} {f['comprobante_visible']}",
                "CAE": f["cae"],
                "Período": f["periodo"],
                "DNI Afiliado": f["afiliado_dni"],
                "Cód. SSSalud": "S/R" if f.get("sin_recupero") else f["codigo_nomenclador"],
                "Prestación": f["prestacion"],
                "Cant.": f["cantidad"],
                "Importe": formatear_moneda_ars(f["importe_total"]),
            }
            for i, f in enumerate(registradas)
        ])
        st.dataframe(df_conf, use_container_width=True, hide_index=True)

        c_dl1, c_dl2, c_dl3 = st.columns(3)
        with c_dl1:
            txt_bytes = generar_txt_sssalud(registradas)
            st.download_button(
                "📑 Descargar TXT Rendición SSSalud (Con Recupero)",
                data=txt_bytes,
                file_name=f"DATOS_SSSALUD_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                mime="text/plain",
                use_container_width=True,
            )
        with c_dl2:
            zip_reg = generar_zip_sssalud_adaptado(registradas, rechazadas)
            st.download_button(
                "📦 Descargar ZIP PDFs 1 Página (SSSalud)",
                data=zip_reg,
                file_name=f"PDFs_SSSalud_Autorizados_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                mime="application/zip",
                use_container_width=True,
            )
        with c_dl3:
            if st.button("🔄 Iniciar Nueva Tanda de Facturas", use_container_width=True):
                reiniciar_proceso_completo()
                st.rerun()


# ==============================================================================
# EJECUCIÓN DEL FLUJO EN 3 PASOS
# ==============================================================================
render_cabecera_y_stepper()

if st.session_state["paso_actual"] == "seleccion":
    render_paso_1_seleccion()
elif st.session_state["paso_actual"] == "informe":
    render_paso_2_informe()
elif st.session_state["paso_actual"] == "revision":
    render_paso_3_revision()
else:
    reiniciar_proceso_completo()
    st.rerun()
