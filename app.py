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
    formatear_periodo_legible,
    generar_csv_revision,
    generar_excel_revision,
    generar_txt_sssalud,
    generar_zip_sssalud_adaptado,
    obtener_opciones_selector_prestacion,
    recalcular_estados_lote,
    renderizar_preview_pdf_png,
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

    /* Centrado y simetría de tarjetas y controles en Paso 1 */
    .sapi-card-centered {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 1.25rem 1.5rem;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
        margin-bottom: 0.9rem;
        text-align: center;
        width: 100%;
        box-sizing: border-box;
    }
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
        min-height: 150px !important;
        padding: 1.5rem 1.25rem !important;
        gap: 0.75rem !important;
        border-radius: 14px !important;
        margin: 0 auto !important;
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
    [data-testid="stFileUploaderDropzoneInstructions"] > div {
        align-items: center !important;
        text-align: center !important;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] span {
        text-align: center !important;
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
</style>
""",
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
        "facturas_procesadas": [],
        "facturas_rechazadas": [],
        "descargar_formato_sssalud": False,
        "zip_sssalud_bytes": None,
        "lote_confirmado": False,
        "facturas_registradas_final": [],
        "uploader_key": 0,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


inicializar_estado()


def reiniciar_proceso_completo() -> None:
    st.session_state["paso_actual"] = "seleccion"
    st.session_state["facturas_procesadas"] = []
    st.session_state["facturas_rechazadas"] = []
    st.session_state["zip_sssalud_bytes"] = None
    st.session_state["lote_confirmado"] = False
    st.session_state["facturas_registradas_final"] = []
    st.session_state["uploader_key"] += 1


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

    st.markdown(
        f"""
        <div class="sapi-card-centered">
            <p class="sapi-card-title" style="text-align: center;">📂 Selección de Facturas Digitales en PDF</p>
            <p class="sapi-card-sub" style="text-align: center; margin-bottom: 0;">
                Seleccione o arrastre los archivos PDF de las facturas electrónicas AFIP/ARCA.<br>
                <b>Límite de versión Demo: hasta {MAX_PDFS_POR_TANDA_DEMO} archivos PDF por tanda</b> (máx. 15 MB por archivo).
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_files = st.file_uploader(
        f"Seleccione hasta {MAX_PDFS_POR_TANDA_DEMO} archivos PDF",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"pdf_uploader_{st.session_state['uploader_key']}",
    )

    if uploaded_files:
        total_subidos = len(uploaded_files)
        excede_limite = total_subidos > MAX_PDFS_POR_TANDA_DEMO

        if excede_limite:
            st.warning(
                f"⚠️ **Límite de Versión Demo ({MAX_PDFS_POR_TANDA_DEMO} PDFs por tanda):** "
                f"Ha seleccionado **{total_subidos} archivos**. Se procesarán únicamente los **primeros {MAX_PDFS_POR_TANDA_DEMO} PDFs** "
                f"y se omitirán los {total_subidos - MAX_PDFS_POR_TANDA_DEMO} restantes."
            )

        archivos_activos = uploaded_files[:MAX_PDFS_POR_TANDA_DEMO]
        archivos_omitidos = uploaded_files[MAX_PDFS_POR_TANDA_DEMO:]

        # Tabla resumen de archivos seleccionados
        df_files = pd.DataFrame([
            {
                "#": idx + 1,
                "Archivo PDF": f.name,
                "Tamaño (KB)": f"{len(f.getvalue()) / 1024:.1f} KB",
                "Estado en Tanda": "Incluido en la tanda (1-15)",
            }
            for idx, f in enumerate(archivos_activos)
        ])
        st.dataframe(df_files, use_container_width=True, hide_index=True, height=min(280, 45 + 35 * len(archivos_activos)))

        if archivos_omitidos:
            with st.expander(f"Ver {len(archivos_omitidos)} archivo(s) omitido(s) por el límite de {MAX_PDFS_POR_TANDA_DEMO} PDFs"):
                for f_om in archivos_omitidos:
                    st.caption(f"• {f_om.name} ({len(f_om.getvalue()) / 1024:.1f} KB)")

        col_btn_clear, col_btn_scan = st.columns([1, 2])
        with col_btn_clear:
            if st.button("🗑️ Vaciar lista de PDFs", use_container_width=True):
                st.session_state["uploader_key"] += 1
                st.rerun()

        with col_btn_scan:
            boton_deshabilitado = not cuit_valido
            if st.button(
                f"🔍 Escanear y Validar {len(archivos_activos)} Factura(s) Digital(es) →",
                type="primary",
                use_container_width=True,
                disabled=boton_deshabilitado,
                help=(
                    "Corrija el CUIT Institucional (11 dígitos) o déjelo vacío para continuar."
                    if boton_deshabilitado
                    else None
                ),
            ):
                barra = st.progress(0, text="Iniciando lectura multicapa de códigos QR AFIP/ARCA...")
                lote_tuplas = [(f.name, f.getvalue()) for f in archivos_activos]

                def _actualizar_progreso(actual: int, total: int, nombre: str) -> None:
                    porc = int((actual / max(total, 1)) * 100)
                    barra.progress(porc, text=f"Escaneando ({actual}/{total}): {nombre}...")

                resultado_lote = escanear_lote_pdfs(
                    archivos_pdf=lote_tuplas,
                    cuit_institucional=cuit_limpio if cuit_valido else "",
                    progreso_callback=_actualizar_progreso,
                )
                barra.empty()

                st.session_state["cuit_institucional_validado"] = cuit_limpio if cuit_valido else ""
                st.session_state["facturas_procesadas"] = resultado_lote["procesadas"]
                st.session_state["facturas_rechazadas"] = resultado_lote["rechazadas"]
                st.session_state["zip_sssalud_bytes"] = generar_zip_sssalud_adaptado(
                    resultado_lote["procesadas"],
                    resultado_lote["rechazadas"],
                )
                st.session_state["lote_confirmado"] = False
                st.session_state["paso_actual"] = "informe"
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


# ==============================================================================
# PASO 2: INFORME DE VALIDACIÓN
# ==============================================================================
def render_paso_2_informe() -> None:
    procesadas: List[Dict[str, Any]] = st.session_state["facturas_procesadas"]
    rechazadas: List[Dict[str, Any]] = st.session_state["facturas_rechazadas"]
    cuit_inst = st.session_state.get("cuit_institucional_validado", "")

    importe_total_proc = sum(float(f.get("importe_total") or 0.0) for f in procesadas)

    # Tarjetas Resumen de Procesadas vs Rechazadas (igual a SAPI Web)
    st.markdown(
        f"""
        <div class="kpi-grid">
            <div class="kpi-box-ok">
                <div>
                    <p style="margin:0; font-size:0.95rem; font-weight:700; color:#166534;">
                        ✅ Facturas Procesadas con Éxito
                    </p>
                    <p style="margin:0.2rem 0 0 0; font-size:0.82rem; color:#15803d;">
                        QR AFIP/ARCA decodificado · Importe acumulado: <b>{formatear_moneda_ars(importe_total_proc)}</b>
                    </p>
                </div>
                <div class="kpi-num-ok">{len(procesadas)}</div>
            </div>
            <div class="kpi-box-err">
                <div>
                    <p style="margin:0; font-size:0.95rem; font-weight:700; color:#991b1b;">
                        ❌ Facturas No Procesadas / Rechazadas
                    </p>
                    <p style="margin:0.2rem 0 0 0; font-size:0.82rem; color:#b91c1c;">
                        {"CUIT Institucional verificado: " + formatear_cuit_guiones(cuit_inst) if cuit_inst else "Comprobantes sin QR válido o Notas de Crédito"}
                    </p>
                </div>
                <div class="kpi-num-err">{len(rechazadas)}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col_proc, col_rech = st.columns(2, gap="large")

    # Columna 1: Tabla de Facturas Procesadas
    with col_proc:
        st.markdown("#### ✅ Facturas Procesadas")
        if procesadas:
            df_proc = pd.DataFrame([
                {
                    "Archivo": f["archivo"],
                    "Prestador": f"{f['prestador_nombre']} ({formatear_cuit_guiones(f['cuit_emisor'])})",
                    "Comprobante SSSalud": f["registro_identificador"],
                    "Tipo": f["tipo_comprobante_etiqueta"],
                    "CAE": f["cae"],
                    "Emisión": f["fecha_emision_ddmmaaaa"],
                    "Importe": formatear_moneda_ars(f["importe_total"]),
                }
                for f in procesadas
            ])
            st.dataframe(df_proc, use_container_width=True, hide_index=True, height=290)
        else:
            st.warning("Ninguna factura pudo ser procesada con éxito en este lote.")

    # Columna 2: Tabla de Facturas No Procesadas / Rechazadas
    with col_rech:
        st.markdown("#### ❌ Facturas No Procesadas")
        if rechazadas:
            etiquetas_err = {
                "nota_credito": "🚫 Nota de Crédito",
                "cuit_invalido": "🏛️ Entidad No Coincide",
                "sin_qr": "🔍 Sin QR / Ilegible",
            }
            df_rech = pd.DataFrame([
                {
                    "Archivo": r["archivo"],
                    "Clasificación": etiquetas_err.get(r.get("tipo_error", "sin_qr"), "Error"),
                    "Motivo del Rechazo": r["motivo"],
                }
                for r in rechazadas
            ])
            st.dataframe(df_rech, use_container_width=True, hide_index=True, height=290)
        else:
            st.success("🎉 ¡Excelente! Todas las facturas de la tanda fueron validadas sin rechazos.")

    # Opción SSSalud: Adaptar facturas a 1 página y descargar en 3 carpetas (igual a SAPI Web)
    st.markdown("---")
    with st.container():
        col_chk, col_dl_zip = st.columns([1.8, 1.2], gap="medium")
        with col_chk:
            descargar_sssalud = st.checkbox(
                "📦 Descargar las facturas adaptadas en formato de la SSSalud (1 página · 3 carpetas)",
                value=st.session_state["descargar_formato_sssalud"],
                help=(
                    "Adapta automáticamente cada PDF dejando únicamente la primera página y los empaqueta en un archivo .ZIP "
                    "dividido en: 'Facturas Procesadas' (renombradas como CUIT_TIPO_PTO_NRO.pdf), "
                    "'No Procesadas - CUIT No Corresponde a la Entidad' y 'No Procesadas - Sin QR'."
                ),
            )
            st.session_state["descargar_formato_sssalud"] = descargar_sssalud
            st.caption(
                "Recorta cada PDF a **1 sola página** y los clasifica en **3 carpetas**: "
                "`Facturas Procesadas` (`CUIT_TIPO_PTO_NRO.pdf`), "
                "`No Procesadas - CUIT No Corresponde a la Entidad` y `No Procesadas - Sin QR`."
            )

        with col_dl_zip:
            zip_bytes = st.session_state.get("zip_sssalud_bytes")
            if not zip_bytes:
                zip_bytes = generar_zip_sssalud_adaptado(procesadas, rechazadas)
                st.session_state["zip_sssalud_bytes"] = zip_bytes

            st.download_button(
                label="⬇️ Descargar Facturas_SSSalud_Adaptadas.zip (1 pág.)",
                data=zip_bytes,
                file_name=f"Facturas_SSSalud_Adaptadas_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                mime="application/zip",
                use_container_width=True,
                type="secondary",
            )

    st.markdown("---")
    col_back, col_spacer, col_next = st.columns([1.1, 0.8, 1.4])
    with col_back:
        if st.button("← Seleccionar otros archivos", use_container_width=True):
            reiniciar_proceso_completo()
            st.rerun()

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

    # Banner informativo sobre modo de validación sin listado de autorizaciones
    st.markdown(
        """
        <div class="sapi-card" style="background: #f8fafc; border-left: 4px solid #2563eb; padding: 0.95rem 1.25rem;">
            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:0.5rem;">
                <div>
                    <p style="margin:0; font-size:0.92rem; font-weight:700; color:#1e3a8a;">
                        📋 Revisión y Autorización de Comprobantes (Modo sin cruce de autorizaciones)
                    </p>
                    <p style="margin:0.2rem 0 0 0; font-size:0.82rem; color:#475569;">
                        Verifique o complete el <b>Período (AAAAMM)</b>, el <b>DNI del Afiliado (7-8 dígitos)</b> y la
                        <b>Codificación del Nomenclador SSSalud (Códigos 1 al 99 o Sin Recupero)</b>.
                        El sistema valida en tiempo real duplicados dentro del lote y consistencia de datos.
                    </p>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Si en el Paso 2 el usuario dejó tildado "Descargar formato SSSalud", mostrar acceso directo destacado
    if st.session_state.get("descargar_formato_sssalud"):
        zip_bytes = generar_zip_sssalud_adaptado(procesadas, rechazadas)
        st.session_state["zip_sssalud_bytes"] = zip_bytes
        col_aviso_zip, col_btn_zip_top = st.columns([2.2, 1])
        with col_aviso_zip:
            st.info("📦 **Formato SSSalud (1 página · 3 carpetas) preparado:** Puede descargar el archivo `.zip` adaptado en cualquier momento.")
        with col_btn_zip_top:
            st.download_button(
                "⬇️ Descargar ZIP SSSalud Adaptado",
                data=zip_bytes,
                file_name=f"Facturas_SSSalud_Adaptadas_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                mime="application/zip",
                use_container_width=True,
            )

    # Contadores en tiempo real y acciones rápidas de selección
    listas = [f for f in procesadas if f.get("estado") == "Listo"]
    con_error = [f for f in procesadas if f.get("estado") != "Listo"]
    seleccionadas = [f for f in procesadas if f.get("seleccionada") and f.get("estado") == "Listo"]
    importe_seleccionado = sum(float(f.get("importe_total") or 0.0) for f in seleccionadas)

    c_stat1, c_stat2, c_stat3, c_stat4, c_act1, c_act2 = st.columns([1, 1, 1, 1.3, 1.1, 1.1])
    c_stat1.metric("Total en Revisión", len(procesadas))
    c_stat2.metric("Listas (Válidas)", len(listas))
    c_stat3.metric("Con Observaciones", len(con_error))
    c_stat4.metric("Importe Seleccionado", formatear_moneda_ars(importe_seleccionado))

    with c_act1:
        if st.button("☑️ Seleccionar todas las Listas", use_container_width=True):
            for f in procesadas:
                f["seleccionada"] = (f.get("estado") == "Listo")
            st.rerun()
    with c_act2:
        if st.button("⬜ Desmarcar todas", use_container_width=True):
            for f in procesadas:
                f["seleccionada"] = False
            st.rerun()

    st.markdown("---")

    # Listado interactivo de revisión de cada factura (Estilo SAPI Web)
    id_a_eliminar = None

    for idx, fila in enumerate(procesadas):
        fid = fila["id"]
        es_listo = fila.get("estado") == "Listo"
        borde_color = "#86efac" if es_listo else "#fca5a5"
        fondo_header = "#f0fdf4" if es_listo else "#fef2f2"

        with st.container():
            st.markdown(
                f"""
                <div style="border: 1px solid {borde_color}; border-radius: 12px; padding: 0.7rem 1rem; background: {fondo_header}; margin-bottom: 0.35rem; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:0.5rem;">
                    <div style="display:flex; align-items:center; gap:0.65rem; flex-wrap:wrap;">
                        <span style="font-weight:800; color:#0f172a; font-size:0.9rem;">#{idx + 1}</span>
                        <span style="font-weight:700; color:#1e293b; font-size:0.9rem;">📄 {fila['archivo']}</span>
                        <span class="badge-info" title="Identificador estructurado SSSalud (CUIT_TIPO_PTO_NRO)">{fila['registro_identificador']}</span>
                        <span style="font-size:0.82rem; color:#334155;">
                            <b>{fila['tipo_comprobante_etiqueta']}</b> ({fila['comprobante_visible']}) ·
                            CAE: <code>{fila['cae']}</code> ·
                            Emisión: <b>{fila['fecha_emision_ddmmaaaa']}</b> ·
                            CUIT Emisor: <code>{formatear_cuit_guiones(fila['cuit_emisor'])}</code>
                            {f" · CUIT Receptor: <code>{formatear_cuit_guiones(fila['cuit_receptor'])}</code>" if fila.get('cuit_receptor') else ""}
                        </span>
                    </div>
                    <div style="display:flex; align-items:center; gap:0.75rem;">
                        <span style="font-size:1.05rem; font-weight:800; color:#0f172a;">{formatear_moneda_ars(fila['importe_total'])}</span>
                        {"<span class='badge-ok'>✓ Listo</span>" if es_listo else f"<span class='badge-err'>✕ {fila.get('motivo_estado', 'Error')}</span>"}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            # Controles editables de la fila
            col_sel, col_prest, col_per, col_dni, col_af_nom, col_nomen, col_cant, col_sr, col_del = st.columns(
                [0.55, 1.55, 0.95, 1.0, 1.25, 2.35, 0.7, 0.75, 0.55],
                gap="small",
            )

            with col_sel:
                sel_val = st.checkbox(
                    "Aut.",
                    value=bool(fila.get("seleccionada", False)) and es_listo,
                    key=f"sel_{fid}",
                    disabled=not es_listo,
                    help="Incluir esta factura en la autorización y exportación final.",
                )
                fila["seleccionada"] = sel_val

            with col_prest:
                nuevo_prestador = st.text_input(
                    "Prestador (Razón Social)",
                    value=str(fila.get("prestador_nombre") or ""),
                    key=f"prest_{fid}",
                )
                fila["prestador_nombre"] = nuevo_prestador.strip() or f"Prestador ({formatear_cuit_guiones(fila['cuit_emisor'])})"

            with col_per:
                nuevo_periodo = st.text_input(
                    "Período (AAAAMM)",
                    value=str(fila.get("periodo") or ""),
                    max_chars=6,
                    placeholder="202603",
                    key=f"per_{fid}",
                    help=f"Período prestacional en formato AAAAMM ({formatear_periodo_legible(fila.get('periodo', ''))}).",
                )
                fila["periodo"] = nuevo_periodo.strip()

            with col_dni:
                nuevo_dni = st.text_input(
                    "DNI Afiliado",
                    value=str(fila.get("afiliado_dni") or ""),
                    max_chars=10,
                    placeholder="7 u 8 dígitos",
                    key=f"dni_{fid}",
                    help="DNI del afiliado/beneficiario (7 u 8 dígitos numéricos).",
                )
                fila["afiliado_dni"] = "".join(ch for ch in nuevo_dni if ch.isdigit())

            with col_af_nom:
                nuevo_af_nombre = st.text_input(
                    "Nombre Afiliado (Opc.)",
                    value=str(fila.get("afiliado_nombre") or ""),
                    placeholder="Apellido y Nombre",
                    key=f"afnom_{fid}",
                )
                fila["afiliado_nombre"] = nuevo_af_nombre.strip()

            with col_nomen:
                label_actual = _obtener_label_actual(fila)
                idx_op = labels_opciones.index(label_actual) if label_actual in labels_opciones else 0
                label_elegido = st.selectbox(
                    "Cód. Nomenclador SSSalud / Prestación",
                    options=labels_opciones,
                    index=idx_op,
                    key=f"nom_{fid}",
                )
                op_elegida = mapa_label_a_op[label_elegido]
                cambio_prestacion = (
                    op_elegida["codigo"] != fila.get("codigo_nomenclador")
                    or op_elegida["sin_recupero"] != fila.get("sin_recupero")
                    or op_elegida["descripcion"] != fila.get("prestacion")
                )
                fila["codigo_nomenclador"] = op_elegida["codigo"]
                fila["prestacion"] = op_elegida["descripcion"]
                fila["sin_recupero"] = op_elegida["sin_recupero"]
                if cambio_prestacion:
                    if op_elegida["codigo"] in CODIGOS_CON_CANTIDAD_CUATRO:
                        fila["cantidad"] = 4
                    elif not op_elegida["sin_recupero"] and op_elegida["codigo"] != "96":
                        fila["cantidad"] = int(op_elegida.get("cantidad_default") or 1)

            with col_cant:
                cod_actual = str(fila.get("codigo_nomenclador") or "")
                es_fijo_4 = (not fila.get("sin_recupero")) and (cod_actual in CODIGOS_CON_CANTIDAD_CUATRO)
                if es_fijo_4:
                    fila["cantidad"] = 4
                nueva_cant = st.number_input(
                    "Cant.",
                    min_value=1,
                    max_value=99999,
                    value=int(fila.get("cantidad") or 1),
                    step=1,
                    disabled=es_fijo_4,
                    key=f"cant_{fid}_{ es_fijo_4 }",
                    help="Para códigos 90 y 91 es fijo (4). Para Transporte (96), Apoyo (86) o Sin Recupero puede editarse.",
                )
                fila["cantidad"] = 4 if es_fijo_4 else int(nueva_cant)

            with col_sr:
                admite_dep = (not fila.get("sin_recupero")) and (str(fila.get("codigo_nomenclador") or "") in CODIGOS_DEPENDENCIA_VALIDOS)
                if admite_dep:
                    dep_val = st.checkbox(
                        "Dep.",
                        value=bool(fila.get("dependencia", False)),
                        key=f"dep_{fid}",
                        help="Adicional por Dependencia (habilitado para este código SSSalud).",
                    )
                    fila["dependencia"] = dep_val
                else:
                    fila["dependencia"] = False
                    st.caption("S/R: **Sí**" if fila.get("sin_recupero") else f"Cód: **{fila.get('codigo_nomenclador') or '-'}**")

            with col_del:
                st.write("")
                if st.button("🗑️", key=f"del_{fid}", help="Quitar esta factura del lote"):
                    id_a_eliminar = fid

            # Visor de PDF integrado bajo demanda y advertencias de cronograma
            col_prev, col_adv = st.columns([1.2, 2.8])
            with col_prev:
                with st.expander(f"👁️ Ver Factura PDF ({fila['archivo']})"):
                    png_preview = renderizar_preview_pdf_png(fila.get("pdf_bytes", b""), zoom=1.35)
                    if png_preview:
                        st.image(png_preview, caption=f"Página 1 · {fila['archivo']} ({fila['registro_identificador']})", use_container_width=True)
                    else:
                        st.warning("No se pudo generar la vista previa gráfica del PDF.")
            with col_adv:
                if fila.get("advertencias"):
                    for adv in fila["advertencias"]:
                        st.markdown(f"<span class='badge-warn'>⚠️ {adv}</span>", unsafe_allow_html=True)

            st.markdown("<div style='margin-bottom: 0.65rem;'></div>", unsafe_allow_html=True)

    if id_a_eliminar is not None:
        st.session_state["facturas_procesadas"] = [f for f in procesadas if f["id"] != id_a_eliminar]
        st.rerun()

    # Recalcular después de las ediciones interactivas
    recalcular_estados_lote(procesadas)
    listas_post = [f for f in procesadas if f.get("estado") == "Listo"]
    seleccionadas_post = [f for f in procesadas if f.get("seleccionada") and f.get("estado") == "Listo"]

    # Barra inferior de acciones y exportaciones (igual a SAPI Web)
    st.markdown("---")
    col_nav_back, col_exp_csv, col_exp_xlsx, col_exp_zip, col_confirm = st.columns([1.05, 1.05, 1.15, 1.25, 1.5], gap="small")

    with col_nav_back:
        if st.button("← Volver a Informe", use_container_width=True):
            st.session_state["paso_actual"] = "informe"
            st.rerun()

    with col_exp_csv:
        csv_bytes = generar_csv_revision(procesadas)
        st.download_button(
            "📄 Exportar a CSV",
            data=csv_bytes,
            file_name=f"facturas_revision_{datetime.now().strftime('%Y-%m-%d')}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with col_exp_xlsx:
        excel_bytes = generar_excel_revision(procesadas, rechazadas, cuit_inst)
        st.download_button(
            "📊 Exportar a Excel (.xlsx)",
            data=excel_bytes,
            file_name=f"SAPI_Carga_Masiva_Facturas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_exp_zip:
        zip_actualizado = generar_zip_sssalud_adaptado(procesadas, rechazadas)
        st.session_state["zip_sssalud_bytes"] = zip_actualizado
        st.download_button(
            "📦 ZIP SSSalud (1 pág.)",
            data=zip_actualizado,
            file_name=f"Facturas_SSSalud_Adaptadas_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
            mime="application/zip",
            use_container_width=True,
        )

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
