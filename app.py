import streamlit as st
import os
import shutil
import fitz  # PyMuPDF
import numpy as np
import cv2
import base64
import re
import json
import pandas as pd
from urllib.parse import urlparse, parse_qs
import pytesseract
import unicodedata
import tempfile
import zipfile
import io

# --------- Configuración de Página ---------
st.set_page_config(
    page_title="Escáner de Facturas QR AFIP - OSPIF",
    page_icon="📂",
    layout="wide"
)

# --------- Estilos CSS (Estética Portal OSPIF) ---------
LOGO_BASE64 = "data:image/png;base64,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

st.markdown(f"""
    <style>
        :root {
            --primary: #3B86B6;
            --primary-dark: #2A658C;
            --accent: #F2C12E;
            --bg-light: #F4F7F9;
            --white: #FFFFFF;
            --text-main: #2C3E50;
            --text-muted: #596A7A;
            --border-color: #E2E8F0;
            --shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
            --radius: 8px;
        }

        /* Fondo y tipografía general */
        .stApp {
            background-color: var(--bg-light);
            font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            color: var(--text-main);
        }

        /* Ocultar elementos nativos innecesarios */
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}

        .block-container {
            padding-top: 1.5rem !important;
            padding-bottom: 3rem !important;
            max-width: 1100px !important;
        }

        /* Banner Principal estilo Hero OSPIF */
        .ospif-hero {
            background: linear-gradient(135deg, var(--primary-dark) 0%, var(--primary) 100%);
            color: var(--white);
            padding: 36px 24px;
            text-align: center;
            border-radius: var(--radius);
            border-bottom: 5px solid var(--accent);
            box-shadow: var(--shadow);
            margin-bottom: 28px;
        }

        .ospif-hero img {
            width: 90px;
            height: 90px;
            background: var(--white);
            border-radius: 50%;
            padding: 5px;
            margin-bottom: 12px;
            box-shadow: 0 4px 14px rgba(0,0,0,0.2);
            object-fit: contain;
        }

        .ospif-hero h1 {
            color: var(--white) !important;
            font-size: 2rem !important;
            font-weight: 700 !important;
            margin-bottom: 8px !important;
            letter-spacing: -0.5px;
            padding: 0 !important;
        }

        .ospif-hero p {
            color: rgba(255, 255, 255, 0.95) !important;
            font-size: 1.05rem !important;
            max-width: 680px;
            margin: 0 auto !important;
        }

        /* Tarjetas blancas con borde superior primario */
        .ospif-card {
            background: var(--white);
            border-radius: var(--radius);
            padding: 24px 28px;
            box-shadow: var(--shadow);
            border-top: 4px solid var(--primary);
            margin-bottom: 24px;
            border-left: 1px solid var(--border-color);
            border-right: 1px solid var(--border-color);
            border-bottom: 1px solid var(--border-color);
        }

        .ospif-card h3 {
            color: var(--primary-dark) !important;
            font-size: 1.3rem !important;
            font-weight: 700 !important;
            margin-bottom: 6px !important;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .ospif-card p {
            color: var(--text-muted) !important;
            font-size: 0.95rem !important;
            margin-bottom: 16px !important;
        }

        .step-badge {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 30px;
            height: 30px;
            background-color: var(--accent);
            color: var(--text-main);
            font-weight: 700;
            font-size: 0.95rem;
            border-radius: 50%;
        }

        /* Personalización de File Uploader */
        [data-testid="stFileUploader"] {
            background-color: var(--white);
            padding: 16px;
            border-radius: var(--radius);
            border: 1px dashed var(--primary);
        }

        /* Botón Primario (Amarillo Acento o Azul Primario) */
        div.stButton > button[kind="primary"] {
            background-color: var(--primary);
            color: var(--white);
            font-weight: 600;
            padding: 10px 24px;
            border-radius: 6px;
            border: none;
            transition: all 0.2s ease;
            width: 100%;
            font-size: 1rem;
        }

        div.stButton > button[kind="primary"]:hover {
            background-color: var(--primary-dark);
            color: var(--white);
        }

        /* Botón de Descarga Destacado (Estilo btn-accent) */
        div.stDownloadButton > button {
            background-color: var(--accent) !important;
            color: #1A252F !important;
            font-weight: 700 !important;
            padding: 12px 24px !important;
            border-radius: 6px !important;
            border: none !important;
            box-shadow: var(--shadow) !important;
            width: 100% !important;
            font-size: 1.05rem !important;
            transition: background-color 0.2s ease !important;
        }

        div.stDownloadButton > button:hover {
            background-color: #dfb021 !important;
            color: #1A252F !important;
        }

        /* Pie de página institucional */
        .ospif-footer {
            background-color: #1F2D3D;
            color: #CBD5E1;
            padding: 20px 24px;
            border-radius: var(--radius);
            margin-top: 40px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 16px;
            font-size: 0.88rem;
        }

        .ospif-footer-brand {
            display: flex;
            align-items: center;
            gap: 12px;
        }

        .ospif-footer-brand img {
            height: 40px;
            background: var(--white);
            border-radius: 50%;
            padding: 2px;
        }

        .ospif-footer-brand h4 {
            color: var(--white) !important;
            font-size: 1rem !important;
            margin: 0 !important;
        }
    </style>
""", unsafe_allow_html=True)

# --------- Constantes ---------
ZOOM = 4.0
ROI_CENTRAL = (0.1, 0.3, 0.8, 0.4)
TIPO_MAP = {'1':'1','4':'2','6':'3','9':'4','11':'5','15':'6','51':'7','54':'8'}
REGIONES_QR = [
    (0.00, 0.00, 0.25, 0.25), (0.00, 0.00, 0.50, 0.50),
    (0.00, 0.75, 0.25, 0.25), (0.00, 0.75, 0.50, 0.25),
    (0.00, 0.75, 1.00, 0.25), (0.00, 0.50, 0.50, 0.50)
]
MES_MAP = {'enero':'01','febrero':'02','marzo':'03','abril':'04','mayo':'05','junio':'06',
           'julio':'07','agosto':'08','septiembre':'09','setiembre':'09','octubre':'10',
           'noviembre':'11','diciembre':'12'}
ABBR_MAP = {'ene':'01','feb':'02','mar':'03','abr':'04','may':'05','jun':'06','jul':'07',
            'ago':'08','sep':'09','sept':'09','set':'09','oct':'10','nov':'11','dic':'12'}

# --------- Funciones Auxiliares ---------
def _normalize_digits(s):
    return re.sub(r"\D", "", str(s)) if s else ""

def _normalize_str(s):
    if s is None: return ""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.strip().lower()

def extraer_texto_pdf_y_roi(ruta_pdf):
    texto = ""
    try:
        doc = fitz.open(ruta_pdf)
        for p in doc: texto += p.get_text()
        doc.close()
    except: pass
    
    try:
        doc = fitz.open(ruta_pdf)
        for page in doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(ZOOM, ZOOM))
            img = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, pix.n)
            h, w = img.shape[:2]
            xf, yf, wf, hf = ROI_CENTRAL
            roi = img[int(h*yf):int(h*(yf+hf)), int(w*xf):int(w*(xf+wf))]
            gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
            texto += pytesseract.image_to_string(gray, lang='spa') + "\n"
        doc.close()
    except: pass
    return texto

# --- Lógica de DNI y Periodo (Sin Padrón) ---
KEYWORDS_DNI = r"(dni|documento|doc\.?|nro\.?\s*doc|número\s*de\s*documento)"
REGEX_DNI_KEY = re.compile(rf"{KEYWORDS_DNI}\D{{0,30}}(\d{{1,2}}[.,]\d{{3}}[.,]\d{{3}}|\d{{8}})", re.IGNORECASE)
REGEX_DNI_PUNTOS = re.compile(r"\b(\d{1,2}\.\d{3}\.\d{3})\b")
REGEX_DNI_COMAS  = re.compile(r"\b(\d{1,2},\d{3},\d{3})\b")
REGEX_DNI_CORRIDO = re.compile(r"\b(\d{8})\b")

def es_dni_valido(d):
    if len(d) != 8 or not d.isdigit(): return False
    if d.startswith('0000') or int(d) < 10000000: return False
    return True

def limpiar_dni(cad): return cad.replace('.', '').replace(',', '').replace(' ', '')

def extraer_dni_de_pdf(ruta_pdf):
    texto = extraer_texto_pdf_y_roi(ruta_pdf)
    candidatos = []
    for m in REGEX_DNI_KEY.finditer(texto):
        d = limpiar_dni(m.group(1))
        if es_dni_valido(d): candidatos.append(d)
    if not candidatos:
        for m in REGEX_DNI_PUNTOS.finditer(texto):
            d = limpiar_dni(m.group(1))
            if es_dni_valido(d): candidatos.append(d)
    if not candidatos:
        for m in REGEX_DNI_COMAS.finditer(texto):
            d = limpiar_dni(m.group(1))
            if es_dni_valido(d): candidatos.append(d)
    if not candidatos:
        for m in REGEX_DNI_CORRIDO.finditer(texto):
            d = m.group(1)
            if es_dni_valido(d): candidatos.append(d)
            
    return candidatos[0] if candidatos else ''

# --- Periodo ---
REGEX_PERIODO_AAAAMM = re.compile(r"\b(20\d{2})(0[1-9]|1[0-2])\b")
REGEX_PERIODO_AAAA_MM = re.compile(r"\b(20\d{2})-(0[1-9]|1[0-2])\b")
_MONTH_KEYS = list(MES_MAP.keys()) + list(ABBR_MAP.keys())
_MONTH_KEYS.sort(key=len, reverse=True)
MONTH_PATTERN = r"(" + "|".join(map(re.escape, _MONTH_KEYS)) + r")"
REGEX_PERIODO_TEXTO = re.compile(rf"\b{MONTH_PATTERN}\s*(?:,|\s+de)?\s*(20\d{{2}})\b", re.IGNORECASE)

def extraer_periodo_de_pdf(ruta_pdf):
    raw = extraer_texto_pdf_y_roi(ruta_pdf)
    texto = _normalize_str(raw)
    m = REGEX_PERIODO_AAAAMM.search(texto)
    if m: return m.group(1) + m.group(2)
    m = REGEX_PERIODO_AAAA_MM.search(texto)
    if m: return m.group(1) + m.group(2)
    m = REGEX_PERIODO_TEXTO.search(texto)
    if m:
        mes_txt = _normalize_str(m.group(1))
        yyyy = m.group(2)
        mm = MES_MAP.get(mes_txt, ABBR_MAP.get(mes_txt))
        if mm: return yyyy + mm
    return ''

# --- QR ---
def extraer_url_de_pdf(ruta_pdf):
    try:
        doc = fitz.open(ruta_pdf)
        pix = doc[0].get_pixmap(matrix=fitz.Matrix(ZOOM,ZOOM))
        img = np.frombuffer(pix.samples, np.uint8).reshape(pix.height,pix.width,pix.n)
        if pix.n == 4: img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
        doc.close()
        detector = cv2.QRCodeDetector()
        
        d, _, _ = detector.detectAndDecode(img)
        if d and d.startswith(('http','https')): return d
        
        for xf,yf,wf,hf in REGIONES_QR:
             h,w = img.shape[:2]
             roi = img[int(h*yf):int(h*(yf+hf)), int(w*xf):int(w*(xf+wf))]
             d, _, _ = detector.detectAndDecode(roi)
             if d and d.startswith(('http','https')): return d
    except: pass
    return None

def parsear_campos_qr(url):
    qs = parse_qs(urlparse(url).query)
    p = qs.get('p',[None])[0] or (re.search(r'[?&]p=([^&]+)',url) or [None])[0]
    if not p: return None,None,None,None
    try:
        dec = base64.b64decode(p).decode('utf-8')
    except: return None,None,None,None
    if dec.strip().startswith('{'):
        try:
            d = json.loads(dec)
            return (str(d.get('cuit','')), str(d.get('tipoCmp','')), str(d.get('ptoVta','')), str(d.get('nroCmp','')))
        except: return None,None,None,None
    parts = dec.split('|')
    return (parts[0], parts[1], parts[2], parts[3] if len(parts)>=4 else None)


# ================== INTERFAZ VISUAL ESTILO OSPIF ==================

# Encabezado Hero
st.markdown(f"""
    <div class="ospif-hero">
        <img src="{LOGO_BASE64}" alt="Escudo OSPIF">
        <h1>Escáner de Facturas QR AFIP</h1>
        <p>Portal de Discapacidad OSPIF — Herramienta digital para la lectura automatizada de comprobantes, decodificación de QR, detección de DNI y período prestacional.</p>
    </div>
""", unsafe_allow_html=True)

# Tarjeta 1: Carga de Documentación
st.markdown("""
    <div class="ospif-card">
        <h3><span class="step-badge">1</span> Carga de Facturas en PDF</h3>
        <p>Seleccioná o arrastrá los comprobantes en formato PDF que deseás auditar y renombrar.</p>
    </div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Subir Facturas PDF",
    type="pdf",
    accept_multiple_files=True,
    label_visibility="collapsed"
)

st.write("")

# Tarjeta 2: Configuración y Opciones
st.markdown("""
    <div class="ospif-card">
        <h3><span class="step-badge">2</span> Opciones de Procesamiento</h3>
        <p>Definí los parámetros de salida para los archivos PDF procesados y el reporte de auditoría.</p>
    </div>
""", unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)
with col1:
    opt_renombrar = st.checkbox("Renombrar archivos PDF (CUIT_TIPO_PTO_NRO)", value=True)
with col2:
    opt_excel = st.checkbox("Generar reporte detallado en Excel", value=True)
with col3:
    opt_solo_original = st.checkbox("Conservar solo la 1° hoja (Factura original)", value=True)

st.write("")

# Botón de Ejecución
if st.button("🚀 Procesar Facturas Cargadas", type="primary"):
    if not uploaded_files:
        st.warning("⚠️ Por favor, cargá al menos un archivo PDF antes de iniciar el procesamiento.")
    else:
        with tempfile.TemporaryDirectory() as temp_dir:
            paths = []
            for uploaded_file in uploaded_files:
                p = os.path.join(temp_dir, uploaded_file.name)
                with open(p, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                paths.append(p)

            progress_bar = st.progress(0)
            status_text = st.empty()
            
            rows = []
            files_to_zip = []

            for i, path in enumerate(paths):
                fn = os.path.basename(path)
                status_text.markdown(f"**⏳ Procesando comprobante ({i+1}/{len(paths)}):** `{fn}`")
                
                error, url, dni, periodo = '', '', '', ''
                cuit, tipo, pto, nro, sss = '', '', '', '', ''

                # Extracciones sin padrón
                dni = extraer_dni_de_pdf(path)
                periodo = extraer_periodo_de_pdf(path)
                url_detect = extraer_url_de_pdf(path)

                if not url_detect:
                    error = 'QR no detectado'
                else:
                    url = url_detect
                    cuit, tipo, pto, nro = parsear_campos_qr(url)
                    sss = TIPO_MAP.get(tipo,'')
                    missing = [n for n,v in zip(['CUIT','Tipo','Pto','Nro'], [cuit,tipo,pto,nro]) if not v]
                    if missing: error = 'Faltan campos: ' + ', '.join(missing)

                if dni and nro and dni.lstrip('0') == (nro or '').lstrip('0'): dni = ''
                if dni and cuit:
                    cuit_digits = re.sub(r'\D', '', cuit)
                    if len(cuit_digits) == 11:
                        dni_prestador = cuit_digits[2:10]
                        if dni == dni_prestador: dni = ''

                nombre_calculado = f"{cuit}_{sss}_{pto}_{nro}" if (cuit and sss and pto and nro) else ""

                try:
                    doc = fitz.open(path); new = fitz.open()
                    new.insert_pdf(doc)
                    
                    if opt_solo_original:
                        for p in range(new.page_count-1, 0, -1):
                            new.delete_page(p)
                    
                    if opt_renombrar and not error and nombre_calculado:
                        nombre_final_zip = nombre_calculado + ".pdf"
                    else:
                        nombre_final_zip = fn
                    
                    out_path = os.path.join(temp_dir, "temp_processed_" + fn)
                    new.save(out_path)
                    new.close(); doc.close()
                    
                    files_to_zip.append((nombre_final_zip, out_path))
                    
                except Exception as e:
                    error = f"Error guardando PDF: {e}" if not error else error + f" | Error PDF: {e}"
                    files_to_zip.append((fn, path))

                rows.append({
                    'Archivo Original': fn,
                    'Archivo Generado': nombre_calculado + ".pdf" if (not error and nombre_calculado) else '',
                    'Codigo QR': url,
                    'DNI Detectado': dni,
                    'Periodo': periodo,
                    'Estado / Error': error if error else 'OK'
                })
                
                progress_bar.progress((i + 1) / len(paths))

            status_text.empty()
            progress_bar.empty()

            df_res = pd.DataFrame(rows, columns=['Archivo Original', 'Archivo Generado', 'Codigo QR', 'DNI Detectado', 'Periodo', 'Estado / Error'])

            # Generar archivo ZIP en memoria
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                if opt_excel:
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        df_res.to_excel(writer, index=False)
                    zf.writestr("Reporte_Procesamiento.xlsx", excel_buffer.getvalue())
                
                for name, filepath in files_to_zip:
                    zf.write(filepath, name)
            
            st.success("✅ ¡Procesamiento de facturas finalizado correctamente!")

            # Vista previa de la tabla de resultados
            st.markdown("""
                <div class="ospif-card">
                    <h3><span class="step-badge">3</span> Resumen de Auditoría y Descarga</h3>
                    <p>Verificá los datos extraídos en la tabla inferior y descargá el paquete ZIP con los comprobantes listos.</p>
                </div>
            """, unsafe_allow_html=True)

            st.dataframe(df_res, use_container_width=True)

            st.write("")
            st.download_button(
                label="📥 Descargar Paquete de Facturas Procesadas (.ZIP)",
                data=zip_buffer.getvalue(),
                file_name="facturas_procesadas_ospif.zip",
                mime="application/zip"
            )

# Pie de página institucional
st.markdown(f"""
    <div class="ospif-footer">
        <div class="ospif-footer-brand">
            <img src="{LOGO_BASE64}" alt="Logo OSPIF">
            <div>
                <h4>Portal de Discapacidad OSPIF</h4>
                <span>Obra Social del Personal de la Industria Fideera</span>
            </div>
        </div>
        <div>
            <span>Módulo de Auditoría y Digitalización de Facturación</span>
        </div>
    </div>
""", unsafe_allow_html=True)
