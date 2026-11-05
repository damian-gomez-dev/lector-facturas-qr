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

# --------- Configuración Tesseract ---------
# Si lo corres en local y no detecta tesseract, descomenta la siguiente línea:
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

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

# --- Lógica de DNI y Periodo ---
KEYWORDS_DNI = r"(dni|documento|doc\.?-?|nro\.?\s*doc|número\s*de\s*documento)"
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
        
        # Intento rápido
        d, _, _ = detector.detectAndDecode(img)
        if d and d.startswith(('http','https')): return d
        
        # Intento con ROI
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


# ================== INTERFAZ STREAMLIT ==================

st.set_page_config(page_title="Portal de Escaneo de Facturas - Discapacidad", page_icon="📂", layout="centered")

st.markdown("""
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
    }
    .stApp {
        background-color: var(--bg-light);
        color: var(--text-main);
        font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    }
    .hero-generic {
        background: linear-gradient(135deg, #2A658C 0%, #3B86B6 100%);
        color: #FFFFFF;
        padding: 36px 24px;
        text-align: center;
        border-radius: 8px;
        border-bottom: 5px solid #F2C12E;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
        margin-bottom: 15px;
    }
    .hero-generic h1 {
        color: #FFFFFF !important;
        font-size: 2rem;
        font-weight: 700;
        margin-bottom: 8px;
        letter-spacing: -0.5px;
        padding-top: 0;
    }
    .hero-generic p {
        color: rgba(255, 255, 255, 0.95);
        font-size: 1rem;
        margin: 0;
    }
    .section-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-top: 4px solid #3B86B6;
        border-radius: 8px;
        padding: 20px 24px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.06);
        margin-bottom: 15px;
    }
    .section-card h3 {
        color: #2A658C !important;
        font-size: 1.25rem;
        font-weight: 700;
        margin-bottom: 6px;
        padding-top: 0;
    }
    .section-card p {
        color: #596A7A;
        font-size: 0.92rem;
        margin-bottom: 0;
    }
    [data-testid="stFileUploader"] {
        background-color: #FFFFFF;
        padding: 24px;
        border-radius: 8px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
        margin-bottom: 16px;
    }
    [data-testid="stFileUploader"] section {
        text-align: center;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
    }
    [data-testid="stFileUploader"] section > div {
        text-align: center;
    }
    [data-testid="stFileUploader"] small {
        display: block;
        text-align: center;
        width: 100%;
    }
    
        /* Estilo para que el botón ocupe el 100% del ancho del contenedor y esté alineado con las tarjetas */
    .stButton {
        width: 100% !important;
    }
    .stButton > button {
        width: 100% !important;
        display: block !important;
        background-color: #3B86B6 !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 16px 32px !important;
        font-size: 1.05rem !important;
        box-shadow: 0 4px 10px rgba(59, 134, 182, 0.3) !important;
        box-sizing: border-box !important;
    }    .stButton > button {
        background-color: #3B86B6 !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 16px 32px !important;
        font-size: 1.05rem !important;
        box-shadow: 0 4px 10px rgba(59, 134, 182, 0.3) !important;
        box-sizing: border-box !important;
        margin: 0 auto !important;
    }
    div.stButton > button:hover {
        background-color: #2A658C;
        color: #FFFFFF;
        box-shadow: 0 6px 14px rgba(42, 101, 140, 0.4);
    }
    div.stDownloadButton > button {
        background-color: #F2C12E;
        color: #1A252F;
        font-weight: 700;
        border-radius: 6px;
        border: none;
        padding: 14px 28px;
        width: 100%;
        font-size: 1.05rem;
    }
    div.stDownloadButton > button:hover {
        background-color: #dfb021;
        color: #1A252F;
    }
    .footer-generic {
        background-color: #1F2D3D;
        color: #CBD5E1;
        padding: 18px;
        text-align: center;
        border-radius: 8px;
        margin-top: 15px;
        font-size: 0.88rem;
        width: 100%;
        box-sizing: border-box;
    }
</style>

<div class="hero-generic">
    <h1>📂 Portal de Escaneo de Facturas Discapacidad</h1>
    <p>Sistema de procesamiento automatizado para facturas de Mecanismo de Integración</p>
</div>

<div class="section-card">
    <h3>📄 Carga de Comprobantes</h3>
    <p>Subí tus facturas en formato PDF para extraer automáticamente sus datos, incluyendo período de prestación y DNI del afiliado.</p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader("", type="pdf", accept_multiple_files=True, help="Arrastrá y soltá tus archivos PDF aquí")

# --- OPCIONES ---
col1, col2, col3 = st.columns(3)
with col1:
    opt_renombrar = st.checkbox("Renombrar archivos PDF", value=True, help="Si se marca, los archivos se renombrarán a CUIT_TIPO_PTO_NRO.pdf")
with col2:
    opt_excel = st.checkbox("Generar Excel reporte", value=True, help="Genera un Excel con el detalle.")
with col3:
    opt_solo_original = st.checkbox("Factura original", value=True, help="Si se marca, se eliminan las hojas extra, dejando solo la primera página.")

st.write("") # Espaciador estético

procesar_facturas = st.button("🚀 Procesar Facturas")

if procesar_facturas and uploaded_files:
    
    with tempfile.TemporaryDirectory() as temp_dir:
        
        # Guardar archivos subidos en temp
        paths = []
        for uploaded_file in uploaded_files:
            p = os.path.join(temp_dir, uploaded_file.name)
            with open(p, "wb") as f:
                f.write(uploaded_file.getbuffer())
            paths.append(p)

        progress_bar = st.progress(0)
        status_text = st.empty()
        
        rows = []
        files_to_zip = [] # Tuplas (nombre_en_zip, ruta_fisica)

        for i, path in enumerate(paths):
            fn = os.path.basename(path)
            status_text.text(f"Procesando ({i+1}/{len(paths)}): {fn}")
            
            error, url, dni, periodo = '', '', '', ''
            cuit, tipo, pto, nro, sss = '', '', '', '', ''

            # Extracciones
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

            # Determinar el nombre base detectado
            nombre_calculado = f"{cuit}_{sss}_{pto}_{nro}" if (cuit and sss and pto and nro) else ""

            # Lógica de guardado y manipulación PDF
            try:
                doc = fitz.open(path); new = fitz.open()
                new.insert_pdf(doc)
                
                # Eliminar hojas extra SOLO si el usuario lo pide
                if opt_solo_original:
                    for p in range(new.page_count-1,0,-1):
                        new.delete_page(p)
                
                # Decidir nombre final
                if opt_renombrar and not error and nombre_calculado:
                    nombre_final_zip = nombre_calculado + ".pdf"
                else:
                    nombre_final_zip = fn # Nombre original
                
                out_path = os.path.join(temp_dir, "temp_processed_" + fn)
                new.save(out_path)
                new.close(); doc.close()
                
                files_to_zip.append((nombre_final_zip, out_path))
                
            except Exception as e:
                # Fallback en caso de error crítico con el PDF
                error = f"Error guardando PDF: {e}" if not error else error + f" | Error PDF: {e}"
                files_to_zip.append((fn, path))

            rows.append({
                'Archivo Original': fn,
                'Archivo Generado': nombre_calculado + ".pdf" if (not error and nombre_calculado) else '',
                'Codigo QR': url,
                'DNI': dni,
                'Periodo': periodo,
                'Error': error
            })
            
            progress_bar.progress((i + 1) / len(paths))

        status_text.empty()

        # Generar ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            
            if opt_excel:
                df = pd.DataFrame(rows, columns=['Archivo Original', 'Archivo Generado', 'Codigo QR', 'DNI', 'Periodo', 'Error'])
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                zf.writestr("Reporte_Procesamiento.xlsx", excel_buffer.getvalue())
            
            for name, filepath in files_to_zip:
                zf.write(filepath, name)
        
        st.success("✅ ¡Procesamiento completado con éxito!")
        st.download_button(
            label="⬇️ Descargar Facturas Procesadas (.ZIP)",
            data=zip_buffer.getvalue(),
            file_name="facturas_procesadas.zip",
            mime="application/zip"
        )

st.markdown("""
<div class="footer-generic">
    Portal de Procesamiento de Facturas de Discapacidad
</div>
""", unsafe_allow_html=True)
