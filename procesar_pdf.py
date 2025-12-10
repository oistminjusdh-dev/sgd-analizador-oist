# procesar_pdf.py (versión final: extractor por código, asunto limpio y aprendizaje)
import os
import sys
import re
import json
from io import BytesIO
from datetime import datetime
import pandas as pd

# pdf parsing
try:
    import pdfplumber
except:
    pdfplumber = None

# OCR opcional
try:
    import pytesseract
    from PIL import Image
    OCR_AVAILABLE = True
except:
    OCR_AVAILABLE = False

# resource helper
def resource_path(relative):
    try:
        base = sys._MEIPASS
    except:
        base = os.path.abspath(".")
    return os.path.join(base, relative)

PERSONNEL_MASTER_FILE = resource_path("personal_oist.csv")
APRENDIZAJE_FILE = resource_path("modelo_aprendizaje.json")
ASUNTOS_FILE = resource_path("asuntos_aprendidos.json")

def cargar_json(path):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}

def guardar_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

# aprendizaje ligero
def registrar_correccion_nombre(original, corregido):
    data = cargar_json(APRENDIZAJE_FILE)
    key = original.upper().strip()
    val = corregido.title().strip()
    if key not in data:
        data[key] = {"corregido_a": val, "correcciones": 1}
    else:
        data[key]["corregido_a"] = val
        data[key]["correcciones"] += 1
    guardar_json(APRENDIZAJE_FILE, data)

def registrar_correccion_asunto(original, corregido):
    data = cargar_json(ASUNTOS_FILE)
    key = original.upper().strip()
    val = corregido.strip()
    if key not in data:
        data[key] = {"corregido_a": val, "correcciones": 1}
    else:
        data[key]["corregido_a"] = val
        data[key]["correcciones"] += 1
    guardar_json(ASUNTOS_FILE, data)

def aplicar_correcciones_nombre(nombre):
    data = cargar_json(APRENDIZAJE_FILE)
    key = (nombre or "").upper().strip()
    return data[key]["corregido_a"] if key in data else nombre

def aplicar_correcciones_asunto(asunto):
    data = cargar_json(ASUNTOS_FILE)
    key = (asunto or "").upper().strip()
    return data[key]["corregido_a"] if key in data else asunto

# lista maestra
def cargar_lista_maestra():
    if os.path.exists(PERSONNEL_MASTER_FILE):
        try:
            df = pd.read_csv(PERSONNEL_MASTER_FILE, header=None, names=["Nombre"])
            return df["Nombre"].fillna("").astype(str).str.upper().tolist()
        except:
            return []
    return []

def guardar_lista_maestra(lista):
    df = pd.DataFrame(list(set(lista)), columns=["Nombre"])
    df.to_csv(PERSONNEL_MASTER_FILE, index=False, header=False, encoding="utf-8")

def actualizar_lista_maestra(nombres):
    actuales = cargar_lista_maestra()
    nuevos = [n.upper() for n in nombres if n and n.upper() not in actuales]
    if nuevos:
        actuales.extend(nuevos)
        guardar_lista_maestra(actuales)
    return [n.title() for n in actuales]

# OCR fallback
def ocr_pdf_bytes(pdf_bytes):
    if not OCR_AVAILABLE:
        return ""
    try:
        from pdf2image import convert_from_bytes
    except:
        return ""
    try:
        pages = convert_from_bytes(pdf_bytes)
        text = []
        for img in pages:
            text.append(pytesseract.image_to_string(img, lang="spa"))
        return "\n".join(text)
    except:
        return ""

# asunto limpio (opción A: todo en una línea)
def extract_clean_asunto(block, codigo=None):
    if not block or not isinstance(block, str):
        return "Asunto no visible"
    t = block
    t = t.replace("\r", "\n")
    t = re.sub(r"\n{2,}", "\n", t)
    t = re.sub(r"[ \t]+", " ", t)
    # unir líneas (excepto cuando comienzan con ORIGEN/DESTINO/PROVEIDO/FECHA)
    t = re.sub(r"\n(?!\s*(ORIGEN|DESTINO|PROVEIDO|FECHA INGRESO))", " ", t, flags=re.IGNORECASE)
    # preferir ATENCION
    m = re.search(r"ATENCION\s*[:\-]?\s*(.*)", t, flags=re.IGNORECASE | re.DOTALL)
    if m:
        candidate = m.group(1).strip()
    else:
        # fallback: quitar destino/origen y tomar el resto
        candidate = re.sub(r"(DESTINO|ORIGEN).*", "", t, flags=re.IGNORECASE|re.DOTALL).strip()
    # cortar antes de PROVEIDO / FECHA INGRESO / ORIGEN / DESTINO / TOTAL DOCUMENTOS
    candidate = re.split(r"PROVEIDO|FECHA INGRESO|ORIGEN|DESTINO|TOTAL DOCUMENTOS|SISTEMA DE GESTION DOCUMENTAL", candidate, flags=re.IGNORECASE)[0]
    # quitar "días(s)" y fechas/hours
    candidate = re.sub(r"\d+\s*d[ií]as?\(s\)", "", candidate, flags=re.IGNORECASE)
    candidate = re.sub(r"\d{1,2}/\d{1,2}/\d{4}\s*\d{1,2}:\d{2}", "", candidate)
    candidate = re.sub(r"\s{2,}", " ", candidate).strip()
    # Resultado en una línea (opción A)
    candidate = candidate.strip()
    # Capitalizar de forma conservadora (no romper acrónimos)
    # Dejamos mayúsculas originales en palabras como "PIM" si aparecen
    # Convertimos a título pero preservamos siglas comúnmente uppercase:
    candidate = candidate.replace(" PIM ", " PIM ")
    if candidate:
        return candidate.strip()
    return "Asunto no visible"

# limpieza y formato final
def clean_extracted_df(df, fecha_hoy):
    if df is None or df.empty:
        return pd.DataFrame(columns=["Nombre_Personal","Codigo","Estado","Fecha_Recepcion_Str","Dias_En_Bandeja","Asunto","Fecha_Recepcion"])
    df = df.copy()
    df["Fecha_Recepcion"] = pd.to_datetime(df["Fecha_Recepcion_Str"], format="%d/%m/%Y %H:%M", errors="coerce")
    df["Fecha_Recepcion"] = df["Fecha_Recepcion"].fillna(pd.Timestamp.now())
    df["Dias_En_Bandeja"] = (fecha_hoy - df["Fecha_Recepcion"]).dt.days.astype(int)
    # asunto limpio por fila
    df["Asunto"] = df.apply(lambda r: extract_clean_asunto(r["Asunto"], r.get("Codigo")), axis=1)
    # limpiar nombre
    def clean_personal(n):
        if not n or not isinstance(n, str):
            return "N/A"
        s = n.strip()
        if "/" in s:
            s = s.split("/")[-1].strip()
        if "," in s:
            a,b = [x.strip() for x in s.split(",",1)]
            s = f"{b} {a}"
        return s.title()
    df["Nombre_Personal"] = df["Nombre_Personal"].apply(clean_personal)
    # aplicar correcciones aprendidas
    df["Nombre_Personal"] = df["Nombre_Personal"].apply(aplicar_correcciones_nombre)
    df["Asunto"] = df["Asunto"].apply(aplicar_correcciones_asunto)
    df = df.sort_values(["Nombre_Personal","Dias_En_Bandeja"], ascending=[True,False])
    df["Fecha_Recepcion_Str"] = df["Fecha_Recepcion"].dt.strftime("%d/%m/%Y %H:%M")
    return df[["Nombre_Personal","Codigo","Estado","Fecha_Recepcion_Str","Dias_En_Bandeja","Asunto","Fecha_Recepcion"]]

# extractor por CÓDIGO (pattern fijo)
PATRON_CODIGO = r"\d{4}\s*USC[-\s]?\d+|\d{4}USC-\d+|\d{4}USC\d+"

def extraer_datos_pdf(pdf_bytes):
    texto = ""
    if pdfplumber:
        try:
            with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
                for p in pdf.pages:
                    texto += (p.extract_text() or "") + "\n"
        except:
            texto = ""
    if not texto.strip():
        texto = ocr_pdf_bytes(pdf_bytes)
    texto = re.sub(r"\r", "\n", texto)
    texto = re.sub(r"\n{2,}", "\n", texto)
    texto = re.sub(r"[ \t]+", " ", texto).strip()
    texto_up = texto.upper()
    # rango fechas
    m = re.search(r"(\d{2}/\d{2}/\d{4})\s+AL\s+(\d{2}/\d{2}/\d{4})", texto_up)
    date_range = f"{m.group(1)} Al {m.group(2)}" if m else "Fecha no disponible"
    # eliminar pie si aparece repetido
    if "TOTAL DOCUMENTOS" in texto_up:
        texto = texto_up.split("TOTAL DOCUMENTOS")[0]
        texto_up = texto.upper()
    # buscar códigos
    codigos = list(re.finditer(r"\d{4}USC[- ]?\d+", texto_up))
    if not codigos:
        # fallback: intentar códigos genéricos (menos ideal)
        codigos = list(re.finditer(r"[A-Z0-9\-]{8,}", texto_up))
    posiciones = [m.start() for m in codigos]
    codigos_str = [m.group(0).replace(" ", "").replace("-", "") if m else "SIN_CODIGO" for m in codigos]
    if not posiciones:
        return pd.DataFrame(), set(), date_range, "OIST MINJUSDH"
    posiciones.append(len(texto_up))
    registros = []
    nombres_extraidos = set()
    vistos = set()
    for i in range(len(posiciones)-1):
        blo = texto[posiciones[i]:posiciones[i+1]].strip()
        blo_up = blo.upper()
        codigo_match = re.search(r"\d{4}USC[- ]?\d+", blo_up)
        codigo = codigo_match.group(0).replace(" ", "").replace("-", "") if codigo_match else "SIN_CODIGO"
        # evitar duplicados exactos por codigo
        if codigo != "SIN_CODIGO" and codigo in vistos:
            continue
        vistos.add(codigo)
        # estado
        estado = "PENDIENTE"
        if "POR RECIBIR" in blo_up:
            estado = "POR RECIBIR"
        elif "PENDIENTE" in blo_up:
            estado = "PENDIENTE"
        # fecha preferente (FECHA INGRESO or first date)
        m_fecha_ing = re.search(r"FECHA INGRESO\s*:\s*(\d{2}/\d{2}/\d{4}\s*\d{1,2}:\d{2})", blo, flags=re.IGNORECASE)
        if not m_fecha_ing:
            m_fecha_ing = re.search(r"(\d{2}/\d{2}/\d{4})\s*(\d{1,2}:\d{2})", blo)
        fecha_str = m_fecha_ing.group(0) if m_fecha_ing else datetime.now().strftime("%d/%m/%Y %H:%M")
        # nombre (buscar DESTINO, sino buscar linea final con '/ APELLIDO, NOMBRE')
        nombre = "N/A"
        m_dest = re.search(r"DESTINO\s*:\s*(.*)", blo, flags=re.IGNORECASE)
        if m_dest:
            linea = m_dest.group(1).strip()
            # si el destino ocupa más líneas, intenta limpiar
            if "/" in linea:
                posible = linea.split("/")[-1].strip()
            else:
                posible = linea
            posible = re.sub(r"EN LAS ENTIDADES.*", "", posible, flags=re.IGNORECASE).strip()
            if "," in posible:
                a,b = [x.strip() for x in posible.split(",",1)]
                nombre = f"{b} {a}".title()
            else:
                nombre = posible.title()
        else:
            # buscar patrón / APELLIDO, NOMBRE en las últimas 3 líneas
            lines = [ln.strip() for ln in blo.split("\n") if ln.strip()]
            tail = "\n".join(lines[-4:])
            m_tail = re.search(r"/\s*([A-ZÁÉÍÓÚÑ\s,]+)", tail, flags=re.IGNORECASE)
            if m_tail:
                posible = m_tail.group(1).strip()
                if "," in posible:
                    a,b = [x.strip() for x in posible.split(",",1)]
                    nombre = f"{b} {a}".title()
                else:
                    nombre = posible.title()
        registros.append({
            "Estado": estado,
            "Fecha_Recepcion_Str": fecha_str,
            "Codigo": codigo,
            "Nombre_Personal": nombre,
            "Asunto": blo
        })
        if nombre and nombre.upper() not in ("N/A",""):
            nombres_extraidos.add(nombre.upper())
    df = pd.DataFrame(registros, columns=["Estado","Fecha_Recepcion_Str","Codigo","Nombre_Personal","Asunto"])
    return df, nombres_extraidos, date_range, "OIST MINJUSDH"

# generar excel
def procesar_y_generar_excel(pdf_bytes, fecha_hoy):
    df_raw, nombres, date_range, office = extraer_datos_pdf(pdf_bytes)
    actualizar_lista_maestra(list(nombres))
    df_clean = clean_extracted_df(df_raw, fecha_hoy)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_clean.to_excel(writer, sheet_name="1. Completo", index=False)
        df_clean.sort_values(by="Dias_En_Bandeja", ascending=False).to_excel(writer, sheet_name="2. Antiguos", index=False)
        df_clean.groupby("Nombre_Personal").size().reset_index(name="Total de Documentos").to_excel(writer, sheet_name="3. Totales", index=False)
    output.seek(0)
    return output

# get dashboard data
def get_dashboard_data(pdf_bytes, fecha_hoy):
    df_raw, nombres, date_range, office = extraer_datos_pdf(pdf_bytes)
    actualizar_lista_maestra(list(nombres))
    df_clean = clean_extracted_df(df_raw, fecha_hoy)
    # ensure POR RECIBIR and PENDIENTE columns exist aggregated
    try:
        totals = df_clean.groupby('Nombre_Personal')['Estado'].value_counts().unstack(fill_value=0).reset_index()
        if 'PENDIENTE' not in totals.columns:
            totals['PENDIENTE'] = 0
        if 'POR RECIBIR' not in totals.columns:
            totals['POR RECIBIR'] = 0
        df = pd.merge(df_clean, totals[['Nombre_Personal','PENDIENTE','POR RECIBIR']], on='Nombre_Personal', how='left')
    except Exception:
        df = df_clean.copy()
        df['PENDIENTE'] = df['Estado'].eq('PENDIENTE').astype(int)
        df['POR RECIBIR'] = df['Estado'].eq('POR RECIBIR').astype(int)
    return df, date_range, office
