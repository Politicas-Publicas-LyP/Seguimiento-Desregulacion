import pandas as pd
import requests
from bs4 import BeautifulSoup
from datetime import date
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import sys
import re
import unicodedata
import time

# --- CONFIGURACIÓN DE CREDENCIALES (GitHub Secrets) ---
EMAIL_ORIGEN = os.environ.get('EMAIL_ORIGEN')
PASSWORD_APP = os.environ.get('PASSWORD_APP')
# Soporta múltiples destinatarios separados por coma:
#   EMAIL_DESTINO="correo1@gmail.com,correo2@gmail.com"
EMAIL_DESTINO = os.environ.get('EMAIL_DESTINO', '')

# El envío usa SMTP. Por defecto apunta a Gmail, pero el servidor es configurable:
# así podés usar CUALQUIER proveedor (Gmail, GMX, Zoho, etc.) sin tocar el código,
# definiendo SMTP_HOST y SMTP_PORT como variables/secrets.
#   - Gmail:  SMTP_HOST=smtp.gmail.com  (requiere contraseña de aplicación)
#   - GMX:    SMTP_HOST=mail.gmx.com    (permite la contraseña normal de la cuenta)
# EMAIL_ORIGEN = dirección remitente; PASSWORD_APP = la contraseña SMTP que pida
# el proveedor (de aplicación en Gmail, o la normal en GMX).
SMTP_HOST = os.environ.get('SMTP_HOST') or 'smtp.gmail.com'
SMTP_PORT = int(os.environ.get('SMTP_PORT') or '587')

# ─── CAPA DE IA (confirmación de coincidencias con Gemini) ───────────────────
# El filtro de keywords PRESELECCIONA candidatos; la IA confirma cuáles son
# coincidencias reales (la norma efectivamente modifica/deroga/crea lo del caso)
# y descarta las que solo comparten vocabulario. Apagada por defecto: si USAR_IA
# no es "true", el radar funciona igual que ahora (solo keywords).
USAR_IA = (os.environ.get('USAR_IA', 'false').lower() == 'true')
GEMINI_API_KEY = os.environ.get('GEMINI_API_KEY', '')
# Modelo configurable. Por defecto el mejor equilibrio dentro del tier gratuito.
# Alternativas: "gemini-3.5-flash" (más nuevo), "gemini-2.5-flash-lite" (más liviano).
GEMINI_MODEL = os.environ.get('GEMINI_MODEL') or 'gemini-2.5-flash'
GEMINI_API_BASE = "https://generativelanguage.googleapis.com/v1beta/models"
IA_DELAY_SEGUNDOS = 1.5      # pausa entre consultas para respetar el límite gratuito
IA_MAX_CHARS_NORMA = 8000    # recorte del texto de la norma que se manda a la IA

# --- CONFIGURACIÓN GENERAL ---
NOMBRE_ARCHIVO_BASE = 'seguimiento_desregulacion_estandarizado'
NOMBRE_HOJA_EXCEL = 'Radar'

# ─── FUENTE DE LA BASE DE DATOS ──────────────────────────────────────────────
# "excel_local"  → lee el archivo .xlsx del repositorio (comportamiento original).
# "google_sheets" → lee la planilla en la nube (se actualiza sin re-subir el archivo).
# Cuando termines el setup de Google, cambiá esto a "google_sheets" (o definí el
# secret/variable FUENTE_DATOS en GitHub).
FUENTE_DATOS = os.environ.get('FUENTE_DATOS', 'excel_local')

# ID de la planilla de Google (está en la URL:
#   https://docs.google.com/spreadsheets/d/ESTE_ES_EL_ID/edit ).
GOOGLE_SHEET_ID = os.environ.get('GOOGLE_SHEET_ID', '')

# Contenido del JSON de la cuenta de servicio (se carga como secret en GitHub,
# pegando el archivo .json completo como valor del secret GOOGLE_CREDENTIALS_JSON).
GOOGLE_CREDENTIALS_JSON = os.environ.get('GOOGLE_CREDENTIALS_JSON', '')
URL_BORA_BASE = "https://www.boletinoficial.gob.ar"

# Solo escaneamos la Primera Sección (Decretos, Resoluciones, Leyes, etc.)
URL_PRIMERA_SECCION = f"{URL_BORA_BASE}/seccion/primera"

# Incluir listado completo de normas del día en el email (sin IA, solo títulos y links)
INCLUIR_LISTADO_BOLETIN = True

# Umbral de score para considerar un match (bajado de 3.5 a 2.5 para no perder coincidencias)
SCORE_MINIMO = 2.5

# Estados que se consideran INACTIVOS (no se cruzan contra el BORA).
# Cualquier otro estado —incluido 'PENDIENTE' o vacío— se trata como ACTIVO.
ESTADOS_INACTIVOS = {"INACTIVO", "BAJA", "CERRADO", "ARCHIVADO", "DESCARTADO"}

# Si el escaneo del BORA devuelve menos normas que este mínimo, se asume que algo
# falló (bloqueo de IP, HTML incompleto) y se dispara una alerta en lugar de
# enviar un "sin novedades" engañoso.
MINIMO_NORMAS_ESPERADAS = 5

# Si en un día aparecen más coincidencias que esto, casi seguro hay una regresión
# (keywords genéricas, ruido). Se avisa para recalibrar en vez de naturalizar una
# avalancha de alertas (premortem: avalancha de ruido).
MAX_ALERTAS_RAZONABLE = 25

# ─── FASE 2: LECTURA DE TEXTO COMPLETO ───────────────────────────────────────
# Si es True, el radar entra a la ficha de cada norma sustantiva y matchea contra
# el TEXTO COMPLETO (título + resumen + cuerpo), no solo el título del índice.
# Poner en False para volver al comportamiento anterior (solo títulos).
LEER_TEXTO_COMPLETO = True

# Pausa (segundos) entre cada pedido a la ficha de una norma, para no saturar el
# BORA ni que bloqueen la IP de GitHub Actions (premortem: fallo de throttling).
BORA_DELAY_SEGUNDOS = 1.2

# Si el cuerpo de una norma trae menos caracteres que esto, se asume lectura fallida
# (shell vacío / HTML incompleto) y se cae al título + resumen como respaldo.
MIN_CARACTERES_CUERPO = 250

# Si más de esta fracción de las normas no se pudieron leer, se dispara alerta:
# probablemente el BORA cambió el HTML o está bloqueando (fallar ruidoso).
MAX_FRACCION_FALLOS_LECTURA = 0.5

# Umbral de score para texto completo.
SCORE_MINIMO_TEXTO = 3.0

# Reglas anti-falsos-positivos para texto completo. Una coincidencia exige:
#  - score por encima del umbral, y
#  - una FRASE multipalabra O al menos 2 keywords DISTINTIVAS (términos
#    específicos, no genéricos ni del armazón del BORA).
# Así un texto largo no dispara alerta por compartir una o dos palabras comunes
# con un caso (premortem: avalancha de ruido). Calibrado con normas reales del
# 09/06/2026: bajó de 142 coincidencias a ~1 por norma, conservando las correctas.
EXIGIR_KEYWORD_DISTINTIVA = True
LONGITUD_KEYWORD_DISTINTIVA = 7   # palabras sueltas con 7+ letras pueden ser distintivas
MIN_KEYWORDS_DISTINTIVAS = 2      # mínimo de keywords distintivas (si no hay frase)

# --- FILTROS ---
# Normas de personal / actos administrativos rutinarios: si el TÍTULO contiene
# alguno de estos términos, la norma se ignora (no es una desregulación).
EXCLUSIONES_RRHH = [
    "DESIGNASE", "DASE POR ASIGNADA", "DASE POR PRORROGADA",
    "PRORROGASE", "ACEPTASE LA RENUNCIA", "CONTRATACIONES",
    "LICENCIA", "TRASLADO", "NOMBRAMIENTO", "CESE",
    # Personal militar/fuerzas de seguridad y otros actos de personal.
    "PROMOCIONES", "ASCENSO", "ASCENSOS", "CONDECORACION", "CONDECORACIONES",
    "INCORPORACION", "INCORPORASE", "INCORPORANSE", "PASE A RETIRO",
    "RETIRO", "RETIROS", "HABER DE RETIRO", "JUBILACION", "RECONOCIMIENTO DE SERVICIOS",
    "DISTINCIONES", "FELICITACIONES",
]

# Tipos de norma que NO se cruzan contra la base (rutinarios: edictos, balances,
# convocatorias, notificaciones). Los avisos oficiales casi nunca son una
# desregulación, pero hacen ruido porque el organismo nombrado coincide con un caso.
TIPOS_NO_MATCHEABLES = {"Aviso"}

# Stoplist (palabras de ruido). Incluye el ARMAZÓN del sitio del BORA (menú, pie)
# y vocabulario legal/administrativo UBICUO que aparece en casi toda norma. NO
# incluye términos de dominio (VENTANILLA, COMERCIO, IMPORTACION, ARANCEL,
# JUGUETES, ARBITRAJE...) que sí distinguen un caso. Una palabra acá vale 0 y
# nunca cuenta como distintiva.
PALABRAS_RUIDO = {
    # Originales
    "TODAS", "TODOS", "OPTAR", "UNICO", "UNICA", "FIJOS", "BANCO",
    "NACION", "DECRETO", "DECRETOS", "NORMA", "NORMAS", "NORMATIVA", "NORMATIVAS",
    "ACTO", "ACTOS", "ENTIDADES", "ORGANISMOS", "ORGANISMO", "PERSONAL", "GENERAL",
    "NACIONAL", "NACIONALES", "ARGENTINA", "ARGENTINAS", "ARGENTINO", "ARGENTINOS",
    "CADA", "SERA", "DICHO", "DICHA", "PARTE", "PARTES",
    # Conectores, adjetivos y verbos genéricos
    "SERIA", "CONVENIENTE", "CONVENDRIA", "CONVENDRÍA", "INADECUADO", "INADECUADA",
    "ENCIMA", "CUANDO", "ENTRE", "MENOR", "MAYOR", "AMPLIAR", "BAJA", "PERMITE",
    "ELIMINAR", "ELIMINA", "DEROGAR", "DEROGACION", "DEROGASE", "MODIFICAR",
    "MODIFICACION", "MODIFICACIONES", "REQUIERE", "EXIGE", "PROHIBE", "PROHIBIR",
    "CREANDO", "FACILITE", "INCENTIVE", "ENCARECEN", "ENCARECIENDO", "ESTABLECIO",
    "ESTABLECESE", "ESTABLECIDO", "FIGURA", "REALIZAR", "ENCUENTRA", "MODELO",
    "INCISO", "VENCIDO", "PREVIA", "RESPECTO", "APLICACION",
    # Vocabulario legal/administrativo ubicuo (aparece en casi toda norma)
    "OBLIGACION", "ESTADO", "SISTEMA", "REGIMEN", "ANEXO", "REFORMA", "REGISTRO",
    "INFORMACION", "AGENCIA", "CONDICIONES", "PLAZOS", "COSTOS", "SERIE", "SERIAN",
    "DEBERIA", "PODRIA", "CONVENIENTES",
    "RESOLUCION", "RESOLUCIONES", "OFICIAL", "BOLETIN", "SECCION", "DERECHOS",
    "HUMANOS", "AUTORIDADES", "AUTORIDAD", "MEDIDAS", "MEDIDA", "DIRECCION",
    "DIRECCIONES", "MINISTERIO", "SECRETARIA", "GESTION", "RECURSOS", "CIUDADANOS",
    "JURIDICA", "JURIDICO", "VIGENTE", "VIGENCIA", "MARCO", "REGLAMENTO",
    "REGLAMENTARIA", "DISPOSICION", "DISPOSICIONES", "LEY", "LEYES", "EXPEDIENTE",
    "PUBLICO", "PUBLICA", "PUBLICOS", "PUBLICAS", "SERVICIOS", "SERVICIO",
    "INTERVENCION", "ARTICULOS", "ARTICULO", "PROCEDIMIENTO", "PROCEDIMIENTOS",
    "PERSONA", "PERSONAS", "FISICAS", "PRESTADORES", "TITULARES", "PRIVADO",
    "PRIVADA", "IMPLEMENTACION", "CONTRATO", "DESARROLLO", "INVESTIGACION",
    "SEGURIDAD", "APARTADO", "TECNICO", "TECNICOS", "TECNICA", "ESTUDIOS",
    "ESTUDIO", "DESPACHO", "ASIGNACION", "INFRAESTRUCTURA", "TERCEROS",
}

# Siglas cortas que sí son relevantes y no deben descartarse
SIGLAS_PERMITIDAS = {
    "DNU", "IVA", "UIF", "ANR", "SRT", "ART", "CNV", "UBA",
    "IGJ", "SSN", "SSS", "SAS", "BCE", "FCI", "PPP",
    "EIA", "UVA", "CEA", "APN", "CFI", "BNA", "YPF",
    "AGN", "PIB", "PBI", "ONP", "AFI", "FMI", "BCRA",
    "AFIP", "ARCA", "ANAC", "EANA", "ENACOM", "ANMAT",
    "SENASA", "ENRE", "CNRT", "ORSNA", "ENARGAS",
    # Siglas y códigos específicos de los casos de desregulación seguidos.
    "VUCE", "VUCEA", "CUD", "DDP", "BK", "BIT", "ZAP", "CIBU", "OMS",
    "OACI", "NCM", "LCM", "RENPI", "RENFO", "INADI", "INAES", "INET",
    "UNDEF", "INPRES", "CONETEC", "CONICET", "ARICCAME", "ENOHSA",
    "MIPYME", "PSAV", "ENOSHA", "INAFCI",
}

# --- Tipos de norma y sus patrones de detección ---
TIPOS_NORMA = [
    ("Decreto de Necesidad y Urgencia", [r"\bDNU\b", r"\bDECRETO DE NECESIDAD Y URGENCIA\b"]),
    ("Decreto",                         [r"\bDECRETO\b"]),
    ("Decisión Administrativa",         [r"\bDECISION ADMINISTRATIVA\b"]),
    ("Resolución Conjunta",             [r"\bRESOLUCION CONJUNTA\b"]),
    ("Resolución General",              [r"\bRESOLUCION GENERAL\b"]),
    ("Resolución",                      [r"\bRESOLUCION\b"]),
    ("Disposición",                     [r"\bDISPOSICION\b"]),
    ("Comunicación",                    [r"\bCOMUNICACION\b"]),
    ("Aviso",                           [r"\bAVISO\b"]),
    ("Ley",                             [r"\bLEY\b"]),
    ("Acordada",                        [r"\bACORDADA\b"]),
]

COLORES_TIPO = {
    "Decreto de Necesidad y Urgencia": ("#c0392b", "#fdecea"),
    "Decreto":                         ("#2c3e50", "#eaf2f8"),
    "Decisión Administrativa":         ("#8e44ad", "#f4ecf7"),
    "Resolución Conjunta":             ("#2980b9", "#ebf5fb"),
    "Resolución General":              ("#2980b9", "#ebf5fb"),
    "Resolución":                      ("#2471a3", "#d6eaf8"),
    "Disposición":                     ("#27ae60", "#eafaf1"),
    "Comunicación":                    ("#f39c12", "#fef9e7"),
    "Aviso":                           ("#95a5a6", "#f2f4f4"),
    "Ley":                             ("#c0392b", "#fdedec"),
    "Acordada":                        ("#7d3c98", "#f5eef8"),
    "Otros":                           ("#7f8c8d", "#f2f3f4"),
    "Personal (RRHH)":                 ("#bdc3c7", "#f8f9f9"),
}

ORDEN_TIPOS = [
    "Ley", "Decreto de Necesidad y Urgencia", "Decreto",
    "Decisión Administrativa", "Resolución Conjunta",
    "Resolución General", "Resolución", "Disposición",
    "Acordada", "Comunicación", "Aviso", "Otros", "Personal (RRHH)"
]


# ═══════════════════════════════════════════════════════════════════════════════
# FUNCIONES AUXILIARES
# ═══════════════════════════════════════════════════════════════════════════════

def normalizar_texto(texto):
    if not isinstance(texto, str):
        return ""
    texto = ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )
    texto = re.sub(r'[^A-Z0-9\s]', ' ', texto.upper())
    return re.sub(r'\s+', ' ', texto).strip()


def peso_keyword(keyword: str) -> float:
    k = keyword.strip().upper()
    if k in PALABRAS_RUIDO:
        return 0.0
    # Frase de varias palabras: señal muy fuerte (p. ej. "CONDICIONES HABILITANTES",
    # "CODIGO CIVIL Y COMERCIAL", "CONVENCIONES COLECTIVAS").
    if ' ' in k:
        return 3.0
    # Siglas relevantes (CONICET, VUCE, ARICCAME, ENARGAS, ...).
    if k in SIGLAS_PERMITIDAS:
        return 1.5
    n = len(k)
    if n <= 4:
        return 0.0
    elif n <= 6:
        return 0.8
    elif n <= 9:
        return 1.5
    elif n <= 12:
        return 2.0
    else:
        return 2.6


def es_distintiva(keyword: str) -> bool:
    """Una keyword es DISTINTIVA si es una frase de varias palabras o un término
    específico y largo. Las genéricas (ruido, palabras cortas) no lo son.
    Sirve para exigir al menos una señal fuerte antes de declarar coincidencia."""
    k = keyword.strip().upper()
    if not k or k in PALABRAS_RUIDO:
        return False
    if ' ' in k:
        return True
    if k in SIGLAS_PERMITIDAS:
        return True
    return len(k) >= LONGITUD_KEYWORD_DISTINTIVA


def es_keyword_valida(keyword: str) -> bool:
    """Determina si una keyword debe incluirse en el matching."""
    k = keyword.strip().upper()
    if not k:
        return False
    if len(k) <= 3:
        return k in SIGLAS_PERMITIDAS
    return True


def get_con_reintentos(url, headers, timeout=15, intentos=3, espera=5):
    for intento in range(1, intentos + 1):
        try:
            r = requests.get(url, headers=headers, timeout=timeout)
            if r.status_code < 500:
                return r
            print(f"  ⚠️  Status {r.status_code} en intento {intento}/{intentos} — {url}")
        except requests.exceptions.RequestException as e:
            print(f"  ⚠️  Error de red en intento {intento}/{intentos}: {e}")
        if intento < intentos:
            time.sleep(espera)
    print(f"  ❌ Todos los intentos fallaron para {url}")
    return None


def obtener_texto_norma(url, headers):
    """Entra a la ficha de una norma y devuelve (resumen_oficial, cuerpo_texto).

    - resumen_oficial: el meta-description del BORA (siempre server-rendered;
      contiene organismo + tipo + título extendido de la norma).
    - cuerpo_texto: el texto visible de la página tras quitar menús y pie.
    Si la lectura falla, devuelve ('', '') y quien llama cae al título como respaldo.
    """
    r = get_con_reintentos(url, headers, timeout=20, intentos=2, espera=4)
    if r is None or r.status_code != 200:
        return "", ""

    try:
        soup = BeautifulSoup(r.text, 'html.parser')
        resumen = ""
        meta = soup.find('meta', attrs={'name': 'description'})
        if meta and meta.get('content'):
            resumen = meta['content']
        for tag in soup(['script', 'style', 'nav', 'footer', 'header', 'noscript']):
            tag.decompose()
        texto_pagina = soup.get_text(" ", strip=True)

        # CRÍTICO: recortar SOLO el cuerpo del aviso. La página del BORA trae un
        # armazón (menú "Boletín Oficial", pie con "Ministerio de Justicia y
        # Derechos Humanos", etc.) que, si se matchea, genera falsos positivos en
        # TODAS las normas. El cuerpo real va entre "Ver texto del aviso" y el
        # widget "Compartir por email" (o "Fecha de publicación" como respaldo).
        ini = texto_pagina.find("Ver texto del aviso")
        ini = ini + len("Ver texto del aviso") if ini != -1 else 0
        fin = texto_pagina.find("Compartir por email", ini)
        if fin == -1:
            m = re.search(r"Fecha de publicaci", texto_pagina[ini:])
            fin = ini + m.start() if m else len(texto_pagina)
        cuerpo = texto_pagina[ini:fin].strip()
        return resumen, cuerpo
    except Exception as e:
        print(f"  ⚠️  Error al parsear la ficha {url}: {e}")
        return "", ""


def clasificar_tipo_norma(titulo_normalizado: str) -> str:
    for nombre_tipo, patrones in TIPOS_NORMA:
        for patron in patrones:
            if re.search(patron, titulo_normalizado):
                return nombre_tipo
    return "Otros"


def es_norma_de_rrhh(titulo_normalizado: str) -> bool:
    return any(excl in titulo_normalizado for excl in EXCLUSIONES_RRHH)


def parsear_destinatarios(email_destino_raw: str) -> list:
    """Convierte la variable EMAIL_DESTINO en una lista de direcciones."""
    if not email_destino_raw:
        return []
    return [e.strip() for e in email_destino_raw.split(',') if e.strip()]


# ═══════════════════════════════════════════════════════════════════════════════
# MÓDULO 1: CARGA DE LA BASE (Excel local o Google Sheets)
# ═══════════════════════════════════════════════════════════════════════════════

def leer_base_excel_local():
    """Lee la hoja Radar desde el archivo .xlsx del repositorio."""
    archivo_excel = f"{NOMBRE_ARCHIVO_BASE}.xlsx"
    if not os.path.exists(archivo_excel):
        print(f"❌ Error: No se encontró el archivo '{archivo_excel}'.")
        sys.exit(1)
    try:
        df = pd.read_excel(archivo_excel, sheet_name=NOMBRE_HOJA_EXCEL)
        print(f"📄 Base leída desde Excel local: {len(df)} filas.")
        return df
    except Exception as e:
        print(f"❌ Error al leer el Excel: {e}")
        enviar_alerta_error("No se pudo leer el Excel local", str(e))
        sys.exit(1)


def leer_base_google_sheets():
    """Lee la hoja Radar desde Google Sheets con una cuenta de servicio.
    Requiere los secrets GOOGLE_SHEET_ID y GOOGLE_CREDENTIALS_JSON."""
    if not GOOGLE_SHEET_ID or not GOOGLE_CREDENTIALS_JSON:
        msg = ("Falta GOOGLE_SHEET_ID y/o GOOGLE_CREDENTIALS_JSON. No se puede leer "
               "la base desde Google Sheets. Cargá ambos secrets en GitHub o volvé a "
               "FUENTE_DATOS='excel_local'.")
        print(f"❌ {msg}")
        enviar_alerta_error("Faltan credenciales de Google Sheets", msg)
        sys.exit(1)
    try:
        import json
        import gspread
        from google.oauth2.service_account import Credentials
        info = json.loads(GOOGLE_CREDENTIALS_JSON)
        scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
        creds = Credentials.from_service_account_info(info, scopes=scopes)
        gc = gspread.authorize(creds)
        ws = gc.open_by_key(GOOGLE_SHEET_ID).worksheet(NOMBRE_HOJA_EXCEL)
        df = pd.DataFrame(ws.get_all_records())  # fila 1 = encabezados
        if df.empty:
            raise ValueError("La hoja se leyó pero no tiene filas de datos.")
        print(f"☁️  Base leída desde Google Sheets: {len(df)} filas.")
        return df
    except Exception as e:
        msg = (f"No se pudo leer la base desde Google Sheets: {e}. Verificá que la "
               f"planilla esté compartida con el email de la cuenta de servicio, que el "
               f"ID y el nombre de hoja ('{NOMBRE_HOJA_EXCEL}') sean correctos, y que la "
               f"API de Google Sheets esté habilitada.")
        print(f"❌ {msg}")
        enviar_alerta_error("Error leyendo Google Sheets", msg)
        sys.exit(1)


def cargar_archivo_robusto():
    # Elegir la fuente según la configuración (no rompe nada: por defecto Excel local).
    if FUENTE_DATOS == 'google_sheets':
        print("🔗 Fuente de datos: Google Sheets")
        df = leer_base_google_sheets()
    else:
        print("🔗 Fuente de datos: Excel local")
        df = leer_base_excel_local()

    df.columns = [c.upper().strip() for c in df.columns]
    mapeo = {
        'ACCION_ESPERADA': 'ACCION',
        'ACCION ESPERADA': 'ACCION',
        'PALABRAS CLAVE': 'PALABRAS_CLAVE'
    }
    df = df.rename(columns=mapeo)

    if 'PALABRAS_CLAVE' not in df.columns:
        print("❌ Error: No se encontró la columna PALABRAS_CLAVE.")
        sys.exit(1)

    df['PALABRAS_CLAVE'] = df['PALABRAS_CLAVE'].fillna('').astype(str).str.upper()
    df['ACCION'] = df['ACCION'].fillna('REVISAR').astype(str).str.upper() if 'ACCION' in df.columns else 'REVISAR'

    if 'ID_CASO' not in df.columns:
        df['ID_CASO'] = [f"CASO-{i+1}" for i in range(len(df))]

    # --- Filtro de estado (CORREGIDO) ---
    # Antes el código activaba solo ESTADO == 'ACTIVO', pero en la base los estados
    # reales son 'Pendiente' e 'Inactivo': ninguno coincidía con 'ACTIVO', de modo
    # que se cruzaban apenas las filas con estado vacío. Ahora se ACTIVA todo lo que
    # no esté explícitamente inactivo (los 'Pendiente' y los vacíos pasan a activos).
    if 'ESTADO' in df.columns:
        df['ESTADO'] = df['ESTADO'].fillna('PENDIENTE').astype(str).str.upper().str.strip()
        total_antes = len(df)
        df = df[~df['ESTADO'].isin(ESTADOS_INACTIVOS)]
        excluidos = total_antes - len(df)
        print(f"   Estados: {total_antes} filas en la base → {len(df)} activas "
              f"({excluidos} inactivas excluidas).")
    else:
        print("   ⚠️  No hay columna ESTADO: se consideran activos todos los casos.")

    # --- Alarma de base vacía (fallar ruidoso, no silencioso) ---
    if len(df) == 0:
        msg = ("La base de datos quedó con 0 casos activos tras aplicar los filtros. "
               "Esto suele indicar un cambio en los nombres de columnas, en los valores "
               "de la columna ESTADO, o un archivo vacío. El radar se detiene para no "
               "enviar un 'sin coincidencias' engañoso.")
        print(f"❌ {msg}")
        enviar_alerta_error("Base de datos vacía (0 casos activos)", msg)
        sys.exit(1)

    print(f"✅ Base de datos cargada: {len(df)} casos activos.")
    return df


# ═══════════════════════════════════════════════════════════════════════════════
# MÓDULO 2: ESCANEO DEL BORA (solo Primera Sección)
# ═══════════════════════════════════════════════════════════════════════════════

def escanear_boletin():
    fecha_hoy = date.today().strftime('%d/%m/%Y')
    # User-Agent realista + headers de navegador: reduce timeouts y bloqueos del BORA.
    headers = {
        'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                       '(KHTML, like Gecko) Chrome/124.0 Safari/537.36'),
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'es-AR,es;q=0.9',
    }
    normas = []
    urls_vistas = set()

    print(f"📡 Escaneando Primera Sección del BORA ({fecha_hoy})...")
    # Más tiempo y más reintentos: el timeout suele ser transitorio desde GitHub.
    r = get_con_reintentos(URL_PRIMERA_SECCION, headers, timeout=30, intentos=4, espera=8)

    if r is None or r.status_code != 200:
        print("  ⚠️  No se pudo acceder a la Primera Sección del BORA.")
        return normas, False   # leido_ok=False: no se pudo leer el BORA

    soup = BeautifulSoup(r.text, 'html.parser')

    for link in soup.find_all('a', href=True):
        href = link['href']
        texto = link.get_text(" ", strip=True).upper()
        if 'detalleAviso' in href and len(texto) > 10:
            full_url = URL_BORA_BASE + href if href.startswith('/') else href
            if full_url not in urls_vistas:
                urls_vistas.add(full_url)
                titulo_norm = normalizar_texto(texto)
                normas.append({
                    'TEXTO': texto,
                    'TEXTO_NORM': titulo_norm,
                    'TEXTO_MATCH': titulo_norm,   # por defecto = título; se enriquece abajo
                    'TEXTO_ORIGINAL': texto,      # texto sin normalizar, para la IA
                    'URL': full_url,
                    'TIPO': clasificar_tipo_norma(titulo_norm),
                    'ES_RRHH': es_norma_de_rrhh(titulo_norm),
                })

    print(f"  ✅ {len(normas)} normas encontradas en la Primera Sección.")

    # --- FASE 2: leer el texto completo de cada norma sustantiva ---
    if LEER_TEXTO_COMPLETO:
        sustantivas = [n for n in normas if not n['ES_RRHH']]
        print(f"  📖 Leyendo el texto completo de {len(sustantivas)} normas "
              f"sustantivas (pausa de {BORA_DELAY_SEGUNDOS}s entre cada una)...")
        fallos = 0
        for i, norma in enumerate(sustantivas, 1):
            resumen, cuerpo = obtener_texto_norma(norma['URL'], headers)
            lectura_ok = len(cuerpo) >= MIN_CARACTERES_CUERPO
            if not lectura_ok:
                fallos += 1
            # El texto a matchear combina título + resumen oficial + cuerpo.
            # Aunque falle el cuerpo, el resumen ya mejora respecto al título solo.
            combinado = f"{norma['TEXTO']} {resumen} {cuerpo}"
            norma['TEXTO_MATCH'] = normalizar_texto(combinado)
            norma['TEXTO_ORIGINAL'] = combinado   # sin normalizar, para la IA
            if i % 15 == 0:
                print(f"     ... {i}/{len(sustantivas)} leídas ({fallos} fallidas)")
            time.sleep(BORA_DELAY_SEGUNDOS)

        total = max(1, len(sustantivas))
        frac_fallos = fallos / total
        print(f"  ✅ Texto completo leído. Fallos de lectura: {fallos}/{total} "
              f"({frac_fallos:.0%}).")
        if frac_fallos > MAX_FRACCION_FALLOS_LECTURA:
            msg = (f"No se pudo leer el cuerpo de {fallos} de {total} normas "
                   f"({frac_fallos:.0%}). Posible bloqueo del BORA o cambio en el HTML "
                   f"de las fichas. Las coincidencias de hoy pueden ser incompletas: "
                   f"conviene revisar manualmente.")
            print(f"  ⚠️  {msg}")
            enviar_alerta_error("Lectura de texto completo degradada", msg)

    return normas, True   # leido_ok=True: el índice del BORA se leyó correctamente


# ═══════════════════════════════════════════════════════════════════════════════
# MÓDULO 3: MATCHING CON LA BASE DE DATOS
# ═══════════════════════════════════════════════════════════════════════════════

def cruzar_con_base(normas, df):
    alertas = []
    normas_ignoradas = 0

    for norma in normas:
        if norma['ES_RRHH']:
            normas_ignoradas += 1
            continue

        # Saltar avisos oficiales y otros tipos rutinarios (no son desregulaciones;
        # solo generan ruido porque el organismo emisor coincide con un caso).
        if norma.get('TIPO') in TIPOS_NO_MATCHEABLES:
            continue

        # Matchear contra el texto completo si está disponible; si no, contra el título.
        texto_norma = norma.get('TEXTO_MATCH') or norma['TEXTO_NORM']
        umbral = SCORE_MINIMO_TEXTO if LEER_TEXTO_COMPLETO else SCORE_MINIMO

        for _, row in df.iterrows():
            id_caso = row['ID_CASO']
            accion = row['ACCION']

            keywords_crudas = str(row['PALABRAS_CLAVE']).replace(';', ',').split(',')
            keywords = [normalizar_texto(k) for k in keywords_crudas if es_keyword_valida(k)]
            if not keywords:
                continue

            score_total = 0.0
            keywords_encontradas = []
            distintivas_encontradas = []

            for k in keywords:
                patron = r'\b' + re.escape(k) + r'\b'
                if re.search(patron, texto_norma):
                    w = peso_keyword(k)
                    if w > 0:
                        score_total += w
                        keywords_encontradas.append(k)
                        if es_distintiva(k):
                            distintivas_encontradas.append(k)

            # Reglas anti-falsos-positivos:
            # 1) el score debe superar el umbral, y
            # 2) debe coincidir una FRASE multipalabra O al menos
            #    MIN_KEYWORDS_DISTINTIVAS keywords distintivas (términos
            #    específicos). Compartir una o dos palabras genéricas con el caso
            #    ya no alcanza para disparar una alerta.
            tiene_frase = any(' ' in k for k in keywords_encontradas)
            if score_total < umbral:
                continue
            if EXIGIR_KEYWORD_DISTINTIVA and not (
                tiene_frase or len(distintivas_encontradas) >= MIN_KEYWORDS_DISTINTIVAS
            ):
                continue

            ya_alertado = any(
                a['ID'] == id_caso and a['URL'] == norma['URL']
                for a in alertas
            )
            if ya_alertado:
                continue

            print(
                f"  🎯 Match: {id_caso} | score={score_total:.1f} "
                f"| distintivas={distintivas_encontradas} | todas={keywords_encontradas}"
            )

            alertas.append({
                'ID': id_caso,
                'ACCION': accion,
                'URL': norma['URL'],
                'TITULO': norma['TEXTO'][:150],
                'SCORE': score_total,
                'KEYWORDS_ENCONTRADAS': keywords_encontradas,
                'KEYWORDS_DISTINTIVAS': distintivas_encontradas,
                'TIPO': norma['TIPO'],
                # Contexto para la capa de IA:
                'ORGANISMO': str(row.get('ORGANISMO', '') or ''),
                'NUMERO_NORMA': str(row.get('NUMERO_NORMA', '') or ''),
                'TEXTO_NORMA': norma.get('TEXTO_ORIGINAL', norma['TEXTO']),
            })

    return alertas, normas_ignoradas


# ═══════════════════════════════════════════════════════════════════════════════
# MÓDULO 3-BIS: CONFIRMACIÓN CON IA (Gemini)
# ═══════════════════════════════════════════════════════════════════════════════

def _consultar_gemini(prompt: str) -> str:
    """Hace una consulta a la API de Gemini y devuelve el texto de la respuesta.
    Lanza una excepción con el motivo concreto si algo falla (para diagnóstico)."""
    url = f"{GEMINI_API_BASE}/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
    generation_config = {"temperature": 0, "maxOutputTokens": 600}
    # Los modelos 2.5/3 "piensan" por defecto y se comen el presupuesto de tokens
    # dejando la respuesta vacía. Desactivamos ese pensamiento para esta tarea simple.
    if any(s in GEMINI_MODEL for s in ("2.5", "3.")) or "flash-lite" in GEMINI_MODEL:
        generation_config["thinkingConfig"] = {"thinkingBudget": 0}
    body = {"contents": [{"parts": [{"text": prompt}]}], "generationConfig": generation_config}

    ultimo_error = None
    for intento in range(1, 3):  # 1 reintento ante fallos transitorios (DNS/red)
        try:
            r = requests.post(url, json=body, timeout=40)
            if r.status_code != 200:
                # Surface el mensaje real de la API (key inválida, modelo inexistente, cuota…)
                raise RuntimeError(f"HTTP {r.status_code}: {r.text[:200]}")
            data = r.json()
            cands = data.get("candidates") or []
            if not cands:
                raise RuntimeError(f"sin candidates: {str(data)[:200]}")
            parts = (cands[0].get("content") or {}).get("parts") or []
            textos = [p.get("text", "") for p in parts if p.get("text")]
            if not textos:
                raise RuntimeError(f"respuesta sin texto (finishReason="
                                   f"{cands[0].get('finishReason', '?')})")
            return " ".join(textos)
        except Exception as e:
            ultimo_error = e
            if intento < 2:
                time.sleep(3)
    raise ultimo_error


def confirmar_con_ia(id_caso, accion, desc_caso, texto_norma):
    """Le pregunta a Gemini si la norma se refiere ESPECÍFICAMENTE al caso.
    Devuelve (coincide, razon). coincide es True/False, o None si la IA falló."""
    import json
    prompt = f"""Sos analista de políticas públicas de la Fundación Libertad y Progreso. \
Seguimos una lista de normas que proponemos desregular (modificar, derogar, eliminar o reglamentar).

Te paso UN caso que seguimos y el TEXTO de una norma publicada hoy en el Boletín Oficial. \
Decidí si la norma del Boletín se refiere ESPECÍFICAMENTE a la norma, el organismo o el tema del caso \
(la modifica, deroga, reglamenta, crea o disuelve ese organismo, o cambia ese régimen). \
Que solo compartan palabras sueltas o temas generales NO alcanza: tiene que ser sobre lo mismo.

CASO {id_caso} (acción que proponemos: {accion}):
{desc_caso}

NORMA DEL BOLETÍN:
{texto_norma[:IA_MAX_CHARS_NORMA]}

Respondé SOLO con un JSON válido, sin texto adicional:
{{"coincide": true, "razon": "<una frase breve>"}}  ó  {{"coincide": false, "razon": "<una frase breve>"}}"""
    try:
        txt = _consultar_gemini(prompt).strip()
        m = re.search(r'\{.*\}', txt, re.DOTALL)
        if not m:
            # Respuesta sin JSON: incierto → None (se conserva para revisión manual).
            return None, f"respuesta no interpretable: {txt[:120]}"
        obj = json.loads(m.group(0))
        # Solo descartamos ante un 'false' EXPLÍCITO; cualquier otra cosa = incierto.
        if 'coincide' not in obj:
            return None, f"sin campo coincide: {txt[:120]}"
        return bool(obj.get('coincide')), str(obj.get('razon', ''))[:300]
    except Exception as e:
        return None, f"error: {e}"


def confirmar_alertas_con_ia(alertas):
    """Revisa cada candidata con la IA. Devuelve (confirmadas, descartadas).
    Ante un error de IA, NO descarta la candidata (la conserva marcada para revisión
    manual): preferimos un falso positivo a perder una coincidencia real."""
    if not GEMINI_API_KEY:
        msg = ("USAR_IA está activado pero falta GEMINI_API_KEY. Se envían las "
               "coincidencias por keywords sin el filtro de IA. Cargá el secret "
               "GEMINI_API_KEY o poné USAR_IA=false.")
        print(f"  ⚠️  {msg}")
        enviar_alerta_error("IA activada sin GEMINI_API_KEY", msg)
        return alertas, []

    confirmadas, descartadas, errores = [], [], 0
    for a in alertas:
        desc = (f"Organismo: {a.get('ORGANISMO', '')}. "
                f"Norma objetivo: {a.get('NUMERO_NORMA', '')}. "
                f"Términos del caso: {', '.join(a['KEYWORDS_ENCONTRADAS'])}.")
        coincide, razon = confirmar_con_ia(a['ID'], a['ACCION'], desc, a.get('TEXTO_NORMA', ''))
        a['IA_RAZON'] = razon
        if coincide is None:
            errores += 1
            a['IA_RAZON'] = f"IA no disponible ({razon}); revisar manualmente"
            confirmadas.append(a)
        elif coincide:
            confirmadas.append(a)
        else:
            descartadas.append(a)
        if coincide is None:
            print(f"  🤖 {a['ID']}: ⚠️ sin respuesta — {razon}")
        else:
            print(f"  🤖 {a['ID']}: {'✅ confirma' if coincide else '✖ descarta'} — {razon}")
        time.sleep(IA_DELAY_SEGUNDOS)

    print(f"  🤖 IA: {len(confirmadas)} confirmadas, {len(descartadas)} descartadas, "
          f"{errores} sin respuesta (de {len(alertas)} candidatas).")
    if alertas and errores == len(alertas):
        enviar_alerta_error(
            "La IA no respondió en ninguna consulta",
            "Ninguna consulta a Gemini tuvo éxito (revisá GEMINI_API_KEY, el modelo o la "
            "cuota). Se enviaron las coincidencias por keywords sin filtrar.")
    return confirmadas, descartadas


# ═══════════════════════════════════════════════════════════════════════════════
# MÓDULO 4: CONSTRUCCIÓN DEL EMAIL
# ═══════════════════════════════════════════════════════════════════════════════

def construir_html_listado(normas):
    """Genera un listado limpio de normas agrupadas por tipo (sin IA)."""
    if not normas:
        return ""

    # Agrupar por tipo
    por_tipo = {}
    total_rrhh = 0
    for n in normas:
        if n['ES_RRHH']:
            total_rrhh += 1
            continue
        tipo = n['TIPO']
        if tipo not in por_tipo:
            por_tipo[tipo] = []
        por_tipo[tipo].append(n)

    total_sustantivas = sum(len(v) for v in por_tipo.values())

    html = f"""
    <div style="margin-top: 30px; border-top: 3px solid #2c3e50; padding-top: 20px;">
        <h2 style="color: #2c3e50; margin-bottom: 5px;">
            📋 Normas publicadas hoy en la Primera Sección
        </h2>
        <p style="color: #7f8c8d; margin-top: 0;">
            {total_sustantivas} normas sustantivas — {total_rrhh} de personal/RRHH filtradas
        </p>
    """

    for tipo in ORDEN_TIPOS:
        if tipo == "Personal (RRHH)" or tipo not in por_tipo:
            continue
        lista = por_tipo[tipo]
        color_borde, color_fondo = COLORES_TIPO.get(tipo, ("#7f8c8d", "#f2f3f4"))

        html += f"""
        <div style="margin: 15px 0;">
            <h3 style="color: {color_borde}; border-bottom: 2px solid {color_borde};
                       padding-bottom: 5px; margin-bottom: 8px;">
                {tipo} ({len(lista)})
            </h3>
        """

        for norma in lista:
            titulo = norma['TEXTO'][:140]
            if len(norma['TEXTO']) > 140:
                titulo += "..."

            html += f"""
            <div style="background-color: {color_fondo}; padding: 10px 14px;
                        border-left: 4px solid {color_borde}; margin: 6px 0 6px 10px;
                        border-radius: 0 4px 4px 0;">
                <p style="margin: 0 0 4px 0; font-size: 13px; color: #2c3e50;">
                    {titulo}
                </p>
                <a href="{norma['URL']}" style="font-size: 12px; color: {color_borde};">
                    🔗 Ver norma completa
                </a>
            </div>
            """

        html += "</div>"

    if total_rrhh > 0:
        html += f"""
        <p style="color: #95a5a6; font-style: italic; margin-top: 15px;">
            + {total_rrhh} resoluciones de personal filtradas
            (designaciones, prórrogas, renuncias, etc.)
        </p>
        """

    html += "</div>"
    return html


def construir_email_completo(alertas, normas_ignoradas, listado_html="", descartadas_ia=None):
    fecha_hoy = date.today().strftime('%d/%m/%Y')

    cuerpo = f"""
    <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 800px; margin: 0 auto;">
        <h2 style="color: #2c3e50;">Radar Legislativo — {fecha_hoy}</h2>
        <p><a href='{URL_PRIMERA_SECCION}'>Ver Primera Sección en el BORA</a></p>
        <hr style="border: 1px solid #ecf0f1;">
    """

    # Coincidencias con la base de datos
    if alertas:
        cuerpo += f"""
        <div style="background-color: #fdf2f2; border: 1px solid #e74c3c;
                    padding: 15px; border-radius: 6px; margin: 15px 0;">
            <h3 style="color: #e74c3c; margin-top: 0;">
                🎯 {len(alertas)} coincidencia(s) con tu base de datos
            </h3>
        """
        for a in alertas:
            cuerpo += f"""
            <div style="background: white; padding: 12px; margin: 10px 0;
                        border-left: 4px solid #e74c3c; border-radius: 0 4px 4px 0;">
                <h4 style="margin: 0 0 5px 0; color: #2c3e50;">
                    Caso: {a['ID']} — Acción: {a['ACCION']}
                </h4>
                <p style="margin: 3px 0; font-size: 13px;">
                    <b>Tipo:</b> {a['TIPO']}
                </p>
                <p style="margin: 3px 0; font-size: 13px;">
                    <b>Score:</b> {a['SCORE']:.1f} —
                    <b>Coincidencia por:</b> {', '.join(a.get('KEYWORDS_DISTINTIVAS') or a['KEYWORDS_ENCONTRADAS'])}
                </p>
                <p style="margin: 3px 0; font-size: 12px; color: #7f8c8d;">
                    Otras keywords: {', '.join(a['KEYWORDS_ENCONTRADAS'])}
                </p>
                <p style="margin: 3px 0; font-size: 13px; color: #16a085;">
                    {('🤖 <b>IA:</b> ' + a['IA_RAZON']) if a.get('IA_RAZON') else ''}
                </p>
                <p style="margin: 3px 0; font-size: 13px;">
                    <b>Título:</b> {a['TITULO']}
                </p>
                <p style="margin: 5px 0 0 0;">
                    <a href='{a['URL']}' style="font-size: 12px;">🔗 Leer norma completa</a>
                </p>
            </div>
            """
        cuerpo += "</div>"
    else:
        cuerpo += """
        <p style="color: #27ae60; font-weight: bold; margin-bottom: 4px;">
            ✅ El Boletín de hoy se leyó con éxito, pero no hubo coincidencias con la base.
        </p>
        <p style="color: #7f8c8d; font-size: 13px; margin-top: 0;">
            Igual se sugiere una revisión manual del listado de abajo por las dudas.
        </p>
        """

    # Candidatas que la IA descartó (se muestran para poder auditar la IA).
    if descartadas_ia:
        cuerpo += f"""
        <div style="background-color: #f8f9f9; border: 1px solid #d5dbdb;
                    padding: 12px 15px; border-radius: 6px; margin: 15px 0;">
            <h4 style="color: #7f8c8d; margin: 0 0 8px 0;">
                🤖 {len(descartadas_ia)} candidata(s) revisada(s) y descartada(s) por la IA
            </h4>
            <p style="color: #95a5a6; font-size: 12px; margin-top: 0;">
                Coincidieron por palabras clave pero la IA las consideró no relevantes.
                Listadas por si querés verificar.
            </p>
        """
        for d in descartadas_ia:
            cuerpo += f"""
            <div style="margin: 6px 0 6px 10px; font-size: 12px; color: #7f8c8d;">
                <b>{d['ID']}</b> — {d['TITULO']}<br>
                🤖 {d.get('IA_RAZON', '')}
                — <a href='{d['URL']}' style="font-size: 11px;">ver norma</a>
            </div>
            """
        cuerpo += "</div>"

    # Listado del boletín
    if listado_html:
        cuerpo += listado_html

    # Footer
    cuerpo += f"""
        <hr style="border: 1px solid #ecf0f1; margin-top: 25px;">
        <p style="color: #95a5a6; font-size: 12px;">
            Se filtraron {normas_ignoradas} resoluciones de personal (ruido administrativo).
            <br>Generado automáticamente por Radar Desregulación.
        </p>
    </div>
    """
    return cuerpo


# ═══════════════════════════════════════════════════════════════════════════════
# MÓDULO 5: ENVÍO DE EMAIL (soporta múltiples destinatarios)
# ═══════════════════════════════════════════════════════════════════════════════

def enviar_html(asunto: str, cuerpo_html: str) -> bool:
    """Envía un email HTML por SMTP de Gmail. Devuelve True si se envió.
    Requiere EMAIL_ORIGEN (cuenta Gmail) y PASSWORD_APP (contraseña de aplicación)."""
    destinatarios = parsear_destinatarios(EMAIL_DESTINO)
    if not all([EMAIL_ORIGEN, PASSWORD_APP]) or not destinatarios:
        print("❌ Faltan credenciales de email (EMAIL_ORIGEN/PASSWORD_APP) o destinatarios.")
        return False

    msg = MIMEMultipart()
    msg['From'] = EMAIL_ORIGEN
    msg['To'] = ', '.join(destinatarios)
    msg['Subject'] = asunto
    msg.attach(MIMEText(cuerpo_html, 'html'))

    try:
        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_ORIGEN, PASSWORD_APP)
        server.sendmail(EMAIL_ORIGEN, destinatarios, msg.as_string())
        server.quit()
        print(f"✅ Email enviado a: {', '.join(destinatarios)}")
        return True
    except Exception as e:
        print(f"❌ Error al enviar email: {e}")
        return False


def enviar_alerta_error(titulo_error: str, detalle: str):
    """Envía un email de ALERTA cuando el radar no puede operar con normalidad.

    El objetivo es 'fallar ruidoso': si algo está mal (base vacía, BORA caído),
    el equipo debe enterarse, en lugar de recibir un 'sin novedades' que oculta
    el problema. Si no hay credenciales, al menos queda registrado en el log.
    """
    fecha_hoy = date.today().strftime('%d/%m/%Y')
    cuerpo_html = f"""
    <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 700px; margin: 0 auto;">
        <div style="background-color: #fdf2f2; border: 2px solid #c0392b;
                    padding: 18px; border-radius: 6px;">
            <h2 style="color: #c0392b; margin-top: 0;">🚨 Radar Desregulación — ALERTA</h2>
            <p style="font-size: 14px;"><b>{titulo_error}</b></p>
            <p style="font-size: 13px; color: #2c3e50;">{detalle}</p>
            <p style="font-size: 12px; color: #7f8c8d; margin-top: 15px;">
                Fecha: {fecha_hoy}. El radar se detuvo deliberadamente: este aviso
                significa que NO se hizo el escaneo normal y conviene revisarlo.
            </p>
        </div>
    </div>
    """
    if not enviar_html(f"🚨 RADAR: ERROR — {titulo_error} ({fecha_hoy})", cuerpo_html):
        print(f"🚨 ALERTA (no se pudo enviar por email): {titulo_error} — {detalle}")


def enviar_aviso(titulo: str, detalle: str):
    """Envía un email INFORMATIVO (no de error) cuando el radar corrió pero no pudo
    llegar a un resultado concluyente hoy (BORA inaccesible, día sin novedades, etc.).
    No es una falla del sistema: solo avisa que conviene una revisión manual."""
    fecha_hoy = date.today().strftime('%d/%m/%Y')
    cuerpo_html = f"""
    <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 700px; margin: 0 auto;">
        <div style="background-color: #eef5fb; border: 1px solid #2980b9;
                    padding: 18px; border-radius: 6px;">
            <h2 style="color: #2471a3; margin-top: 0;">ℹ️ Radar Desregulación — {fecha_hoy}</h2>
            <p style="font-size: 14px;"><b>{titulo}</b></p>
            <p style="font-size: 13px; color: #2c3e50;">{detalle}</p>
            <p style="font-size: 13px; color: #2c3e50; margin-top: 12px;">
                👉 Se recomienda una <b>revisión manual</b> del Boletín de hoy:
                <a href="{URL_PRIMERA_SECCION}">ver Primera Sección en el BORA</a>.
            </p>
        </div>
    </div>
    """
    if not enviar_html(f"ℹ️ RADAR: revisión manual sugerida ({fecha_hoy})", cuerpo_html):
        print(f"ℹ️ AVISO (no se pudo enviar por email): {titulo} — {detalle}")


def enviar_email(alertas, normas_ignoradas, listado_html="", descartadas_ia=None):
    fecha_hoy = date.today().strftime('%d/%m/%Y')

    if alertas:
        asunto = f"🔴 RADAR: {len(alertas)} coincidencia(s) en el BORA ({fecha_hoy})"
    elif listado_html:
        asunto = f"📋 Resumen del Boletín Oficial ({fecha_hoy})"
    else:
        asunto = f"✅ RADAR: Sin novedades ({fecha_hoy})"

    cuerpo_html = construir_email_completo(alertas, normas_ignoradas, listado_html, descartadas_ia)
    enviar_html(asunto, cuerpo_html)


# ═══════════════════════════════════════════════════════════════════════════════
# EJECUCIÓN PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════

def ejecutar_radar():
    print("=" * 60)
    print("  RADAR DESREGULACIÓN")
    print("=" * 60)

    df = cargar_archivo_robusto()
    normas, leido_ok = escanear_boletin()

    # Caso 1: no se pudo ni acceder al BORA (timeout / sitio caído). NO es una falla
    # del sistema: se manda un aviso informativo y se termina limpio (sin error rojo).
    if not leido_ok:
        enviar_aviso(
            "No se pudo acceder al Boletín Oficial hoy",
            "El radar intentó leer la Primera Sección del BORA pero no obtuvo respuesta "
            "(timeout o sitio caído). Suele ser un problema transitorio de red. No se "
            "pudo verificar el boletín de hoy.")
        print("ℹ️ BORA inaccesible: se envió aviso de revisión manual.")
        return

    # Caso 2: se accedió pero con muy pocas normas (día no hábil / sin edición, o un
    # posible cambio en el sitio). Aviso informativo, sin cortar con error.
    if len(normas) < MINIMO_NORMAS_ESPERADAS:
        enviar_aviso(
            "El Boletín de hoy trae muy pocas o ninguna norma",
            f"Se accedió al BORA correctamente, pero solo se encontraron {len(normas)} "
            f"norma(s) en la Primera Sección. Puede ser un día no hábil o sin edición. "
            f"Si esperabas que hubiera boletín, podría indicar un cambio en el sitio.")
        print(f"ℹ️ Solo {len(normas)} norma(s): se envió aviso de revisión manual.")
        return

    print(f"\n🔎 Cruzando {len(df)} casos contra {len(normas)} normas...")
    alertas, normas_ignoradas = cruzar_con_base(normas, df)

    print(f"\n📊 Coincidencias por keywords: {len(alertas)} | {normas_ignoradas} de RRHH filtradas.")

    # --- FASE 4: confirmación con IA (Gemini) ---
    # El filtro de keywords ya preseleccionó candidatas; la IA descarta las que
    # solo comparten vocabulario. Apagada por defecto (USAR_IA).
    descartadas_ia = []
    if USAR_IA and alertas:
        print(f"\n🤖 Confirmando {len(alertas)} candidata(s) con IA ({GEMINI_MODEL})...")
        alertas, descartadas_ia = confirmar_alertas_con_ia(alertas)
        print(f"📊 Tras IA: {len(alertas)} confirmadas, {len(descartadas_ia)} descartadas.")

    # Salvaguarda: un número anómalo de coincidencias sugiere una regresión.
    if len(alertas) > MAX_ALERTAS_RAZONABLE:
        msg = (f"El radar generó {len(alertas)} coincidencias en un día, muy por encima "
               f"de lo normal ({MAX_ALERTAS_RAZONABLE}). Suele indicar keywords demasiado "
               f"genéricas o un cambio que reintrodujo ruido. Conviene revisar antes de "
               f"confiar en el resultado de hoy.")
        print(f"  ⚠️  {msg}")
        enviar_alerta_error("Número anómalo de coincidencias", msg)

    listado_html = ""
    if INCLUIR_LISTADO_BOLETIN:
        listado_html = construir_html_listado(normas)

    enviar_email(alertas, normas_ignoradas, listado_html, descartadas_ia)


if __name__ == "__main__":
    ejecutar_radar()
