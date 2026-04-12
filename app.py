"""
app.py — DocVerify Backend
Flask API REST
Listo para deploy en Render.com
"""

import os
import re
import io
import signal
import tempfile
from datetime import datetime

from flask import Flask, request, jsonify
from flask_cors import CORS
from docx import Document
from spellchecker import SpellChecker

# ─────────────────────────────────────────────────
# FLASK SETUP
# ─────────────────────────────────────────────────
app = Flask(__name__)

CORS(app, origins=["https://ever186.github.io/docverify/"])

app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10 MB máximo

ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# ─────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────
PLACEHOLDER = re.compile(r'^<[^>]+>$')

def es_placeholder(val: str) -> bool:
    return bool(PLACEHOLDER.match(val.strip()))

def normalizar(texto: str) -> str:
    """Normaliza guiones Unicode y espacios para comparación robusta."""
    return texto.lower().replace('\u2013', '-').replace('\u2014', '-').strip()


# ─────────────────────────────────────────────────
# 1. EXTRACCIÓN
# ─────────────────────────────────────────────────
def extraer_contenido(file_bytes: bytes) -> dict:
    """
    Lee el .docx desde bytes en memoria
    """
    stream = io.BytesIO(file_bytes)
    doc = Document(stream)
    props = doc.core_properties

    metadatos = {
        "autor_meta":     (props.author or "").strip(),
        "modificado_por": (props.last_modified_by or "").strip(),
        "modificado":     str(props.modified),
    }

    encabezados = []
    for section in doc.sections:
        for para in section.header.paragraphs:
            t = para.text.strip()
            if t:
                encabezados.append(t)

    parrafos = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    tablas = []
    for tabla in doc.tables:
        filas = []
        for row in tabla.rows:
            fila = [cell.text.strip() for cell in row.cells]
            filas.append(fila)
        tablas.append(filas)

    texto_completo = "\n".join(
        encabezados + parrafos +
        [c for t in tablas for f in t for c in f]
    )

    return {
        "metadatos":      metadatos,
        "encabezados":    encabezados,
        "parrafos":       parrafos,
        "tablas":         tablas,
        "texto_completo": texto_completo,
    }


# ─────────────────────────────────────────────────
# 2. PARÁMETROS
# ─────────────────────────────────────────────────
def verificar_parametros(contenido, titulo_p, id_tarea, consecutivo):
    resultados = []
    tablas = contenido["tablas"]
    encabezados = contenido["encabezados"]

    def celda(t, f, c):
        try:
            return tablas[t][f][c].strip()
        except IndexError:
            return ""

    def idx_conclusiones(tablas):
        for i, tabla in enumerate(tablas):
            for fila in tabla:
                for celda_val in fila:
                    if "id azure" in celda_val.lower() or "consecutivo" in celda_val.lower():
                        return i
        return 4

    t_conc = idx_conclusiones(tablas)

    # ── TÍTULO
    if not titulo_p:
        resultados.append({"estado":"WARN","detalle":"No se ingresó título"})
    else:
        titulo_norm  = normalizar(titulo_p)
        en_header    = any(titulo_norm in normalizar(e) for e in encabezados)
        en_historial = titulo_norm in normalizar(celda(2, 1, 3))

        if en_header and en_historial:
            resultados.append({"estado":"OK","detalle":f'"{titulo_p}" presente en encabezado y historial'})
        elif en_header:
            val_hist = celda(2, 1, 3)
            resultados.append({"estado":"WARN","detalle":f'"{titulo_p}" solo en encabezado. Historial contiene: "{val_hist[:80]}"'})
        elif en_historial:
            resultados.append({"estado":"WARN","detalle":f'"{titulo_p}" solo en historial, falta en encabezado'})
        else:
            resultados.append({"estado":"ERROR","detalle":f'"{titulo_p}" no encontrado en ninguna ubicación del documento'})

    # ── ID TAREA
    if not id_tarea:
        resultados.append({"estado":"WARN","detalle":"No se ingresó ID de tarea"})
    else:
        v = celda(t_conc, 1, 1)
        if normalizar(id_tarea) in normalizar(v):
            resultados.append({"estado":"OK","detalle":f'"{id_tarea}" encontrado en tabla conclusiones (ID Azure)'})
        elif es_placeholder(v):
            resultados.append({"estado":"WARN","detalle":f'Celda ID Azure es placeholder "{v}", no "{id_tarea}"'})
        elif v == "":
            resultados.append({"estado":"ERROR","detalle":"Celda ID Azure vacía — revisar estructura del documento"})
        else:
            resultados.append({"estado":"ERROR","detalle":f'Se esperaba "{id_tarea}", se encontró "{v}"'})

    # ── CONSECUTIVO
    if not consecutivo:
        resultados.append({"estado":"WARN","detalle":"No se ingresó consecutivo"})
    else:
        v1 = celda(1, 0, 1)
        v2 = celda(t_conc, 2, 1)
        encontrado_en, placeholder_en = [], []

        for v, etq in [(v1, "tabla_info→Código"), (v2, "tabla_conclusiones→Consecutivo")]:
            if normalizar(consecutivo) in normalizar(v):
                encontrado_en.append(etq)
            elif es_placeholder(v):
                placeholder_en.append(etq)

        if len(encontrado_en) == 2:
            resultados.append({"estado":"OK","detalle":f'"{consecutivo}" coherente en tabla_info y tabla_conclusiones'})
        elif len(encontrado_en) == 1:
            resultados.append({"estado":"WARN","detalle":f'"{consecutivo}" solo encontrado en {encontrado_en[0]}'})
        elif placeholder_en:
            resultados.append({"estado":"WARN","detalle":f'Celdas de consecutivo sin rellenar: {placeholder_en}'})
        else:
            resultados.append({"estado":"ERROR","detalle":f'"{consecutivo}" no coincide. tabla_info="{v1}", conclusiones="{v2}"'})

    return resultados


# ─────────────────────────────────────────────────
# 3. FECHAS
# ─────────────────────────────────────────────────
PATRONES_FECHA = [
    r'\b\d{1,2}/\d{1,2}/\d{4}\b',
    r'\b\d{1,2}-\d{1,2}-\d{4}\b',
    r'\b\d{4}-\d{2}-\d{2}\b',
    r'\b\d{1,2}\s+de\s+\w+\s+de\s+\d{4}\b',
    r'\b\d{1,2}\s+\w+\s+\d{4}\b',
]

def extraer_fechas(texto):
    fechas = []
    for p in PATRONES_FECHA:
        fechas += re.findall(p, texto, re.IGNORECASE)
    return list(set(fechas))

def verificar_fechas(contenido):
    resultados = []
    tablas = contenido["tablas"]
    encabezados = contenido["encabezados"]
    hallazgos = {}

    for enc in encabezados:
        for f in extraer_fechas(enc):
            hallazgos.setdefault("encabezado", []).append(f)

    for fi in [1, 2]:
        try:
            v = tablas[2][fi][0]
            if not es_placeholder(v) and v:
                for f in extraer_fechas(v):
                    hallazgos.setdefault(f"historial_fila{fi}", []).append(f)
        except IndexError:
            pass

    if not hallazgos:
        resultados.append({"estado":"WARN","detalle":"Sin fechas concretas — celdas probablemente son placeholders"})
        return resultados

    todas = []
    for ubicacion, fechas in hallazgos.items():
        for f in fechas:
            resultados.append({"estado":"INFO","detalle":f'Fecha en [{ubicacion}]: "{f}"'})
            todas.append(f)

    uniq = list(set(todas))
    if len(uniq) == 1:
        resultados.append({"estado":"OK","detalle":f'Fechas consistentes en todo el documento: "{uniq[0]}"'})
    elif len(uniq) > 1:
        resultados.append({"estado":"ERROR","detalle":f'Fechas distintas encontradas: {uniq}'})

    return resultados


# ─────────────────────────────────────────────────
# 4. AUTORES
# ─────────────────────────────────────────────────
def verificar_autores(contenido):
    resultados = []
    tablas = contenido["tablas"]
    meta_autor = contenido["metadatos"]["autor_meta"]
    meta_mod   = contenido["metadatos"]["modificado_por"]

    resultados.append({"estado":"INFO","detalle":f'Autor en metadatos del archivo: "{meta_autor}"'})
    resultados.append({"estado":"INFO","detalle":f'Modificado por (metadatos): "{meta_mod}"'})

    ubicaciones = [
        (1, 1, 1, "tabla_info→Autor"),
        (2, 1, 4, "historial→fila1→Autor"),
        (2, 2, 4, "historial→fila2→Autor"),
    ]

    autores_doc = {}
    for t_i, f_i, c_i, etq in ubicaciones:
        try:
            v = tablas[t_i][f_i][c_i].strip()
            if v and not es_placeholder(v):
                autores_doc[etq] = v
                resultados.append({"estado":"INFO","detalle":f'Autor en [{etq}]: "{v}"'})
            elif es_placeholder(v):
                resultados.append({"estado":"WARN","detalle":f'[{etq}] es placeholder sin rellenar: "{v}"'})
        except IndexError:
            pass

    if not autores_doc:
        resultados.append({"estado":"WARN","detalle":"Todas las celdas de autor son placeholders"})
        return resultados

    valores = list(set(autores_doc.values()))
    if len(valores) == 1:
        resultados.append({"estado":"OK","detalle":f'Autor consistente en todo el documento: "{valores[0]}"'})
    else:
        resultados.append({"estado":"ERROR","detalle":f'Autores inconsistentes en el documento: {valores}'})

    for val in valores:
        coincide = meta_autor and (
            meta_autor.lower() in val.lower() or val.lower() in meta_autor.lower()
        )
        if meta_autor and not coincide:
            resultados.append({"estado":"WARN","detalle":f'Autor del documento "{val}" difiere del metadato del archivo "{meta_autor}"'})
        elif meta_autor and coincide:
            resultados.append({"estado":"OK","detalle":f'"{val}" coincide con metadatos del archivo'})

    return resultados


# ─────────────────────────────────────────────────
# 5. ORTOGRAFÍA
# ─────────────────────────────────────────────────
IGNORAR_PALABRAS = {
    "sanitizar","canonicalización","canonicalizar","csrf","xss","sql",
    "api","backend","frontend","payload","bypass","exploit","fuzzing",
    "pentesting","https","http","rate","limiting","aplicaciones","aplicación",
    "autenticación","autorización","controles","control","funcionalidades",
    "funcionalidad","vulnerabilidades","caracteres",
    "conclusiones","autorizados","apropiados","aquellos","corresponden",
    "deben","datos","ficheros","directo","encontradas","entreguen",
    "evaluada","incluyendo","implementar","periódica","confidencial",
    "infraestructura","telecomunicaciones","informática","módulos",
    "resultados","solicitaron","distribución","codificar","contextualmente",
    "directamente","publicación","suministrado","específicamente",
    "provista","ninguna","provisto","subversión","versión","sección",
    "también","través","según","módulo","opción","opciones",
    "deberán","tendrán","están","básicamente","acceso",
    "usuario","usuarios","servidor","servidores","sistema","sistemas",
    "seguridad","prueba","pruebas","informe","informes","semestral",
    "documentación","actualización","revisión","creación","generación",
    "aprobación","modificación","versiones","fecha","registro","registros",
    "información","gerencia","dirección","recomendaciones","implementación",
    "validación","verificación","proteger","enviar","exigir","perfilar",
    "restringir","asignar","utilizar","implementando","modificaciones",
    "observaciones","obtenidos","realizadas","relacionadas","revisiones",
    "meses","necesarios","otros","pasados","privilegios","recursos",
    "salidas","sesiones","siguientes","todas","valores","áreas","tarea",
}

def verificar_ortografia(contenido):
    resultados = []

    texto = contenido["texto_completo"]
    texto = re.sub(r'<[^>]+>', ' ', texto)
    texto = re.sub(r'https?://\S+', ' ', texto)
    texto = re.sub(r'\b[A-Z]{2,}\b', ' ', texto)
    texto = re.sub(r'\b\d[\d.,/-]*\b', ' ', texto)
    texto = re.sub(r'[^\w\sáéíóúÁÉÍÓÚñÑüÜ]', ' ', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()

    palabras = [p.lower() for p in texto.split() if len(p) > 3]
    palabras_a_revisar = [p for p in set(palabras) if p not in IGNORAR_PALABRAS]

    if not palabras_a_revisar:
        resultados.append({"estado":"OK","detalle":"Sin errores ortográficos detectados"})
        return resultados

    try:
        def timeout_handler(sig, frame):
            raise TimeoutError()

        signal.signal(signal.SIGALRM, timeout_handler)
        signal.alarm(15)

        spell = SpellChecker(language='es')
        desconocidas = spell.unknown(palabras_a_revisar)
        errores = [p for p in desconocidas if p not in IGNORAR_PALABRAS]

        signal.alarm(0)

        if not errores:
            resultados.append({"estado":"OK","detalle":"Sin errores ortográficos detectados"})
        else:
            resultados.append({"estado":"INFO","detalle":f"{len(errores)} posible(s) error(es) ortográfico(s)"})
            for palabra in sorted(errores)[:20]:
                sugerencia = spell.correction(palabra)
                if sugerencia and sugerencia != palabra:
                    resultados.append({"estado":"WARN","detalle":f'"{palabra}" → sugerencia: "{sugerencia}"'})
                else:
                    resultados.append({"estado":"WARN","detalle":f'"{palabra}" no reconocida en diccionario español'})

    except TimeoutError:
        resultados.append({"estado":"WARN","detalle":"Verificación ortográfica omitida (tiempo límite excedido)"})
    except Exception as e:
        resultados.append({"estado":"WARN","detalle":f"Error en verificación ortográfica: {str(e)}"})

    return resultados


# ─────────────────────────────────────────────────
# ENDPOINT PRINCIPAL
# ─────────────────────────────────────────────────
@app.route('/verificar', methods=['POST'])
def verificar():
    # ── Validar archivo
    if 'file' not in request.files:
        return jsonify({"error": "No se recibió ningún archivo"}), 400

    file = request.files['file']

    if not file or not file.filename:
        return jsonify({"error": "Archivo vacío"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Solo se aceptan archivos .docx"}), 400

    # ── Leer en memoria (ZERO RETENTION — nunca toca disco)
    file_bytes = file.read()
    if len(file_bytes) == 0:
        return jsonify({"error": "El archivo está vacío"}), 400

    # ── Parámetros
    titulo_p    = request.form.get('titulo', '').strip()
    id_tarea    = request.form.get('id_tarea', '').strip()
    consecutivo = request.form.get('consecutivo', '').strip()

    # ── Procesar
    try:
        contenido = extraer_contenido(file_bytes)
    except Exception as e:
        return jsonify({"error": f"No se pudo leer el documento: {str(e)}"}), 422

    resultados = {
        "parametros": verificar_parametros(contenido, titulo_p, id_tarea, consecutivo),
        "fechas":     verificar_fechas(contenido),
        "autores":    verificar_autores(contenido),
        "ortografia": verificar_ortografia(contenido),
    }

    # ── Conteos resumen
    total_ok   = sum(1 for rs in resultados.values() for r in rs if r["estado"] == "OK")
    total_warn = sum(1 for rs in resultados.values() for r in rs if r["estado"] == "WARN")
    total_err  = sum(1 for rs in resultados.values() for r in rs if r["estado"] == "ERROR")

    return jsonify({
        "resultados": resultados,
        "resumen": {
            "ok":      total_ok,
            "alertas": total_warn,
            "errores": total_err,
        }
    })


@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok", "service": "DocVerify"})


# ─────────────────────────────────────────────────
# RUN
# ─────────────────────────────────────────────────
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_ENV') == 'development'
    app.run(host='0.0.0.0', port=port, debug=debug)
