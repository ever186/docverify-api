"""
app.py — Verificador de coherencia .docx
Flask API REST — Zero data retention
"""

import os
import re
import io
import signal
import threading

from flask import Flask, request, jsonify
from flask_cors import CORS
from docx import Document
from spellchecker import SpellChecker

app = Flask(__name__)
CORS(app,
    origins=["https://ever186.github.io"],
    methods=["GET", "POST", "OPTIONS"],
    allow_headers=["Content-Type", "Authorization"],
    expose_headers=["Content-Type"],
    max_age=600
)
app.config['MAX_CONTENT_LENGTH'] = 15 * 1024 * 1024  # 15 MB

@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": "El archivo es demasiado grande (máx 15 MB)"}), 413

@app.errorhandler(Exception)
def handle_exception(e):
    return jsonify({"error": f"Error interno: {str(e)}"}), 500

ALLOWED_EXTENSIONS = {'docx'}
def allowed_file(f): return '.' in f and f.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
PLACEHOLDER = re.compile(r'^<[^>]+>$')
def es_placeholder(v): return bool(PLACEHOLDER.match(v.strip()))

def normalizar(t):
    return t.lower().replace('\u2013','-').replace('\u2014','-').strip()

PATRONES_FECHA = [
    r'\b\d{1,2}/\d{1,2}/\d{4}\b',
    r'\b\d{4}/\d{2}/\d{2}\b',
    r'\b\d{1,2}-\d{1,2}-\d{4}\b',
    r'\b\d{4}-\d{2}-\d{2}\b',
    r'\b\d{1,2}\s+de\s+\w+\s+de\s+\d{4}\b',
    r'\b\d{1,2}\s+\w+\s+\d{4}\b',
]

def extraer_fechas(texto):
    fechas = []
    for p in PATRONES_FECHA:
        fechas += re.findall(p, texto, re.IGNORECASE)
    fechas = [f for f in fechas if not re.match(r'^\d{1,2}\.\d{1,2}$', f)]
    return list(set(fechas))

def idx_conclusiones(tablas):
    for i, tabla in enumerate(tablas):
        for fila in tabla:
            for c in fila:
                if 'id azure' in c.lower() or ('consecutivo' in c.lower() and 'consecutivo' != c.lower()):
                    return i
    # fallback: buscar tabla con campo "consecutivo"
    for i, tabla in enumerate(tablas):
        for fila in tabla:
            if any(c.lower() == 'consecutivo' for c in fila):
                return i
    return -1

def celda_safe(tablas, t, f, c):
    try: return tablas[t][f][c].strip()
    except: return ""

# ─────────────────────────────────────────────
# EXTRACCIÓN
# ─────────────────────────────────────────────
def extraer_contenido(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    props = doc.core_properties

    metadatos = {
        "autor_meta":     (props.author or "").strip(),
        "modificado_por": (props.last_modified_by or "").strip(),
    }

    encabezados = []
    for section in doc.sections:
        vistos = set()
        for para in section.header.paragraphs:
            t = para.text.strip()
            if t and t not in vistos:
                encabezados.append(t)
                vistos.add(t)

    parrafos = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    # Namespace de Word
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    def _texto_tc(tc):
        """Extrae todo el texto de una celda incluyendo runs anidados."""
        return ''.join(t.text or '' for t in tc.iter(f'{{{W}}}t')).strip()

    def _leer_fila(row):
        """
        Lee celdas de una fila manejando:
        - Celdas normales (w:tc)
        - Date pickers y controles ricos (w:sdt que contienen w:tc)
        - Celdas fusionadas (merged) — deduplica por identidad de objeto
        """
        celdas_raw = []
        for child in row._tr:
            local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if local == 'tc':
                celdas_raw.append(_texto_tc(child))
            elif local == 'sdt':
                # Buscar w:tc dentro del sdtContent (date pickers, etc.)
                sdt_ns = f'{{{W}}}sdtContent'
                tc_ns  = f'{{{W}}}tc'
                sdt_content = child.find(f'.//{sdt_ns}')
                if sdt_content is not None:
                    tc = sdt_content.find(f'.//{tc_ns}')
                    if tc is not None:
                        celdas_raw.append(_texto_tc(tc))
                        continue
                # Fallback: extraer texto directo del SDT
                celdas_raw.append(''.join(t.text or '' for t in child.iter(f'{{{W}}}t')).strip())

        # Deduplicar celdas horizontalmente fusionadas
        seen_ids = []
        celdas_unicas = []
        for cell in row.cells:
            cid = id(cell._tc)
            if cid not in seen_ids:
                seen_ids.append(cid)
                celdas_unicas.append(cell)

        # Usar la lectura SDT si tiene más columnas que la deduplicada
        if len(celdas_raw) >= len(celdas_unicas):
            return celdas_raw
        return [cell.text.strip() for cell in celdas_unicas]

    tablas = []
    for tabla in doc.tables:
        filas = [_leer_fila(row) for row in tabla.rows]
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

# ─────────────────────────────────────────────
# SECCIÓN 1 — ENCABEZADO
# ─────────────────────────────────────────────
def seccion_encabezado(contenido, titulo_p):
    """Muestra el encabezado completo y valida el título."""
    encabezados = contenido["encabezados"]
    validaciones = []
    fragmento = []

    for enc in encabezados:
        fragmento.append(enc)

    # Validar título en encabezado
    if titulo_p:
        en_header = any(normalizar(titulo_p) in normalizar(e) for e in encabezados)
        if en_header:
            validaciones.append({"estado":"OK","detalle":f'Título "{titulo_p}" ✔ presente en encabezado'})
        else:
            validaciones.append({"estado":"ERROR","detalle":f'Título "{titulo_p}" ✘ NO encontrado en encabezado'})

    # Validar fechas en encabezado
    fechas_header = []
    for enc in encabezados:
        fechas_header += extraer_fechas(enc)
    if fechas_header:
        validaciones.append({"estado":"INFO","detalle":f'Fecha en encabezado: {list(set(fechas_header))}'})
    else:
        validaciones.append({"estado":"WARN","detalle":"Sin fecha concreta en encabezado (puede ser placeholder)"})

    estado_general = "ERROR" if any(v["estado"]=="ERROR" for v in validaciones) else \
                     "WARN"  if any(v["estado"]=="WARN"  for v in validaciones) else "OK"

    return {
        "titulo":      "Encabezado del documento",
        "estado":      estado_general,
        "fragmento":   fragmento,
        "validaciones": validaciones,
    }

# ─────────────────────────────────────────────
# SECCIÓN 2 — INFORMACIÓN DEL DOCUMENTO
# ─────────────────────────────────────────────
def seccion_info_documento(contenido, consecutivo):
    """Muestra la tabla de información completa y valida consecutivo y autor."""
    tablas    = contenido["tablas"]
    metadatos = contenido["metadatos"]
    validaciones = []
    fragmento = []   # lista de {campo, valor, estado}

    try:
        tabla_info = tablas[1]
    except IndexError:
        return {"titulo":"Información del documento","estado":"WARN",
                "fragmento":[],"validaciones":[{"estado":"WARN","detalle":"Tabla de información no encontrada"}]}

    # Construir fragmento visual con TODAS las filas de la tabla
    campos_importantes = {}
    for fila in tabla_info:
        if len(fila) >= 2:
            campo = fila[0].rstrip(':').strip()
            valor = fila[1].strip()
        elif len(fila) == 1:
            campo = fila[0].rstrip(':').strip()
            valor = ""
        else:
            continue

        es_ph = es_placeholder(valor) if valor else False
        fragmento.append({
            "campo": campo,
            "valor": valor if valor else "—",
            "es_placeholder": es_ph
        })
        campos_importantes[campo.lower()] = valor

    # Validar consecutivo
    if consecutivo:
        val_cod = campos_importantes.get("código", "") or campos_importantes.get("codigo", "")
        if normalizar(consecutivo) in normalizar(val_cod):
            validaciones.append({"estado":"OK","detalle":f'Consecutivo "{consecutivo}" ✔ coincide con Código'})
        elif es_placeholder(val_cod):
            validaciones.append({"estado":"WARN","detalle":f'Código aún es placeholder: "{val_cod}"'})
        else:
            validaciones.append({"estado":"ERROR","detalle":f'Consecutivo "{consecutivo}" ✘ no coincide — Código contiene: "{val_cod}"'})

    # Validar autor vs metadatos
    autor_doc = campos_importantes.get("autor", "")
    meta_autor = metadatos["autor_meta"]
    meta_mod   = metadatos["modificado_por"]

    if autor_doc and not es_placeholder(autor_doc):
        validaciones.append({"estado":"INFO","detalle":f'Autor en metadatos del archivo: "{meta_autor}"'})
        validaciones.append({"estado":"INFO","detalle":f'Última modificación por: "{meta_mod}"'})
        coincide = meta_autor and (meta_autor.lower() in autor_doc.lower() or autor_doc.lower() in meta_autor.lower())
        if coincide:
            validaciones.append({"estado":"OK","detalle":f'Autor "{autor_doc}" ✔ coincide con metadatos del archivo'})
        else:
            validaciones.append({"estado":"WARN","detalle":f'Autor "{autor_doc}" ≠ metadato del archivo "{meta_autor}" — verificar si es intencional'})
    else:
        validaciones.append({"estado":"WARN","detalle":"Campo Autor es placeholder o está vacío"})

    estado_general = "ERROR" if any(v["estado"]=="ERROR" for v in validaciones) else \
                     "WARN"  if any(v["estado"]=="WARN"  for v in validaciones) else "OK"

    return {
        "titulo":       "Información del documento",
        "estado":       estado_general,
        "fragmento":    fragmento,
        "validaciones": validaciones,
        "tipo":         "tabla_info",
    }

# ─────────────────────────────────────────────
# SECCIÓN 3 — HISTORIAL DE REVISIONES
# ─────────────────────────────────────────────
def seccion_historial(contenido, titulo_p):
    """Muestra el historial completo y valida título, fechas y autor."""
    tablas = contenido["tablas"]
    validaciones = []
    fragmento = []

    try:
        tabla_hist = tablas[2]
    except IndexError:
        return {"titulo":"Historial de revisiones","estado":"WARN",
                "fragmento":[],"validaciones":[{"estado":"WARN","detalle":"Historial no encontrado"}]}

    # Encabezado del historial
    if tabla_hist:
        encabezado_hist = tabla_hist[0]
        filas_data      = tabla_hist[1:]
    else:
        return {"titulo":"Historial de revisiones","estado":"WARN",
                "fragmento":[],"validaciones":[{"estado":"WARN","detalle":"Historial vacío"}]}

    fragmento = {
        "encabezados": encabezado_hist,
        "filas":       filas_data
    }

    # ── Validar título en historial
    if titulo_p:
        titulo_norm = normalizar(titulo_p)
        en_hist = any(titulo_norm in normalizar(c) for fila in filas_data for c in fila)
        if en_hist:
            validaciones.append({"estado":"OK","detalle":f'Título "{titulo_p}" ✔ presente en historial'})
        else:
            validaciones.append({"estado":"ERROR","detalle":f'Título "{titulo_p}" ✘ NO encontrado en historial'})

    # ── Validar fechas en historial
    fechas_hist = []
    for fila in filas_data:
        for celda in fila:
            if not es_placeholder(celda):
                fechas_hist += extraer_fechas(celda)
    fechas_hist = list(set(fechas_hist))

    if fechas_hist:
        if len(fechas_hist) == 1:
            validaciones.append({"estado":"OK","detalle":f'Fechas en historial coherentes: "{fechas_hist[0]}"'})
        else:
            validaciones.append({"estado":"ERROR","detalle":f'Fechas inconsistentes en historial: {fechas_hist}'})
    else:
        validaciones.append({"estado":"WARN","detalle":"Sin fechas concretas en historial (placeholders o celdas fusionadas)"})

    # ── Validar coherencia de autores en historial
    autores_hist = []
    for fila in filas_data:
        if fila:
            ultimo = fila[-1]
            if ultimo and not es_placeholder(ultimo) and len(ultimo) > 3:
                autores_hist.append(ultimo)
    autores_uniq = list(set(autores_hist))

    if len(autores_uniq) == 1:
        validaciones.append({"estado":"OK","detalle":f'Autor coherente en historial: "{autores_uniq[0]}"'})
    elif len(autores_uniq) > 1:
        validaciones.append({"estado":"ERROR","detalle":f'Autores inconsistentes en historial: {autores_uniq}'})
    else:
        validaciones.append({"estado":"WARN","detalle":"Autores en historial son placeholders"})

    estado_general = "ERROR" if any(v["estado"]=="ERROR" for v in validaciones) else \
                     "WARN"  if any(v["estado"]=="WARN"  for v in validaciones) else "OK"

    return {
        "titulo":       "Historial de revisiones",
        "estado":       estado_general,
        "fragmento":    fragmento,
        "validaciones": validaciones,
        "tipo":         "tabla_historial",
    }

# ─────────────────────────────────────────────
# SECCIÓN 4 — CONCLUSIONES
# ─────────────────────────────────────────────
def seccion_conclusiones(contenido, titulo_p, id_tarea, consecutivo):
    """Muestra la tabla de conclusiones completa + párrafo y valida todos los campos."""
    tablas   = contenido["tablas"]
    parrafos = contenido["parrafos"]
    validaciones = []
    t_conc = idx_conclusiones(tablas)

    # ── Párrafo de conclusiones (contexto textual)
    parrafo_conc = ""
    for para in parrafos:
        if "corresponden a la aplicación" in para.lower() or "resultados obtenidos" in para.lower():
            parrafo_conc = para
            break

    # ── Tabla conclusiones
    tabla_conc_data = []
    if t_conc >= 0:
        try:
            tabla_conc_data = tablas[t_conc]
        except IndexError:
            pass

    fragmento = {
        "parrafo":       parrafo_conc,
        "tabla_campos":  []
    }

    # Construir vista de la tabla de conclusiones
    valores_conc = {}
    if tabla_conc_data:
        for fila in tabla_conc_data[1:]:  # skip header
            if len(fila) >= 2:
                campo = fila[0].strip()
                valor = fila[1].strip()
                es_ph = es_placeholder(valor) if valor else False
                fragmento["tabla_campos"].append({
                    "campo": campo,
                    "valor": valor if valor else "—",
                    "es_placeholder": es_ph
                })
                valores_conc[campo.lower()] = valor

    # ── Validar título en párrafo de conclusiones
    if titulo_p:
        titulo_norm = normalizar(titulo_p)
        en_parrafo = any(titulo_norm in normalizar(p) for p in parrafos)
        en_tabla   = any(titulo_norm in normalizar(c) for fila in tabla_conc_data for c in fila)
        if en_parrafo or en_tabla:
            validaciones.append({"estado":"OK","detalle":f'Título "{titulo_p}" ✔ presente en sección Conclusiones'})
        else:
            validaciones.append({"estado":"WARN","detalle":f'Título "{titulo_p}" ✘ NO aparece en Conclusiones — revisar párrafo introductorio'})

    # ── Validar ID Azure
    if id_tarea:
        val_id = valores_conc.get("id azure", "")
        if normalizar(id_tarea) in normalizar(val_id):
            validaciones.append({"estado":"OK","detalle":f'ID Azure "{id_tarea}" ✔ coincide'})
        elif es_placeholder(val_id):
            validaciones.append({"estado":"WARN","detalle":f'ID Azure aún es placeholder: "{val_id}"'})
        elif val_id == "":
            validaciones.append({"estado":"ERROR","detalle":"Celda ID Azure vacía"})
        else:
            validaciones.append({"estado":"ERROR","detalle":f'ID Azure esperado "{id_tarea}" ✘ no coincide — contiene: "{val_id}"'})

    # ── Validar consecutivo en conclusiones
    if consecutivo:
        val_consec = valores_conc.get("consecutivo", "")
        if normalizar(consecutivo) in normalizar(val_consec):
            validaciones.append({"estado":"OK","detalle":f'Consecutivo "{consecutivo}" ✔ coincide en conclusiones'})
        elif es_placeholder(val_consec):
            validaciones.append({"estado":"WARN","detalle":f'Consecutivo en conclusiones es placeholder: "{val_consec}"'})
        else:
            validaciones.append({"estado":"ERROR","detalle":f'Consecutivo esperado "{consecutivo}" ✘ no coincide — contiene: "{val_consec}"'})

    # ── Mostrar resultado y observaciones como info
    resultado = valores_conc.get("resultado", "")
    if resultado and not es_placeholder(resultado):
        estado_res = "OK" if resultado.upper() == "CERTIFICADA" else "WARN"
        validaciones.append({"estado":estado_res,"detalle":f'Resultado de la prueba: "{resultado}"'})

    estado_general = "ERROR" if any(v["estado"]=="ERROR" for v in validaciones) else \
                     "WARN"  if any(v["estado"]=="WARN"  for v in validaciones) else "OK"

    return {
        "titulo":       "Conclusiones",
        "estado":       estado_general,
        "fragmento":    fragmento,
        "validaciones": validaciones,
        "tipo":         "conclusiones",
    }

# ─────────────────────────────────────────────
# SECCIÓN 5 — COHERENCIA GLOBAL DE FECHAS
# ─────────────────────────────────────────────
def seccion_fechas(contenido):
    """Reúne TODAS las fechas del documento y verifica que sean coherentes."""
    tablas     = contenido["tablas"]
    encabezados= contenido["encabezados"]
    parrafos   = contenido["parrafos"]
    validaciones = []
    hallazgos  = {}  # { descripcion_ubicacion: fecha }

    # Encabezados
    for enc in encabezados:
        for f in extraer_fechas(enc):
            hallazgos[f"Encabezado — {enc[:50]}"] = f

    # Historial — todas las celdas
    try:
        for fi in [1, 2]:
            for ci, v in enumerate(tablas[2][fi]):
                if not es_placeholder(v) and v:
                    for f in extraer_fechas(v):
                        hallazgos[f"Historial fila {fi} — {v[:40]}"] = f
    except (IndexError, KeyError):
        pass

    # Resto de tablas
    for ti, tabla in enumerate(tablas):
        if ti == 2: continue
        for fi, fila in enumerate(tabla):
            for ci, v in enumerate(fila):
                if not es_placeholder(v) and v:
                    for f in extraer_fechas(v):
                        nombre_campo = tablas[ti][0][ci] if fi > 0 and tablas[ti] and ci < len(tablas[ti][0]) else f"col{ci}"
                        hallazgos[f"Tabla {ti}, {nombre_campo} — {v[:30]}"] = f

    # Párrafos
    for i, para in enumerate(parrafos):
        for f in extraer_fechas(para):
            hallazgos[f"Párrafo — {para[:50]}"] = f

    fragmento = [{"ubicacion": ub, "fecha": f} for ub, f in hallazgos.items()]

    if not hallazgos:
        return {
            "titulo":       "Coherencia de fechas",
            "estado":       "WARN",
            "fragmento":    [],
            "validaciones": [{"estado":"WARN","detalle":"Sin fechas concretas en el documento — todos los campos de fecha son placeholders"}],
        }

    todas = list(hallazgos.values())
    uniq  = list(set(todas))

    if len(uniq) == 1:
        validaciones.append({"estado":"OK","detalle":f'Todas las fechas son coherentes: "{uniq[0]}"'})
    else:
        validaciones.append({"estado":"ERROR","detalle":f'Fechas inconsistentes detectadas: {uniq}'})
        for ub, f in hallazgos.items():
            validaciones.append({"estado":"INFO","detalle":f'"{f}" en {ub}'})

    estado_general = "ERROR" if any(v["estado"]=="ERROR" for v in validaciones) else \
                     "WARN"  if any(v["estado"]=="WARN"  for v in validaciones) else "OK"

    return {
        "titulo":       "Coherencia de fechas",
        "estado":       estado_general,
        "fragmento":    fragmento,
        "validaciones": validaciones,
        "tipo":         "fechas",
    }

# ─────────────────────────────────────────────
# SECCIÓN 6 — ORTOGRAFÍA
# ─────────────────────────────────────────────
IGNORAR_PALABRAS = {
    # técnico seguridad
    "sanitizar","canonicalización","csrf","xss","sql","api","backend","frontend",
    "payload","bypass","exploit","fuzzing","pentesting","apikey","endpoint",
    "endpoints","token","tokens","bearer","oauth","jwt","https","http","cors",
    "json","xml","rest","soap","curl","headers","header","response","request",
    "client","server","proxy","timeout","redirect","cookie","cookies","script",
    "injection","buffer","overflow","encoding","hashing","hash","hmac","rsa",
    # inglés técnico común
    "access","application","applicant","mobile","cloud","devops","pipeline",
    # español formal que pyspellchecker falla
    "aplicaciones","autenticación","autorización","controles","funcionalidades",
    "vulnerabilidades","caracteres","conclusiones","autorizados","apropiados",
    "corresponden","ficheros","evaluada","incluyendo","periódica","confidencial",
    "telecomunicaciones","informática","módulos","distribución","codificar",
    "subversión","versión","sección","también","través","según","opciones",
    "deberán","tendrán","están","acceso","usuarios","servidores","sistemas",
    "seguridad","informe","documentación","actualización","revisión","creación",
    "generación","aprobación","modificación","versiones","registros","información",
    "gerencia","dirección","recomendaciones","implementación","validación",
    "implementando","modificaciones","observaciones","obtenidos","realizadas",
    "relacionadas","revisiones","necesarios","privilegios","recursos","salidas",
    "sesiones","siguientes","áreas","adicionalmente","aceptando","accesibles",
    "ambientes","amplifica","amplía","aparece","aislados","credenciales",
    "canales","clientes","críticas","debidas","derivados","activación","activos",
    # hashes y IDs — ignorar todo lo que parece hex
}

def es_hex_o_id(palabra):
    return bool(re.match(r'^[0-9a-f]{4,}$', palabra))

def verificar_ortografia(contenido):
    texto = contenido["texto_completo"]
    texto = re.sub(r'<[^>]+>', ' ', texto)
    texto = re.sub(r'https?://\S+', ' ', texto)
    texto = re.sub(r'\b[A-Z]{2,}\b', ' ', texto)
    texto = re.sub(r'\b\d[\d.,/-]*\b', ' ', texto)
    texto = re.sub(r'[^\w\sáéíóúÁÉÍÓÚñÑüÜ]', ' ', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()

    palabras = [p.lower() for p in texto.split() if len(p) > 3]
    palabras_revisar = [
        p for p in set(palabras)
        if p not in IGNORAR_PALABRAS and not es_hex_o_id(p) and len(p) > 3
    ]

    if not palabras_revisar:
        return {"titulo":"Ortografía","estado":"OK","fragmento":[],
                "validaciones":[{"estado":"OK","detalle":"Sin errores ortográficos detectados"}],
                "tipo":"ortografia"}

    sugerencias = []
    try:
        resultado_spell = {"data": None, "done": False}
        def run_spell():
            try:
                spell = SpellChecker(language='es')
                desc  = spell.unknown(palabras_revisar)
                errores = [p for p in desc if p not in IGNORAR_PALABRAS and not es_hex_o_id(p)]
                resultado_spell["data"] = (spell, errores)
            except Exception as ex:
                resultado_spell["data"] = ex
            finally:
                resultado_spell["done"] = True

        t = threading.Thread(target=run_spell, daemon=True)
        t.start()
        t.join(timeout=12)

        if not resultado_spell["done"] or isinstance(resultado_spell["data"], Exception):
            return {"titulo":"Ortografía","estado":"WARN","fragmento":[],
                    "validaciones":[{"estado":"WARN","detalle":"Revisión ortográfica omitida por timeout"}],
                    "tipo":"ortografia"}

        spell, errores = resultado_spell["data"]
        for palabra in sorted(errores)[:25]:
            sug = spell.correction(palabra)
            if sug and sug != palabra:
                sugerencias.append({"palabra": palabra, "sugerencia": sug})
            else:
                sugerencias.append({"palabra": palabra, "sugerencia": None})

    except Exception as e:
        return {"titulo":"Ortografía","estado":"WARN","fragmento":[],
                "validaciones":[{"estado":"WARN","detalle":f"Error: {str(e)}"}],
                "tipo":"ortografia"}

    validaciones = []
    if not sugerencias:
        validaciones.append({"estado":"OK","detalle":"Sin errores ortográficos detectados"})
    else:
        validaciones.append({"estado":"INFO","detalle":f"{len(sugerencias)} posible(s) error(es) — solo sugerencias, verificar manualmente"})

    return {
        "titulo":       "Ortografía",
        "estado":       "OK" if not sugerencias else "INFO",
        "fragmento":    sugerencias,
        "validaciones": validaciones,
        "tipo":         "ortografia",
    }

# ─────────────────────────────────────────────
# ENDPOINT PRINCIPAL
# ─────────────────────────────────────────────
@app.route('/verificar', methods=['POST', 'OPTIONS'])
def verificar():
    # Responder preflight CORS
    if request.method == 'OPTIONS':
        return jsonify({}), 200

    if 'file' not in request.files:
        return jsonify({"error": "No se recibió ningún archivo"}), 400

    file = request.files['file']
    if not file or not file.filename or not allowed_file(file.filename):
        return jsonify({"error": "Solo se aceptan archivos .docx"}), 400

    file_bytes = file.read()
    if len(file_bytes) == 0:
        return jsonify({"error": "El archivo está vacío"}), 400

    titulo_p    = request.form.get('titulo', '').strip()
    id_tarea    = request.form.get('id_tarea', '').strip()
    consecutivo = request.form.get('consecutivo', '').strip()

    try:
        contenido = extraer_contenido(file_bytes)
    except Exception as e:
        return jsonify({"error": f"No se pudo leer el documento: {str(e)}"}), 422

    secciones = [
        seccion_encabezado(contenido, titulo_p),
        seccion_info_documento(contenido, consecutivo),
        seccion_historial(contenido, titulo_p),
        seccion_conclusiones(contenido, titulo_p, id_tarea, consecutivo),
        seccion_fechas(contenido),
        verificar_ortografia(contenido),
    ]

    total_ok   = sum(1 for s in secciones if s["estado"] == "OK")
    total_warn = sum(1 for s in secciones if s["estado"] == "WARN")
    total_err  = sum(1 for s in secciones if s["estado"] == "ERROR")

    return jsonify({
        "secciones": secciones,
        "resumen": {"ok": total_ok, "alertas": total_warn, "errores": total_err}
    })

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok"})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
