"""
Sistema de Registro de Reportes de Operación de Equipos
─────────────────────────────────────────────────────────
Objetivos resueltos en esta versión:
  1. Rendimiento: stats en caché con TTL, read_only para lecturas,
     diagnóstico de cold-start documentado.
  2. Prompt Engineering: deducción contextual para OCR en Ecuador.
  3. Plantilla Excel: usa copia.xlsx como base, hoja "CONTROL MAP",
     datos desde fila 7, bordes fieles al original.
"""

import os, re, json, base64, logging, requests, smtplib, ssl
import zipfile, threading, uuid, time, shutil, queue
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from flask import Flask, request, jsonify, render_template, send_file, Response, stream_with_context
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# ──────────────────────────────────────────────
#  Logging
# ──────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

# ──────────────────────────────────────────────
#  Flask & config
# ──────────────────────────────────────────────
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

UPLOAD_FOLDER  = 'uploads'
TEMPLATE_FILE  = 'template.xlsx'          # fuente inmutable — NUNCA se modifica
EXCEL_FILE     = 'outputs/Reporte_Operacion_Equipos.xlsx'
SHEET_NAME     = 'CONTROL MAP'
FIRST_DATA_ROW = 7                        # fila donde empiezan los datos (1-indexed)

GROQ_API_KEY    = os.environ.get('GROQ_API_KEY',    '')
SMTP_EMAIL      = os.environ.get('SMTP_EMAIL',      '')
SMTP_PASSWORD   = os.environ.get('SMTP_PASSWORD',   '')
VALIDATOR_EMAIL = os.environ.get('VALIDATOR_EMAIL', '')
ADMIN_PIN       = os.environ.get('ADMIN_PIN',       '1234')

# ──────────────────────────────────────────────
#  In-memory jobs & SSE
# ──────────────────────────────────────────────
jobs        = {}
sse_clients = []

# ──────────────────────────────────────────────
#  Stats cache (evita abrir el Excel en cada request)
#  Performance fix #1: TTL de 30 segundos
# ──────────────────────────────────────────────
_stats_cache      = {'registros': 0, 'imagenes': 0}
_stats_cache_time = 0.0
_STATS_TTL        = 30.0   # segundos

def _get_stats(force: bool = False) -> dict:
    global _stats_cache, _stats_cache_time
    now = time.monotonic()
    if force or (now - _stats_cache_time) > _STATS_TTL:
        try:
            if os.path.exists(EXCEL_FILE):
                wb    = openpyxl.load_workbook(EXCEL_FILE, read_only=True)  # read_only = rápido
                ws    = wb[SHEET_NAME]
                # Cuenta por col 2 (CAMPAMENTO) — más confiable que col 1 (MEGAZONA)
                count = sum(
                    1 for row in ws.iter_rows(
                        min_row=FIRST_DATA_ROW, min_col=2, max_col=2, values_only=True
                    ) if row[0] is not None
                )
                wb.close()
            else:
                count = 0
            img_count = len(os.listdir(UPLOAD_FOLDER)) if os.path.exists(UPLOAD_FOLDER) else 0
            _stats_cache      = {'registros': count, 'imagenes': img_count}
            _stats_cache_time = now
        except Exception as exc:
            log.warning("Stats read error: %s", exc)
    return _stats_cache

# ──────────────────────────────────────────────
#  Mapeo exacto: clave JSON → columna Excel (1-indexed)
#  Basado en row 6 del template copia.xlsx
# ──────────────────────────────────────────────
COLUMN_MAP = {
    "MEGAZONA":                1,   # A
    "CAMPAMENTO":              2,   # B
    "SECTOR":                  3,   # C
    "PISCINA":                 4,   # D
    "HECTÁREAS":               5,   # E
    "FECHA REAL DE INICIO":    6,   # F
    "FECHA DIARIO":            7,   # G  ← FECHA DE TRABAJO DIARIO
    "CATEGORÍA DE TRABAJO":    8,   # H
    "DESCRIPCIÓN DEL TRABAJO": 9,   # I
    "FECHA REAL DE FIN":       10,  # J
    "TIPO DE MAQUINARIA":      11,  # K
    "CÓDIGO DE MAQUINARIA":    12,  # L
    "CLASE DE MAQUINARIA":     13,  # M
    "PROVEEDOR":               14,  # N
    "No. COMPROBANTE":         15,  # O
    "RESPONSABLE DE REGISTRO": 16,  # P
    "HOROMETRO INICIAL":       17,  # Q
    "HOROMETRO FINAL":         18,  # R
    "MAÑANA hora inicio":      19,  # S
    "MAÑANA hora fin":         20,  # T
    "TARDE hora inicio":       21,  # U
    "TARDE hora fin":          22,  # V
    "NOCHE hora inicio":       23,  # W
    "NOCHE hora fin":          24,  # X
    "TOTAL HORAS":             25,  # Y
    "HORAS EXTRAS":            26,  # Z
    "CONSUMO DE DIESEL":       27,  # AA
    "% DE AVANCE":             28,  # AB
    "OBSERVACIONES":           29,  # AC
}

# Columnas con borde LEFT=medium (límites de sección en el template)
_MEDIUM_LEFT  = {1, 11, 17, 27}
# Columnas con borde RIGHT=medium
_MEDIUM_RIGHT = {5, 10, 16, 26, 29}

def _make_border(col: int) -> Border:
    thin   = Side(style='thin')
    medium = Side(style='medium')
    left   = medium if col in _MEDIUM_LEFT  else thin
    right  = medium if col in _MEDIUM_RIGHT else thin
    return Border(left=left, right=right, top=thin, bottom=thin)

# ──────────────────────────────────────────────
#  Excel — inicialización desde plantilla
# ──────────────────────────────────────────────
def init_excel():
    os.makedirs('outputs', exist_ok=True)
    if not os.path.exists(EXCEL_FILE):
        if os.path.exists(TEMPLATE_FILE):
            shutil.copy2(TEMPLATE_FILE, EXCEL_FILE)
            log.info("Excel inicializado desde plantilla: %s", EXCEL_FILE)
        else:
            # Fallback si no hay template: crear uno básico
            log.warning("template.xlsx no encontrado — creando Excel genérico")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = SHEET_NAME
            hdr_fill  = PatternFill("solid", fgColor="1F4E79")
            hdr_font  = Font(color="FFFFFF", bold=True, size=9, name="Calibri")
            hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for key, col_idx in COLUMN_MAP.items():
                cell           = ws.cell(row=FIRST_DATA_ROW - 1, column=col_idx, value=key)
                cell.fill      = hdr_fill
                cell.font      = hdr_font
                cell.alignment = hdr_align
                cell.border    = _make_border(col_idx)
            ws.freeze_panes = f"A{FIRST_DATA_ROW}"
            wb.save(EXCEL_FILE)


def append_to_excel(data: dict) -> int:
    """Añade una fila de datos al Excel y devuelve el número de registro (1-based)."""
    init_excel()
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb[SHEET_NAME]

    # Busca la primera fila vacía a partir de FIRST_DATA_ROW
    # Estrategia: usa max_row como base, luego verifica si tiene datos
    next_row = max(ws.max_row, FIRST_DATA_ROW)
    row_has_data = any(
        ws.cell(row=next_row, column=c).value is not None
        for c in range(1, len(COLUMN_MAP) + 1)
    )
    if row_has_data:
        next_row += 1

    # Estilo de fila: alternancia de fondo
    even_fill = PatternFill("solid", fgColor="EBF3FB")
    odd_fill  = PatternFill("solid", fgColor="FFFFFF")
    row_fill  = even_fill if (next_row % 2 == 0) else odd_fill
    data_font = Font(size=9, name="Calibri")
    data_align = Alignment(horizontal="center", vertical="center", wrap_text=False)

    for key, col_idx in COLUMN_MAP.items():
        value = data.get(key)
        cell  = ws.cell(row=next_row, column=col_idx, value=value)
        cell.font      = data_font
        cell.alignment = data_align
        cell.border    = _make_border(col_idx)
        cell.fill      = row_fill

    ws.row_dimensions[next_row].height = 18
    wb.save(EXCEL_FILE)

    record_num = next_row - FIRST_DATA_ROW + 1   # registro #1 = fila 7
    log.info("Excel actualizado — fila %s, registro #%s", next_row, record_num)
    _get_stats(force=True)   # invalida caché inmediatamente
    return record_num


# ──────────────────────────────────────────────
#  Prompt de extracción (Prompt Engineering avanzado)
#  Objetivo: deducción lógica + contexto Ecuador
# ──────────────────────────────────────────────
EXTRACT_PROMPT = r"""Eres un experto en digitalización de formularios físicos de maquinaria pesada para una empresa ecuatoriana de movimiento de tierras, excavación, nivelación y construcción.

## CONTEXTO GEOGRÁFICO Y OPERATIVO
- Empresa opera en Ecuador, principalmente en provincias del Guayas, Los Ríos, Manabí y Pichincha.
- Ciudades frecuentes en los formularios: Guayaquil, Durán, Daule, Samborondón, Naranjito, Milagro, El Triunfo, Quevedo, Babahoyo, Vinces, Ventanas, Santo Domingo.
- Tipos de maquinaria habituales: excavadora, retroexcavadora, bulldozer, motoniveladora, compactadora, vibrocompactador, volquete, cargadora frontal, minicargadora, grúa, generador.
- Tipos de trabajo habituales: excavación, relleno, compactación, nivelación, lastrado, desbroce, carga y transporte, instalación de tubería, movimiento de tierras.
- Proveedores son empresas o personas con RUC ecuatoriano (número de 13 dígitos).

## REGLAS DE CORRECCIÓN OCR (aplica deducción lógica)
1. Si lees un nombre de ciudad que suena similar a una ciudad ecuatoriana real, corrígelo (ej. "Debian" → "Durán", "Daula" → "Daule", "Guayaquil" con tildes raras → "Guayaquil").
2. Los horómetros son números enteros o decimales de 4-6 dígitos (ej. 1723, 4521.5).
3. Las horas de trabajo son del formato HH:MM o descripciones como "8am", "12:00", "13:00".
4. Las fechas están escritas en español ("3 de marzo" → "03/03/2025") o en formato dd/mm/aaaa.
5. Si lees letras ambiguas en nombres propios, elige la opción que tenga más sentido como nombre real en español.
6. El campo "% DE AVANCE" es un porcentaje entre 0 y 100.
7. "CONSUMO DE DIESEL" es un número en galones o litros.
8. "CLASE DE MAQUINARIA" suele ser "Propia" o "Alquilada".

## MAPEO DE CAMPOS DEL FORMULARIO FÍSICO → CLAVES JSON
El formulario físico usa etiquetas distintas a las claves JSON. Aplica este mapeo:
| Etiqueta en el formulario | Clave JSON destino |
|---|---|
| CAMPAMENTO | CAMPAMENTO |
| FECHA / FECHA DE TRABAJO | FECHA DIARIO |
| RUC (del encabezado) | No. COMPROBANTE |
| OBRA / Descripción de obra | DESCRIPCIÓN DEL TRABAJO |
| EQUIPO / Máquina | TIPO DE MAQUINARIA |
| TIPO DE TRABAJO / Actividad | CATEGORÍA DE TRABAJO |
| COMBUSTIBLE / DIESEL | CONSUMO DE DIESEL |
| HORÓMETRO I / Inicial | HOROMETRO INICIAL |
| HORÓMETRO F / Final | HOROMETRO FINAL |
| MAÑANA DE → | MAÑANA hora inicio |
| MAÑANA A → | MAÑANA hora fin |
| TARDE DE → | TARDE hora inicio |
| TARDE A → | TARDE hora fin |
| NOCHE DE → | NOCHE hora inicio |
| NOCHE A → | NOCHE hora fin |
| TOTAL HORAS TRABAJADAS | TOTAL HORAS |
| OBSERVACIONES | OBSERVACIONES |
| No / Nº (número del reporte) | No. COMPROBANTE |

## INSTRUCCIONES DE SALIDA
Devuelve ÚNICAMENTE un bloque JSON válido con exactamente estas 29 claves.
Usa null si el campo no es legible o no existe en el formulario.
NO incluyas texto adicional, explicaciones ni bloques markdown.

{
  "MEGAZONA": null,
  "CAMPAMENTO": null,
  "SECTOR": null,
  "PISCINA": null,
  "HECTÁREAS": null,
  "FECHA REAL DE INICIO": null,
  "FECHA DIARIO": null,
  "CATEGORÍA DE TRABAJO": null,
  "DESCRIPCIÓN DEL TRABAJO": null,
  "FECHA REAL DE FIN": null,
  "TIPO DE MAQUINARIA": null,
  "CÓDIGO DE MAQUINARIA": null,
  "CLASE DE MAQUINARIA": null,
  "PROVEEDOR": null,
  "No. COMPROBANTE": null,
  "RESPONSABLE DE REGISTRO": null,
  "HOROMETRO INICIAL": null,
  "HOROMETRO FINAL": null,
  "MAÑANA hora inicio": null,
  "MAÑANA hora fin": null,
  "TARDE hora inicio": null,
  "TARDE hora fin": null,
  "NOCHE hora inicio": null,
  "NOCHE hora fin": null,
  "TOTAL HORAS": null,
  "HORAS EXTRAS": null,
  "CONSUMO DE DIESEL": null,
  "% DE AVANCE": null,
  "OBSERVACIONES": null
}"""

# Modelos en orden de preferencia (fallback automático)
VISION_MODELS = [
    "meta-llama/llama-4-scout-17b-16e-instruct",
    "meta-llama/llama-4-maverick-17b-128e-instruct",
]


def _parse_json_safe(text: str) -> dict:
    text = text.strip()
    # 1. directo
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    # 2. sin markdown fences
    clean = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
    clean = re.sub(r"\s*```$", "", clean).strip()
    try:
        return json.loads(clean)
    except json.JSONDecodeError:
        pass
    # 3. extrae primer objeto JSON con regex
    m = re.search(r"\{[\s\S]*\}", text)
    if m:
        try:
            return json.loads(m.group())
        except json.JSONDecodeError:
            pass
    raise ValueError(f"No se pudo parsear JSON. Respuesta (primeros 300 chars): {text[:300]}")


def extract_with_groq(image_base64: str, mime_type: str) -> dict:
    if not GROQ_API_KEY:
        raise RuntimeError("GROQ_API_KEY no configurado en las variables de entorno.")

    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    payload = {
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image_url",
                 "image_url": {"url": f"data:{mime_type};base64,{image_base64}"}},
                {"type": "text", "text": EXTRACT_PROMPT},
            ],
        }],
        "max_tokens": 1200,
        "temperature": 0.05,
    }

    last_error = None
    for model in VISION_MODELS:
        payload["model"] = model
        log.info("Intentando modelo: %s", model)
        try:
            resp = requests.post(
                "https://api.groq.com/openai/v1/chat/completions",
                headers=headers, json=payload, timeout=120,
            )
            if resp.status_code == 200:
                content = resp.json()["choices"][0]["message"]["content"]
                log.info("Respuesta IA (preview): %s", content[:300])
                return _parse_json_safe(content)
            last_error = f"HTTP {resp.status_code}: {resp.text[:300]}"
            log.warning("Modelo %s falló: %s", model, last_error)
        except Exception as exc:
            last_error = str(exc)
            log.warning("Modelo %s excepción: %s", model, exc)

    raise RuntimeError(f"Todos los modelos fallaron. Último error: {last_error}")


# ──────────────────────────────────────────────
#  Email (3 estrategias SMTP)
# ──────────────────────────────────────────────
def send_email_with_image(image_data: bytes, filename: str, proveedor_name: str):
    if not all([SMTP_EMAIL, SMTP_PASSWORD, VALIDATOR_EMAIL]):
        log.warning("Email omitido — credenciales SMTP incompletas.")
        return

    msg           = MIMEMultipart()
    msg['From']   = SMTP_EMAIL
    msg['To']     = VALIDATOR_EMAIL
    msg['Subject'] = f"Nuevo Reporte — {proveedor_name} — {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    body = (
        f"Se registró un nuevo reporte de operación de equipos.\n\n"
        f"Proveedor  : {proveedor_name}\n"
        f"Fecha/Hora : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n"
        f"Archivo    : {filename}\n\n"
        f"Los datos ya fueron guardados en el Excel acumulativo."
    )
    msg.attach(MIMEText(body, 'plain', 'utf-8'))
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(image_data)
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
    msg.attach(part)
    raw = msg.as_string()

    attempts = [
        ("Office365 STARTTLS 587", lambda: _smtp_starttls("smtp.office365.com", 587, raw)),
        ("Gmail STARTTLS 587",     lambda: _smtp_starttls("smtp.gmail.com",      587, raw)),
        ("Office365 SSL 465",      lambda: _smtp_ssl("smtp.office365.com",       465, raw)),
    ]
    for label, fn in attempts:
        try:
            fn()
            log.info("Email enviado via %s a %s", label, VALIDATOR_EMAIL)
            return
        except Exception as exc:
            log.warning("Email strategy '%s' falló: %s", label, exc)
    log.error("Todas las estrategias de email fallaron.")


def _smtp_starttls(host, port, raw):
    ctx = ssl.create_default_context()
    with smtplib.SMTP(host, port, timeout=30) as s:
        s.ehlo(); s.starttls(context=ctx); s.ehlo()
        s.login(SMTP_EMAIL, SMTP_PASSWORD)
        s.sendmail(SMTP_EMAIL, VALIDATOR_EMAIL, raw)


def _smtp_ssl(host, port, raw):
    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL(host, port, context=ctx, timeout=30) as s:
        s.login(SMTP_EMAIL, SMTP_PASSWORD)
        s.sendmail(SMTP_EMAIL, VALIDATOR_EMAIL, raw)


# ──────────────────────────────────────────────
#  SSE broadcast
# ──────────────────────────────────────────────
def _broadcast_stats():
    stats   = _get_stats(force=True)
    payload = f"data: {json.dumps(stats)}\n\n"
    for q in list(sse_clients):
        try:
            q.put_nowait(payload)
        except Exception:
            pass


# ──────────────────────────────────────────────
#  Background job
# ──────────────────────────────────────────────
def process_job(job_id, image_base64, mime_type, image_data, filename, proveedor_name):
    try:
        jobs[job_id]['status'] = 'processing'
        extracted  = extract_with_groq(image_base64, mime_type)
        record_num = append_to_excel(extracted)
        # Email en hilo aparte para no bloquear el resultado
        threading.Thread(
            target=send_email_with_image,
            args=(image_data, filename, proveedor_name),
            daemon=True,
        ).start()
        jobs[job_id] = {'status': 'done', 'record': record_num, 'data': extracted}
        log.info("Job %s completado → registro #%s", job_id, record_num)
        _broadcast_stats()
    except Exception as exc:
        log.error("Job %s error: %s", job_id, exc)
        jobs[job_id] = {'status': 'error', 'error': str(exc)}


# ──────────────────────────────────────────────
#  Rutas públicas
# ──────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No se recibió ningún archivo'}), 400
    file = request.files['file']
    if not file.filename:
        return jsonify({'error': 'Archivo vacío'}), 400

    ext = file.filename.rsplit('.', 1)[-1].lower()
    if ext not in {'png', 'jpg', 'jpeg', 'webp'}:
        return jsonify({'error': 'Solo se aceptan imágenes (JPG, PNG, WEBP)'}), 400

    mime_map  = {'jpg': 'image/jpeg', 'jpeg': 'image/jpeg', 'png': 'image/png', 'webp': 'image/webp'}
    mime_type = mime_map[ext]

    proveedor_name = (request.form.get('proveedor', '') or 'Desconocido').strip()
    image_data     = file.read()
    image_base64   = base64.b64encode(image_data).decode()

    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    ts            = datetime.now().strftime('%Y%m%d_%H%M%S')
    safe_prov     = re.sub(r'[^\w\-]', '_', proveedor_name)[:40]
    filename      = f"{ts}_{safe_prov}.{ext}"
    with open(os.path.join(UPLOAD_FOLDER, filename), 'wb') as f:
        f.write(image_data)

    job_id       = str(uuid.uuid4())
    jobs[job_id] = {'status': 'queued'}
    threading.Thread(
        target=process_job,
        args=(job_id, image_base64, mime_type, image_data, filename, proveedor_name),
        daemon=True,
    ).start()
    log.info("Job %s encolado para '%s'", job_id, proveedor_name)
    return jsonify({'job_id': job_id})


@app.route('/status/<job_id>')
def status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({'status': 'not_found'}), 404
    return jsonify(job)


@app.route('/stats')
def stats():
    return jsonify(_get_stats())


@app.route('/stats/stream')
def stats_stream():
    def generator():
        q = queue.Queue()
        sse_clients.append(q)
        try:
            # envía estado actual al conectar
            payload = f"data: {json.dumps(_get_stats())}\n\n"
            yield payload
            while True:
                try:
                    data = q.get(timeout=25)
                    yield data
                except queue.Empty:
                    yield ": heartbeat\n\n"
        finally:
            try:
                sse_clients.remove(q)
            except ValueError:
                pass

    return Response(
        stream_with_context(generator()),
        mimetype='text/event-stream',
        headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'},
    )


# ──────────────────────────────────────────────
#  Rutas admin (protegidas por PIN)
# ──────────────────────────────────────────────
@app.route('/admin')
def admin():
    return render_template('admin.html')


@app.route('/admin/verify', methods=['POST'])
def verify_pin():
    data = request.get_json() or {}
    if data.get('pin') == ADMIN_PIN:
        return jsonify({'success': True})
    return jsonify({'success': False, 'error': 'PIN incorrecto'}), 401


def _pin_ok(req) -> bool:
    pin = req.headers.get('X-Admin-Pin') or req.args.get('pin', '')
    return pin == ADMIN_PIN


@app.route('/download-excel')
def download_excel():
    if not _pin_ok(request):
        return jsonify({'error': 'No autorizado'}), 401
    init_excel()
    return send_file(EXCEL_FILE, as_attachment=True,
                     download_name='Reporte_Operacion_Equipos.xlsx')


@app.route('/download-zip')
def download_zip():
    if not _pin_ok(request):
        return jsonify({'error': 'No autorizado'}), 401
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    zip_path = 'outputs/imagenes_reportes.zip'
    with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        for fname in os.listdir(UPLOAD_FOLDER):
            zf.write(os.path.join(UPLOAD_FOLDER, fname), fname)
    return send_file(zip_path, as_attachment=True,
                     download_name=f'imagenes_reportes_{datetime.now().strftime("%Y%m%d")}.zip')


# ──────────────────────────────────────────────
#  Entry point
# ──────────────────────────────────────────────
if __name__ == '__main__':
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs('outputs', exist_ok=True)
    init_excel()
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
