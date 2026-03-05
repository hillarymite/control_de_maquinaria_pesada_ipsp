import os
import re
import json
import base64
import logging
import requests
import smtplib
import ssl
import zipfile
import threading
import uuid
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from flask import Flask, request, jsonify, render_template, send_file, Response, stream_with_context
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, GradientFill
from openpyxl.utils import get_column_letter
from datetime import datetime
import time

# ──────────────────────────────────────────────
#  Logging
# ──────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
log = logging.getLogger(__name__)

# ──────────────────────────────────────────────
#  App & Config
# ──────────────────────────────────────────────
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

UPLOAD_FOLDER = 'uploads'
EXCEL_FILE    = 'outputs/Reporte_Operacion_Equipos.xlsx'

GROQ_API_KEY    = os.environ.get('GROQ_API_KEY', '')
SMTP_EMAIL      = os.environ.get('SMTP_EMAIL', '')
SMTP_PASSWORD   = os.environ.get('SMTP_PASSWORD', '')
VALIDATOR_EMAIL = os.environ.get('VALIDATOR_EMAIL', '')
ADMIN_PIN       = os.environ.get('ADMIN_PIN', '1234')

# In-memory job store & SSE subscribers
jobs        = {}
sse_clients = []          # list of queue.Queue() objects

# ──────────────────────────────────────────────
#  Column definitions
# ──────────────────────────────────────────────
COLUMNS = [
    "MEGAZONA", "CAMPAMENTO", "SECTOR", "PISCINA", "HECTÁREAS",
    "FECHA REAL DE INICIO", "FECHA DE TRABAJO DIARIO", "CATEGORÍA DE TRABAJO",
    "DESCRIPCIÓN DEL TRABAJO", "FECHA REAL DE FIN", "TIPO DE MAQUINARIA",
    "CÓDIGO DE MAQUINARIA", "CLASE DE MAQUINARIA", "PROVEEDOR",
    "No. COMPROBANTE", "RESPONSABLE DE REGISTRO", "HOROMETRO INICIAL",
    "HOROMETRO FINAL", "MAÑANA hora inicio", "MAÑANA hora fin",
    "TARDE hora inicio", "TARDE hora fin", "NOCHE hora inicio", "NOCHE hora fin",
    "TOTAL HORAS", "HORAS EXTRAS", "CONSUMO DE DIESEL", "% DE AVANCE",
    "OBSERVACIONES",
]

# Custom widths for specific columns (others default to 16)
COLUMN_WIDTHS = {
    "MEGAZONA": 12, "CAMPAMENTO": 16, "SECTOR": 14, "PISCINA": 12,
    "HECTÁREAS": 11, "FECHA REAL DE INICIO": 18, "FECHA DE TRABAJO DIARIO": 20,
    "CATEGORÍA DE TRABAJO": 22, "DESCRIPCIÓN DEL TRABAJO": 30, "FECHA REAL DE FIN": 18,
    "TIPO DE MAQUINARIA": 22, "CÓDIGO DE MAQUINARIA": 20, "CLASE DE MAQUINARIA": 20,
    "PROVEEDOR": 24, "No. COMPROBANTE": 16, "RESPONSABLE DE REGISTRO": 24,
    "HOROMETRO INICIAL": 18, "HOROMETRO FINAL": 18,
    "MAÑANA hora inicio": 18, "MAÑANA hora fin": 16,
    "TARDE hora inicio": 16, "TARDE hora fin": 14,
    "NOCHE hora inicio": 16, "NOCHE hora fin": 14,
    "TOTAL HORAS": 14, "HORAS EXTRAS": 14,
    "CONSUMO DE DIESEL": 20, "% DE AVANCE": 14, "OBSERVACIONES": 32,
}

# ──────────────────────────────────────────────
#  Excel helpers
# ──────────────────────────────────────────────
def init_excel():
    os.makedirs('outputs', exist_ok=True)
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Reportes"

        # Header style
        hdr_fill  = PatternFill("solid", fgColor="1F4E79")
        hdr_font  = Font(color="FFFFFF", bold=True, size=9, name="Calibri")
        hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin      = Side(style="thin", color="AAAAAA")
        border    = Border(left=thin, right=thin, top=thin, bottom=thin)

        for col_idx, col_name in enumerate(COLUMNS, 1):
            cell           = ws.cell(row=1, column=col_idx, value=col_name)
            cell.fill      = hdr_fill
            cell.font      = hdr_font
            cell.alignment = hdr_align
            cell.border    = border
            ws.column_dimensions[get_column_letter(col_idx)].width = COLUMN_WIDTHS.get(col_name, 16)

        ws.row_dimensions[1].height = 50
        ws.freeze_panes = "A2"
        wb.save(EXCEL_FILE)
        log.info("Excel file created: %s", EXCEL_FILE)


def append_to_excel(data: dict) -> int:
    init_excel()
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    next_row = ws.max_row + 1

    thin      = Side(style="thin", color="DDDDDD")
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)
    even_fill = PatternFill("solid", fgColor="EBF3FB")
    odd_fill  = PatternFill("solid", fgColor="FFFFFF")
    row_fill  = even_fill if next_row % 2 == 0 else odd_fill

    for col_idx, col_name in enumerate(COLUMNS, 1):
        value         = data.get(col_name)
        cell          = ws.cell(row=next_row, column=col_idx, value=value)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border   = border
        cell.fill     = row_fill
        cell.font     = Font(size=9, name="Calibri")

    ws.row_dimensions[next_row].height = 20
    wb.save(EXCEL_FILE)
    log.info("Excel updated — row %s", next_row)
    return next_row - 1   # record number (1-based, excluding header)


# ──────────────────────────────────────────────
#  AI extraction (Groq — Llama 4 Scout vision)
# ──────────────────────────────────────────────
EXTRACT_PROMPT = """Eres un asistente especializado en digitalizar formularios físicos.
Analiza esta imagen de un Reporte de Operación de Equipos y extrae TODOS los campos visibles con máxima precisión.

Devuelve ÚNICAMENTE un JSON válido con exactamente estas claves (usa null si el campo no es legible):

{
  "MEGAZONA": null,
  "CAMPAMENTO": null,
  "SECTOR": null,
  "PISCINA": null,
  "HECTÁREAS": null,
  "FECHA REAL DE INICIO": null,
  "FECHA DE TRABAJO DIARIO": null,
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
}

Reglas importantes:
- Responde SOLO con el JSON. Cero texto adicional, cero bloques markdown, cero explicaciones.
- Las fechas deben quedar en formato DD/MM/YYYY si son legibles.
- Los valores numéricos (horas, horómetros) deben quedar como número o cadena numérica, sin unidades.
- Si un campo está vacío en el formulario, devuelve null.
"""

VISION_MODELS = [
    "meta-llama/llama-4-scout-17b-16e-instruct",
    "meta-llama/llama-4-maverick-17b-128e-instruct",
]


def _parse_json_response(text: str) -> dict:
    """Robust JSON extraction: handles markdown fences and surrounding text."""
    text = text.strip()

    # 1. Direct parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # 2. Strip markdown fences
    fenced = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
    fenced = re.sub(r"\s*```$", "", fenced).strip()
    try:
        return json.loads(fenced)
    except json.JSONDecodeError:
        pass

    # 3. Extract first JSON object via regex
    match = re.search(r"\{[\s\S]*\}", text)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass

    raise ValueError(f"Cannot parse JSON from response: {text[:300]}")


def extract_with_groq(image_base64: str, mime_type: str) -> dict:
    if not GROQ_API_KEY:
        raise RuntimeError("GROQ_API_KEY not configured")

    headers = {
        "Authorization": f"Bearer {GROQ_API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:{mime_type};base64,{image_base64}"},
                    },
                    {"type": "text", "text": EXTRACT_PROMPT},
                ],
            }
        ],
        "max_tokens": 1024,
        "temperature": 0.05,
    }

    last_error = None
    for model in VISION_MODELS:
        payload["model"] = model
        log.info("Trying model: %s", model)
        try:
            resp = requests.post(
                "https://api.groq.com/openai/v1/chat/completions",
                headers=headers,
                json=payload,
                timeout=120,
            )
            if resp.status_code == 200:
                content = resp.json()["choices"][0]["message"]["content"]
                log.info("Raw AI response: %s", content[:400])
                return _parse_json_response(content)
            else:
                last_error = f"HTTP {resp.status_code}: {resp.text[:300]}"
                log.warning("Model %s failed: %s", model, last_error)
        except Exception as exc:
            last_error = str(exc)
            log.warning("Model %s exception: %s", model, exc)

    raise RuntimeError(f"All models failed. Last error: {last_error}")


# ──────────────────────────────────────────────
#  Email (Office 365 / STARTTLS + SSL fallback)
# ──────────────────────────────────────────────
def send_email_with_image(image_data: bytes, filename: str, proveedor_name: str):
    if not all([SMTP_EMAIL, SMTP_PASSWORD, VALIDATOR_EMAIL]):
        log.warning("Email skipped — SMTP credentials not fully configured.")
        return

    msg             = MIMEMultipart()
    msg['From']     = SMTP_EMAIL
    msg['To']       = VALIDATOR_EMAIL
    msg['Subject']  = (
        f"Nuevo Reporte — {proveedor_name} — "
        f"{datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )
    body = (
        f"Se ha registrado un nuevo reporte de operación de equipos.\n\n"
        f"Proveedor  : {proveedor_name}\n"
        f"Fecha/Hora : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n"
        f"Archivo    : {filename}\n\n"
        f"Los datos extraídos ya fueron guardados en el Excel acumulativo.\n"
        f"Adjunto encontrará la imagen original del formulario."
    )
    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(image_data)
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
    msg.attach(part)

    raw = msg.as_string()

    # Strategy 1 — Office 365 STARTTLS (port 587)
    try:
        ctx = ssl.create_default_context()
        with smtplib.SMTP("smtp.office365.com", 587, timeout=30) as server:
            server.ehlo()
            server.starttls(context=ctx)
            server.ehlo()
            server.login(SMTP_EMAIL, SMTP_PASSWORD)
            server.sendmail(SMTP_EMAIL, VALIDATOR_EMAIL, raw)
        log.info("Email sent via Office365 STARTTLS to %s", VALIDATOR_EMAIL)
        return
    except Exception as exc:
        log.warning("Office365 STARTTLS failed: %s — trying Gmail STARTTLS", exc)

    # Strategy 2 — Gmail STARTTLS (port 587)
    try:
        ctx = ssl.create_default_context()
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as server:
            server.ehlo()
            server.starttls(context=ctx)
            server.ehlo()
            server.login(SMTP_EMAIL, SMTP_PASSWORD)
            server.sendmail(SMTP_EMAIL, VALIDATOR_EMAIL, raw)
        log.info("Email sent via Gmail STARTTLS to %s", VALIDATOR_EMAIL)
        return
    except Exception as exc:
        log.warning("Gmail STARTTLS failed: %s — trying SSL port 465", exc)

    # Strategy 3 — SSL port 465 (Office 365 / generic)
    try:
        ctx = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.office365.com", 465, context=ctx, timeout=30) as server:
            server.login(SMTP_EMAIL, SMTP_PASSWORD)
            server.sendmail(SMTP_EMAIL, VALIDATOR_EMAIL, raw)
        log.info("Email sent via SSL:465 to %s", VALIDATOR_EMAIL)
    except Exception as exc:
        log.error("All email strategies failed: %s", exc)


# ──────────────────────────────────────────────
#  SSE broadcast helper
# ──────────────────────────────────────────────
def _broadcast_stats():
    """Push updated stats to all connected SSE clients."""
    try:
        init_excel()
        wb       = openpyxl.load_workbook(EXCEL_FILE, read_only=True)
        ws       = wb.active
        count    = ws.max_row - 1
        img_count = len(os.listdir(UPLOAD_FOLDER)) if os.path.exists(UPLOAD_FOLDER) else 0
        wb.close()
        payload = json.dumps({"registros": count, "imagenes": img_count})
        data    = f"data: {payload}\n\n"
        for q in list(sse_clients):
            try:
                q.put_nowait(data)
            except Exception:
                pass
    except Exception as exc:
        log.warning("SSE broadcast error: %s", exc)


# ──────────────────────────────────────────────
#  Background job
# ──────────────────────────────────────────────
def process_job(job_id, image_base64, mime_type, image_data, filename, proveedor_name):
    try:
        jobs[job_id]['status'] = 'processing'
        log.info("Job %s — extracting with AI...", job_id)

        extracted  = extract_with_groq(image_base64, mime_type)
        record_num = append_to_excel(extracted)

        # Email in a separate daemon thread so it doesn't block the job result
        email_thread = threading.Thread(
            target=send_email_with_image,
            args=(image_data, filename, proveedor_name),
            daemon=True,
        )
        email_thread.start()

        jobs[job_id] = {'status': 'done', 'record': record_num, 'data': extracted}
        log.info("Job %s — done, record #%s", job_id, record_num)
        _broadcast_stats()

    except Exception as exc:
        log.error("Job %s — error: %s", job_id, exc)
        jobs[job_id] = {'status': 'error', 'error': str(exc)}


# ──────────────────────────────────────────────
#  Routes — public
# ──────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No se recibió ningún archivo'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Archivo vacío'}), 400

    allowed = {'png', 'jpg', 'jpeg', 'webp'}
    ext = file.filename.rsplit('.', 1)[-1].lower()
    if ext not in allowed:
        return jsonify({'error': 'Solo se aceptan imágenes (JPG, PNG, WEBP)'}), 400

    mime_map  = {'jpg': 'image/jpeg', 'jpeg': 'image/jpeg', 'png': 'image/png', 'webp': 'image/webp'}
    mime_type = mime_map.get(ext, 'image/jpeg')

    proveedor_name = request.form.get('proveedor', 'Desconocido').strip() or 'Desconocido'
    image_data     = file.read()
    image_base64   = base64.b64encode(image_data).decode('utf-8')

    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    timestamp      = datetime.now().strftime('%Y%m%d_%H%M%S')
    safe_proveedor = re.sub(r'[^\w\s\-]', '', proveedor_name).strip().replace(' ', '_')
    filename       = f"{timestamp}_{safe_proveedor}.{ext}"
    with open(os.path.join(UPLOAD_FOLDER, filename), 'wb') as f:
        f.write(image_data)

    job_id        = str(uuid.uuid4())
    jobs[job_id]  = {'status': 'queued'}
    t = threading.Thread(
        target=process_job,
        args=(job_id, image_base64, mime_type, image_data, filename, proveedor_name),
        daemon=True,
    )
    t.start()

    log.info("Job %s queued for proveedor '%s'", job_id, proveedor_name)
    return jsonify({'job_id': job_id})


@app.route('/status/<job_id>')
def status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({'status': 'not_found'}), 404
    return jsonify(job)


@app.route('/stats')
def stats():
    init_excel()
    wb        = openpyxl.load_workbook(EXCEL_FILE, read_only=True)
    ws        = wb.active
    count     = ws.max_row - 1
    img_count = len(os.listdir(UPLOAD_FOLDER)) if os.path.exists(UPLOAD_FOLDER) else 0
    wb.close()
    return jsonify({'registros': count, 'imagenes': img_count})


# ── Server-Sent Events endpoint for real-time stats ──
@app.route('/stats/stream')
def stats_stream():
    import queue

    def event_generator():
        q = queue.Queue()
        sse_clients.append(q)
        try:
            # Send current stats immediately on connect
            _broadcast_stats()
            while True:
                try:
                    data = q.get(timeout=25)
                    yield data
                except queue.Empty:
                    yield ": heartbeat\n\n"   # keep-alive comment
        finally:
            try:
                sse_clients.remove(q)
            except ValueError:
                pass

    return Response(
        stream_with_context(event_generator()),
        mimetype='text/event-stream',
        headers={
            'Cache-Control': 'no-cache',
            'X-Accel-Buffering': 'no',   # disable Nginx buffering on Render
        },
    )


# ──────────────────────────────────────────────
#  Routes — admin (PIN-protected)
# ──────────────────────────────────────────────
@app.route('/admin')
def admin():
    return render_template('admin.html')


@app.route('/admin/verify', methods=['POST'])
def verify_pin():
    data = request.get_json()
    if data and data.get('pin') == ADMIN_PIN:
        return jsonify({'success': True})
    return jsonify({'success': False, 'error': 'PIN incorrecto'}), 401


def _check_pin(req) -> bool:
    pin = req.headers.get('X-Admin-Pin') or req.args.get('pin')
    return pin == ADMIN_PIN


@app.route('/download-excel')
def download_excel():
    if not _check_pin(request):
        return jsonify({'error': 'No autorizado'}), 401
    init_excel()
    return send_file(
        EXCEL_FILE,
        as_attachment=True,
        download_name='Reporte_Operacion_Equipos.xlsx',
    )


@app.route('/download-zip')
def download_zip():
    if not _check_pin(request):
        return jsonify({'error': 'No autorizado'}), 401
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    zip_path = 'outputs/imagenes_reportes.zip'
    with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        for fname in os.listdir(UPLOAD_FOLDER):
            zf.write(os.path.join(UPLOAD_FOLDER, fname), fname)
    return send_file(
        zip_path,
        as_attachment=True,
        download_name=f'imagenes_reportes_{datetime.now().strftime("%Y%m%d")}.zip',
    )


# ──────────────────────────────────────────────
#  Entry point
# ──────────────────────────────────────────────
if __name__ == '__main__':
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs('outputs', exist_ok=True)
    init_excel()
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
