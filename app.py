import os
import json
import base64
import requests
import smtplib
import zipfile
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from flask import Flask, request, jsonify, render_template, send_file
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

UPLOAD_FOLDER = 'uploads'
EXCEL_FILE = 'outputs/Reporte_Operacion_Equipos.xlsx'

GROQ_API_KEY      = os.environ.get('GROQ_API_KEY', '')
SMTP_EMAIL        = os.environ.get('SMTP_EMAIL', '')
SMTP_PASSWORD     = os.environ.get('SMTP_PASSWORD', '')
VALIDATOR_EMAIL   = os.environ.get('VALIDATOR_EMAIL', '')
ADMIN_PIN         = os.environ.get('ADMIN_PIN', '1234')  # Cambia esto en Render

COLUMNS = [
    "MEGAZONA", "CAMPAMENTO", "SECTOR", "PISCINA", "HECTÁREAS",
    "FECHA REAL DE INICIO", "FECHA DE TRABAJO DIARIO", "CATEGORÍA DE TRABAJO",
    "DESCRIPCIÓN DEL TRABAJO", "FECHA REAL DE FIN", "TIPO DE MAQUINARIA",
    "CÓDIGO DE MAQUINARIA", "CLASE DE MAQUINARIA", "PROVEEDOR",
    "No. COMPROBANTE", "RESPONSABLE DE REGISTRO", "HOROMETRO INICIAL",
    "HOROMETRO FINAL", "MAÑANA hora inicio", "MAÑANA hora fin",
    "TARDE hora inicio", "TARDE hora fin", "NOCHE hora inicio", "NOCHE hora fin",
    "TOTAL HORAS", "HORAS EXTRAS", "CONSUMO DE DIESEL", "% DE AVANCE",
    "OBSERVACIONES"
]

def init_excel():
    os.makedirs('outputs', exist_ok=True)
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Reportes"
        header_fill = PatternFill("solid", fgColor="1F4E79")
        header_font = Font(color="FFFFFF", bold=True, size=9)
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin = Side(style="thin", color="AAAAAA")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for col_idx, col_name in enumerate(COLUMNS, 1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_align
            cell.border = border
            ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 16
        ws.row_dimensions[1].height = 45
        ws.freeze_panes = "A2"
        wb.save(EXCEL_FILE)

def extract_with_groq(image_base64, mime_type):
    prompt = """Analiza esta imagen de un Reporte de Operación de Equipos y extrae TODOS los campos visibles.
Devuelve ÚNICAMENTE un JSON válido con estos campos exactos (usa null si no puedes leer el valor):

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

Responde SOLO con el JSON, sin texto adicional, sin markdown."""

    headers = {
        "Authorization": f"Bearer {GROQ_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "meta-llama/llama-4-scout-17b-16e-instruct",
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{image_base64}"}},
                {"type": "text", "text": prompt}
            ]
        }],
        "max_tokens": 1000,
        "temperature": 0.1
    }
    response = requests.post(
        "https://api.groq.com/openai/v1/chat/completions",
        headers=headers, json=payload, timeout=30
    )
    if response.status_code != 200:
        raise Exception(f"Error Groq API: {response.text}")
    content = response.json()["choices"][0]["message"]["content"].strip()
    if content.startswith("```"):
        content = content.split("```")[1]
        if content.startswith("json"):
            content = content[4:]
    return json.loads(content)

def append_to_excel(data):
    init_excel()
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    next_row = ws.max_row + 1
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    row_fill = PatternFill("solid", fgColor="EBF3FB") if next_row % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
    for col_idx, col_name in enumerate(COLUMNS, 1):
        value = data.get(col_name, None)
        cell = ws.cell(row=next_row, column=col_idx, value=value)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
        cell.fill = row_fill
    ws.row_dimensions[next_row].height = 20
    wb.save(EXCEL_FILE)
    return next_row - 1

def send_email_with_image(image_data, filename, proveedor_name):
    if not all([SMTP_EMAIL, SMTP_PASSWORD, VALIDATOR_EMAIL]):
        print("Email no configurado, saltando envío.")
        return
    try:
        msg = MIMEMultipart()
        msg['From'] = SMTP_EMAIL
        msg['To'] = VALIDATOR_EMAIL
        msg['Subject'] = f"📋 Nuevo Reporte de Equipo - {proveedor_name} - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        body = f"""Se ha recibido un nuevo reporte de operación de equipo.

Proveedor: {proveedor_name}
Fecha y hora de recepción: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
Archivo: {filename}

Los datos ya han sido extraídos y añadidos al Excel acumulativo.
Puede descargar el Excel y el ZIP de imágenes desde la plataforma.
"""
        msg.attach(MIMEText(body, 'plain'))
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(image_data)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
        msg.attach(part)
        with smtplib.SMTP('smtp.office365.com', 587) as server:
            server.starttls()
            server.login(SMTP_EMAIL, SMTP_PASSWORD)
            server.sendmail(SMTP_EMAIL, VALIDATOR_EMAIL, msg.as_string())
    except Exception as e:
        print(f"Error enviando email: {e}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/admin')
def admin():
    return render_template('admin.html')

@app.route('/admin/verify', methods=['POST'])
def verify_pin():
    data = request.get_json()
    if data.get('pin') == ADMIN_PIN:
        return jsonify({'success': True})
    return jsonify({'success': False, 'error': 'PIN incorrecto'}), 401

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

    mime_types = {'jpg': 'image/jpeg', 'jpeg': 'image/jpeg', 'png': 'image/png', 'webp': 'image/webp'}
    mime_type = mime_types.get(ext, 'image/jpeg')

    proveedor_name = request.form.get('proveedor', 'Desconocido')
    image_data = file.read()
    image_base64 = base64.b64encode(image_data).decode('utf-8')

    # Save image
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    safe_proveedor = "".join(c for c in proveedor_name if c.isalnum() or c in (' ', '-', '_')).strip()
    filename = f"{timestamp}_{safe_proveedor}.{ext}"
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    with open(filepath, 'wb') as f:
        f.write(image_data)

    try:
        extracted = extract_with_groq(image_base64, mime_type)
        record_num = append_to_excel(extracted)
        send_email_with_image(image_data, filename, proveedor_name)
        return jsonify({
            'success': True,
            'record': record_num,
            'data': extracted
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download-excel')
def download_excel():
    pin = request.headers.get('X-Admin-Pin') or request.args.get('pin')
    if pin != ADMIN_PIN:
        return jsonify({'error': 'No autorizado'}), 401
    init_excel()
    return send_file(EXCEL_FILE, as_attachment=True,
                     download_name='Reporte_Operacion_Equipos.xlsx')

@app.route('/download-zip')
def download_zip():
    pin = request.headers.get('X-Admin-Pin') or request.args.get('pin')
    if pin != ADMIN_PIN:
        return jsonify({'error': 'No autorizado'}), 401
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    zip_path = 'outputs/imagenes_reportes.zip'
    with zipfile.ZipFile(zip_path, 'w') as zf:
        for fname in os.listdir(UPLOAD_FOLDER):
            zf.write(os.path.join(UPLOAD_FOLDER, fname), fname)
    return send_file(zip_path, as_attachment=True,
                     download_name=f'imagenes_reportes_{datetime.now().strftime("%Y%m%d")}.zip')

@app.route('/stats')
def stats():
    init_excel()
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    count = ws.max_row - 1
    img_count = len(os.listdir(UPLOAD_FOLDER)) if os.path.exists(UPLOAD_FOLDER) else 0
    return jsonify({'registros': count, 'imagenes': img_count})

if __name__ == '__main__':
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs('outputs', exist_ok=True)
    init_excel()
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
