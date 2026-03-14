import os
import uuid
import shutil
import tempfile
import zipfile
from io import BytesIO

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from flask import Flask, render_template, request, jsonify, send_file, send_from_directory
import json

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max upload

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), 'uploads')
FONTS_DIR = os.path.join(os.path.dirname(__file__), 'fonts')
os.makedirs(UPLOAD_DIR, exist_ok=True)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload-template', methods=['POST'])
def upload_template():
    """Upload a certificate template image. Returns image URL and dimensions."""
    if 'template' not in request.files:
        return jsonify(error='No template file provided'), 400

    f = request.files['template']
    if not f.filename:
        return jsonify(error='Empty filename'), 400

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ('.png', '.jpg', '.jpeg', '.webp'):
        return jsonify(error='Unsupported image format. Use PNG, JPG, or WEBP.'), 400

    filename = f'{uuid.uuid4().hex}{ext}'
    filepath = os.path.join(UPLOAD_DIR, filename)
    f.save(filepath)

    with Image.open(filepath) as img:
        width, height = img.size

    return jsonify(url=f'/uploads/{filename}', filename=filename, width=width, height=height)


@app.route('/upload-data', methods=['POST'])
def upload_data():
    """Upload CSV/XLSX data file. Returns detected column names."""
    if 'datafile' not in request.files:
        return jsonify(error='No data file provided'), 400

    f = request.files['datafile']
    if not f.filename:
        return jsonify(error='Empty filename'), 400

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ('.csv', '.xlsx', '.xls'):
        return jsonify(error='Unsupported file format. Use CSV or XLSX.'), 400

    filename = f'{uuid.uuid4().hex}{ext}'
    filepath = os.path.join(UPLOAD_DIR, filename)
    f.save(filepath)

    try:
        if ext == '.csv':
            df = pd.read_csv(filepath, nrows=5)
        else:
            df = pd.read_excel(filepath, nrows=5)
    except Exception as e:
        os.remove(filepath)
        return jsonify(error=f'Failed to read file: {e}'), 400

    columns = [str(c) for c in df.columns]
    preview = df.head(3).fillna('').astype(str).values.tolist()

    return jsonify(filename=filename, columns=columns, preview=preview)


@app.route('/uploads/<path:filename>')
def serve_upload(filename):
    return send_from_directory(UPLOAD_DIR, filename)


@app.route('/fonts-list')
def fonts_list():
    """Return available font files."""
    fonts = []
    if os.path.isdir(FONTS_DIR):
        for fname in sorted(os.listdir(FONTS_DIR)):
            if fname.lower().endswith(('.ttf', '.otf')):
                fonts.append(fname)
    return jsonify(fonts=fonts)


@app.route('/upload-font', methods=['POST'])
def upload_font():
    """Upload a custom font file to the fonts directory."""
    if 'font' not in request.files:
        return jsonify(error='No font file provided'), 400

    f = request.files['font']
    if not f.filename:
        return jsonify(error='Empty filename'), 400

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ('.ttf', '.otf'):
        return jsonify(error='Unsupported font format. Use TTF or OTF.'), 400

    # Keep original filename for readability
    safe_name = f.filename.replace(' ', '_')
    filepath = os.path.join(FONTS_DIR, safe_name)
    f.save(filepath)

    return jsonify(filename=safe_name)


def load_font(path, size):
    try:
        return ImageFont.truetype(path, size)
    except Exception:
        return ImageFont.load_default()


def draw_centered(draw, text, center_x, y, font, fill):
    if not text:
        return
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    x = int(center_x - text_width / 2)
    draw.text((x, y), text, font=font, fill=fill)


@app.route('/generate', methods=['POST'])
def generate():
    """Generate certificates and return as ZIP."""
    template_filename = request.form.get('template_filename')
    data_filename = request.form.get('data_filename')
    fields_json = request.form.get('fields')

    if not template_filename or not data_filename or not fields_json:
        return jsonify(error='Missing required fields'), 400

    template_path = os.path.join(UPLOAD_DIR, template_filename)
    data_path = os.path.join(UPLOAD_DIR, data_filename)

    if not os.path.exists(template_path) or not os.path.exists(data_path):
        return jsonify(error='Uploaded files not found. Please re-upload.'), 400

    try:
        fields = json.loads(fields_json)
    except json.JSONDecodeError:
        return jsonify(error='Invalid fields JSON'), 400

    if not fields:
        return jsonify(error='No fields mapped. Click on the template to place at least one field.'), 400

    # Read data
    ext = os.path.splitext(data_path)[1].lower()
    try:
        if ext == '.csv':
            df = pd.read_csv(data_path)
        else:
            df = pd.read_excel(data_path)
    except Exception as e:
        return jsonify(error=f'Failed to read data file: {e}'), 400

    columns = list(df.columns)

    # Preload fonts per field
    font_cache = {}
    for field in fields:
        font_name = field.get('font', '')
        font_size = int(field.get('fontSize', 30))
        cache_key = (font_name, font_size)
        if cache_key not in font_cache:
            font_path = os.path.join(FONTS_DIR, font_name) if font_name else ''
            if font_path and os.path.exists(font_path):
                font_cache[cache_key] = load_font(font_path, font_size)
            else:
                # Try first available font
                fallback = None
                if os.path.isdir(FONTS_DIR):
                    for fname in os.listdir(FONTS_DIR):
                        if fname.lower().endswith(('.ttf', '.otf')):
                            fallback = os.path.join(FONTS_DIR, fname)
                            break
                font_cache[cache_key] = load_font(fallback, font_size) if fallback else ImageFont.load_default()

    # Generate certificates
    tmp_dir = tempfile.mkdtemp()
    errors = []
    count = 0

    try:
        for idx, row in df.iterrows():
            try:
                img = Image.open(template_path).convert('RGB')
                draw = ImageDraw.Draw(img)

                # Use first mapped field's column value for filename
                cert_label = str(idx)
                for field in fields:
                    col = field.get('column', '')
                    val = ''
                    # Find matching column (case-insensitive)
                    for c in columns:
                        if c.strip().lower() == col.strip().lower():
                            val = row[c]
                            break
                    if pd.isna(val):
                        val = ''
                    else:
                        val = str(val).strip()

                    if val and cert_label == str(idx):
                        cert_label = '_'.join(val.split())

                    x = int(field.get('x', 0))
                    y = int(field.get('y', 0))
                    font_name = field.get('font', '')
                    font_size = int(field.get('fontSize', 30))
                    color = field.get('color', '#000000')

                    font = font_cache.get((font_name, font_size), ImageFont.load_default())
                    draw_centered(draw, val, x, y, font, fill=color)

                out_path = os.path.join(tmp_dir, f'certificate_{cert_label}.png')
                img.save(out_path, 'PNG')
                count += 1

            except Exception as e:
                errors.append(f'Row {idx}: {e}')

        if count == 0:
            return jsonify(error='No certificates generated. ' + '; '.join(errors)), 400

        # Create ZIP
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            for fname in os.listdir(tmp_dir):
                fpath = os.path.join(tmp_dir, fname)
                zf.write(fpath, fname)

        zip_buffer.seek(0)
        return send_file(zip_buffer, mimetype='application/zip',
                         as_attachment=True, download_name='certificates.zip')

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == '__main__':
    app.run(debug=True, port=8080)
