import os
import smtplib
from email.message import EmailMessage
import hashlib
import hmac
import mimetypes
import math
import zipfile
import json
import re
import base64
import pymupdf as fitz  # PyMuPDF for PDF operations & text-to-pdf rendering
from flask import Flask, render_template, request, send_file, jsonify , Response
from pypdf import PdfReader, PdfWriter
from PIL import Image, ImageOps, ImageFilter, ImageEnhance, ImageDraw
from functools import lru_cache
from uuid import uuid4
from io import BytesIO
import docx
import pdfplumber
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from werkzeug.utils import secure_filename
from optimize_ocr_backend import TOOLS as OPTIMIZE_OCR_TOOLS, process_optimize_ocr
from edit_secure_backend import TOOLS as EDIT_SECURE_TOOLS, process as process_edit_secure
from archive_backend import process_archive
# Additional Create PDF converters
import csv
import html
import xml.etree.ElementTree as ET
import ipaddress
import socket
import time
from urllib.parse import urlparse
from tempfile import TemporaryDirectory
from markdown import markdown as markdown_to_html
from bs4 import BeautifulSoup
from ebooklib import epub, ITEM_DOCUMENT
from weasyprint import HTML
import cairosvg


app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 2 * 1024 * 1024 * 1024  # 2 GB limit

# Lightweight per-process contact-form throttling. For multi-worker deployments,
# use a shared rate-limit store at the reverse-proxy/application layer.
_CONTACT_ATTEMPTS = {}
_CONTACT_WINDOW_SECONDS = 600
_CONTACT_MAX_ATTEMPTS = 5

UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@lru_cache(maxsize=1)
def get_bg_session():
    # Reuse the AI model between requests instead of loading it every time.
    from rembg import new_session
    return new_session('u2net')

# --- Home Route ---
@app.route('/')
def home():
    return render_template('index.html')

@app.route('/contact', methods=['POST'])
def contact_message():
    """Send homepage contact form messages to the FileVellum inbox."""
    name = request.form.get('name', '').strip()
    sender_email = request.form.get('email', '').strip()
    phone = request.form.get('phone', '').strip()
    source = request.form.get('source', '').strip()
    message = request.form.get('message', '').strip()

    if not name or not sender_email or not message:
        return jsonify({'success': False, 'error': 'Please fill in your name, email and message.'}), 400

    # Keep form input bounded and reject malformed email/header input.
    if len(name) > 100 or len(sender_email) > 254 or len(phone) > 40 or len(source) > 120 or len(message) > 10000:
        return jsonify({'success': False, 'error': 'One or more fields are too long. Please shorten your message and try again.'}), 400
    if not re.fullmatch(r"[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+", sender_email):
        return jsonify({'success': False, 'error': 'Please enter a valid email address.'}), 400
    if any(ch in name for ch in '\r\n') or any(ch in sender_email for ch in '\r\n'):
        return jsonify({'success': False, 'error': 'Please check the name and email fields.'}), 400

    # Basic abuse protection. This is intentionally conservative and does not
    # change the contact form's appearance or normal submission flow.
    now = time.monotonic()
    client_ip = request.remote_addr or 'unknown'
    attempts = [stamp for stamp in _CONTACT_ATTEMPTS.get(client_ip, []) if now - stamp < _CONTACT_WINDOW_SECONDS]
    if len(attempts) >= _CONTACT_MAX_ATTEMPTS:
        _CONTACT_ATTEMPTS[client_ip] = attempts
        return jsonify({'success': False, 'error': 'Too many messages have been sent. Please wait a few minutes and try again.'}), 429
    attempts.append(now)
    _CONTACT_ATTEMPTS[client_ip] = attempts
    if len(_CONTACT_ATTEMPTS) > 10000:
        # Prevent unbounded growth in a long-running process.
        for ip, stamps in list(_CONTACT_ATTEMPTS.items()):
            fresh = [stamp for stamp in stamps if now - stamp < _CONTACT_WINDOW_SECONDS]
            if fresh:
                _CONTACT_ATTEMPTS[ip] = fresh
            else:
                _CONTACT_ATTEMPTS.pop(ip, None)

    # Credentials are supplied through environment variables; never hard-code Gmail passwords.
    smtp_user = os.environ.get('FILEVELLUM_EMAIL_USER', '').strip()
    smtp_password = os.environ.get('FILEVELLUM_EMAIL_APP_PASSWORD', '').strip().replace(' ', '')
    recipient = os.environ.get('FILEVELLUM_CONTACT_TO', 'filevellum@gmail.com').strip()

    if not smtp_user or not smtp_password:
        app.logger.error('Contact email is not configured: missing FILEVELLUM_EMAIL_USER or FILEVELLUM_EMAIL_APP_PASSWORD.')
        return jsonify({'success': False, 'error': 'Email service is not configured yet. Please try again later.'}), 503

    email_message = EmailMessage()
    email_message['Subject'] = f'FileVellum Contact Message from {name}'
    email_message['From'] = smtp_user
    email_message['To'] = recipient
    email_message['Reply-To'] = sender_email
    email_message.set_content(
        f'New message received from the FileVellum website.\n\n'
        f'Name: {name}\n'
        f'Email: {sender_email}\n'
        f'Phone: {phone or "Not provided"}\n'
        f'How they found us: {source or "Not provided"}\n\n'
        f'Message:\n{message}\n'
    )

    try:
        with smtplib.SMTP('smtp.gmail.com', 587, timeout=20) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(smtp_user, smtp_password)
            server.send_message(email_message)
        return jsonify({'success': True, 'message': 'Your message has been sent successfully.'})
    except Exception:
        app.logger.exception('Could not send FileVellum contact email.')
        return jsonify({'success': False, 'error': 'Message could not be sent right now. Please try again later.'}), 502

@app.route('/about-founders')
def about_founders():
    return send_file(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'about_founders.html'))

# --- All Dedicated Tool Page Routes ---
@app.route('/compressor')
def compressor(): return render_template('tool_template.html', tool_name="ZIP Compressor", tool_type="compressor")

@app.route('/extract-zip')
def extract_zip(): return render_template('tool_template.html', tool_name="Extract ZIP", tool_type="extract_zip")

@app.route('/edit-image')
def edit_image(): return render_template('tool_template.html', tool_name='Edit Image', tool_type='edit_image')

@app.route('/merge-pdf')
def merge_pdf(): return render_template('tool_template.html', tool_name="Merge PDF", tool_type="merge")

@app.route('/split-pdf')
def split_pdf(): return render_template('tool_template.html', tool_name="Split PDF", tool_type="split")

@app.route('/compress-pdf')
def compress_pdf(): return render_template('tool_template.html', tool_name="Compress PDF", tool_type="compress_pdf")

@app.route('/compress-file')
def compress_file(): return render_template('tool_template.html', tool_name="Compress PDF, JPG & PNG", tool_type="compress_file")

@app.route('/pdf-to-word')
def pdf_to_word(): return render_template('tool_template.html', tool_name="PDF to Word", tool_type="pdf_to_word")

@app.route('/pdf-to-excel')
def pdf_to_excel(): return render_template('tool_template.html', tool_name="PDF to Excel", tool_type="pdf_to_excel")

@app.route('/pdf-to-powerpoint')
def pdf_to_ppt(): return render_template('tool_template.html', tool_name="PDF to PowerPoint", tool_type="pdf_to_powerpoint")

@app.route('/pdf-to-jpg')
def pdf_to_jpg(): return render_template('tool_template.html', tool_name="PDF to JPG", tool_type="pdf_to_jpg")

@app.route('/pdf-to-png')
def pdf_to_png(): return render_template('tool_template.html', tool_name="PDF to PNG", tool_type="pdf_to_png")

@app.route('/word-to-pdf')
def word_to_pdf(): return render_template('tool_template.html', tool_name="Word to PDF", tool_type="word_to_pdf")

@app.route('/excel-to-pdf')
def excel_to_pdf(): return render_template('tool_template.html', tool_name="Excel to PDF", tool_type="excel_to_pdf")

@app.route('/powerpoint-to-pdf')
def ppt_to_pdf(): return render_template('tool_template.html', tool_name="PowerPoint to PDF", tool_type="powerpoint_to_pdf")

@app.route('/jpg-to-pdf')
def jpg_to_pdf(): return render_template('tool_template.html', tool_name="JPG to PDF", tool_type="jpg_to_pdf")

@app.route('/png-to-pdf')
def png_to_pdf(): return render_template('tool_template.html', tool_name="PNG to PDF", tool_type="png_to_pdf")

@app.route('/png-to-jpg')
def png_to_jpg(): return render_template('tool_template.html', tool_name="PNG to JPG", tool_type="png_to_jpg")

@app.route('/jpg-to-png')
def jpg_to_png(): return render_template('tool_template.html', tool_name="JPG to PNG", tool_type="jpg_to_png")

@app.route('/edit-pdf')
def edit_pdf(): return render_template('tool_template.html', tool_name="Edit PDF", tool_type="edit_pdf")

@app.route('/sign-pdf')
def sign_pdf(): return render_template('tool_template.html', tool_name='Sign PDF', tool_type='sign_pdf')

@app.route('/watermark')
def watermark(): return render_template('tool_template.html', tool_name='Watermark', tool_type='watermark')

@app.route('/rotate-pdf')
def rotate_pdf(): return render_template('tool_template.html', tool_name='Rotate PDF', tool_type='rotate_pdf')

@app.route('/protect-pdf')
def protect_pdf(): return render_template('tool_template.html', tool_name='Protect PDF', tool_type='protect_pdf')

@app.route('/unlock-pdf')
def unlock_pdf(): return render_template('tool_template.html', tool_name='Unlock PDF', tool_type='unlock_pdf')

@app.route('/organize-pdf')
def organize_pdf(): return render_template('tool_template.html', tool_name="Organize Pages", tool_type="organize_pdf")

# Background Removal Route
@app.route('/remove-bg')
def remove_bg(): return render_template('tool_template.html', tool_name="Background Removal", tool_type="remove_bg")



# ZIP & Archives tools use the existing shared template and /api/process pipeline.
ARCHIVE_TOOLS = {
    'create-zip-archive': ('Create ZIP Archive', 'create_zip_archive'),
    'extract-zip-archive': ('Extract ZIP Archive', 'extract_zip_archive'),
    'compress-files-to-zip': ('Compress Files to ZIP', 'compress_files_to_zip'),
    'extract-multiple-files': ('Extract Multiple Files', 'extract_multiple_files'),
    'merge-zip-archives': ('Merge ZIP Archives', 'merge_zip_archives'),
    'split-large-files-into-parts': ('Split Large Files into Parts', 'split_large_files_into_parts'),
    'password-protect-zip': ('Password Protect ZIP', 'password_protect_zip'),
    'extract-password-protected-zip': ('Extract Password-Protected ZIP', 'extract_password_protected_zip'),
    'create-tar-archive': ('Create TAR Archive', 'create_tar_archive'),
    'extract-tar-archive': ('Extract TAR Archive', 'extract_tar_archive'),
    'create-tar-gz-archive': ('Create TAR.GZ Archive', 'create_tar_gz_archive'),
    'extract-tar-gz-archive': ('Extract TAR.GZ Archive', 'extract_tar_gz_archive'),
    'zip-file-information': ('ZIP File Information', 'zip_file_information'),
    'zip-to-tar': ('Convert ZIP to TAR', 'zip_to_tar'),
    'tar-to-zip': ('Convert TAR to ZIP', 'tar_to_zip'),
}

@app.route('/archive-tool/<tool_slug>')
def archive_tool_page(tool_slug):
    item = ARCHIVE_TOOLS.get(tool_slug)
    if not item:
        return 'Tool not found', 404
    return render_template('tool_template.html', tool_name=item[0], tool_type=item[1])

# Image Editing tools — all render through the existing shared tool_template.html
IMAGE_EDIT_TOOLS = {
    'image-compressor': 'Image Compressor',
    'image-resizer': 'Image Resizer',
    'image-cropper': 'Image Cropper',
    'image-rotator': 'Image Rotator',
    'image-flipper': 'Image Flipper',
    'image-watermark': 'Image Watermark',
    'passport-photo-maker': 'Passport Photo Maker',
    'blur-image': 'Blur Image',
    'sharpen-image': 'Sharpen Image',
    'image-filters': 'Image Color Filters',
    'image-to-text': 'Image to Text OCR',
    'image-metadata-viewer': 'Image Metadata Viewer',
    'remove-image-metadata': 'Remove Image Metadata',
    'image-background-changer': 'Image Background Color Changer',
    'image-border-maker': 'Image Border Maker',
    'image-collage-maker': 'Image Collage Maker',
    'image-upscaler': 'Image Upscaler',
    'image-color-palette': 'Image Color Palette Extractor',
}

@app.route('/image-tool/<tool_slug>')
def image_edit_tool_page(tool_slug):
    tool_type = tool_slug.replace('-', '_')
    if tool_slug not in IMAGE_EDIT_TOOLS:
        return 'Tool not found', 404
    return render_template('tool_template.html', tool_name=IMAGE_EDIT_TOOLS[tool_slug], tool_type=tool_type)

# --- Image Conversion tools ---
IMAGE_CONVERSION_TOOLS = {
    'jpg-to-png': ('JPG to PNG', 'jpg_to_png'),
    'png-to-jpg': ('PNG to JPG', 'png_to_jpg'),
    'images-to-webp': ('Images to WebP', 'images_to_webp'),
    'webp-to-jpg': ('WebP to JPG', 'webp_to_jpg'),
    'webp-to-png': ('WebP to PNG', 'webp_to_png'),
    'svg-to-png': ('SVG to PNG', 'svg_to_png'),
    'png-to-svg': ('PNG to SVG', 'png_to_svg'),
    'heic-to-jpg': ('HEIC to JPG', 'heic_to_jpg'),
    'image-to-base64': ('Image to Base64', 'image_to_base64'),
    'base64-to-image': ('Base64 to Image', 'base64_to_image'),
    'gif-to-jpg': ('GIF to JPG', 'gif_to_jpg'),
    'gif-maker': ('GIF Maker', 'gif_maker'),
    'avif-to-jpg': ('AVIF to JPG', 'avif_to_jpg'),
    'jpg-to-avif': ('JPG to AVIF', 'jpg_to_avif'),
    'webp-to-avif': ('WebP to AVIF', 'webp_to_avif'),
    'avif-to-png': ('AVIF to PNG', 'avif_to_png'),
    'image-to-ico': ('Image to ICO', 'image_to_ico'),
    'ico-to-png': ('ICO to PNG', 'ico_to_png'),
    'favicon-generator': ('Favicon Generator', 'favicon_generator'),
    'bmp-to-png': ('BMP to PNG', 'bmp_to_png'),
}

@app.route('/image-convert/<tool_slug>')
def image_conversion_page(tool_slug):
    if tool_slug not in IMAGE_CONVERSION_TOOLS:
        return 'Tool not found', 404
    tool_name, tool_type = IMAGE_CONVERSION_TOOLS[tool_slug]
    return render_template('tool_template.html', tool_name=tool_name, tool_type=tool_type)


def process_image_conversion(tool_type, files, form, upload_folder):
    """Convert image formats for the Image Conversion catalog."""
    from PIL import Image
    from io import BytesIO
    import zipfile
    import base64
    import uuid
    try:
        import pillow_avif  # noqa: F401 - registers AVIF support with Pillow
    except ImportError:
        pass

    valid_images = {'.jpg', '.jpeg', '.png', '.webp', '.bmp', '.gif', '.ico', '.avif', '.heic', '.heif', '.tif', '.tiff'}
    named = [f for f in files if f and f.filename]
    if tool_type == 'base64_to_image':
        raw_text = (form.get('base64_data') or '').strip()
        if not raw_text:
            return jsonify(error='Paste a Base64 image string first.'), 400
        if ',' in raw_text and raw_text.lower().startswith('data:'):
            raw_text = raw_text.split(',', 1)[1]
        try:
            payload = base64.b64decode(raw_text, validate=True)
            image = Image.open(BytesIO(payload))
            image.verify()
        except Exception:
            return jsonify(error='The Base64 text is invalid or does not contain a supported image.'), 400
        image = Image.open(BytesIO(payload))
        ext = (image.format or 'PNG').lower()
        ext = 'jpg' if ext == 'jpeg' else ext
        filename = f'{uuid.uuid4().hex}_converted.{ext}'
        image.save(os.path.join(upload_folder, filename))
        return jsonify(success=True, filename=filename)

    if not named:
        return jsonify(error='Please upload an image file.'), 400
    if any(os.path.splitext(f.filename)[1].lower() in ('.heic', '.heif') for f in named):
        try:
            import pillow_heif
            pillow_heif.register_heif_opener()
        except ImportError:
            return jsonify(error='HEIC support is not installed. Run: python -m pip install pillow-heif'), 503
    if tool_type == 'gif_maker' and len(named) < 2:
        return jsonify(error='Select at least two images to create a GIF.'), 400
    if tool_type == 'images_to_webp' and len(named) > 1:
        output_names = []
        for f in named:
            ext = os.path.splitext(f.filename)[1].lower()
            if ext not in valid_images:
                return jsonify(error='One or more files are not supported images.'), 400
            try:
                im = Image.open(f).convert('RGB')
                outname = f'{uuid.uuid4().hex}_{os.path.splitext(os.path.basename(f.filename))[0]}.webp'
                im.save(os.path.join(upload_folder, outname), 'WEBP', quality=90, method=6)
                output_names.append(outname)
            except Exception as exc:
                return jsonify(error='Could not convert one of the images: ' + str(exc)[:150]), 400
        zipname = f'{uuid.uuid4().hex}_webp_images.zip'
        with zipfile.ZipFile(os.path.join(upload_folder, zipname), 'w', zipfile.ZIP_DEFLATED) as zf:
            for name in output_names:
                zf.write(os.path.join(upload_folder, name), name)
        return jsonify(success=True, filename=zipname)

    f = named[0]
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in valid_images and not (tool_type == 'svg_to_png' and ext == '.svg'):
        return jsonify(error='Unsupported image format. Please upload a supported image.'), 400
    raw = f.read()
    if not raw or len(raw) > 50 * 1024 * 1024:
        return jsonify(error='Image must be non-empty and no larger than 50 MB.'), 400
    stem = os.path.splitext(os.path.basename(f.filename))[0] or 'image'
    out = None
    out_format = None

    try:
        if tool_type == 'svg_to_png':
            import cairosvg
            out_bytes = cairosvg.svg2png(bytestring=raw)
            outname = f'{uuid.uuid4().hex}_{stem}.png'
            with open(os.path.join(upload_folder, outname), 'wb') as fp:
                fp.write(out_bytes)
            return jsonify(success=True, filename=outname)

        if tool_type == 'png_to_svg':
            image = Image.open(BytesIO(raw)).convert('RGBA')
            import base64 as b64
            buffer = BytesIO(); image.save(buffer, format='PNG')
            data = b64.b64encode(buffer.getvalue()).decode('ascii')
            svg = f'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="{image.width}" height="{image.height}" viewBox="0 0 {image.width} {image.height}"><image width="{image.width}" height="{image.height}" xlink:href="data:image/png;base64,{data}"/></svg>'
            outname = f'{uuid.uuid4().hex}_{stem}.svg'
            with open(os.path.join(upload_folder, outname), 'w', encoding='utf-8') as fp:
                fp.write(svg)
            return jsonify(success=True, filename=outname)

        if tool_type == 'image_to_base64':
            mime = Image.MIME.get(Image.open(BytesIO(raw)).format, 'application/octet-stream')
            encoded = base64.b64encode(raw).decode('ascii')
            outname = f'{uuid.uuid4().hex}_{stem}_base64.txt'
            with open(os.path.join(upload_folder, outname), 'w', encoding='utf-8') as fp:
                fp.write(f'data:{mime};base64,{encoded}')
            return jsonify(success=True, filename=outname)

        image = Image.open(BytesIO(raw))
        if tool_type == 'gif_maker':
            frames = []
            for item in named:
                try:
                    item.stream.seek(0)
                    frame = Image.open(item).convert('RGBA')
                    frame.thumbnail((1200, 1200))
                    frames.append(frame.copy())
                except Exception:
                    return jsonify(error='One of the selected files could not be read as an image.'), 400
            outname = f'{uuid.uuid4().hex}_animation.gif'
            frames[0].save(os.path.join(upload_folder, outname), save_all=True, append_images=frames[1:], duration=500, loop=0, optimize=True)
            return jsonify(success=True, filename=outname)
        if tool_type == 'gif_to_jpg':
            image.seek(0); out = image.convert('RGB'); out_format = 'JPEG'
        elif tool_type in ('jpg_to_png', 'webp_to_png', 'avif_to_png', 'bmp_to_png', 'ico_to_png'):
            out = image.convert('RGBA' if 'A' in image.getbands() else 'RGB'); out_format = 'PNG'
        elif tool_type in ('png_to_jpg', 'webp_to_jpg', 'heic_to_jpg', 'avif_to_jpg'):
            out = image.convert('RGB'); out_format = 'JPEG'
        elif tool_type in ('jpg_to_avif', 'webp_to_avif'):
            out = image.convert('RGB'); out_format = 'AVIF'
        elif tool_type == 'image_to_ico':
            out = image.convert('RGBA'); out_format = 'ICO'
        elif tool_type == 'favicon_generator':
            out = image.convert('RGBA'); out_format = 'ICO'
        elif tool_type == 'images_to_webp':
            out = image.convert('RGB'); out_format = 'WEBP'
        else:
            return jsonify(error='Unsupported image conversion.'), 400

        outname = f'{uuid.uuid4().hex}_{stem}.{out_format.lower()}'
        save_options = {}
        if out_format == 'ICO': save_options['sizes'] = [(16,16),(32,32),(48,48),(64,64),(128,128),(256,256)]
        if out_format == 'WEBP': save_options.update(quality=90, method=6)
        if out_format == 'JPEG': save_options['quality'] = 95
        try:
            out.save(os.path.join(upload_folder, outname), format=out_format, **save_options)
        except (OSError, KeyError) as exc:
            if out_format == 'AVIF':
                return jsonify(error='AVIF encoding is not available in this server Pillow build.'), 503
            raise exc
        return jsonify(success=True, filename=outname)
    except Exception as exc:
        return jsonify(error='Conversion failed: ' + str(exc)[:180]), 400


# --- Additional Create PDF pages ---
@app.route('/html-to-pdf')
def html_to_pdf_page(): return render_template('tool_template.html', tool_name='HTML to PDF', tool_type='html_to_pdf')

@app.route('/text-to-pdf')
def text_to_pdf_page(): return render_template('tool_template.html', tool_name='Text to PDF', tool_type='text_to_pdf')

@app.route('/images-to-pdf')
def images_to_pdf_page(): return render_template('tool_template.html', tool_name='Images to PDF', tool_type='images_to_pdf')

@app.route('/epub-to-pdf')
def epub_to_pdf_page(): return render_template('tool_template.html', tool_name='EPUB to PDF', tool_type='epub_to_pdf')

@app.route('/tiff-to-pdf')
def tiff_to_pdf_page(): return render_template('tool_template.html', tool_name='TIFF to PDF', tool_type='tiff_to_pdf')

@app.route('/svg-to-pdf')
def svg_to_pdf_page(): return render_template('tool_template.html', tool_name='SVG to PDF', tool_type='svg_to_pdf')

@app.route('/markdown-to-pdf')
def markdown_to_pdf_page(): return render_template('tool_template.html', tool_name='Markdown to PDF', tool_type='markdown_to_pdf')

@app.route('/csv-to-pdf')
def csv_to_pdf_page(): return render_template('tool_template.html', tool_name='CSV to PDF', tool_type='csv_to_pdf')

@app.route('/webpage-to-pdf')
def webpage_to_pdf_page(): return render_template('tool_template.html', tool_name='Webpage to PDF', tool_type='webpage_to_pdf')



# --- Convert PDF: additional tool pages ---
@app.route('/pdf-to-text')
def page_pdf_to_text():
    return render_template('tool_template.html', tool_name='PDF to Text', tool_type='pdf_to_text')

@app.route('/pdf-to-html')
def page_pdf_to_html():
    return render_template('tool_template.html', tool_name='PDF to HTML', tool_type='pdf_to_html')

@app.route('/pdf-to-json')
def page_pdf_to_json():
    return render_template('tool_template.html', tool_name='PDF to JSON', tool_type='pdf_to_json')

@app.route('/pdf-to-svg')
def page_pdf_to_svg():
    return render_template('tool_template.html', tool_name='PDF to SVG', tool_type='pdf_to_svg')

@app.route('/pdf-to-tiff')
def page_pdf_to_tiff():
    return render_template('tool_template.html', tool_name='PDF to TIFF', tool_type='pdf_to_tiff')

@app.route('/pdf-to-epub')
def page_pdf_to_epub():
    return render_template('tool_template.html', tool_name='PDF to EPUB', tool_type='pdf_to_epub')

@app.route('/pdf-to-images-zip')
def page_pdf_to_images_zip():
    return render_template('tool_template.html', tool_name='PDF to Images ZIP', tool_type='pdf_to_images_zip')

@app.route('/pdf-to-markdown')
def page_pdf_to_markdown():
    return render_template('tool_template.html', tool_name='PDF to Markdown', tool_type='pdf_to_markdown')

@app.route('/pdf-to-csv')
def page_pdf_to_csv():
    return render_template('tool_template.html', tool_name='PDF to CSV', tool_type='pdf_to_csv')

@app.route('/pdf-to-pdfa')
def page_pdf_to_pdfa():
    return render_template('tool_template.html', tool_name='PDF to PDF/A', tool_type='pdf_to_pdfa')

@app.route('/pdfa-to-pdf')
def page_pdfa_to_pdf():
    return render_template('tool_template.html', tool_name='PDF/A to PDF', tool_type='pdfa_to_pdf')


@app.route('/document-convert/<tool_slug>')
def document_conversion_page(tool_slug):
    tool_type = tool_slug.replace('-', '_')
    if tool_type not in DOCUMENT_CONVERSION_TOOLS:
        return 'Tool not found', 404
    tool_name = DOCUMENT_CONVERSION_TOOLS[tool_type][0]
    return render_template('tool_template.html', tool_name=tool_name, tool_type=tool_type)

def safe_public_url(raw_url):
    """Disallow local/private network addresses for user-submitted URLs."""
    url = urlparse(raw_url)
    if url.scheme not in ('https', 'http') or not url.hostname or url.username or url.password:
        raise ValueError('Enter a valid public http/https URL.')
    if url.port not in (None, 80, 443):
        raise ValueError('Custom ports are not supported.')
    addresses = socket.getaddrinfo(url.hostname, None)
    if not addresses or any(not ipaddress.ip_address(info[4][0]).is_global for info in addresses):
        raise ValueError('Local or private network URLs are not allowed.')
    return raw_url



# PDF conversion utilities (additive: existing converters remain unchanged)
import json as _json
import re as _re
import shutil as _shutil
import subprocess as _subprocess
from xml.sax.saxutils import escape as _xml_escape
from flask import current_app

CONVERT_PDF_EXTRA = {"pdf_to_text", "pdf_to_html", "pdf_to_json", "pdf_to_svg", "pdf_to_tiff", "pdf_to_epub", "pdf_to_images_zip", "pdf_to_markdown", "pdf_to_csv", "pdf_to_pdfa", "pdfa_to_pdf"}

def process_convert_pdf_extra(kind, files):
    if not files or not files[0].filename or not files[0].filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Please upload a valid PDF file.'}), 400
    uploaded = files[0]
    raw = uploaded.read()
    if not raw.startswith(b'%PDF-'):
        return jsonify({'error': 'This is not a PDF file.'}), 400
    if len(raw) > 150 * 1024 * 1024:
        return jsonify({'error': 'Maximum PDF size for this converter is 150 MB.'}), 413
    try:
        doc = fitz.open(stream=raw, filetype='pdf')
        if doc.needs_pass:
            return jsonify({'error': 'Password-protected PDF: unlock it first.'}), 400
        if doc.page_count == 0 or doc.page_count > 300:
            return jsonify({'error': 'PDF must have 1 to 300 pages.'}), 400
        name = (os.path.splitext(secure_filename(uploaded.filename))[0] or 'converted')[:65]
        ext = {'pdf_to_text':'txt','pdf_to_html':'html','pdf_to_json':'json','pdf_to_svg':'zip',
               'pdf_to_tiff':'tiff','pdf_to_epub':'epub','pdf_to_images_zip':'zip',
               'pdf_to_markdown':'md','pdf_to_csv':'csv','pdf_to_pdfa':'pdf','pdfa_to_pdf':'pdf'}[kind]
        filename = f'{name}_{kind}_{uuid4().hex[:10]}.{ext}'
        out = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        def page_texts():
            return [p.get_text(sort=True) for p in doc]
        if kind == 'pdf_to_text':
            with open(out,'w',encoding='utf-8') as f:
                f.write('\n\n'.join(page_texts()))
        elif kind == 'pdf_to_html':
            pages = []
            for i,p in enumerate(doc):
                pages.append(f'<section><h2>Page {i+1}</h2>'+p.get_text('html')+'</section>')
            with open(out,'w',encoding='utf-8') as f:
                f.write('<!doctype html><html><head><meta charset="utf-8"><title>PDF conversion</title></head><body>'+''.join(pages)+'</body></html>')
        elif kind == 'pdf_to_json':
            data={'pages':[{'page':i+1,'width':round(p.rect.width,2),'height':round(p.rect.height,2),'text':p.get_text(sort=True),'blocks':p.get_text('blocks')[:2000]} for i,p in enumerate(doc)]}
            with open(out,'w',encoding='utf-8') as f:
                _json.dump(data,f,ensure_ascii=False,indent=2)
        elif kind == 'pdf_to_svg':
            with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
                for i,p in enumerate(doc):
                    z.writestr(f'page_{i+1}.svg',p.get_svg_image(text_as_path=False))
        elif kind == 'pdf_to_tiff':
            # Multipage TIFF at 110 DPI; limit raster size to protect server memory
            frames=[]
            try:
                for p in doc:
                    pix=p.get_pixmap(dpi=110,alpha=False)
                    if pix.width*pix.height>20_000_000: raise ValueError('Page too large to rasterize.')
                    frames.append(Image.frombytes('RGB',(pix.width,pix.height),pix.samples))
                frames[0].save(out,format='TIFF',save_all=True,append_images=frames[1:],compression='tiff_deflate')
            finally:
                for frame in frames: frame.close()
        elif kind == 'pdf_to_images_zip':
            with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
                for i,p in enumerate(doc):
                    pix=p.get_pixmap(dpi=140,alpha=False)
                    z.writestr(f'page_{i+1}.png',pix.tobytes('png'))
        elif kind == 'pdf_to_markdown':
            with open(out,'w',encoding='utf-8') as f:
                for i,p in enumerate(doc):
                    if i:f.write('\n\n---\n\n')
                    f.write(f'## Page {i+1}\n\n'+p.get_text(sort=True).strip()+'\n')
        elif kind == 'pdf_to_csv':
            with open(out,'w',encoding='utf-8-sig',newline='') as f:
                writer=csv.writer(f);writer.writerow(['page','row','column','value'])
                for i,p in enumerate(doc):
                    tables=p.find_tables()
                    found=False
                    for table in tables.tables:
                        found=True
                        for ri,row in enumerate(table.extract()):
                            for ci,val in enumerate(row):
                                writer.writerow([i+1,ri+1,ci+1,val or ''])
                    if not found:
                        for ri,line in enumerate(p.get_text(sort=True).splitlines()):
                            if line.strip():writer.writerow([i+1,ri+1,1,line.strip()])
        elif kind == 'pdf_to_epub':
            book=epub.EpubBook();book.set_identifier(uuid4().hex)
            book.set_title(name);book.set_language('en')
            chapters=[]
            for i,p in enumerate(doc):
                ch=epub.EpubHtml(title=f'Page {i+1}',file_name=f'page_{i+1}.xhtml',lang='en')
                ch.content=f'<html><body><h1>Page {i+1}</h1><pre style="white-space:pre-wrap">{_xml_escape(p.get_text(sort=True))}</pre></body></html>'
                book.add_item(ch);chapters.append(ch)
            book.toc=tuple(chapters);book.spine=['nav']+chapters
            book.add_item(epub.EpubNcx());book.add_item(epub.EpubNav())
            epub.write_epub(out,book)
        elif kind == 'pdfa_to_pdf':
            # PDF/A is already a PDF: resave as ordinary PDF (not guaranteed to remove all PDF/A metadata).
            doc.save(out,garbage=3,deflate=True)
        elif kind == 'pdf_to_pdfa':
            # Ghostscript is needed for standards-based PDF/A-2b conversion.
            gs=_shutil.which('gswin64c') or _shutil.which('gs') or _shutil.which('gswin32c')
            if not gs:
                return jsonify({'error':'PDF/A conversion requires Ghostscript installed and on PATH. Other PDF converters work without it.'}), 503
            with TemporaryDirectory() as td:
                source=os.path.join(td,'source.pdf')
                with open(source,'wb') as f:f.write(raw)
                command=[gs,'-dPDFA=2','-dBATCH','-dNOPAUSE','-dSAFER','-sDEVICE=pdfwrite',
                         '-dPDFACompatibilityPolicy=2','-sColorConversionStrategy=RGB',
                         '-sOutputFile='+out,source]
                result=_subprocess.run(command,stdout=_subprocess.PIPE,stderr=_subprocess.PIPE,timeout=120)
                if result.returncode!=0 or not os.path.isfile(out):
                    return jsonify({'error':'Ghostscript PDF/A conversion failed. See server logs.'}), 422
        doc.close()
        return jsonify({'success':True,'filename':filename})
    except Exception:
        app.logger.exception('Convert PDF failed: %s',kind)
        return jsonify({'error':'Conversion failed. Check the PDF and server logs.'}), 422


# Organize PDF additional tool pages
@app.route('/remove-pdf-pages')
def page_remove_pdf_pages(): return render_template('tool_template.html', tool_name='Remove Pdf Pages', tool_type='remove_pdf_pages')

@app.route('/extract-pdf-pages')
def page_extract_pdf_pages(): return render_template('tool_template.html', tool_name='Extract Pdf Pages', tool_type='extract_pdf_pages')

@app.route('/crop-pdf')
def page_crop_pdf(): return render_template('tool_template.html', tool_name='Crop Pdf', tool_type='crop_pdf')

@app.route('/resize-pdf-pages')
def page_resize_pdf_pages(): return render_template('tool_template.html', tool_name='Resize Pdf Pages', tool_type='resize_pdf_pages')

@app.route('/add-pdf-page-numbers')
def page_add_pdf_page_numbers(): return render_template('tool_template.html', tool_name='Add Pdf Page Numbers', tool_type='add_pdf_page_numbers')

@app.route('/reverse-pdf-pages')
def page_reverse_pdf_pages(): return render_template('tool_template.html', tool_name='Reverse Pdf Pages', tool_type='reverse_pdf_pages')

@app.route('/duplicate-pdf-pages')
def page_duplicate_pdf_pages(): return render_template('tool_template.html', tool_name='Duplicate Pdf Pages', tool_type='duplicate_pdf_pages')

@app.route('/reorder-pdf-pages')
def page_reorder_pdf_pages(): return render_template('tool_template.html', tool_name='Reorder Pdf Pages', tool_type='reorder_pdf_pages')

@app.route('/insert-pdf-pages')
def page_insert_pdf_pages(): return render_template('tool_template.html', tool_name='Insert Pdf Pages', tool_type='insert_pdf_pages')

@app.route('/replace-pdf-pages')
def page_replace_pdf_pages(): return render_template('tool_template.html', tool_name='Replace Pdf Pages', tool_type='replace_pdf_pages')

@app.route('/pdf-booklet-maker')
def page_pdf_booklet_maker(): return render_template('tool_template.html', tool_name='Pdf Booklet Maker', tool_type='pdf_booklet_maker')

@app.route('/remove-blank-pdf-pages')
def page_remove_blank_pdf_pages(): return render_template('tool_template.html', tool_name='Remove Blank Pdf Pages', tool_type='remove_blank_pdf_pages')



# Additional Organize PDF tools. The original routes and processing are preserved.
ORGANIZE_EXTRA_TOOLS = {"remove_pdf_pages", "extract_pdf_pages", "crop_pdf", "resize_pdf_pages",
 "add_pdf_page_numbers", "reverse_pdf_pages", "duplicate_pdf_pages", "reorder_pdf_pages",
 "insert_pdf_pages", "replace_pdf_pages", "pdf_booklet_maker", "remove_blank_pdf_pages"}

def _pages(spec, count, required=False):
    spec = (spec or '').strip()
    if not spec:
        if required: raise ValueError('Enter page numbers, e.g. 1,3-5.')
        return list(range(count))
    result=[]
    for chunk in spec.split(','):
        chunk=chunk.strip()
        if not chunk: continue
        parts=chunk.split('-')
        if len(parts)==1: a=b=int(parts[0])
        elif len(parts)==2: a,b=map(int,parts)
        else: raise ValueError('Invalid page range.')
        if a<1 or b<1 or a>count or b>count: raise ValueError(f'Page numbers must be 1 to {count}.')
        result.extend(range(a-1,b if a<=b else b-2,1 if a<=b else -1))
    if not result: raise ValueError('Select at least one page.')
    return result

def process_organize_extra(tool_type, files):
    try:
        if not files or not files[0].filename or not files[0].filename.lower().endswith('.pdf'):
            return jsonify({'error':'Upload a valid PDF file.'}),400
        from pypdf import Transformation
        import math
        from reportlab.pdfgen import canvas
        first=PdfReader(files[0].stream)
        n=len(first.pages)
        if not n: raise ValueError('PDF has no pages.')
        spec=request.form.get('page_range','')
        writer=PdfWriter()
        if tool_type in ('remove_pdf_pages','extract_pdf_pages','duplicate_pdf_pages','reorder_pdf_pages'):
            chosen=_pages(spec,n,True)
            if tool_type=='remove_pdf_pages': order=[i for i in range(n) if i not in set(chosen)]
            elif tool_type=='duplicate_pdf_pages':
                order=[]
                for i in range(n):
                    order.append(i)
                    if i in chosen: order.append(i)
            else: order=chosen
            if not order: raise ValueError('No pages would remain.')
            for i in order:writer.add_page(first.pages[i])
        elif tool_type=='reverse_pdf_pages':
            for i in reversed(range(n)):writer.add_page(first.pages[i])
        elif tool_type in ('insert_pdf_pages','replace_pdf_pages'):
            if len(files)<2 or not files[1].filename.lower().endswith('.pdf'):
                raise ValueError('Upload two PDF files: original first, pages to insert/replace second.')
            second=PdfReader(files[1].stream)
            if not second.pages: raise ValueError('Second PDF has no pages.')
            pos=int(request.form.get('position','1'))
            if tool_type=='insert_pdf_pages':
                if not 1<=pos<=n+1:raise ValueError(f'Insert position must be 1 to {n+1}.')
                for i in range(n+1):
                    if i==pos-1:
                        for p in second.pages:writer.add_page(p)
                    if i<n:writer.add_page(first.pages[i])
            else:
                if not 1<=pos<=n:raise ValueError(f'Replace position must be 1 to {n}.')
                for i in range(n):
                    if i==pos-1:
                        for p in second.pages:writer.add_page(p)
                    if not(pos-1<=i<pos-1+len(second.pages)):writer.add_page(first.pages[i])
        elif tool_type=='remove_blank_pdf_pages':
            # Text-only blank detection is insufficient; also inspect drawings and images.
            doc=fitz.open(stream=files[0].stream.read(),filetype='pdf')
            for i,p in enumerate(doc):
                if p.get_text('text').strip() or p.get_images(full=True) or p.get_drawings():writer.add_page(first.pages[i])
            doc.close()
            if not writer.pages:raise ValueError('Every page appears blank.')
        elif tool_type=='pdf_booklet_maker':
            # Booklet imposition on landscape sheets, 2 source pages per side.
            from pypdf import PageObject
            total=((n+3)//4)*4
            sheetw,sheeth=1190,842
            for k in range(total//4):
                for pair in ((total-2*k-1,2*k),(2*k+1,total-2*k-2)):
                    sheet=PageObject.create_blank_page(width=sheetw,height=sheeth)
                    for side,idx in enumerate(pair):
                        if idx>=n:continue
                        p=first.pages[idx]
                        w=float(p.mediabox.width);h=float(p.mediabox.height)
                        scale=min(sheetw/2/w,sheeth/h)*.96
                        tx=side*(sheetw/2)+(sheetw/2-w*scale)/2
                        ty=(sheeth-h*scale)/2
                        sheet.merge_transformed_page(p,Transformation().scale(scale).translate(tx,ty))
                    writer.add_page(sheet)
        elif tool_type in ('crop_pdf','resize_pdf_pages','add_pdf_page_numbers'):
            if tool_type=='add_pdf_page_numbers':
                from reportlab.pdfgen import canvas
                for i,p in enumerate(first.pages):
                    w,h=float(p.mediabox.width),float(p.mediabox.height)
                    buf=BytesIO();c=canvas.Canvas(buf,pagesize=(w,h));c.setFont('Helvetica',11)
                    c.drawCentredString(w/2,22,str(i+1));c.save();buf.seek(0)
                    p.merge_page(PdfReader(buf).pages[0]);writer.add_page(p)
            elif tool_type=='crop_pdf':
                margin=float(request.form.get('margin','20'))
                if margin<0:raise ValueError('Margin cannot be negative.')
                for p in first.pages:
                    l,b,r,t=map(float,(p.mediabox.left,p.mediabox.bottom,p.mediabox.right,p.mediabox.top))
                    if r-l<=2*margin or t-b<=2*margin:raise ValueError('Crop margin too large.')
                    p.cropbox.lower_left=(l+margin,b+margin)
                    p.cropbox.upper_right=(r-margin,t-margin)
                    writer.add_page(p)
            else:
                width=float(request.form.get('width','595'));height=float(request.form.get('height','842'))
                if not 72<=width<=3000 or not 72<=height<=3000:raise ValueError('Width/height must be 72–3000 points.')
                from pypdf import PageObject
                for p in first.pages:
                    new=PageObject.create_blank_page(width=width,height=height)
                    pw=float(p.mediabox.width);ph=float(p.mediabox.height)
                    scale=min(width/pw,height/ph)
                    new.merge_transformed_page(p,Transformation().scale(scale).translate((width-pw*scale)/2,(height-ph*scale)/2))
                    writer.add_page(new)
        else:raise ValueError('Unknown tool.')
        filename=f'filevellum_{tool_type}_{uuid4().hex[:10]}.pdf'
        with open(os.path.join(app.config['UPLOAD_FOLDER'],filename),'wb') as f:writer.write(f)
        return jsonify({'success':True,'filename':filename})
    except (ValueError,IndexError,TypeError) as exc:
        return jsonify({'error':str(exc)}),400
    except Exception:
        app.logger.exception('Organize PDF processing failed')
        return jsonify({'error':'PDF processing failed. Check that the file is valid and not encrypted.'}),500

# --- Real Processing Logic with Strict Validations ---

# --- Optimize & OCR PDF tool pages ---
@app.route('/grayscale-pdf')
def grayscale_pdf_page(): return render_template('tool_template.html', tool_name='Grayscale PDF', tool_type='grayscale_pdf')
@app.route('/ocr-pdf')
def ocr_pdf_page(): return render_template('tool_template.html', tool_name='OCR PDF', tool_type='ocr_pdf')
@app.route('/make-pdf-searchable')
def make_pdf_searchable_page(): return render_template('tool_template.html', tool_name='Make PDF Searchable', tool_type='make_pdf_searchable')
@app.route('/optimize-pdf-images')
def optimize_pdf_images_page(): return render_template('tool_template.html', tool_name='Optimize PDF Images', tool_type='optimize_pdf_images')
@app.route('/pdf-file-size-analyzer')
def pdf_file_size_analyzer_page(): return render_template('tool_template.html', tool_name='PDF File Size Analyzer', tool_type='pdf_file_size_analyzer')
@app.route('/pdf-page-size-checker')
def pdf_page_size_checker_page(): return render_template('tool_template.html', tool_name='PDF Page Size Checker', tool_type='pdf_page_size_checker')
@app.route('/pdf-font-inspector')
def pdf_font_inspector_page(): return render_template('tool_template.html', tool_name='PDF Font Inspector', tool_type='pdf_font_inspector')
@app.route('/pdf-color-converter')
def pdf_color_converter_page(): return render_template('tool_template.html', tool_name='PDF Color Converter', tool_type='pdf_color_converter')
@app.route('/pdf-image-extractor')
def pdf_image_extractor_page(): return render_template('tool_template.html', tool_name='PDF Image Extractor', tool_type='pdf_image_extractor')

@app.route('/edit-secure/<tool_slug>')
def edit_secure_page(tool_slug):
    tool_type = tool_slug.replace('-', '_')
    if tool_type not in EDIT_SECURE_TOOLS:
        return 'Tool not found', 404
    return render_template('tool_template.html', tool_name=EDIT_SECURE_TOOLS[tool_type], tool_type=tool_type)


IMAGE_PROCESS_TYPES = {v.replace('-', '_') for v in [
    'image-compressor','image-resizer','image-cropper','image-rotator','image-flipper',
    'image-watermark','passport-photo-maker','blur-image','sharpen-image','image-filters',
    'image-to-text','image-metadata-viewer','remove-image-metadata','image-background-changer',
    'image-border-maker','image-collage-maker','image-upscaler','image-color-palette'
]} | {'remove_bg'}

def process_image_tool(tool_type, files, form, upload_folder):
    """Validated image processing for Image Editing tools."""
    allowed = {'.jpg','.jpeg','.png','.webp','.bmp','.tif','.tiff'}
    if not files or not files[0] or not files[0].filename:
        return jsonify({'error':'Please upload an image.'}), 400
    if tool_type == 'image_collage_maker':
        chosen = [f for f in files if f and f.filename]
        if len(chosen) < 2:
            return jsonify({'error':'Please select at least two images for a collage.'}),400
    else:
        chosen = [files[0]]
    try:
        images=[]
        for f in chosen:
            ext=os.path.splitext(secure_filename(f.filename))[1].lower()
            if ext not in allowed: return jsonify({'error':'Supported formats: JPG, PNG, WEBP, BMP and TIFF.'}),400
            raw=f.read()
            if not raw or len(raw)>30*1024*1024: return jsonify({'error':'Each image must be non-empty and no larger than 30 MB.'}),400
            im=Image.open(BytesIO(raw))
            im.verify()
            im=Image.open(BytesIO(raw))
            if im.width*im.height>25_000_000: return jsonify({'error':'Image exceeds 25 megapixels.'}),400
            images.append((im.copy(),f.filename))
        img, original=images[0]
        base=os.path.splitext(secure_filename(original))[0] or 'image'
        out_ext='.png'
        if tool_type=='image_metadata_viewer':
            info={'filename':original,'format':img.format,'width':img.width,'height':img.height,'mode':img.mode,
                  'file_size_bytes':len(raw),'metadata':{str(k):str(v)[:500] for k,v in img.info.items() if k not in ('icc_profile','exif')}}
            name=f'{uuid4().hex}_image_metadata.json'
            with open(os.path.join(upload_folder,name),'w',encoding='utf-8') as fp: json.dump(info,fp,indent=2,ensure_ascii=False)
            return jsonify(success=True,filename=name)
        if tool_type=='image_to_text':
            try:
                import pytesseract
                text=pytesseract.image_to_string(img,lang=form.get('language','eng'))
            except Exception as exc:
                return jsonify(error='OCR requires pytesseract and the Tesseract OCR engine. '+str(exc)[:180]),503
            name=f'{uuid4().hex}_{base}_text.txt'
            with open(os.path.join(upload_folder,name),'w',encoding='utf-8') as fp: fp.write(text)
            return jsonify(success=True,filename=name)
        if tool_type=='image_color_palette':
            colors=img.convert('RGB').resize((max(1,min(300,img.width)),max(1,min(300,img.height)))).quantize(colors=8).convert('RGB')
            palette=colors.getcolors(maxcolors=1000000) or []
            palette=sorted(palette,reverse=True)[:8]
            report={'filename':original,'colors':[{'hex':'#%02X%02X%02X'%tuple(rgb),'pixels':count} for count,rgb in palette]}
            name=f'{uuid4().hex}_{base}_palette.json'
            with open(os.path.join(upload_folder,name),'w',encoding='utf-8') as fp: json.dump(report,fp,indent=2)
            return jsonify(success=True,filename=name)
        if tool_type=='remove_bg':
            try:
                from rembg import remove
                output=remove(raw,session=get_bg_session())
                out=Image.open(BytesIO(output)).convert('RGBA')
            except Exception as exc:
                return jsonify(error='Background removal needs rembg[cpu] and its model. '+str(exc)[:180]),503
        elif tool_type=='image_compressor':
            quality=max(20,min(95,int(form.get('quality',70))))
            out=img.convert('RGB'); out_ext='.jpg'
            name=f'{uuid4().hex}_{base}_compressed.jpg'
            out.save(os.path.join(upload_folder,name),'JPEG',quality=quality,optimize=True)
            return jsonify(success=True,filename=name)
        elif tool_type=='image_resizer':
            w=int(form.get('width') or 0); h=int(form.get('height') or 0)
            if not 1<=w<=12000 or not 1<=h<=12000:return jsonify(error='Enter width and height between 1 and 12000 pixels.'),400
            out=img.resize((w,h),Image.Resampling.LANCZOS)
        elif tool_type=='image_cropper':
            x=int(form.get('x') or 0); y=int(form.get('y') or 0)
            w=int(form.get('width') or 0); h=int(form.get('height') or 0)
            if w<1 or h<1 or x<0 or y<0 or x+w>img.width or y+h>img.height:return jsonify(error='Crop area is outside the image or invalid.'),400
            out=img.crop((x,y,x+w,y+h))
        elif tool_type=='image_rotator':
            angle=int(form.get('angle') or 90)%360
            out=img.rotate(-angle,expand=True)
        elif tool_type=='image_flipper':
            out=ImageOps.mirror(img) if form.get('direction')=='horizontal' else ImageOps.flip(img)
        elif tool_type=='image_watermark':
            out=img.convert('RGBA'); layer=Image.new('RGBA',out.size,(0,0,0,0)); draw=ImageDraw.Draw(layer)
            text=(form.get('text') or 'FileVellum').strip()[:80]
            draw.text((20,max(10,out.height-55)),text,fill=(255,255,255,150),stroke_width=1,stroke_fill=(0,0,0,120))
            out=Image.alpha_composite(out,layer)
        elif tool_type=='passport_photo_maker':
            w=int(form.get('width') or 413); h=int(form.get('height') or 531)
            if not 100<=w<=2000 or not 100<=h<=2000:return jsonify(error='Photo dimensions must be 100–2000 px.'),400
            out=ImageOps.fit(img.convert('RGB'),(w,h),method=Image.Resampling.LANCZOS,centering=(.5,.35))
        elif tool_type=='blur_image':
            radius=max(1,min(30,float(form.get('radius') or 3)))
            out=img.filter(ImageFilter.GaussianBlur(radius))
        elif tool_type=='sharpen_image':
            out=img.filter(ImageFilter.UnsharpMask(radius=2,percent=160,threshold=3))
        elif tool_type=='image_filters':
            mode=form.get('filter','grayscale')
            if mode=='grayscale': out=ImageOps.grayscale(img)
            elif mode=='sepia':
                gray=ImageOps.grayscale(img).convert('RGB')
                out=Image.merge('RGB',tuple(gray.getchannel('RGB'[i]).point(lambda p, c=i: min(255,int(p*[1.0,.95,.82][c]))) for i in range(3)))
            elif mode=='invert': out=ImageOps.invert(img.convert('RGB'))
            elif mode=='bright': out=ImageEnhance.Brightness(img).enhance(1.25)
            elif mode=='contrast': out=ImageEnhance.Contrast(img).enhance(1.35)
            else: out=img.copy()
        elif tool_type=='remove_image_metadata':
            out=Image.new(img.mode,img.size); out.putdata(list(img.getdata()))
        elif tool_type=='image_background_changer':
            color=form.get('color','#ffffff').strip()
            if not re.fullmatch(r'#[0-9a-fA-F]{6}',color):return jsonify(error='Choose a valid HEX color.'),400
            rgba=img.convert('RGBA'); bg=Image.new('RGBA',rgba.size,color); bg.alpha_composite(rgba); out=bg.convert('RGB')
        elif tool_type=='image_border_maker':
            size=max(1,min(300,int(form.get('size') or 20)))
            color=form.get('color','#ffeb18')
            if not re.fullmatch(r'#[0-9a-fA-F]{6}',color):return jsonify(error='Choose a valid HEX color.'),400
            out=ImageOps.expand(img,border=size,fill=color)
        elif tool_type=='image_collage_maker':
            thumbs=[]
            for im,_ in images:
                im.thumbnail((500,500)); thumbs.append(im.convert('RGB'))
            cols=2; rows=(len(thumbs)+1)//2
            out=Image.new('RGB',(cols*520,rows*520),'white')
            for i,im in enumerate(thumbs): out.paste(im,((i%2)*520+10,(i//2)*520+10))
        elif tool_type=='image_upscaler':
            factor=max(2,min(4,int(form.get('scale') or 2)))
            if img.width*factor>12000 or img.height*factor>12000:return jsonify(error='Upscaled dimensions would exceed 12000 px.'),400
            out=img.resize((img.width*factor,img.height*factor),Image.Resampling.LANCZOS)
        else:
            return jsonify(error='Unsupported image tool.'),400
        name=f'{uuid4().hex}_{base}_{tool_type}.png'
        out.save(os.path.join(upload_folder,name),'PNG',optimize=True)
        return jsonify(success=True,filename=name)
    except (OSError,ValueError,TypeError) as exc:
        return jsonify(error='Unable to process image: '+str(exc)[:180]),400


# --- Document Conversion tools ---
DOCUMENT_CONVERSION_TOOLS = {
    'docx_to_txt': ('DOCX to TXT', '.docx', '.txt'),
    'txt_to_docx': ('TXT to DOCX', '.txt', '.docx'),
    'docx_to_html': ('DOCX to HTML', '.docx', '.html'),
    'html_to_docx': ('HTML to DOCX', ('.html', '.htm'), '.docx'),
    'csv_to_excel': ('CSV to Excel', '.csv', '.xlsx'),
    'excel_to_csv': ('Excel to CSV', ('.xlsx', '.xls'), '.csv'),
    'json_to_csv': ('JSON to CSV', '.json', '.csv'),
    'csv_to_json': ('CSV to JSON', '.csv', '.json'),
    'xml_to_json': ('XML to JSON', '.xml', '.json'),
    'json_to_xml': ('JSON to XML', '.json', '.xml'),
    'rtf_to_docx': ('RTF to DOCX', '.rtf', '.docx'),
    'docx_to_rtf': ('DOCX to RTF', '.docx', '.rtf'),
    'odt_to_docx': ('ODT to DOCX', '.odt', '.docx'),
    'docx_to_odt': ('DOCX to ODT', '.docx', '.odt'),
}


def _xml_element_to_value(element):
    """Convert an XML element tree into JSON-compatible values."""
    children = list(element)
    value = (element.text or '').strip()
    if not children:
        return value
    result = {}
    for child in children:
        child_value = _xml_element_to_value(child)
        if child.tag in result:
            if not isinstance(result[child.tag], list):
                result[child.tag] = [result[child.tag]]
            result[child.tag].append(child_value)
        else:
            result[child.tag] = child_value
    if value:
        result['_text'] = value
    if element.attrib:
        result['_attributes'] = element.attrib
    return result


def _json_to_xml_element(parent, key, value):
    child = ET.SubElement(parent, str(key))
    if isinstance(value, dict):
        for sub_key, sub_value in value.items():
            _json_to_xml_element(child, sub_key, sub_value)
    elif isinstance(value, list):
        for item in value:
            _json_to_xml_element(child, 'item', item)
    elif value is not None:
        child.text = str(value)
    return child


def _rtf_to_text(raw):
    """Extract readable text from common RTF documents without executing content."""
    text = raw.decode('cp1252', errors='replace')
    text = re.sub(r"\\'[0-9a-fA-F]{2}", lambda m: bytes([int(m.group(0)[2:], 16)]).decode('cp1252'), text)
    text = re.sub(r'\\u(-?\d+)\??', lambda m: chr(int(m.group(1)) % 65536), text)
    text = re.sub(r'\\(par|line|row)\b', '\n', text)
    text = re.sub(r'\\tab\b', '\t', text)
    text = re.sub(r'\\[a-zA-Z]+-?\d* ?', '', text)
    text = re.sub(r'\\[^a-zA-Z]', '', text)
    text = text.replace('{', '').replace('}', '')
    return re.sub(r'\n{3,}', '\n\n', text).strip()


def _docx_paragraphs_and_tables(document):
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    body = document.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield ('paragraph', Paragraph(child, document))
        elif isinstance(child, CT_Tbl):
            yield ('table', Table(child, document))


def process_document_conversion(kind, files):
    """Convert the 14 document formats listed in the Document Conversion catalog."""
    if kind not in DOCUMENT_CONVERSION_TOOLS:
        return jsonify({'error': 'Unsupported document conversion.'}), 400
    if not files or not files[0] or not files[0].filename:
        return jsonify({'error': 'Please select a file first.'}), 400
    file = files[0]
    _, allowed_ext, output_ext = DOCUMENT_CONVERSION_TOOLS[kind]
    original_name = secure_filename(file.filename)
    extension = os.path.splitext(original_name)[1].lower()
    allowed = allowed_ext if isinstance(allowed_ext, tuple) else (allowed_ext,)
    if extension not in allowed:
        return jsonify({'error': f"This tool accepts: {', '.join(allowed)}"}), 400
    raw = file.read()
    if not raw:
        return jsonify({'error': 'The selected file is empty.'}), 400
    if len(raw) > 50 * 1024 * 1024:
        return jsonify({'error': 'Maximum file size is 50 MB.'}), 413

    stem = (os.path.splitext(original_name)[0] or 'converted')[:80]
    out_ext = output_ext
    output_name = f'{stem}_{kind}_{uuid4().hex[:8]}{out_ext}'
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_name)
    try:
        if kind == 'docx_to_txt':
            document = docx.Document(BytesIO(raw))
            chunks = []
            for item_type, item in _docx_paragraphs_and_tables(document):
                if item_type == 'paragraph':
                    chunks.append(item.text)
                else:
                    chunks.extend(' | '.join(cell.text for cell in row.cells) for row in item.rows)
            with open(output_path, 'w', encoding='utf-8-sig') as out:
                out.write('\n'.join(chunks))

        elif kind == 'txt_to_docx':
            text = raw.decode('utf-8-sig')
            document = docx.Document()
            for line in text.splitlines() or ['']:
                document.add_paragraph(line)
            document.save(output_path)

        elif kind == 'docx_to_html':
            document = docx.Document(BytesIO(raw))
            parts = ['<!doctype html><html><head><meta charset="utf-8"><title>Converted document</title></head><body>']
            for item_type, item in _docx_paragraphs_and_tables(document):
                if item_type == 'paragraph':
                    value = html.escape(item.text)
                    style = (item.style.name or '').lower()
                    if style.startswith('heading'):
                        match = re.search(r'(\d+)', style)
                        level = min(6, max(1, int(match.group(1)) if match else 1))
                        parts.append(f'<h{level}>{value}</h{level}>')
                    else:
                        parts.append(f'<p>{value}</p>')
                else:
                    parts.append('<table border="1"><tbody>')
                    for row in item.rows:
                        parts.append('<tr>' + ''.join(f'<td>{html.escape(cell.text)}</td>' for cell in row.cells) + '</tr>')
                    parts.append('</tbody></table>')
            parts.append('</body></html>')
            with open(output_path, 'w', encoding='utf-8') as out:
                out.write('\n'.join(parts))

        elif kind == 'html_to_docx':
            soup = BeautifulSoup(raw.decode('utf-8-sig', errors='replace'), 'html.parser')
            document = docx.Document()
            root = soup.body or soup
            for node in root.find_all(['h1','h2','h3','h4','h5','h6','p','li','pre','blockquote','br'], recursive=True):
                if node.name == 'br':
                    continue
                value = node.get_text(' ', strip=True)
                if not value:
                    continue
                if node.name.startswith('h') and len(node.name) == 2:
                    document.add_heading(value, level=min(9, int(node.name[1])))
                elif node.name == 'li':
                    document.add_paragraph(value, style='List Bullet')
                else:
                    document.add_paragraph(value)
            if not document.paragraphs:
                document.add_paragraph(root.get_text('\n', strip=True))
            document.save(output_path)

        elif kind == 'csv_to_excel':
            frame = pd.read_csv(BytesIO(raw), encoding='utf-8-sig')
            frame.to_excel(output_path, index=False, engine='openpyxl')

        elif kind == 'excel_to_csv':
            frame = pd.read_excel(BytesIO(raw), sheet_name=0, engine='xlrd' if extension == '.xls' else 'openpyxl')
            frame.to_csv(output_path, index=False, encoding='utf-8-sig')

        elif kind == 'json_to_csv':
            data = json.loads(raw.decode('utf-8-sig'))
            if isinstance(data, dict):
                data = data.get('data', [data]) if isinstance(data.get('data', [data]), list) else [data]
            if not isinstance(data, list):
                return jsonify({'error': 'JSON must contain an array of objects or a single object.'}), 400
            if data and not all(isinstance(row, dict) for row in data):
                return jsonify({'error': 'JSON array items must be objects to convert to CSV.'}), 400
            columns = list(dict.fromkeys(key for row in data for key in row.keys()))
            with open(output_path, 'w', newline='', encoding='utf-8-sig') as out:
                writer = csv.DictWriter(out, fieldnames=columns or ['value'])
                writer.writeheader()
                for row in data:
                    writer.writerow({key: json.dumps(row.get(key), ensure_ascii=False) if isinstance(row.get(key), (dict, list)) else row.get(key, '') for key in columns})

        elif kind == 'csv_to_json':
            with open(output_path, 'w', encoding='utf-8') as out:
                rows = list(csv.DictReader(__import__('io').StringIO(raw.decode('utf-8-sig'))))
                json.dump(rows, out, ensure_ascii=False, indent=2)

        elif kind == 'xml_to_json':
            root = ET.fromstring(raw)
            with open(output_path, 'w', encoding='utf-8') as out:
                json.dump({root.tag: _xml_element_to_value(root)}, out, ensure_ascii=False, indent=2)

        elif kind == 'json_to_xml':
            data = json.loads(raw.decode('utf-8-sig'))
            root = ET.Element('root')
            if isinstance(data, dict):
                for key, value in data.items():
                    _json_to_xml_element(root, key, value)
            elif isinstance(data, list):
                for value in data:
                    _json_to_xml_element(root, 'item', value)
            else:
                root.text = str(data)
            ET.ElementTree(root).write(output_path, encoding='utf-8', xml_declaration=True)

        elif kind == 'rtf_to_docx':
            text = _rtf_to_text(raw)
            document = docx.Document()
            for line in text.splitlines() or ['']:
                document.add_paragraph(line)
            document.save(output_path)

        elif kind == 'docx_to_rtf':
            document = docx.Document(BytesIO(raw))
            paragraphs = []
            for item_type, item in _docx_paragraphs_and_tables(document):
                if item_type == 'paragraph':
                    paragraphs.append(item.text)
                else:
                    paragraphs.extend(' \\tab '.join(cell.text for cell in row.cells) for row in item.rows)
            def rtf_escape(value):
                result=[]
                for char in value:
                    code=ord(char)
                    if char in r'\\{}': result.append('\\'+char)
                    elif char == '\n': result.append('\\line ')
                    elif code > 127: result.append(f'\\u{code if code < 32768 else code-65536}?')
                    else: result.append(char)
                return ''.join(result)
            with open(output_path, 'w', encoding='ascii', errors='xmlcharrefreplace') as out:
                out.write(r'{\rtf1\ansi\deff0 ' + r'\par '.join(rtf_escape(p) for p in paragraphs) + '}')

        elif kind == 'odt_to_docx':
            with zipfile.ZipFile(BytesIO(raw), 'r') as archive:
                content = ET.fromstring(archive.read('content.xml'))
            ns = {'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'}
            document = docx.Document()
            for paragraph in content.findall('.//text:p', ns) + content.findall('.//text:h', ns):
                value = ''.join(paragraph.itertext()).strip()
                document.add_paragraph(value)
            document.save(output_path)

        elif kind == 'docx_to_odt':
            document = docx.Document(BytesIO(raw))
            paragraphs = []
            for item_type, item in _docx_paragraphs_and_tables(document):
                if item_type == 'paragraph':
                    paragraphs.append(item.text)
                else:
                    paragraphs.extend(' | '.join(cell.text for cell in row.cells) for row in item.rows)
            office = 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'
            text_ns = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'
            ET.register_namespace('office', office); ET.register_namespace('text', text_ns)
            root = ET.Element(f'{{{office}}}document-content', {f'{{{office}}}version':'1.2'})
            body = ET.SubElement(root, f'{{{office}}}body')
            text_body = ET.SubElement(body, f'{{{office}}}text')
            for value in paragraphs:
                paragraph = ET.SubElement(text_body, f'{{{text_ns}}}p')
                paragraph.text = value
            content_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            with zipfile.ZipFile(output_path, 'w') as archive:
                archive.writestr('mimetype', 'application/vnd.oasis.opendocument.text', compress_type=zipfile.ZIP_STORED)
                archive.writestr('content.xml', content_xml, compress_type=zipfile.ZIP_DEFLATED)
                archive.writestr('META-INF/manifest.xml', '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>', compress_type=zipfile.ZIP_DEFLATED)

        return jsonify({'success': True, 'filename': output_name})
    except (UnicodeDecodeError, json.JSONDecodeError, ET.ParseError, zipfile.BadZipFile, ValueError, KeyError) as exc:
        app.logger.info('Document conversion validation error (%s): %s', kind, exc)
        if os.path.exists(output_path): os.remove(output_path)
        return jsonify({'error': 'Could not convert this file. Please check that it is valid and not corrupted.'}), 400
    except Exception:
        app.logger.exception('Document conversion failed: %s', kind)
        if os.path.exists(output_path): os.remove(output_path)
        return jsonify({'error': 'Conversion failed. Please check the file and try again.'}), 422




# File Utilities: dedicated pages for the existing homepage catalog.
FILE_UTILITY_TOOLS = {
    'file_size_converter': 'File Size Converter',
    'file_hash_generator': 'File Hash Generator',
    'file_type_detector': 'File Type Detector',
    'file_renamer': 'File Renamer',
    'batch_file_renamer': 'Batch File Renamer',
    'file_size_analyzer': 'File Size Analyzer',
    'file_extension_checker': 'File Extension Checker',
    'file_checksum_verifier': 'File Checksum Verifier',
    'file_to_base64': 'File to Base64',
    'base64_to_file': 'Base64 to File',
}

@app.route('/file-utility/<tool_slug>')
def file_utility_page(tool_slug):
    tool_type = tool_slug.replace('-', '_')
    if tool_type not in FILE_UTILITY_TOOLS:
        return 'Tool not found', 404
    return render_template('tool_template.html',
                           tool_name=FILE_UTILITY_TOOLS[tool_type],
                           tool_type=tool_type)


def _file_utility_process(tool_type, files, form, upload_folder):
    """Process the pre-existing File Utilities catalog tools."""
    from werkzeug.utils import secure_filename
    from flask import jsonify
    import base64
    import zipfile
    import io

    def save_report(suffix, content, binary=False):
        filename = f'filevellum_{tool_type}_{uuid4().hex[:12]}{suffix}'
        mode = 'wb' if binary else 'w'
        kwargs = {} if binary else {'encoding': 'utf-8'}
        with open(os.path.join(upload_folder, filename), mode, **kwargs) as fp:
            fp.write(content)
        return jsonify(success=True, filename=filename)

    valid = [f for f in (files or []) if f and f.filename]
    if tool_type == 'file_size_converter':
        try:
            value = float(form.get('size_value', ''))
            unit = form.get('size_unit', 'bytes').lower()
            if not math.isfinite(value) or value < 0:
                raise ValueError
            factors = {'bytes':1, 'kb':1024, 'mb':1024**2, 'gb':1024**3, 'tb':1024**4}
            if unit not in factors: raise ValueError
            size_bytes = value * factors[unit]
            report = (f'Input: {value:g} {unit.upper()}\\n'
                      f'Bytes: {size_bytes:.0f}\\n'
                      f'KB: {size_bytes/1024:.6f}\\n'
                      f'MB: {size_bytes/1024**2:.9f}\\n'
                      f'GB: {size_bytes/1024**3:.12f}\\n'
                      f'TB: {size_bytes/1024**4:.12f}\\n')
            return save_report('.txt', report)
        except (ValueError, OverflowError):
            return jsonify(error='Enter a valid non-negative file size.'), 400

    if tool_type == 'base64_to_file':
        raw_text = (form.get('base64_data') or '').strip()
        if not raw_text:
            return jsonify(error='Paste Base64 data first.'), 400
        if raw_text.startswith('data:') and ',' in raw_text:
            raw_text = raw_text.split(',', 1)[1]
        try:
            payload = base64.b64decode(''.join(raw_text.split()), validate=True)
        except Exception:
            return jsonify(error='Invalid Base64 data.'), 400
        if not payload or len(payload) > 100*1024*1024:
            return jsonify(error='Decoded file must be between 1 byte and 100 MB.'), 400
        requested = secure_filename(form.get('output_filename', '').strip()) or 'decoded_file.bin'
        if '.' not in requested: requested += '.bin'
        filename = f'filevellum_{uuid4().hex[:10]}_{requested}'
        with open(os.path.join(upload_folder, filename), 'wb') as fp: fp.write(payload)
        return jsonify(success=True, filename=filename)

    if not valid:
        return jsonify(error='Please select at least one file.'), 400
    if any(f.content_length and f.content_length > 100*1024*1024 for f in valid):
        return jsonify(error='Each file must be no larger than 100 MB.'), 413

    # Read once; the same bytes are used by all operations.
    items=[]
    total=0
    for f in valid:
        data=f.read()
        total += len(data)
        if not data: return jsonify(error=f'{f.filename} is empty.'), 400
        if total > 250*1024*1024:
            return jsonify(error='Combined upload size must not exceed 250 MB.'), 413
        items.append((f.filename, data))

    if tool_type in ('file_hash_generator','file_checksum_verifier'):
        name, data = items[0]
        algorithm = (form.get('hash_algorithm') or 'sha256').lower()
        if algorithm not in ('md5','sha1','sha256','sha512'):
            return jsonify(error='Choose MD5, SHA-1, SHA-256 or SHA-512.'), 400
        digest = hashlib.new(algorithm, data).hexdigest()
        if tool_type == 'file_hash_generator':
            report = {'filename':name,'size_bytes':len(data),'algorithm':algorithm,'hash':digest}
            return save_report('.json', json.dumps(report, indent=2))
        expected = (form.get('expected_hash') or '').strip().lower()
        if not expected:
            return jsonify(error='Enter the expected checksum to verify.'), 400
        matched = hmac.compare_digest(digest, expected)
        report = (f'File: {name}\\nAlgorithm: {algorithm.upper()}\\n'
                  f'Calculated: {digest}\\nExpected: {expected}\\n'
                  f'Result: {"MATCH" if matched else "NO MATCH"}\\n')
        return save_report('.txt', report)

    if tool_type == 'file_type_detector':
        name, data = items[0]
        signatures = [
            (b'%PDF-', 'PDF document', 'application/pdf'),
            (b'PK\\x03\\x04', 'ZIP-based archive/document', 'application/zip'),
            (b'\\x1f\\x8b', 'GZIP archive', 'application/gzip'),
            (b'\\x89PNG\\r\\n\\x1a\\n', 'PNG image', 'image/png'),
            (b'\\xff\\xd8\\xff', 'JPEG image', 'image/jpeg'),
            (b'GIF87a', 'GIF image', 'image/gif'), (b'GIF89a', 'GIF image', 'image/gif'),
            (b'BM', 'BMP image', 'image/bmp'),
            (b'\\x00\\x00\\x01\\x00', 'ICO image', 'image/x-icon'),
            (b'RIFF', 'RIFF container (WEBP/WAV/AVI)', 'application/octet-stream'),
            (b'OggS', 'Ogg media', 'application/ogg'),
            (b'ID3', 'MP3 audio', 'audio/mpeg'),
        ]
        detected = next((x for sig,*x in signatures if data.startswith(sig)), None)
        if detected is None and len(data)>12 and data[4:8] == b'ftyp':
            detected = ['ISO Base Media (MP4/HEIF family)', 'video/mp4']
        if detected is None:
            detected = ['Unknown / text / custom binary', 'application/octet-stream']
        report={'filename':name,'extension':os.path.splitext(name)[1].lower(),
                'detected_type':detected[0],'mime_type':detected[1],
                'size_bytes':len(data)}
        return save_report('.json', json.dumps(report, indent=2))

    if tool_type in ('file_renamer','batch_file_renamer'):
        prefix = form.get('new_name', '').strip()
        if not prefix:
            return jsonify(error='Enter the new filename.'), 400
        prefix = secure_filename(prefix)
        if not prefix: return jsonify(error='Enter a valid filename.'), 400
        if tool_type == 'file_renamer':
            if len(items) != 1: return jsonify(error='Select exactly one file.'), 400
            original, data = items[0]
            ext = os.path.splitext(original)[1]
            stem = os.path.splitext(prefix)[0] if os.path.splitext(prefix)[1] else prefix
            outname = secure_filename(stem) + ext
            filename=f'filevellum_{uuid4().hex[:8]}_{outname}'
            with open(os.path.join(upload_folder,filename),'wb') as fp: fp.write(data)
            return jsonify(success=True, filename=filename)
        stem = os.path.splitext(prefix)[0] if os.path.splitext(prefix)[1] else prefix
        filename=f'filevellum_batch_rename_{uuid4().hex[:10]}.zip'
        used=set()
        with zipfile.ZipFile(os.path.join(upload_folder,filename),'w',zipfile.ZIP_DEFLATED) as zf:
            for i,(original,data) in enumerate(items,1):
                ext=os.path.splitext(original)[1]
                candidate=f'{stem}_{i}{ext}'
                while candidate.casefold() in used:
                    candidate=f'{stem}_{i}_{len(used)}{ext}'
                used.add(candidate.casefold())
                zf.writestr(candidate,data)
        return jsonify(success=True,filename=filename)

    if tool_type in ('file_size_analyzer','file_extension_checker'):
        rows=[]
        for name,data in items:
            ext=os.path.splitext(name)[1].lower()
            rows.append({'filename':name,'extension':ext or '(none)',
                         'size_bytes':len(data),'size_kb':round(len(data)/1024,3),
                         'size_mb':round(len(data)/1024**2,6),
                         'mime_type_guess':mimetypes.guess_type(name)[0] or 'unknown',
                         'extension_recognized':bool(ext and mimetypes.guess_type(name)[0])})
        report={'files':rows,'total_files':len(rows),
                'total_size_bytes':sum(len(data) for _,data in items)}
        return save_report('.json', json.dumps(report, indent=2))

    if tool_type == 'file_to_base64':
        name,data=items[0]
        encoded=base64.b64encode(data).decode('ascii')
        return save_report('.txt', f'Filename: {name}\\nBase64:\\n{encoded}\\n')
    return jsonify(error='Unsupported File Utilities tool.'), 400

# Text Utilities pages (the tools run locally in the browser; no file upload is required)
TEXT_UTILITY_TOOLS = {
    'word-counter': 'Word Counter',
    'character-counter': 'Character Counter',
    'text-case-converter': 'Text Case Converter',
    'remove-duplicate-lines': 'Remove Duplicate Lines',
    'text-sorter': 'Text Sorter',
    'find-replace-text': 'Find & Replace Text',
    'text-to-speech': 'Text to Speech',
    'speech-to-text': 'Speech to Text',
    'text-diff-checker': 'Text Diff Checker',
    'remove-extra-spaces': 'Remove Extra Spaces',
    'line-counter': 'Line Counter',
    'text-reverser': 'Text Reverser',
    'lorem-ipsum-generator': 'Lorem Ipsum Generator',
    'text-to-binary': 'Text to Binary',
    'binary-to-text': 'Binary to Text',
}

@app.route('/text-utilities/<tool_slug>')
def text_utility_page(tool_slug):
    tool_name = TEXT_UTILITY_TOOLS.get(tool_slug)
    if not tool_name:
        return 'Tool not found', 404
    return render_template('tool_template.html', tool_name=tool_name, tool_type='text_' + tool_slug.replace('-', '_'))


# QR & Barcode tools
QR_BARCODE_TOOLS = {
    'qr_code_generator': 'QR Code Generator',
    'qr_code_scanner': 'QR Code Scanner',
    'wifi_qr_code_generator': 'Wi-Fi QR Code Generator',
    'vcard_qr_code_generator': 'vCard QR Code Generator',
    'barcode_generator': 'Barcode Generator',
    'barcode_scanner': 'Barcode Scanner',
    'qr_code_with_logo': 'QR Code with Logo',
    'qr_code_color_customizer': 'QR Code Color Customizer',
}



# Developer Tools pages (processing is performed locally in the user's browser)
DEVELOPER_TOOLS = {
    'json-formatter': 'JSON Formatter',
    'json-validator': 'JSON Validator',
    'html-formatter': 'HTML Formatter',
    'css-minifier': 'CSS Minifier',
    'javascript-minifier': 'JavaScript Minifier',
    'url-encoder-decoder': 'URL Encoder / Decoder',
    'markdown-to-html': 'Markdown to HTML',
    'password-generator': 'Password Generator',
    'color-picker': 'Color Picker',
    'base64-text-encoder-decoder': 'Base64 Text Encoder / Decoder',
    'html-minifier': 'HTML Minifier',
    'css-formatter': 'CSS Formatter',
    'javascript-formatter': 'JavaScript Formatter',
    'xml-formatter': 'XML Formatter',
    'xml-validator': 'XML Validator',
    'regex-tester': 'Regex Tester',
    'jwt-decoder': 'JWT Decoder',
    'uuid-generator': 'UUID Generator',
}

@app.route('/developer-tools/<tool_slug>')
def developer_tool_page(tool_slug):
    tool_name = DEVELOPER_TOOLS.get(tool_slug)
    if not tool_name:
        return 'Tool not found', 404
    return render_template('tool_template.html', tool_name=tool_name,
                           tool_type='dev_' + tool_slug.replace('-', '_'))

@app.route('/qr-barcode/<tool_slug>')
def qr_barcode_page(tool_slug):
    tool_type = tool_slug.replace('-', '_')
    if tool_type not in QR_BARCODE_TOOLS:
        return 'Tool not found', 404
    return render_template('tool_template.html', tool_name=QR_BARCODE_TOOLS[tool_type], tool_type=tool_type)

def _qr_barcode_process(tool_type, files, form, upload_folder):
    """Generate QR/barcodes or decode uploaded QR/barcode images."""
    from flask import jsonify
    from PIL import Image
    from uuid import uuid4
    from io import BytesIO
    import os

    def save_image(image, suffix='png'):
        filename = f'filevellum_{tool_type}_{uuid4().hex[:12]}.{suffix}'
        path = os.path.join(upload_folder, filename)
        image.save(path, format=suffix.upper())
        return jsonify({'success': True, 'filename': filename})

    if tool_type in ('qr_code_scanner', 'barcode_scanner'):
        if not files or not files[0] or not files[0].filename:
            return jsonify(error='Please upload an image containing a QR code or barcode.'), 400
        try:
            image = Image.open(files[0].stream).convert('RGB')
        except Exception:
            return jsonify(error='The uploaded file is not a valid image.'), 400
        decoded = []
        if tool_type == 'qr_code_scanner':
            try:
                import cv2
                import numpy as np
                arr = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
                detector = cv2.QRCodeDetector()
                value, points, _ = detector.detectAndDecode(arr)
                if value:
                    decoded.append({'type': 'QR Code', 'value': value})
            except Exception:
                pass
        else:
            try:
                from pyzbar.pyzbar import decode
                decoded = [{'type': item.type, 'value': item.data.decode('utf-8', errors='replace')} for item in decode(image)]
            except Exception:
                return jsonify(error='Barcode scanning needs the ZBar library installed on the server (and the Python package pyzbar).'), 501
        if not decoded:
            return jsonify(error='No readable code was detected. Try a clearer, well-lit image.'), 422
        return jsonify({'success': True, 'result': decoded})

    content = form.get('text', '').strip()
    if tool_type == 'wifi_qr_code_generator':
        ssid = form.get('ssid', '').strip()
        password = form.get('wifi_password', '')
        security = form.get('security', 'WPA')
        if not ssid:
            return jsonify(error='Please enter the Wi-Fi network name (SSID).'), 400
        content = f'WIFI:T:{security};S:{ssid};P:{password};;'
    elif tool_type == 'vcard_qr_code_generator':
        name = form.get('contact_name', '').strip()
        phone = form.get('contact_phone', '').strip()
        email = form.get('contact_email', '').strip()
        org = form.get('contact_org', '').strip()
        if not name and not phone and not email:
            return jsonify(error='Enter at least a contact name, phone number, or email.'), 400
        content = 'BEGIN:VCARD\nVERSION:3.0\n' + (f'FN:{name}\n' if name else '') + (f'TEL:{phone}\n' if phone else '') + (f'EMAIL:{email}\n' if email else '') + (f'ORG:{org}\n' if org else '') + 'END:VCARD'
    elif tool_type == 'barcode_generator':
        value = form.get('barcode_value', '').strip()
        barcode_type = form.get('barcode_format', 'code128').lower()
        if not value:
            return jsonify(error='Please enter the value to encode.'), 400
        try:
            import barcode
            from barcode.writer import ImageWriter
            code_class = barcode.get_barcode_class(barcode_type)
            output = BytesIO()
            code_class(value, writer=ImageWriter()).write(output, options={'write_text': True, 'module_height': 15, 'quiet_zone': 3})
            output.seek(0)
            image = Image.open(output).convert('RGB')
            return save_image(image, 'png')
        except ImportError:
            return jsonify(error='Barcode generation dependency missing. Install: pip install "python-barcode[images]"'), 501
        except Exception as exc:
            return jsonify(error=f'Could not generate this barcode. Check the value and selected format. ({exc})'), 400

    if not content:
        return jsonify(error='Please enter the text or URL for your QR code.'), 400
    try:
        import qrcode
        from qrcode.constants import ERROR_CORRECT_H
        from qrcode.exceptions import DataOverflowError
        qr = qrcode.QRCode(version=None, error_correction=ERROR_CORRECT_H, box_size=10, border=4)
        qr.add_data(content)
        qr.make(fit=True)
        fill = form.get('qr_color', '#111111') if tool_type == 'qr_code_color_customizer' else '#111111'
        back = form.get('background_color', '#ffffff') if tool_type == 'qr_code_color_customizer' else '#ffffff'
        image = qr.make_image(fill_color=fill, back_color=back).convert('RGB')
        if tool_type == 'qr_code_with_logo':
            logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static', 'images', 'yashraj-gurjar-founder.png')
            if os.path.isfile(logo_path):
                logo = Image.open(logo_path).convert('RGBA')
                logo.thumbnail((image.width // 4, image.height // 4))
                pad = 10
                tile = Image.new('RGB', (logo.width + pad * 2, logo.height + pad * 2), 'white')
                tile.paste(logo, (pad, pad), logo)
                image.paste(tile, ((image.width - tile.width)//2, (image.height - tile.height)//2))
        return save_image(image, 'png')
    except ImportError:
        return jsonify(error='QR generation dependency missing. Install: pip install "qrcode[pil]"'), 501
    except Exception:
        app.logger.exception('QR generation failed')
        return jsonify(error='Could not generate the QR code. Please check the entered content.'), 400




# Extra Utilities: client-side workspaces plus optional server-side font conversion.
EXTRA_UTILITY_TOOLS = {
    'font-file-converter': ('Font File Converter', 'extra_font_file_converter'),
    'font-previewer': ('Font Previewer', 'extra_font_previewer'),
    'color-contrast-checker': ('Color Contrast Checker', 'extra_color_contrast_checker'),
    'unit-converter': ('Unit Converter', 'extra_unit_converter'),
    'date-calculator': ('Date Calculator', 'extra_date_calculator'),
    'timestamp-converter': ('Timestamp Converter', 'extra_timestamp_converter'),
    'url-parser': ('URL Parser', 'extra_url_parser'),
    'html-entity-encoder-decoder': ('HTML Entity Encoder / Decoder', 'extra_html_entity_encoder_decoder'),
    'css-gradient-generator': ('CSS Gradient Generator', 'extra_css_gradient_generator'),
    'meta-tag-generator': ('Meta Tag Generator', 'extra_meta_tag_generator'),
    'robots-txt-generator': ('Robots.txt Generator', 'extra_robots_txt_generator'),
    'sitemap-generator': ('Sitemap Generator', 'extra_sitemap_generator'),
    'open-graph-preview': ('Open Graph Preview', 'extra_open_graph_preview'),
    'markdown-editor': ('Markdown Editor', 'extra_markdown_editor'),
    'html-live-preview': ('HTML Live Preview', 'extra_html_live_preview'),
    'text-file-merger': ('Text File Merger', 'extra_text_file_merger'),
    'csv-viewer': ('CSV Viewer', 'extra_csv_viewer'),
    'json-viewer': ('JSON Viewer', 'extra_json_viewer'),
    'xml-viewer': ('XML Viewer', 'extra_xml_viewer'),
    'html-table-generator': ('HTML Table Generator', 'extra_html_table_generator'),
}

@app.route('/extra-utilities/<tool_slug>')
def extra_utility_page(tool_slug):
    item = EXTRA_UTILITY_TOOLS.get(tool_slug)
    if not item:
        return 'Tool not found', 404
    return render_template('tool_template.html', tool_name=item[0], tool_type=item[1])


@app.route('/extra-utilities/font-file-converter', methods=['POST'])
def extra_font_file_convert():
    uploaded = request.files.get('file')
    target = (request.form.get('target_format') or '').lower().strip()
    if not uploaded or not uploaded.filename:
        return jsonify({'error': 'Please choose a font file.'}), 400
    if target not in {'ttf', 'otf', 'woff', 'woff2'}:
        return jsonify({'error': 'Choose TTF, OTF, WOFF or WOFF2 as the output format.'}), 400
    try:
        from fontTools.ttLib import TTFont
        from io import BytesIO
        font = TTFont(BytesIO(uploaded.read()))
        font.flavor = {'ttf': None, 'otf': None, 'woff': 'woff', 'woff2': 'woff2'}[target]
        out = BytesIO()
        font.save(out)
        out.seek(0)
        return send_file(out, as_attachment=True,
                         download_name=(os.path.splitext(uploaded.filename)[0] + '.' + target),
                         mimetype='font/woff2' if target == 'woff2' else ('font/woff' if target == 'woff' else 'application/octet-stream'))
    except ImportError:
        return jsonify({'error': 'Font conversion dependency missing. Install it with: pip install fonttools brotli'}), 501
    except Exception as exc:
        return jsonify({'error': 'Font conversion failed: ' + str(exc)[:240]}), 400


# Image Inspection tools — share the existing FileVellum tool page and API pipeline.
IMAGE_INSPECTION_TOOLS = {
    'image-dpi-checker': ('Image DPI Checker', 'image_dpi_checker'),
    'image-aspect-ratio-calculator': ('Image Aspect Ratio Calculator', 'image_aspect_ratio_calculator'),
    'image-exif-viewer': ('Image EXIF Viewer', 'image_exif_viewer'),
    'image-dimension-checker': ('Image Dimension Checker', 'image_dimension_checker'),
    'image-histogram': ('Image Histogram', 'image_histogram'),
    'image-color-picker': ('Image Color Picker', 'image_color_picker'),
    'image-transparency-checker': ('Image Transparency Checker', 'image_transparency_checker'),
    'image-format-detector': ('Image Format Detector', 'image_format_detector'),
    'image-quality-analyzer': ('Image Quality Analyzer', 'image_quality_analyzer'),
    'image-comparison-tool': ('Image Comparison Tool', 'image_comparison_tool'),
}

@app.route('/image-inspection/<tool_slug>')
def image_inspection_page(tool_slug):
    item = IMAGE_INSPECTION_TOOLS.get(tool_slug)
    if not item:
        return 'Tool not found', 404
    return render_template('tool_template.html', tool_name=item[0], tool_type=item[1])


def process_image_inspection(tool_type, files, form):
    """Return safe, structured inspection results for uploaded images."""
    from PIL import Image, ImageChops, ImageFilter, ImageStat, ExifTags
    from io import BytesIO
    from fractions import Fraction
    import math

    if not files or not files[0] or not files[0].filename:
        return jsonify({'error': 'Please select an image first.'}), 400
    try:
        raw = files[0].read()
        if not raw:
            return jsonify({'error': 'The selected image is empty.'}), 400
        img = Image.open(BytesIO(raw))
        img.load()
        width, height = img.size
        fmt = img.format or 'Unknown'
        report = {'File name': files[0].filename, 'Format': fmt, 'Dimensions': f'{width} × {height} px', 'File size': f'{len(raw):,} bytes'}

        if tool_type == 'image_dpi_checker':
            dpi = img.info.get('dpi')
            if isinstance(dpi, (tuple, list)) and len(dpi) >= 2:
                report.update({'Horizontal DPI': round(float(dpi[0]), 2), 'Vertical DPI': round(float(dpi[1]), 2), 'Print size': f'{width/max(float(dpi[0]),1):.2f} × {height/max(float(dpi[1]),1):.2f} inches'})
            else:
                report.update({'DPI': 'Not embedded in this image', 'Note': 'DPI cannot be reliably inferred from pixel dimensions alone.'})
        elif tool_type == 'image_aspect_ratio_calculator':
            ratio = Fraction(width, height)
            report.update({'Aspect ratio': f'{ratio.numerator}:{ratio.denominator}', 'Decimal ratio': round(width / height, 6), 'Orientation': 'Landscape' if width > height else ('Portrait' if height > width else 'Square'), 'Megapixels': round((width*height)/1_000_000, 3)})
        elif tool_type == 'image_exif_viewer':
            exif = img.getexif()
            tags = {}
            for key, value in exif.items():
                tag = ExifTags.TAGS.get(key, str(key))
                if key == 34853 and hasattr(value, 'items'):
                    gps = {ExifTags.GPSTAGS.get(k, str(k)): str(v)[:180] for k, v in value.items()}
                    tags['GPSInfo'] = gps
                else:
                    tags[tag] = str(value)[:500]
            report.update({'EXIF fields found': len(tags), 'EXIF data': tags if tags else 'No EXIF metadata found'})
        elif tool_type == 'image_dimension_checker':
            report.update({'Width': f'{width:,} pixels', 'Height': f'{height:,} pixels', 'Total pixels': f'{width*height:,}', 'Megapixels': round(width*height/1_000_000, 3), 'Mode': img.mode})
        elif tool_type == 'image_histogram':
            rgb = img.convert('RGB')
            hist = rgb.histogram()
            bins = {}
            for channel, offset in [('Red',0),('Green',256),('Blue',512)]:
                bins[channel] = [sum(hist[offset+i*16:offset+(i+1)*16]) for i in range(16)]
            report.update({'Histogram bins': '16 bins per RGB channel', 'Red channel': bins['Red'], 'Green channel': bins['Green'], 'Blue channel': bins['Blue']})
        elif tool_type == 'image_color_picker':
            try:
                x, y = int(form.get('pixel_x', width//2)), int(form.get('pixel_y', height//2))
            except (TypeError, ValueError):
                return jsonify({'error': 'Enter valid pixel coordinates.'}), 400
            if x < 0 or y < 0 or x >= width or y >= height:
                return jsonify({'error': f'Coordinates must be within x: 0–{width-1}, y: 0–{height-1}.'}), 400
            rgba = img.convert('RGBA').getpixel((x,y))
            report.update({'Selected pixel': f'({x}, {y})', 'HEX': '#%02X%02X%02X' % rgba[:3], 'RGB': f'{rgba[0]}, {rgba[1]}, {rgba[2]}', 'Alpha': rgba[3], 'Tip': 'Use pixel coordinates measured from the image top-left corner.'})
        elif tool_type == 'image_transparency_checker':
            has_alpha = 'A' in img.getbands() or 'transparency' in img.info
            transparent_pixels = 0
            if has_alpha:
                alpha = img.convert('RGBA').getchannel('A')
                transparent_pixels = sum(1 for value in alpha.getdata() if value < 255)
            report.update({'Has transparency': 'Yes' if transparent_pixels else 'No', 'Transparent / semi-transparent pixels': f'{transparent_pixels:,}', 'Total pixels': f'{width*height:,}', 'Transparency percentage': round(transparent_pixels*100/(width*height), 3)})
        elif tool_type == 'image_format_detector':
            report.update({'Detected format': fmt, 'MIME type': Image.MIME.get(fmt, 'Unknown'), 'File extension': ('.' + fmt.lower()) if fmt != 'Unknown' else 'Unknown', 'Color mode': img.mode})
        elif tool_type == 'image_quality_analyzer':
            gray = img.convert('L')
            stat = ImageStat.Stat(gray)
            edge_stat = ImageStat.Stat(gray.filter(ImageFilter.FIND_EDGES))
            entropy = gray.entropy()
            sharpness = round(edge_stat.var[0], 2)
            report.update({'Brightness (0–255)': round(stat.mean[0], 2), 'Contrast (standard deviation)': round(stat.stddev[0], 2), 'Entropy (0–8)': round(entropy, 3), 'Edge detail score': sharpness, 'Analysis note': 'These are image characteristics, not a guaranteed subjective quality or print-quality rating.'})
        elif tool_type == 'image_comparison_tool':
            second = request.files.get('second_image')
            if not second or not second.filename:
                return jsonify({'error': 'Please select a second image to compare.'}), 400
            raw2 = second.read()
            im2 = Image.open(BytesIO(raw2)); im2.load()
            report.update({'Image A dimensions': f'{width} × {height}', 'Image B': second.filename, 'Image B dimensions': f'{im2.width} × {im2.height}', 'Image A size': f'{len(raw):,} bytes', 'Image B size': f'{len(raw2):,} bytes', 'Same dimensions': 'Yes' if img.size == im2.size else 'No'})
            a = img.convert('RGB').resize((256,256)); b = im2.convert('RGB').resize((256,256))
            diff = ImageChops.difference(a,b)
            mean_diff = sum(ImageStat.Stat(diff).mean)/3
            report.update({'Average pixel difference (0–255)': round(mean_diff, 3), 'Approximate similarity': f'{max(0,100-mean_diff/255*100):.2f}%', 'Comparison method': 'Images resized to 256 × 256 for a normalized pixel comparison.'})
        else:
            return jsonify({'error': 'Unsupported image inspection tool.'}), 400
        return jsonify({'success': True, 'inspection': True, 'tool': tool_type, 'report': report})
    except Exception as exc:
        return jsonify({'error': 'Unable to inspect this image. Please try a valid JPG, PNG, WEBP, TIFF, BMP or GIF file. ' + str(exc)[:140]}), 400


@app.route('/api/process', methods=['POST'])
def api_process():
    files = request.files.getlist('file') or request.files.getlist('file[]')
    tool_type = request.form.get('tool_type', '').strip()
    if tool_type in {v[1] for v in IMAGE_CONVERSION_TOOLS.values()}:
        return process_image_conversion(tool_type, files, request.form, app.config['UPLOAD_FOLDER'])
    if tool_type in DOCUMENT_CONVERSION_TOOLS:
        return process_document_conversion(tool_type, files)
    if tool_type in FILE_UTILITY_TOOLS:
        return _file_utility_process(tool_type, files, request.form, app.config['UPLOAD_FOLDER'])
    if tool_type in QR_BARCODE_TOOLS:
        return _qr_barcode_process(tool_type, files, request.form, app.config['UPLOAD_FOLDER'])
    if tool_type in {v[1] for v in IMAGE_INSPECTION_TOOLS.values()}:
        return process_image_inspection(tool_type, files, request.form)
    if tool_type in {v[1] for v in ARCHIVE_TOOLS.values()} or tool_type in ('compressor', 'extract_zip'):
        result = process_archive(tool_type, files, request.form, app.config['UPLOAD_FOLDER'])
        if isinstance(result, tuple):
            payload, status = result
            return jsonify(payload), status
        return jsonify(result)
    if 'file' not in request.files and 'file[]' not in request.files and tool_type != 'webpage_to_pdf':
        return jsonify({'error': 'No files uploaded'}), 400
    
    try:
        if tool_type in IMAGE_PROCESS_TYPES:
            return process_image_tool(tool_type, files, request.form, app.config['UPLOAD_FOLDER'])
        if tool_type in EDIT_SECURE_TOOLS:
            return process_edit_secure(tool_type, files, app.config['UPLOAD_FOLDER'])
        if tool_type in OPTIMIZE_OCR_TOOLS:
            return process_optimize_ocr(tool_type, files, request.form, app.config['UPLOAD_FOLDER'])
        if tool_type in ORGANIZE_EXTRA_TOOLS:
            return process_organize_extra(tool_type, files)

        # 1. Image Editor Tool (Manual Crop, Resize, Rotate)
        if tool_type == 'edit_image':
            if not files or not files[0] or not files[0].filename:
                return jsonify({'success': False, 'error': 'Please upload an image first.'}), 400

            uploaded = files[0]
            original_name = secure_filename(uploaded.filename)
            extension = os.path.splitext(original_name)[1].lower()
            if extension not in {'.jpg', '.jpeg', '.png'}:
                return jsonify({'success': False, 'error': 'Please upload a JPG or PNG image.'}), 400

            raw = uploaded.read()
            if not raw:
                return jsonify({'success': False, 'error': 'The uploaded image is empty.'}), 400
            if len(raw) > 30 * 1024 * 1024:
                return jsonify({'success': False, 'error': 'Image size must not exceed 30 MB.'}), 413

            try:
                with Image.open(BytesIO(raw)) as source:
                    source.verify()
                with Image.open(BytesIO(raw)) as source:
                    img = ImageOps.exif_transpose(source)
                    if img.width < 1 or img.height < 1 or img.width * img.height > 25_000_000:
                        return jsonify({'success': False, 'error': 'Image dimensions must not exceed 25 megapixels.'}), 400
                    img = img.copy()
            except (OSError, ValueError, Image.DecompressionBombError):
                return jsonify({'success': False, 'error': 'The uploaded file is not a valid JPG or PNG image.'}), 400

            # Crop coordinates are reported against the original preview, so crop first.
            try:
                crop_values = []
                for key in ('crop_x', 'crop_y', 'crop_w', 'crop_h'):
                    raw_value = request.form.get(key, '')
                    value = float(raw_value) if raw_value not in (None, '') else None
                    if value is not None and (not math.isfinite(value) or abs(value) > 25_000):
                        raise ValueError('Invalid crop coordinates.')
                    crop_values.append(value)

                cx, cy, cw, ch = crop_values
                if any(value is not None for value in crop_values):
                    if None in crop_values or cw <= 0 or ch <= 0 or cx < 0 or cy < 0:
                        return jsonify({'success': False, 'error': 'Please provide a valid crop area.'}), 400
                    right, bottom = cx + cw, cy + ch
                    if right > img.width + 1 or bottom > img.height + 1:
                        return jsonify({'success': False, 'error': 'The selected crop area is outside the image.'}), 400
                    left, top = max(0, int(round(cx))), max(0, int(round(cy)))
                    right, bottom = min(img.width, int(round(right))), min(img.height, int(round(bottom)))
                    if right <= left or bottom <= top:
                        return jsonify({'success': False, 'error': 'The selected crop area is too small.'}), 400
                    img = img.crop((left, top, right, bottom))
            except (TypeError, ValueError, OverflowError):
                return jsonify({'success': False, 'error': 'Crop coordinates must be valid numbers.'}), 400

            try:
                rotate_raw = request.form.get('rotate', '0')
                rotate_value = float(rotate_raw)
                if not math.isfinite(rotate_value) or rotate_value not in (0, 90, 180, 270):
                    raise ValueError
                rotate_angle = int(rotate_value)
            except (TypeError, ValueError, OverflowError):
                return jsonify({'success': False, 'error': 'Rotation must be 0, 90, 180, or 270 degrees.'}), 400

            if rotate_angle:
                img = img.rotate(-rotate_angle, expand=True)

            try:
                width_raw = (request.form.get('width') or '').strip()
                height_raw = (request.form.get('height') or '').strip()
                if width_raw or height_raw:
                    width = int(width_raw) if width_raw else None
                    height = int(height_raw) if height_raw else None
                    if ((width is not None and not 1 <= width <= 12000) or
                            (height is not None and not 1 <= height <= 12000)):
                        raise ValueError
                    if width is None:
                        width = max(1, int(round(img.width * height / img.height)))
                    elif height is None:
                        height = max(1, int(round(img.height * width / img.width)))
                    if width * height > 25_000_000:
                        return jsonify({'success': False, 'error': 'Output dimensions must not exceed 25 megapixels.'}), 400
                    img = img.resize((width, height), Image.Resampling.LANCZOS)
            except (TypeError, ValueError, OverflowError, ZeroDivisionError):
                return jsonify({'success': False, 'error': 'Width and height must be whole numbers between 1 and 12000 pixels.'}), 400

            base_name = os.path.splitext(original_name)[0] or 'image'
            output_extension = '.png' if extension == '.png' else '.jpg'
            out_filename = f"{uuid4().hex}_{base_name}_edited{output_extension}"
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            try:
                if output_extension == '.png':
                    img.save(out_path, 'PNG', optimize=True)
                else:
                    if img.mode not in ('RGB', 'L'):
                        background = Image.new('RGB', img.size, (255, 255, 255))
                        if 'A' in img.getbands():
                            background.paste(img, mask=img.getchannel('A'))
                        else:
                            background.paste(img.convert('RGB'))
                        img = background
                    else:
                        img = img.convert('RGB')
                    img.save(out_path, 'JPEG', quality=95, optimize=True)
            except OSError:
                return jsonify({'success': False, 'error': 'Unable to save the processed image. Please try again.'}), 500

            return jsonify({'success': True, 'filename': out_filename})

        # ZIP & Archives are dispatched to archive_backend above.

        # 4. PDF to JPG Converter
        elif tool_type in ['pdf_to_jpg', 'pdf-to-jpg']:
            file = files[0]
            if not file.filename.lower().endswith('.pdf'):
                return jsonify({'error': 'Please upload a valid PDF file.'}), 400
                
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
            file.save(pdf_path)
            
            doc = fitz.open(pdf_path)
            image_filenames = []
            base_name = file.filename.rsplit('.', 1)[0]
            
            for idx, page in enumerate(doc):
                pix = page.get_pixmap(dpi=150)
                img_filename = f"{base_name}_page_{idx+1}.jpg"
                img_path = os.path.join(app.config['UPLOAD_FOLDER'], img_filename)
                Image.frombytes('RGB', (pix.width, pix.height), pix.samples).save(img_path, 'JPEG', quality=90)
                image_filenames.append(img_filename)
            
            if len(image_filenames) == 1:
                return jsonify({'success': True, 'filename': image_filenames[0]})
            else:
                zip_filename = f"{base_name}_images.zip"
                zip_path = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for img_name in image_filenames:
                        zipf.write(os.path.join(app.config['UPLOAD_FOLDER'], img_name), img_name)
                return jsonify({'success': True, 'filename': zip_filename})

        # 5. PDF to PNG Converter
        elif tool_type in ['pdf_to_png', 'pdf-to-png']:
            file = files[0]
            if not file.filename.lower().endswith('.pdf'):
                return jsonify({'error': 'Please upload a valid PDF file.'}), 400
                
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
            file.save(pdf_path)
            
            doc = fitz.open(pdf_path)
            image_filenames = []
            base_name = file.filename.rsplit('.', 1)[0]
            
            for idx, page in enumerate(doc):
                pix = page.get_pixmap(dpi=150)
                img_filename = f"{base_name}_page_{idx+1}.png"
                img_path = os.path.join(app.config['UPLOAD_FOLDER'], img_filename)
                pix.save(img_path)
                image_filenames.append(img_filename)
            
            if len(image_filenames) == 1:
                return jsonify({'success': True, 'filename': image_filenames[0]})
            else:
                zip_filename = f"{base_name}_png_images.zip"
                zip_path = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for img_name in image_filenames:
                        zipf.write(os.path.join(app.config['UPLOAD_FOLDER'], img_name), img_name)
                return jsonify({'success': True, 'filename': zip_filename})

        # 6. PDF to PowerPoint Converter
        elif tool_type in ['pdf_to_powerpoint', 'pdf-to-powerpoint', 'pdf-to-ppt']:
            file = files[0]
            if not file.filename.lower().endswith('.pdf'):
                return jsonify({'error': 'Please upload a valid PDF file.'}), 400
                
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
            file.save(pdf_path)
            
            ppt_filename = file.filename.rsplit('.', 1)[0] + '.pptx'
            ppt_path = os.path.join(app.config['UPLOAD_FOLDER'], ppt_filename)
            
            prs = Presentation()
            blank_slide_layout = prs.slide_layouts[6]
            
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text(layout=True)
                    slide = prs.slides.add_slide(blank_slide_layout)
                    
                    if text:
                        left = Inches(0.8)
                        top = Inches(0.8)
                        width = Inches(11.5)
                        height = Inches(6.5)
                        
                        txBox = slide.shapes.add_textbox(left, top, width, height)
                        tf = txBox.text_frame
                        tf.word_wrap = True
                        
                        lines = text.split('\n')
                        for idx, line in enumerate(lines):
                            if line.strip():
                                p = tf.add_paragraph() if idx > 0 else tf.paragraphs[0]
                                p.text = line.rstrip()
                                p.font.size = Pt(13)

            prs.save(ppt_path)
            return jsonify({'success': True, 'filename': ppt_filename})

        # 7. PDF to Excel Converter
        elif tool_type in ['pdf_to_excel', 'pdf-to-excel']:
            file = files[0]
            if not file.filename.lower().endswith('.pdf'):
                return jsonify({'error': 'Please upload a valid PDF file.'}), 400
                
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
            file.save(pdf_path)
            
            excel_filename = file.filename.rsplit('.', 1)[0] + '.xlsx'
            excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
            
            rows_data = []
            
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages, start=1):
                    tables = page.extract_tables()
                    table_extracted = False
                    
                    if tables:
                        for table in tables:
                            if table:
                                table_extracted = True
                                for row in table:
                                    cleaned_row = [str(cell).strip() if cell is not None else "" for cell in row]
                                    rows_data.append([f"Page {page_num}"] + cleaned_row)
                    
                    if not table_extracted:
                        text = page.extract_text(layout=True)
                        if text:
                            lines = text.split('\n')
                            for line in lines:
                                if line.strip():
                                    rows_data.append([f"Page {page_num}", line.strip()])

            if not rows_data:
                rows_data = [["Status", "No text or data was found in the PDF."]]

            df = pd.DataFrame(rows_data)
            
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Extracted Data', index=False, header=False)

            return jsonify({'success': True, 'filename': excel_filename})

        # 8. PDF to Word Converter
        elif tool_type in ['pdf_to_word', 'pdf-to-word']:
            file = files[0]
            if not file.filename.lower().endswith('.pdf'):
                return jsonify({'error': 'Please upload a valid PDF file.'}), 400
                
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
            file.save(pdf_path)
            
            word_filename = file.filename.rsplit('.', 1)[0] + '.docx'
            word_path = os.path.join(app.config['UPLOAD_FOLDER'], word_filename)
            
            doc = docx.Document()
            
            with pdfplumber.open(pdf_path) as pdf:
                for idx, page in enumerate(pdf.pages):
                    text = page.extract_text(layout=True)
                    if text:
                        lines = text.split('\n')
                        for line in lines:
                            if line.strip():
                                doc.add_paragraph(line.rstrip())
                    
                    if idx < len(pdf.pages) - 1:
                        doc.add_page_break()

            doc.save(word_path)
            return jsonify({'success': True, 'filename': word_filename})

        # 9. Merge PDF Tool
        elif tool_type == 'merge':
            out_filename = 'merged_output.pdf'
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            writer = PdfWriter()
            for file in files:
                if not file.filename.lower().endswith('.pdf'):
                    return jsonify({'error': 'Sabhi files PDF honi chahiye'}), 400
                reader = PdfReader(file)
                for page in reader.pages:
                    writer.add_page(page)
            with open(out_path, 'wb') as f:
                writer.write(f)
            return jsonify({'success': True, 'filename': out_filename})

        # 10. Split PDF Tool
        elif tool_type == 'split':
            file = files[0]
            if not file.filename.lower().endswith('.pdf'):
                return jsonify({'error': 'Please upload a PDF file.'}), 400
            range_str = request.form.get('split_range', '').strip()
            reader = PdfReader(file)
            writer = PdfWriter()
            total_pages = len(reader.pages)
            
            if not range_str:
                return jsonify({'error': 'Please select a page.'}), 400

            try:
                p_num = int(range_str.split(',')[0].strip())
                p_idx = p_num - 1
                if 0 <= p_idx < total_pages:
                    writer.add_page(reader.pages[p_idx])
                else:
                    return jsonify({'error': f'Invalid page. Total pages: {total_pages}'}), 400
            except Exception as e:
                return jsonify({'error': 'Unable to select the page. Please try again.'}), 400

            out_filename = 'split_output.pdf'
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            with open(out_path, 'wb') as f:
                writer.write(f)
            return jsonify({'success': True, 'filename': out_filename})

        # 11. Organize Pages Tool
        elif tool_type == 'organize_pdf':
            writer = PdfWriter()
            all_pages = []
            
            for file in files:
                if not file.filename.lower().endswith('.pdf'):
                    return jsonify({'error': 'Sabhi files PDF honi chahiye'}), 400
                reader = PdfReader(file)
                for page in reader.pages:
                    all_pages.append(page)
            
            total_pages = len(all_pages)
            order_str = request.form.get('split_range', '').strip()
            
            if not order_str:
                return jsonify({'error': 'Please select one or more pages.'}), 400

            try:
                parts = order_str.split(',')
                for part in parts:
                    if part.strip():
                        p_num = int(part.strip())
                        p_idx = p_num - 1
                        if 0 <= p_idx < total_pages:
                            writer.add_page(all_pages[p_idx])
            except Exception as to_err:
                return jsonify({'error': 'Unable to update the page order. Please try again.'}), 400

            out_filename = 'organized_output.pdf'
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            with open(out_path, 'wb') as f:
                writer.write(f)
            return jsonify({'success': True, 'filename': out_filename})

        # 12. Compress PDF Tool
        elif tool_type in ['compress-pdf', 'compress_pdf']:
            file = files[0]
            if not file.filename.lower().endswith('.pdf'):
                return jsonify({'error': 'Please upload a PDF file.'}), 400
            reader = PdfReader(file)
            writer = PdfWriter()
            for page in reader.pages:
                page.compress_content_streams()
                writer.add_page(page)
            out_filename = 'compressed_output.pdf'
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            with open(out_path, 'wb') as f:
                writer.write(f)
            return jsonify({'success': True, 'filename': out_filename})

        # 13. Universal Compression Tool (PDF, JPG, PNG)
        elif tool_type == 'compress_file':
            file = files[0]
            filename = secure_filename(file.filename)
            file_ext = filename.lower().rsplit('.', 1)[1]
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(input_path)
            
            if file_ext == 'pdf':
                reader = PdfReader(input_path)
                writer = PdfWriter()
                for page in reader.pages:
                    page.compress_content_streams()
                    writer.add_page(page)
                out_filename = f"compressed_{filename}"
                out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
                with open(out_path, 'wb') as f:
                    writer.write(f)
                return jsonify({'success': True, 'filename': out_filename})
            
            elif file_ext in ['jpg', 'jpeg']:
                out_filename = f"compressed_{filename}"
                out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
                img = Image.open(input_path)
                img.save(out_path, 'JPEG', quality=65, optimize=True)
                return jsonify({'success': True, 'filename': out_filename})
            
            elif file_ext == 'png':
                out_filename = f"compressed_{filename}"
                out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
                img = Image.open(input_path)
                img.save(out_path, 'PNG', optimize=True)
                return jsonify({'success': True, 'filename': out_filename})
            
            else:
                return jsonify({'error': 'Please upload a valid PDF, JPG, or PNG file.'}), 400

        # 14. Background Removal Tool (rembg / U2Net)
        elif tool_type in ('remove_bg', 'remove-bg'):
            file = files[0] if files else None
            if file is None or not file.filename or not file.filename.lower().endswith(('.jpg', '.jpeg', '.png', '.webp')):
                return jsonify({'error': 'Please upload a JPG, PNG or WEBP image.'}), 400

            try:
                # Validate actual image content, not just its extension.
                source = file.read()
                with Image.open(BytesIO(source)) as image:
                    image.verify()
                with Image.open(BytesIO(source)) as image:
                    if image.width * image.height > 25_000_000:
                        return jsonify({'error': 'Image too large. Maximum 25 megapixels.'}), 400
            except (OSError, ValueError, Image.DecompressionBombError):
                return jsonify({'error': 'Invalid or damaged image file.'}), 400

            try:
                from rembg import remove
                output_bytes = remove(source, session=get_bg_session())
            except (ImportError, ModuleNotFoundError) as exc:
                app.logger.exception('Background removal dependency missing')
                return jsonify({'error': 'Background removal is not installed. Install rembg[cpu] and restart the server.', 'details': str(exc)}), 503
            except Exception as exc:
                app.logger.exception('Background removal failed')
                return jsonify({'error': 'Background removal failed. Check server logs and AI model download.', 'details': str(exc)}), 500

            base_name = os.path.splitext(secure_filename(file.filename))[0] or 'image'
            out_filename = f"bg_removed_{base_name}_{uuid4().hex[:10]}.png"
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            with Image.open(BytesIO(output_bytes)) as output_image:
                output_image.convert('RGBA').save(out_path, 'PNG')
            return jsonify({'success': True, 'filename': out_filename})

        # 15. Universal Repeating Full-Page Watermark Tool
        elif tool_type == 'watermark':
            file = files[0] if files else None
            if not file:
                return "Please upload a file.", 400
                
            filename = secure_filename(file.filename)
            file_ext = filename.lower().rsplit('.', 1)[1] if '.' in filename else ''
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(input_path)
            
            wm_text = request.form.get('watermark_text', 'CONFIDENTIAL').strip()
            brightness = int(request.form.get('brightness', 30))  # 10 to 100
            
            out_filename = 'watermarked_' + filename
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            
            if file_ext == 'pdf':
                doc = fitz.open(input_path)
                color_val = 1.0 - (brightness / 100.0 * 0.6)
                
                for page in doc:
                    rect = page.rect
                    w, h = rect.width, rect.height
                    step_x = 180
                    step_y = 120
                    
                    y = -50
                    while y < h + 100:
                        x = -50
                        while x < w + 100:
                            point = fitz.Point(x, y)
                            page.insert_text(point, wm_text, fontsize=18, color=(color_val, color_val, color_val), rotate=35)
                            x += step_x
                        y += step_y
                        
                doc.save(out_path)
                doc.close()
                
            elif file_ext in ['jpg', 'jpeg', 'png']:
                from PIL import ImageDraw, ImageFont
                base_img = Image.open(input_path).convert('RGBA')
                w, h = base_img.size
                
                overlay_size = int(max(w, h) * 2)
                txt_img = Image.new('RGBA', (overlay_size, overlay_size), (255, 255, 255, 0))
                draw = ImageDraw.Draw(txt_img)
                
                alpha_val = int((brightness / 100.0) * 150) + 15
                font_size = int(min(w, h) / 25) or 20
                
                try:
                    font = ImageFont.truetype("arial.ttf", font_size)
                except IOError:
                    font = ImageFont.load_default()
                
                step_x = font_size * len(wm_text) * 0.8 + 60
                step_y = font_size * 4
                
                y = 0
                while y < overlay_size:
                    x = 0
                    while x < overlay_size:
                        draw.text((x, y), wm_text, fill=(120, 120, 120, alpha_val), font=font)
                        x += step_x
                    y += step_y
                
                rotated_txt = txt_img.rotate(35, expand=False)
                crop_x = (overlay_size - w) // 2
                crop_y = (overlay_size - h) // 2
                cropped_overlay = rotated_txt.crop((crop_x, crop_y, crop_x + w, crop_y + h))
                
                watermarked = Image.alpha_composite(base_img, cropped_overlay)
                
                if filename.lower().endswith('.png'):
                    watermarked.save(out_path, 'PNG')
                else:
                    watermarked.convert('RGB').save(out_path, 'JPEG', quality=95)
            else:
                return "Unsupported file format!", 400
                
            return send_file(out_path, as_attachment=True, download_name=out_filename)

        # 16. Sign PDF or Image Tool
        elif tool_type == 'sign_pdf':
            file = files[0] if files else None
            if not file or not file.filename:
                return jsonify({'error': 'Please upload a PDF or image file.'}), 400

            filename = secure_filename(file.filename)
            file_ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else ''
            if file_ext not in {'pdf', 'jpg', 'jpeg', 'png'}:
                return jsonify({'error': 'Supported document formats: PDF, JPG and PNG.'}), 400

            raw_document = file.read()
            max_document_size = 100 * 1024 * 1024 if file_ext == 'pdf' else 30 * 1024 * 1024
            if not raw_document or len(raw_document) > max_document_size:
                return jsonify({'error': 'The uploaded file is empty or exceeds the allowed size.'}), 400

            # The signature is sent as a data URL by both the drawing pad and upload tab.
            drawn_data = (request.form.get('drawn_signature') or '').strip()
            if not drawn_data:
                return jsonify({'error': 'Draw or upload your signature before applying it.'}), 400

            try:
                header, encoded = drawn_data.split(',', 1)
                if header not in ('data:image/png;base64', 'data:image/jpeg;base64', 'data:image/jpg;base64'):
                    return jsonify({'error': 'Signature must be a PNG or JPG image.'}), 400
                signature_bytes = base64.b64decode(encoded, validate=True)
                if not signature_bytes or len(signature_bytes) > 5 * 1024 * 1024:
                    return jsonify({'error': 'Signature image is empty or larger than 5 MB.'}), 400
                with Image.open(BytesIO(signature_bytes)) as signature_source:
                    signature_source.load()
                    signature_image = signature_source.convert('RGBA')
                    signature_buffer = BytesIO()
                    signature_image.save(signature_buffer, format='PNG')
                    signature_png = signature_buffer.getvalue()
            except (ValueError, TypeError, OSError, base64.binascii.Error):
                return jsonify({'error': 'The signature image is invalid. Please draw or upload it again.'}), 400

            try:
                x = float(request.form.get('pos_x', 50))
                y = float(request.form.get('pos_y', 50))
                w = float(request.form.get('width', 150))
                h = float(request.form.get('height', 60))
                scale = float(request.form.get('render_scale', 1.2))
            except (TypeError, ValueError):
                return jsonify({'error': 'Invalid signature position or size.'}), 400

            if not all(map(lambda value: value == value and abs(value) != float('inf'), (x, y, w, h, scale))):
                return jsonify({'error': 'Invalid signature position or size.'}), 400
            if scale <= 0 or w <= 0 or h <= 0:
                return jsonify({'error': 'Signature size and preview scale must be greater than zero.'}), 400

            output_stem = os.path.splitext(filename)[0] or 'document'
            out_filename = f"signed_{output_stem}_{uuid4().hex[:10]}.{file_ext}"
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)

            if file_ext == 'pdf':
                if not raw_document.startswith(b'%PDF-'):
                    return jsonify({'error': 'The uploaded file is not a valid PDF.'}), 400
                try:
                    doc = fitz.open(stream=raw_document, filetype='pdf')
                    if doc.needs_pass:
                        doc.close()
                        return jsonify({'error': 'Password-protected PDFs must be unlocked before signing.'}), 400
                    if doc.page_count < 1:
                        doc.close()
                        return jsonify({'error': 'The PDF contains no pages.'}), 400

                    # The current preview displays page 1, so apply the signature to page 1.
                    page = doc[0]
                    page_rect = page.rect
                    pdf_x, pdf_y = x / scale, y / scale
                    pdf_w, pdf_h = w / scale, h / scale
                    left = min(max(page_rect.x0, pdf_x), page_rect.x1)
                    top = min(max(page_rect.y0, pdf_y), page_rect.y1)
                    right = min(page_rect.x1, left + pdf_w)
                    bottom = min(page_rect.y1, top + pdf_h)
                    if right <= left or bottom <= top:
                        doc.close()
                        return jsonify({'error': 'The signature is outside the visible PDF page. Move it onto the page and try again.'}), 400

                    page.insert_image(fitz.Rect(left, top, right, bottom), stream=signature_png, keep_proportion=True, overlay=True)
                    doc.save(out_path, garbage=4, deflate=True)
                    doc.close()
                except Exception:
                    try:
                        doc.close()
                    except Exception:
                        pass
                    raise

            else:
                try:
                    with Image.open(BytesIO(raw_document)) as source_image:
                        source_image.load()
                        base_image = source_image.convert('RGBA')
                    image_scale = base_image.width / max(1, int(request.form.get('preview_width', base_image.width) or base_image.width))
                    # Prefer the preview's explicit natural-pixels-per-CSS-pixel scale.
                    image_scale = scale if scale > 0 else image_scale
                    image_x = max(0, min(base_image.width, int(round(x * image_scale))))
                    image_y = max(0, min(base_image.height, int(round(y * image_scale))))
                    image_w = max(1, min(base_image.width - image_x, int(round(w * image_scale))))
                    image_h = max(1, min(base_image.height - image_y, int(round(h * image_scale))))
                    with Image.open(BytesIO(signature_png)) as sig:
                        sig_image = sig.convert('RGBA').resize((image_w, image_h), Image.Resampling.LANCZOS)
                    base_image.alpha_composite(sig_image, (image_x, image_y))
                    if file_ext == 'png':
                        base_image.save(out_path, 'PNG', optimize=True)
                    else:
                        base_image.convert('RGB').save(out_path, 'JPEG', quality=95, optimize=True)
                except (OSError, ValueError):
                    return jsonify({'error': 'Unable to read the uploaded image. Please try another JPG or PNG.'}), 400

            return jsonify({'success': True, 'filename': out_filename})

       # 17. Universal Rotate Tool
        elif tool_type in ['rotate-pdf', 'rotate_pdf']:
            file = files[0] if files else None
            if not file:
                return "Please upload a file.", 400
                
            filename = secure_filename(file.filename)
            file_ext = filename.lower().rsplit('.', 1)[1] if '.' in filename else ''
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(input_path)
            
            angle = int(request.form.get('angle', 90))
            out_filename = 'rotated_' + filename
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            
            if file_ext == 'pdf':
                doc = fitz.open(input_path)
                for page in doc:
                    page.set_rotation((page.rotation + angle) % 360)
                doc.save(out_path)
                doc.close()
                
            elif file_ext in ['jpg', 'jpeg', 'png']:
                img = Image.open(input_path)
                rotated_img = img.rotate(-angle, expand=True)
                
                if filename.lower().endswith('.png'):
                    rotated_img.save(out_path, 'PNG')
                else:
                    rotated_img.convert('RGB').save(out_path, 'JPEG', quality=95)
            else:
                return "Unsupported file format!", 400
                
            return send_file(out_path, as_attachment=True, download_name=out_filename)

       # 18. Protect PDF Tool
        elif tool_type in ['protect-pdf', 'protect_pdf']:
            file = files[0] if files else None
            if not file or not file.filename.lower().endswith('.pdf'):
                return "Please upload a valid PDF file.", 400
            
            password = request.form.get('password', '').strip()
            if not password:
                return "Please enter the security password.", 400
                
            filename = secure_filename(file.filename)
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(input_path)
            
            doc = fitz.open(input_path)
            out_filename = 'protected_' + filename
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            
            doc.save(out_path, encryption=fitz.PDF_ENCRYPT_AES_128, user_pw=password, owner_pw=password)
            doc.close()
            
            return send_file(out_path, as_attachment=True, download_name=out_filename)

        # 19. Unlock PDF Tool
        elif tool_type in ['unlock-pdf', 'unlock_pdf']:
            file = files[0] if files else None
            if not file or not file.filename.lower().endswith('.pdf'):
                return "Please upload a valid PDF file.", 400
            
            password = request.form.get('password', '').strip()
            filename = secure_filename(file.filename)
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(input_path)
            
            doc = fitz.open(input_path)
            if doc.is_encrypted:
                auth = doc.authenticate(password)
                if not auth:
                    doc.close()
                    return "Incorrect password. Please enter the correct password.", 400
            
            out_filename = 'unlocked_' + filename
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            
            doc.save(out_path)
            doc.close()
            
            return send_file(out_path, as_attachment=True, download_name=out_filename)

        # 20. Strict JPG to PDF Tool
        elif tool_type in ['jpg_to_pdf', 'jpg-to-pdf']:
            image_list = []
            for file in files:
                if not file.filename.lower().endswith(('.jpg', '.jpeg')):
                    return jsonify({'error': 'Please upload JPG or JPEG images only.'}), 400
                img = Image.open(file).convert('RGB')
                image_list.append(img)
            
            out_filename = 'jpg_to_pdf.pdf'
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            if image_list:
                image_list[0].save(out_path, save_all=True, append_images=image_list[1:])
            return jsonify({'success': True, 'filename': out_filename})

        # 21. Strict PNG to PDF Tool
        elif tool_type in ['png_to_pdf', 'png-to-pdf']:
            image_list = []
            for file in files:
                if not file.filename.lower().endswith('.png'):
                    return jsonify({'error': 'Please upload PNG images only.'}), 400
                
                img = Image.open(file)
                if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
                    background = Image.new('RGB', img.size, (255, 255, 255))
                    background.paste(img, mask=img.convert('RGBA').split()[3])
                    rgb_img = background
                else:
                    rgb_img = img.convert('RGB')
                
                image_list.append(rgb_img)
            
            out_filename = 'png_to_pdf.pdf'
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            if image_list:
                image_list[0].save(out_path, save_all=True, append_images=image_list[1:])
            return jsonify({'success': True, 'filename': out_filename})

        # 22. Strict PNG to JPG Tool
        elif tool_type in ['png_to_jpg', 'png-to-jpg']:
            file = files[0]
            if not file.filename.lower().endswith('.png'):
                return jsonify({'error': 'Please upload a PNG image.'}), 400
            
            base_name = file.filename.rsplit('.', 1)[0]
            out_filename = f"{base_name}.jpg"
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            
            img = Image.open(file)
            if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
                background = Image.new('RGB', img.size, (255, 255, 255))
                background.paste(img, mask=img.convert('RGBA').split()[3])
                rgb_img = background
            else:
                rgb_img = img.convert('RGB')
                
            rgb_img.save(out_path, 'JPEG', quality=95)
            return jsonify({'success': True, 'filename': out_filename})

        # 23. Strict JPG to PNG Tool
        elif tool_type in ['jpg_to_png', 'jpg-to-png']:
            file = files[0]
            if not file.filename.lower().endswith(('.jpg', '.jpeg')):
                return jsonify({'error': 'Please upload a JPG or JPEG image.'}), 400
            
            base_name = file.filename.rsplit('.', 1)[0]
            out_filename = f"{base_name}.png"
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            
            img = Image.open(file)
            img.save(out_path, 'PNG')
            return jsonify({'success': True, 'filename': out_filename})

        # 24. Strict Word to PDF Tool
        elif tool_type in ['word_to_pdf', 'word-to-pdf']:
            file = files[0]
            if not file.filename.lower().endswith(('.doc', '.docx')):
                return jsonify({'error': 'Please upload a Word (.doc or .docx) file.'}), 400
            
            input_filename = secure_filename(file.filename)
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
            file.save(input_path)
            
            doc_obj = docx.Document(input_path)
            pdf_doc = fitz.open()
            page = pdf_doc.new_page()
            
            y_position = 50
            for para in doc_obj.paragraphs:
                text = para.text.strip()
                if text:
                    page.insert_text((50, y_position), text, fontsize=11, color=(0, 0, 0))
                    y_position += 20
                    if y_position > 750:
                        page = pdf_doc.new_page()
                        y_position = 50
            
            out_filename = input_filename.rsplit('.', 1)[0] + '.pdf'
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            pdf_doc.save(out_path)
            pdf_doc.close()
            
            return jsonify({'success': True, 'filename': out_filename})

        # 25. Strict Excel to PDF Tool
        elif tool_type in ['excel_to_pdf', 'excel-to-pdf']:
            file = files[0]
            if not file.filename.lower().endswith(('.xls', '.xlsx')):
                return jsonify({'error': 'Please upload an Excel (.xls or .xlsx) file.'}), 400
            
            input_filename = secure_filename(file.filename)
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
            file.save(input_path)
            
            df = pd.read_excel(input_path)
            pdf_doc = fitz.open()
            page = pdf_doc.new_page()
            
            y_position = 50
            header_text = " | ".join([str(col) for col in df.columns])
            page.insert_text((50, y_position), header_text, fontsize=10, color=(0, 0, 0))
            y_position += 25
            
            for index, row in df.iterrows():
                row_text = " | ".join([str(val) for val in row.values])
                page.insert_text((50, y_position), row_text, fontsize=9, color=(0, 0, 0))
                y_position += 18
                if y_position > 750:
                    page = pdf_doc.new_page()
                    y_position = 50
            
            out_filename = input_filename.rsplit('.', 1)[0] + '.pdf'
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            pdf_doc.save(out_path)
            pdf_doc.close()
            
            return jsonify({'success': True, 'filename': out_filename})

        # 26. Strict PowerPoint to PDF Tool
        elif tool_type in ['powerpoint_to_pdf', 'powerpoint-to-pdf', 'ppt_to_pdf', 'ppt-to-pdf']:
            file = files[0]
            if not file.filename.lower().endswith(('.ppt', '.pptx')):
                return jsonify({'error': 'Please upload a PowerPoint (.ppt or .pptx) file.'}), 400
            
            input_filename = secure_filename(file.filename)
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
            file.save(input_path)
            
            prs = Presentation(input_path)
            pdf_doc = fitz.open()
            
            for idx, slide in enumerate(prs.slides):
                page = pdf_doc.new_page()
                y_position = 50
                
                page.insert_text((50, y_position), f"--- Slide {idx + 1} ---", fontsize=12, color=(0.4, 0.4, 0.4))
                y_position += 30
                
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            text = paragraph.text.strip()
                            if text:
                                page.insert_text((50, y_position), text, fontsize=11, color=(0, 0, 0))
                                y_position += 20
                                if y_position > 750:
                                    page = pdf_doc.new_page()
                                    y_position = 50
            
            out_filename = input_filename.rsplit('.', 1)[0] + '.pdf'
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            pdf_doc.save(out_path)
            pdf_doc.close()
            
            return jsonify({'success': True, 'filename': out_filename})

        # 27-35. Additional Create PDF converters
        elif tool_type in ('html_to_pdf', 'text_to_pdf', 'images_to_pdf', 'epub_to_pdf',
                           'tiff_to_pdf', 'svg_to_pdf', 'markdown_to_pdf', 'csv_to_pdf', 'webpage_to_pdf'):
            allowed = {
                'html_to_pdf': ('.html', '.htm'),
                'text_to_pdf': ('.txt',),
                'images_to_pdf': ('.jpg', '.jpeg', '.png', '.webp', '.bmp', '.tif', '.tiff'),
                'epub_to_pdf': ('.epub',),
                'tiff_to_pdf': ('.tif', '.tiff'),
                'svg_to_pdf': ('.svg',),
                'markdown_to_pdf': ('.md', '.markdown'),
                'csv_to_pdf': ('.csv',),
            }
            if tool_type != 'webpage_to_pdf':
                if not files or not files[0].filename:
                    return jsonify({'error': 'Please upload a file.'}), 400
                if any(not f.filename or not f.filename.lower().endswith(allowed[tool_type]) for f in files):
                    return jsonify({'error': 'Unsupported file format for this tool.'}), 400
            out_filename = f"filevellum_{tool_type}_{uuid4().hex[:12]}.pdf"
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            if tool_type == 'webpage_to_pdf':
                # Fetch HTML only; external resources are intentionally not fetched.
                import requests
                target = safe_public_url(request.form.get('webpage_url', '').strip())
                with requests.get(target, timeout=12, allow_redirects=False,
                                  headers={'User-Agent': 'FileVellum/1.0'}, stream=True) as resp:
                    if resp.status_code in (301, 302, 303, 307, 308):
                        return jsonify({'error': 'Redirecting URLs are not supported.'}), 400
                    resp.raise_for_status()
                    if 'text/html' not in resp.headers.get('Content-Type', '').lower():
                        return jsonify({'error': 'URL must point to an HTML page.'}), 400
                    content = resp.raw.read(2_000_001)
                    if len(content) > 2_000_000:
                        return jsonify({'error': 'Webpage HTML is too large.'}), 400
                    page_html = content.decode(resp.encoding or 'utf-8', errors='replace')
                HTML(string=page_html, url_fetcher=lambda url: (_ for _ in ()).throw(ValueError('External resources disabled'))).write_pdf(out_path)
            elif tool_type == 'html_to_pdf':
                page_html = files[0].read().decode('utf-8-sig', errors='replace')
                HTML(string=page_html, url_fetcher=lambda url: (_ for _ in ()).throw(ValueError('External resources disabled'))).write_pdf(out_path)
            elif tool_type == 'svg_to_pdf':
                cairosvg.svg2pdf(bytestring=files[0].read(), write_to=out_path)
            elif tool_type in ('images_to_pdf', 'tiff_to_pdf'):
                images = []
                for file in files:
                    with Image.open(file.stream) as img:
                        frames = getattr(img, 'n_frames', 1)
                        for idx in range(frames):
                            img.seek(idx)
                            frame = img.convert('RGBA')
                            background = Image.new('RGB', frame.size, 'white')
                            background.paste(frame, mask=frame.getchannel('A'))
                            images.append(background)
                if not images:
                    return jsonify({'error': 'No valid images found.'}), 400
                images[0].save(out_path, 'PDF', save_all=True, append_images=images[1:], resolution=100.0)
            elif tool_type == 'epub_to_pdf':
                with TemporaryDirectory() as tmp:
                    epub_path = os.path.join(tmp, 'book.epub')
                    files[0].save(epub_path)
                    book = epub.read_epub(epub_path)
                    chapters = []
                    for item in book.get_items():
                        if item.get_type() == ITEM_DOCUMENT:
                            soup = BeautifulSoup(item.get_content(), 'html.parser')
                            for tag in soup(['script', 'iframe', 'object']): tag.decompose()
                            chapters.append(str(soup.body or soup))
                    if not chapters:
                        return jsonify({'error': 'No readable EPUB chapters found.'}), 400
                    HTML(string='<html><body>' + '<div style="break-after:page">'.join(chapters) + '</body></html>',
                         url_fetcher=lambda url: (_ for _ in ()).throw(ValueError('External resources disabled'))).write_pdf(out_path)
            else:
                content = files[0].read().decode('utf-8-sig', errors='replace')
                if tool_type == 'markdown_to_pdf':
                    body = markdown_to_html(content, extensions=['tables', 'fenced_code'])
                elif tool_type == 'csv_to_pdf':
                    from io import StringIO
                    reader = csv.reader(StringIO(content))
                    rows = []
                    for idx, row in enumerate(reader):
                        if idx >= 5000: break
                        rows.append('<tr>' + ''.join('<td>' + html.escape(cell) + '</td>' for cell in row[:50]) + '</tr>')
                    body = '<table border="1" cellspacing="0" cellpadding="5">' + ''.join(rows) + '</table>'
                else:
                    body = '<pre style="white-space:pre-wrap;overflow-wrap:anywhere">' + html.escape(content) + '</pre>'
                HTML(string='<html><head><meta charset="utf-8"><style>body{font-family:sans-serif}table{border-collapse:collapse;font-size:10px}td{max-width:160px;overflow-wrap:anywhere}</style></head><body>' + body + '</body></html>',
                     url_fetcher=lambda url: (_ for _ in ()).throw(ValueError('External resources disabled'))).write_pdf(out_path)
            return jsonify({'success': True, 'filename': out_filename})

        # Convert PDF: additional conversion formats
        elif tool_type in CONVERT_PDF_EXTRA:
            return process_convert_pdf_extra(tool_type, files)

        else:
            return jsonify({'error': 'Invalid tool or unsupported operation!'}), 400

    except Exception as e:
        app.logger.exception('FileVellum processing error')
        return jsonify({'error': 'Processing failed. Please check the file or server logs.'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    # Only allow a single safe filename located directly inside the uploads folder.
    safe_name = secure_filename(filename)
    if not filename or safe_name != filename or filename in {'.', '..'}:
        return "Invalid download filename.", 400

    upload_root = os.path.realpath(app.config['UPLOAD_FOLDER'])
    file_path = os.path.realpath(os.path.join(upload_root, safe_name))
    if os.path.commonpath([upload_root, file_path]) != upload_root or not os.path.isfile(file_path):
        return "File not found.", 404

    return send_file(file_path, as_attachment=True, download_name=safe_name)

@app.route('/robots.txt')
def robots():
    content = "User-agent: *\nAllow: /\nSitemap: https://filevellum.onrender.com/sitemap.xml"
    return Response(content, mimetype='text/plain')

@app.route('/sitemap.xml')
def sitemap():
    base_url = "https://filevellum.onrender.com"
    urls = [f"{base_url}/", f"{base_url}/about-founders"]

    # 1. Static / Direct GET Routes
    excluded = {'home', 'about_founders', 'robots', 'sitemap', 'static'}
    for rule in app.url_map.iter_rules():
        if 'GET' in rule.methods and not rule.arguments and rule.endpoint not in excluded:
            urls.append(f"{base_url}{rule.rule}")

    # 2. All Dynamic Catalog Tools
    catalogs = [
        ('/archive-tool/', ARCHIVE_TOOLS.keys(), False),
        ('/image-tool/', IMAGE_EDIT_TOOLS.keys(), False),
        ('/image-convert/', IMAGE_CONVERSION_TOOLS.keys(), False),
        ('/document-convert/', DOCUMENT_CONVERSION_TOOLS.keys(), True),
        ('/file-utility/', FILE_UTILITY_TOOLS.keys(), True),
        ('/text-utilities/', TEXT_UTILITY_TOOLS.keys(), False),
        ('/developer-tools/', DEVELOPER_TOOLS.keys(), False),
        ('/qr-barcode/', QR_BARCODE_TOOLS.keys(), True),
        ('/extra-utilities/', EXTRA_UTILITY_TOOLS.keys(), False),
        ('/image-inspection/', IMAGE_INSPECTION_TOOLS.keys(), False),
        ('/edit-secure/', EDIT_SECURE_TOOLS.keys(), True),
    ]

    for prefix, keys, to_hyphen in catalogs:
        for key in keys:
            slug = key.replace('_', '-') if to_hyphen else key
            urls.append(f"{base_url}{prefix}{slug}")

    # Remove duplicates while preserving order
    unique_urls = list(dict.fromkeys(urls))

    xml_items = []
    for u in unique_urls:
        priority = "1.0" if u == f"{base_url}/" else "0.8"
        xml_items.append(f"    <url>\n        <loc>{u}</loc>\n        <priority>{priority}</priority>\n    </url>")

    xml = '<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n' + '\n'.join(xml_items) + '\n</urlset>'
    return Response(xml, mimetype='application/xml')

if __name__ == '__main__':
    # Debug mode is opt-in; keep it disabled by default for safer local/public runs.
    debug_mode = os.environ.get('FILEVELLUM_DEBUG', '0').strip().lower() in {'1', 'true', 'yes'}
    app.run(debug=debug_mode, use_reloader=False)