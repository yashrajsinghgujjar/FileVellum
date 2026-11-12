import os, io, json, zipfile
from uuid import uuid4
from PIL import Image
import pymupdf as fitz
from flask import jsonify

TOOLS = {
 'grayscale_pdf','ocr_pdf','make_pdf_searchable','optimize_pdf_images',
 'pdf_file_size_analyzer','pdf_page_size_checker','pdf_font_inspector',
 'pdf_color_converter','pdf_image_extractor'
}

def _save_bytes(data, folder, name):
    name = f"{uuid4().hex[:10]}_{name}"
    path = os.path.join(folder, name)
    with open(path, 'wb') as f: f.write(data)
    return name

def _pdf(files):
    if not files or not files[0] or not files[0].filename:
        raise ValueError('Please upload a PDF file.')
    f=files[0]
    if not f.filename.lower().endswith('.pdf'):
        raise ValueError('Please upload a valid PDF file.')
    data=f.read()
    if not data: raise ValueError('The uploaded PDF is empty.')
    try: doc=fitz.open(stream=data, filetype='pdf')
    except Exception as e: raise ValueError('This PDF is invalid, damaged, or password protected.') from e
    return data, doc, f.filename

def process_optimize_ocr(tool, files, form, upload_folder):
    try:
        data, doc, original = _pdf(files)
        base=os.path.splitext(os.path.basename(original))[0]
        if tool == 'pdf_file_size_analyzer':
            pages=[]
            for i,p in enumerate(doc):
                pages.append({'page':i+1,'width_pt':round(p.rect.width,2),'height_pt':round(p.rect.height,2),'images':len(p.get_images(full=True)),'text_chars':len(p.get_text())})
            report={'file':original,'size_bytes':len(data),'size_kb':round(len(data)/1024,2),'pages':doc.page_count,'metadata':doc.metadata,'page_details':pages}
            name=_save_bytes(json.dumps(report,indent=2,ensure_ascii=False).encode(),upload_folder,base+'_file_analysis.json')
            return jsonify(success=True,filename=name)
        if tool == 'pdf_page_size_checker':
            pages=[{'page':i+1,'width_pt':round(p.rect.width,2),'height_pt':round(p.rect.height,2),'width_mm':round(p.rect.width*25.4/72,1),'height_mm':round(p.rect.height*25.4/72,1),'orientation':'Landscape' if p.rect.width>p.rect.height else 'Portrait' if p.rect.height>p.rect.width else 'Square'} for i,p in enumerate(doc)]
            name=_save_bytes(json.dumps({'file':original,'pages':pages},indent=2).encode(),upload_folder,base+'_page_sizes.json')
            return jsonify(success=True,filename=name)
        if tool == 'pdf_font_inspector':
            fonts=[]
            for i,p in enumerate(doc):
                seen=set()
                for f in p.get_fonts(full=True):
                    key=(f[0],f[3],f[4])
                    if key not in seen:
                        seen.add(key); fonts.append({'page':i+1,'xref':f[0],'name':f[3],'type':f[4],'encoding':f[5]})
            name=_save_bytes(json.dumps({'file':original,'fonts':fonts},indent=2).encode(),upload_folder,base+'_font_report.json')
            return jsonify(success=True,filename=name)
        if tool == 'pdf_image_extractor':
            name=f"{uuid4().hex[:10]}_{base}_images.zip"; path=os.path.join(upload_folder,name)
            count=0
            with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
                for pi,page in enumerate(doc):
                    for ii,img in enumerate(page.get_images(full=True),1):
                        xref=img[0]
                        try:
                            info=doc.extract_image(xref); ext=info.get('ext','bin')
                            z.writestr(f'page_{pi+1}/image_{ii}.{ext}',info['image']); count+=1
                        except Exception: pass
                z.writestr('README.txt',f'Extracted {count} image(s) from {original}.')
            return jsonify(success=True,filename=name)
        if tool == 'grayscale_pdf' or tool == 'pdf_color_converter':
            mode=form.get('color_mode','grayscale').lower()
            if tool=='grayscale_pdf': mode='grayscale'
            if mode not in ('grayscale','rgb'): mode='grayscale'
            out=fitz.open()
            dpi=max(72,min(200,int(form.get('dpi',150) or 150)))
            for page in doc:
                pix=page.get_pixmap(matrix=fitz.Matrix(dpi/72,dpi/72),colorspace=fitz.csGRAY if mode=='grayscale' else fitz.csRGB,alpha=False)
                img=fitz.open('png',pix.tobytes('png'))
                rect=page.rect
                p=out.new_page(width=rect.width,height=rect.height)
                p.insert_image(p.rect,stream=img.convert_to_pdf())
            payload=out.tobytes(garbage=4,deflate=True)
            name=_save_bytes(payload,upload_folder,base+('_grayscale.pdf' if mode=='grayscale' else '_rgb.pdf'))
            out.close(); doc.close(); return jsonify(success=True,filename=name)
        if tool == 'optimize_pdf_images':
            quality=max(35,min(95,int(form.get('quality',65) or 65)))
            max_dpi=max(72,min(300,int(form.get('dpi',150) or 150)))
            replaced=set()
            for page in doc:
                for image in page.get_images(full=True):
                    xref=image[0]
                    if xref in replaced: continue
                    try:
                        info=doc.extract_image(xref)
                        im=Image.open(io.BytesIO(info['image'])).convert('RGB')
                        # Downsample only when image resolution materially exceeds requested target.
                        im.thumbnail((max_dpi*20,max_dpi*20),Image.Resampling.LANCZOS)
                        buff=io.BytesIO(); im.save(buff,format='JPEG',quality=quality,optimize=True)
                        page.replace_image(xref,stream=buff.getvalue()); replaced.add(xref)
                    except Exception: continue
            payload=doc.tobytes(garbage=4,deflate=True,clean=True)
            name=_save_bytes(payload,upload_folder,base+'_optimized.pdf')
            doc.close(); return jsonify(success=True,filename=name)
        if tool in ('ocr_pdf','make_pdf_searchable'):
            try:
                import pytesseract
                from pytesseract import TesseractNotFoundError
            except ImportError as e:
                raise RuntimeError('OCR dependency missing. Install pytesseract and Tesseract OCR.') from e
            lang=(form.get('ocr_language') or 'eng').strip()
            dpi=max(150,min(300,int(form.get('dpi',200) or 200)))
            output=fitz.open()
            for page in doc:
                # Keep existing text for searchable-PDF mode; OCR only pages with no meaningful text.
                if tool=='make_pdf_searchable' and len(page.get_text().strip())>20:
                    output.insert_pdf(doc,from_page=page.number,to_page=page.number)
                    continue
                pix=page.get_pixmap(matrix=fitz.Matrix(dpi/72,dpi/72),alpha=False)
                image=Image.open(io.BytesIO(pix.tobytes('png'))).convert('RGB')
                try:
                    searchable=pytesseract.image_to_pdf_or_hocr(image,extension='pdf',lang=lang)
                except TesseractNotFoundError as e:
                    raise RuntimeError('Tesseract OCR is not installed or not on PATH. Install Tesseract OCR, then restart Flask.') from e
                one=fitz.open('pdf',searchable); output.insert_pdf(one); one.close()
            name=_save_bytes(output.tobytes(garbage=4,deflate=True),upload_folder,base+'_searchable.pdf')
            output.close(); doc.close(); return jsonify(success=True,filename=name)
        raise ValueError('Unknown Optimize & OCR PDF tool.')
    except ValueError as e:
        return jsonify(error=str(e)),400
    except RuntimeError as e:
        return jsonify(error=str(e)),500
    except Exception:
        return jsonify(error='PDF processing failed. Please verify the PDF and try again.'),500
