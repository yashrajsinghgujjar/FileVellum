"""FileVellum Edit & Secure PDF additional converters."""
import io, os, re, zipfile, json
from uuid import uuid4
import pymupdf as fitz
from flask import request, jsonify
from werkzeug.utils import secure_filename

TOOLS = {
 'edit_pdf_metadata':'Edit PDF Metadata','remove_pdf_metadata':'Remove PDF Metadata',
 'add_pdf_header_footer':'Add PDF Header & Footer','add_text_to_pdf':'Add Text to PDF',
 'add_image_to_pdf':'Add Image to PDF','redact_pdf':'Redact PDF',
 'pdf_form_filler':'PDF Form Filler','pdf_form_creator':'PDF Form Creator',
 'flatten_pdf':'Flatten PDF','compare_pdf_files':'Compare PDF Files',
 'pdf_repair':'PDF Repair','pdf_password_checker':'PDF Password Checker',
 'pdf_permissions_viewer':'PDF Permissions Viewer',
 'pdf_attachment_extractor':'PDF Attachment Extractor',
}

def process(tool, files, upload_folder):
    if not files or not files[0].filename.lower().endswith('.pdf'):
        return jsonify(error='Upload a PDF file.'),400
    raw=files[0].read()
    if not raw.startswith(b'%PDF-') or len(raw)>100*1024*1024:
        return jsonify(error='Invalid PDF or file larger than 100 MB.'),400
    def opt(k,default=''):return request.form.get(k,default).strip()
    def limit(s,n=500):return s[:n]
    try:
        doc=fitz.open(stream=raw,filetype='pdf')
        if doc.needs_pass:
            pw=opt('password')
            if tool=='pdf_password_checker':
                return _report({'password_protected':True,'pages':'Password required'},upload_folder,tool)
            if not pw or not doc.authenticate(pw):
                return jsonify(error='This PDF is encrypted. Enter its correct password.'),400
        if tool=='pdf_password_checker':return _report({'password_protected':False,'pages':doc.page_count},upload_folder,tool)
        if tool=='pdf_permissions_viewer':
            return _report({'encrypted':bool(doc.is_encrypted),'permissions_integer':doc.permissions,'page_count':doc.page_count,'metadata':doc.metadata},upload_folder,tool)
        if not doc.page_count:return jsonify(error='PDF has no pages.'),400
        if doc.page_count>300:return jsonify(error='Maximum 300 pages.'),400
        if tool=='edit_pdf_metadata':
            meta={k:doc.metadata.get(k,'') or '' for k in ('title','author','subject','keywords','creator')}
            for k in meta:
                if opt(k):meta[k]=limit(opt(k),250)
            doc.set_metadata(meta)
        elif tool=='remove_pdf_metadata':doc.set_metadata({})
        elif tool=='add_pdf_header_footer':
            header=limit(opt('header'),150);footer=limit(opt('footer'),150)
            for i,p in enumerate(doc):
                if header:p.insert_text((40,28),header.replace('{page}',str(i+1)),fontsize=10)
                if footer:p.insert_text((40,p.rect.height-25),footer.replace('{page}',str(i+1)),fontsize=10)
        elif tool=='add_text_to_pdf':
            text=limit(opt('text'),3000)
            if not text:return jsonify(error='Enter text to add.'),400
            page=int(opt('page','1'))-1
            if not 0<=page<len(doc):return jsonify(error='Invalid page number.'),400
            p=doc[page];x=float(opt('x','40'));y=float(opt('y','60'))
            p.insert_textbox(fitz.Rect(x,y,min(x+500,p.rect.width),min(y+400,p.rect.height)),text,fontsize=12)
        elif tool=='add_image_to_pdf':
            img=request.files.get('overlay_image')
            if not img or not img.filename.lower().endswith(('.png','.jpg','.jpeg')):return jsonify(error='Upload JPG or PNG overlay image.'),400
            data=img.read()
            if len(data)>15*1024*1024:return jsonify(error='Image maximum 15 MB.'),400
            p=doc[0];p.insert_image(fitz.Rect(40,40,min(240,p.rect.width),min(240,p.rect.height)),stream=data,keep_proportion=True)
        elif tool=='redact_pdf':
            target=limit(opt('text'),200)
            if not target:return jsonify(error='Enter the exact text to redact.'),400
            matches=0
            for p in doc:
                for rect in p.search_for(target):p.add_redact_annot(rect,fill=(0,0,0));matches+=1
                p.apply_redactions()
            if not matches:return jsonify(error='Text not found. Scanned PDFs need OCR first.'),400
        elif tool=='pdf_form_filler':
            field=opt('field_name');value=limit(opt('field_value'),500);found=0
            for p in doc:
                for widget in list(p.widgets() or []):
                    if widget.field_name==field:
                        widget.field_value=value;widget.update();found+=1
            if not found:return jsonify(error='Field not found. Enter an existing PDF form field name.'),400
        elif tool=='pdf_form_creator':
            field=re.sub(r'[^a-zA-Z0-9_\-]','_',opt('field_name','Your_Name'))[:80]
            page=doc[0];w=fitz.Widget();w.field_name=field;w.field_type=fitz.PDF_WIDGET_TYPE_TEXT
            w.rect=fitz.Rect(50,80,300,110);w.text_fontsize=12;w.border_width=1;page.add_widget(w)
        elif tool=='flatten_pdf':
            # Render pages to fixed image-only pages (annotations and forms included).
            flat=fitz.open()
            for p in doc:
                pix=p.get_pixmap(matrix=fitz.Matrix(1.5,1.5),annots=True)
                outp=flat.new_page(width=p.rect.width,height=p.rect.height)
                outp.insert_image(outp.rect,stream=pix.tobytes('png'))
            doc.close();doc=flat
        elif tool=='compare_pdf_files':
            second=request.files.get('second_pdf')
            if not second or not second.filename.lower().endswith('.pdf'):return jsonify(error='Upload second PDF.'),400
            other=fitz.open(stream=second.read(),filetype='pdf')
            if other.needs_pass:return jsonify(error='Second PDF is encrypted.'),400
            diffs=[]
            for i in range(max(len(doc),len(other))):
                a=doc[i].get_text(sort=True) if i<len(doc) else ''
                b=other[i].get_text(sort=True) if i<len(other) else ''
                if a!=b:diffs.append({'page':i+1,'first_excerpt':a[:700],'second_excerpt':b[:700]})
            return _report({'comparison':'Extracted text only (not visual layout)','different_pages':diffs,'first_pages':len(doc),'second_pages':len(other)},upload_folder,tool)
        elif tool=='pdf_repair':
            # PyMuPDF can reconstruct readable PDF structure when opening succeeds.
            pass
        elif tool=='pdf_attachment_extractor':
            if doc.embfile_count()==0:return jsonify(error='This PDF has no embedded attachments.'),400
            name=f'{uuid4().hex}_attachments.zip';path=os.path.join(upload_folder,name)
            with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
                for i in range(min(doc.embfile_count(),100)):
                    info=doc.embfile_info(i);original=secure_filename(info.get('filename') or info.get('name') or f'attachment_{i}')
                    z.writestr(f'{i+1}_{original}',doc.embfile_get(i))
            return jsonify(success=True,filename=name)
        else:return jsonify(error='Unsupported tool.'),400
        name=f'{uuid4().hex}_{tool}.pdf';doc.save(os.path.join(upload_folder,name),garbage=4,deflate=True,encryption=fitz.PDF_ENCRYPT_NONE)
        return jsonify(success=True,filename=name)
    except (ValueError,TypeError,RuntimeError,IndexError) as e:
        return jsonify(error=f'Unable to process PDF: {str(e)[:160]}'),400

def _report(data,folder,tool):
    name=f'{uuid4().hex}_{tool}.json'
    with open(os.path.join(folder,name),'w',encoding='utf8') as f:json.dump(data,f,indent=2,ensure_ascii=False)
    return jsonify(success=True,filename=name)
