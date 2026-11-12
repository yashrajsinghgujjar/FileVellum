import os, io, json, zipfile
from uuid import uuid4
from PIL import Image, ImageFilter, ImageEnhance, ImageOps, ImageDraw
from werkzeug.utils import secure_filename

TOOLS = {
    'remove_bg','edit_image','image_compressor','image_resizer','image_cropper',
    'image_rotator','image_flipper','image_watermark','passport_photo_maker',
    'blur_image','sharpen_image','image_color_filters','image_to_text_ocr',
    'image_metadata_viewer','remove_image_metadata','image_background_color_changer',
    'image_border_maker','image_collage_maker','image_upscaler',
    'image_color_palette_extractor'
}
NAMES = {
'remove_bg':'Background Remover','edit_image':'Image Editor','image_compressor':'Image Compressor',
'image_resizer':'Image Resizer','image_cropper':'Image Cropper','image_rotator':'Image Rotator',
'image_flipper':'Image Flipper','image_watermark':'Image Watermark','passport_photo_maker':'Passport Photo Maker',
'blur_image':'Blur Image','sharpen_image':'Sharpen Image','image_color_filters':'Image Color Filters',
'image_to_text_ocr':'Image to Text OCR','image_metadata_viewer':'Image Metadata Viewer',
'remove_image_metadata':'Remove Image Metadata','image_background_color_changer':'Image Background Color Changer',
'image_border_maker':'Image Border Maker','image_collage_maker':'Image Collage Maker',
'image_upscaler':'Image Upscaler','image_color_palette_extractor':'Image Color Palette Extractor'
}
def _get_image(file):
    if not file or not file.filename: raise ValueError('Please select an image first.')
    try:
        im=Image.open(file.stream)
        im.load()
        return im
    except Exception as e: raise ValueError('Invalid or unsupported image file.') from e
def _save(im, folder, base, fmt=None, quality=90):
    fmt=(fmt or im.format or 'PNG').upper()
    if fmt not in ('JPEG','JPG','PNG','WEBP','BMP','TIFF'): fmt='PNG'
    if fmt=='JPG': fmt='JPEG'
    ext={'JPEG':'jpg','PNG':'png','WEBP':'webp','BMP':'bmp','TIFF':'tiff'}[fmt]
    name=f"{uuid4().hex[:10]}_{secure_filename(os.path.splitext(base)[0]) or 'image'}.{ext}"
    path=os.path.join(folder,name)
    if fmt=='JPEG' and im.mode in ('RGBA','LA'):
        bg=Image.new('RGB',im.size,'white'); bg.paste(im,mask=im.getchannel('A')); im=bg
    elif fmt=='JPEG': im=im.convert('RGB')
    im.save(path,format=fmt,quality=quality,optimize=(fmt in ('JPEG','PNG','WEBP')))
    return name
def process_image_tool(tool, files, form, folder):
    try:
        if tool=='image_collage_maker':
            valid=[f for f in files if f and f.filename]
            if len(valid)<2: raise ValueError('Please upload at least two images for a collage.')
            ims=[]
            for f in valid:
                im=_get_image(f).convert('RGB'); im.thumbnail((600,600)); ims.append(im)
            cols=max(1,int(len(ims)**0.5+0.999)); rows=(len(ims)+cols-1)//cols
            cellw=max(i.width for i in ims); cellh=max(i.height for i in ims)
            canvas=Image.new('RGB',(cols*cellw+(cols+1)*12,rows*cellh+(rows+1)*12),'white')
            for n,im in enumerate(ims):
                x=12+(n%cols)* (cellw+12); y=12+(n//cols)*(cellh+12)
                canvas.paste(im,(x,y))
            return {'success':True,'filename':_save(canvas,folder,'image_collage.jpg','JPEG',92)}
        f=files[0] if files else None
        im=_get_image(f); original=f.filename
        fmt=im.format or 'PNG'
        if tool=='image_metadata_viewer':
            info={'filename':original,'format':fmt,'width':im.width,'height':im.height,'mode':im.mode,
                  'aspect_ratio':round(im.width/im.height,4) if im.height else None,
                  'dpi':im.info.get('dpi'),'metadata_keys':list(im.info.keys())}
            name=f"{uuid4().hex[:10]}_image_metadata.json"
            with open(os.path.join(folder,name),'w',encoding='utf-8') as w: json.dump(info,w,indent=2,ensure_ascii=False)
            return {'success':True,'filename':name}
        if tool=='image_to_text_ocr':
            try:
                import pytesseract
                lang=form.get('ocr_language','eng') or 'eng'
                content=pytesseract.image_to_string(im,lang=lang)
            except Exception as e: raise RuntimeError('OCR requires Tesseract OCR and pytesseract. Install both and configure Tesseract PATH.') from e
            name=f"{uuid4().hex[:10]}_image_text.txt"
            with open(os.path.join(folder,name),'w',encoding='utf-8') as w: w.write(content)
            return {'success':True,'filename':name}
        if tool=='remove_bg':
            try:
                from rembg import remove, new_session
                result=remove(im.convert('RGBA'),session=new_session('u2net'))
                return {'success':True,'filename':_save(result,folder,original,'PNG')}
            except Exception as e: raise RuntimeError('Background Remover needs rembg. Install with: pip install rembg') from e
        if tool=='remove_image_metadata':
            clean=Image.new(im.mode,im.size); clean.putdata(list(im.getdata())); im=clean
        elif tool in ('edit_image','image_cropper'):
            x=max(0,int(float(form.get('crop_x',0) or 0))); y=max(0,int(float(form.get('crop_y',0) or 0)))
            w=int(float(form.get('crop_w',0) or 0)); h=int(float(form.get('crop_h',0) or 0))
            if w>0 and h>0: im=im.crop((x,y,min(x+w,im.width),min(y+h,im.height)))
            angle=int(form.get('rotate',0) or 0)%360
            if angle: im=im.rotate(-angle,expand=True)
            rw=int(form.get('width',0) or 0); rh=int(form.get('height',0) or 0)
            if rw>0 and rh>0: im=im.resize((min(rw,12000),min(rh,12000)),Image.Resampling.LANCZOS)
        elif tool=='image_resizer':
            w=int(form.get('width',im.width) or im.width); h=int(form.get('height',im.height) or im.height)
            if w<1 or h<1: raise ValueError('Width and height must be greater than zero.')
            im=im.resize((min(w,12000),min(h,12000)),Image.Resampling.LANCZOS)
        elif tool=='image_rotator':
            im=im.rotate(-int(form.get('rotate',90) or 90),expand=True)
        elif tool=='image_flipper':
            im=ImageOps.mirror(im) if form.get('flip','horizontal')=='horizontal' else ImageOps.flip(im)
        elif tool=='image_watermark':
            im=im.convert('RGBA'); layer=Image.new('RGBA',im.size,(0,0,0,0)); d=ImageDraw.Draw(layer)
            txt=(form.get('watermark_text') or 'FileVellum').strip()[:100]
            d.text((max(8,im.width//30),max(8,im.height- max(35,im.height//12))),txt,fill=(255,255,255,170),stroke_width=1,stroke_fill=(0,0,0,100))
            im=Image.alpha_composite(im,layer)
        elif tool=='passport_photo_maker':
            im=ImageOps.fit(im.convert('RGB'),(413,531),method=Image.Resampling.LANCZOS,centering=(0.5,0.35))
        elif tool=='blur_image':
            im=im.filter(ImageFilter.GaussianBlur(radius=max(1,min(20,float(form.get('strength',3) or 3)))))
        elif tool=='sharpen_image':
            im=im.filter(ImageFilter.UnsharpMask(radius=2,percent=max(100,min(500,int(form.get('strength',180) or 180))),threshold=3))
        elif tool=='image_color_filters':
            mode=form.get('filter','grayscale')
            if mode=='grayscale': im=ImageOps.grayscale(im)
            elif mode=='sepia':
                g=ImageOps.grayscale(im).convert('RGB'); im=ImageOps.colorize(ImageOps.grayscale(g),'#34200f','#e8c58a')
            elif mode=='invert': im=ImageOps.invert(im.convert('RGB'))
            elif mode=='bright': im=ImageEnhance.Brightness(im).enhance(1.25)
            elif mode=='contrast': im=ImageEnhance.Contrast(im).enhance(1.4)
        elif tool=='image_background_color_changer':
            color=(form.get('background_color') or '#ffffff').strip()
            if not (len(color)==7 and color[0]=='#'): color='#ffffff'
            rgba=im.convert('RGBA'); bg=Image.new('RGBA',rgba.size,color)
            if 'A' in rgba.getbands(): bg.alpha_composite(rgba)
            else: bg=rgba
            im=bg
        elif tool=='image_border_maker':
            px=max(1,min(300,int(form.get('border',20) or 20)))
            color=form.get('border_color') or '#ffffff'
            if len(color)!=7 or not color.startswith('#'): color='#ffffff'
            im=ImageOps.expand(im,border=px,fill=color)
        elif tool=='image_upscaler':
            scale=max(2,min(4,int(form.get('scale',2) or 2)))
            im=im.resize((min(im.width*scale,12000),min(im.height*scale,12000)),Image.Resampling.LANCZOS)
            im=im.filter(ImageFilter.UnsharpMask(radius=1.5,percent=120,threshold=3))
        elif tool=='image_color_palette_extractor':
            colors=im.convert('RGB').quantize(colors=8).convert('RGB').getcolors(im.width*im.height)
            colors=sorted(colors or [],reverse=True)[:8]
            report={'filename':original,'dominant_colors':[{'hex':'#%02x%02x%02x'%c,'pixels':n} for n,c in colors]}
            name=f"{uuid4().hex[:10]}_color_palette.json"
            with open(os.path.join(folder,name),'w') as w: json.dump(report,w,indent=2)
            return {'success':True,'filename':name}
        quality=max(20,min(95,int(form.get('quality',82) or 82)))
        outfmt=form.get('output_format') or ('PNG' if tool in ('remove_bg','remove_image_metadata') else fmt)
        return {'success':True,'filename':_save(im,folder,original,outfmt,quality)}
    except (ValueError, RuntimeError):
        raise
    except Exception as e:
        raise RuntimeError('Image processing failed. Check the image and options, then try again.') from e
