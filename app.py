"""PDFMint Studio - FastAPI application. Run: uvicorn app:app --reload"""
from __future__ import annotations

import json
import os
import re
import secrets
import shutil
import tempfile
import time
from pathlib import Path

import fitz
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask
from starlette.datastructures import UploadFile

from catalog import TOOLS, TOOL_BY_ID
from engine import process, ProcessingError, apply_editor_actions, open_pdf, external_binary

ROOT=Path(__file__).resolve().parent
APP_VERSION='2026.11.14.17'
MAX_TOTAL_MB=int(os.getenv('MAX_UPLOAD_MB','80'))
MAX_BYTES=MAX_TOTAL_MB*1024*1024
MAX_FILES=40
EDITOR_ROOT=Path(tempfile.gettempdir())/'pdfmint-editor-sessions'
EDITOR_MAX_AGE=2*60*60

app=FastAPI(title='PDFMint Studio',version='1.17.0',description='Free, self-hostable PDF tools')


def pdfmint_download_name(name:str, fallback:str='result')->str:
    """Return a safe user-facing download filename branded with PDFMint."""
    raw=Path(str(name or fallback).replace('\\','/').split('/')[-1]).name.strip() or fallback
    # Avoid stacking the brand if an upstream operation already used it.
    if re.match(r'(?i)^pdfmint(?:[_\-\s]|$)', raw):
        return raw
    return f'PDFMint_{raw}'

@app.middleware("http")
async def no_stale_cache(request:Request, call_next):
    response=await call_next(request)
    if request.url.path.startswith('/api/') or request.url.path in ('/app.js','/styles.css','/index.html','/'):
        response.headers['Cache-Control']='no-store, no-cache, must-revalidate, max-age=0'
        response.headers['Pragma']='no-cache'
        response.headers['Expires']='0'
    response.headers['X-PDFMint-Version']=APP_VERSION
    return response

@app.get('/api/tools')
async def get_tools():
    return {'tools':TOOLS, 'limits':{'total_upload_mb':MAX_TOTAL_MB,'max_files':MAX_FILES}}

@app.get('/api/health')
async def health():
    return {'status':'ok','version':APP_VERSION,'tools':len(TOOLS), 'capabilities':{
        'libreoffice':bool(external_binary('libreoffice')),
        'modern_office_fallback':True,
        'tesseract':bool(external_binary('tesseract')),
        'ghostscript':bool(external_binary('ghostscript')),
        'ai_configured':bool(os.getenv('OPENAI_API_KEY')),
    }}


def _cleanup_editor_sessions():
    try: EDITOR_ROOT.mkdir(parents=True,exist_ok=True)
    except Exception: return
    now=time.time()
    for child in EDITOR_ROOT.iterdir():
        try:
            if child.is_dir() and now-child.stat().st_mtime>EDITOR_MAX_AGE:
                shutil.rmtree(child,ignore_errors=True)
        except OSError: pass


def _editor_dir(editor_id:str)->Path:
    if not re.fullmatch(r'[A-Za-z0-9_-]{16,80}',editor_id or ''):
        raise HTTPException(404,'Editor session not found')
    target=EDITOR_ROOT/editor_id
    if not target.is_dir(): raise HTTPException(404,'Editor session expired. Upload the PDF again.')
    try: os.utime(target,None)
    except OSError: pass
    return target


def _span_color(value):
    try: value=int(value)&0xFFFFFF
    except Exception: value=0
    return f'#{value:06x}'


def _page_info(page):
    spans=[]
    try: blocks=page.get_text('dict',sort=True).get('blocks',[])
    except Exception: blocks=[]
    for block in blocks:
        for line in block.get('lines',[]):
            for span in line.get('spans',[]):
                text=str(span.get('text',''))
                bbox=span.get('bbox')
                if not text.strip() or not bbox or len(bbox)!=4: continue
                origin=span.get('origin') or (bbox[0], bbox[3])
                spans.append({'text':text[:1200],'bbox':[round(float(v),2) for v in bbox], 'origin':[round(float(v),2) for v in origin[:2]], 'size':round(float(span.get('size',11) or 11),2), 'font':str(span.get('font',''))[:100], 'color':_span_color(span.get('color',0))})
                if len(spans)>=1600: break
            if len(spans)>=1600: break
        if len(spans)>=1600: break
    r=page.rect
    return {'width':round(r.width,2),'height':round(r.height,2),'text':spans}


@app.post('/api/editor/prepare')
async def prepare_editor(request:Request):
    """Create a short-lived local editing session and return page/text metadata."""
    _cleanup_editor_sessions()
    tmp=None
    try:
        form=await request.form(max_files=1,max_fields=8)
        item=form.get('file')
        password=str(form.get('password') or '')
        if not isinstance(item,UploadFile): raise ProcessingError('Choose one PDF to edit.')
        name=Path((item.filename or 'document.pdf').replace('\\','/').split('/')[-1]).name
        if Path(name).suffix.lower()!='.pdf': raise ProcessingError('Edit PDF accepts a PDF file.')
        editor_id=secrets.token_urlsafe(18).replace('-','_')
        tmp=EDITOR_ROOT/editor_id;tmp.mkdir(parents=True,exist_ok=False)
        src=tmp/'source.pdf';size=0
        with src.open('wb') as stream:
            while True:
                chunk=await item.read(1024*1024)
                if not chunk: break
                size+=len(chunk)
                if size>MAX_BYTES: raise ProcessingError(f'PDF cannot exceed {MAX_TOTAL_MB} MB.')
                stream.write(chunk)
        if size==0 or b'%PDF-' not in src.open('rb').read(1024): raise ProcessingError('The selected file is not a valid PDF.')

        # Password-protected PDFs must be authenticated before any page/text data
        # is exposed. Once authenticated, save a decrypted temporary working copy
        # inside the short-lived editor session so later render/save requests do
        # not need to retain the user's password.
        try:
            probe=fitz.open(str(src))
        except Exception as exc:
            raise ProcessingError(f'Cannot read {name}: {exc}') from exc
        try:
            if probe.needs_pass:
                if not password:
                    probe.close()
                    shutil.rmtree(tmp,ignore_errors=True); tmp=None
                    return JSONResponse({'error':'This PDF is password-protected. Enter its password to edit it.','password_required':True},status_code=423)
                if probe.authenticate(password)<=0:
                    probe.close()
                    shutil.rmtree(tmp,ignore_errors=True); tmp=None
                    return JSONResponse({'error':'Incorrect PDF password. Please try again.','password_required':True,'invalid_password':True},status_code=401)
                unlocked=tmp/'source-unlocked.pdf'
                probe.save(str(unlocked),garbage=4,deflate=True,encryption=fitz.PDF_ENCRYPT_NONE)
                probe.close()
                src.unlink(missing_ok=True)
                unlocked.replace(src)
                probe=fitz.open(str(src))
            if not len(probe): raise ProcessingError('PDF has no pages.')
            if len(probe)>180: raise ProcessingError('Interactive editing is limited to 180 pages per file.')
            pages=[_page_info(pg) for pg in probe]
        finally:
            try: probe.close()
            except Exception: pass
        (tmp/'meta.json').write_text(json.dumps({'name':name,'pages':pages},ensure_ascii=False),encoding='utf-8')
        return {'editor_id':editor_id,'name':name,'pages':pages}
    except ProcessingError as exc:
        if tmp: shutil.rmtree(tmp,ignore_errors=True)
        return JSONResponse({'error':str(exc)},status_code=400)
    except Exception as exc:
        if tmp: shutil.rmtree(tmp,ignore_errors=True)
        return JSONResponse({'error':f'Could not prepare PDF editor: {str(exc)[:180]}'},status_code=500)


def _render_editor_page(work:Path,page_number:int,clean_text:bool=False)->bytes:
    cache=work/f"page-{page_number:04d}{'-clean' if clean_text else ''}.png"
    if cache.exists():
        return cache.read_bytes()
    src=work/'source.pdf'
    with open_pdf(src) as doc:
        if page_number<1 or page_number>len(doc): raise HTTPException(404,'Page not found')
        page=doc[page_number-1]
        if clean_text:
            # Remove only text glyphs from a temporary in-memory copy while keeping
            # vector graphics / table lines / images. The browser paints the
            # extracted text back as editable HTML boxes on top.
            try:
                for block in page.get_text('dict',sort=True).get('blocks',[]):
                    for line in block.get('lines',[]):
                        for span in line.get('spans',[]):
                            text=str(span.get('text',''))
                            bbox=span.get('bbox')
                            if text.strip() and bbox and len(bbox)==4:
                                page.add_redact_annot(fitz.Rect(*map(float,bbox)),fill=None)
                page.apply_redactions(images=0,graphics=0,text=0)
            except Exception:
                pass
        pix=page.get_pixmap(matrix=fitz.Matrix(1.65,1.65),alpha=False,colorspace=fitz.csRGB)
        data=pix.tobytes('png')
    try: cache.write_bytes(data)
    except OSError: pass
    return data


@app.get('/api/editor/{editor_id}/page/{page_number}.png')
async def editor_page(editor_id:str,page_number:int):
    work=_editor_dir(editor_id)
    try:
        data=_render_editor_page(work,page_number,False)
        return Response(content=data,media_type='image/png',headers={'Cache-Control':'private, max-age=1800','X-Content-Type-Options':'nosniff'})
    except HTTPException: raise
    except Exception as exc: raise HTTPException(500,f'Could not render page: {str(exc)[:160]}')


@app.get('/api/editor/{editor_id}/page/{page_number}/clean.png')
async def editor_clean_page(editor_id:str,page_number:int):
    work=_editor_dir(editor_id)
    try:
        data=_render_editor_page(work,page_number,True)
        return Response(content=data,media_type='image/png',headers={'Cache-Control':'private, max-age=1800','X-Content-Type-Options':'nosniff'})
    except HTTPException: raise
    except Exception as exc: raise HTTPException(500,f'Could not render editable page: {str(exc)[:160]}')


@app.post('/api/editor/{editor_id}/save')
async def save_editor(request:Request,editor_id:str):
    work=_editor_dir(editor_id);src=work/'source.pdf';out=work/'edited.pdf'
    try:
        body=await request.json()
        actions=body.get('actions',[]) if isinstance(body,dict) else []
        apply_editor_actions(src,actions,out)
        return FileResponse(str(out),filename=pdfmint_download_name('edited.pdf'),media_type='application/pdf',headers={'Cache-Control':'no-store','X-Content-Type-Options':'nosniff'})
    except ProcessingError as exc:
        return JSONResponse({'error':str(exc)},status_code=400)
    except Exception as exc:
        return JSONResponse({'error':f'Could not save edited PDF: {str(exc)[:180]}'},status_code=500)

@app.post('/api/process/{tool_id}')
async def handle_file(request:Request,tool_id:str):
    conf=TOOL_BY_ID.get(tool_id)
    if not conf: raise HTTPException(404,'Unknown PDF tool')
    tmp=Path(tempfile.mkdtemp(prefix='pdfmint-'))
    keep_temp=False
    try:
        # Avoid memory buffering: starlette's multipart UploadFile uses a spool file.
        form=await request.form(max_files=MAX_FILES,max_fields=30)
        upload_files=form.getlist('files')
        if not upload_files or any(not isinstance(f,UploadFile) for f in upload_files):
            raise ProcessingError('Select at least one document.')
        if len(upload_files)<conf['min_files']:
            raise ProcessingError(f"{conf['name']} needs at least {conf['min_files']} files.")
        if not conf['multiple'] and len(upload_files)!=1:
            raise ProcessingError('Choose only one file for this tool.')
        if len(upload_files)>MAX_FILES: raise ProcessingError('Too many files.')
        try: opts=json.loads(str(form.get('options','{}')))
        except (ValueError,TypeError): raise ProcessingError('Invalid tool settings.')
        if not isinstance(opts,dict): raise ProcessingError('Invalid tool settings.')
        paths=[];size=0
        for i,item in enumerate(upload_files):
            name=Path((item.filename or 'document').replace('\\','/').split('/')[-1]).name
            if not re.fullmatch(r'[\w.()\- ]{1,150}',name,re.UNICODE):
                name='upload'+Path(name).suffix.lower()[:12]
            ext=Path(name).suffix.lower()
            accept=conf['accept']
            allowed=['.jpg','.jpeg','.png','.webp','.bmp','.tif','.tiff','.heic'] if accept=='image/*' else [v.strip().lower() for v in accept.split(',')]
            if ext not in allowed:
                raise ProcessingError(f'{name}: expected {"image" if accept=="image/*" else accept} files.')
            # Preserve the original basename for batch-conversion output names while
            # isolating every upload in its own generated directory to avoid collisions.
            item_dir=tmp/f'input-{i+1:03d}'
            item_dir.mkdir(exist_ok=True)
            path=item_dir/name
            with path.open('wb') as stream:
                while True:
                    chunk=await item.read(1024*1024)
                    if not chunk: break
                    size+=len(chunk)
                    if size>MAX_BYTES:
                        raise ProcessingError(f'Total uploads cannot exceed {MAX_TOTAL_MB} MB.')
                    stream.write(chunk)
            if path.stat().st_size==0: raise ProcessingError(f'{name} is empty.')
            if ext=='.pdf' and b'%PDF-' not in path.open('rb').read(1024):
                raise ProcessingError(f'{name} does not appear to be a valid PDF.')
            paths.append(path)
        out=process(tool_id,paths,opts,tmp)
        if not out.exists() or not out.is_file() or out.stat().st_size==0:
            raise ProcessingError('Processing did not produce a file.')
        if out.stat().st_size>MAX_BYTES*2:
            raise ProcessingError('The output is too large to download in one request.')
        keep_temp=True
        return FileResponse(path=str(out),filename=pdfmint_download_name(out.name),media_type='application/octet-stream',background=BackgroundTask(shutil.rmtree,tmp,ignore_errors=True),headers={'Cache-Control':'no-store','X-Content-Type-Options':'nosniff'})
    except ProcessingError as exc:
        return JSONResponse({'error':str(exc)},status_code=400)
    except HTTPException:
        raise
    except Exception as exc:
        # User-facing error omits internal paths and stack traces.
        return JSONResponse({'error':f'Could not process this file: {str(exc)[:180]}'},status_code=500)
    finally:
        if not keep_temp: shutil.rmtree(tmp,ignore_errors=True)

app.mount('/',StaticFiles(directory=ROOT/'static',html=True),name='static')
