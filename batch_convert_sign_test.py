"""v16 regression: batch convert tools + custom Sign PDF placement."""
from pathlib import Path
import io, sys, tempfile, zipfile
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import fitz
from docx import Document
from openpyxl import Workbook
from pptx import Presentation
from fastapi.testclient import TestClient
from app import app
from engine import process

base=Path(tempfile.mkdtemp(prefix='pdfmint-v16-'))
pdfs=[]
for i in range(2):
    path=base/f'doc{i+1}.pdf'; doc=fitz.open(); page=doc.new_page(width=595,height=842); page.insert_text((70,80),f'Batch PDF {i+1}'); doc.save(path); doc.close(); pdfs.append(path)
words=[]; ppts=[]; sheets=[]; htmls=[]
for i in range(2):
    path=base/f'word{i+1}.docx'; doc=Document(); doc.add_paragraph(f'Word batch {i+1}'); doc.save(path); words.append(path)
    path=base/f'presentation{i+1}.pptx'; prs=Presentation(); prs.slides.add_slide(prs.slide_layouts[5]); prs.save(path); ppts.append(path)
    path=base/f'sheet{i+1}.xlsx'; wb=Workbook(); wb.active.append(['value',i+1]); wb.save(path); sheets.append(path)
    path=base/f'page{i+1}.html'; path.write_text(f'<h1>HTML {i+1}</h1>',encoding='utf-8'); htmls.append(path)

for tool,inputs in [
    ('word-to-pdf',words),('ppt-to-pdf',ppts),('excel-to-pdf',sheets),('html-to-pdf',htmls),
    ('pdf-to-jpg',pdfs),('pdf-to-word',pdfs),('pdf-to-ppt',pdfs),('pdf-to-excel',pdfs),('pdf-to-pdfa',pdfs),
]:
    work=base/f'work-{tool}'; work.mkdir(); out=process(tool,inputs,{},work)
    assert out.suffix=='.zip' and out.exists(),(tool,out)
    with zipfile.ZipFile(out) as archive:
        names=archive.namelist(); assert len(names)==2 and len(set(names))==2,(tool,names)
    print('PASS batch',tool,names)

work=base/'sign'; work.mkdir()
out=process('sign',[pdfs[0]],{'signature':'SALMAN','signature_style':'script','color':'#0a3d9e','size':44,'page':1,'x_pct':10,'y_pct':20,'w_pct':35,'h_pct':12},work)
with fitz.open(out) as doc:
    page=doc[0]; rects=[]
    for image in page.get_images(full=True): rects.extend(page.get_image_rects(image[0]))
    assert rects
    target=max(rects,key=lambda r:r.get_area()); allowed=fitz.Rect(59.5,168.4,267.75,269.44)
    assert target.x0>=allowed.x0-1 and target.y0>=allowed.y0-1 and target.x1<=allowed.x1+1 and target.y1<=allowed.y1+1,(target,allowed)
print('PASS custom sign placement',target)

client=TestClient(app)
src=pdfs[0].read_bytes()
response=client.post('/api/process/pdf-to-word',files=[('files',('a.pdf',src,'application/pdf')),('files',('b.pdf',src,'application/pdf'))],data={'options':'{}'})
assert response.status_code==200,response.text[:300]
with zipfile.ZipFile(io.BytesIO(response.content)) as archive:
    assert archive.namelist()==['PDFMint_a.docx','PDFMint_b.docx'],archive.namelist()
assert 'PDFMint_converted-files.zip' in (response.headers.get('content-disposition') or ''), response.headers.get('content-disposition')
print('PASS API batch PDFMint filenames')
