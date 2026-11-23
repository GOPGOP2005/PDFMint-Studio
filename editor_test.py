"""Interactive editor and built-in Office fallback tests."""
from pathlib import Path
import base64, io, tempfile, sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))

import fitz
from fastapi.testclient import TestClient
from docx import Document
from openpyxl import Workbook
from pptx import Presentation
from pptx.util import Inches
from PIL import Image

from app import app
from engine import _office_fallback

client=TestClient(app)

def sample_pdf():
    d=fitz.open(); p=d.new_page(width=595,height=842)
    p.insert_text((72,100),'Old Name',fontsize=18,color=(0,0,0))
    p.insert_text((72,140),'Invoice 123',fontsize=12,color=(0,0,0))
    raw=d.tobytes();d.close();return raw

src=sample_pdf()
prep=client.post('/api/editor/prepare',files={'file':('edit.pdf',src,'application/pdf')})
assert prep.status_code==200,prep.text[:300]
data=prep.json(); assert data['pages'] and data['pages'][0]['text']
editor_id=data['editor_id']
sp=next(s for s in data['pages'][0]['text'] if 'Old Name' in s['text']); assert sp.get('origin') and sp.get('font')
page=client.get(f'/api/editor/{editor_id}/page/1.png')
assert page.status_code==200 and page.content.startswith(b'\x89PNG')
# 1x1 red PNG for image insertion
bio=io.BytesIO();Image.new('RGB',(20,20),(220,40,40)).save(bio,'PNG')
img='data:image/png;base64,'+base64.b64encode(bio.getvalue()).decode()
actions=[
    {'type':'replace_text','page':0,'rect':sp['bbox'],'origin':sp['origin'],'font':sp['font'],'text':'New Name','size':18,'color':'#111111'},
    {'type':'add_text','page':0,'rect':[72,170,260,205],'text':'Added note','size':12,'color':'#e74755'},
    {'type':'highlight','page':0,'rect':[68,125,180,150]},
    {'type':'rectangle','page':0,'rect':[65,215,260,260],'color':'#3366cc','width':2},
    {'type':'image','page':0,'rect':[400,60,450,110],'data':img},
]
saved=client.post(f'/api/editor/{editor_id}/save',json={'actions':actions})
assert saved.status_code==200,saved.text[:300]
assert saved.content.startswith(b'%PDF-')
with fitz.open(stream=saved.content,filetype='pdf') as d:
    text=d[0].get_text()
    assert 'New Name' in text,text
    assert 'Added note' in text,text
    assert 'Old Name' not in text,text
    assert len(d[0].get_images())>=1
print('PASS interactive editor replace/add/highlight/shape/image')

with tempfile.TemporaryDirectory() as td:
    td=Path(td)
    doc=Document();doc.add_heading('Fallback DOCX',1);doc.add_paragraph('Works without LibreOffice.');docx=td/'a.docx';doc.save(docx)
    wb=Workbook();ws=wb.active;ws.append(['Item','Price']);ws.append(['Test',99]);xlsx=td/'a.xlsx';wb.save(xlsx)
    prs=Presentation();slide=prs.slides.add_slide(prs.slide_layouts[6]);box=slide.shapes.add_textbox(Inches(1),Inches(1),Inches(5),Inches(1));box.text='Fallback PPTX';pptx=td/'a.pptx';prs.save(pptx)
    for srcfile in (docx,xlsx,pptx):
        out=td/(srcfile.stem+'.pdf')
        _office_fallback(srcfile,out)
        assert out.exists() and out.stat().st_size>500,srcfile
        with fitz.open(out) as d: assert len(d)>=1
        print('PASS office fallback',srcfile.suffix,out.stat().st_size)

# v10 regression: replacing text must never delete table/page line art, and
# whiteout must be saved as a fill with no black stroke.
d=fitz.open();p=d.new_page(width=300,height=220)
p.draw_rect(fitz.Rect(20,20,280,200),color=(0,0,0),width=2)
p.draw_line(fitz.Point(20,80),fitz.Point(280,80),color=(0,0,0),width=1)
p.insert_text((40,78),'Border Text',fontsize=14)
raw=d.tobytes();d.close()
prep2=client.post('/api/editor/prepare',files={'file':('border.pdf',raw,'application/pdf')})
assert prep2.status_code==200,prep2.text
j=prep2.json(); eid=j['editor_id']; sp2=next(s for s in j['pages'][0]['text'] if 'Border Text' in s['text'])
saved2=client.post(f'/api/editor/{eid}/save',json={'actions':[
    {'type':'replace_text','page':0,'rect':sp2['bbox'],'origin':sp2['origin'],'font':sp2['font'],'text':'Edited Text','size':14,'color':'#111111'},
    {'type':'whiteout','page':0,'rect':[210,120,260,145]},
]})
assert saved2.status_code==200,saved2.text[:300]
with fitz.open(stream=saved2.content,filetype='pdf') as outdoc:
    pg=outdoc[0]; drawings=pg.get_drawings()
    # Original rectangle + rule survive; whiteout adds a fill-only drawing.
    assert any(x.get('color') == (0.0,0.0,0.0) for x in drawings),drawings
    assert any(x.get('color') is None and x.get('fill') == (1.0,1.0,1.0) for x in drawings),drawings
    assert 'Edited Text' in pg.get_text() and 'Border Text' not in pg.get_text()
print('PASS v10 border preservation + whiteout fill-only output')
