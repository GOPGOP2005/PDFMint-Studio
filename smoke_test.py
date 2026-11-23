"""Smoke-test actual outputs for the major tools. Run: python tests/smoke_test.py"""
from pathlib import Path
import io,sys,tempfile,json
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import fitz
from PIL import Image
from docx import Document
from openpyxl import Workbook
from pptx import Presentation
from engine import process,ProcessingError

base=Path(tempfile.mkdtemp(prefix='pdfmint-smoke-'))
pdf=base/'sample.pdf'; doc=fitz.open()
for i in range(3):
    p=doc.new_page(width=595,height=842)
    p.insert_text((70,80),f'Test PDF page {i+1} - this is a sample PDF about productivity.',fontsize=12)
    p.insert_text((70,105),f'This line contains PRIVATE SECRET on page {i+1}.',fontsize=12)
    p.insert_text((70,135),f'Line two discusses page {i+1} and business work.',fontsize=12)
doc.save(pdf);doc.close()
other=base/'other.pdf';doc=fitz.open();p=doc.new_page();p.insert_text((50,50),'Other doc changed text.');doc.save(other);doc.close()
image=base/'image.jpg';Image.new('RGB',(640,480),'#a8d4ff').save(image)
word=base/'test.docx';d=Document();d.add_heading('Test Word',0);d.add_paragraph('Hello PDFMint');d.save(word)
excel=base/'test.xlsx';wb=Workbook();wb.active.append(['One','Two']);wb.active.append([1,2]);wb.save(excel)
ppt=base/'test.pptx';presentation=Presentation();slide=presentation.slides.add_slide(presentation.slide_layouts[5]);presentation.save(ppt)
html=base/'test.html';html.write_text('<h1>PDFMint HTML</h1><p>Conversion success.</p>')
form=base/'form.pdf';fd=fitz.open();pg=fd.new_page();w=fitz.Widget();w.field_name='Name';w.field_type=fitz.PDF_WIDGET_TYPE_TEXT;w.rect=fitz.Rect(100,100,300,120);pg.add_widget(w);fd.save(form);fd.close()

cases=[
 ('merge',[pdf,other],{}),('split',[pdf],{'mode':'each'}),('split',[pdf],{'mode':'ranges','pages':'1-2,3'}),('split',[pdf],{'mode':'extract','pages':'2,3'}),
 ('remove-pages',[pdf],{'pages':'2'}),('extract-pages',[pdf],{'pages':'2-3'}),('organize',[pdf],{'pages':'3,1,2,1'}),
 ('scan',[image],{'enhance':'grayscale'}),('compress',[pdf],{'quality':'strong'}),('repair',[pdf],{}),
 ('jpg-to-pdf',[image],{}),('word-to-pdf',[word],{}),('ppt-to-pdf',[ppt],{}),('excel-to-pdf',[excel],{}),('html-to-pdf',[html],{}),
 ('pdf-to-jpg',[pdf],{}),('pdf-to-word',[pdf],{}),('pdf-to-ppt',[pdf],{}),('pdf-to-excel',[pdf],{}),('pdf-to-pdfa',[pdf],{}),
 ('rotate',[pdf],{'angle':'90'}),('page-numbers',[pdf],{}),('watermark',[pdf],{'text':'SAMPLE'}),('crop',[pdf],{'left':20,'top':10,'bottom':10,'right':20}),
 ('edit',[pdf],{'text':'Reviewed','page':1,'x':12,'y':40}),('forms',[form],{'action':'inspect'}),('forms',[form],{'action':'fill','values':'{"Name":"Alice"}'}),
 ('protect',[pdf],{'password':'test1234'}),('sign',[pdf],{'signature':'A User'}),('redact',[pdf],{'text':'PRIVATE SECRET'}),
 ('compare',[pdf,other],{}),('summarize',[pdf],{}),('markdown',[pdf],{}),('ocr',[pdf],{'language':'eng'})
]
fail=[]
for k,(tool,inputs,opts) in enumerate(cases):
    work=base/f'task-{k}';work.mkdir()
    try:
        res=process(tool,inputs,opts,work)
        assert res.is_file() and res.stat().st_size>0
        if res.suffix=='.pdf':
            with fitz.open(res) as check:
                assert len(check)>0
        print(f'PASS {tool:15s} {res.name:25s} {res.stat().st_size:>8d} B')
    except Exception as ex:
        fail.append((tool,str(ex)));print('FAIL',tool,repr(ex))
print('FINISHED',len(cases)-len(fail),'/',len(cases))
for name,err in fail:print('  FAILED',name,err)
if fail:sys.exit(1)
