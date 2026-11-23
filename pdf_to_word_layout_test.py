"""Regression test for PDF -> Word layout reconstruction."""
from pathlib import Path
import sys,tempfile
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import fitz
from docx import Document
from engine import pdf_to_word

base=Path(tempfile.mkdtemp(prefix='pdfmint-word-layout-'))
src=base/'layout.pdf'; out=base/'layout.docx'
doc=fitz.open(); page=doc.new_page(width=595,height=842)
page.insert_text((42,60),'RESUME',fontsize=15,fontname='hebo')
page.insert_text((42,90),'Education',fontsize=15,fontname='hebo')
# Draw a simple 2x2 table and cell text.
xs=[42,260,553]; ys=[120,150,185]
for x in xs: page.draw_line((x,ys[0]),(x,ys[-1]),color=(0,0,0),width=.6)
for y in ys: page.draw_line((xs[0],y),(xs[-1],y),color=(0,0,0),width=.6)
for x,y,text in [(48,140,'Qualification'),(266,140,'Year'),(48,173,'High School'),(266,173,'2024')]:
    page.insert_text((x,y),text,fontsize=10.5)
doc.save(src); doc.close()
pdf_to_word(src,out)
word=Document(out)
assert len(word.tables)>=1, 'Detected PDF table was not preserved in DOCX.'
assert len(word.tables[0].rows)==2 and len(word.tables[0].columns)==2
assert any('RESUME' in p.text for p in word.paragraphs)
print('PASS pdf-to-word layout', out.stat().st_size, 'B')
