from pathlib import Path
import re, tempfile, zipfile
import fitz
import engine

with tempfile.TemporaryDirectory() as td:
    td=Path(td)
    src=td/'sample.pdf'
    doc=fitz.open(); page=doc.new_page(width=595,height=842)
    page.insert_text((60,70),'Office 2007 Compatibility',fontsize=16)
    page.draw_rect(fitz.Rect(60,110,500,180))
    page.insert_text((70,135),'Table-like content',fontsize=11)
    doc.save(src); doc.close()
    out=td/'converted.docx'
    engine._pdf_to_word_layout(src,out)
    with zipfile.ZipFile(out) as z:
        assert z.testzip() is None
        settings=z.read('word/settings.xml').decode('utf-8')
        xml=z.read('word/document.xml').decode('utf-8')
    assert re.search(r'compatibilityMode[^>]*w:val="12"', settings)
    assert not re.search(r'<w:tbl\s+[^>]*w:tblInd=', xml)
print('OFFICE_2007_DOCX_TEST_PASS')
