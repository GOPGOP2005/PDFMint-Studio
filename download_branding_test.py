"""v17 regression: every user-facing output name carries the PDFMint brand."""
from pathlib import Path
import io, sys, zipfile
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import fitz
from fastapi.testclient import TestClient
from app import app

client=TestClient(app)
doc=fitz.open(); page=doc.new_page(); page.insert_text((70,70),'Brand me'); src=doc.tobytes(); doc.close()

def filename_from(response):
    cd=response.headers.get('content-disposition','')
    assert 'filename=' in cd or "filename*=utf-8''" in cd, cd
    return cd

# Direct PDF output.
r=client.post('/api/process/rotate',files={'files':('photo.pdf',src,'application/pdf')},data={'options':'{"angle":"90"}'})
assert r.status_code==200,r.text[:300]
assert 'PDFMint_' in filename_from(r),filename_from(r)

# Direct image output filename.
r=client.post('/api/process/pdf-to-jpg',files={'files':('photo.pdf',src,'application/pdf')},data={'options':'{}'})
assert r.status_code==200,r.text[:300]
assert 'PDFMint_' in filename_from(r),filename_from(r)

# Multi-file image export ZIP plus every extracted image filename.
r=client.post('/api/process/pdf-to-jpg',files=[('files',('a.pdf',src,'application/pdf')),('files',('b.pdf',src,'application/pdf'))],data={'options':'{}'})
assert r.status_code==200,r.text[:300]
assert 'PDFMint_' in filename_from(r),filename_from(r)
with zipfile.ZipFile(io.BytesIO(r.content)) as z:
    assert z.namelist() and all(Path(n).name.startswith('PDFMint_') for n in z.namelist()),z.namelist()

# Editor save filename.
prep=client.post('/api/editor/prepare',files={'file':('photo.pdf',src,'application/pdf')})
assert prep.status_code==200,prep.text[:300]
eid=prep.json()['editor_id']
r=client.post(f'/api/editor/{eid}/save',json={'actions':[]})
assert r.status_code==200,r.text[:300]
assert 'PDFMint_edited.pdf' in filename_from(r),filename_from(r)
print('PASS PDFMint download branding: direct PDF, image ZIP entries, editor save')
