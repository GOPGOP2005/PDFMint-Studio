"""End-to-end HTTP smoke test with real in-memory documents."""
from pathlib import Path
import io,sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from fastapi.testclient import TestClient
import fitz
from app import app

client=TestClient(app)
assert client.get('/').status_code==200
info=client.get('/api/tools').json()
assert len(info['tools'])==34
assert client.get('/api/health').json()['status']=='ok'

doc=fitz.open();pg=doc.new_page();pg.insert_text((70,70),'Top Secret PDFMint');src=doc.tobytes();doc.close()
response=client.post('/api/process/rotate', files={'files':('original.pdf',src,'application/pdf')},data={'options':'{"angle":"90"}'})
assert response.status_code==200,response.text[:400]
assert response.content.startswith(b'%PDF-')
with fitz.open(stream=response.content,filetype='pdf') as output:
    assert output[0].rotation==90
response=client.post('/api/process/protect',files={'files':('test.pdf',src,'application/pdf')},data={'options':'{"password":"test123"}'})
assert response.status_code==200,response.text[:300]
with fitz.open(stream=response.content,filetype='pdf') as locked:
    assert locked.needs_pass
    assert locked.authenticate('test123')>0
protected=response.content
response=client.post('/api/process/unlock',files={'files':('locked.pdf',protected,'application/pdf')},data={'options':'{"password":"test123"}'})
assert response.status_code==200,response.text[:300]
with fitz.open(stream=response.content,filetype='pdf') as unlocked:
    assert not unlocked.needs_pass, 'Unlock should produce unencrypted PDF'
response=client.post('/api/process/redact',files={'files':('original.pdf',src,'application/pdf')},data={'options':'{"text":"Top Secret"}'})
assert response.status_code==200,response.text[:300]
with fitz.open(stream=response.content,filetype='pdf') as clean:
    assert 'Top Secret' not in clean[0].get_text()
response=client.post('/api/process/workflow',files=[('files',('first.pdf',src,'application/pdf')),('files',('second.pdf',src,'application/pdf'))],data={'options':'{"preset":"merge-compress"}'})
assert response.status_code==200,response.text[:300]
with fitz.open(stream=response.content,filetype='pdf') as merged:
    assert len(merged)==2
wrong=client.post('/api/process/merge',files={'files':('sample.pdf',src,'application/pdf')},data={'options':'{}'})
assert wrong.status_code==400
print('API test: 34 tools, homepage, rotate, protect/unlock, redact, workflow, validation PASS')
