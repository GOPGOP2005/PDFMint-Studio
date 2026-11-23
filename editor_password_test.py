from pathlib import Path
import tempfile,sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import fitz
from fastapi.testclient import TestClient
from app import app


def main():
    with tempfile.TemporaryDirectory(prefix='pdfmint-password-test-') as td:
        path=Path(td)/'locked.pdf'
        doc=fitz.open(); page=doc.new_page(); page.insert_text((72,72),'PASSWORD EDITOR TEST',fontsize=14)
        doc.save(str(path), encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw='owner123', user_pw='pass123')
        doc.close()
        raw=path.read_bytes(); client=TestClient(app)
        r=client.post('/api/editor/prepare',files={'file':('locked.pdf',raw,'application/pdf')})
        assert r.status_code==423 and r.json().get('password_required') is True
        r=client.post('/api/editor/prepare',files={'file':('locked.pdf',raw,'application/pdf')},data={'password':'wrong'})
        assert r.status_code==401 and r.json().get('invalid_password') is True
        r=client.post('/api/editor/prepare',files={'file':('locked.pdf',raw,'application/pdf')},data={'password':'pass123'})
        assert r.status_code==200
        data=r.json(); assert data['pages'][0]['text'][0]['text']=='PASSWORD EDITOR TEST'
        eid=data['editor_id']
        assert client.get(f'/api/editor/{eid}/page/1/clean.png').status_code==200
        print('PASS editor password gate/authentication')

if __name__=='__main__': main()
