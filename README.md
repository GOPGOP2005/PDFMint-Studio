# PDFMint Studio — Your all-in-one PDF website

A complete, original-design PDF tools web application with a clean, tool-first workflow inspired by modern online PDF utilities. It is built with FastAPI, PyMuPDF and a responsive HTML/CSS/JavaScript frontend. **Not affiliated with iLovePDF or Sejda.** The interface uses PDFMint branding and original code.

## What you get

- A fully responsive clean white/red home page with a compact navigation bar, category pills, search and a five-column tool catalog on desktop.
- Simple tool upload screens: title + description + one large file button first; settings appear only after a file is selected.
- A dedicated Edit PDF application layout with page thumbnails on the left, editing tools across the top, document canvas in the middle, and save/help controls on the right.
- **34 named tools** across organize, optimize, convert, edit, protect, intelligence and automated workflows.
- Actual back-end PDF processing. It is **not** a static mockup or clickable-only design.
- Self-hostable code, Windows startup helper, a Dockerfile, end-to-end tests and no required user account.

### Feature list

| Organize | Optimize | Convert to PDF | Convert from PDF |
|---|---|---|---|
| Merge PDF | Compress PDF | JPG to PDF | PDF to JPG |
| Split PDF | Repair PDF | Word to PDF | PDF to Word |
| Remove pages | OCR PDF | PowerPoint to PDF | PDF to PowerPoint |
| Extract pages | | Excel to PDF | PDF to Excel |
| Organize pages | | Uploaded HTML to PDF | PDF to PDF/A (best-effort) |
| Scan images to PDF | | | |

| Edit | Security | Intelligence | Workflows |
|---|---|---|---|
| Rotate PDF | Unlock PDF | Smart summary | Merge → Compress |
| Page numbers | Protect PDF | Translate text | Watermark → Compress |
| Watermark | Visible signature mark | PDF to Markdown | Rotate → Compress |
| Crop PDF | Permanent text redaction | | Compress → Protect |
| Full interactive editor | Compare PDF text | | |
| PDF Forms | | | |

**Editor scope:** Edit PDF now uses an extracted-text editing layer. The server renders a temporary clean page with text glyphs removed while preserving vector borders, table rules and images; the extracted text is placed back in the browser as visible editable text boxes. Click any visible text and type directly in place. Hover shows a light blue box, focus shows a blue caret/outline, zoom runs from 25% to 500%, and the Hand tool pans the page. Whiteout, Highlight and Shape show a live guide continuously while the mouse button is held (even if the pointer leaves the page because pointer capture is used); releasing the mouse finalizes the selection. Preview guide outlines are not written into the saved PDF. Existing-text replacement removes only text glyphs and preserves page borders/graphics. Complex embedded fonts, curved text and highly designed typography may not reproduce perfectly.

**Other honest scope notes:** PDF-to-PowerPoint creates image-backed slides (not editable slide objects). PDF-to-Word uses `pdf2docx` when installed, otherwise PDFMint reconstructs an editable DOCX with detected typography, page flow and tables; very complex layouts can still need cleanup. Signing in the standalone Sign tool is a **visible signature stamp**, not a certificate-based digital signature. PDF comparison is text-based, not image-based. HTML to PDF supports uploaded HTML, **not arbitrary external webpage URLs**. Translating extracted text uses a server-side AI API key and exports DOCX without recreating the original layout. PDF/A should be validated with a specialist conformance checker before archival use.



## v16 batch conversion + custom signature placement

- Every tool in **Convert to PDF** and **Convert from PDF** now accepts multiple source files. Batch conversions return one ZIP containing one converted result per source file with readable filenames. JPG → PDF keeps its existing behavior of combining selected images into one PDF.
- **Sign PDF** now opens a real page preview after upload. The handwritten signature is a draggable object: move it anywhere on the page, resize it from the corner handle, and switch pages before processing. The backend receives page-relative coordinates so the final PDF uses the chosen location on any page size.
- The old automatic bottom-right signature behavior remains as a backend fallback for older clients that do not send custom placement coordinates.

## v12 interface refresh

The v12 frontend uses a much simpler interaction model:

- Home starts with Merge, Split, Compress, PDF to Word and PDF to PowerPoint, followed by the rest of the tools.
- Tool pages do not show an intimidating settings panel before upload; options appear after files are selected.
- Edit PDF switches into a dedicated three-column editor with page thumbnails, a top toolbar, extracted editable text, zoom/page navigation and a persistent Save changes control.
- Whiteout/Highlight/Shape show a live drag guide until pointer release; guide outlines are editor-only and are not burned into Whiteout output.

## Run on Windows (easy)

1. Install **Python 3.11 or newer**. During installation, enable **Add Python to PATH**.
2. Extract this ZIP to a normal folder, e.g. `Desktop\PDFMint-Studio`.
3. Double-click **`start_windows.bat`**. The first run installs packages into a project-local `.venv` folder.
4. Open **http://127.0.0.1:8000** on the same PC. The startup script also opens it automatically.
5. Keep the command window open while the local website is running. Press **Ctrl+C** to stop.

**Office conversion no longer requires LibreOffice for modern files:** DOCX, XLSX and PPTX have built-in local PDF fallbacks. If LibreOffice is installed, PDFMint automatically prefers it for higher-fidelity conversion and it is still required for legacy DOC/XLS/PPT/ODT/ODS/ODP files.

For OCR and PDF/A support, double-click **`install_optional_windows.bat`** once. It installs Tesseract OCR, Ghostscript and LibreOffice through Windows Package Manager (winget). PDFMint also checks their common Windows installation folders, so they do not have to be manually added to PATH. Restart `start_windows.bat` after installation. For improved PDF → DOCX layout, run `pip install -r requirements-optional.txt` inside `.venv`. WeasyPrint on Windows may require Pango/GTK runtime dependencies.

## Run in PowerShell manually

```powershell
cd "C:\path\to\PDFMint-Studio"
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m uvicorn app:app --host 127.0.0.1 --port 8000
```

Open: http://127.0.0.1:8000

If Windows blocks `.ps1` activation, use `.venv\Scripts\python.exe -m pip install -r requirements.txt` and `.venv\Scripts\python.exe -m uvicorn app:app --host 127.0.0.1 --port 8000` without activating.

## Run on Linux or macOS

```bash
cd PDFMint-Studio
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app:app --host 127.0.0.1 --port 8000
```

For document conversion, install `libreoffice`; for OCR, `tesseract` and relevant language packs; for archival PDF generation, `ghostscript`.

## Deploy online using Docker

```bash
docker build -t pdfmint-studio .
docker run --rm -p 8000:8000 pdfmint-studio
```

Then open http://localhost:8000. Most server providers that accept Docker deployments can run this image. The provided Dockerfile installs LibreOffice, Tesseract with English/Hindi language data, Ghostscript, fonts and rendering libraries. Set `PORT` if your host requires another listening port, and configure HTTPS in your hosting platform.

**Important:** Uploading only the `/static` folder to Netlify will show the UI but **will not process PDF files**. The app needs its FastAPI/Python service (either a Docker host or a separate back-end API).

## Optional AI features

Core tools do **not** need an API key. Summarize PDF runs a simple local extractive summary when no key is configured. For true generative summaries and translation, set these **server-side environment variables**, not in JavaScript:

```
OPENAI_API_KEY=your_secret_key
OPENAI_MODEL=gpt-4o-mini
```

AI usage may have a charge from the provider. Translation needs an API key; the site displays a helpful error when none is set. Never put this key in an uploaded document, Git repo or frontend assets.

## API

- `GET /api/tools` — returns public tool catalog + upload limits.
- `GET /api/health` — server status and installed converter availability.
- `POST /api/process/{tool_id}` — multipart form-data: one or more `files` and a JSON string field `options`.
- `POST /api/editor/prepare` — starts a short-lived interactive PDF editing session.
- `GET /api/editor/{id}/page/{n}.png` — renders a page preview for the editor.
- `GET /api/editor/{id}/page/{n}/clean.png` — renders the temporary text-removed background used by the extracted-text editor.
- `POST /api/editor/{id}/save` — applies editor actions and returns the edited PDF.

Example:

```bash
curl -X POST http://127.0.0.1:8000/api/process/rotate \
  -F 'files=@example.pdf' \
  -F 'options={"angle":"90"}' -o rotated.pdf
```

## File privacy and production safety

Each request gets a separate temporary directory, deleted after the output has been served. No user accounts, permanent storage or analytics are added by this code. The default total upload limit is **80 MB** and there is a **40-file** cap. Configure `MAX_UPLOAD_MB` on the server to adjust the upload threshold (the frontend currently displays the default 80 MB limit; update it if you change the server setting).

**Before running this publicly at scale:** isolate document processing in sandboxed workers/containers, add authentication or request rate limiting, file scanning, CPU/time/memory quotas and monitoring; put the API behind HTTPS and a reverse proxy. PDFs and Office documents can be malicious, and conversions are CPU-heavy. The shipped server is a working, self-hostable starter, **not a fully hardened, zero-maintenance enterprise service**. Large conversion requests on cheap hosting can time out.

## Test

```bash
python tests/smoke_test.py
python tests/api_test.py
python tests/editor_test.py
```

The smoke test exercises 34 representative operation cases with real PDF/Office/image inputs. The API test verifies the actual HTTP response and major PDF behaviors, including password lock/unlock and irreversible redaction. The editor test verifies existing-text replacement, new text, image insertion, annotation actions and DOCX/XLSX/PPTX conversion without LibreOffice.

## Folder structure

```text
PDFMint-Studio/
├── app.py                 FastAPI server, upload handling, download responses
├── engine.py              Real document processing
├── catalog.py             All tool metadata and settings
├── static/
│   ├── index.html         Responsive frontend
│   ├── styles.css         Complete visual design
│   ├── app.js             Interactive frontend
│   └── favicon.svg        Original project icon
├── requirements.txt       Core Python packages
├── requirements-optional.txt
├── start_windows.bat      Double-click launcher for Windows
├── install_optional_windows.bat  One-click OCR/PDF-A/legacy Office helpers
├── start_linux.sh         Linux starter
├── Dockerfile             Deployment image
└── tests/                 Executable integration tests
```

**Project brand:** PDFMint Studio (independent). Do not use a third party's logo or claim affiliation with iLovePDF in deployment, advertising or pricing.


## PDF to Word quality

PDFMint now includes a built-in layout-preserving PDF to Word converter. It reconstructs detected tables, font sizes/bold text, page size, and document flow even when the optional `pdf2docx` package is not installed. Complex magazine-style PDFs can still differ from the source because PDF has no native Word document structure.

## v8 Simple PDF Editor

The Edit PDF workspace now uses a simplified, Sejda-inspired interaction model while keeping PDFMint branding:

- One **Text** tool: click existing text to edit it in place; click blank space to add new text.
- The old right-side settings panel is removed from the editor.
- Compact top toolbar for Text, Image, Sign, Whiteout, Highlight, Draw, Shape and Hand.
- Context formatting appears only when text is selected (font, size, color, delete).
- Page navigation, zoom and **Apply changes** stay in the top bar.
- Real 25%–500% zoom and Ctrl+mouse-wheel zoom remain supported.
- Existing text replacements are removed and reinserted by the backend when saving; they are not just visual browser overlays.

For scanned/image-only PDFs, run OCR first if you need selectable/editable text.

## v9 extracted-text PDF editor

The Edit PDF workspace now renders a clean page background with original text glyphs removed, then places extracted PDF text as editable HTML text boxes on top. Click any text and type directly. Changes are written back to the original PDF when you choose Apply changes. Scanned/image PDFs still need OCR first.


### v10 editor fixes
- Original page raster remains intact in the editor, preventing page borders / table rules from shifting or disappearing.
- Text replacement preserves vector graphics and images in the output PDF.
- Whiteout shows a dashed black guide in the editor, but the saved PDF contains only the white patch (no guide border).
- Added overlays receive a visible editor selection outline.


## v11 extracted text editor
- Edit PDF uses a text-clean page raster, preserving vector borders/tables/images.
- Extracted PDF text is rendered as visible contenteditable HTML at PDF coordinates.
- Hover shows a blue guide; click shows a blue edit box/caret; typing updates the saved PDF action.
- Whiteout/shape guides remain preview-only.

## Edit PDF direct-open flow (v13)

- Selecting a PDF in **Edit PDF** immediately switches to the full editor workspace; the generic selected-files/options card is skipped.
- Password-protected PDFs show an in-editor password gate. The editor is not prepared until the correct password is supplied.
- After authentication, the server creates only a short-lived decrypted working copy in the local temporary editor session; the password is not stored in editor metadata.
- Incorrect passwords stay on the password gate and can be retried without leaving the Edit PDF screen.


## v14 UI update
- Every tool now opens on the same clean, centered upload screen.
- Upload button labels adapt to PDF/Word/PowerPoint/Excel/image inputs.
- Merge PDF shows **Select PDF files** and **or drop PDFs here** before any options.
- Edit PDF still opens directly into the editor after upload, with the password gate for protected PDFs.
- Sign in the PDF editor now creates a handwritten-style transparent signature image instead of plain text.
## v16 unified download/results screen

Every PDFMint tool now ends on one consistent result page: a large contextual Download button, a back action, start-over/delete controls, filename/size, and three relevant “Continue to…” tools. The result screen is used by Merge, Split, Compress, Convert, Protect, Sign, Edit, and all other processing tools.


## PDFMint-branded download filenames (v17)

Every user-facing download now includes the `PDFMint_` prefix. This applies to PDFs, JPG/PNG image exports, DOCX/XLSX/PPTX outputs, and batch ZIP downloads. Files stored inside batch ZIP archives are branded too (for example `PDFMint_invoice.docx` or `PDFMint_page-1.jpg`). Existing names that already start with PDFMint are not double-prefixed.
