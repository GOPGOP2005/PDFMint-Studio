"""One source of truth for PDFMint tools shown by the UI and processed by the API."""

def opt(key, label, kind='text', default='', choices=None, hint=None, minimum=None, maximum=None):
    d = dict(key=key,label=label,type=kind,default=default)
    if choices is not None: d['choices']=[{'value':v,'label':n} for v,n in choices]
    if hint: d['hint']=hint
    if minimum is not None: d['min']=minimum
    if maximum is not None: d['max']=maximum
    return d

def tool(id, name, desc, category, icon, accept='.pdf', multiple=False, min_files=1, options=None, note=None):
    d=dict(id=id,name=name,description=desc,category=category,icon=icon,accept=accept,multiple=multiple,min_files=min_files,options=options or [])
    if note: d['note']=note
    return d

TOOLS = [
    tool('workflow','PDF Workflows','Run multiple PDF actions in one step.','Workflows','workflow',multiple=True,options=[opt('preset','Workflow preset','select','merge-compress',[('merge-compress','Merge → Compress'),('watermark-compress','Watermark → Compress'),('rotate-compress','Rotate → Compress'),('compress-protect','Compress → Protect')]),opt('text','Watermark text (if used)','text','CONFIDENTIAL'),opt('angle','Rotate angle (if used)','select','90',[('90','90°'),('180','180°'),('270','270°')]),opt('password','New password (if used)','password','')],note='Choose the preset, upload PDF documents and process them in sequence.'),
    tool('merge','Merge PDF','Combine multiple PDF files in your chosen order.','Organize','files',multiple=True,min_files=2),
    tool('split','Split PDF','Create PDFs from pages or page ranges.','Organize','scissors',options=[opt('mode','Split mode','select','each',[('each','Every page (ZIP)'),('ranges','Selected ranges (ZIP)'),('extract','Extract selected pages (one PDF)')]),opt('pages','Pages / ranges','text','',hint='For ranges: 1-3,5,8-10. Leave blank to include all.')]),
    tool('remove-pages','Remove pages','Delete selected pages from your document.','Organize','file-minus-2',options=[opt('pages','Pages to remove','text','2,4',hint='Example: 2,4-6')]),
    tool('extract-pages','Extract pages','Keep only the pages that you specify.','Organize','file-output',options=[opt('pages','Pages to extract','text','1-3',hint='Example: 1-3,7')]),
    tool('organize','Organize PDF','Reorder and duplicate pages using page numbers.','Organize','layout-grid',options=[opt('pages','New page order','text','3,1,2',hint='Example: 3,1,2,4-6')]),
    tool('scan','Scan to PDF','Turn camera photos into clean PDF scans.','Organize','scan-line',accept='image/*',multiple=True,options=[opt('enhance','Document enhancement','select','grayscale',[('grayscale','Grayscale + contrast'),('color','Original color'),('bw','Black and white')])]),
    tool('compress','Compress PDF','Reduce PDF size with adjustable image quality.','Optimize','minimize-2',options=[opt('quality','Compression level','select','balanced',[('light','Light (best quality)'),('balanced','Balanced'),('strong','Strong (smallest size)')])]),
    tool('repair','Repair PDF','Attempt to rebuild a readable damaged PDF.','Optimize','wrench'),
    tool('ocr','OCR PDF','Make scanned pages searchable with text recognition.','Optimize','scan-text',options=[opt('language','OCR language','text','eng',hint='Tesseract language code, e.g. eng or eng+hin.')],note='Requires Tesseract OCR and the requested language data.'),
    tool('jpg-to-pdf','JPG to PDF','Convert JPG, PNG and WEBP images into one PDF.','Convert to PDF','images',accept='image/*',multiple=True,options=[opt('paper','Page size','select','a4',[('a4','A4'),('letter','US Letter'),('fit','Fit each image')]),opt('margin','Margin (pt)','number',20,minimum=0,maximum=120)]),
    tool('word-to-pdf','Word to PDF','Convert Word documents to PDF.','Convert to PDF','file-text',accept='.doc,.docx,.odt',multiple=True,note='DOCX works locally. Legacy DOC/ODT uses LibreOffice when installed.'),
    tool('ppt-to-pdf','PowerPoint to PDF','Export presentations as PDF.','Convert to PDF','presentation',accept='.ppt,.pptx,.odp',multiple=True,note='PPTX works locally. Legacy PPT/ODP uses LibreOffice when installed.'),
    tool('excel-to-pdf','Excel to PDF','Turn spreadsheets into PDF.','Convert to PDF','table-2',accept='.xls,.xlsx,.ods',multiple=True,note='XLSX works locally. Legacy XLS/ODS uses LibreOffice when installed.'),
    tool('html-to-pdf','HTML to PDF','Convert uploaded HTML documents into PDFs.','Convert to PDF','code-2',accept='.html,.htm',multiple=True,note='Converts uploaded HTML; external web URLs are not fetched.'),
    tool('pdf-to-jpg','PDF to JPG','Convert one or more PDFs into JPG images.','Convert from PDF','image',multiple=True,options=[opt('dpi','Image resolution (DPI)','select','130',[('90','90 DPI'),('130','130 DPI'),('200','200 DPI')])]),
    tool('pdf-to-word','PDF to Word','Convert one or more PDFs to editable DOCX files.','Convert from PDF','file-text',multiple=True,note='Preserves detected tables, typography and page flow with PDFMint built-in layout reconstruction; uses pdf2docx automatically if available.'),
    tool('pdf-to-ppt','PDF to PowerPoint','Convert one or more PDFs into PowerPoint presentations.','Convert from PDF','presentation',multiple=True,note='Slides preserve page appearance as images; elements are not individually editable.'),
    tool('pdf-to-excel','PDF to Excel','Convert one or more PDFs into Excel workbooks.','Convert from PDF','sheet',multiple=True),
    tool('pdf-to-pdfa','PDF to PDF/A','Convert one or more PDFs to archival-style PDF/A outputs.','Convert from PDF','archive',multiple=True,note='Requires Ghostscript; validate with a PDF/A compliance checker before formal archiving.'),
    tool('rotate','Rotate PDF','Rotate every PDF page clockwise.','Edit','rotate-cw',options=[opt('angle','Clockwise rotation','select','90',[('90','90°'),('180','180°'),('270','270°')])]),
    tool('page-numbers','Page numbers','Automatically number PDF pages.','Edit','list-ordered',options=[opt('start','Starting number','number',1,minimum=1,maximum=100000),opt('position','Position','select','bottom-center',[('bottom-center','Bottom center'),('bottom-right','Bottom right'),('top-right','Top right')]),opt('size','Font size','number',12,minimum=6,maximum=40)]),
    tool('watermark','Add watermark','Stamp transparent text over every page.','Edit','stamp',options=[opt('text','Watermark text','text','CONFIDENTIAL'),opt('opacity','Opacity (%)','number',22,minimum=5,maximum=100),opt('font_size','Font size','number',48,minimum=12,maximum=100)]),
    tool('crop','Crop PDF','Trim margins from every page in points.','Edit','crop',options=[opt('left','Left margin (pt)','number',20,minimum=0,maximum=500),opt('top','Top margin (pt)','number',20,minimum=0,maximum=500),opt('right','Right margin (pt)','number',20,minimum=0,maximum=500),opt('bottom','Bottom margin (pt)','number',20,minimum=0,maximum=500)]),
    tool('edit','Edit PDF','Edit existing text and add text, images, shapes, highlights, drawings and redactions.','Edit','pencil',options=[],note='Interactive editor: select existing text to replace it, or use annotation tools directly on the page preview.'),
    tool('forms','PDF Forms','Inspect or fill existing interactive form fields.','Edit','clipboard-list',options=[opt('action','Action','select','inspect',[('inspect','Inspect form fields'),('fill','Fill form fields')]),opt('values','Field values (JSON object)','textarea','{"Name": "Your name"}',hint='First inspect fields to learn their exact names.')]),
    tool('unlock','Unlock PDF','Remove an existing password with permission.','Security','unlock-keyhole',options=[opt('password','Current PDF password','password','')],note='You must supply the correct current password.'),
    tool('protect','Protect PDF','Encrypt your PDF with a new password.','Security','lock-keyhole',options=[opt('password','New password','password','')]),
    tool('sign','Sign PDF','Add a handwritten-style visible signature to the document.','Security','pen-tool',options=[opt('signature','Signature name','text','Your Name'),opt('signature_style','Signature style','select','script',[('script','Classic handwritten'),('elegant','Elegant cursive'),('hand','Natural handwriting'),('bold','Bold signature')]),opt('color','Signature color','color','#0a3d9e'),opt('size','Signature size','number',44,minimum=18,maximum=90),opt('page','Page number','number',1,minimum=1,maximum=100000)],note='Drag and resize the signature object on the PDF preview. Visible signature only, NOT a cryptographic/digital certificate signature.'),
    tool('redact','Redact PDF','Permanently remove all matches of sensitive text.','Security','shield-check',options=[opt('text','Text to remove','text','',hint='Exact text match, case-insensitive. Redaction cannot be undone.')]),
    tool('compare','Compare PDF','Generate a per-page text difference report.','Security','git-compare-arrows',multiple=True,min_files=2,note='Compares extracted text, not pixel-perfect graphical differences.'),
    tool('summarize','Smart PDF summary','Create a concise summary from the document.','Intelligence','sparkles',options=[opt('length','Summary length','select','short',[('short','Short'),('medium','Medium'),('long','Detailed')])],note='Uses optional AI API if configured; otherwise local extractive summary.'),
    tool('translate','Translate PDF','Translate extracted PDF text to another language.','Intelligence','languages',options=[opt('language','Target language','select','Hindi',[('Hindi','Hindi'),('English','English'),('Spanish','Spanish'),('French','French'),('Arabic','Arabic')])],note='Requires OPENAI_API_KEY. Produces an editable DOCX; original layout is not preserved.'),
    tool('markdown','PDF to Markdown','Export document text as Markdown.','Intelligence','file-code-2',note='Text-focused export; complex tables may need manual cleanup.'),
]
TOOL_BY_ID = {t['id']:t for t in TOOLS}
