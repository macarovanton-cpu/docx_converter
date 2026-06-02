# docx_converter Project Status

Last updated: 2026-06-02

## Current State

Project: `docx_converter`

Type: Streamlit application.

Original capability: Markdown to DOCX.

Current import capability:

- files to Markdown through Microsoft MarkItDown;
- batch upload for PDF, DOCX, XLSX, and PPTX;
- PDF page ranges;
- PDF diagnostics and image-only page detection;
- individual `.md` downloads;
- ZIP download;
- `combined.md`;
- shared page range for PDF batch conversion;
- cached PDF diagnostics;
- oversized page range protection;
- temporary file cleanup.

## Current Phase

OCR PoC / preparing OCR backend wrapper.

## Next Recommended Task

Create `ocr_converter.py` with an OCRmyPDF subprocess wrapper.

Before creating it, check the working tree: a local uncommitted PoC file named
`ocr_converter.py` may already exist from the previous OCR backend step.

Expected wrapper behavior:

- call OCRmyPDF through `sys.executable -m ocrmypdf`;
- support `rus+eng` by default;
- validate input PDF exists;
- validate output PDF is created;
- raise readable `RuntimeError` messages with stdout/stderr on subprocess
  failure;
- include `check_ocr_dependencies()` that reports all dependency statuses
  without failing on the first error.

## Local OCR Dependencies

- Tesseract path: `C:\Program Files\Tesseract-OCR\tesseract.exe`
- Ghostscript path: `C:\Program Files\gs\gs10.07.1\bin\gswin64c.exe`
- OCR languages: `eng`, `rus`, `osd`
- OCRmyPDF version: `17.5.0`

Manual OCR checks already completed:

- Tesseract is installed locally.
- Tesseract languages `eng` and `rus` are available.
- Ghostscript is installed.
- OCRmyPDF is installed in `.venv`.
- OCRmyPDF can create a searchable PDF.
- The current app can convert an OCR PDF to Markdown.
- Raw OCR Markdown quality is not good enough without cleanup.

## Do Not Commit

Do not commit:

- `.venv`;
- `test_files`;
- `__pycache__`;
- `.streamlit/secrets.toml`;
- temporary PDF files;
- temporary Markdown files;
- temporary ZIP files.

## OCR Environment Check Commands

From the project root:

```powershell
.\.venv\Scripts\python.exe -m ocrmypdf --version
```

```powershell
& "C:\Program Files\Tesseract-OCR\tesseract.exe" --version
```

```powershell
& "C:\Program Files\Tesseract-OCR\tesseract.exe" --list-langs
```

```powershell
& "C:\Program Files\gs\gs10.07.1\bin\gswin64c.exe" --version
```

Manual OCR smoke test:

```powershell
.\.venv\Scripts\python.exe -m ocrmypdf --skip-text --deskew --rotate-pages -l rus+eng input.pdf output_ocr.pdf
```

If using `ocr_converter.py`:

```powershell
.\.venv\Scripts\python.exe -c "from ocr_converter import check_ocr_dependencies; print(check_ocr_dependencies())"
```

```powershell
.\.venv\Scripts\python.exe -c "from ocr_converter import ocr_pdf_to_searchable_pdf; ocr_pdf_to_searchable_pdf(r'input.pdf', r'output_ocr.pdf')"
```

## Recent Work Log

2026-06-02:

- Created project memory files: `PROJECT_PLAN.md`, `PROJECT_STATUS.md`, and
  `AGENTS.md`.
- Files changed in this task: project memory markdown files only.
- Checks run for this task: `git diff --stat`.
- Next step: finish or verify the OCR backend wrapper, then update this file.

## Next Prompt Target For Codex

Implement or verify `ocr_converter.py` as a backend-only OCR PoC:

- do not change `app.py`;
- do not change `file_converter.py`;
- do not change `convert.py`;
- do not add OCR to UI;
- do not add dependencies;
- run `python -m py_compile app.py file_converter.py convert.py ocr_converter.py`;
- update `PROJECT_STATUS.md` with changed files, checks run, and the next step.
