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
- shared page range for PDF batch conversion;
- PDF diagnostics and image-only page detection;
- image-only warnings;
- individual `.md` downloads;
- ZIP download;
- `combined.md`;
- cached PDF diagnostics;
- oversized page range protection;
- UI validation for page ranges;
- temporary file cleanup.

The `feature/markitdown-import` PR has been merged into `main`.

## Current Phase

Preparing OCR UI auto mode.

## Branch Context

Current/last branch: `chore/project-status-docs`

Next recommended branch: `feature/ocr-ui-auto`

## Next Recommended Task

Add OCR mode to the Streamlit UI:

- supported modes: `off` / `auto`;
- `off`: keep the current MarkItDown-only behavior;
- `auto`: apply OCR only for PDF files without a text layer / image-only PDFs;
- start with a minimal plan for OCR UI auto mode before touching code.

## Do Not Do Now

- Do not add LLM cleanup.
- Do not do a large refactor.
- Do not change `convert.py`.
- Do not add OCR to the main `requirements.txt` without a packaging decision.
- Do not commit `test_files`.

## Project Files

Created project memory files:

- `PROJECT_PLAN.md`;
- `PROJECT_STATUS.md`;
- `AGENTS.md`.

Do not change `PROJECT_PLAN.md` unless the roadmap changes.

## OCR Backend PoC

`ocr_converter.py` has been added as a backend-only OCR PoC helper.

Local OCR setup verified:

- Tesseract is installed.
- Tesseract path: `C:\Program Files\Tesseract-OCR\tesseract.exe`
- Ghostscript path: `C:\Program Files\gs\gs10.07.1\bin\gswin64c.exe`
- OCR languages: `eng`, `rus`, `osd`
- OCRmyPDF version: `17.5.0`
- `ocr_pdf_to_searchable_pdf()` created
  `test_files\ocr_sample_from_wrapper.pdf`.

Important OCR PoC conclusion:

- OCRmyPDF + MarkItDown produces raw Markdown;
- good Markdown quality will need cleanup / LLM cleanup later;
- the next stage is not cleanup, but minimal OCR UI auto mode.

## Do Not Commit

Do not commit:

- `.venv`;
- `test_files`;
- `__pycache__`;
- `.streamlit/secrets.toml`;
- temporary PDF files;
- temporary Markdown files;
- temporary ZIP files.

Use explicit `git add <file>` commands instead of `git add .` when unrelated
local files exist.

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

## How to Resume

1. `git checkout main`
2. `git pull`
3. `git checkout -b feature/ocr-ui-auto`
4. read `PROJECT_PLAN.md` `PROJECT_STATUS.md` `AGENTS.md`
5. start with minimal plan for OCR UI auto mode

## Recent Work Log

2026-06-02:

- `feature/markitdown-import` was merged into `main`.
- MarkItDown mode "Files -> Markdown" works.
- Batch import, PDF page ranges, PDF diagnostics, image-only warnings, ZIP,
  `combined.md`, diagnostic caching, oversized range protection, and UI range
  validation are implemented.
- OCR backend PoC helper `ocr_converter.py` was added and checked locally.
- Files changed in this status update: `PROJECT_STATUS.md` only.
- Checks to run for this status update: `git diff --stat` and
  `git diff -- PROJECT_STATUS.md`.
- Recommended next step: create `feature/ocr-ui-auto` from updated `main` and
  add minimal OCR mode `off` / `auto` in the Streamlit UI.

## Next Prompt Target For Codex

Start `feature/ocr-ui-auto`:

- add OCR mode to the Streamlit UI: `off` / `auto`;
- in `auto`, apply OCR only for PDF files without a text layer / image-only;
- keep the change minimal and scoped;
- do not add LLM cleanup;
- do not do a large refactor;
- do not change `convert.py` unless clearly necessary;
- do not add OCR to the main `requirements.txt` without a packaging decision;
- do not commit `test_files`;
- preserve the existing Markdown to DOCX workflow;
- preserve the existing MarkItDown import workflow.
