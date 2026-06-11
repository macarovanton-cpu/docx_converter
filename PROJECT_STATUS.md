# docx_converter Project Status

Last updated: 2026-06-11

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
- OCR mode UI switch (`off` / `auto`) for the files-to-Markdown mode;
- OCR auto candidate display based on existing PDF diagnostics;
- OCR backend integration for PDF candidates without a text layer;
- OCR auto respects selected PDF page ranges when deciding whether OCR is
  needed;
- backend-only deterministic OCR Markdown cleanup PoC, not connected to the UI
  or conversion flow;
- oversized page range protection;
- UI validation for page ranges;
- temporary file cleanup.

The `feature/markitdown-import` PR has been merged into `main`.

## Current Phase

Backend-only deterministic cleanup PoC for OCR Markdown refined after checking
real OCR Markdown. The cleanup is not connected to the UI or existing
conversion flow.

## Branch Context

Current branch: `feature/ocr-markdown-cleanup-poc`

## Next Recommended Task

- rerun `cleanup_ocr_markdown()` on the real `ТТ.md` OCR Markdown and evaluate
  readability;
- keep cleanup disconnected from conversion until a separate explicit task;
- keep LLM cleanup as a separate future task.

## Do Not Do Now

- Do not add LLM cleanup.
- Do not connect cleanup to the UI or conversion flow without a separate task.
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

2026-06-11:

- Refined the backend-only deterministic OCR Markdown cleanup PoC after
  checking a real OCR Markdown file.
- Added conservative handling for inline OCR bullet markers and glued technical
  subheadings in `markdown_cleanup.py`.
- Added focused tests for inline `;e` bullets and glued `Навес:` headings in
  `tests/test_markdown_cleanup.py`.
- UI was not changed; cleanup is still not connected to conversion.
- Recommended next step: rerun cleanup on the real `ТТ.md` and evaluate
  readability.

2026-06-11 previous cleanup step:

- Added backend-only PoC deterministic cleanup for OCR Markdown in
  `markdown_cleanup.py`.
- `cleanup_ocr_markdown()` performs conservative normalization and cleanup:
  line ending normalization, form feed removal, intra-line space cleanup, blank
  line compression, common OCR bullet marker conversion, a small OCR artifact
  dictionary, and line breaks before obvious large numbered sections.
- Added focused unit tests in `tests/test_markdown_cleanup.py`.
- UI was not changed.
- Cleanup is not connected to conversion yet.
- Files changed: `markdown_cleanup.py`, `tests/test_markdown_cleanup.py`,
  `PROJECT_STATUS.md`.
- Checks run: `python -m unittest tests.test_markdown_cleanup` and
  `python -m py_compile markdown_cleanup.py tests/test_markdown_cleanup.py`.
- Recommended next step: manually run cleanup on a real OCR-generated `.md` and
  evaluate quality.

2026-06-11 previous step:

- Connected OCR `auto` mode to `ocr_converter.ocr_pdf_to_searchable_pdf()`
  for PDF files whose diagnostics show pages without a text layer.
- `off` still uses the existing MarkItDown-only conversion path.
- Text-layer PDFs in `auto` skip OCR and continue through the existing
  MarkItDown flow.
- OCR candidate PDFs are written to a temporary searchable PDF, converted
  through MarkItDown, and the temporary OCR PDF is removed in `finally`.
- OCR `auto` now checks only the selected PDF pages when `page_range` is set,
  so image-only pages outside the selected range do not trigger OCR.
- The per-file OCR candidate UI status uses the selected `page_range` as well.
- OCR errors are returned as the current per-file result error, so the batch
  can continue processing other files.
- Added focused unit tests for OCR candidate detection, OCR skip/apply
  behavior, page-range-aware OCR decisions, temp OCR PDF cleanup, and OCR
  backend error wrapping.
- Files changed: `app.py`, `ocr_auto_mode.py`, `tests/test_ocr_auto_mode.py`,
  `PROJECT_STATUS.md`.
- Checks run: `python -m unittest tests.test_ocr_auto_mode` and
  `python -m py_compile app.py file_converter.py ocr_converter.py
  ocr_auto_mode.py tests/test_ocr_auto_mode.py`.
- Recommended next step: manual Streamlit smoke check with real text-layer and
  image-only PDFs before committing.

2026-06-11 earlier step:

- Added OCR mode UI in `Файлы -> Markdown` with `off` / `auto`; default is
  `off`.
- `off` keeps the existing MarkItDown conversion behavior unchanged.
- `auto` does not call `ocr_pdf_to_searchable_pdf()` yet.
- `auto` uses the existing cached PDF diagnostics to show whether PDF files
  are OCR candidates because they contain pages without a text layer, or
  whether OCR is not needed because a text layer is present.
- Files changed: `app.py`, `PROJECT_STATUS.md`.
- Checks to run: `python -m py_compile app.py file_converter.py` and a manual
  Streamlit smoke check of `Файлы -> Markdown` with OCR `off` and `auto`.
- Recommended next step: in a separate task, connect OCR `auto` to the backend
  helper for OCR candidates only.

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

Manual-check `feature/ocr-ui-auto`:

- run the Streamlit app locally;
- verify `OCR mode = off` keeps the existing MarkItDown flow;
- verify `OCR mode = auto` skips OCR for text-layer PDFs;
- verify `OCR mode = auto` with `page_range` skips OCR when selected pages have
  text, even if another PDF page is image-only;
- verify `OCR mode = auto` applies OCR for image-only PDFs and still returns
  other batch files if one file errors;
- do not add LLM cleanup;
- do not add deterministic cleanup;
- do not do a large refactor;
- do not change `convert.py`;
- preserve the existing Markdown to DOCX workflow;
- preserve the existing MarkItDown import workflow.
