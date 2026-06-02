# docx_converter Project Plan

Last updated: 2026-06-02

## Roadmap

### 1. Markdown to DOCX

Status: done.

The original project converts Markdown into styled DOCX documents using the
existing `convert.py` pipeline. This path must remain stable while new import
features are added.

### 2. MarkItDown Import

Status: done.

The `feature/markitdown-import` work added:

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

### 3. OCR For Scanned PDFs

Status: next major phase.

Goal: support scanned/image-only PDFs by producing a searchable OCR PDF first,
then converting that OCR PDF to Markdown through the existing import flow.

Small stages:

1. OCR backend wrapper
   - Add a small `ocr_converter.py` module.
   - Call OCRmyPDF through `sys.executable -m ocrmypdf`.
   - Check local OCR dependencies without changing the UI.

2. OCR mode in UI
   - Add an explicit OCR mode after the backend wrapper is stable.
   - Keep OCR opt-in.
   - Do not change the default Markdown to DOCX workflow.

3. OCR raw Markdown
   - Convert searchable OCR PDFs through the current PDF to Markdown path.
   - Preserve raw OCR Markdown as a downloadable/debuggable output.

4. OCR cleanup / LLM cleanup
   - Evaluate cleanup separately after raw OCR output is visible.
   - Decide whether cleanup is deterministic post-processing, LLM-assisted, or
     both.
   - Keep cleanup optional until quality and costs are understood.

5. Tests
   - Add focused tests for OCR wrapper command construction and errors.
   - Add integration/manual checks for local OCR dependencies.
   - Protect existing Markdown to DOCX and MarkItDown import behavior.

## Not Now

- Do not do a large refactor.
- Do not add OCR directly to the main requirements without a packaging decision.
- Do not break Markdown to DOCX.
- Do not mix OCR cleanup or LLM cleanup into the first OCR backend wrapper step.
- Do not make OCR the default conversion path.
