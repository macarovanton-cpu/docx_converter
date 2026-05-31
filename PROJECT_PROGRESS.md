# Progress: MarkItDown import feature

Branch: `feature/markitdown-import`

## Implemented

- Added Microsoft MarkItDown backend conversion.
- Added PDF page range support, e.g. `1-3, 7, 10-12`.
- Added PDF page diagnostics:
  - page count;
  - text layer detection;
  - image-only page warning.
- Added Streamlit UI mode switch:
  - `Markdown -> DOCX`;
  - `Файлы -> Markdown`.
- Added batch upload for PDF/DOCX/XLSX/PPTX.
- Added per-file page range input.
- Added shared PDF page range for many typical commercial proposals.
- Added individual `.md` downloads.
- Added ZIP download for all successful `.md` results.
- Added combined Markdown download with separators by source file.
- Updated README.

## Known limitations

- OCR is not implemented.
- Image-only PDF pages are detected but not converted to text.
- Page ranges are supported reliably only for PDF.
- DOCX/XLSX/PPTX are converted as whole files.
- Old `Markdown -> DOCX` path still depends on Google Drive/local DOCX template.

## Tested manually

- Streamlit starts locally.
- `Файлы -> Markdown` works with a 13-page commercial PDF.
- PDF page range `1-2` works.
- PDF page range `4` works for extracting specification pages.
- Page 13 image-only detection works.
- Batch upload of 3 commercial PDFs works.
- ZIP download works.
- `combined.md` works and separates content by source file.

## Next tasks

1. Run Claude Opus architectural review.
2. Review critical and medium issues only.
3. Avoid large refactor before Pull Request.
4. Consider adding:
   - better combined Markdown header/table of contents;
   - local template fallback for Markdown -> DOCX;
   - OCR as a separate future stage.