# AGENTS.md

Project: `docx_converter`

## Command Rule

Always prefix shell commands with `rtk` when possible. If a command needs
PowerShell builtins, use:

```powershell
rtk powershell -NoProfile -Command "<command>"
```

## Working Rules For Codex / Claude

- Work in small, reviewable steps.
- Keep changes scoped to the user's current request.
- Do not do broad refactors unless explicitly requested.
- Do not change `convert.py` unless clearly necessary for the task.
- Do not change `app.py` or `file_converter.py` unless the task explicitly
  requires it.
- Do not add OCR to the UI without a separate explicit task.
- Do not add LLM cleanup without a separate explicit task.
- Do not add new dependencies unless the user explicitly asks for packaging or
  dependency changes.
- Preserve the existing Markdown to DOCX workflow.
- Preserve the existing MarkItDown import workflow unless the task is about
  changing it.

## Files Not To Commit

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

## Project Memory

After each task, update `PROJECT_STATUS.md` with:

- what was done;
- which files changed;
- which checks were run;
- the recommended next step.

Change `PROJECT_PLAN.md` only when the roadmap changes.

## OCR Notes

The OCR phase is currently a PoC for scanned PDFs.

Known local OCR setup:

- Tesseract: `C:\Program Files\Tesseract-OCR\tesseract.exe`
- Ghostscript: `C:\Program Files\gs\gs10.07.1\bin\gswin64c.exe`
- OCRmyPDF: installed in `.venv`
- OCR languages: `eng`, `rus`, `osd`

Do not assume OCR binaries are available on PATH. Prefer explicit local paths
for manual checks, and use `sys.executable -m ocrmypdf` inside Python code.

## Required Final Report Format

At the end of each task, report:

1. changed files;
2. checks run;
3. manual checks needed;
4. safe to commit or not.
