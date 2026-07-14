# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Working principles

These are the rules of engagement. Read them before touching code.

**Plan first, then execute.** Start non-trivial work in plan mode. Iterate on the plan until it is right, then switch to auto-accept and let the implementation run in one go. A good plan is the highest-leverage artifact in the session — a bad plan produces 40 changes nobody asked for.

**Verify your own work.** Never hand back code you have not checked. Run `pytest -v`. Run the audit battery in `scratchpad/audit/`. Generate a real `.docx` and inspect it programmatically with `python-docx` (styles, alignment, breaks) — do not assume the output is correct because the code looks correct. Self-verification is worth more than confidence.

**Every mistake becomes a rule.** When a bug or a wrong assumption is found, do not just fix the code — add the rule to this file so it is not repeated. This file is the project's memory across sessions. Keep it under ~200 lines so it is actually read.

**Small, independent changes.** One `# ПРАВКА #N` = one problem = one commit with a descriptive message. Do not bundle unrelated fixes. Do not opportunistically refactor code you were not asked to touch.

**The human stays in the review seat.** Architectural decisions, brand rules, and what ships to clients are the human's call. Ask one clarifying question when the spec is ambiguous — do not guess.

**AI-written code is statistically dirty.** Plausible-looking code that passes tests can still be conceptually wrong: redundant loops, silently dropped data, edge cases that never fire. Bugs here are conceptual, not syntactic. Assume this about your own output and look for it.

## Forbidden patterns

- Do **not** change the signature `convert_md_to_docx(md_text, output_filename, template_path=None, images=None)` — `app.py` depends on it. Extend with optional parameters only.
- Do **not** collapse the numbered `# ПРАВКА #N` edits into a general-purpose constructor. The flat numbered structure is deliberate and is what makes debugging possible.
- Do **not** split `convert.py` into multiple files. The monolith is a conscious choice.
- Do **not** propose a different stack (Pandoc, Quarto, mdbook). The stack is chosen and works.
- Do **not** add dependencies to `requirements.txt` without explicit agreement — Streamlit Cloud does a cold build.
- Do **not** silently swallow data. If a table row has extra cells, a URL is malformed, or an image fails to decode — surface it, do not drop it. Silent corruption in a client-facing КП is the worst failure mode this project has.

## Running the app

```bash
pip install -r requirements.txt
streamlit run app.py
```

Local-run gotchas (learned 2026-07-15):
- MD→DOCX needs a template: `.streamlit/secrets.toml` here is **empty (0 bytes)**, so Drive is unavailable and the app falls back to `C:\Users\tonik\Desktop\docx_converter\template.docx` — that file must exist, otherwise «Шаблон не найден».
- `convert.py` prints emoji (`✅`); on a cp1251 Windows console this raises `UnicodeEncodeError` and the UI shows it as a conversion error. Run with `python -X utf8 -m streamlit run app.py` (or set `PYTHONIOENCODING=utf-8`).

The dev container auto-starts the app on port 8501 after attach (`postAttachCommand` in `.devcontainer/devcontainer.json`).

Run `convert.py` standalone for local batch testing:
```bash
python convert.py   # uses INPUT_FILE / OUTPUT_FILE / TEMPLATE_FILE at the top of the file
```

## Deployment

Push to `main` → Streamlit Cloud picks it up automatically. No CI step required. **Claude Code does not push — the human pushes.**

## Architecture

Seven Python modules:

- **`app.py`** — Streamlit UI. Downloads the `.docx` template from Google Drive (via service account in `st.secrets`), calls `convert_md_to_docx`, and serves the result as a file download. Falls back to a `local_path` if Drive credentials are absent. `DOC_TYPES` dict at the top controls available document types.
- **`convert.py`** — Core Markdown → DOCX engine (~1050 lines). Single public entry point: `convert_md_to_docx(md_text, output_filename, template_path=None, images=None)`. Parses MD into blocks split on `\n\n`, dispatches each block to a typed renderer, writes via `python-docx`.
- **`file_converter.py`** — Reverse direction: DOCX / PDF / TXT → Markdown. Entry point: `convert_file_to_md(file_bytes, filename) → (md_text, images)`. Also hosts the `Файлы -> Markdown` MarkItDown layer (`convert_with_markitdown`) and PDF diagnostics (`analyze_pdf_pages`, via pypdf).
- **`markdown_cleanup.py`** — Deterministic OCR Markdown cleanup (`cleanup_ocr_markdown`). Covered by tests but **not connected to the UI or conversion flow** — backend-only.
- **`ocr_auto_mode.py`** — «OCR or not» orchestrator (`convert_pdf_with_optional_ocr`). Decides whether a PDF needs OCR, honoring the selected page range.
- **`ocr_converter.py`** — OCRmyPDF wrapper via `subprocess`. Also provides `check_ocr_dependencies`.
- **`pdf_core.py`** — Provider-agnostic PDF → Markdown core. Public entry points: `pdf_to_markdown(pdf_bytes, *, page_range, mode, provider) -> str` and `pdf_to_markdown_with_status(...) -> (str, status_dict | None)` (the latter is what `app.py` uses — the UI shows `ocr_status`). Owns bytes→tempfile plumbing; no Streamlit, no caches. Defines the `OcrProvider` protocol (`ocr_pdf_to_markdown(pdf_bytes, page_range) -> str`) with one implementation, `OcrmypdfProvider`; `provider=None` routes through `ocr_auto_mode.convert_pdf_with_optional_ocr` unchanged. A second (cloud vision) provider is a planned separate PR.

OCR pipeline (mode `auto` in `Файлы -> Markdown`): `pdf_core.pdf_to_markdown_with_status` → `analyze_pdf_pages` (pypdf) → `ocr_auto_mode.convert_pdf_with_optional_ocr` → `ocr_converter.ocr_pdf_to_searchable_pdf` (`ocrmypdf --skip-text --deskew --rotate-pages -l rus+eng`) → `convert_with_markitdown` over the OCR text layer. Wired into the UI through `app.py` (`_convert_uploaded_file`).

## Brand constants (convert.py)

```python
BRAND_BLUE   = "015198"   # headings, accents
BRAND_RED    = "D04514"   # H2, decorative underline, signature rule
BRAND_ORANGE = "EF7F1A"   # blockquotes, photo placeholders
TEXT_DARK    = "1A1A1A"   # body text
```

Fonts: **PT Sans** (body, 12 pt), **PT Sans Narrow** (all headings).
Page margins: left 2 cm, right 1.5 cm → `CONTENT_WIDTH_CM = 17.5`.

## Block rendering map (convert.py)

| Markdown input | Renderer |
|---|---|
| `# …` | H1: PT Sans Narrow 18 pt BRAND_BLUE + red underline rule |
| `## …` | H2: PT Sans Narrow 14 pt BRAND_RED |
| `### …` | H3: PT Sans Narrow 11 pt TEXT_DARK bold |
| First `\n\n` block after H1 | `add_intro_paragraph` — left blue border accent |
| `> …` | Blockquote: orange left border, light grey fill, italic |
| `!! text !!` | `add_callout_box` — light blue fill table with border |
| `\| … \|` table | Styled table: BRAND_BLUE header row, zebra rows, 3-col gets BG_LIGHT_BLUE last column |
| `- / * / 1.` list | Bullet / numbered list, 1.5 cm indent |
| `**Кому:** …` | Requisites block: light blue fill |
| `С уважением` | Signature block: red top rule, kept together |
| `📷 / [Место для фото` | Photo placeholder: orange left border |
| `**Стадия/Фаза/Шаг/Этап/ВАЖНО` | Stage block: blue left border, light blue fill |
| `---` | Ignored (visual separator only) |

Table cells with «Да», «Нет», «Отсутствует» get automatic ✓/✗ icons.

## Numbered edits convention

`convert.py` uses numbered comments `# ПРАВКА #N: …` to mark deliberate changes. New edits are numbered strictly ascending and marked the same way.

**Known gap: `#25` does not exist in the code.** The file contains #1–#24, #26–#33. Column alignment from `:----` separators was never implemented — the separator row is simply filtered out. Do not assume README's edit list is accurate; verify against the code.

## Known issues

The P0 audit findings were fixed in the #28–#33 cycle (CRLF normalization; #26 vs requisites/stage blocks; callout spacer; numbered-list detection and restart; `![alt](src)` garbage hyperlinks; table cell split/padding). The audit battery lives in the session scratchpad (`run_audit*.py`) — rerun it after touching block dispatch, lists, tables, or the inline parser.

Documented long-standing limits: column alignment from `:----` separators is not implemented; list markers are capped at 2 digits (`^\d{1,2}\. `) so years like «2025.» are not eaten as list items; pseudo-headings without applied Word styles (mammoth cannot detect them); double-digit page numbers render vertically in LibreOffice.

## Constraints

- **`template.docx`** lives on Google Drive — do not add it to the repo and do not modify it.
- **`app.py`, `file_converter.py`, `requirements.txt`, `.devcontainer/*`** — edit only on explicit request.
- Code must run on both Windows (local) and Linux (prod).

## Known production limitation

OCR `auto` is implemented and wired into the UI, but `ocrmypdf` is **not** in `requirements.txt` and there is no `packages.txt`. On Streamlit Community Cloud the system Tesseract/Ghostscript binaries are unavailable, so OCR `auto` currently fails in production. Open question, resolved by a separate packaging task — do not add `ocrmypdf` to `requirements.txt` or create `packages.txt` as part of unrelated work.

## Secrets

Local dev requires `.streamlit/secrets.toml` (git-ignored):

```toml
[gcp_service_account]
type = "service_account"
project_id = "..."
private_key = "..."
client_email = "..."
# … remaining service account fields
```

Without this key, `use_drive` is `False` and the app falls back to the `local_path` in `DOC_TYPES`.

## Adding a document type

Add an entry to `DOC_TYPES` in `app.py`:

```python
"🔖 Имя типа": {
    "drive_id":    "<Google Drive file ID>",
    "local_path":  "<absolute local path to .docx template>",
    "output_name": "<filename stem>",
    "hint":        "<Markdown structure hint shown in the UI>",
},
```
