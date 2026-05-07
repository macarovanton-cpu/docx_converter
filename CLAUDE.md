# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the app

```bash
pip install -r requirements.txt
streamlit run app.py
```

The dev container auto-starts the app on port 8501 after attach (`postAttachCommand` in `.devcontainer/devcontainer.json`).

Run `convert.py` standalone for local batch testing:
```bash
python convert.py   # uses INPUT_FILE / OUTPUT_FILE / TEMPLATE_FILE at the top of the file
```

## Deployment

Push to `main` → Streamlit Cloud picks it up automatically. No CI step required.

## Architecture

Three Python modules:

- **`app.py`** — Streamlit UI. Downloads the `.docx` template from Google Drive (via service account in `st.secrets`), calls `convert_md_to_docx`, and serves the result as a file download. Falls back to a `local_path` if Drive credentials are absent. `DOC_TYPES` dict at the top controls available document types (label, Drive file ID, local fallback path, filename stem, Markdown hint).
- **`convert.py`** — Core Markdown → DOCX engine (~705 lines). Single public entry point: `convert_md_to_docx(md_text, output_filename, template_path=None)`. This signature must not change — `app.py` depends on it. Parses MD into blocks split on `\n\n`, dispatches each block to a typed renderer, and writes the result via `python-docx`.
- **`file_converter.py`** — Reverse direction: DOCX / PDF / TXT → Markdown. Entry point: `convert_file_to_md(file_bytes, filename) → (md_text, images)`. Used by `app.py` to pre-fill the editor when the user uploads an existing document.

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

`convert.py` uses numbered comments `# ПРАВКА #N: …` to mark deliberate changes to defaults. Current highest: **#11** (signature block page-break lock). New edits must be numbered starting at **#12** and marked the same way.

## Constraints

- **`template.docx`** lives on Google Drive — do not add it to the repo and do not modify it.
- **`app.py`, `file_converter.py`, `requirements.txt`, `.devcontainer/*`** — edit only on explicit request.
- **No new `requirements.txt` dependencies** without prior agreement.
- `convert_md_to_docx(md_text, output_filename, template_path=None)` signature is fixed.

## Secrets

Local dev requires `.streamlit/secrets.toml` (git-ignored):

```toml
[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
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
