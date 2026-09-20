# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Working principles

These are the rules of engagement. Read them before touching code.

**Plan first, then execute.** Start non-trivial work in plan mode. Iterate on the plan until it is right, then switch to auto-accept and let the implementation run in one go. A good plan is the highest-leverage artifact in the session — a bad plan produces 40 changes nobody asked for.

**Verify your own work.** Never hand back code you have not checked. Run `pytest -v`. Generate a real `.docx` and inspect it programmatically with `python-docx` (styles, alignment, breaks) — do not assume the output is correct because the code looks correct. Self-verification is worth more than confidence.

**Every mistake becomes a rule.** When a bug or a wrong assumption is found, do not just fix the code — add the rule to this file so it is not repeated. This file is the project's memory across sessions. Keep it under ~200 lines so it is actually read.

**Small, independent changes.** One `# ПРАВКА #N` = one problem = one commit with a descriptive message. Do not bundle unrelated fixes. Do not opportunistically refactor code you were not asked to touch.

**The human stays in the review seat.** Architectural decisions, brand rules, and what ships to clients are the human's call. Ask one clarifying question when the spec is ambiguous — do not guess.

**AI-written code is statistically dirty.** Plausible-looking code that passes tests can still be conceptually wrong: redundant loops, silently dropped data, edge cases that never fire. Bugs here are conceptual, not syntactic. Assume this about your own output and look for it.

## Verifying results

- Any check done programmatically gets turned into a test in `tests/`. A one-off script in a session scratch folder is lost the moment the session ends — it verified nothing for the next session.
- Parsing `word/document.xml` verifies document *structure* (styles applied, elements present, correct order). It does not verify *appearance*. Only a human opening the file in real Word can confirm it looks right.
- An agent's report of what it did is not evidence. `git log` and `git status` are.

## Forbidden patterns

- Do **not** change the signature `convert_md_to_docx(md_text, output_filename, template_path=None, images=None, doc_style='pz')` — `app.py` depends on it. Extend with optional parameters only.
- Do **not** read profile keys with `style.get(...)` — only `style['key']`. `.get` returns `None`, and `Pt(None)` fails far from the typo.
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
- MD→DOCX needs a template: without a `[gcp_service_account]` section in `.streamlit/secrets.toml`, Drive is unavailable and the app falls back to `C:\Users\tonik\Desktop\docx_converter\template.docx` — that file must exist, otherwise «Шаблон не найден».
- Бланка письма локально нет: `template_letter.docx` рядом с `app.py` отсутствует, Drive выключен. Без него письмо проверяется только по ветке без шаблона — полей и колонтитула бланка в таком выводе не будет.

The dev container auto-starts the app on port 8501 after attach (`postAttachCommand` in `.devcontainer/devcontainer.json`).

Run `convert.py` standalone for local batch testing:
```bash
python convert.py   # uses INPUT_FILE / OUTPUT_FILE / TEMPLATE_FILE at the top of the file
```

## Deployment

Push to `main` → Streamlit Cloud picks it up automatically. No CI step required. **Claude Code does not push — the human pushes.**

## Architecture

Seven Python modules plus the `ocr/` package:

- **`app.py`** — Streamlit UI. Downloads the `.docx` template from Google Drive (via service account in `st.secrets`), calls `convert_md_to_docx`, and serves the result as a file download. Falls back to a `local_path` if Drive credentials are absent. `DOC_TYPES` dict at the top controls available document types.
- **`convert.py`** — Core Markdown → DOCX engine (1316 lines). Single public entry point: `convert_md_to_docx(md_text, output_filename, template_path=None, images=None)`. Parses MD into blocks split on `\n\n`, dispatches each block to a typed renderer, writes via `python-docx`.
- **`file_converter.py`** — Reverse direction: DOCX / PDF / TXT → Markdown. Entry point: `convert_file_to_md(file_bytes, filename) → (md_text, images)`. Also hosts the `Файлы -> Markdown` MarkItDown layer (`convert_with_markitdown`) and PDF diagnostics (`analyze_pdf_pages`, via pypdf).
- **`markdown_cleanup.py`** — Deterministic OCR Markdown cleanup (`cleanup_ocr_markdown`). Covered by tests but **not connected to the UI or conversion flow** — backend-only.
- **`ocr_auto_mode.py`** — «OCR or not» orchestrator (`convert_pdf_with_optional_ocr`). Decides whether a PDF needs OCR, honoring the selected page range.
- **`ocr_converter.py`** — OCRmyPDF wrapper via `subprocess`. Both runs are bounded: `OCR_TIMEOUT_SEC = 300` for ocrmypdf, `DEPENDENCY_TIMEOUT_SEC = 15` for `--version` probes. `TimeoutExpired` surfaces as a readable message, never a traceback (no manual `kill()` — `subprocess.run` already kills the child before raising). Also provides `check_ocr_dependencies`, which locates Ghostscript via `shutil.which` over platform candidates (`gswin64c` / `gswin32c` on Windows, `gs` elsewhere) — never a hardcoded name.
- **`pdf_core.py`** — Provider-agnostic PDF → Markdown core. Public entry points: `pdf_to_markdown(pdf_bytes, *, page_range, mode, provider) -> str` and `pdf_to_markdown_with_status(...) -> (str, status_dict | None)` (the latter is what `app.py` uses — the UI shows `ocr_status`). Owns bytes→tempfile plumbing; no Streamlit, no caches. Defines the `OcrProvider` protocol (`ocr_pdf_to_markdown(pdf_bytes, page_range) -> str`) with one implementation, `OcrmypdfProvider`; `provider=None` routes through `ocr_auto_mode.convert_pdf_with_optional_ocr` unchanged. A second (cloud vision) provider is a planned separate PR.
- **`ocr/`** — MinerU OCR pipeline (`ocr/__init__.py` holds the shared `Finding` dataclass and `SEVERITIES`). Eight modules:
  - **`ocr/mineru_provider.py`** (#61) — MinerU cloud API v4 behind the `pdf_core.OcrProvider` protocol: PDF → raw zip → `OcrResult`. `result_from_zip` unpacks a zip without touching the network — that is what the cache reuses.
  - **`ocr/cache.py`** (#62) — raw provider response cache (zip + `meta.json`): `cache_key`, `build_meta`, `LocalCache`, `make_cache`. `make_cache("drive")` raises `NotImplementedError` until stage 8.
  - **`ocr/postprocess.py`** (#63) — deterministic Markdown cleanup after OCR: HTML tables → pipe tables, page-split tables merged, homoglyphs, `No` → `№`; everything doubtful becomes a `Finding`, nothing is silently fixed.
  - **`ocr/validate.py`** (#64) — checks over the postprocessed Markdown (ИНН/ОГРН/КПП, ГОСТ, units, table totals) plus `build_report` (`report.json` schema v1), `annotate` and `strip_annotations`.
  - **`ocr/diff.py`** (#65) — word-level comparison of two runs (`vlm` vs `pipeline`); every divergence becomes a `low_confidence` finding. Text is never changed, no side is declared right.
  - **`ocr/cli.py`** (#66) — `python -m ocr.cli`: `run_pipeline` wires the whole tract (cache → provider → postprocess → validate → optional diff → report), `main` writes `out.md` / `report.json` and prints one JSON line. Stage 8's UI calls `run_pipeline` unchanged.
  - **`ocr/ingest.py`** (#81) — single entry for `.pdf` / `.docx` / `.xlsx`: `detect_route` (scan / text with tables / text / office), MinerU routes go to `run_pipeline` unchanged, MarkItDown routes go through the same `postprocess` + `validate` + `build_report`. `ocr.cli.main` calls `ingest`; `app.py` still calls `run_pipeline`.
  - **`ocr/board.py`** (#82) — `python -m ocr.board`: offline quality board over the six fixtures (findings by rule, table integrity, size; `count_diffs` is `null` until spec 10). Never overwrites an existing `<stem>.errors.txt`.

OCR pipeline (mode `auto` in `Файлы -> Markdown`): `pdf_core.pdf_to_markdown_with_status` → `analyze_pdf_pages` (pypdf) → `ocr_auto_mode.convert_pdf_with_optional_ocr` → `ocr_converter.ocr_pdf_to_searchable_pdf` (`ocrmypdf --skip-text --deskew --rotate-pages -l rus+eng`) → `convert_with_markitdown` over the OCR text layer. Wired into the UI through `app.py` (`_convert_uploaded_file`).

## Brand constants (convert.py)

```python
BRAND_BLUE   = "015198"   # headings, accents
BRAND_RED    = "D04514"   # H2, decorative underline, signature rule
BRAND_ORANGE = "EF7F1A"   # blockquotes, photo placeholders
TEXT_DARK    = "1A1A1A"   # body text
```

Page margins: left 2 cm, right 1.5 cm → `CONTENT_WIDTH_CM = 17.5`.

## Style profiles (ПРАВКА #55)

Оформление задают два плоских словаря, `STYLE_PZ` и `STYLE_LETTER` (46 ключей:
шрифты и кегли, цвета, флаги декора). Профиль выбирается аргументом
`doc_style='pz'|'letter'` и доезжает до функций **явным параметром `style`**, а
не модульной переменной: Streamlit обслуживает сессии потоками одного процесса,
и глобал при двух одновременных конвертациях разных типов молча отдал бы
клиенту письмо в стиле ПЗ. Неизвестный `doc_style` — `KeyError`, не фолбэк.

Профиль получают шесть функций (`add_intro_paragraph`, `add_callout_box`,
`add_table_cell_content`, `ensure_list_numbering`, `add_hyperlink_run` через
`link_color`, `parse_inline_markdown` — только ради цвета ссылки) и сам
`convert_md_to_docx`. Дефолт `style=None` → `STYLE_PZ`.

- **ПЗ**: PT Sans 12 pt, PT Sans Narrow в заголовках, фирменные цвета.
- **Письмо**: строгий бланк, PT Sans 10.5 pt, только чёрный, маркер «—»,
  «ПИСЬМО» по центру, тема курсивом, адресат вправо, подпись с табуляцией.

Новый ключ добавляется **в оба словаря** — их наборы сверяются на импорте
явным `raise` (не `assert`: тот вырезается под `python -O`).

Регрессию ПЗ ловит `tests/test_golden_pz.py` — посимвольное сравнение
`document.xml`, `numbering.xml`, `settings.xml` и стиля `Normal` с эталоном.
**Эталон перезаписывается только осознанно**, вместе с правкой, которая
намеренно меняет оформление ПЗ, и диффом в ревью:
`python tests/test_golden_pz.py --update`.

## Block rendering map (convert.py, профиль ПЗ)

| Markdown input | Renderer |
|---|---|
| `# …` | H1: PT Sans Narrow 18 pt BRAND_BLUE + red underline rule |
| `## …` | H2: PT Sans Narrow 14 pt BRAND_RED |
| `### …` | H3: PT Sans Narrow 13 pt TEXT_DARK bold |
| `#### … / ##### … / ###### …` | H4–H6: PT Sans Narrow 12 pt TEXT_DARK bold, no decorative lines (all three levels render identically) |
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

В профиле письма те же блоки рисуются иначе, плюс одна своя конструкция:
`**Дата:**` / `**Исх.:**` и первый `**Кому:**` выбираются **пре-проходом** до
основного цикла (`_pop_block`) и собираются в шапку — безрамочную таблицу 1×2
(`add_letter_header`). Порядок этих блоков в markdown значения не имеет, шапка
всегда идёт первой. Состояния «мету видели, ждём Кому» в цикле нет и заводить
его не надо.

## Numbered edits convention

`convert.py` uses numbered comments `# ПРАВКА #N: …` to mark deliberate changes. New edits are numbered strictly ascending and marked the same way. This flat, in-file numbering *is* the edit history — there is no separate changelog or list elsewhere, README included.

**Known gap: `#25` does not exist in the code, and never did.** The file contains #1–#24, #26–#58 (#52–#53 в `ocr_converter.py`, #54 и #59 в `app.py`). Дальше нумерация продолжается вне `convert.py`: #60 в `pdf_core.py`, #61–#66 в `ocr/` (по одной правке на модуль, см. список модулей выше), #81 — `ocr/ingest.py`, #82 — `ocr/board.py`. Do not assign #25 retroactively and do not treat its absence as something to "fix" — it is a permanently skipped number, not a missing edit to restore. Column alignment from `:----` separators was never implemented — the separator row is simply filtered out.

## Known issues

The P0 audit findings were fixed in the #28–#33 cycle (CRLF normalization; #26 vs requisites/stage blocks; callout spacer; numbered-list detection and restart; `![alt](src)` garbage hyperlinks; table cell split/padding). Rerun `pytest -v` after touching block dispatch, lists, tables, or the inline parser.

Documented long-standing limits: column alignment from `:----` separators is not implemented; list markers are capped at 2 digits (`^\d{1,2}\. `) so years like «2025.» are not eaten as list items; pseudo-headings without applied Word styles (mammoth cannot detect them); double-digit page numbers render vertically in LibreOffice.

## OCR CLI (инструмент для агента)

**Назначение:** тендерный PDF (в т.ч. скан) → Markdown со структурой + отчёт о сомнительных местах.
Смысл исходника не меняется: чинятся только известные артефакты OCR, остальное помечается.

**Вызов (из корня репозитория):**
`python -m ocr.cli ВХОД.(pdf|docx|xlsx) --out ПАПКА [--engine mineru|ocrmypdf] [--mode vlm|pipeline] [--verify] [--annotate] [--annotate-all] [--cache local|drive]`

**Вход (ПРАВКА #81):** `.pdf`, `.docx`, `.xlsx` — через `ocr.ingest.ingest`. PDF без текстового слоя хотя бы
на одной странице и текстовый PDF с таблицами идут в MinerU; текстовый PDF без таблиц, DOCX и XLSX —
в MarkItDown (без сети и ключа; `--verify` там — ошибка). Любой маршрут проходит `postprocess` + `validate`
и даёт тот же `report.json`.

**Табло качества (ПРАВКА #82):** `python -X utf8 -m ocr.board` — строго оффлайн, шесть фикстур из
`_test/fixtures/ocr/` (PDF — из `vlm_raw*.zip` через временный кэш, промах — ошибка, не облако).
Пишет `_test/quality_board.json`, черновики `_test/board/<имя>/out.md` + `report.json` и заготовки
`_test/fixtures/ocr/<имя>.errors.txt` (существующие не перезаписывает — в них ручной труд).
Код `0` — табло построено, `1` — исключение; критичные находки кода не меняют.

**Нужно:** переменная окружения `MINERU_API_KEY` (для `--engine mineru`). PDF до 200 МБ и 200 страниц.
Документ уходит в облако mineru.net; повторный прогон того же файла берётся из `.cache/ocr/`.

**Выход:** `ПАПКА/out.md`, `ПАПКА/report.json`, в stdout — одна строка JSON:
`{"status","out_md","report","cache_hit","findings":{"critical","warning","info"},"error"}`.

**Коды выхода:** `0` — критичных находок нет; `2` — есть критичные находки, `out.md` написан,
читать `report.json`; `1` — тракт не отработал, читать `error`.

**Флаги:** `--verify` — второй прогон другим движком MinerU, расхождения → находки `low_confidence`
(дольше, вдвое больше квоты). `--annotate` — находки вставлены в `out.md` как `!! ПРОВЕРИТЬ: … !!`;
перед конвертацией в DOCX снять через `ocr.validate.strip_annotations`. `low_confidence` в текст
не вставляются (их десятки, в `report.json` они есть все) — для полной картины `--annotate-all`.
Находки, чей фрагмент в тексте не нашёлся, уходят в конец файла, в раздел «Не привязанные находки».
`--cache drive` пока не реализован.

**report.json:** `findings[]` = `{id, rule, severity, page, snippet, suggestion}`.
`snippet` — дословный фрагмент `out.md`; `suggestion` — что предлагает тракт или что увидел второй прогон;
`page` — страница PDF или `null`.

## Constraints

- **`template.docx`** lives on Google Drive — do not add it to the repo and do not modify it.
- **`app.py`, `file_converter.py`, `requirements.txt`, `.devcontainer/*`** — edit only on explicit request.
- Code must run on both Windows (local) and Linux (prod).

## Known production limitation

OCR `auto` is implemented and wired into the UI, but `ocrmypdf` is **not** in `requirements.txt` and there is no `packages.txt`. On Streamlit Community Cloud the system Tesseract/Ghostscript binaries are unavailable, so OCR `auto` currently fails in production. Open question, resolved by a separate packaging task — do not add `ocrmypdf` to `requirements.txt` or create `packages.txt` as part of unrelated work. Since #53 the Ghostscript check is at least honest on Linux: if `gs` is on PATH it reports `ok`, instead of always failing on the Windows-only `gswin64c`.

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
    "style":       "pz",   # ключ из convert.STYLES — обязателен (#59)
    "hint":        "<Markdown structure hint shown in the UI>",
},
```

Новому оформлению нужен свой профиль в `convert.py`, а не `if` по типу
документа в коде рендера.
