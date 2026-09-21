# 08 — MinerU в UI (ретроспектива)

**# ПРАВКА #73.** Ретроспектива: написана 2026-09 по коду app.py, после исполнения.
Этап 8 выполнен по явному запросу человека без спеки; этот файл фиксирует факт, а не задаёт работу.

Номера строк `app.py` — на 21.09.2026 (коммит `fdcf87e`); якорь — имя функции, не номер.

## Цель (что было сделано)

В режим «Файлы → Markdown» добавлен третий вариант OCR — MinerU (облако). PDF уходил в
`ocr.cli.run_pipeline` без изменения тракта; под результатом показывались находки и
отдавался `report.json`. Сигнатура `run_pipeline` (спека 07) осталась нетронутой: всё,
чего ей не хватало для UI (диапазон страниц, прогресс), было решено на стороне `app.py`.

## Условия входа (какими они были)

Спеки 00–07 исполнены (#60–#66), исправления тракта #67–#72 закоммичены. `app.py`
правился по явному запросу человека — общий запрет автономного прогона на этот файл
тем самым был снят для одной правки.

## Трогали

- `app.py` — блок «ПРАВКА #73» (строки 241–420), выбор режима и чекбоксы в
  `render_files_to_markdown_mode` (631–673), `st.status` в цикле конвертации (787–800),
  блок находок и кнопка `report.json` в цикле результатов (872–898)
- `tests/test_app_fixes.py` — блок «ПРАВКА #73», 11 функций

## Не трогали

`ocr/*` целиком (в том числе сигнатуру `run_pipeline`), `pdf_core.py`, `file_converter.py`,
`requirements.txt`. `drive_client.py` не создавался, Drive-код из `app.py` не выносился,
Drive-кэш не делался (`make_cache("drive")` по-прежнему `NotImplementedError`).

## Интерфейсы (по факту)

| Что | Где (`app.py`) |
|---|---|
| `ocr_mode`: `off` / `auto` / `mineru`. `_ocr_mode_options(ocrmypdf_ok)` возвращал `["off"] + (["auto"] если бинарники найдены) + ["mineru"]`; проверка бинарников — `_ocrmypdf_available()` под `st.cache_resource` (три `--version` один раз на процесс); подписи — `_OCR_MODE_LABELS` («Без OCR», «OCRmyPDF (локально)», «MinerU (облако)») | `_ocr_mode_options` (257), `_ocrmypdf_available` (263), `_OCR_MODE_LABELS` (245) |
| ключ: `st.secrets["MINERU_API_KEY"]`, затем `os.environ.get("MINERU_API_KEY")`; `FileNotFoundError` (нет `secrets.toml`) и `KeyError` (нет ключа) → окружение; нет нигде — `None` | `_mineru_api_key` (268) |
| диапазон страниц: `run_pipeline` диапазона не принимал, выбранные страницы вырезались `pypdf` (`PdfReader` → `PdfWriter`) до отправки; `page_range=None` — байты как есть; страница вне PDF → `ValueError("Страница N вне диапазона PDF: в файле всего M стр.")`. Следствие: номера страниц в находках считались от вырезки, не от исходного файла (об этом говорила подпись в UI) | `_pdf_page_subset` (276) |
| прогресс без колбэка в тракте: `_mineru_provider_factory(api_key, status)` возвращал `build(model_version)`; внутри — подкласс `_StatusProvider(MineruProvider)`, который метил `fetch_raw_zip` («загрузка файла…» до, «Постобработка и проверка…» после), и инжектируемый `sleep`, ставивший метку «ожидание результата…». `status=None` допустим — метки тогда не ставились | `_mineru_provider_factory` (299) |
| вызов: `run_pipeline(pdf_bytes, source_name=uploaded_file.name, work_dir=Path(<tmp>), verify=verify, annotate=annotate, cache=LocalCache(_OCR_CACHE_ROOT), provider_factory=_mineru_provider_factory(_mineru_api_key(), status))` — только при `ext == "pdf" and ocr_mode == "mineru"`; остальные PDF шли в `pdf_to_markdown_with_status`, не-PDF — в `convert_with_markitdown` | `_convert_uploaded_file` (364) |
| ошибки: любое исключение → `result["error"] = str(e) or type(e).__name__`; у `MineruAuthError` к тексту добавлялся `_MINERU_KEY_HINT`; трейсбека в UI не было | там же (402–405), `_MINERU_KEY_HINT` (251) |
| результат: в словаре результата появился ключ `report` — отчёт `run_pipeline`; `None` вне режима `mineru` и при ошибке | там же (399, 413) |
| UI: чекбоксы «Сверка вторым прогоном» (`value=False`) и «Пометки в тексте» (`value=True`) — только в режиме `mineru`; предупреждение `st.warning`, если ключ не найден; `st.status` на каждый PDF с итоговой меткой «готово» / «ошибка» | `render_files_to_markdown_mode` (616; 646–673, 787–800) |
| находки: `_split_findings` делил отчёт на обычные и `low_confidence`; `_render_ocr_report` показывал сводку (`st.error` при critical, иначе `st.info`, с пометкой «результат из кэша»), таблицу находок (раскрыта при critical) и `low_confidence` отдельным свёрнутым блоком; при `ANNOTATION_PREFIX` в тексте — подпись, что пометки «!! ПРОВЕРИТЬ: … !!» надо снять перед DOCX | `_split_findings` (328), `_findings_table` (335), `_render_ocr_report` (341), цикл результатов (872–877) |
| `report.json` — отдельная кнопка `<stem>.report.json` рядом с «Скачать .md»; в общий ZIP и в объединённый Markdown отчёт не входил | цикл результатов (888–898) |
| кэш: `_OCR_CACHE_ROOT = <папка app.py>/.cache/ocr` — от расположения файла, не от cwd; на Streamlit Cloud эфемерный, осознанно | `_OCR_CACHE_ROOT` (243) |

## Приёмочные тесты (какими они были)

`tests/test_app_fixes.py`, блок «ПРАВКА #73» — 11 функций, фейковый провайдер, без сети:
`test_ocr_mode_options_hide_ocrmypdf_without_binaries`, `test_mineru_api_key_prefers_secrets`,
`test_mineru_api_key_falls_back_to_env`, `test_mineru_api_key_absent`,
`test_mineru_mode_runs_pipeline_with_fake_provider`, `test_mineru_mode_verify_runs_second_model`,
`test_mineru_auth_error_gets_key_hint`, `test_mineru_error_shown_as_text_without_hint`,
`test_pdf_page_subset_cuts_selected_pages`, `test_pdf_page_subset_rejects_page_outside_pdf`,
`test_split_findings_separates_low_confidence`.

## Известные ограничения

- DOCX/XLSX в режиме `mineru` шли голым MarkItDown без `postprocess`/`validate` и без отчёта
  (чинит спека 12).
- Живой прогон MinerU из UI правкой #73 не проверялся: тесты — на фейковом провайдере.
- `sleep` зовётся и в ретраях загрузки — метка «ожидание» может мигнуть раньше времени
  (комментарий `# ponytail:` в `_mineru_provider_factory`). Косметика; точнее — только
  колбэком внутри `ocr/`.
- Номера страниц в находках при заданном диапазоне — от вырезки, не от исходного PDF.

## PLACEHOLDER-ы

1. Номера строк `app.py` — на день написания; якорь — имя функции.

## Коммит

`b5f5849` — «ПРАВКА #73: MinerU в режиме «Файлы → Markdown» — run_pipeline из UI, находки, report.json»
