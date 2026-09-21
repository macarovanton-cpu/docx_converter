# 12 — UI на едином входе `ingest`

**# ПРАВКА #85.** Зависит от: 09 (`ocr.ingest`), 08 (ретроспектива #73 — что именно сохраняется),
11 (CLAUDE.md уже сверен). `app.py` **разрешён** этой спекой. Сеть не нужна.

## Цель

В режиме `mineru` UI зовёт `ocr.ingest.ingest` вместо `ocr.cli.run_pipeline`. Следствие:
DOCX и XLSX проходят постпроцессор, валидатор и получают `report.json` так же, как PDF —
сейчас они в том же режиме идут голым MarkItDown без единой проверки (`app.py`, ветка `else`
в `_convert_uploaded_file`), хотя переключатель обещает «отчёт о сомнительных местах».

Сохраняется без изменений поведения: переключатель `ocr_mode` и его три значения, прогресс
через `st.status`, подмножество страниц (`_pdf_page_subset`), обработка `MineruError` /
`MineruAuthError`, блок находок, кнопка `report.json`. **Стиль UI не меняется**: те же
виджеты, та же раскладка; меняются только тексты подписей там, где они стали неправдой.

### Что изменится для человека (записать в docs_sync дословно)

1. DOCX/XLSX в режиме `mineru`: появляются находки и кнопка `report.json`; в облако они
   по-прежнему **не уходят** (маршрут `office`, `provider = "markitdown"`).
2. Текстовый PDF без таблиц в режиме `mineru` **перестаёт уходить в облако**: `detect_route`
   даёт `text`, конвертирует MarkItDown. Раньше в MinerU шёл любой PDF. Это правило спеки 09,
   UI его наследует, а не вводит своё.
3. Маршрут считается **по вырезке страниц**, а не по исходному файлу: диапазон «1-2» из
   смешанного PDF может дать `text`, хотя весь файл — `scan`. Это верно: в облако идёт вырезка.
4. PPTX — как раньше: MarkItDown, без отчёта (`ingest` его не принимает).

## Шаг 0 — условия входа (иначе стоп)

1. Спека 11 закоммичена, `pytest -q` зелёный.
2. `grep -n "run_pipeline" app.py` — ровно импорт, один вызов и упоминания в комментариях/докстрингах
   блока #73. Другие вызовы — стоп: спека писалась по коду, где вызов один.
3. `ingest` и `detect_route` имеют сигнатуры спеки 09 (`ocr/ingest.py`). Расхождение — стоп.

## Трогать

- `app.py` — импорт, `_convert_uploaded_file`, цикл конвертации и подпись режима в
  `render_files_to_markdown_mode`, строка маршрута в карточке результата (#85)
- `tests/test_app_fixes.py` — перевод mineru-тестов на настоящий PDF, новые тесты
- `CLAUDE.md` — три фразы: описание `app.py`, `ocr/cli.py` («Stage 8's UI calls `run_pipeline` unchanged»),
  `ocr/ingest.py` («`app.py` still calls `run_pipeline`»)
- `docs/specs/ocr/README.md` — строка таблицы спек; `docs/specs/ocr/08-ui-mineru.md` — одна строка-ссылка
  «вызов тракта заменён спекой 12», текст ретроспективы не переписывать
- `docs/docx_converter_docs_sync.md` — `### ПРАВКА #85` (четыре пункта выше + схема вызова), снимок «следующий — #86»

## Не трогать

`ocr/*` целиком (**`ingest`, `run_pipeline`, `detect_route` не меняются**: нужен новый параметр —
стоп), `pdf_core.py`, `file_converter.py`, `requirements.txt`, `convert.py`, режим «MD → DOCX»,
ветки `off` и `auto`, `_pdf_page_subset`, `_mineru_api_key`, `_mineru_provider_factory` (имя,
сигнатура, швы прогресса — на них стоят тесты #73), `_render_ocr_report`, `_split_findings`,
`_findings_table`, схема `report.json`, ключи результата, кроме одного добавленного.

## Интерфейсы (дословно)

### `run_pipeline` остаётся публичной

Решение: **да**. Её зовёт сам `ingest` (маршруты `scan`, `text_tables` и `engine="ocrmypdf"`),
на ней стоят `tests/test_ocr_cli.py` и сравнение `ingest == run_pipeline` в `tests/test_ocr_ingest.py`,
ею может пользоваться агент. Меняется одно: **UI её больше не зовёт и не импортирует**.
`from ocr.cli import run_pipeline` из `app.py` удаляется; докстринг `run_pipeline`
(«UI этапа 8 зовёт ту же функцию») не правится — `ocr/cli.py` вне «Трогать», неточность
записывается в docs_sync.

### `app.py` (#85)

```python
from ocr.ingest import detect_route, ingest

_INGEST_EXTS = ("pdf", "docx", "xlsx")      # ПРАВКА #85: что в режиме mineru идёт через ocr.ingest
_MINERU_ROUTES = ("scan", "text_tables")    # маршруты, на которых есть облако и второй прогон

_ROUTE_LABELS = {
    "scan": "скан → MinerU (облако)",
    "text_tables": "текстовый PDF с таблицами → MinerU (облако)",
    "text": "текстовый PDF без таблиц → MarkItDown (без облака)",
    "office": "DOCX/XLSX → MarkItDown (без облака)",
}
```

`_convert_uploaded_file` — сигнатура прежняя. Ветка `mineru`:

```python
if ocr_mode == "mineru" and ext in _INGEST_EXTS:
    data = uploaded_file.getvalue()
    if ext == "pdf":
        data = _pdf_page_subset(data, page_range)
    # ponytail: detect_route зовётся дважды (здесь и внутри ingest) — второй проход
    # pdfplumber по страницам. Убирается только параметром route у ingest, а он заморожен спекой 09.
    route = detect_route(data, uploaded_file.name)
    with tempfile.TemporaryDirectory() as work_dir:
        markdown, report = ingest(
            data, source_name=uploaded_file.name, work_dir=Path(work_dir),
            verify=verify and route in _MINERU_ROUTES, annotate=annotate,
            cache=LocalCache(_OCR_CACHE_ROOT),
            provider_factory=_mineru_provider_factory(_mineru_api_key(), status))
```

- Ветки `elif ext == "pdf"` (режимы `off`/`auto`) и `else` (MarkItDown) — без изменений; в `else`
  теперь попадают PPTX в любом режиме и DOCX/XLSX вне `mineru`.
- Результат получает **один новый ключ** `route`: строка из `ROUTES` либо `None` (вне ветки
  `mineru` и при ошибке до определения маршрута). Остальные восемь ключей и их смысл прежние.
  `_build_markdown_zip` и `_build_combined_markdown` новый ключ не читают.
- `verify` на маршрутах `text`/`office` **не передаётся**: `ingest` на нём бросает `ValueError`
  с текстом про CLI-флаг, а человек ставил галочку на пачку файлов. Это не молчаливое гашение:
  в карточке результата при `verify and route not in _MINERU_ROUTES` стоит
  `st.caption("Сверка вторым прогоном не применялась: файл не шёл через MinerU.")`, а
  `report["verified"]` честно `False`. Для этого в результат рядом с `route` ничего не добавляется —
  `verify` UI знает из чекбокса (`st.session_state.get("files_to_md_mineru_verify")` — `.get`, потому что
  вне режима `mineru` чекбокс не рисуется и ключа может не быть).
- Ошибки: блок `except` не меняется. `ValueError` из `detect_route` («Неподдерживаемый формат»)
  недостижим — расширение отфильтровано `_INGEST_EXTS`; битый PDF даёт исключение `pypdf`,
  оно показывается текстом, как любое другое.
- `status` (метки этапов) на маршрутах MarkItDown не обновляется провайдером — провайдер не
  создаётся. Итоговую метку «готово»/«ошибка» ставит цикл, как сейчас.

`render_files_to_markdown_mode`:

- условие `st.status`: `ocr_mode == "mineru" and _file_ext(name) == "pdf"` →
  `ocr_mode == "mineru" and _file_ext(name) in _INGEST_EXTS`; заголовок статуса
  `f"MinerU: {name}…"` → `f"OCR-тракт: {name}…"` (и две метки исхода) — для DOCX слово «MinerU» было бы неправдой;
- подпись режима (`st.caption` в ветке `else` переключателя) — новый текст:
  «В облако mineru.net уходят сканы и текстовые PDF с таблицами (до 200 МБ и 200 стр.). Текстовые PDF
  без таблиц, DOCX и XLSX конвертируются локально, но проходят ту же проверку и получают отчёт.
  PPTX — как обычно, без отчёта. При заданном диапазоне отправляются только выбранные страницы, и
  номера страниц в находках считаются от этой вырезки.»;
- `help` переключателя: «MinerU: облачное распознавание PDF с отчётом…» → «MinerU: распознавание и
  проверка PDF/DOCX/XLSX с отчётом о сомнительных местах.» Ярлык `_OCR_MODE_LABELS["mineru"]` не меняется;
- карточка результата: при `result.get("route")` — `st.caption(f"Маршрут: {_ROUTE_LABELS[route]}")`
  перед блоком находок; ниже — подпись про несостоявшуюся сверку (см. выше). `.get` здесь допустим:
  в `st.session_state` могут лежать результаты прошлой версии без ключа.

## Приёмочные тесты

`tests/test_app_fixes.py`. Сеть заглушена фикстурой `no_network` (как в `tests/test_ocr_golden.py`),
область — только тесты блока #85/#73. Кэш — `tmp_path` (`_fake_mineru` уже подменяет `_OCR_CACHE_ROOT`).

### Перевод существующих

Тесты #73 шлют `_FakeUpload("скан.pdf")` с байтами `b"fake"` — `detect_route` на них упадёт в `pypdf`.
Во всех четырёх (`…runs_pipeline_with_fake_provider`, `…verify_runs_second_model`, `…auth_error…`,
`…error_shown_as_text…`) вход → `_FakeUpload("скан.pdf", _blank_pdf(1))`: пустая страница без
текстового слоя = маршрут `scan`. Assert-ы не меняются; первый тест переименовать в
`test_mineru_mode_runs_ingest_with_fake_provider` и добавить `assert result["route"] == "scan"`.
`_blank_pdf` поднять выше по файлу — тело не менять.

### Новые

```python
def _docx_bytes() -> bytes:          # python-docx: заголовок, абзац, таблица 2x2
def _xlsx_bytes() -> bytes:          # openpyxl (стоит с markitdown[all]): лист 3x2 с числами


@pytest.mark.parametrize("name, data", [("док.docx", _docx_bytes()), ("табл.xlsx", _xlsx_bytes())])
def test_mineru_mode_office_goes_through_ingest(monkeypatch, tmp_path, name, data):
    calls = _fake_mineru(monkeypatch, tmp_path, vlm="не должен понадобиться")
    result = app._convert_uploaded_file(_FakeUpload(name, data), "1-3", ocr_mode="mineru", verify=True)
    assert result["error"] is None and result["markdown"].strip()
    assert result["route"] == "office" and result["page_range"] == "all"
    assert result["report"]["provider"] == "markitdown"
    assert result["report"]["verified"] is False               # verify не передан, не ValueError
    assert list(result["report"]) == REPORT_KEYS                # те же десять ключей, что у PDF
    assert calls == []                                          # провайдер не создавался


def test_mineru_mode_text_pdf_stays_local(monkeypatch, tmp_path):
    pdf = require_fixture("textpdf1.pdf").read_bytes()          # стр. 1-2: слой есть, таблиц нет (спека 09)
    calls = _fake_mineru(monkeypatch, tmp_path, vlm="x")
    result = app._convert_uploaded_file(_FakeUpload("т.pdf", pdf), "1-2", ocr_mode="mineru")
    assert result["route"] == "text" and calls == []            # маршрут — по вырезке
    assert result["report"]["provider"] == "markitdown"


def test_mineru_mode_pptx_keeps_plain_markitdown(monkeypatch):
    monkeypatch.setattr(app, "convert_with_markitdown", lambda path, page_range=None: "# Слайд")
    result = app._convert_uploaded_file(_FakeUpload("през.pptx"), None, ocr_mode="mineru")
    assert result["markdown"] == "# Слайд" and result["report"] is None and result["route"] is None


def test_office_outside_mineru_mode_unchanged(monkeypatch):     # off: отчёта нет, как было
    monkeypatch.setattr(app, "convert_with_markitdown", lambda path, page_range=None: "# Д")
    result = app._convert_uploaded_file(_FakeUpload("д.docx"), None, ocr_mode="off")
    assert result["report"] is None and result["route"] is None


def test_result_keys_closed():
    assert set(result) == {"filename", "download_name", "file_type", "page_range", "ocr_status",
                           "markdown", "report", "route", "error"}      # и на успехе, и на ошибке


def test_app_no_longer_imports_run_pipeline():
    assert not hasattr(app, "run_pipeline")
```

`REPORT_KEYS` — список из `tests/test_ocr_ingest.py` (`KEYS`), дословно. `test_page_range_ignored_for_non_pdf`
(режим `off`) обязан остаться зелёным без правок.

### Руками (в отчёт прогона, не в тест)

`streamlit run app.py`, режим MinerU, три файла разом — `bakeoff.pdf` (из кэша), `docx1.docx`, `xlsx1.xlsx`:
у всех трёх — строка маршрута, сводка находок и кнопка `report.json`; сводки совпадают со строками табло
(`5/7/2`, `0/1/0`, `0/4/0` на 21.09.2026). Раскладка не поехала. Если `bakeoff.pdf` в рабочем кэше нет —
PDF пропустить и написать об этом, в облако ради проверки не ходить.

## Готово, когда

- `pytest -v` зелёный целиком; `tests/test_ocr_cli.py`, `tests/test_ocr_ingest.py`, `tests/test_ocr_board.py` не менялись.
- `grep -n "run_pipeline" app.py` — только в комментариях/докстрингах, импорта и вызова нет.
- `git diff --stat`: `app.py`, `tests/test_app_fixes.py`, документы. Ни одного файла из `ocr/`.
- Маркер `# ПРАВКА #85` стоит у каждого изменённого места `app.py`; маркеры #73 не удалены.
- В docs_sync — `### ПРАВКА #85`: схема вызова, четыре изменения поведения, запись о неточном докстринге
  `run_pipeline`, снимок «следующий свободный номер — **#86**».

## PLACEHOLDER-ы

1. `detect_route` вызывается дважды на файл (UI и `ingest`). На 200-страничном текстовом PDF это второй
   полный проход `pdfplumber.find_tables`. Замерить при ручной проверке; заметно — отдельная правка
   `ingest` (вернуть маршрут), не обход в `app.py`.
2. Живой прогон MinerU из UI по-прежнему не проверялся (как и на #73) — только кэш и фейковый провайдер.
3. `pdf_has_tables` не видит безрамочных таблиц (PLACEHOLDER 1 спеки 09): такой PDF в режиме `mineru`
   теперь молча останется локальным. Строка маршрута в карточке — единственное, что это показывает человеку.
4. Тексты подписей — рабочие; человек вправе переписать их без новой спеки.

## Коммит

```
ПРАВКА #85: UI зовёт ocr.ingest — DOCX и XLSX получают проверку и report.json

app.py: в режиме MinerU PDF, DOCX и XLSX идут через ocr.ingest.ingest вместо
run_pipeline: постпроцессор, валидатор и report.json одинаковы для всех трёх
форматов. Текстовый PDF без таблиц остаётся локальным (маршрут text), маршрут
считается по вырезке страниц и показан в карточке результата; сверка вторым
прогоном передаётся только на маршрутах MinerU. Переключатель, st.status,
_pdf_page_subset, обработка MineruError и кнопка report.json — прежние.
run_pipeline остаётся публичной (её зовёт ingest и CLI), UI её не импортирует.
tests/test_app_fixes.py: фейковый провайдер, без сети; пути DOCX, XLSX,
текстового PDF и PPTX.
```
