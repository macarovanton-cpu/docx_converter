# 09 — Единый вход (ingest) и табло качества

**# ПРАВКА #81** (+ #82). Зависит от: 00–07. Шаг 0 номера не расходует. Сеть — только в шаге 0.
Номер 08 зарезервирован за ретроспективной спекой этапа 8 (MinerU в UI, ПРАВКА #73).

## Цель

Этап A плана OCR, первая половина: **измерение и единый выход**.

1. Один вход для `.pdf` / `.docx` / `.xlsx` — `ocr.ingest.ingest`. Какой бы конвертер
   ни отработал, результат проходит `postprocess` + `validate` и даёт один и тот же
   `report.json` (схема v1 спеки 05, без новых полей).
2. Табло качества `python -m ocr.board` — строго оффлайн, одна метрика на все форматы.
3. Продукт для человека: черновые `out.md` по шести фикстурам и пустые заготовки
   `<имя>.errors.txt`. Ошибки человек выписывает сам, в паузе между спеками 09 и 10.

Эталонов на этой спеке ещё нет: колонки `count_diffs` в табло пустые (`null`).

## Шаг 0 — ревизия фикстур (иначе стоп; единственное место, где разрешена сеть)

Сырые прогоны человек делал руками и мог ошибиться в именах и копировании.
Ожидаемое соответствие:

| Фикстура | Сырой выход |
|---|---|
| `bakeoff.pdf` | `vlm_raw.zip` |
| `bakeoff2.pdf` | `vlm_raw2.zip` |
| `bakeoff3.pdf` | `vlm_raw3.zip` |
| `textpdf1.pdf` | `vlm_raw_textpdf1.zip` |
| `docx1.docx`, `xlsx1.xlsx` | нет — родная конвертация, должны просто существовать |

Для каждого PDF из `_test/fixtures/ocr/`:

1. `sha = sha256(pdf)`. В `.cache/ocr/` есть запись `{sha}-mineru-vlm-pall-v1`
   (`.zip` + `.meta.json`), и `meta["sha256"] == sha`.
2. `vlm_raw*.zip` фикстуры **побайтно** равен zip этой записи кэша
   (сравнивать sha256 файлов, не размер).
3. Внутри zip есть `full.md` и ровно один член с суффиксом `content_list.json`
   (`content_list_v2.json` под суффикс не подходит — читать через
   `tests/ocr_fixtures.read_raw`).
4. `max(block["page_idx"]) + 1` по `content_list` равно
   `file_converter.get_pdf_page_count(pdf)`.

Исполнитель вправе:

- заменить неверно названный или чужой zip правильным из `.cache/ocr/`;
- прогнать отсутствующий (нет ни в фикстурах, ни в кэше):
  `python -X utf8 -m ocr.cli _test/fixtures/ocr/<файл> --out _test/out_<имя> --cache local`
  (`MINERU_API_KEY` в окружении), затем скопировать zip из `.cache/ocr/` под ожидаемым именем.

Больше сеть не нужна нигде. `bakeoff.pdf`/`vlm_raw.zip` при расхождении **не заменять**:
на нём стоят `golden.md` и тесты спек 04–07 — стоп, решает человек.

Итог шага — таблица «фикстура → zip → sha совпадает → страниц (PDF / content_list)»
в отчёте прогона и в `docs/docx_converter_docs_sync.md`, новый раздел
`## Фикстуры этапа A (<дата>)`. Требование: после шага 0 у каждой PDF-фикстуры есть
зафиксированный сырой выход, и дальше весь тракт оффлайн.

## Трогать

- `ocr/ingest.py` — создать (#81)
- `ocr/cli.py` — только `main` и `_parse_args` (#81)
- `ocr/board.py` — создать (#82)
- `tests/test_ocr_ingest.py`, `tests/test_ocr_board.py` — создать
- `CLAUDE.md`, `README.md` — раздел «OCR CLI» (вход `.pdf/.docx/.xlsx`, `python -m ocr.board`) и список модулей
- `docs/specs/ocr/README.md` — таблица спек (+09: #81, #82), таблица фикстур (шесть входов и четыре zip)
- `docs/docx_converter_docs_sync.md` — раздел шага 0, `### ПРАВКА #81`, `### ПРАВКА #82`, снимок «следующий свободный номер — #83»

## Не трогать

`app.py` (никак — он продолжает звать `run_pipeline`), `file_converter.py` (только
вызывается), `pdf_core.py` (`pdf_to_markdown_with_status` не меняется), `ocr_auto_mode.py`,
`ocr/cli.py::run_pipeline` и `_mineru_result`, остальные модули `ocr/*`, схема
`report.json`, имена и серьёзность правил, `requirements.txt` (новых зависимостей нет:
`pdfplumber`, `pypdf`, `markitdown[all]` уже стоят), `tests/test_ocr_cli.py` (обязан
остаться зелёным без правок), `convert.py`, `template*.docx`, `.devcontainer/*`,
`.github/*`, всё из общего списка запретов README.

## Интерфейсы (дословно)

### `ocr/ingest.py`

```python
"""ПРАВКА #81: единый вход тракта: PDF/DOCX/XLSX -> (markdown, report)."""

ROUTES = ("scan", "text_tables", "text", "office")
OFFICE_EXTS = (".docx", ".xlsx")
MIN_TABLE_ROWS, MIN_TABLE_COLS = 2, 2


def pdf_has_tables(pdf_path: str) -> bool: ...


def detect_route(data: bytes, filename: str) -> str: ...


def ingest(data: bytes, *, source_name: str, work_dir: Path,
           engine: str = "mineru", mode: str = "vlm",
           verify: bool = False, annotate: bool = False,
           annotate_all: bool = False,
           cache: "CacheBackend | None" = None,
           provider_factory=None) -> tuple[str, dict]: ...
```

Параметры `ingest` после `work_dir` — те же и с тем же смыслом, что у `run_pipeline`
(спека 07); на маршрутах MinerU они передаются в него без изменений.

### Маршруты

| Маршрут | Условие | Конвертер | `provider` / `model_version` в отчёте |
|---|---|---|---|
| `scan` | `.pdf`, хотя бы одна страница без текстового слоя | `run_pipeline(engine="mineru")` | `mineru` / `mode` |
| `text_tables` | `.pdf`, слой есть на всех страницах, `pdf_has_tables` | `run_pipeline(engine="mineru")` — тот же вызов | `mineru` / `mode` |
| `text` | `.pdf`, слой везде, таблиц нет | `pdf_core.pdf_to_markdown_with_status(data, mode="auto")` | `markitdown` / `None` |
| `office` | `.docx`, `.xlsx` | `file_converter.convert_with_markitdown(tmp_path)` | `markitdown` / `None` |

- `is_ocr=false` для `text_tables` отдельным параметром **не задаётся**: `MineruProvider.fetch_raw_zip`
  уже считает `is_ocr = any(not page.has_text_layer …)` — у текстового PDF это `False`.
  Поэтому `scan` и `text_tables` — один вызов, ключ кэша один (`…-mineru-vlm-pall-v1`),
  провайдер не правится. Смешанный PDF (часть страниц без слоя) — `scan`.
- `detect_route`: расширение по `filename` (регистр не важен); не `.pdf` и не из
  `OFFICE_EXTS` → `ValueError("Неподдерживаемый формат: …")`. Для PDF байты пишутся во
  временный файл, `file_converter.analyze_pdf_pages` → `ocr_auto_mode.pdf_pages_without_text_layer(pages)`:
  непусто → `scan`; иначе `pdf_has_tables` → `text_tables` / `text`. Временный файл удаляется в `finally`.
- `engine="ocrmypdf"` — явный выбор человека: детектор не зовётся, PDF идёт прямо в
  `run_pipeline` (все его `ValueError`, включая `--verify`, остаются как были).
- `verify=True` на маршрутах `text` и `office` → `ValueError("--verify доступен только для маршрутов MinerU")`:
  второго прогона там нет, молча игнорировать нельзя.
- Маршруты `text` и `office` после конвертера:
  `md, findings = postprocess(md, None)`; `findings += validate(md, None)`;
  `build_report(source=source_name, sha256=sha256(data), provider="markitdown", model_version=None,
  cache_hit=False, verified=False, findings=findings, content_list=None)`; затем `annotate`
  ровно как в `run_pipeline` (ПРАВКА #70). `page` у находок — `None`: страниц нет.
- `office`: временный файл с родным расширением (по нему `convert_with_markitdown`
  выбирает конвертер), `page_range` не передаётся, файл удаляется в `finally`.
  Исключения конвертера не глотать.

### Детектор таблиц — почему новый

Скан/текст различает существующий код (`analyze_pdf_pages` + `pdf_pages_without_text_layer`) —
он пригоден и используется как есть. Детектора **таблиц** в проекте нет вообще:
`extract_tables` / `find_tables` не зовётся нигде, `ocr_auto_mode` решает только «OCR или нет».
Новый — минимальный, на уже установленном `pdfplumber`:

```python
def pdf_has_tables(pdf_path: str) -> bool:
    """True, если хоть на одной странице page.find_tables() нашёл таблицу
    не меньше MIN_TABLE_ROWS x MIN_TABLE_COLS. Выход на первой найденной."""
```

Стратегия `find_tables` по умолчанию (линии). PLACEHOLDER 1 — см. ниже.

### `ocr/cli.py` (#81)

```
python -m ocr.cli INPUT.(pdf|docx|xlsx) --out DIR [те же флаги, что в спеке 07]
```

Меняется только `main`: вместо `run_pipeline(pdf_bytes, …)` — `ingest(data, …)` с теми же
аргументами; help позиционного аргумента — `входной PDF/DOCX/XLSX`. Строка JSON в stdout
(шесть ключей), коды `0 / 2 / 1`, запись `out.md`/`report.json` только после успеха —
**без изменений**. `run_pipeline` остаётся публичной: её зовёт `app.py`.

### `ocr/board.py` (#82)

```python
"""ПРАВКА #82: табло качества по фикстурам. Строго оффлайн."""

BOARD_SCHEMA_VERSION = 1
FIXTURES = Path(__file__).resolve().parents[1] / "_test" / "fixtures" / "ocr"
BOARD_JSON = FIXTURES.parents[1] / "quality_board.json"      # _test/quality_board.json
DRAFTS = FIXTURES.parents[1] / "board"                       # _test/board/<stem>/
ERRORS_HEADER = "# страница | было | надо\n"

BOARD = (("bakeoff.pdf", "vlm_raw.zip"), ("bakeoff2.pdf", "vlm_raw2.zip"),
         ("bakeoff3.pdf", "vlm_raw3.zip"), ("textpdf1.pdf", "vlm_raw_textpdf1.zip"),
         ("docx1.docx", None), ("xlsx1.xlsx", None))


def board_row(name: str, md: str, report: dict, route: str) -> dict: ...


def build_board(work_dir: Path, *, fixtures: Path = FIXTURES) -> tuple[dict, dict[str, tuple[str, dict]]]: ...


def main(argv: "list[str] | None" = None) -> int: ...
```

- `build_board`: временный `LocalCache(work_dir / "cache")`; для каждой пары с zip —
  `key = cache_key(pdf, "mineru", "vlm")`, `cache.put(key, zip_bytes, build_meta(key, pdf, zip_bytes,
  provider="mineru", model_version="vlm", page_range=None, pages=<страниц PDF>))`; затем
  `ingest(data, source_name=name, work_dir=work_dir / stem, cache=cache, provider_factory=_offline)`.
  `_offline(model_version)` бросает `RuntimeError(f"табло оффлайн: нет сырого zip для {…}")` —
  промах кэша обязан упасть, а не уйти в облако. Рабочий `.cache/ocr/` табло не читает и не пишет.
  Возврат: `(board, {name: (md, report)})`.
- Отсутствующая фикстура — `FileNotFoundError` с путём (в `main` — код 1); в тесте пропуск
  делает `require_fixture`, не табло.
- `board`:
  ```json
  {"schema_version": 1, "created_at": "2026-09-20T12:00:00Z", "rows": [ … ]}
  ```
- Строка (`board_row`), набор ключей закрытый, порядок такой:

  | Ключ | Значение |
  |---|---|
  | `fixture` | имя файла |
  | `route` | из `detect_route` |
  | `provider` | `report["provider"]` |
  | `critical`, `warning`, `info` | из `report["summary"]` |
  | `by_rule` | `{rule: число}` по `report["findings"]`, ключи отсортированы |
  | `tables` | `len(parse_pipe_tables(md))` |
  | `table_rows` | сумма строк всех таблиц |
  | `table_broken` | число строк, где ячеек не столько, сколько в первой строке своей таблицы, **плюс** число находок `html_table_unparsed` и `table_merge_failed` |
  | `chars` | `len(md)` |
  | `tokens` | `len(md.split())` |
  | `count_diffs`, `threshold` | `None` — заполняет спека 10 |

  Метрика одна на все форматы: ни один ключ не зависит от маршрута.
- `main`: печатает таблицу в консоль (колонки: fixture, route, critical, warning, info,
  tables, table_broken, chars, count_diffs), пишет `BOARD_JSON` (`utf-8`, `ensure_ascii=False`,
  `indent=2`), черновики `DRAFTS/<stem>/out.md` и `report.json` (перезаписываются), и
  заготовки `FIXTURES/<stem>.errors.txt` для пяти фикстур без эталона (`bakeoff.pdf` не нужна —
  у неё `golden.md`). **Существующий `errors.txt` не перезаписывается никогда** — в нём ручной труд.
  Код выхода: `0` — табло построено; `1` — исключение (traceback в stderr). Критичные
  находки кода не меняют: табло меряет, а не судит.

### Формат `<stem>.errors.txt`

```
# страница | было | надо
# «было» — дословный фрагмент _test/board/<stem>/out.md, встречается в нём ровно один раз
#   (не уникален — расширить контекстом). «надо» — чем заменить; пусто = удалить.
# Вставка потерянного: «было» = соседний текст, «надо» = он же со вставкой.
# Черта внутри текста — \| . Страница — номер в PDF; для DOCX/XLSX — «-» или имя листа.
# Строки с # и пустые игнорируются. Ошибок нет — оставить файл как есть.
# DOCX/XLSX: только потери и искажения (объединённые ячейки, нумерация, колонтитулы,
#   потерянные строки). Стиль не правим.
```

Первая строка — `ERRORS_HEADER`; остальное — комментарии той же заготовки.

## Приёмочные тесты

Сеть заглушена в обоих файлах: `monkeypatch.setattr(socket, "socket", <бросает AssertionError>)`.
Пропуски — только через `require_fixture`. Временный кэш — в `tmp_path`.

### `tests/test_ocr_ingest.py`

```python
# маршруты на фикстурах
for name, route in {"bakeoff.pdf": "scan", "bakeoff2.pdf": "scan", "bakeoff3.pdf": "scan",
                    "textpdf1.pdf": "text_tables",
                    "docx1.docx": "office", "xlsx1.xlsx": "office"}.items():
    assert detect_route(require_fixture(name).read_bytes(), name) == route
assert detect_route(Path("test_files/sample.pdf").read_bytes(), "SAMPLE.PDF") == "text"
with pytest.raises(ValueError, match="Неподдерживаемый формат"):
    detect_route(b"x", "a.pptx")

# единый report на всех четырёх маршрутах
KEYS = ["schema_version", "source", "sha256", "provider", "model_version",
        "cache_hit", "verified", "created_at", "summary", "findings"]
for name, provider in (("bakeoff.pdf", "mineru"), ("textpdf1.pdf", "mineru"),
                       ("docx1.docx", "markitdown"), ("xlsx1.xlsx", "markitdown"),
                       ("sample.pdf", "markitdown")):
    md, report = ingest(data, source_name=name, work_dir=tmp_path / name,
                        cache=seeded_cache, provider_factory=offline)
    assert list(report) == KEYS and report["provider"] == provider
    assert report["sha256"] == hashlib.sha256(data).hexdigest()
    assert set(report["summary"]) == {"critical", "warning", "info"}
    assert md.strip()
    assert postprocess(md)[0] == md                         # постпроцессор прошёл и идемпотентен

# MinerU-маршрут — тот же результат, что у run_pipeline
assert ingest(bakeoff, …)[0] == run_pipeline(bakeoff, …)[0]
assert ingest(bakeoff, …)[1]["cache_hit"] is True           # провайдер не создавался

# markitdown-маршруты
assert report_docx["model_version"] is None and report_docx["cache_hit"] is False
assert all(item["page"] is None for item in report_docx["findings"])
assert "|" in md_xlsx                                        # листы доехали таблицами

# verify / engine
with pytest.raises(ValueError, match="MinerU"):
    ingest(docx, source_name="docx1.docx", work_dir=tmp_path, verify=True)
with pytest.raises(ValueError, match="--verify"):
    ingest(bakeoff, source_name="bakeoff.pdf", work_dir=tmp_path, engine="ocrmypdf", verify=True)

# CLI: office-вход, коды и JSON как в спеке 07
code = main([str(require_fixture("docx1.docx")), "--out", str(tmp_path / "d")])
payload = json.loads(capsys.readouterr().out)
assert code in (0, 2) and payload["error"] is None
assert list(payload) == ["status", "out_md", "report", "cache_hit", "findings", "error"]
assert (tmp_path / "d" / "out.md").exists() and (tmp_path / "d" / "report.json").exists()
assert main([str(tmp_path / "a.pptx"), "--out", str(tmp_path / "p")]) == 1
```

### `tests/test_ocr_board.py`

```python
board, outputs = build_board(tmp_path)
rows = board["rows"]
assert board["schema_version"] == 1
assert [r["fixture"] for r in rows] == [name for name, _ in BOARD] and len(rows) == 6
assert all(list(r) == ROW_KEYS for r in rows)               # одна метрика на все форматы
assert all(r["count_diffs"] is None and r["threshold"] is None for r in rows)
for r in rows:
    md, report = outputs[r["fixture"]]
    assert (r["critical"], r["warning"], r["info"]) == tuple(report["summary"][s] for s in SEVERITIES)
    assert sum(r["by_rule"].values()) == len(report["findings"])
    assert r["chars"] == len(md) > 0
assert next(r for r in rows if r["fixture"] == "bakeoff.pdf")["critical"] == 5   # как в test_ocr_cli
assert next(r for r in rows if r["fixture"] == "xlsx1.xlsx")["tables"] >= 1

# промах кэша падает, а не идёт в сеть
with pytest.raises(RuntimeError, match="оффлайн"):
    build_board(tmp_path / "x", fixtures=<копия без vlm_raw3.zip>)

# main: файлы и неприкосновенность ручного труда (пути подменены monkeypatch на tmp_path)
assert main([]) == 0 and json.loads(BOARD_JSON.read_text(encoding="utf-8"))["rows"]
assert (DRAFTS / "docx1" / "out.md").exists()
errors = FIXTURES_TMP / "docx1.errors.txt"
assert errors.read_text(encoding="utf-8").startswith(ERRORS_HEADER)
assert not (FIXTURES_TMP / "bakeoff.errors.txt").exists()
errors.write_text(ERRORS_HEADER + "3 | а | б\n", encoding="utf-8")
main([])
assert errors.read_text(encoding="utf-8").endswith("3 | а | б\n")
```

## Готово, когда

- Шаг 0: таблица ревизии в отчёте прогона и в docs_sync, все четыре zip на месте под ожидаемыми именами.
- `pytest -v` зелёный целиком; `tests/test_ocr_cli.py` и `tests/test_app_fixes.py` не менялись.
- `python -X utf8 -m ocr.board` с выдернутой сетью: таблица из 6 строк, `_test/quality_board.json`,
  шесть `_test/board/<stem>/out.md`, пять `_test/fixtures/ocr/<stem>.errors.txt`.
- `git diff --stat` не содержит `app.py`, `file_converter.py`, `pdf_core.py`, `requirements.txt`;
  в `ocr/cli.py` диф только в `main` / `_parse_args`.
- В отчёте прогона — таблица табло как есть и список: что человеку открыть
  (`_test/board/*/out.md`) и куда писать ошибки. После этого **стоп**: спека 10 не начинается,
  пока человек не заполнил `errors.txt`.

## PLACEHOLDER-ы

1. `pdf_has_tables` на линиях: безрамочную таблицу не увидит, и такой PDF уйдёт в `text`
   (MarkItDown, плоский текст). Приёмка держится на `textpdf1.pdf → text_tables`; если он
   даёт `text` — **стоп**, не подбирать стратегию `text` наугад, решает человек.
   `# ponytail:` комментарий с этим потолком — в коде у функции.
2. Пороги `MIN_TABLE_ROWS/COLS = 2` — стартовые, на одной фикстуре.
3. XLSX через MarkItDown: объединённые ячейки и пустые клетки приходят как `NaN`/пустые.
   Тракт это не чинит; если человек сочтёт это потерей — строка в `xlsx1.errors.txt`,
   решение о своём рендере (openpyxl) — отдельная спека.
4. Шкала ожиданий этапа A (сканы 8–8,5; текстовый PDF ~9; DOCX/XLSX 9,5+) формулы не имеет.
   Табло даёт факты (`count_diffs`, находки, целостность таблиц); балл ставит человек.
5. Фикстуры на маршрут `text` нет: он закрыт только юнит-тестом на `test_files/sample.pdf`
   и в табло не входит.

## Коммит

Три коммита, в этом порядке.

```
OCR-тракт, этап A: ревизия фикстур

docs/docx_converter_docs_sync.md: таблица «фикстура → zip → sha → страниц».
Сырые zip лежат в _test/ (gitignored); дальше тракт оффлайн.
```

```
ПРАВКА #81: единый вход ocr.ingest — PDF/DOCX/XLSX в один report.json

ocr/ingest.py: detect_route (скан / текст с таблицами / текст / офис) на
analyze_pdf_pages + pdfplumber.find_tables; MinerU-маршруты идут в run_pipeline
без изменений, MarkItDown-маршруты — через тот же postprocess + validate +
build_report. ocr/cli.py: main принимает .pdf/.docx/.xlsx и зовёт ingest;
run_pipeline, коды выхода и JSON-итог прежние. app.py не тронут.
```

```
ПРАВКА #82: табло качества python -m ocr.board

ocr/board.py: шесть фикстур, оффлайн из vlm_raw*.zip (временный кэш, промах —
ошибка) и родных конвертаций DOCX/XLSX. Находки по правилам, целостность таблиц,
размер; count_diffs пуст до эталонов. Пишет _test/quality_board.json, черновые
out.md и заготовки errors.txt (существующие не перезаписывает).
```
