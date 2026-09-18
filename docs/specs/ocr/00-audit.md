# 00 — Аудит, эталон, тестовая инфраструктура

**# ПРАВКА:** нет. Продуктовый код не меняется, номер не расходуется.
Следующий свободный номер после этой спеки — #60.

## Цель

1. Зафиксировать факты о коде, на которые опираются спеки 01–07.
2. Собрать `golden.md` из `vlm.md` по закрытому списку правок.
3. Зафиксировать регрессионную метрику `VLM_TO_GOLDEN_DIFFS`.
4. Дать остальным спекам общие тестовые помощники и маркер `live`.
5. Создать `docs/docx_converter_docs_sync.md` (файла с таким именем в
   репозитории нет и не было — создаётся с нуля, по факту аудита).

## Шаг 0 — условия входа (иначе стоп)

- `test_files/sample.pdf` и `test_files/ocr_sample.pdf` существуют.
- `git status` работает (нет `fatal: detected dubious ownership`).
- `pytest -q` зелёный. Базовый прогон на момент подготовки спек без `test_files/`:
  `119 passed, 2 skipped, 3 failed, 2 errors` — все пять из `tests/test_pdf_core.py`,
  все `FileNotFoundError`.

## Трогать

- `_test/fixtures/ocr/golden.md` — создать (папка в `.gitignore`, в коммит не попадёт)
- `tests/ocr_fixtures.py` — создать
- `tests/test_ocr_fixtures.py` — создать
- `conftest.py` (корень, сейчас пустой) — маркер `live`
- `docs/docx_converter_docs_sync.md` — создать

## Не трогать

Всё остальное. В частности `vlm.md`, `pipeline.md`, `bakeoff.pdf` — только чтение.

## Факты аудита (перепроверить и перенести в docs_sync)

| Вопрос PLAN | Ответ по коду |
|---|---|
| Последняя правка | #59, `app.py:284`. `grep -rn "ПРАВКА #" --include=*.py .` |
| Контракт `OcrProvider` дословно | `pdf_core.py:16-20`: `def ocr_pdf_to_markdown(self, pdf_bytes: bytes, page_range: str \| None = None) -> str: ...` |
| Как вызывается тракт | `app._convert_uploaded_file` → `pdf_core.pdf_to_markdown_with_status(bytes, page_range=…, mode=…)` → `(str, dict \| None)`. Аргумент `provider` из UI не передаётся никогда; единственные потребители протокола — тесты |
| Есть ли `requests` | В `requirements.txt` нет. В `.venv` стоит транзитивно |
| Что делает `ocr_converter.py` | subprocess-обёртка `ocrmypdf`: PDF → searchable PDF (`ocr_pdf_to_searchable_pdf`), плюс `check_ocr_dependencies`. Markdown не производит, сети, кэша и отчёта нет |
| Дублирует ли будущий тракт | **Нет.** Остаётся оффлайн-движком за `OcrmypdfProvider` (`--engine ocrmypdf`) |
| `markdown_cleanup.py` | `cleanup_ocr_markdown` уже делает `\bNo\.?\s*(?=\d)` → `№ ` (с пробелом). `Noп/п` не ловит. Ни к чему не подключён |
| `convert.py` и сырая HTML-таблица | Проверить одним прогоном (см. ниже), результат словами — в docs_sync |

### Поправки к PLAN (найдены аудитом, учтены в спеках)

- `ocr_auto_mode.py` протокол провайдера не использует — этап 1 его не трогает.
- Вторая из трёх разорванных таблиц в `vlm.md` имеет **2 колонки**, а не 3: правило
  «одинаковое число колонок» на ней не сработает. Правило склейки — в спеке 04.
- `РоЕ` в `vlm.md` целиком кириллический, это не «смешанный алфавит внутри токена».
- В блоке подписей 6 должностей и 5 ФИО — пары не восстанавливаются.
- «сыручими» и «IR-камерами» правилами не ловятся — только сверкой прогонов (06).
- Лимит страниц MinerU по доке API — 200, не 600.
- Отдельного шага «создание задачи» в API нет: задача ставится сама после PUT.

### Прогон `convert.py` на сырой HTML-таблице

Взять из `vlm.md` заголовок `## ТЕХНИЧЕСКОЕ ЗАДАНИЕ` и первые ~600 символов первого
`<table>` (закрыть теги вручную), прогнать
`convert_md_to_docx(md, out, template_path=None)` во временную папку, открыть
`python-docx`-ом. Записать в docs_sync одной-двумя фразами: таблица появилась /
текст с тегами ушёл абзацем / упало с исключением. `convert.py` не править при
любом исходе — это входные данные для спеки 04, не баг.

## Сборка `golden.md`

Исходник — `vlm.md`. Правки — **закрытым списком**, больше ничего не трогать
(`АТОМОБИЛЕЙ`, `Мрх`, нумерация `1. 3. 4.` в п. 41, `Exel` — остаются: это либо
опечатки исходника, либо неподтверждённое).

| # | Было в `vlm.md` | Стало в `golden.md` | Кол-во |
|---|---|---|---|
| 1 | `Noп/п`, `No1`, `No384-ФЗ`, `No7-ФЗ` | `№ п/п`, `№ 1`, `№ 384-ФЗ`, `№ 7-ФЗ` (с пробелом — так уже делает `cleanup_ocr_markdown`) | 4 |
| 2 | `+50  $C^{\circ}$ ;` и `+35  $C^{\circ}$ .` | `+50 °C;` и `+35 °C.` | 2 |
| 3 | `РоЕ` (кириллица) | `PoE` (латиница) | 2 |
| 4 | `сыручими` | `сыпучими` | 1 |
| 5 | `уличными РоЕ IR-камерами` | `уличными PoE IP-камерами` — только это место; остальные `IP-` и `IP камеры` не трогать | 1 |
| 6 | подписи, см. ниже | | 5 строк |
| 7 | три `<table>` | одна pipe-таблица, см. ниже | |

### Правка 6 — подписи: откуда взято каждое значение

Фамилии Сиражитдинов, Ямалов, Кустова **уже стоят кириллицей в самом `vlm.md`** —
их не восстанавливаем, правим только инициалы.

| Строка `vlm.md` | Было | Стало | Основание |
|---|---|---|---|
| 35 | `A.II. Taipov` | `А.Ш. Таипов` | фамилия — `docs/PLAN_OCR.md:93` и `pipeline.md:33` (`A.Ш. Таипов`); `II` — разобранная на палки `Ш`, по `pipeline.md:33` |
| 37 | `P.P. Hypeeb` | `Р.Р. Нуреев` | фамилия — `docs/PLAN_OCR.md:93`; инициалы — гомоглифы `P`→`Р` |
| 39 | `A.P. Сиражитдинов` | `А.Р. Сиражитдинов` | фамилия дословно из `vlm.md:39`; инициалы — гомоглифы `A`→`А`, `P`→`Р` |
| 41 | `A.III. Ямалов` | `А.Ш. Ямалов` | фамилия дословно из `vlm.md:41`; `III` → `Ш` по `pipeline.md:39` (`A.Ш. Ямалов`) |
| 43 | `E.K. Кустова` | `Е.К. Кустова` | фамилия дословно из `vlm.md:43`; инициалы — гомоглифы `E`→`Е`, `K`→`К` |

PLACEHOLDER: `Ш.` в строках 35 и 41 подтверждён только вторым OCR-прогоном, не
глазами. Человек сверяет со стр. 9 `bakeoff.pdf` и, если нужно, правит golden
вместе с `GOLDEN_SHA256`.

Порядок и разбиение строк блока подписей **сохранить как в `vlm.md`**: должностей
6, ФИО 5, сопоставить их по тексту нельзя.

### Правка 7 — таблица

- Три `<table>` → **одна** pipe-таблица в 3 колонки. Первая строка — шапка,
  за ней `|---|---|---|`.
- Таблица 2 (`vlm.md:17`, 2 ячейки: пустая + текст) и первая строка таблицы 3
  (`vlm.md:19`, 3 ячейки: пустая, пустая, текст) — продолжения пункта **25**.
  Их текст дописывается через один пробел в конец третьей ячейки строки `25`,
  в порядке следования. Отдельных строк с пустым номером в golden нет.
- `<tr><td colspan="3">I. Общие данные</td></tr>` → `| I. Общие данные |  |  |`.
- Текст ячеек — дословно, с учётом правок 1–5. Переносов строк внутри ячеек в
  `vlm.md` нет, символа `|` нет (проверено: 0 вхождений).
- Текст до и после таблицы, пустые строки между абзацами — как в `vlm.md`.

## Интерфейсы (дословно)

`tests/ocr_fixtures.py`:

```python
"""Общие помощники тестов OCR-тракта: пути к фикстурам, пропуски, метрика расхождений."""

import difflib
import re
from pathlib import Path

import pytest

FIXTURES = Path(__file__).resolve().parents[1] / "_test" / "fixtures" / "ocr"


def require_fixture(name: str) -> Path:
    """Путь к фикстуре; pytest.skip, если файла нет (папка в .gitignore)."""


def read_fixture(name: str) -> str:
    """require_fixture + read_text(encoding='utf-8')."""


def text_tokens(md: str) -> list[str]:
    """Токены текста без разметки.

    1. строки-разделители pipe-таблиц (^\s*\|?[\s:|-]+\|?\s*$ с хотя бы одним '-') удалить;
    2. HTML-теги <[^>]+> заменить пробелом;
    3. символ '|' заменить пробелом;
    4. str.split().
    """


def count_diffs(a: str, b: str) -> int:
    """Число не-'equal' опкодов
    difflib.SequenceMatcher(None, text_tokens(a), text_tokens(b), autojunk=False)."""
```

`conftest.py` (корень):

```python
import os

import pytest


def pytest_configure(config):
    config.addinivalue_line(
        "markers", "live: ходит в сеть (MinerU); без MINERU_API_KEY пропускается")


def pytest_collection_modifyitems(config, items):
    if os.environ.get("MINERU_API_KEY"):
        return
    skip = pytest.mark.skip(reason="нет MINERU_API_KEY")
    for item in items:
        if "live" in item.keywords:
            item.add_marker(skip)
```

`tests/test_ocr_fixtures.py` — константы:

```python
VLM_TO_GOLDEN_DIFFS = 14      # факт. значение ставит исполнитель; расчётное — 14
GOLDEN_SHA256 = "<hex>"       # golden.md вне git — sha ловит тихую подмену эталона
```

Расчёт 14: `No`×4 + `°C`×2 + (`РоЕ IR-камерами` — соседние токены, один опкод) +
второй `РоЕ` + `сыручими` + 5 строк подписей. Склейка таблиц опкодов не даёт —
`text_tokens` снимает разметку. Получилось иначе — не подгонять golden, а
выписать в docs_sync каждое расхождение (`get_opcodes`) и объяснить.

### Разделы `docs/docx_converter_docs_sync.md`

`# Синхронизация документации с кодом (аудит <дата>)` → `## Модули` (по факту,
с числом строк) → `## Нумерация правок` → `## Контракт OcrProvider` →
`## Как вызывается OCR-тракт` → `## ocr_converter.py и новый тракт` →
`## convert.py и сырая HTML-таблица` → `## Базовый прогон тестов` →
`## Метрика vlm→golden` → `## Расхождения с CLAUDE.md / README.md / PLAN_OCR.md`.
Сами `CLAUDE.md` и `README.md` в этой спеке не править — только перечислить расхождения.

## Приёмочные тесты (`tests/test_ocr_fixtures.py`)

```python
vlm, golden = read_fixture("vlm.md"), read_fixture("golden.md")

# метрика и неизменность эталона
assert count_diffs(vlm, golden) == VLM_TO_GOLDEN_DIFFS
assert count_diffs(golden, golden) == 0
assert hashlib.sha256(require_fixture("golden.md").read_bytes()).hexdigest() == GOLDEN_SHA256

# правки 1–5
assert not [t for t in text_tokens(golden) if t.startswith("No")]
assert golden.count("№ ") == 4 and "№ п/п" in golden and "№ 384-ФЗ" in golden
assert "$" not in golden and "\\circ" not in golden
assert "+50 °C;" in golden and "+35 °C." in golden
assert "РоЕ" not in golden            # кириллица
assert golden.count("PoE") == 2       # латиница
assert "сыпучими" in golden and "сыручими" not in golden
assert "IR-" not in golden
assert golden.count("IP-камерами") == 1
assert golden.count("IP-") == vlm.count("IP-") + 1      # остальные IP- не тронуты

# правка 6
for name in ("А.Ш. Таипов", "Р.Р. Нуреев", "А.Р. Сиражитдинов", "А.Ш. Ямалов", "Е.К. Кустова"):
    assert name in golden
tail = golden[golden.index("СОГЛАСОВАНО:"):]
assert not re.search(r"[A-Za-z]", tail)

# правка 7
assert "<table" not in golden and "<td" not in golden
lines = golden.split("\n")
idx = [i for i, l in enumerate(lines) if l.startswith("|")]
assert idx == list(range(idx[0], idx[-1] + 1))            # один сплошной блок
rows = [l for l in lines if l.startswith("|")]
assert all(len(r.strip().strip("|").split("|")) == 3 for r in rows)
assert re.fullmatch(r"\|\s*-+\s*\|\s*-+\s*\|\s*-+\s*\|", rows[1])
first = [r.strip().strip("|").split("|")[0].strip() for r in rows[2:]]
assert "" not in first                                     # номер пункта не потерян
assert first.count("25") == 1 and first.count("26") == 1
row25 = next(r for r in rows if r.lstrip("| ").startswith("25 "))
assert "первичная поверка весов" in row25                   # хвост со стр. 6
assert "ITV Интеллект - УРММ" in row25                      # хвост со стр. 7
assert any(r.startswith("| I. Общие данные |") for r in rows)

# помощники
assert text_tokens("<td>a</td><td>b|c</td>\n|---|---|\n| d |") == ["a", "b", "c", "d"]
```

## Готово, когда

- `pytest -v` зелёный; `tests/test_ocr_fixtures.py` не пропущен (фикстуры на месте).
- `golden.md` существует, `GOLDEN_SHA256` и `VLM_TO_GOLDEN_DIFFS` вписаны фактические.
- `docs/docx_converter_docs_sync.md` содержит все разделы, включая словесный
  результат прогона HTML-таблицы через `convert.py`.
- `git status`: изменены только файлы из «Трогать» (golden в статусе не виден — ignored).

## Коммит

```
OCR-тракт, этап 0: аудит, метрика vlm→golden, тестовые помощники

- tests/ocr_fixtures.py: пути к фикстурам, text_tokens, count_diffs
- tests/test_ocr_fixtures.py: проверки golden.md, VLM_TO_GOLDEN_DIFFS, sha эталона
- conftest.py: маркер live (пропуск без MINERU_API_KEY)
- docs/docx_converter_docs_sync.md: факты аудита кода
golden.md лежит в _test/ (gitignored), его sha256 зафиксирован в тесте.
```
