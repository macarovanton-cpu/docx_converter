# 16 — Сверка по картинке в тракте: весь документ полосами, подсказки в report.json

**# ПРАВКА #91**:
- `ocr/vision.py` — новый модуль: окна по всему документу, вызов `Verifier`, локальный diff, фильтры, находки;
- `ocr/ingest.py` — параметры `vision` / `vision_progress`;
- `ocr/cli.py` — флаг `--vision [МОДЕЛЬ]`;
- `ocr/__init__.py` — `Finding.reading` / `Finding.model`;
- `ocr/validate.py` — `build_report`, `REPORT_SCHEMA_VERSION = 2`.

Зависит от: 15 (замер v3, #90). Наш код в сеть не ходит. `claude` запускается только в живой проверке — её
делает человек. В pytest `claude` не вызывается ни разу.

## Цель

Замер v3 (docs_sync, «Замер v3 (22.09.2026), живой добор»):
- Opus нашёл 41 из 46 опкодов остатка с полосами;
- `wrong_fix` — 1, и тот — артефакт сравнения;
- ложных тревог — 2 из 50 контрольных окон;
- склейки `merge` — 29 из 31.

Вывод v2 «склейки модель не видит» снят. Сверка идёт в тракт: модель переписывает страницы скана, `judge`
сравнивает прочитанное с нашим текстом, расхождения становятся подсказками в `report.json`.

Три свойства подсказки:
- **Текст не меняет.** Подсказка — только находка `vision_diff` в `report.json`. При включённых пометках
  (`--annotate`, в UI — «Пометки в тексте») она ещё и пометка `!! ПРОВЕРИТЬ: … !!` в markdown. Без пометок
  markdown **байт в байт** тот же, что без сверки.
- **Сверяется весь документ**, а не места находок валидатора: 38 из 46 ошибок v3 не ловило ни одно правило.
- **Разница ищется локально** — diff по токенам, тот же `judge` и та же нарезка, что в замере (физическая
  страница из `*_model.json`, стык страниц — #90). `ocr/measure.py` не меняется ни строкой: вызывается как
  библиотека.

Работает только локально, где `claude` установлен и залогинен. На Streamlit Cloud сверки нет: UI её не
предлагает (спека 17), а в тракте любая её неудача становится находкой `vision_skipped`, и конвертация идёт
дальше.

## Решения человека (не пересматриваются)

22.09.2026, по итогам замера v3:

- **Модель по умолчанию — `claude-opus-5`** через `claude -p` (`ClaudeCodeVerifier`). Sonnet — только
  параметром (`--vision claude-sonnet-5`).
- **Сверяется весь документ полосами** — как нарезка замера, с физической страницей и стыками (#90).
- **Разница — локальный diff по токенам**, как в замере.
- **Не подсказываются:**
  - разница только в гомоглифах: латиница и кириллица одного начертания по картинке неразличимы, ими
    занимается `mixed_alphabet_unknown`;
  - разница только в пунктуации и тире.
- **Подсказка текст не меняет.**
- **Только локально.** Нет `claude`, нет входа, лимит, таймаут — конвертация не падает, в отчёте находка
  «сверка не выполнена: причина».
- **Не трогать** `postprocess`, `measure`, `gemini_verifier`, `convert.py`, `file_converter`,
  `requirements.txt`, `app.py`.

23.09.2026, ответы на вопросы спеки:

1. **Ложные подсказки: порога и лимита на документ нет; есть предохранитель страницы.**
   - Оценка по v3. Opus на контрольной выборке — 2 ложные тревоги из 50 окон, это 348 токенов. После
     фильтров этой спеки остаётся одна: bakeoff2-c10, «организации.-Произвести». Вторая (textpdf1-c11,
     `°C` / `°С`) снимается фильтром гомоглифов.
   - Это ≈2.9 ложной подсказки на 1000 токенов. На фикстурах ~255 токенов на страницу (9183 токена на
     36 страниц), значит ≈0.7 ложной подсказки на страницу. Событие одно, поэтому 95%-интервал широкий: от 0
     до ~4 на страницу.
   - Настоящих подсказок на тех же фикстурах после фильтров ~35 на 46 опкодов остатка, ≈1 на страницу.
     Точность ≈55–60 %; у документа, где тракт уже хорош, доля ложных наибольшая.
   - Лимит на документ отбрасывал бы настоящие подсказки вслепую: ранжировать их нечем (`confidence` не
     спрашивается, `ratio` уже отсечён `MIN_RATIO`).
   - Предохранитель: больше `PAGE_HINT_LIMIT = 15` подсказок на одной странице — это сбой нарезки или скана
     (в v2 так выглядели 48 из 100 случаев с чужой страницы). Подсказки такой страницы не выдаются, вместо
     них одна находка `vision_skipped` с их числом.
2. **`report.json` — схема v2.**
   - У `Finding` два новых необязательных поля: `reading` (окно прочтения модели) и `model`, по умолчанию
     `None`.
   - У каждой находки ровно 8 ключей; у прочих правил эти два — `null`.
   - Список ключей остаётся закрытым, `REPORT_SCHEMA_VERSION = 2`.
3. **Пропущенный пробел у знака препинания подсказывается** — это склейка, а не пунктуация.
   - Фильтр «только пунктуация» убирает знаки **внутри** токенов и сохраняет границы токенов.
   - `комплексом).В` → `комплексом). В` остаётся подсказкой.
   - На ответах Opus v3 это +4 настоящих: `комплексом).В`, `комплекса)Предусмотреть`,
     `документация(документация`, `ІД(по`. И +2 ложных: `организации.-Произвести`, `2020).Температура`.

## Шаг 0 — условия входа (иначе стоп)

1. `grep -rn "ПРАВКА #9[12]" --include=*.py .` (без `.venv`) пуст; последняя правка в коде — #90.
2. `pytest -v` зелёный; `python -X utf8 -m ocr.board` оффлайн даёт шесть строк с `diffs == thr`.
3. `_test/verify_measure.v3.claude-code.claude-opus-5.json` на месте:
   - `totals.found == 41`, `false_alarm == 2`, `complete == true`;
   - `ops` у `bakeoff2-c10` — `[["replace", "организации.-Произвести", "организации. -Произвести", "merge"]]`.
4. **Проба — одноразовая, в репозиторий не входит.** Алгоритм этой спеки прогоняется на черновиках
   `build_board`. Используются `build_windows`, `run_measure`, `hints` и два фейка из раздела «Приёмочные
   тесты»: «эталон» — точный, «эхо» — с реалистичной полосой. Ожидаемо (проба 22.09.2026):

   | фикстура | токенов md / plain | окон | с полосами | unlocated | unresolved | narrow | seam | стр. / полос |
   |---|---|---|---|---|---|---|---|---|
   | bakeoff.pdf | 2560 / 2559 | 853 | 829 | 22 | 2 | 9 | 8 | 9 / 38 |
   | bakeoff2.pdf | 2073 / 2068 | 690 | 673 | 17 | 0 | 18 | 21 | 8 / 40 |
   | bakeoff3.pdf | 2329 / 2328 | 776 | 753 | 23 | 0 | 41 | 6 | 9 / 43 |
   | textpdf1.pdf | 2221 / 2186 | 729 | 602 | 124 | 3 | 144 | 0 | 10 / 97 |

   | фикстура | опкодов эталона | косметических | к покрытию | покрыто | «эталон»: подсказок вне опкодов | «эхо»: подсказок |
   |---|---|---|---|---|---|---|
   | bakeoff.pdf | 13 | 3 | 10 | 7 | 0 | 0 |
   | bakeoff2.pdf | 20 | 4 | 16 | 15 | 0 | 0 |
   | bakeoff3.pdf | 8 | 1 | 7 | 7 | 0 | 0 |
   | textpdf1.pdf | 13 | 5 | 8 | 6 | 0 | 0 |

   Опкоды считаются по plain-токенам (`plain_tokens` черновика против `plain_tokens` эталона), а не по
   `text_tokens` табло; число совпало с табло на всех четырёх.

   **Косметические** — `is_cosmetic` истинно:
   - bakeoff: `помещений-`, `помещениях-`, `E.K.`;
   - bakeoff2: `ІД(по`→`IД(по`, `C-21`, `A.C.`, `A. C.`;
   - bakeoff3: `(CPO),`;
   - textpdf1: `M0601`, `M0808`, `M70,`, `MB150; C;– H4,`, `H22С;`.

   **Не покрыты — 6, закрытый список.** Каждого касается окно без полос (`unlocated` / `unresolved`):
   - bakeoff: `IR-камерами` → `IP-камерами`; `A.II. Taipov P.P. Hypeeb A.P.` (подписи); `A.III.` (подписи);
   - bakeoff2: вставка `3.`;
   - textpdf1: вставка строки модификаций `М0601 … М0808`; вставка `Федеральное государственное унитарное`.

   Время на машине человека:
   - растеризация — 7–11 с на фикстуру;
   - вердикт на точном фейке «эталон» (полоса — текст всего блока) — 75–103 с на скан, 6 с на textpdf1;
   - на «эхе» с реалистичной полосой — 25–31 с на скан (≈3 с на страницу), 6 с на textpdf1.

   Другие числа — стоп с распечаткой. Алгоритм под числа не подкручивать.
5. Каталог `.cache/ocr/verify/` — не чистить (там ответы v3, часть полос тракта совпадёт с ними байт в байт).

## Трогать

- `ocr/vision.py` — создать (#91).
- `ocr/ingest.py` — `ingest`: параметры `vision`, `vision_progress`, ветка сверки (#91).
- `ocr/cli.py` — только `_parse_args` (флаг `--vision`) и `main` (передача в `ingest`, прогресс в stderr) (#91).
- `ocr/__init__.py` — два поля `Finding` (#91).
- `ocr/validate.py` — только `REPORT_SCHEMA_VERSION` и `build_report` (#91).
- `tests/test_ocr_vision.py` — создать (#91).
- `tests/test_ocr_validate.py::test_report_schema`, `tests/test_ocr_cli.py` — `schema_version == 2` и восемь
  ключей находки. Больше ничего в этих файлах.
- `CLAUDE.md`:
  - список модулей `ocr/`: +`ocr/vision.py`, «Eleven» → «Twelve»;
  - абзац «OCR CLI»: флаг `--vision`, ключи находки `report.json`;
  - «Known production limitation»: сверка по картинке — только локально;
  - фраза о нумерации: `#91 — ocr/vision.py + ocr/ingest.py + ocr/cli.py + ocr/__init__.py + ocr/validate.py`;
  - строка-урок: «Подсказка сверки по картинке текст не меняет: только находка `vision_diff`; гомоглифы и
    пунктуацию по скану не подсказывать (#91)».
- `docs/specs/ocr/README.md`:
  - строка спеки 16 в таблице;
  - `Finding` в «Общих типах» — с двумя новыми полями;
  - в таблицу правил — `vision_diff` и `vision_skipped` (оба `warning`, выдаёт 16);
  - сводка PLACEHOLDER-ов — п. 11.
- `docs/PLAN_OCR.md` — «Этап B, шаг 4 (#91)» и строка статуса.
- `docs/docx_converter_docs_sync.md` — `### ПРАВКА #91` и снимок «следующий — #92».

## Не трогать

- `ocr/measure.py` — ни строкой. Из него только импорт: `Case`, `TRANSCRIBE_QUESTION`, `model_pages`, `norm`,
  `page_tiles`, `physical_place`, `run_measure`, `split_tiles`. Приватные имена (`_split`, `_touches`,
  `_content_list`) не импортируются — по конвенции #90 («файл заморожен, приватное имя не импортируем»).
  Их логика в `ocr/vision.py` — копия с пометкой.
- `ocr/gemini_verifier.py` — только импорт `CONTEXT_TOKENS`, `crop_block`, `locate_block`.
- `ocr/claude_code_verifier.py` — только импорт `ClaudeCodeVerifier`.
- `ocr/postprocess.py` — только импорт `HOMOGLYPHS_CYR_TO_LAT`.
- `ocr/board.py` — только импорт `text_tokens`.
- `ocr/diff.py`, `ocr/mineru_provider.py`, `ocr/cache.py`, `pdf_core.py`.
- В `ocr/cli.py` — `run_pipeline`, `_mineru_result` и коды выхода.
- В `ocr/validate.py` — все правила, `page_of`, `annotate`, `strip_annotations`.
- `app.py` (это спека 17), `file_converter.py`, `convert.py`, `requirements.txt`, `template*.docx`, `conftest.py`.
- Фикстуры, эталоны, `*.errors.txt`, zip MinerU, `_test/verify_measure*`.
- `.cache/ocr/` и `.cache/ocr/verify/` — не чистить и руками не править.

## Интерфейсы (дословно)

### `ocr/__init__.py` (#91)

```python
@dataclass(frozen=True)
class Finding:
    rule: str
    severity: str            # одно из SEVERITIES
    page: int | None         # 1-based; None — страница неизвестна
    snippet: str             # дословный фрагмент из markdown, по нему ищется место пометки
    suggestion: str | None = None
    reading: str | None = None     # ПРАВКА #91: vision_diff — окно прочтения модели по скану
    model: str | None = None       # ПРАВКА #91: vision_* — модель сверки по картинке
```

Докстринг модуля — `"""ПРАВКА #61/#63: OCR-тракт MinerU. Общие типы. ПРАВКА #91: reading, model."""`.
Все существующие `Finding(...)` в коде именованные, новые поля с умолчанием — вызовы не меняются.

### `ocr/validate.py` (#91)

```python
REPORT_SCHEMA_VERSION = 2        # ПРАВКА #91: у находки + reading, model
```

`build_report`: элемент `findings[]` —
`{"id", "rule", "severity", "page", "snippet", "suggestion", "reading", "model"}`. Порядок ключей ровно такой.
`reading` и `model` берутся из `Finding`. Верх отчёта не меняется: те же десять ключей.

### `ocr/vision.py` (#91)

```python
"""ПРАВКА #91: сверка по картинке в тракте. Весь документ режется на окна, каждое окно — полосы своего блока
на физической странице; модель (claude -p) переписывает страницу, разница — локальный diff по токенам тем же
judge, что в замере (ocr.measure не меняется). Подсказка текст не меняет: только находка vision_diff.
Только локально: без claude, входа или лимита — находка vision_skipped, конвертация идёт дальше."""

import json
import re
import sys
import traceback
import zipfile
from collections import Counter
from dataclasses import dataclass
from difflib import SequenceMatcher
from io import BytesIO

from ocr import Finding
from ocr.board import text_tokens
from ocr.claude_code_verifier import ClaudeCodeVerifier
from ocr.gemini_verifier import CONTEXT_TOKENS, crop_block, locate_block
from ocr.measure import (TRANSCRIBE_QUESTION, Case, model_pages, norm, page_tiles, physical_place, run_measure,
                         split_tiles)
from ocr.postprocess import HOMOGLYPHS_CYR_TO_LAT
from pdf_core import Verifier

VISION_MODEL = "claude-opus-5"   # решение человека 22.09.2026: Opus по умолчанию, Sonnet — параметром
VISION_STRIDE = 3                # PLACEHOLDER: токенов в пролёте окна; окно — пролёт ± CONTEXT_TOKENS
PAGE_HINT_LIMIT = 15             # решение человека 23.09.2026: больше на странице — сбой нарезки или скана
SEVERITY = {"vision_diff": "warning", "vision_skipped": "warning"}
_MARKUP_RE = re.compile(r"#{1,6}|!?\[[^\]]*\]\([^)]*\)")    # заголовок, картинка, ссылка: на скане их нет
_TO_CYRILLIC = str.maketrans({lat: cyr for cyr, lat in HOMOGLYPHS_CYR_TO_LAT.items()})   # как в ocr/measure.py


@dataclass(frozen=True)
class Hint:
    start: int           # пролёт «нашей» стороны в plain-токенах, [start, end); у вставки — соседний токен
    end: int
    suggestion: str      # прочитанное по скану, токены через пробел; "" — у нас лишнее
    page: int            # физическая страница окна (у стыка — первая из двух)
    reading: str         # окно прочтения, по которому судилось окно


def make_verifier(model: str) -> Verifier: ...
def plain_tokens(md: str) -> tuple[list[str], list[tuple[int, int]]]: ...
def is_cosmetic(ours: list[str], read: list[str]) -> bool: ...
def build_windows(tokens: list[str], pdf_bytes: bytes, content_list: list, pages: list[str]) -> tuple[list[Case], Counter]: ...
def hints(rows: list[dict]) -> list[Hint]: ...
def limit_pages(found: list[Hint]) -> tuple[list[Hint], dict[int, int]]: ...
def skipped_finding(page: int | None, text: str, model: str) -> Finding: ...
def vision_findings(md: str, pdf_bytes: bytes, zip_bytes: bytes, *, source_name: str, model: str = VISION_MODEL,
                    progress=None) -> list[Finding]: ...
```

Импорт `ocr.vision` тянет `ocr.measure` → `ocr.board` → `ocr.ingest`. Поэтому `ocr.ingest` и `ocr.cli`
импортируют `ocr.vision` **внутри функции**, иначе цикл импортов. Прецедент — `ocr.cli.main`, который так же
импортирует `ocr.ingest` (#81).

**`make_verifier(model)`** → `ClaudeCodeVerifier(model=model, live=True)`.
- `live=True`: флаг `--vision` (в UI — галочка) и есть согласие на расход подписки. `CLAUDE_CODE_LIVE` —
  предохранитель замера, в тракте он не нужен.
- Кэш — умолчание `ClaudeCodeVerifier`, то есть тот же `.cache/ocr/verify/`, те же ключи.
- Полоса блока `(page_idx, bbox)` режется так же, как в замере, байт в байт. Где тракт режет блок, уже
  нарезанный замером, ответ v3 берётся из кэша.
- Тесты подменяют `ocr.vision.make_verifier` (`monkeypatch.setattr`). `vision_findings` ищет имя в модуле в
  момент вызова.

**`plain_tokens(md)`** → `(tokens, spans)`, токены окон и срез каждого в `md`.
- Идём по `text_tokens(md)` по порядку:
  - `start = md.find(token, pos)`, `pos = start + len(token)`;
  - `start == -1` → `ValueError` (не должно случаться: токены `text_tokens` — подстроки `md` в том же порядке).
- Токен **выбрасывается**, если `_MARKUP_RE.fullmatch(token)` или `token.replace("\\", "")` пуст.
- Иначе в `tokens` идёт `token.replace("\\", "")` (`\-` → `-`), в `spans` — `(start, pos)`.
- Зачем: в content_list и на скане нет `#`, `##`, `\-` и `![](images/…)`.
  - Без фильтра `locate_block` не находил окно: у textpdf1 не привязывались 139 окон из 741.
  - Ссылка на картинку дала бы ложное «у нас лишнее».
- Срез может захватить строку-разделитель таблицы, если `find` споткнулся о `-` в ней. Он всё равно
  дословный и лежит в том же блоке; для пометки этого хватает.

**`is_cosmetic(ours, read)`** — разница только в гомоглифах, пунктуации и тире.
```python
def _letters(tokens):
    return [t for t in (re.sub(r"[\W_]", "", x.translate(_TO_CYRILLIC)) for x in tokens) if t]
return _letters(ours) == _letters(read)
```
- Свёртка латиницы в кириллицу — до удаления знаков.
- Знаки убираются внутри токена; границы токенов остаются (решение 3).
- Регистр не трогается.

**`build_windows(tokens, pdf_bytes, content_list, pages)`** → `(cases, skipped)`. `pages` — `model_pages`.
- Пролёты: `i` в `range(0, len(tokens), VISION_STRIDE)`, `e = min(len(tokens), i + VISION_STRIDE)`,
  `lo = max(0, i - CONTEXT_TOKENS)`, `hi = min(len(tokens), e + CONTEXT_TOKENS)`. Каждый токен лежит ровно
  в одном пролёте.
- `needle, before, after = tokens[i:e], tokens[lo:i], tokens[e:hi]` → `locate_block(needle, before, after,
  content_list)`.
  - Блока нет → `skipped["unlocated"] += 1`.
- `index` — позиция блока в `content_list` (по `is`, как в `build_cases`) →
  `physical_place(needle, before, after, content_list, index, pages)`:
  - `"unresolved"` → `skipped["unresolved"] += 1`;
  - `"narrow"` → **окно судится по полосам своего блока** (`blocks = [index]`). Это отличие от замера, где
    `narrow` — служебный исход. В документе много коротких текстовых блоков: заголовки, колонтитулы,
    строки списков. Выбросить их — потерять покрытие. От ложной подсказки защищают `MIN_RATIO` (окно шире
    блока даёт `ratio < 0.5` → не прочитано) и `clipped` (контекст за краем блока — срезанный край);
  - `"seam"` → `next_tiles` — полосы блока `blocks[1]`.
- Полосы — `split_tiles(crop_block(pdf_bytes, page_idx, bbox))` блока `blocks[0]`, кэш по
  `(page_idx, tuple(bbox))`, как `tiles_of` в `build_cases`.
- Окно — `Case`:

  ```python
  Case(id=f"w{i:05d}", fixture="", kind="error", tag=None, fragment=" ".join(tokens[lo:hi]),
       expected=" ".join(tokens[lo:hi]), span=(i - lo, e - lo), gold_span=(i - lo, e - lo), error_kind=None,
       rules=(), page=blocks[0] page_idx + 1, block_type=block["type"], ambiguous=ambiguous, tiles=…,
       next_tiles=…, page_source=source, list_page=block["page_idx"] + 1)
  ```

  - Эталона в тракте нет, поэтому `expected == fragment`. Тогда `judge` отдаёт `wrong_fix` ровно тогда, когда
    живой опкод «наше ↔ прочитанное» касается пролёта. `found` невозможен.
  - `id` хранит начало пролёта: `i = int(id[1:])`, `lo = i - span[0]`.
  - `fixture=""`: `run_measure` смотрит `BOARD` только у случаев без полос, а их в `cases` нет.
- Возврат: `cases` только с полосами, `skipped` — `Counter` с ключами `unlocated`, `unresolved`. Всего окон —
  `len(range(0, len(tokens), VISION_STRIDE))`.

**`hints(rows)`** — подсказки из строк `run_measure(...)["cases"]`.
- Берутся строки с `outcome == "wrong_fix"`.
- `raw = row["fragment"].split()`, `(s, e) = row["span"]`, `i = int(row["id"][1:])`, `lo = i - s`.
- Живые токены — как `norm` потокенно (`norm` поштучный, поэтому это то же, что `norm(raw)`):
  `kept = [k for k, t in enumerate(raw) if norm([t])]`, `f = [norm([raw[k]])[0] for k in kept]`.
- Пролёт в `f`: `fs = sum(k < s for k in kept)`, `fe = sum(k < e for k in kept)`.
  Это копия `ocr.measure._split`, приватное имя не импортируется.
- `W = row["reading"].split()`: `reading` — `" ".join` токенов окна прочтения, внутри токенов пробелов нет.
- Опкоды `SequenceMatcher(None, f, W, autojunk=False)`, кроме `equal`, у которых `_touches(i1, i2, fs, fe)`.
  `_touches` — копия `ocr.measure._touches`.
- Опкод с `is_cosmetic(f[i1:i2], W[j1:j2])` пропускается.
- Пролёт в plain-токенах:
  - `i2 > i1` → `(lo + kept[i1], lo + kept[i2 - 1] + 1)`;
  - вставка → соседний токен слева `(lo + kept[i1 - 1], …+1)`, при `i1 == 0` — справа `(lo + kept[0], …+1)`.
- `Hint(start, end, " ".join(W[j1:j2]), row["page"], row["reading"])`.
- Повторы (опкод на границе двух пролётов попадает в оба окна) схлопываются по `(start, end, suggestion)`,
  первый остаётся.
- Порядок — по `(start, end)`.

**`limit_pages(found)`** → `(kept, flooded)`.
- `flooded = {page: n}` — страницы, где подсказок больше `PAGE_HINT_LIMIT`.
- `kept` — остальные подсказки, в прежнем порядке.

**`vision_findings(md, pdf_bytes, zip_bytes, *, source_name, model=VISION_MODEL, progress=None)`**. Порядок:

1. **Всё — в одном `try`.** Любое исключение (zip без `*_model.json`, сбой растеризации, дефект кода) →
   `traceback.print_exc(file=sys.stderr)` и `[skipped_finding(None, f"сверка не выполнена: {type(exc).__name__}: {exc}", model)]`.
   Исключение наружу не выходит никогда.
2. `content_list` — единственный член zip на `content_list.json` (копия `ocr.measure._content_list`, но по
   байтам). Не один → `ValueError`.
3. `pages = model_pages(archive)`, где `archive = BytesIO(zip_bytes)` и `archive.name = f"{source_name}: zip MinerU"`.
   `model_pages` открывает `zipfile.ZipFile(zip_path)` и берёт `zip_path.name` только в тексте ошибки.
4. `tokens, spans = plain_tokens(md)`; `cases, skipped = build_windows(tokens, pdf_bytes, content_list, pages)`.
5. `verifier = make_verifier(model)` — **после** шагов 2–4: без `content_list` верификатор не создаётся.
6. `keys = list(page_tiles(cases))`;
   `measure = run_measure(cases, _Progress(verifier, keys, progress))`. `_Progress` — приватная обёртка:
   - перед каждым `transcribe` зовёт `progress(page, done, total)` (`page` — `keys[k][1]`, `done` — `k + 1`,
     `total` — `len(keys)`), если `progress` не `None`;
   - после успешного вызова переносит `last_cache_hits` верификатора (или `[False] * n`);
   - `run_measure` читает его через `getattr`, как в замере.
7. `found, flooded = limit_pages(hints(measure["cases"]))`.
8. Находки, в этом порядке:
   - `vision_diff` на каждую `Hint`: `page=hint.page`, `snippet=md[spans[start][0]:spans[end - 1][1]]`,
     `suggestion=hint.suggestion`, `reading=hint.reading`, `model=model`;
   - `vision_skipped` по сбоям и покрытию — таблица ниже.
9. В stderr одна строка: `сверка по картинке ({model}): стр. {len(keys)}, вызовов claude {calls}, подсказок
   {len(found)}`, где `calls = getattr(verifier, "network_calls", 0)`.

`skipped_finding(page, text, model)` → `Finding(rule="vision_skipped", severity="warning", page=page,
snippet="", suggestion=text, model=model)`. Пустой `snippet` при пометках уводит находку в хвост «## Не привязанные
находки» — так задумано в #70.

| когда | `page` | `suggestion` |
|---|---|---|
| исключение (шаг 1) | `None` | `сверка не выполнена: {Тип}: {текст}` |
| строка с `error_reason == "stop"` (нет бинаря, не залогинен, лимит, не та модель) | стр. остановки | `сверка остановлена на стр. {p}: {error}; не сверены стр. {p, …}` — страница остановки и все ключи `keys` после неё |
| строки с прочей `error_reason` (`timeout`, `cli`, `parse`, `count`, `other`, `miss`) — по одной на пару `(row["page"], error)` | эта | `стр. {p} не сверена: {error}` |
| `flooded` — по одной на страницу | эта | `стр. {p}: подсказок {n} — больше {PAGE_HINT_LIMIT}, похоже на сбой нарезки или скана; подсказки страницы не выданы, сверить страницу глазами` |
| не сверено окон > 0 | `None` | `не сверено окон {k} из {total}: не привязаны к скану — {unlocated + unresolved}, не прочитаны — {unreadable}, у края полосы — {clipped}` |

- `unreadable` — строки с `verdict == "unreadable"`, `clipped` — строки с `reading == "clipped"`. Окна
  несверенных страниц сюда не входят: у них своя находка.
- `error` — поле строки `run_measure`, `f"{Тип}: {текст}"`.
- Нет бинаря — это `ClaudeCodeMissingError` (`VerifierConfigError` → `STOP_ERRORS`): `run_measure`
  останавливается на первой странице, находка — «сверка остановлена на стр. 1: ClaudeCodeMissingError: не
  найден claude…».

### `ocr/ingest.py` (#91)

```python
def ingest(data: bytes, *, source_name: str, work_dir: Path,
           engine: str = "mineru", mode: str = "vlm",
           verify: bool = False, annotate: bool = False,
           annotate_all: bool = False,
           cache: "CacheBackend | None" = None,
           provider_factory=None,
           vision: str | None = None,          # ПРАВКА #91: модель сверки по картинке; None — сверки нет
           vision_progress=None) -> tuple[str, dict]:   # ПРАВКА #91: (page, done, total) перед каждой страницей
```

- `vision is None` — поведение прежнее, **ни одна ветка не меняется**.
- `vision` — это не `verify`: `verify` — второй прогон MinerU (`low_confidence`), `vision` — сверка по
  картинке. Независимы, можно вместе.
- `vision` на маршрутах `text` / `office` и при `engine == "ocrmypdf"` → `ValueError("сверка по картинке
  доступна только для маршрутов MinerU")`. Как `--verify`: проверка до конвертации.
- `vision` при `cache is None` → `ValueError("сверке по картинке нужен кэш: сырой ответ MinerU берётся из него")`.
  CLI и UI кэш передают всегда.
- Маршруты `scan` / `text_tables` с `vision`:
  1. `md, report = run_pipeline(data, **{**pipeline, "annotate": False, "annotate_all": False})`;
  2. `from ocr.vision import vision_findings` — здесь, не наверху модуля;
  3. `cached = cache.get(cache_key(data, "mineru", mode))`. После `run_pipeline` запись есть всегда. Нет —
     `skipped_finding(None, "сверка не выполнена: сырого ответа MinerU нет в кэше", vision)`;
  4. `extra = vision_findings(md, data, cached[0], source_name=source_name, model=vision,
     progress=vision_progress)`;
  5. отчёт пересобирается: `build_report(source=…, sha256=…, provider=…, model_version=…, cache_hit=…,
     verified=…` — всё из `report` — `, findings=[Finding(**{k: f[k] for k in FINDING_FIELDS}) for f in
     report["findings"]] + extra)`. `content_list` не передаётся: страницы у прежних находок уже
     проставлены;
  6. `annotate or annotate_all` → `annotate_md(md, report, include_low_confidence=annotate_all)`.
- `FINDING_FIELDS = ("rule", "severity", "page", "snippet", "suggestion", "reading", "model")` — константа
  `ocr/ingest.py`.
- `run_pipeline` не меняется.

### `ocr/cli.py` (#91)

```python
    parser.add_argument("--vision", nargs="?", const=VISION_MODEL, default=None, metavar="МОДЕЛЬ",
                        help="сверка по картинке через Claude Code (claude -p; только локально, где claude "
                             "установлен и залогинен). Без значения — claude-opus-5; подсказки -> находки "
                             "vision_diff, текст не меняется")
```

- `VISION_MODEL` импортируется внутри `_parse_args`: `from ocr.vision import VISION_MODEL`.
- `main` передаёт в `ingest` `vision=args.vision` и `vision_progress`, который пишет в stderr
  `сверка по картинке: стр. {page} ({done} из {total})`.
- stdout — по-прежнему одна строка JSON; коды выхода прежние. Подсказки — `warning`, кода `2` не дают.

## Приёмочные тесты

Сеть и `claude` недоступны: в autouse-фикстуре `socket.socket` и `subprocess.run` — заглушки с
`AssertionError`, переменные `CLAUDE_CODE_LIVE`, `VERIFIER`, `GEMINI_LIVE` сняты через `delenv`. Фикстуры —
через `require_fixture`. `make_verifier` подменяется во всех тестах, где доходит до вызова.

### `tests/test_ocr_vision.py` (#91)

```python
# is_cosmetic: пары из v3 и из пробы
assert is_cosmetic(["помещений-"], ["помещений"]) and is_cosmetic(["°C:"], ["°С:"])
assert is_cosmetic(["M0601"], ["М0601"]) and is_cosmetic(["C;–"], ["C;"]) and is_cosmetic(["II–"], ["II-"])
assert is_cosmetic(["MB150;", "C;–", "H4,"], ["МВ150;", "С;–", "Н4,"]) and is_cosmetic(["\\-"], [])
assert not is_cosmetic(["IP65не"], ["IP65", "не"]) and not is_cosmetic(["сыручими"], ["сыпучими"])
assert not is_cosmetic(["комплексом).В"], ["комплексом).", "В"])                      # решение 3
assert not is_cosmetic(["организации.-Произвести"], ["организации.", "-Произвести"])    # c10 — остаётся
assert not is_cosmetic([], ["Значение"]) and not is_cosmetic(["IR-камерами"], ["IP-камерами"])

# plain_tokens
md = "## Заголовок\n\n| № | Имя |\n|---|---|\n| 1 | \\- датчик ![](images/a.jpg) |\n"
tokens, spans = plain_tokens(md)
assert tokens == ["Заголовок", "№", "Имя", "1", "-", "датчик"]
assert [md[s:e] for s, e in spans] == ["Заголовок", "№", "Имя", "1", "\\-", "датчик"]

# limit_pages
h = lambda p, n: Hint(n, n + 1, "x", p, "x")
kept, flooded = limit_pages([h(2, n) for n in range(16)] + [h(3, 99)])
assert flooded == {2: 16} and kept == [h(3, 99)]
assert limit_pages([h(2, n) for n in range(15)])[1] == {}                                # 15 — ещё можно
```

**Четыре PDF-фикстуры** — фикстура модуля (`scope="module"`), считается один раз:
- `build_board` во временной папке;
- для каждой `(pdf, zip)` из `BOARD` с zip: `tokens` = `plain_tokens(черновик)`, `content_list`, `pages`,
  `cases, skipped = build_windows(...)`.

Фейки (оба — `transcribe(images, question)`, `question == TRANSCRIBE_QUESTION`):
- **«эталон», точный.** `g = plain_tokens(эталон)[0]`, опкоды `SequenceMatcher(None, tokens, g,
  autojunk=False)`. `to_b` — как в `ocr.measure.build_cases`: `i == len` → `len(g)`, `i == 0` → `0`.
  Полоса → `" ".join(g[to_b(lo_min):to_b(hi_max)])`, где `lo_min` / `hi_max` — крайние `lo` / `hi` окон, у
  которых полоса в `tiles` или `next_tiles`.
- **«эхо», реалистичная полоса.** Для каждого кортежа полос блока `T` (у окна — `tiles` или `next_tiles`)
  диапазон `[lo, hi]` — по окнам с этим `T`. `n = len(T)`, `L = hi - lo`,
  `ov = 2 * (VISION_STRIDE + 2 * CONTEXT_TOKENS)`. Полоса `T[k]` получает отрезок
  `[max(lo, lo + k*L//n - ov), min(hi, lo + (k+1)*L//n + ov))`. Текст полосы — `" ¶ ".join` отрезков из
  `tokens`. Он в ~3 раза короче точного фейка: так вердикт идёт ≈3 с на страницу, как в жизни.

```python
# состав — числа Шага 0, поштучно по фикстурам
assert len(range(0, len(tokens), VISION_STRIDE)) == WINDOWS[name]          # 853 / 690 / 776 / 729
assert len(cases) == PLACED[name] and dict(skipped) == SKIPPED[name]         # 829 / 673 / 753 / 602
assert Counter(c.page_source for c in cases) == SOURCES[name]                # narrow / seam / … по таблице
assert all(c.tiles for c in cases) and all(bool(c.next_tiles) == (c.page_source == "seam") for c in cases)

# «эхо»: ни одной подсказки
assert hints(run_measure(cases, echo)["cases"]) == []

# «эталон»: подсказки только у опкодов эталона; все некосметические покрыты, кроме закрытого списка
found = hints(run_measure(cases, golden)["cases"])
ops = [op for op in opcodes if op[0] != "equal"]
near = lambda h, op: _touches(h.start, h.end, op[1] - (op[1] == op[2]), op[2] + (op[1] == op[2]))
assert all(any(near(h, op) for op in ops) for h in found)                    # ложных — 0
want = [op for op in ops if not is_cosmetic(tokens[op[1]:op[2]], g[op[3]:op[4]])]
missed = [" ".join(tokens[op[1]:op[2]]) or "+" + " ".join(g[op[3]:op[4]]) for op in want
          if not any(near(h, op) for h in found)]
assert missed == MISSED[name]
# MISSED: bakeoff ["IR-камерами", "A.II. Taipov P.P. Hypeeb A.P.", "A.III."]; bakeoff2 ["+3."];
#   bakeoff3 []; textpdf1 ["+М0601 i 30, i 35, i 40-SS МИ ВДА/12Я МИ ВДА/12ЯС DIS2116 М0808",
#   "+Федеральное государственное унитарное"]
assert len(want) - len(missed) == COVERED[name]                              # 7 / 15 / 7 / 6
# каждый непокрытый касается окна без полос — это «не привязано», а не «не увидели»
```

`_touches` в тесте — копия из `ocr/vision.py`, импорт приватного имени из своего модуля допустим.
`docx1` / `xlsx1`: `ingest(..., vision=VISION_MODEL)` → `ValueError` «только для маршрутов MinerU».

**`vision_findings` и `ingest` на textpdf1** (маршрут `text_tables`, вердикт — секунды). Кэш засевается, как
`seeded_cache` в `tests/test_ocr_ingest.py`; провайдер — `offline`.

```python
# фейки строятся тем же build_windows на тех же входах (детерминирован), «эхо» — с реалистичной полосой
# эхо: vision_diff нет; одна vision_skipped покрытия, в ней «из 729» и «не привязаны к скану — 127»
# «одна правка»: w — первый токен пролёта первого окна с block_type "text", из одних букв, длиной >= 6;
#   в тексте полос этого окна (эхо) каждое вхождение w заменено на w + "ъ" ->
#   vision_diff есть, у всех snippet == w (дословно из md), suggestion == w + "ъ", model == "fake-model",
#   reading содержит w + "ъ"; у первой page == page окна
# md, отчёт = ingest(..., vision="fake-model", annotate=False); md0, _ = ingest(..., annotate=False)
assert md == md0                                                             # текст не меняется
assert report["schema_version"] == 2
assert all(list(f) == ["id", "rule", "severity", "page", "snippet", "suggestion", "reading", "model"]
           for f in report["findings"])
assert all(f["reading"] is None and f["model"] is None for f in report["findings"]
           if not f["rule"].startswith("vision_"))
# annotate=True (фейк «одна правка»): пометка «!! ПРОВЕРИТЬ: [N] vision_diff: <w>ъ !!» стоит после блока с w;
#   strip_annotations(md_annotated) == md0
# progress: вызовы (page, done, total) — total == 10, done == 1..10 по порядку, page — ключи page_tiles
# сбои фейка (VerifierError-наследники из ocr.claude_code_verifier), md каждый раз == md0:
#   ClaudeCodeMissingError на 1-м вызове -> одна «сверка остановлена на стр. 1: ClaudeCodeMissingError: …;
#     не сверены стр. 1, …, 10», vision_diff нет
#   ClaudeCodeLimitError на 3-м -> «сверка остановлена на стр. P3: …; не сверены стр.» — ключи page_tiles с 3-го;
#     первых двух ключей page_tiles в «не сверены» нет, находок «не сверена» у них тоже нет
#   ClaudeCodeTimeoutError на одной странице -> «стр. N не сверена: ClaudeCodeTimeoutError: …», прочие сверены
#   RuntimeError("boom") в transcribe -> «сверка не выполнена: RuntimeError: boom», трейсбек в stderr
#   zip без *_model.json -> «сверка не выполнена: ValueError: …», make_verifier не вызывался
# cache=None + vision -> ValueError «нужен кэш»; engine="ocrmypdf" + vision -> ValueError
# vision=None -> отчёт тот же, что до #91, кроме schema_version и двух null-ключей (сравнить rule/snippet/suggestion)
# make_verifier("claude-opus-5") -> ClaudeCodeVerifier, .model == "claude-opus-5", ._live is True (без вызова)
```

**CLI** — по образцу `tests/test_ocr_cli.py::test_main_writes_files_and_one_json_line`: `chdir(tmp_path)`,
кэш `.cache/ocr` в `tmp_path` засеян zip textpdf1, `provider_factory=offline`, `make_verifier` подменён.

```python
# _parse_args: без флага -> None; «--vision» -> VISION_MODEL; «--vision claude-sonnet-5» -> "claude-sonnet-5"
# main([...textpdf1.pdf, "--out", out, "--vision"]) -> код 0, stdout — одна строка JSON,
#   report.json: schema_version 2, есть vision_skipped; stderr: «сверка по картинке: стр. 1 (1 из 10)» и итоговая строка
```

`tests/test_ocr_validate.py::test_report_schema` и `tests/test_ocr_cli.py`: `schema_version == 2`, восемь
ключей находки. Остальное в этих файлах не меняется.

**Время.** Модуль `tests/test_ocr_vision.py` по пробе идёт около 7 минут: три скан-фикстуры на точном
«эталоне» — 75–103 с каждая. Больше 10 минут на машине человека — стоп, решает человек (PLACEHOLDER 6).

## Проверка

1. Шаг 0 — числа пробы совпали.
2. `pytest -v` зелёный целиком. `claude` и сеть не тронуты.
3. `python -X utf8 -m ocr.board` — шесть строк, `diffs == thr`: табло о сверке не знает.
4. `git diff --stat` — только файлы из «Трогать». `git diff ocr/measure.py ocr/gemini_verifier.py
   ocr/claude_code_verifier.py app.py` пуст.

## Живая проверка (только руками человека, после зелёного `pytest` и коммита #91)

Исполнитель эти команды не запускает. Вызовы тратят подписку человека. `CLAUDE_CODE_LIVE` не нужен: флаг
`--vision` сам по себе согласие.

Перед запуском: сырой ответ MinerU для фикстуры лежит в `.cache/ocr/` (`{sha}-mineru-vlm-pall-v1.zip`,
docs_sync «Фикстуры этапа A»). Иначе CLI пойдёт в MinerU — нужен `MINERU_API_KEY`, и это квота MinerU.

```powershell
python -X utf8 -m ocr.cli _test/fixtures/ocr/bakeoff.pdf --out _test/vision/bakeoff --vision
python -X utf8 -m ocr.cli _test/fixtures/ocr/bakeoff.pdf --out _test/vision/bakeoff --vision              # повтор: из кэша, «вызовов claude 0»
python -X utf8 -m ocr.cli _test/fixtures/ocr/bakeoff.pdf --out _test/vision/bakeoff-annot --vision --annotate
```

- Ожидаемо: до 9 вызовов Opus, по одному на страницу. Страницы, где все полосы уже нарезал замер v3,
  берутся из кэша. По прайсу ≈ $0.08 за вызов, реально платит подписка. По v3 — ~15 с на вызов плюс
  ~3 с на страницу вердикта.
- Смотреть в `report.json` находки `vision_diff`, каждую — против 13 опкодов остатка bakeoff (docs_sync, #83,
  таблица «Сырой путь против эталона»). Ожидаемо по пробе и v3: подсказки у `сыручими`, склеек пунктов
  13–14, `комплексом).В`, `комплекса)Предусмотреть`. Нет подсказок у `помещений-` (тире), `E.K.`
  (гомоглиф), подписей (не привязаны).
- Подсказку вне опкодов человек сверяет со сканом: ложная она или пропуск эталона.
- `vision_skipped` покрытия — числа сверить с Шагом 0.
- В `out.md` без `--annotate` пометок нет; в `bakeoff-annot/out.md` — пометки `vision_diff`.
- Дальше то же для `bakeoff2.pdf`, `bakeoff3.pdf`, `textpdf1.pdf`. Sonnet — по желанию:
  `--vision claude-sonnet-5`.
- Лимит подписки → находка «сверка остановлена…», конвертация цела. Повтор после сброса доберёт с места
  остановки: отвеченные полосы — в кэше.

После проверки, по команде человека, — документационный коммит с разделом «Сверка по картинке в тракте:
живая проверка» в docs_sync. На каждую фикстуру:
- вызовов claude (строка stderr);
- подсказок `vision_diff`: совпали с опкодом эталона / ложные по скану / пропуск эталона;
- непокрытые опкоды;
- `vision_skipped`;
- время прогона.

Ожидаемых чисел нет, есть оценка решения 1: ≈0.7 ложной на страницу.

## Готово, когда

- `pytest -v` зелёный целиком; сеть и `claude` не тронуты; в `git diff --stat` нет файлов из «Не трогать».
- Проверка, п. 1–4, дала ожидаемое.
- В docs_sync `### ПРАВКА #91`:
  - отклонения от буквы спеки, если были;
  - таблицы Шага 0 (состав окон и покрытие) и время модуля тестов;
  - строка «`narrow` в тракте судится, в замере — служебный исход» с причиной;
  - строка «report.json v2: у находки + `reading`, `model`»;
  - снимок нумерации: «следующий — #92».
- **Стоп.** Живая проверка — человек. Спека 17 (UI) — после коммита #91.

## PLACEHOLDER-ы

1. **Непривязанные окна** — 17 % окон у textpdf1 (124 + 3), 2–3 % у сканов. Основная причина: markdown прошёл
   `postprocess` (`No` → `№`, гомоглифы, `$C^{\circ}$` → `°C`, склейки таблиц), а `locate_block` ищет по
   тексту content_list. Путь улучшения — нормализовать обе стороны при привязке; `gemini_verifier`
   заморожен, так что это отдельная правка. Такие окна честно считаются в `vision_skipped` покрытия.
2. **`narrow` судится** (отличие от замера). На фикстурах ложных подсказок это не дало (фейки), живьём — по
   проверке человека. Путь улучшения — объединённый `bbox` соседних блоков (PLACEHOLDER 1 спеки 15).
3. **`VISION_STRIDE = 3`**: окно — 9 токенов, как у опкода замера с контекстом ±3. Больше пролёт — меньше
   вердиктов, но выше риск двух склеек в окне (`ratio` < 0.5 → не прочитано).
4. **Страница у стыка** — первая из двух, даже если опкод на второй.
5. **Удаление** («у нас лишнее», `suggestion == ""`): пометка показывает `snippet` — `_annotation` берёт
   `suggestion or snippet`. В UI это «Стало: ∅» (спека 17).
6. **Время.**
   - Вердикт — ≈3 с на страницу скана на машине человека, растеризация — ≈1 с на страницу.
   - Модуль тестов — ≈7 минут (точный «эталон» на трёх сканах).
   - Путь ускорения — судить окно только по полосе, где нашлись его токены. Для этого нужен свой цикл вместо
     `run_measure`: отдельная правка, после живой проверки.
7. **`PAGE_HINT_LIMIT = 15`** — по оценке решения 1 (фон ≈0.7 ложных + 1–2.5 настоящих на страницу).
   Пересмотреть по живой проверке.
8. **Нет бинаря `claude`** обнаруживается при первом вызове, после растеризации (секунды). В UI галочка без
   `claude` не показывается (спека 17), так что это касается только CLI.
9. **Оценка ложных подсказок** держится на одном событии из 50 окон. Живая проверка на четырёх фикстурах
   (~36 страниц) даст первую настоящую долю.

## Коммит

Коммит кода и документов — один. Раздел живой проверки — отдельным документационным коммитом после неё.
Push делает человек.

```
ПРАВКА #91: сверка по картинке в тракте — весь документ полосами, подсказки в report.json

ocr/vision.py: markdown режется на окна (пролёт 3 токена, контекст ±3), окно
привязывается к блоку content_list на физической странице (*_model.json, стык —
две полосы, как в замере #90); модель (claude -p, по умолчанию claude-opus-5)
переписывает страницу, разница — локальный diff тем же judge, ocr.measure не
тронут. Гомоглифы, пунктуация и тире не подсказываются, пропущенный пробел у
знака — подсказывается. Подсказка — находка vision_diff (страница, наш фрагмент,
прочитанное, замена, модель), текст не меняется. Больше 15 подсказок на
странице — сбой нарезки, одна находка vision_skipped. Нет claude, входа, лимит,
таймаут — vision_skipped, конвертация не падает. ocr/ingest.py: параметр vision
(только scan / text_tables, не путать с verify). ocr/cli.py: --vision [МОДЕЛЬ].
report.json v2: у находки + reading, model.
```

```
docs: сверка по картинке в тракте — живая проверка на фикстурах (ПРАВКА #91)
```
