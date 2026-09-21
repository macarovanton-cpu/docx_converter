# 15 — Замер vision-сверки: физическая страница, стык страниц, разбор ответа

**# ПРАВКА #90**:
- `ocr/measure.py` — страница случая по `*_model.json`, стык страниц, исход `narrow`, приклеенный маркер, промах кэша
  не останавливает замер, причины `error`, выходы `v3`;
- `ocr/claude_code_verifier.py` — только `parse_texts` и `SYSTEM_PROMPT`.

Зависит от: 14. Наш код в сеть не ходит; в этой правке `claude` не запускается ни разу. Живой добор новых страниц
делает человек руками, после зелёного `pytest` и коммита.

## Цель

Живой замер спеки 14 (21.09.2026, `_test/verify_measure.claude-code.*.json`) мерил в основном нарезку, а не модель.
Разбор с уликами — в приложении «Диагностика 22.09.2026». Коротко:

1. **Не та страница.** MinerU склеивает межстраничную таблицу в блок первой страницы. У такого блока `table_body`
   на все страницы, а `bbox` — только первой. Блоки-продолжения на следующих страницах в `content_list.json` есть,
   с `bbox`, но с пустым `table_body`. `locate_block` находит фрагмент в склеенном блоке и режет первую страницу.
   - Итог: 48 из 100 привязанных случаев v2 вырезаны со страницы, где их текста нет. Ещё у 5 окно переходит через
     стык страниц.
   - Отсюда 8/12 `unreadable` и 0/8 склеек на bakeoff, одинаковые числа Opus и Sonnet, 36 из 40 случаев bakeoff2 на
     стр. 1 и `false_alarm` textpdf1-c11.
2. **39 `error` — наш разбор ответа.** На двух оплаченных страницах модель дописала после JSON исправленную копию.
   `parse_texts` режет от первой `{` до последней `}` и падает на `Extra data`.
3. **Окно шире текстового блока** (подписи: bakeoff-e12, e13; bakeoff2-e19, e20). Вырезка — одна строка, `ratio`
   ниже 0.5, хотя исправление модель прочитала («А.Ш.», «Е.К.»).
4. **Приклеенный маркер.** MinerU приклеивает тире списка к предыдущему токену: `DBM14G;–`, `C;–`. Единственный
   `wrong_fix` v2 (textpdf1-e10) — артефакт сравнения.
5. **Промах кэша останавливает замер.** Без `CLAUDE_CODE_LIVE` замер встаёт на первой новой странице, и оплаченные
   страницы после неё оффлайн не пересуживаются.

В `not_found` и `wrong_fix` v2 модель не ошиблась ни разу. Картинка до модели доходит без уменьшения:
PLACEHOLDER 7 спеки 14 закрыт доказательством (раздел PLACEHOLDER-ов).

**#90 чинит измерение, а не модель и не тракт.** `ingest`, `postprocess`, `validate`, `report.json`, CLI и UI о
сверке по-прежнему не знают. Интеграция — только после чисел v3 и решения человека.

Пункты отчёта (приложение) → разделы этой спеки:

| пункт отчёта | где в спеке |
|---|---|
| 1. привязка к физической странице | `model_pages`, `table_run`, `physical_place`, `build_cases` |
| 2. стык страниц | `physical_place` (`seam`), `Case.next_tiles`, `page_tiles`, `run_measure` |
| 3. `parse_texts`, `SYSTEM_PROMPT`, причины `error` | `ocr/claude_code_verifier.py`, `error_reason`, `totals.error_reasons` |
| 4. блок уже фрагмента | служебный исход `narrow` |
| 5. приклеенный маркер | `norm` |
| 6. `neighbor` в сводке | правило docs_sync (раздел «Готово») |
| 7. контроль | та же привязка; неопределённое окно — следующий кандидат |
| 8. PLACEHOLDER 7 спеки 14 | раздел PLACEHOLDER-ов, п. 8 |
| 9. сравнение Opus и Sonnet | «Живой добор», docs_sync v3 |
| 10. цена перезамера | «Живой добор» |

Уточнения при оформлении, без изменения сути:
- Стыку всегда добавляются полосы следующей страницы, поэтому служебный исход `seam` из отчёта не нужен: окно через
  стык судится всегда.
- Из двух вариантов п. 4 выбран `narrow`. Обрезка окна до двух токенов даёт `ratio` ровно 0.5 — на пороге. А текст
  `*_model.json` сырой и расходится с `content_list`, так что обрезка табличных окон по нему испортила бы найденный
  bakeoff-e10.
- Флаг `self_corrected` не заводится: ради него пришлось бы менять `transcribe`. Оба случая описаны в приложении.
- Промах кэша не останавливает замер: иначе проверку «оплаченное пересуживается бесплатно» не выполнить оффлайн.
- Выходы получают `v3` в имени: файлы v2 — улики приложения, перезаписывать их нельзя.

## Решения человека (22.09.2026, не пересматриваются)

- **Одна правка — #90.**
- **Ключи кэша не меняются — это условие приёмки.** Оплаченные ответы пересуживаются бесплатно. Поэтому не
  меняются: `TRANSCRIBE_QUESTION`, `FILES_HEADER`, `verify_cache_key` и её аргументы, `TILE_HEIGHT`, `TILE_OVERLAP`,
  `split_tiles`, `crop_block` и `CROP_*`. Полосы страниц, оплаченных в v2, остаются байт в байт прежними:
  - `bbox` для страницы, где начинается склеенная таблица, — тот же, что был;
  - для страниц-продолжений `bbox` берётся из блока-продолжения `content_list`, а не из `*_model.json`.
- **`ocr/claude_code_verifier.py` — только `SYSTEM_PROMPT` и `parse_texts`.**
  - `SYSTEM_PROMPT` в ключ кэша не входит.
  - `parse_texts` разрешён ответом человека в чате 22.09.2026: без него два оплаченных ответа навсегда остаются
    `error`, а чистить кэш нельзя.
  - Разбор делается при каждом чтении записи, так задумано в #89.
- **Живой добор новых страниц — только руками человека**, после зелёного `pytest` и коммита #90. Исполнитель
  `CLAUDE_CODE_LIVE=1` не выставляет.
- **PLACEHOLDER 7 спеки 14 закрыт** доказательством из приложения.
- **PLACEHOLDER 5 спеки 13 (строки таблиц) остаётся открытым.** Физическая страница теперь берётся из
  `*_model.json`, а строк таблицы по-прежнему нет: у табличного блока и там только `bbox` всей таблицы страницы.

## Шаг 0 — условия входа (иначе стоп)

1. `grep -rn "ПРАВКА #90" --include=*.py .` (без `.venv`) пуст; последняя правка в коде — #89.
2. `pytest -v` зелёный; `python -X utf8 -m ocr.board` оффлайн даёт шесть строк с `diffs == thr`.
3. `_test/verify_measure.claude-code.claude-sonnet-5.json` и `…claude-opus-5.bakeoff.json` (v2) на месте:
   `totals.error == 39`, `unreadable == 10`. Это «до» для docs_sync; файлы не трогать.
4. **Проба — одноразовая, в репозиторий не входит.** Алгоритм `physical_place` этой спеки на случаях текущего
   `build_cases`, текст `*_model.json` пропущен через `html.unescape`:

   | фикстура | страница верна (вкл. `narrow`) | страница `content_list` ≠ физической | окно через стык | не определено | окно шире текстового блока |
   |---|---|---|---|---|---|
   | bakeoff | 8 | 16 | 0 | 0 | e12, e13 |
   | bakeoff2 | 9 | 28 | 3: e07 (4\|5), e08 (5\|6), e11 (7\|8) | 0 | e19, e20 |
   | bakeoff3 | 12 | 2 | 2: e03 (1\|2), e05 (5\|6) | 0 | 0 |
   | textpdf1 | 17 | 2 | 0 | c11 (контроль) | 0 |

   Кроме того:
   - `parse_texts` по правилу этой спеки разбирает все 23 уникальных ответа `claude-code` в кэше без ошибок;
   - у двух ответов (bakeoff2 стр. 1, bakeoff3 стр. 3) объектов два, берётся последний.

   Другие числа — стоп с распечаткой. Алгоритм под числа не подкручивать.
5. Число страниц в `*_model.json` равно числу страниц PDF: 9, 8, 9, 10.

## Трогать

- `ocr/measure.py` — #90.
- `ocr/claude_code_verifier.py` — **только** `SYSTEM_PROMPT` и `parse_texts` (#90).
- `tests/test_ocr_measure.py` — дополнить и поправить числа (#90). Фейки `Transcriber` отвечают и за
  `case.next_tiles`.
- `tests/test_claude_code_verifier.py` — только `test_parse_texts` (#90).
- `CLAUDE.md`:
  - строка `ocr/measure.py` в списке модулей — `(#87, #88, #90)` и одна фраза о физической странице;
  - строка `ocr/claude_code_verifier.py` — «`parse_texts` — последний JSON-объект с ключами по файлам (#90)»;
  - абзац «Замер vision-сверки»:
    - страница — физическая, `*_model.json`;
    - стык — две полосы;
    - `narrow`;
    - промах кэша — не стоп, код `1` и список недостающих страниц;
    - выходы `verify_measure.v3.…`;
  - фраза о нумерации — `#90 — ocr/measure.py + ocr/claude_code_verifier.py`;
  - строка-урок: «MinerU склеивает межстраничную таблицу в блок первой страницы (продолжения — с пустым
    `table_body`): страницу фрагмента брать из `*_model.json`, а не из `content_list` (замер #89)».
- `docs/specs/ocr/README.md`:
  - строка спеки 15 в таблице;
  - сводка PLACEHOLDER-ов: п. 9 — «PLACEHOLDER 7 спеки 14 закрыт спекой 15; PLACEHOLDER 5 спеки 13 — физическая
    страница из `*_model.json` (#90), строк таблиц нет»; новый п. 10 — PLACEHOLDER-ы этой спеки.
- `docs/PLAN_OCR.md` — «Этап B, шаг 3 (#90)» и строка статуса.
- `docs/docx_converter_docs_sync.md`:
  - `### ПРАВКА #90` с таблицей «до/после»;
  - снимок «следующий — #91»;
  - после живого добора — раздел «Замер vision-сверки v3» (см. «Готово»).

## Не трогать

- `ocr/gemini_verifier.py` — целиком. Оттуда только импорт: `locate_block`, `crop_block`, `CONTEXT_TOKENS`, классы
  ошибок.
- `ocr/claude_code_verifier.py` — всё, кроме `SYSTEM_PROMPT` и `parse_texts`: `transcribe`, `_call`,
  `claude_command`, `parse_cli_output`, `FILES_HEADER`, формула ключа.
- `ocr/ingest.py`, `ocr/postprocess.py`, `ocr/validate.py`, `ocr/cli.py`, `ocr/board.py`, `ocr/diff.py`,
  `ocr/mineru_provider.py`, `ocr/cache.py`, `ocr/__init__.py`, `pdf_core.py`.
- `app.py`, `file_converter.py`, `convert.py`, `requirements.txt`, `template*.docx`, `conftest.py`,
  `tests/test_gemini_verifier.py`.
- Фикстуры, эталоны, `*.errors.txt`, zip MinerU.
- `.cache/ocr/verify/` — не чистить и руками не править.
- `_test/verify_measure*.json` и `_test/verify_review*.md` v2 — это «до». Выходы v3 пишутся под другим именем.
- В `measure.py`: `TRANSCRIBE_QUESTION`, `TILE_HEIGHT`, `TILE_OVERLAP`, `split_tiles`, `best_window`, `MIN_RATIO`,
  `WINDOW_SLACK`, таблица исходов `judge`.

## Интерфейсы (дословно)

### `ocr/claude_code_verifier.py` (#90)

```python
# ПРАВКА #90: модель дописывала после JSON исправленную копию; просим ровно один объект (в ключ кэша не входит)
SYSTEM_PROMPT = ("You transcribe document images verbatim. Follow the user instructions exactly. "
                 "Reply with exactly one JSON object and nothing after it.")


def parse_texts(raw: str, names: list[str]) -> dict[str, str]:
    """ПРАВКА #90: ответ — последний JSON-объект с ключами ровно по names (модель дописывает исправленную копию)."""
```

`SYSTEM_PROMPT` — ASCII, без перевода строки, кавычек и `& | < > ^ %`: `claude.cmd` идёт через `cmd.exe`, это
проверяет существующий тест.

**`parse_texts`**:
- Объекты ищутся так. `decoder = json.JSONDecoder()`, `i = raw.find("{")`. Пока `i != -1`:
  - `obj, end = decoder.raw_decode(raw, i)` → объект в список, `i = raw.find("{", end)`;
  - `ValueError` → `i = raw.find("{", i + 1)`.

  С `{` декодируется только объект, поэтому проверка «не `dict`» не нужна. Обёртка ```` ```json ```` пропускается
  сама.
- Объект **подходит**, если выполнены все три условия:
  - `texts = {k.replace("\\", "/").rsplit("/", 1)[-1]: v for k, v in obj.items()}`;
  - `sorted(texts) == sorted(names)` и `len(texts) == len(obj)`;
  - все значения — строки.

  Это прежнее правило #89, по одному объекту.
- Возврат — `texts` **последнего** подходящего объекта.
- Объектов нет → `VerifierError(f"ответ модели не JSON: {raw[:200]!r}")`.
- Объекты есть, подходящих нет → `VerifierError(f"ответ не по файлам: ждали {names}, пришло {list(objects[-1])}:
  {raw[:200]!r}")`.
- Начала сообщений «ответ модели не JSON» и «ответ не по файлам» — часть интерфейса: по ним `error_reason` относит
  сбой к `parse`.

### `ocr/measure.py` (#90)

```python
"""ПРАВКА #87: измерение vision-сверки на остатке эталонов. В тракт не подключено.
ПРАВКА #88: модель переписывает полосы вырезки, вердикт — локальный diff по text_tokens.
ПРАВКА #89: бэкенд — Claude Code (claude -p); --verifier / --model / --fixture.
ПРАВКА #90: страница случая — физическая (*_model.json), окно через стык страниц — по двум полосам, narrow,
приклеенный маркер, промах кэша не останавливает замер."""

import html
import re

from ocr.claude_code_verifier import CLAUDE_MODEL, ClaudeCodeTimeoutError, ClaudeCodeVerifier

MEASURE_SCHEMA_VERSION = 3
SERVICE_OUTCOMES = ("unlocated", "no_image", "narrow", "error")     # ПРАВКА #90: + narrow
PAGE_SOURCES = ("content_list", "model_json", "seam")              # ПРАВКА #90: откуда страница у случая с полосами
ERROR_REASONS = ("miss", "stop", "timeout", "cli", "parse", "count", "other")   # ПРАВКА #90
GLUED = frozenset(";:,.)")   # ПРАВКА #90: знак перед приклеенным маркером списка: "DBM14G;–", "связи).-", "°C:–"
_TAG_RE = re.compile(r"<[^>]+>")   # как в ocr/gemini_verifier.py: файл заморожен, приватное имя не импортируем


@dataclass(frozen=True)
class Case:
    ...                                   # поля #88 — без изменений и в том же порядке
    next_tiles: tuple[bytes, ...] = ()    # ПРАВКА #90: окно через стык — полосы блока следующей страницы; иначе ()
    page_source: str | None = None        # ПРАВКА #90: PAGE_SOURCES, "narrow", "unresolved"; None — блока нет
    list_page: int | None = None          # ПРАВКА #90: страница блока content_list (в v2 она и была page), 1-based


def model_pages(zip_path: Path) -> list[str]: ...
def table_run(content_list: list, index: int) -> list[int]: ...
def physical_place(needle: list[str], before: list[str], after: list[str], content_list: list, index: int,
                   pages: list[str]) -> tuple[str, list[int]]: ...
def norm(tokens: list[str]) -> list[str]: ...           # изменена
def error_reason(exc: Exception) -> str: ...
```

Новые поля `Case` — в конце и со значениями по умолчанию: синтетические случаи в тестах собираются по-старому.

**`model_pages(zip_path)`** — текст таблиц каждой страницы до склейки.
- В zip ровно один член, оканчивающийся на `_model.json`, иначе `ValueError` — как у `_content_list`.
- JSON — список страниц, индекс совпадает с `page_idx`; страница — список блоков.
- Текст страницы собирается так:
  1. `" ".join(block['content'] for block in page if block['type'] == "table")`;
  2. `_TAG_RE.sub(" ", …)`;
  3. `html.unescape`;
  4. `_squash`.
- `html.unescape` обязателен: в `*_model.json` кавычки лежат как `&quot;`. Без него bakeoff2-c13 уходит в
  «не определено».

**`table_run(content_list, index)`** — индексы блоков склеенной таблицы.
- `run = [index]`, `cur = index`.
- Дальше берётся ближайший следующий блок с `type == "table"` (не табличные блоки между ними пропускаются). Он
  входит в `run`, если одновременно:
  - `page_idx` ровно на 1 больше, чем у `cur`;
  - текст пуст: `_squash(_TAG_RE.sub(" ", block.get("table_body") or ""))`. Поле у пустого продолжения может
    отсутствовать, поэтому `.get` — как в `locate_block`.
- Не входит → конец.
- Ожидаемо: bakeoff `[7, 9, 11, 13]`, `[15]`, `[17, 19, 21]`; bakeoff2 `[3, 8, 10, 12, 14, 16, 18, 20]`; bakeoff3
  `[1, 4]`, `[23, 26]`, `[29]`; textpdf1 `[105, 108]`, `[102]`.

**`physical_place(needle, before, after, content_list, index, pages)`** → `(source, блоки)`.
`index` — индекс блока, который вернул `locate_block`; `pages` — `model_pages` фикстуры.

1. **Не таблица.**
   - `_squash(" ".join(before + needle + after))` входит в текст блока (обработка как в `locate_block`: теги →
     пробел, без `unescape`) → `("content_list", [index])`;
   - иначе → `("narrow", [index])`.
2. **Таблица, `table_run` из одного блока** → `("content_list", [index])`. `*_model.json` не спрашивается: склейки
   нет, а сырой текст может расходиться с `content_list`.
3. **Склеенная таблица.** `k` от `CONTEXT_TOKENS` до 0:
   - `nd = _squash("".join((before[-k:] if k else []) + needle + after[:k]))`; пустой `nd` → следующий `k`;
   - блоки `run`, на странице которых `nd in pages[page_idx]`, → первый из них `j`:
     `("content_list" if j == index else "model_json", [j])`;
   - иначе соседние пары `(j, j2)` из `run`, у которых `nd in pages[page_idx(j)] + pages[page_idx(j2)]`, → первая:
     `("seam", [j, j2])`.
4. Ни на одном `k` ничего → `("unresolved", [])`.

Порядок внутри одного `k` — сначала страница, потом стык. Поэтому окно, которое при полном контексте переходит на
следующую страницу, судится по двум полосам, даже если сам опкод целиком на одной.

**`build_cases`** — как в #88, отличия:
- при `zip_name is not None` ещё `pages = model_pages(fixtures / zip_name)`;
- `place(needle, before, after)` отдаёт `(page, block_type, ambiguous, tiles, next_tiles, page_source, list_page)`:
  - блока нет → `(None, None, False, (), (), None, None)`;
  - `index = next(i for i, b in enumerate(content_list) if b is block)`, затем `physical_place`;
  - `narrow` → `(list_page, type, ambiguous, (), (), "narrow", list_page)`;
  - `unresolved` → `(None, type, ambiguous, (), (), "unresolved", list_page)`;
  - иначе полосы — `split_tiles(crop_block(pdf, page_idx, bbox))` блока `блоки[0]`; у `seam` — ещё
    `next_tiles` блока `блоки[1]`; `page = page_idx + 1` блока `блоки[0]`. Кэш вырезок по `(page_idx, bbox)` —
    прежний;
- ошибка идёт в список `located` для контрольной выборки, **если у неё есть блок** (`page_source is not None`),
  даже без полос (`narrow`, `unresolved`). Так окна контроля остаются окнами v2: не меняются ни шаг выборки, ни
  число контрольных (50);
- контрольное окно без полос (`unresolved`) — следующий кандидат, как «не привязался» в #87. На 22.09.2026 это
  одно окно: textpdf1-c11.

**`norm`**:
- токен, у которого последний знак из `MARKERS`, а предпоследний из `GLUED`, делится на `token[:-1]` и
  `token[-1]`;
- дальше, как прежде, выбрасываются токены целиком из `MARKERS`;
- больше ничего не трогается: «слово-» и «-слово» остаются как есть (PLACEHOLDER 2).

`_split` и `judge` работают через `norm`, пролёты считаются по частям — как в #88.

**`error_reason(exc)`**, первое совпадение сверху:

| условие | причина |
|---|---|
| `str(exc)` начинается с «нет в кэше» | `miss` |
| `STOP_ERRORS` | `stop` |
| `ClaudeCodeTimeoutError` | `timeout` |
| текст с «claude -p:» | `cli` |
| текст с «ответов » | `count` (ответ `run_measure` про число текстов) |
| текст с «ответ модели не JSON» или «ответ не по файлам» | `parse` |
| прочее | `other` |

**`page_tiles`**:
- как в #88;
- у случая с `next_tiles` его `next_tiles` добавляются в список ключа `(fixture, page + 1)`, без повторов, в порядке
  первого появления.

**`run_measure`**:
- Сбой страницы записывается в `failures[key] = (error_reason(exc), f"{type(exc).__name__}: {exc}")`.
  - `miss` → в `missing`, **замер идёт дальше**;
  - прочие `STOP_ERRORS` → стоп, как в #89;
  - прочий `VerifierError` → `error` у страницы.
- Ключи случая: `(fixture, page)`, а у стыка ещё `(fixture, page + 1)`. Если один из них в `failures`:
  - `error` и `error_reason` берутся из первого такого ключа;
  - `outcome = None` при `stop` или `miss`, иначе `"error"`.
- Исход без полос:
  - `"no_image"` — у фикстуры нет zip;
  - `"narrow"` — `page_source == "narrow"`;
  - иначе `"unlocated"`.
- **Стык.** `last = pages[(f, p)].index(case.tiles[-1])`, `first = pages[(f, p + 1)].index(case.next_tiles[0])`;
  - `verdict = judge(case, texts_p[last] + "\n" + texts_p1[first])`;
  - `crop` — полоса `last` страницы `p`, `crop_next` — полоса `first` страницы `p + 1`;
  - `cache_hit` — оба попадания.
- Не стык — как в #88. `crop_next = None`.
- Новые поля строки: `page_source`, `list_page`, `crop_next`, `error_reason`.
- `measured` — ошибки с `page is not None` и `outcome != "narrow"`.
- `totals`:
  - `+ "narrow"` после `"unlocated"`;
  - `+ "page_sources": {источник: число строк}` по `PAGE_SOURCES`;
  - `+ "error_reasons": {причина: число строк}` по `ERROR_REASONS`.
- `by_fixture`: `+ "narrow"`, `+ "moved"` (строки с `page_source` `model_json` или `seam`).
- Верх JSON: `+ "missing_pages": [[fixture, page], …]` — страницы с `miss` в порядке `page_tiles`. Это и есть
  точный план живого добора.
- `complete = stop is None and not missing`.

**`write_review`**:
- у стыка после `![](crop)` строка `![](crop_next)`;
- в заголовке случая — `стр. N|N+1`.

**`main`**:
- имя выходов: `verify_measure.v{MEASURE_SCHEMA_VERSION}.{бэкенд}.{модель}[.<stem>+<stem>].json`, так же `verify_review…md`;
- при `complete == False`:
  - есть `stop` → прежняя строка «замер остановлен на …»;
  - есть `missing_pages` → `нет в кэше {N} стр.: {stem pNN, …} — {текст ошибки}` (в тексте есть
    «CLAUDE_CODE_LIVE»);
  - код `1`.

## Приёмочные тесты

Сеть и `claude` недоступны, как в спеке 14: `socket.socket` и `subprocess.run` — заглушки с `AssertionError`;
`CLAUDE_CODE_LIVE`, `VERIFIER`, `GEMINI_LIVE` снимаются через `delenv`.

### `tests/test_claude_code_verifier.py::test_parse_texts` (#90, дополнить)

```python
names = ["01.png", "02.png"]
# прежние случаи #89 — без изменений (обёртка ```json, полные пути, лишний/недостающий файл, нестроковое значение,
#   два ключа на одно имя, "нет json", список)
two = ('{"01.png": "утрежденные", "02.png": "б"}\n\nОдна оговорка: в 01.png я написал «утрежденные».\n\n'
       '{"01.png": "утвержденные", "02.png": "б"}')
assert parse_texts(two, names) == {"01.png": "утвержденные", "02.png": "б"}           # последний, а не срез {…}
assert parse_texts('{"01.png": "а", "02.png": "б"} и ещё {"03.png": "в"}', names) == {"01.png": "а", "02.png": "б"}
assert parse_texts('{"01.png": "a {b} c", "02.png": ""}', names) == {"01.png": "a {b} c", "02.png": ""}
with pytest.raises(VerifierError, match="ответ модели не JSON"):
    parse_texts("нет json", names)
with pytest.raises(VerifierError, match="ответ не по файлам"):
    parse_texts('{"01.png": "а"} {"02.png": "б"}', names)
```

### `tests/test_ocr_measure.py` (#90)

```python
# состав: 63 опкода = 46 с полосами + 4 narrow + 4 unlocated + 9 no_image; контроль — 50, как в v2
located = [c for c in errors if c.tiles]
narrow = [c for c in errors if c.page_source == "narrow"]
unlocated = [c for c in errors if not c.tiles and c.fixture.endswith(".pdf") and c.page_source != "narrow"]
assert len(located) == 46 and len(unlocated) == 4 and sum(c.ambiguous for c in located) == 0
assert {c.id for c in narrow} == {"bakeoff-e12", "bakeoff-e13", "bakeoff2-e19", "bakeoff2-e20"}
assert all(c.page is not None and not c.tiles and not c.next_tiles for c in narrow)
assert Counter(c.error_kind for c in located) == {"merge": 31, "homoglyph": 8, "chars": 7}
assert {t: sum(c.block_type == t for c in located) for t in {c.block_type for c in located}} == {
    "table": 44, "text": 1, "header": 1}
assert len(controls) == 50 and all(c.tiles for c in controls)
assert Counter(c.page_source for c in located) == {"content_list": 20, "model_json": 21, "seam": 5}
seams = [c for c in located if c.page_source == "seam"]
assert {c.id for c in seams} == {"bakeoff2-e07", "bakeoff2-e08", "bakeoff2-e11", "bakeoff3-e03", "bakeoff3-e05"}
assert all(c.next_tiles and c.error_kind == "merge" for c in seams)
assert all(not c.next_tiles for c in cases if c.page_source != "seam")
assert sum(c.page_source == "model_json" for c in controls) >= 27    # 8 bakeoff + 17 bakeoff2 + 2 bakeoff3
assert all(c.page_source in PAGE_SOURCES for c in controls)            # unresolved-окно заменено следующим
assert next(c for c in controls if c.id == "textpdf1-c11").fragment != "+50 Диапазон температуры для приборов °C:– М0601"
page = {c.id: c.page for c in cases}                                   # регрессия v2: физические страницы
assert (page["bakeoff-e02"], page["bakeoff-c13"], page["bakeoff2-e12"], page["textpdf1-e11"]) == (3, 9, 8, 9)
assert (page["bakeoff3-e05"], next(c for c in cases if c.id == "bakeoff-c05").list_page) == (5, 2)
# «после»: текст каждого перенесённого случая есть на его странице *_model.json (у стыка — на стыке двух)
for c in located + controls:
    if c.page_source in ("model_json", "seam"):
        mp = model_pages(FIXTURES / dict(BOARD)[c.fixture])
        text = mp[c.page - 1] + (mp[c.page] if c.page_source == "seam" else "")
        tokens = c.fragment.split()
        needle = "".join(tokens if c.kind == "control" else tokens[slice(*c.span)])
        assert not needle or needle in text, c.id
assert build_cases(outputs) == cases                                   # детерминизм

# table_run / model_pages на настоящих zip (cl — _content_list(FIXTURES / zip) фикстуры)
assert table_run(cl["bakeoff2.pdf"], 3) == [3, 8, 10, 12, 14, 16, 18, 20]
assert table_run(cl["bakeoff.pdf"], 17) == [17, 19, 21] and table_run(cl["textpdf1.pdf"], 102) == [102]
assert [len(model_pages(FIXTURES / dict(BOARD)[n])) for n in ("bakeoff.pdf", "bakeoff2.pdf", "bakeoff3.pdf",
                                                               "textpdf1.pdf")] == [9, 8, 9, 10]

# norm и judge
assert norm(["DBM14G;–", "связи).-", "°C:–", "помещений-", "-АКЗ", "–"]) == [
    "DBM14G;", "связи).", "°C:", "помещений-", "-АКЗ"]
e10 = Case(id="t-e10", fixture="textpdf1.pdf", kind="error", tag="replace", error_kind="homoglyph",
           fragment="740, DHM9B, DBM14G;– MB150; C;– H4, M100;– ZSFY, ZSFY-D,",
           expected="740, DHM9B, DBM14G;– МВ150; С;– Н4, M100;– ZSFY, ZSFY-D,", span=(3, 6), gold_span=(3, 6),
           rules=(), page=8, block_type="table", ambiguous=False, tiles=())
r = judge(e10, "740, DHM9B, DBM14G;\n– MB150; C;\n– H4, M100;\n– ZSFY, ZSFY-D,")
assert r["outcome"] == "not_found" and r["ops"] == [] and r["clipped"] == 0      # в v2 здесь был wrong_fix

# error_reason
assert [error_reason(e) for e in (
    VerifierConfigError("нет в кэше, вызов выключен: нужен CLAUDE_CODE_LIVE=1"), VerifierQuotaError("лимит"),
    ClaudeCodeTimeoutError("claude -p: нет ответа"), VerifierError("claude -p: ответ не JSON: …"),
    VerifierError("ответов 1, картинок 2"), VerifierError("ответ модели не JSON: …"),
    VerifierError("ответ не по файлам: …"), VerifierError("прочее"))] == [
    "miss", "stop", "timeout", "cli", "count", "parse", "parse", "other"]

# полный прогон с фейками (Transcriber отвечает и за next_tiles: полоса -> тексты всех случаев, у кого она в
#   tiles или next_tiles)
echo = Echo(cases)
m = run_measure(cases, echo)
assert len(echo.calls) == len(page_tiles(cases)) and [len(x) for x in echo.calls] == [len(v) for v in page_tiles(cases).values()]
assert m["totals"]["not_found"] == 46 and m["totals"]["agree"] == 50 and m["totals"]["narrow"] == 4
m = run_measure(cases, Golden(cases))
assert m["totals"]["found"] == 46 and m["totals"]["false_alarm"] == 0 and m["totals"]["neighbor"] == 0
sources = m["totals"]["page_sources"]                                  # 46 ошибок + 50 контрольных с полосами
assert sources["seam"] == 5 and sources["model_json"] >= 48 and sum(sources.values()) == 96
row = next(r for r in m["cases"] if r["id"] == "bakeoff3-e05")
assert row["outcome"] == "found" and row["crop"].startswith("verify_crops/bakeoff3/p05-") and row["crop_next"].startswith("verify_crops/bakeoff3/p06-")
# промах кэша на 2-й странице (фейк бросает VerifierConfigError("нет в кэше, …")): не стоп
#   complete False; missing_pages == [ключ 2-й страницы]; вызовов == len(page_tiles(cases)) — спрошены все страницы;
#   у случаев 2-й страницы (и стыков на неё) outcome None, error_reason "miss"; totals.error_reasons["miss"] > 0
# лимит на 3-м вызове: как в #89 (complete False, у 3-й страницы error_reason "stop", дальше None);
#   у случаев двух первых страниц исход есть — стыков на 3-ю среди них нет
# VerifierError на одной странице: outcome "error", error_reason "other", замер complete
# main с подменой (Golden): код 0, файлы verify_measure.v3.fake.golden.json и verify_review.v3.fake.golden.md;
#   --fixture bakeoff.pdf -> verify_measure.v3.fake.golden.bakeoff.json; в review у стыка две картинки
# main с настоящим ClaudeCodeVerifier на холодном кэше без CLAUDE_CODE_LIVE: код 1, «нет в кэше» и
#   «CLAUDE_CODE_LIVE» в stderr, subprocess.run не вызывался, missing_pages == все страницы
```

`Golden` на реальных данных обязан дать `found == 46`: 50 опкодов v2 минус 4 `narrow`, 5 стыков включительно.
Меньше — распечатать случаи и остановиться: значит, стык или страница собраны неверно, и живой добор даст то же
самое. Алгоритм под тест не подкручивать.

## Проверка

1. Шаг 0 — числа пробы совпали.
2. `pytest -v` зелёный целиком.
3. **Оффлайн, без `CLAUDE_CODE_LIVE`** (бесплатно, `claude` не запускается):
   `Remove-Item Env:CLAUDE_CODE_LIVE -ErrorAction SilentlyContinue; python -X utf8 -m ocr.measure`.
   Ожидаемо:
   - код `1`, в stderr — «нет в кэше … стр.»;
   - `missing_pages` в `_test/verify_measure.v3.claude-code.claude-sonnet-5.json`: bakeoff 3, 4, 5, 8, 9;
     bakeoff2 2–8; bakeoff3 2, 6; textpdf1 9. Плюс, возможно, страница окна, заменившего textpdf1-c11 — её
     распечатать;
   - все случаи на страницах, оплаченных в v2, судятся с `cache_hit: true`. Это и есть проверка «ключи кэша не
     изменились»: промах на странице v2 — стоп, найти, что поменяло полосу;
   - bakeoff3-e04 → `found`, bakeoff3-c03 → `agree`, textpdf1-e10 → `not_found`, у случаев bakeoff2 стр. 1 есть
     исходы (не `error`);
   - в `totals.error_reasons` нет `parse`.
4. То же для Opus: `python -X utf8 -m ocr.measure --model claude-opus-5 --fixture bakeoff.pdf`. Код `1`,
   `missing_pages` — bakeoff 3, 4, 5, 8, 9; стр. 2, 6, 7 — из кэша.
5. `git diff --stat` — только файлы из «Трогать». `git diff ocr/claude_code_verifier.py` — только `SYSTEM_PROMPT`
   и `parse_texts`.

## Живой добор (только руками человека, после зелёного `pytest` и коммита #90)

Исполнитель эти команды не запускает. Вызовы тратят подписку человека. Уже оплаченные полосы берутся из кэша: в
пакет страницы уходят только недостающие.

```powershell
$env:CLAUDE_CODE_LIVE = "1"; python -X utf8 -m ocr.measure
$env:CLAUDE_CODE_LIVE = "1"; python -X utf8 -m ocr.measure --model claude-opus-5 --fixture bakeoff.pdf
Remove-Item Env:CLAUDE_CODE_LIVE
```

- Sonnet — около 15 вызовов (+1, если замена textpdf1-c11 легла на новую полосу).
- Opus — 5 вызовов: bakeoff стр. 3, 4, 5, 8, 9. В отчёте стояло 4; пятая — продолжение таблицы на стр. 9 (c13).
- Точный список — `missing_pages` оффлайн-прогона. Цена по прайсу — около $0.05 за вызов Sonnet и $0.13 за вызов
  Opus; реально платит подписка.
- Лимит подписки → `complete: false`, код `1`; повтор после сброса доберёт с места остановки.
- Перед добором человек смотрит новые полосы в `_test/verify_crops/`: они пишутся без модели. Полоса мимо места —
  стоп раньше любого вызова. Особенно стоит проверить стыки: низ страницы N и верх N+1.

## Готово, когда

- `pytest -v` зелёный целиком; сеть и `claude` не тронуты; в `git diff --stat` нет файлов из «Не трогать».
- Проверка, п. 3–5, дала ожидаемое.
- В docs_sync `### ПРАВКА #90`:
  - отклонения от буквы спеки, если были;
  - таблица **«до/после» по привязке к странице**: фикстура → случаев с полосами → «до»: страница `content_list` ≠
    физической → «до»: окно через стык → «после»: расхождений (0, по тесту «после») → `narrow` → окно контроля
    заменено;
  - строка о том, что оплаченные страницы v2 пересуждены из кэша (`cache_hit` у всех), `error` v2 → 0;
  - снимок нумерации: «следующий — #91».
- **Стоп.** Живой добор — человек.
- После добора, по команде человека, — документационный коммит с разделом «Замер vision-сверки v3» в docs_sync:
  - главная таблица по **обеим** моделям: бэкенд и модель → вызовов → входных токенов на вызов → `found` /
    `neighbor` / `not_found` / `wrong_fix` → `agree` / `false_alarm` / `unreadable` → `narrow` → `error`;
  - `by_kind`, `by_block_type`, `error_reasons`;
  - сравнение Sonnet и Opus на bakeoff — только по случаям с полосами;
  - список `false_alarm` со ссылкой на `verify_review.v3…md`.
  - **Головная доля «видит ли модель ошибку»** = `found / opcodes` по `merge` + `chars` из `by_kind`. `neighbor`
    в ней считается как `not_found`: пролёт опкода не задет. Гомоглифы идут отдельной строкой — это догадка, а не
    чтение (PLACEHOLDER 10 спеки 14).
  - Ожидаемых чисел нет.
- Интеграция не начинается, пока человек не посмотрел числа v3.

## PLACEHOLDER-ы

1. **`narrow`** — 4 случая, все инициалы в блоке подписей. Путь улучшения — вырезать объединённый `bbox` соседних
   текстовых блоков страницы, покрывающих окно. Отдельной правкой.
2. **Маркер.** Отделяется только после `; : , . )`. «Слово-» (bakeoff e05/e06, `помещений-`: перенос или
   приклеенный маркер) и «-слово» (bakeoff2-c09, `-АКЗ`) не трогаются — неоднозначно. Решать по сверке v3.
3. **`unresolved`.** Сырой текст `*_model.json` может разойтись с `content_list` так, что окно не найдётся ни на
   одной странице склеенной таблицы. Тогда ошибка → `unlocated`, контроль → следующий кандидат. На 22.09.2026 такое
   окно одно: textpdf1-c11.
4. **Стык.** Текст стыка ищется внизу последней полосы блока на N и вверху первой полосы блока на N+1. Сверяется
   глазами по полосам стыков до добора.
5. **Продолжение таблицы** определяется по пустому `table_body` у табличного блока на следующей странице. Пустая
   таблица, которую MinerU просто не прочитал, тоже попадёт в `run`. Тогда решает `*_model.json`: текста там нет —
   страница не выбирается.
6. **`SYSTEM_PROMPT`.** Уберёт ли он исправленную копию, покажет добор. `parse_texts` выдерживает оба варианта.
7. **PLACEHOLDER 5 спеки 13 (строки таблиц) — открыт.** Физическая страница теперь из `*_model.json` (#90). Строк
   нет и там: у табличного блока страницы только `bbox` всей таблицы.
8. **PLACEHOLDER 7 спеки 14 (масштабирование при чтении) — закрыт.**
   - По документации Claude Code Read уменьшает и пережимает картинку, только если она больше предела модели; после
     этого файл больше 500 КБ перекодируется в JPEG.
   - По документации API у моделей Claude 4.7 и новее (Sonnet 5, Opus 5) предел — 2576 px по длинной стороне и
     4784 визуальных токена; цена картинки — `⌈w/28⌉×⌈h/28⌉`.
   - Полоса 1467×700 = 53×25 = 1325 токенов, самая тяжёлая — 255 678 Б. Значит, ни уменьшения, ни JPEG.
   - Живые вызовы с одной полосой (textpdf1, Sonnet) совпали с формулой до токена:

     | полоса | размер | `⌈w/28⌉×⌈h/28⌉` | `cacheCreationInputTokens` | Δ к p02 | Δ по формуле |
     |---|---|---|---|---|---|
     | p02 | 1421×210 | 51×8 = 408 | 1493 | — | — |
     | p09 | 1421×168 | 51×6 = 306 | 1391 | −102 | −102 |
     | p06 | 1426×554 | 51×20 = 1020 | 2105 | +612 | +612 |
     | p05 | 1424×210 | 51×8 = 408 | 1496 | +3 | 0 (шум вызова) |

   - Жалоб на мелкий или размытый текст нет ни в одном из 23 уникальных ответов.
   - `TILE_HEIGHT` / `TILE_OVERLAP` не меняются: полосы читаются, а новая высота — это новые ключи и повторная
     оплата.
9. **Opus против Sonnet на v2 для сравнения моделей непригодны**: 18 из 24 исходов bakeoff не зависели от модели.
   Сравнение — только на v3.

## Коммит

Коммит кода и документов — один. Таблица замера v3 — отдельным документационным коммитом после живого добора
человеком. Push делает человек.

```
ПРАВКА #90: замер vision-сверки — физическая страница, стык страниц, разбор ответа

ocr/measure.py: MinerU склеивает межстраничную таблицу в блок первой страницы,
и v2 резал полосу не с той страницы (48 из 100 случаев, ещё 5 — через стык).
Страница фрагмента теперь берётся из *_model.json (таблица постранично, до
склейки), bbox — из блока-продолжения content_list: полосы оплаченных страниц
не меняются, ключи кэша те же. Окно через стык страниц судится по последней
полосе страницы и первой полосе следующей. Окно шире текстового блока —
служебный исход narrow. Маркер списка, приклеенный к токену после ;:,.) —
отделяется до сравнения. Промах кэша без CLAUDE_CODE_LIVE не останавливает
замер: оплаченные страницы пересуживаются, недостающие — в missing_pages.
Причины error — в totals. Выходы — verify_measure.v3.…
ocr/claude_code_verifier.py: parse_texts берёт последний JSON-объект с ключами
по файлам (модель дописывала исправленную копию — 39 error в v2); SYSTEM_PROMPT
просит ровно один объект. Тракт не тронут.
```

```
docs: замер vision-сверки v3 — Sonnet и Opus после ПРАВКИ #90
```

---

## Приложение. Диагностика 22.09.2026

Отчёт диагностики замера спеки 14. Без вызовов `claude` и без правок кода: только кэш `.cache/ocr/verify/`, полосы
`_test/verify_crops/`, JSON и MD замера и zip MinerU фикстур; всё посчитано в памяти. Раздел «Правки в спеку 15»
отчёта стал телом этой спеки (соответствие — в «Цели»).

Одна поправка к числам. Первая проба искала текст в `*_model.json` без `html.unescape` и дала у bakeoff2 «27
подтверждено, 4 не определено». Проба Шага 0 с `unescape` даёт 28 перенесённых, 3 окна через стык и 0 не
определённых: не определённые оказались стыками (e07, e08, e11) и кавычкой `&quot;` (c13).

### Контекст

Живой замер #88/#89 (21.09.2026, `_test/verify_measure.claude-code.*.json`) дал три странности:
- 39 `error`, из них 37 из 40 у bakeoff2;
- 8/12 `unreadable` и 0/8 склеек на bakeoff, причём числа Opus и Sonnet одинаковые;
- `not_found` и `wrong_fix` на bakeoff3 и textpdf1.

Главная находка объясняет большую часть всех трёх: **MinerU склеивает межстраничные таблицы**.
- В `content_list.json` весь текст таблицы лежит в блоке первой страницы. У bakeoff2 это блок #3 на стр. 1:
  16 923 знака на стр. 1–8.
- Блоки-продолжения на следующих страницах есть, с `bbox`, но с пустым `table_body`.
- `locate_block` находит фрагмент в склеенном блоке → страница и `bbox` первой страницы → на полосе текста нет.
- Физическую страницу видно в `*_model.json` из того же zip: там таблица лежит постранично, до склейки.

### 1. `error` = 39: разложение по причинам

| причина | случаев | страниц (вызовов) | улики |
|---|---|---|---|
| таймаут (`ClaudeCodeTimeoutError`) | 0 | 0 | самый долгий вызов 27 с при лимите 600 с |
| не-JSON в `result` (`parse_cli_output`) | 0 | 0 | все 16 вызовов `success`, нужная модель есть в `modelUsage` |
| **ошибка разбора ответа (`parse_texts`)** | **39** | **2 из 16** | bakeoff2 стр. 1 → 37 случаев; bakeoff3 стр. 3 → 2 случая |
| ошибка сопоставления полосы с опкодом | 0 (как `error`) | — | но 36 из 40 случаев bakeoff2 привязаны к чужой странице (ниже) |
| исключение в нашем коде | 0 | — | — |

Механика разбора. После первого JSON модель дописала пояснение и **второй, исправленный** объект:
- bakeoff2 стр. 1: `…"}` + `Correction: I made an error in the 03.png entry, so here is the corrected JSON.` + `{…}`;
- bakeoff3 стр. 3: `…"}` + «Одна оговорка: в 01.png я написал «утрежденные», хотя на картинке, судя по всему,
  напечатано «утвержденные»…» + `{…}`.

`parse_texts` делает срез `raw[index("{"):rindex("}")+1]`, захватывает оба объекта, и `json.loads` падает с
`Extra data`. Каждый объект сам по себе валиден и по ключам совпадает с `batch_files`. Тексты объектов различаются
в одной полосе из четырёх (bakeoff2 `p01-03`) и в одной из трёх (bakeoff3 `p03-01`). Правильный ответ —
**последний** объект. Ответы лежат в кэше целиком, поэтому правка разбора бесплатна: так и задумано в #89.

Если пересудить эти ответы в памяти по последнему объекту:
- bakeoff3: e04 → `found` («городе» → «городке»), c03 → `agree`;
- bakeoff2 (стр. 1): `found` 2, `wrong_fix` 2, `not_found` 14, `agree` 2, `unreadable` 17.

У bakeoff2 правка разбора лишь превращает `error` в `not_found`/`unreadable`: текста этих случаев на полосах
стр. 1 нет.

**Почему у bakeoff2 6 полос.** Полосы режутся по парам «страница, блок», в которых есть случаи. У bakeoff2 случаи
нашлись всего на двух страницах: стр. 1 (блок шапки 539×247 и 3 полосы таблицы) и стр. 8 (2 текстовых блока).

| фикстура | стр. | стр. со случаями | полос | склейка таблиц в `content_list` | случаев на чужой странице (по `*_model.json`) |
|---|---|---|---|---|---|
| bakeoff | 9 | 4 (2, 6, 7, 9) | 15 | стр. 2 ← 3, 4, 5; стр. 7 ← 8, 9 | 16 из 24 |
| bakeoff2 | 8 | 2 (1, 8) | 6 | стр. 1 ← 2…8 (вся таблица) | 27 из 40 (ещё 4 табличных не определены; на своей странице 9) |
| bakeoff3 | 9 | 6 | 21 | стр. 1 ← 2; стр. 5 ← 6 | 2 из 16 (и 2 склейки на стыке страниц, см. п. 3) |
| textpdf1 | 10 | 8 | 15 | стр. 8 ← 9 | 3 из 20 (e11, e12, c11) |

`split_tiles` работает верно: полосы покрывают `bbox` блока целиком, с перекрытием. Неверен сам `bbox` — это
первая страница склеенной таблицы.

**Вывод:** 39 `error` — наш разбор ответа (модель дописала исправленную копию JSON), а не таймаут и не модель;
6 полос у bakeoff2 — **нарезка**, точнее привязка блока к странице, а не `split_tiles`.

### 2. `unreadable` на bakeoff.pdf

**Кэш — не один ответ под двумя именами.**
- Записей 15 у Opus и 15 у Sonnet, ключи разные (модель входит в ключ как `claude-code:<модель>`).
- `raw` различается на 3 из 4 страниц: sha 5782…/4e5e… (стр. 2), bf20…/d5fb… (стр. 6), 53b6…/73f7… (стр. 7).
- На стр. 9 ответ совпал: 51 символ, две строки подписи. Обе модели ответили одинаково сами.
- Время создания разное: Sonnet 21:29–21:30, Opus 21:37–21:38 (UTC).
- Видны мелкие различия чтения, например `С°` (Opus) и `C°` (Sonnet) в e06.

| группа (стр. привязки) | случаи | исход (Sonnet = Opus) | физ. стр. (`*_model.json`) | текст на полосе | что прочитала модель |
|---|---|---|---|---|---|
| склейки e02–e09 (стр. 2) | 8 | `not_found` | 3 | нет | верный текст стр. 2: «№ п/п Перечень основных», «3 Инвестор (при наличии)» |
| контроль c02–c06 (стр. 2) | 5 | `unreadable` | 3, 3, 4, 5, 5 | нет | верный текст стр. 2 |
| контроль c10, c12, c13 (стр. 7) | 3 | `unreadable` | 8, 8, 9 | нет | «ИТСО. 26 Требования к», «помех. 3.» — верно по полосе |
| e12, e13 (стр. 9, текст) | 2 | `not_found` | 9 | частично: блок — одна строка подписи, фрагмент шире блока | «А.Ш. Ямалов», «Е.К. Кустова» — **исправление увидено**, но `ratio` 0.22 / 0.29 < 0.5 |
| e01, e10 | 2 | `found` | свои | да | «сыпучими», «IP-камерами» |
| c01, c07, c08, c09 | 4 | `agree` | свои | да | дословно |

- **Глазами** (`bakeoff/p02-01.png`, `p07-04.png` против `bakeoff3/p01-01.png`): чистый скан, кегль около 12 пт,
  строка ~35 px при 200 dpi, читается без усилий. Качество такое же, как у bakeoff3. На `p07-04` внизу «Лист 7 из»,
  а фрагмента c13 («оформлением протокола. 3. Необходимые корректировки», стр. 9) на полосе нет.
- **Масштабирование при чтении (PLACEHOLDER 7):**
  - В документации Claude Code Read сказано: «resizes and recompresses large images to fit the model's image size
    limits»; картинка больше 500 КБ после этого перекодируется в JPEG.
  - В документации API для моделей Claude 4.7+ (сюда входят Sonnet 5 и Opus 5) предел — 2576 px по длинной
    стороне и 4784 визуальных токена; цена картинки — `⌈w/28⌉×⌈h/28⌉`.
  - Полоса 1467×700 — это 53×25 = 1325 токенов, ниже обоих пределов. Самая тяжёлая полоса — 255 678 Б, меньше
    500 КБ. Значит, ни уменьшения, ни JPEG.
  - Токены вызовов это подтверждают точно. Вызовы с одной картинкой (textpdf1):
    - p02 1421×210 = 408 ток. → `cacheCreation` 1493;
    - p09 1421×168 = 306 ток. → 1391, Δ −102 при ожидаемых −102;
    - p06 1426×554 = 1020 ток. → 2105, Δ +612 при ожидаемых +612.

    Модель получает картинку в родном разрешении.
  - В 23 уникальных ответах нет ни одной жалобы на мелкий или размытый текст. Единственная оговорка — модель сама
    нашла у себя описку («утрежденные»), это не вопрос зрения.

**Вывод:** **нарезка**, а точнее привязка блока к странице. 16 из 24 случаев bakeoff вырезаны со страницы, где
их текста нет; ещё 2 — из блока уже фрагмента. Ни картинка, ни масштабирование, ни модель тут ни при чём.
Одинаковые числа у Opus и Sonnet — следствие того же: исход 18 из 24 случаев не зависит от модели, а
оставшиеся 6 обе прочитали одинаково верно.

### 3. `not_found` и `wrong_fix` на измеренном (bakeoff3, textpdf1)

| случай | вид | наш фрагмент (опкод) | эталон | модель написала | исход | что произошло | класс |
|---|---|---|---|---|---|---|---|
| bakeoff3-e03 | merge | `инвестиционный проектE.E1000624». 12.` | `проект E.E1000624»` | `p01-04`: «…ведомостью объемов работ.» инвестиционный проект» — конец полосы | `not_found` (clipped 4) | «проект» — последняя строка стр. 1, «E.E1000624». 12. Требования к» — первая строка стр. 2. Склейка на стыке страниц | нарезка |
| bakeoff3-e05 | merge | `Подрядчик обязанпередать Заказчику` | `обязан передать` | `p05-04`: «…До начала производства работ Подрядчик обязан» — конец | `not_found` (clipped 4) | «передать…» — первая строка стр. 6. Стык страниц | нарезка |
| textpdf1-e07 | homoglyph | `Микросим M0601` (лат.) | `М0601` (кир.) | `Микросим M0601 Микросим M0808 ТП-4 1 2` | `not_found`, ratio 1.0 | прочитано как у нас: гомоглиф на картинке неотличим | не увидела (по определению) |
| textpdf1-e09 | homoglyph | `M70` | `М70` | `BM14G, HM14H1, BM14K, M70, 740, DHM9B,` | `not_found` (clipped 1) | гомоглиф; `clipped` дал приклеенный `DBM14G;–` | не увидела (по определению) |
| textpdf1-e10 | homoglyph | `MB150; C;– H4,` | `МВ150; С;– Н4,` | `DBM14G; MB150; C; H4, M100;` (у модели «–» — маркер новой строки, `norm` его выбросил) | **`wrong_fix`** | гомоглиф не виден, это ожидаемо. `wrong_fix` дали тире, которые MinerU приклеил к предыдущему токену: `C;–` против `C;`. Без этого был бы `not_found` | **сравнение** |
| textpdf1-e11 | homoglyph | `ТИТАН H22С` | `Н22С` | «Таблица 7 Технические характеристики», ratio 0.0 | `not_found` | текст на стр. 9, полоса стр. 8 (склейка таблицы 8 ← 9) | нарезка |
| textpdf1-e12 | merge | `ширина 406 Масса` | `40 6` | то же, ratio 0.0 | `not_found` | стр. 9 | нарезка |

Рядом, в том же объёме:
- bakeoff3-e04 (`error`) после правки разбора даёт `found`: «городе» → «городке».
- textpdf1-e06 `neighbor`: модель вставила «Значение». Это шапка с rowspan/colspan, у MinerU другой порядок
  ячеек. Сам опкод `M0601` прочитан латиницей — не увидела. Класс — сравнение.
- textpdf1-c11 `false_alarm`: фрагмент со стр. 9, полоса стр. 8, где есть похожая строка «Диапазон температуры».
  Класс — нарезка.
- bakeoff3-c02, c06 `unreadable`: физически стр. 2 и 6. Класс — нарезка.

**Вывод:** модель не ошиблась ни разу — ни одного случая «увидела и написала иначе». 4 из 7 — **нарезка**
(2 стыка страниц, 2 чужая страница). 3 из 7 — гомоглифы, неотличимые на картинке по определению. Единственный
`wrong_fix` — артефакт **сравнения** (тире, приклеенное к токену).
