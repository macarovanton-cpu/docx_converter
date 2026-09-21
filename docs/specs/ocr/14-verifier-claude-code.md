# 14 — Vision-сверка: транскрипция вместо вердикта и бэкенд Claude Code

**# ПРАВКА #88** (протокол `Verifier` → транскрипция; `ocr/measure.py`: полосы, локальный вердикт, отчёт для сверки),
**# ПРАВКА #89** (`ocr/claude_code_verifier.py`, выбор бэкенда в `measure`).
Зависит от: 13. Сеть есть только у подпроцесса `claude` при живом прогоне с `CLAUDE_CODE_LIVE=1`. Наш код в сеть
не ходит вообще: ни `requests`, ни сокетов, ни `api.anthropic.com`.

## Цель

Первый замер (спека 13, docs_sync «Замер vision-сверки (21.09.2026)») остановился на `bakeoff.pdf`: кончилась квота
Gemini. И он показал, что измерение мерит не то:

1. строгое сравнение `correction` с эталоном засчитало верные исправления ложными: 2 из 3 `wrong_fix` у 2.5-flash,
   4 из 4 у 3-flash;
2. вырезка «вся таблица страницы» дала 5 `unreadable` из 9 контрольных;
3. окно ±3 токена захватывает соседний опкод; модель отвечает одним словом вместо фрагмента и добавляет ведущее «- »;
4. `confidence` как порог не работает;
5. все 7 `not_found` — склейки на переносах строк.

**#88 чинит измерение.** Модель больше не выносит вердикт по нашему тексту — она **переписывает вырезку дословно**.
Вердикт считается локально: diff по `text_tokens` между транскрипцией и фрагментом тракта. Нашего текста модель не
видит, поэтому соглашаться ей не с чем. Склейки на переносах видны сами: перенос строки в транскрипции — это пробел.

**#89 даёт второй бэкенд** — Claude Code в неинтерактивном режиме — и прогон на нём: все 63 опкода (50 привязанных,
4 `unlocated`, 9 `no_image`) и 50 контрольных на Sonnet; `bakeoff.pdf` — ещё и на Opus. Это измерение, а не порог:
ожидаемых чисел нет.

**Тракт не трогается.** `ingest`, `postprocess`, `validate`, `report.json`, CLI и UI о сверке не знают. Интеграция —
после чисел этой спеки.

Остаток по видам (проба 21.09.2026, 50 привязанных опкодов):

- склейка `merge` — 31 (`IP65не`, `комплексом).В`);
- гомоглиф `homoglyph` — 11 (`A.C.`→`А.С.`, `M0601`→`М0601`);
- прочие буквы `chars` — 8 (`сыручими`, `IR-камерами`).

Гомоглиф на картинке **неотличим** по определению: «найдено» на нём — языковая догадка модели, а не чтение. Поэтому
исходы в отчёте разложены и по видам.

## Решения человека (21.09.2026, не пересматриваются)

- **Бэкенд — Claude Code.** Сверка идёт через официальный неинтерактивный режим `claude -p` под авторизацией, которая
  уже есть у Claude Code на машине. В нашем коде нет OAuth-токенов, API-ключей, шлюзов и HTTP-вызовов к
  `api.anthropic.com` — только `subprocess`. Это «Agent SDK в собственном проекте» по Help Center Anthropic;
  приложение личное. `claude-agent-sdk` в `requirements.txt` **не** добавляется.
- **Только локально.** На Streamlit Cloud этого бэкенда нет и не будет: там нет ни бинаря `claude`, ни авторизации
  подписки. `app.py` и тракт о нём не знают. `python -m ocr.measure` — инструмент человека и агента на локальной машине.
- **Gemini-путь остаётся в коде как есть, в новый замер не подключается.** `ocr/gemini_verifier.py` не меняется ни
  строкой. `GeminiVerifier.verify` (#86) отвечает только на вопрос с вердиктом, транскрипции у него нет;
  `--verifier gemini` / `VERIFIER=gemini` завершаются кодом `1` с этим текстом. 34 ответа Gemini в
  `.cache/ocr/verify/` остаются валидными для своих модели и вопроса; новые вопросы дают новые ключи.
- **Одна страница — один вызов `claude -p`.** По живому прогону человека накладные расходы одного вызова — ~46k входных
  токенов (системный промт Claude Code и содержимое рабочего каталога) при 2 токенах вопроса. Пакет по странице даёт
  ~20 вызовов вместо ~100. Возражение спеки 13 против пакета («модель видит список, где часть заведомо с ошибками»)
  снимается транскрипцией: модель видит только картинки, отличить ошибку от контроля ей не по чему. Кэш ведётся
  **по вырезке**, не по вызову (см. #89).
- **Таблицы режутся на горизонтальные полосы с перекрытием**, не на строки и ячейки. PLACEHOLDER 5 спеки 13 остаётся
  открытым, обоснование ниже.
- **PLACEHOLDER 9 спеки 13 закрыт:** вырезки документов заказчиков разрешено отправлять во внешние модели — так же,
  как документы в MinerU.
- Каталог кэша `.cache/ocr/verify/` и формула `verify_cache_key` не меняются.

**Почему строк таблиц нет** (ответ на «как привязать опкод к строке по `content_list`»):

- у табличного блока в `content_list.json`, `layout.json` и `*_model.json` есть только `bbox` всей таблицы;
- строки `table_body` не совпадают с нарисованными: у 3 из 15 табличных блоков со случаями вся страничная таблица
  записана одним `<tr>`;
- поиск горизонтальных линеек проекцией яркости на сканах (перекос, тонкие линии) нашёл правильное число строк в 0–2
  блоках из 15 (проба 21.09.2026).

Привязка «опкод → строка» потребовала бы распознавания геометрии таблицы. Это отдельная спека — и только если по
числам полосы в `table` окажутся заметно хуже, чем в `text`.

## Шаг 0 — условия входа (иначе стоп)

1. `grep -rn "ПРАВКА #8[89]" --include=*.py .` (без `.venv`) пуст; последняя правка в коде — #87.
2. `pytest -v` зелёный; `python -X utf8 -m ocr.board` оффлайн даёт шесть строк с `diffs == thr`.
3. Проба — одноразовая, в репозиторий не входит; повторяет пробу этой спеки на черновиках `_test/board/*/out.md`:
   - у 50 привязанных опкодов `error_kind`: `merge` 31, `homoglyph` 11, `chars` 8;
   - полосы (`TILE_HEIGHT` 700, `TILE_OVERLAP` 200, режим `L`): самая большая — **249 КБ**, меньше 500 КБ. Выше этого
     порога инструмент Read в Claude Code перекодирует картинку в JPEG;
   - `judge` с фейковой транскрипцией «эталон» даёт 50 `found`, с «эхом» — 50 `not_found` и 50 `agree`;
   - для оценки лимита, не условие: страниц со случаями ~20, полос на страницу — до 6.

   Другие числа видов или «эталона» — стоп с распечаткой; классификацию и `best_window` не подкручивать. Полоса
   ≥ 500 КБ — стоп.
4. Только для живого прогона: `claude --version` — 2.1.263 или новее (проверено человеком 21.09.2026). На Windows
   бинарь — `claude.cmd` (npm) или `claude.exe` (родной установщик). Для pytest `claude` не нужен.

## Трогать

- `pdf_core.py` — протокол `Verifier` (#88). `VERDICTS` и `VerifyResult` не менять: их использует `GeminiVerifier`.
- `ocr/measure.py` — #88 (полосы, транскрипция, локальный вердикт, отчёт для сверки) и #89 (флаги, выбор бэкенда).
- `ocr/claude_code_verifier.py` — создать (#89).
- `tests/test_ocr_measure.py` — переписать (#88), дополнить (#89).
- `tests/test_claude_code_verifier.py` — создать (#89).
- `CLAUDE.md`:
  - список модулей `ocr/`: +1, «Ten» → «Eleven»;
  - строка `ocr/gemini_verifier.py`: «в замер не подключён с #88»;
  - абзац «Замер vision-сверки» в разделе «OCR CLI»;
  - окружение: `CLAUDE_CODE_LIVE`, `VERIFIER`;
  - «Known production limitation»: Claude Code — только локально;
  - фраза о нумерации: #88, #89;
  - одна строка-урок: «вердикт по фрагменту у модели не спрашивать — только транскрипция вырезки, вердикт локально
    (замер #87)».
- `docs/specs/ocr/README.md` — таблица спек и сводка PLACEHOLDER-ов.
- `docs/PLAN_OCR.md` — «этап B, шаг 2».
- `docs/docx_converter_docs_sync.md` — `### ПРАВКА #88`, `### ПРАВКА #89`, раздел замера, снимок «следующий — #90».

## Не трогать

- `ocr/gemini_verifier.py` — целиком, по решению человека; общие имена из него только импортируются.
- `tests/test_gemini_verifier.py`.
- `ocr/postprocess.py` — из него только импорт `HOMOGLYPHS_CYR_TO_LAT`.
- `ocr/ingest.py`, `ocr/validate.py`, `ocr/cli.py`, `ocr/board.py`, `ocr/diff.py`, `ocr/mineru_provider.py`,
  `ocr/cache.py`, `ocr/__init__.py`.
- `app.py`, `file_converter.py`, `convert.py`, `requirements.txt`, `template*.docx`, `conftest.py`.
- Фикстуры, эталоны, `*.errors.txt`.
- `.cache/ocr/verify/` — не чистить; старые `_test/verify_measure*.json`.

Общие для двух бэкендов имена остаются в `ocr/gemini_verifier.py`: `VerifierError` и подклассы, `verify_cache_key`,
`CACHE_ROOT`, `CONTEXT_TOKENS`, `locate_block`, `crop_block`. Обещанный спекой 13 общий модуль потребовал бы править
замороженный файл. Это отдельная правка — когда файл разрешат трогать.

## Интерфейсы (дословно)

### `pdf_core.py` (#88)

```python
# ПРАВКА #88: модель только переписывает вырезки; вердикт считает ocr.measure.judge по тексту тракта.
# GeminiVerifier.verify (#86) в коде остаётся, но этому протоколу не соответствует: в замере не участвует.
class Verifier(Protocol):
    """Вырезки одной страницы + вопрос -> дословный текст каждой, в том же порядке ("" — текста нет)."""

    def transcribe(self, images: list[bytes], question: str) -> list[str]: ...
```

Старый `verify(image_png, fragment, question) -> VerifyResult` из протокола уходит. `VERDICTS` и `VerifyResult` —
на месте, без изменений.

### `ocr/measure.py` (#88)

```python
"""ПРАВКА #87: измерение vision-сверки на остатке эталонов. В тракт не подключено.
ПРАВКА #88: модель переписывает полосы вырезки, вердикт — локальный diff по text_tokens."""

MEASURE_SCHEMA_VERSION = 2
MEASURE_DIR = FIXTURES.parents[1]                  # _test/: verify_measure.<бэкенд>.<модель>.json, verify_review.….md
CROPS_DIR = MEASURE_DIR / "verify_crops"          # <stem>/pNN-MM.png — полоса MM страницы NN
ERROR_OUTCOMES = ("found", "neighbor", "not_found", "wrong_fix")
CONTROL_OUTCOMES = ("agree", "false_alarm", "unreadable")
SERVICE_OUTCOMES = ("unlocated", "no_image", "error")
ERROR_KINDS = ("merge", "homoglyph", "chars")
REVIEW_OUTCOMES = ("false_alarm", "wrong_fix", "neighbor")   # в отчёт для сверки человеком, в этом порядке
TILE_HEIGHT = 700            # PLACEHOLDER: высота полосы, px при CROP_RESOLUTION = 200 (≈ 8.9 см)
TILE_OVERLAP = 200           # PLACEHOLDER: перекрытие соседних полос, px (≈ 2.5 см, шесть строк 12 pt)
WINDOW_SLACK = 3             # окно транскрипции длиннее или короче фрагмента не больше чем на столько токенов
MIN_RATIO = 0.5              # PLACEHOLDER: SequenceMatcher.ratio лучшего окна ниже — место не найдено
MARKERS = frozenset("-–—•*·")   # токен только из этих знаков — маркер списка, в сравнении не участвует

TRANSCRIBE_QUESTION = (...)  # текст ниже, дословно


@dataclass(frozen=True)
class Case:
    id: str                          # "<stem>-e07" / "<stem>-c07"
    fixture: str
    kind: str                        # "error" | "control"
    tag: str | None
    fragment: str                    # окно текста тракта, токены через пробел
    expected: str                    # ПРАВКА #88: то же окно в эталоне, исправлены все опкоды окна; у control == fragment
    span: tuple[int, int] | None     # ПРАВКА #88: опкод — токены fragment.split()[s:e]; у control — None
    gold_span: tuple[int, int] | None   # ПРАВКА #88: его эталонная сторона — expected.split()[s:e]
    error_kind: str | None           # ПРАВКА #88: одно из ERROR_KINDS; у control — None
    rules: tuple[str, ...]
    page: int | None                 # 1-based
    block_type: str | None
    ambiguous: bool
    tiles: tuple[bytes, ...]         # ПРАВКА #88: полосы вырезки блока (PNG); () -> unlocated / no_image


def split_tiles(png: bytes, *, height: int = TILE_HEIGHT, overlap: int = TILE_OVERLAP) -> tuple[bytes, ...]: ...
def error_kind(a_side: list[str], b_side: list[str]) -> str: ...
def norm(tokens: list[str]) -> list[str]: ...
def best_window(fragment: list[str], transcript: list[str]) -> tuple[int, int, float]: ...
def judge(case: Case, transcript: str) -> dict: ...
def page_tiles(cases: list[Case]) -> dict[tuple[str, int], list[bytes]]: ...
def build_cases(outputs: dict, *, fixtures: Path = FIXTURES) -> list[Case]: ...
def run_measure(cases: list[Case], verifier: Verifier) -> dict: ...
def write_review(measure: dict, path: Path) -> None: ...
def make_verifier() -> Verifier: ...        # #88 — заглушка; #89 — make_verifier(name, model)
def main(argv: "list[str] | None" = None) -> int: ...
```

`score` уходит, его место занимает `judge`.

**`TRANSCRIBE_QUESTION`** — дословно. Изменился текст — изменились ключи кэша, и замер оплачивается заново.

```
Тебе даны картинки — вырезки одной страницы документа на русском языке (скан или печать).
Перепиши текст каждой картинки дословно, как он напечатан: буквы, цифры, знаки, латиница и
кириллица, пробелы между словами. Ничего не исправляй, не дополняй и не перефразируй: опечатку
на картинке переписывай как опечатку.
Каждую строку текста начинай с новой строки. Слово, перенесённое со знаком переноса в конце
строки, пиши целиком без знака переноса; дефис внутри слова (технико-экономический) сохраняй.
Строки, обрезанные верхним или нижним краем картинки, пропускай.
Таблицу переписывай ячейка за ячейкой: строки таблицы сверху вниз, ячейки строки слева направо,
строки текста внутри ячейки — по порядку, каждая с новой строки. Без разметки таблиц.
Ответь одним JSON-объектом без пояснений: ключ — имя файла картинки, значение — её текст;
картинка без читаемого текста — пустая строка:
{"01.png": "<текст>", "02.png": "<текст>"}
```

Порядок «ячейка за ячейкой» — тот же, в каком MinerU отдаёт текст таблицы в `table_body`. Окно фрагмента и
транскрипция идут в одном порядке.

**`split_tiles`**

- `PIL.Image.open(BytesIO(png)).convert("L")`, `w, h = size`.
- `h <= height` → одна полоса целиком.
- Иначе начала полос — `list(range(0, h - height, height - overlap)) + [h - height]`; полоса —
  `crop((0, s, w, s + height))`.
- Каждая полоса — `save(buf, "PNG", optimize=True)`.
- Функция детерминирована. Оттенки серого — ради размера (порог 500 КБ, Шаг 0).

**`error_kind`**. `a, b = "".join(a_side), "".join(b_side)`, дальше по порядку:

- `a == b` → `"merge"`;
- `a` и `b` равны после свёртки латиницы в кириллицу → `"homoglyph"`. Свёртка —
  `str.maketrans({lat: cyr for cyr, lat in HOMOGLYPHS_CYR_TO_LAT.items()})`, таблица из `ocr/postprocess.py`, как в
  `ocr/diff.py`;
- иначе → `"chars"` (вставки и удаления — тоже `chars`).

**`norm`**. Выбросить токены, целиком состоящие из знаков `MARKERS`. Больше ничего не трогать: регистр, `ё`, кавычки,
гомоглифы — как есть. Гомоглифы — пятая часть остатка; пробелы нормализует сам `split`.

**`best_window`**

- `n = len(fragment)`.
- Перебор: `length` в `range(max(1, n - WINDOW_SLACK), n + WINDOW_SLACK + 1)`, `start` в
  `range(0, max(1, len(transcript) - length + 1))`.
- Для каждого окна: `ratio = SequenceMatcher(None, fragment, transcript[start:start + length], autojunk=False).ratio()`.
- Возврат — `(start, start + length, ratio)` первого строгого максимума.
- Пустая транскрипция → `(0, 0, 0.0)`.

Для скорости `seq2` — фрагмент, задаётся один раз. На пробе 100 случаев судятся за 2 с.

**`judge(case, transcript)`** → `{"verdict", "outcome", "reading", "ratio", "ops", "clipped"}`.

1. **Токены и пролёты.**
   - У `error`: `raw = case.fragment.split()`, `(s, e) = case.span`, `before, mid, after = raw[:s], raw[s:e], raw[e:]`.
     Из них `f = norm(before) + norm(mid) + norm(after)`; пролёт опкода —
     `(len(norm(before)), len(norm(before)) + len(norm(mid)))`.
   - Так же `expected` по `gold_span` даёт `E` и пролёт эталона.
   - У `control`: `f = norm(case.fragment.split())`, пролётов нет.
2. **Окно.** `t = norm(text_tokens(transcript))`; `i, j, ratio = best_window(f, t)`; `W = t[i:j]`.
   `ratio < MIN_RATIO` → чтение `unreadable`.
3. **Опкоды.** Берутся не-`equal` опкоды `SequenceMatcher(None, f, W, autojunk=False)`.
   - **Срезанные краем** — первый опкод, если это `delete` с `i1 == 0`, и последний, если это `delete` с
     `i2 == len(f)`. Это токены фрагмента за краем полосы или блока. В вердикт они не идут; их число — `clipped`.
   - Остальные опкоды — живые.
4. **Касание.** `touches(i1, i2, s, e)`: если один из отрезков пуст — `s <= i2 and i1 <= e`, иначе — `i1 < e and s < i2`.

Исход по таблице:

| | чтение `unreadable` | живых нет | живые есть, пролёт опкода не задет | задет; `E`↔`W` в пролёте эталона чисто | задет; `E`↔`W` в пролёте эталона расходятся |
|---|---|---|---|---|---|
| `error` | `not_found` | `not_found` | `neighbor` | `found` | `wrong_fix` |
| `control` | `unreadable` | `agree` | `false_alarm` (любой живой опкод) | — | — |

Что значат столбцы:

- «Задет» — живой опкод `f`↔`W` касается пролёта опкода.
- «`E`↔`W` в пролёте эталона» — живые опкоды `SequenceMatcher(None, E, W, autojunk=False)` касаются пролёта эталона;
  срезанные отбрасываются по тому же правилу.
- Если сам пролёт опкода срезан краем → `not_found`, в строке случая `"reading": "clipped"`.

Поля результата:

- `verdict` — локальный, из `VERDICTS`: `unreadable`, `fix` (живые опкоды есть) или `agree`;
- `reading` — `" ".join(W)`;
- `ops` — живые опкоды `f`↔`W`, каждый списком `[tag, " ".join(f[i1:i2]), " ".join(W[j1:j2]), error_kind(...)]`.

Что это чинит из списка «Цели»:

- соседний опкод в окне больше не портит исход: эталон сравнивается только в пролёте, а `expected` исправлен целиком;
- ответ «одним словом» невозможен, ведущее «- » вычищает `norm`;
- `confidence` не спрашивается;
- склейка на переносе — это `replace` одного токена на два.

**`build_cases`** — как в #87 (опкоды, окна, `rules`, контрольная выборка, `locate_block`), с четырьмя отличиями:

1. **`expected` — эталон на весь пролёт окна**: `b[to_b(lo):to_b(min(hi, len(a)))]`.
   - `to_b(i)`: при `i == len(a)` → `len(b)`; иначе берётся опкод `(tag, i1, i2, j1, j2)` с `i1 <= i < i2`; для
     `equal` → `j1 + (i - i1)`, для остальных → `j1`.
   - В окне исправлены **все** опкоды, не только свой. В #87 было `a[lo:i1] + b[j1:j2] + a[i2:hi]`: соседний опкод
     оставался в «эталоне» ошибкой (docs_sync #87, `bakeoff-e01`). Проба показала цену старой формулы: на фейке
     «эталон» два случая (`bakeoff-e13`, `bakeoff2-e20`) уходили в `neighbor`. С новой формулой — 50 из 50 `found`.
2. `span = (i1 - lo, i2 - lo)`, `gold_span = (j1 - to_b(lo), j2 - to_b(lo))`.
3. `error_kind = error_kind(a[i1:i2], b[j1:j2])`.
4. Вместо PNG — `split_tiles(crop_block(...))`; кэш вырезок по `(page_idx, bbox)` остаётся.

У `control` по-прежнему `expected == fragment`: окно лежит внутри `equal`-участка. Сверка «замер и табло считают
одно и то же» — прежняя.

**`page_tiles`** — `{(fixture, page): [уникальные полосы]}` по случаям с непустыми `tiles`, в порядке первого
появления: случай за случаем, полоса за полосой. Номер полосы в этом списке (с 1) — это `MM` в
`CROPS_DIR/<stem>/pNN-MM.png`.

**`run_measure(cases, verifier)`**

- Для каждой пары `(fixture, page)` из `page_tiles` по порядку — **один** `verifier.transcribe(images,
  TRANSCRIBE_QUESTION)`. Длина ответа ≠ числу картинок → `VerifierError`.
- `STOP_ERRORS` (`VerifierQuotaError`, `VerifierAuthError`, `VerifierConfigError`) → `complete = False`. Текст ошибки
  пишется в `error` всех случаев этой страницы, дальше вызовов нет; у оставшихся привязанных случаев `outcome` —
  `None`.
- Прочий `VerifierError` → `outcome = "error"` у всех случаев страницы, замер продолжается.
- Случай судится по той своей полосе, где `best_window` дал наибольший `ratio` (при равенстве — первая).
- `cache_hit` — `verifier.last_cache_hits[номер полосы]`, если такой атрибут есть.

Возврат:

```json
{"schema_version": 2, "created_at": "…Z", "verifier": "claude-code", "model": "claude-sonnet-5",
 "question_sha256": "…", "complete": true,
 "totals": {"opcodes": 63, "measured": 50, "no_image": 9, "unlocated": 4, "control": 50,
            "found": 0, "neighbor": 0, "not_found": 0, "wrong_fix": 0,
            "agree": 0, "false_alarm": 0, "unreadable": 0, "error": 0},
 "by_fixture": ["…"], "by_rule": ["…"], "by_block_type": ["…"],
 "by_kind": [{"error_kind", "opcodes", "found", "neighbor", "not_found", "wrong_fix"}],
 "calls": [{"images", "input_tokens", "output_tokens", "duration_ms", "cost_usd"}],
 "cases": [{"id", "fixture", "kind", "tag", "error_kind", "page", "block_type", "ambiguous", "rules",
            "fragment", "expected", "span", "crop", "verdict", "reading", "ratio", "ops", "clipped",
            "outcome", "cache_hit", "error"}]}
```

- `by_fixture`, `by_rule`, `by_block_type` — как в #87, плюс колонка `neighbor`. В `by_fixture` — только фикстуры,
  попавшие в `cases`, в порядке `BOARD`.
- `verifier` и `model` — атрибуты `name` и `model` объекта; если их нет — имя класса и `None`.
- `calls` — `list(verifier.calls_log)`, если атрибут есть, иначе `[]`.
- `crop` — `"verify_crops/<stem>/pNN-MM.png"` выбранной полосы.
- Картинки в JSON не пишутся.

**`write_review(measure, path)`** — Markdown для сверки человеком. Закрывает PLACEHOLDER 7 спеки 13 процедурой.
В шапке — бэкенд, модель, дата, `totals` одной строкой и процедура. Дальше по `REVIEW_OUTCOMES` — каждый такой случай:

```
### bakeoff-c01 · false_alarm · стр. 2 · table
![](verify_crops/bakeoff/p02-03.png)
- фрагмент: `…`
- эталон: `…`                     ← только у error
- прочитано: `…`
- расхождения: `было` → `прочитано` (merge); …
- [ ] модель права — ошибка в эталоне   [ ] модель ошиблась
```

Процедура в шапке — дословно:

> Каждый случай сверить с картинкой. Модель права — это пропущенная ошибка тракта: строка в `<stem>.errors.txt`,
> потом `--golden --force`, новые sha и порог — отдельной правкой; у `bakeoff.pdf` эталон `golden.md` не трогается
> никогда — только запись в docs_sync. Модель ошиблась — случай остаётся ложной тревогой. Итог — в docs_sync: сколько
> `false_alarm` подтвердилось как ложные.

**`main`** (#88, без флагов):

1. `build_board` во временной папке → `build_cases`.
2. Для каждой фикстуры с полосами — `rmtree(CROPS_DIR / stem)` и запись `page_tiles` в `CROPS_DIR/<stem>/pNN-MM.png`,
   **до** обращения к модели.
3. `verifier = make_verifier()` → `run_measure`.
4. Печать: четыре таблицы, `totals` и строка про вызовы из `calls` — «вызовов N, картинок M, входных токенов на вызов
   в среднем X, выходных всего Y».
5. Запись `MEASURE_DIR / f"verify_measure.{name}.{model}.json"` и `verify_review.{name}.{model}.md`.

Коды выхода — как в #87. В #88 `make_verifier()` бросает `VerifierConfigError("бэкенд транскрипции не подключён:
Gemini отвечает только на вопрос #87, Claude Code — ПРАВКА #89")`. Команда пишет полосы и выходит с кодом `1` — так
полосы можно смотреть уже после #88.

### `ocr/claude_code_verifier.py` (#89)

```python
"""ПРАВКА #89: Verifier на Claude Code — подпроцесс `claude -p` под авторизацией Claude Code на этой машине.
Только локально: на Streamlit Cloud бинаря claude нет и не будет. Ключей, токенов и HTTP в этом коде нет."""

CLAUDE_MODEL = "claude-sonnet-5"         # полное имя, не алиас: ключ кэша не должен уехать вместе с алиасом
CLAUDE_TIMEOUT_SEC = 600                 # PLACEHOLDER: на один вызов (страница, до 6 полос)
SYSTEM_PROMPT = "You transcribe document images verbatim. Follow the user instructions exactly."
FILES_HEADER = "Картинки (прочитай каждую инструментом чтения файлов; ключ ответа — имя файла):"
STRIPPED_ENV = ("ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN")        # только авторизация подписки Claude Code
AUTH_MARKERS = ("/login", "not logged in", "invalid api key")       # PLACEHOLDER: по живому выводу
LIMIT_MARKERS = ("usage limit", "limit reached", "hit your limit")  # PLACEHOLDER: по живому выводу


class ClaudeCodeMissingError(VerifierConfigError): ...   # нет бинаря claude
class ClaudeCodeAuthError(VerifierAuthError): ...        # Claude Code не залогинен
class ClaudeCodeLimitError(VerifierQuotaError): ...      # упёрлись в лимит подписки
class ClaudeCodeTimeoutError(VerifierError): ...         # вызов не уложился в таймаут


def find_claude() -> str: ...
def claude_command(binary: str, model: str, n_images: int) -> list[str]: ...
def parse_cli_output(stdout: str, stderr: str, model: str) -> dict: ...
def parse_texts(raw: str, names: list[str]) -> dict[str, str]: ...


class ClaudeCodeVerifier:
    name = "claude-code"

    def __init__(self, *, model: str = CLAUDE_MODEL, cache_root: Path | None = CACHE_ROOT,
                 live: bool | None = None, timeout_sec: float = CLAUDE_TIMEOUT_SEC): ...

    def transcribe(self, images: list[bytes], question: str) -> list[str]: ...

    last_cache_hits: list[bool]     # по картинкам последнего transcribe
    network_calls: int              # сколько раз запускался claude
    calls_log: list[dict]           # {"images", "input_tokens", "output_tokens", "duration_ms", "cost_usd"}
```

Классы ошибок наследуют существующие классы `gemini_verifier`, поэтому `STOP_ERRORS` в `measure` их уже знает:

- нет бинаря, не залогинен, лимит подписки — **останавливают** замер (повтор доберёт из кэша);
- таймаут — `error` у страницы, замер идёт дальше.

**`find_claude`**

- Кандидаты: при `sys.platform.startswith("win")` — `("claude.exe", "claude.cmd")`, иначе — `("claude",)`.
- Берётся первый найденный через `shutil.which`.
- Не нашли → `ClaudeCodeMissingError("не найден claude (Claude Code) в PATH: этот бэкенд работает только локально,
  где Claude Code установлен и в нём выполнен вход")`.
- Образец — поиск Ghostscript в `ocr_converter.check_ocr_dependencies`.
- На машине человека проверено: `shutil.which("claude.cmd")` → `…\AppData\Roaming\npm\claude.cmd` (npm), `claude.exe`
  нет.

**`claude_command`** — ровно такой список. Флаги взяты дословно из https://code.claude.com/docs/en/cli-reference.md
(21.09.2026).

```python
[binary, "-p", "--output-format", "json", "--model", model,
 "--tools", "Read", "--permission-mode", "dontAsk", "--safe-mode", "--no-session-persistence",
 "--max-turns", str(n_images + 3), "--system-prompt", SYSTEM_PROMPT]
```

| флаг | зачем | что говорит документация |
|---|---|---|
| `-p` | неинтерактивный режим | «Print response without interactive mode» |
| `--output-format json` | один JSON с полем `result` | «`json`: structured JSON with result, session ID, and metadata» |
| `--model` | модель — параметр | «a model alias such as `sonnet`, `opus` … or a model's full name» |
| `--tools "Read"` | из инструментов — только чтение | «Restrict which built-in tools Claude can use … tool names like `"Bash,Edit,Read"`» |
| `--permission-mode dontAsk` | без интерактивных запросов прав | «Auto-denies every call that would otherwise prompt; file reads in your working directories … still run» |
| `--safe-mode` | не грузить CLAUDE.md, хуки, skills, плагины, MCP и auto memory, но сохранить вход | «CLAUDE.md, skills, plugins, hooks, MCP servers … and auto memory do not load. Authentication, model selection, built-in tools, and permissions work normally, which differs from `--bare`» |
| `--no-session-persistence` | вырезки заказчика не оседают в `~/.claude` | «sessions are not saved to disk and cannot be resumed. Print mode only» |
| `--max-turns` | предохранитель от зацикливания | «Limit the number of agentic turns (print mode only). Exits with an error when the limit is reached» |
| `--system-prompt` | заменить системный промт Claude Code коротким — это главная доля ~46k накладных | «Replace the entire system prompt with custom text» |

Не используются:

- `--bare` — «In bare mode, Claude Code never reads OAuth credentials or the system keychain»: с ним подписка не
  работает;
- `--allowedTools` — чтение в рабочем каталоге в `dontAsk` разрешено и так;
- `--dangerously-skip-permissions` — не нужен;
- `--json-schema` — человек задал разбор поля `result`.

На машине человека SessionStart-хуки плагинов добавляют в каждую сессию тысячи токенов; `--safe-mode` их не запускает.

**Промт идёт через stdin, в argv — только ASCII.** `claude.cmd` запускается через `cmd.exe`: перевод строки в
аргументе обрывает команду, а кавычки и `& | < > ^ %` интерпретируются. Поэтому весь русский текст идёт в `input=`
(`claude -p` читает stdin: «Non-interactive mode reads stdin»). `SYSTEM_PROMPT` и остальные значения argv — ASCII без
этих знаков; тест это проверяет.

**`transcribe(images, question)`** — порядок:

1. **Ключи.** `head = f"{question}\n\n{FILES_HEADER}"`; ключ каждой картинки —
   `verify_cache_key(image, "", head, f"claude-code:{self.model}")`.
   - Формула та же, что у Gemini. Фрагмента в промте нет — отсюда `""`.
   - Провайдер и модель передаются в аргументе `model`. Так понято требование задания «тот же ключ с
     provider=claude-code, model_version=<модель>»: `ocr.cache.LocalCache` хранит zip + meta.json ответов MinerU и под
     «PNG → текст» не подходит.
   - Одинаковые картинки дают один ключ: ошибка и контроль на одной полосе делят ответ.
2. **Кэш.** Попадания берутся из `CACHE_ROOT/<key>.json`. Промахи (уникальные, по порядку) идут дальше. Промахов
   нет — ни бинаря, ни вызова.
3. **Живой режим.** `live` (при `None` → `os.environ.get("CLAUDE_CODE_LIVE") == "1"`) ложен →
   `VerifierConfigError("нет в кэше, вызов выключен: нужен CLAUDE_CODE_LIVE=1")`.
4. **Бинарь.** `binary = find_claude()`.
5. **Рабочий каталог.** `tempfile.TemporaryDirectory(prefix="claude-verify-")` — **пустой каталог вне репозитория**:
   выше него нет `CLAUDE.md` и `.claude/settings.json` проекта. Промахи пишутся туда как `01.png`, `02.png`, …
   Промт — `head + "\n" + "\n".join(абсолютные пути)`; картинку Claude Code читает сам, своим инструментом Read.
6. **Вызов.** `subprocess.run(claude_command(binary, self.model, n), input=prompt, capture_output=True, text=True,
   encoding="utf-8", cwd=work, env=env, timeout=self._timeout_sec)`.
   - `env` — `os.environ` без `STRIPPED_ENV`. По документации ключ из окружения важнее входа по подписке («If you
     have an active Claude subscription but also have `ANTHROPIC_API_KEY` set in your environment, Claude Code uses
     the API key»), а решение человека — только подписка.
   - `TimeoutExpired` → `ClaudeCodeTimeoutError(f"claude -p: нет ответа за {timeout} с ({n} картинок)")`.
   - `network_calls += 1` — в `finally`.
7. **Разбор.** `out = parse_cli_output(proc.stdout, proc.stderr, self.model)`. Ошибка — ничего не кэшируется.
8. **Журнал.** `calls_log.append(...)`:
   - `input_tokens` = `inputTokens + cacheReadInputTokens + cacheCreationInputTokens` модели;
   - `output_tokens` = `outputTokens`;
   - `duration_ms`;
   - `cost_usd` = `total_cost_usd` — расчётная по прайсу, реально платит подписка.
9. **Запись в кэш.** На каждый промах — файл `{"key", "model", "image_sha256", "fragment": "", "question": head, "raw":
   out["result"], "created_at", "file": "03.png", "batch_files": [...], "model_usage": out["modelUsage"][model]}`.
   - Это поля Gemini плюс три новых.
   - Хранится **весь** ответ пакета; разбор делается при каждом чтении (`parse_texts(raw, batch_files)[file]`).
     Поэтому правка разбора не тратит лимит.
   - Неразобранный успешный ответ тоже кэшируется — как в #86.
10. **Результат.** Текст каждой картинки берётся тем же разбором записи — из попадания или из только что записанной.
    `last_cache_hits` — по картинкам. `cache_root=None` выключает кэш: записи живут только в памяти вызова.

Повтор после сбоя шлёт только недостающие полосы страницы. Состава пакета в ключе нет — сознательно, иначе повтор
переспрашивал бы уже отвеченное (PLACEHOLDER 9).

**`parse_cli_output`**

- `json.loads(stdout)`; не JSON → `VerifierError(f"claude -p: ответ не JSON: stdout {stdout[:200]!r}, stderr
  {stderr[:200]!r}")`.
- Поля — по `SDKResultMessage` (https://code.claude.com/docs/en/agent-sdk/typescript.md) и живому прогону человека:
  `is_error`, `subtype`, `result`, `modelUsage`, `total_cost_usd`; у ошибок — `errors`, `api_error_status`.
- Ошибка — если `out["is_error"] or out["subtype"] != "success"`. Тогда `text = out.get("result") or
  "; ".join(out.get("errors") or [])` — поля ветки ошибок по документации необязательны, здесь `.get` законен. Дальше
  по порядку:
  1. `api_error_status` 401/403 или маркер из `AUTH_MARKERS` в `text.lower()` → `ClaudeCodeAuthError("Claude Code не
     авторизован: … — запустите claude и выполните /login")`;
  2. 429 или маркер из `LIMIT_MARKERS` → `ClaudeCodeLimitError("лимит подписки Claude Code: …")`;
  3. иначе → `VerifierError(f"claude -p: {subtype}: {text[:200]}")` (`error_max_turns` и прочие).
- Успех, но запрошенной модели нет среди ключей `modelUsage` → `VerifierConfigError(f"ответила не та модель:
  {sorted(modelUsage)}, ждали {model}")`. Это стоп, а не тихая подмена алиасом или fallback-моделью. Лишние ключи
  (вспомогательные модели) допустимы.
- Возврат — `out`.

**`parse_texts`**

- Срез от первой `{` до последней `}`, затем `json.loads`; обёртку ```` ```json ```` снимает этот же срез.
- Не `dict` → ошибка.
- Ключом считается последнее звено пути (`k.replace("\\", "/").rsplit("/", 1)[-1]`): модель может вернуть полный путь.
- Набор ключей ≠ `names` или значение не строка → `VerifierError(f"ответ не по файлам: ждали {names}, пришло {…}:
  {raw[:200]!r}")`.

### `ocr/measure.py` (#89)

```python
from ocr.claude_code_verifier import CLAUDE_MODEL, ClaudeCodeVerifier

VERIFIERS = ("claude-code", "gemini")


def make_verifier(name: str, model: str | None) -> Verifier: ...
```

- `make_verifier("gemini", …)` → `VerifierConfigError("Gemini отвечает только на вопрос #87 (вердикт по фрагменту),
  транскрипции у него нет — в замер не подключён (решение человека 21.09.2026)")`.
- `make_verifier("claude-code", …)` → `ClaudeCodeVerifier(model=model or CLAUDE_MODEL)`.

`main(argv)` — `argparse`:

```
python -X utf8 -m ocr.measure [--verifier claude-code|gemini] [--model ИМЯ] [--fixture ФАЙЛ]...
```

- `--verifier` — по умолчанию `os.environ.get("VERIFIER", "claude-code")`; значение вне `VERIFIERS` → код `1`.
- `--model` — по умолчанию модель бэкенда.
- `--fixture` — повторяемый, имена из `BOARD`. Остаются случаи только этих фикстур; имя выходных файлов получает
  суффикс `.<stem>+<stem>` (`verify_measure.claude-code.claude-opus-5.bakeoff.json`).
- Остальное — как в #88.

## Приёмочные тесты

В обоих файлах сеть и реальный `claude` недоступны:

- `socket.socket` → `AssertionError`;
- `subprocess.run` — только фейковый (в autouse-фикстуре стоит заглушка с `AssertionError`);
- `CLAUDE_CODE_LIVE`, `VERIFIER`, `GEMINI_LIVE` снимаются через `delenv`.

Живого pytest-теста нет: живой прогон — отдельный шаг ниже.

### `tests/test_ocr_measure.py` (#88, переписать)

```python
# состав — прежний, плюс виды, пролёты и полосы
assert len(errors) == sum(r["count_diffs"] for r in board["rows"])              # 63 на 21.09.2026
assert len(located) == 50 and len(unlocated) == 4 and sum(c.ambiguous for c in located) == 2
assert Counter(c.error_kind for c in located) == {"merge": 31, "homoglyph": 11, "chars": 8}
assert all((c.span[0] == c.span[1]) == (c.tag == "insert") for c in located)
assert all((c.gold_span[0] == c.gold_span[1]) == (c.tag == "delete") for c in located)
assert all(c.expected == c.fragment for c in controls) and all(c.expected != c.fragment for c in errors)
assert all(t.startswith(b"\x89PNG") and len(t) < 500_000 for c in located for t in c.tiles)
assert all(PIL.Image.open(BytesIO(t)).size[1] <= TILE_HEIGHT for c in located for t in c.tiles)
assert build_cases(outputs) == cases                                            # детерминизм

# split_tiles: 1500 px -> начала 0, 500, 800; 600 px -> одна полоса; режим L
# judge — синтетические случаи на настоящих строках остатка:
#   склейка при соседней склейке в окне (IP65не … IP54не): found, len(ops) == 2
#   маркер: транскрипция "- оснащение уличными PoE IP-камерами …" -> found (маркер выброшен norm)
#   вставка (insert, пустой пролёт): found; гомоглиф A.C. -> А.С.: found, error_kind "homoglyph"
#   правка не туда: neighbor; правка туда, но не как в эталоне: wrong_fix; согласие: not_found
#   мусор: ratio < MIN_RATIO -> not_found / unreadable; хвост фрагмента за краем: clipped == 1, исход по остальному
#   control: agree / false_alarm / unreadable

# полный прогон с фейками; transcribe записывает вызовы
echo = Echo(cases)          # полоса -> " ".join(fragment всех случаев на ней)
m = run_measure(cases, echo)
assert len(echo.calls) == len(page_tiles(cases))                                     # страница — один вызов
assert [len(x) for x in echo.calls] == [len(v) for v in page_tiles(cases).values()]  # полосы без повторов
assert m["totals"]["not_found"] == 50 and m["totals"]["agree"] == 50 and m["totals"]["false_alarm"] == 0
golden = Golden(cases)      # полоса -> " ¶ ".join(expected всех случаев на ней)
m = run_measure(cases, golden)
assert m["totals"]["found"] == 50 and m["totals"]["false_alarm"] == 0 and m["totals"]["neighbor"] == 0
# «лимит на 3-м вызове»: complete False; у случаев двух первых страниц исход есть, у третьей — error, дальше None
# VerifierError на одной странице: у её случаев outcome "error", замер complete
# main: MEASURE_DIR и make_verifier подменены (golden): код 0, полосы pNN-MM.png, JSON с верным именем,
#   verify_review…md есть; фейк с одной ложной тревогой -> в отчёте её id, ![](verify_crops/…) и фрагмент
# main без подмены (#88) -> код 1, «ПРАВКА #89» в stderr, полосы уже записаны
```

`Golden` на реальных данных обязан дать `found == 50` (проба 21.09.2026). Если меньше — распечатать случаи и
остановиться. Это значит, что `best_window` промахивается мимо места и на живом прогоне будет так же; алгоритм под
тест не подкручивать.

### `tests/test_claude_code_verifier.py` (#89)

```python
OK = {"type": "result", "subtype": "success", "is_error": False, "result": '{"01.png": "текст"}',
      "modelUsage": {"claude-sonnet-5": {"inputTokens": 10, "outputTokens": 5, "cacheReadInputTokens": 0,
                                        "cacheCreationInputTokens": 0, "provider": "firstParty"}},
      "total_cost_usd": 0.01, "duration_ms": 1200}    # PLACEHOLDER 3: заменить образцом живого вывода человека

# parse_cli_output
assert parse_cli_output(json.dumps(OK), "", "claude-sonnet-5")["result"] == '{"01.png": "текст"}'
# is_error + api_error_status 401 -> ClaudeCodeAuthError; result "Not logged in · Please run /login" -> ClaudeCodeAuthError
# api_error_status 429 -> ClaudeCodeLimitError; "Claude AI usage limit reached" -> ClaudeCodeLimitError
# subtype "error_max_turns" -> VerifierError, но не из STOP_ERRORS; не JSON -> VerifierError со stderr в тексте
# модели нет в modelUsage -> VerifierConfigError

# parse_texts: ```json-обёртка; ключи полными путями (C:\\…\\01.png и /tmp/…/01.png);
#   лишний или недостающий файл, нестроковое значение -> VerifierError

# find_claude: win -> claude.exe раньше claude.cmd; linux -> claude; ничего -> ClaudeCodeMissingError

# transcribe с фейковым subprocess.run (пишет argv, input, cwd, env и список файлов в cwd на момент вызова):
v = ClaudeCodeVerifier(cache_root=tmp_path, live=True)
texts = v.transcribe([PNG_A, PNG_B, PNG_A], QUESTION)                 # две уникальные -> один вызов, два файла
assert len(fake.calls) == 1 and fake.calls[0].files == ["01.png", "02.png"] and len(texts) == 3 and texts[0] == texts[2]
assert fake.calls[0].argv[1:] == claude_command("x", CLAUDE_MODEL, 2)[1:]
assert all(a.isascii() and not set(a) & set('\n"&|<>^%') for a in fake.calls[0].argv[1:])   # claude.cmd идёт через cmd.exe
assert "ANTHROPIC_API_KEY" not in fake.calls[0].env                   # monkeypatch.setenv его выставил
assert not Path(fake.calls[0].cwd).resolve().is_relative_to(REPO_ROOT) # каталог вне репозитория
assert not Path(fake.calls[0].cwd).exists()                           # и убран после вызова
assert all(str(Path(fake.calls[0].cwd, n)) in fake.calls[0].input for n in ("01.png", "02.png"))
assert sorted(json.loads(next(tmp_path.iterdir()).read_text("utf-8"))) == sorted(
    ["key", "model", "image_sha256", "fragment", "question", "raw", "created_at", "file", "batch_files", "model_usage"])
assert v.transcribe([PNG_B], QUESTION) == [texts[1]] and len(fake.calls) == 1 and v.last_cache_hits == [True]
v.transcribe([PNG_A, PNG_C], QUESTION)
assert fake.calls[1].files == ["01.png"]                              # уходит только недостающая
# промах без CLAUDE_CODE_LIVE -> VerifierConfigError; find_claude и run не вызывались
# TimeoutExpired -> ClaudeCodeTimeoutError, кэш пуст; ошибка Auth/Limit -> кэш пуст
# неразобранный успешный ответ -> VerifierError, но файл кэша записан (как в #86)
# ключ зависит от модели: sonnet и opus -> разные файлы
# calls_log: input_tokens = inputTokens + cacheRead + cacheCreation
```

Плюс в `tests/test_ocr_measure.py` (#89):

- `make_verifier("gemini", None)` → `VerifierConfigError` с «#87»;
- `main(["--verifier", "gemini"])` → `1`; то же при `VERIFIER=gemini` в окружении;
- `main(["--fixture", "bakeoff.pdf"])` с подменой → в JSON только `bakeoff.pdf`, имя файла с суффиксом `.bakeoff`;
- настоящий `ClaudeCodeVerifier` с холодным кэшем без `CLAUDE_CODE_LIVE` → код `1`, «CLAUDE_CODE_LIVE» в stderr,
  `subprocess.run` не вызывался.

## Живой прогон (после зелёного pytest и просмотра полос человеком)

Запускает человек — или исполнитель, но только по явной команде человека в чате: вызовы тратят подписку человека.

1. Человек смотрит `_test/verify_crops/` — полосы пишутся без модели. Полоса мимо места — стоп раньше любого вызова.
2. Первый срез — `bakeoff.pdf` на Sonnet (4 страницы, 4 вызова):
   `CLAUDE_CODE_LIVE=1 python -X utf8 -m ocr.measure --fixture bakeoff.pdf`. По первому вызову сверяются
   PLACEHOLDER-ы 3–5 и 8: stdin, `--system-prompt` при входе по подписке, Read в `dontAsk`, накладные токены в
   `calls`, ключ модели в `modelUsage`. Непредусмотренная ошибка — стоп, показать вывод человеку.
3. Весь замер на Sonnet: `CLAUDE_CODE_LIVE=1 python -X utf8 -m ocr.measure`. `bakeoff.pdf` при этом берётся из кэша.
4. `bakeoff.pdf` на Opus: `CLAUDE_CODE_LIVE=1 python -X utf8 -m ocr.measure --model claude-opus-5 --fixture bakeoff.pdf`.

В PowerShell: `$env:CLAUDE_CODE_LIVE = "1"; python -X utf8 -m ocr.measure …`. Лимит подписки (`ClaudeCodeLimitError`)
→ `complete: false`, код `1`. Повтор после сброса лимита доберёт с места остановки.

## Готово, когда

- `pytest -v` зелёный целиком; сеть и `claude` не тронуты; в `git diff --stat` нет файлов из «Не трогать».
- После #88 `python -X utf8 -m ocr.measure` → код `1` с «ПРАВКА #89», полосы записаны. После #89 без
  `CLAUDE_CODE_LIVE` → код `1` с «CLAUDE_CODE_LIVE» (кэш холодный).
- В docs_sync:
  - `### ПРАВКА #88` и `### ПРАВКА #89` — с отклонениями от буквы спеки, если они были;
  - раздел «Замер vision-сверки v2 (транскрипция)». Главная таблица: бэкенд и модель → вызовов → входных токенов на
    вызов (в среднем) → выходных всего → `found` / `neighbor` / `not_found` / `wrong_fix` → `agree` / `false_alarm` /
    `unreadable` → `error`;
  - таблицы по `by_kind` и `by_block_type`;
  - список `false_alarm` со ссылкой на `verify_review…md` и отметкой, сколько подтвердил человек;
  - сравнение Sonnet и Opus на `bakeoff.pdf`;
  - строка о том, какие PLACEHOLDER-ы 3–5 и 8 закрыты первым живым вызовом.

  Ожидаемых чисел нет.
- Снимок нумерации: «следующий — #90».
- **После этого стоп**: интеграция не начинается, пока человек не посмотрел числа и отчёт сверки.

## PLACEHOLDER-ы

1. **Размер полос.** `TILE_HEIGHT = 700`, `TILE_OVERLAP = 200` — стартовые значения. Строка или ячейка таблицы
   недоступна (см. «Цель»), PLACEHOLDER 5 спеки 13 остаётся открытым. Цена полос — таблицы переписываются целиком,
   плюс ~30 % на перекрытие. Смена параметров меняет полосы, а с ними ключи — замер оплачивается заново.
2. **Окно.** `MIN_RATIO = 0.5`, `WINDOW_SLACK = 3` — стартовые; `ratio` и `reading` каждого случая пишутся в JSON.
   - Известный край: опечатку модели в первом или последнем токене окна не отличить от края полосы, она считается
     срезанной (`clipped`). Поэтому ложные тревоги на краях окна недосчитываются.
   - Фрагмент с тремя склейками на семь токенов даёт `ratio` около 0.47 и уйдёт в `unreadable`. На пробе таких нет.
3. **Форма вывода `claude -p --output-format json`.** Поля — по документации и живому прогону человека: `is_error`,
   `subtype`, `result`, `modelUsage` (ключ — имя модели, `provider: "firstParty"`), `total_cost_usd`. Образец от
   человека заменит синтетический `OK` в тесте.
4. **Маркеры ошибок.** `AUTH_MARKERS` и `LIMIT_MARKERS` угаданы; основной признак — `api_error_status` из документации.
   Строки сверить по живому выводу `claude`, когда он не залогинен и когда упёрся в лимит.
5. **stdin и системный промт.** Две вещи сверяются первым живым вызовом: промт через stdin при `-p` без позиционного
   промта и `--system-prompt` при входе по подписке.
   - Если `--system-prompt` не работает — заменить на `--append-system-prompt` (одна строка). Ключи не меняются:
     системный промт в ключ не входит, как `generationConfig` у Gemini.
   - Первый вызов покажет, насколько упали накладные ~46k (`calls[].input_tokens`).
6. **Таймаут и ходы.** `CLAUDE_TIMEOUT_SEC = 600`, `--max-turns n + 3`; `--effort` не задаётся (остаётся по
   умолчанию). Всё это стартовые значения.
   - На Windows `claude.cmd` идёт через `cmd.exe`, и по таймауту `subprocess.run` убивает `cmd.exe`, а дочерний `node`
     может остаться. В коде — комментарий `# ponytail:`; если это случится, лечится через `Popen` + `taskkill /T`.
7. **Размер картинки.** Инструмент Read «resizes and recompresses large images … an image that is still larger than
   500KB after that resize is re-encoded as a JPEG». Полосы в режиме `L` весят до 249 КБ. Как Read масштабирует полосу
   1467×700, не документировано — сверяется глазами по качеству ответов.
8. **Модель — полное имя** (`claude-sonnet-5`, `claude-opus-5`), не алиас: алиас переедет на новую модель, а ключ кэша
   останется старым. Модели нет в `modelUsage` → стоп.
9. **Пакет страницы.** Соседние полосы в одном вызове могут влиять на чтение друг друга. Состав пакета в ключ не
   входит: иначе повтор переспрашивал бы уже отвеченное. На холодном кэше пакет детерминирован — все полосы страницы
   в порядке `page_tiles`.
10. **Гомоглифы** (11 из 50) на картинке неотличимы: `found` на них — языковая догадка модели. `by_kind` держит их
    отдельно; выводы о том, видит ли модель ошибку, — только по `merge` и `chars`.
11. **Что закрыто этой спекой** из спеки 13:
    - PLACEHOLDER 4 — `confidence` не спрашивается;
    - PLACEHOLDER 6 — сравнение мягкое (`judge`);
    - PLACEHOLDER 7 — процедура `verify_review`;
    - PLACEHOLDER 9 — решение человека.

    Открытым остаётся PLACEHOLDER 5 (строки таблиц, см. п. 1). PLACEHOLDER-ы 1–3 спеки 13 — про Gemini, остаются
    как есть.

## Коммит

Два коммита кода, в этом порядке. Таблица замера — третьим, документационным коммитом после живого прогона.
Push делает человек.

```
ПРАВКА #88: замер vision-сверки — транскрипция полос, вердикт локально

pdf_core.py: протокол Verifier — transcribe(вырезки страницы, вопрос) -> текст
каждой; вердикт модель больше не выносит. ocr/measure.py: блок режется на
горизонтальные полосы с перекрытием (строк таблиц в content_list нет), модель
переписывает их дословно, judge сравнивает транскрипцию с фрагментом по
text_tokens в лучшем окне: found / neighbor / not_found / wrong_fix, на контроле —
false_alarm. expected — эталон на весь пролёт окна; соседний опкод исход не
портит, маркеры списка выброшены. Виды опкодов merge / homoglyph / chars. Одна
страница — один вызов. verify_review…md — каждая ложная тревога с полосой и
фрагментом для сверки человеком. GeminiVerifier в замер не подключён; тракт не
тронут.
```

```
ПРАВКА #89: ClaudeCodeVerifier — сверка через claude -p, выбор бэкенда в measure

ocr/claude_code_verifier.py: подпроцесс claude -p (--output-format json,
--tools Read, --permission-mode dontAsk, --safe-mode, --no-session-persistence,
короткий --system-prompt) под входом Claude Code по подписке; ключей, токенов и
HTTP в коде нет, ANTHROPIC_API_KEY из окружения подпроцесса убран. Полосы
страницы — во временный пустой каталог, пути — в промте через stdin. Кэш
.cache/ocr/verify/ по полосе, модель в ключе как claude-code:<модель>; вызов
только при CLAUDE_CODE_LIVE=1. Ошибки: нет бинаря, не залогинен, лимит
подписки, таймаут. python -m ocr.measure --verifier/--model/--fixture,
VERIFIER=claude-code по умолчанию. Только локально: на Streamlit Cloud бэкенда нет.
```
