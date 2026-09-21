# 13 — Vision-сверка: протокол, Gemini и измерение

**# ПРАВКА #86** (`Verifier` + `ocr/gemini_verifier.py`), **# ПРАВКА #87** (`ocr/measure.py`).
Зависит от: 10 (эталоны и пороги), 11. От 12 не зависит. Сеть — **только** у человека, при
`GEMINI_LIVE=1`; исполнитель спеки в облако не ходит вообще.

## Цель

Этап B, первый шаг — **ИЗМЕРЕНИЕ, не интеграция**. На шести эталонах известен полный остаток
тракта: 63 опкода (13 + 20 + 8 + 13 + 8 + 1, табло 21.09.2026). Вопрос этой спеки один: если
показать модели вырезку страницы и фрагмент текста — сколько из этих ошибок она найдёт,
сколько исправит неверно и сколько раз поднимет ложную тревогу на заведомо верном тексте.

**В тракт ничего не подключается.** `ingest`, `postprocess`, `validate`, `report.json`, CLI и UI
этой спекой не меняются и о `Verifier` не знают. Интеграция — следующая спека, после чисел:
без них неизвестно, окупает ли сверка свою квоту и не портит ли она текст чаще, чем чинит.

Решения человека (21.09.2026):

- **один опкод — один запрос.** Совпадает с сигнатурой `Verifier` и с ключом кэша; ответ по
  фрагменту не зависит от соседей, ложные тревоги считаются честно. Не больше 54 + 54 = 108 запросов
  один раз (по пробе привязки — 50 + 50, см. ниже), дальше — из кэша. Пакет «все фрагменты страницы одним запросом» отвергнут: модель видит
  список, где часть заведомо с ошибками, и вердикты влияют друг на друга; частота ложных тревог
  в таком замере не переносится на будущий тракт;
- **DOCX/XLSX из замера исключены**: картинки страницы у них нет. Их 9 опкодов (8 + 1) идут в
  таблицу исходом `no_image` и в знаменатели не входят. Меряются опкоды четырёх PDF — их **54**, из них
  к блоку страницы привязывается 50 (остальные 4 — исход `unlocated`, тоже вне знаменателя, поимённо ниже).

## Шаг 0 — условия входа (иначе стоп)

1. Спека 10 закоммичена; `python -X utf8 -m ocr.board` оффлайн даёт шесть строк с `diffs == thr`.
2. `grep -rn "ПРАВКА #8[6-9]" --include=*.py .` (без `.venv`) пуст.
3. В `.venv` работает растеризация без новых зависимостей (проверено 21.09.2026: pdfplumber 0.11.9 +
   pypdfium2, вырезка `bakeoff.pdf` стр. 1 → PNG 508×156):
   `pdfplumber.open(pdf).pages[0].crop(box).to_image(resolution=200).original.save(buf, "PNG")`.
   Не работает — **стоп**: в `requirements.txt` ничего не добавлять.
4. Форма `content_list` (проверено на `vlm_raw.zip` и `vlm_raw_textpdf1.zip`): у каждого блока есть
   `type`, `bbox`, `page_idx`; **`bbox` нормирован в 0–1000 по обеим осям, это не пункты PDF**
   (страница 595×841 pt, максимум `bbox` — 966). Исполнитель перепроверяет на всех четырёх zip:
   `max(bbox) <= 1000`. Иначе — стоп, масштаб вырезки неверен.

## Трогать

- `pdf_core.py` — `VERDICTS`, `VerifyResult`, `Verifier` (#86); существующее не менять
- `ocr/gemini_verifier.py` — создать (#86)
- `ocr/measure.py` — создать (#87)
- `tests/test_gemini_verifier.py`, `tests/test_ocr_measure.py` — создать
- `CLAUDE.md` — список модулей (+2), раздел «OCR CLI»: абзац про `python -m ocr.measure`; «Secrets»/окружение: `GEMINI_API_KEY`, `GEMINI_LIVE`
- `docs/specs/ocr/README.md` — таблица спек; `docs/PLAN_OCR.md` — строка «этап B, шаг 1»
- `docs/docx_converter_docs_sync.md` — `### ПРАВКА #86`, `### ПРАВКА #87`, снимок «следующий — #88»

## Не трогать

`ocr/ingest.py`, `ocr/cli.py`, `ocr/postprocess.py`, `ocr/validate.py`, `ocr/board.py`, `ocr/diff.py`,
`ocr/mineru_provider.py`, `ocr/cache.py`, `ocr/__init__.py` (закрытый список правил **не пополняется**),
`app.py`, схема `report.json`, `requirements.txt` (`requests`, `pdfplumber` уже стоят; SDK Google не ставить),
`conftest.py` (маркер `live` — про MinerU; живого pytest-теста Gemini нет), фикстуры, эталоны, `errors.txt`.

## Интерфейсы (дословно)

### `pdf_core.py` (#86)

```python
# ПРАВКА #86: сверка фрагмента текста с изображением страницы (этап B).
VERDICTS = ("agree", "fix", "unreadable")


@dataclass(frozen=True)
class VerifyResult:
    verdict: str                 # одно из VERDICTS
    correction: str | None       # весь фрагмент в исправленном виде; только при verdict == "fix"
    confidence: float            # 0–1, самооценка модели (см. PLACEHOLDER 4)
    raw: str                     # сырой текст ответа модели


class Verifier(Protocol):
    """Картинка вырезки + фрагмент + вопрос -> вердикт. Один фрагмент — один вызов."""

    def verify(self, image_png: bytes, fragment: str, question: str) -> VerifyResult: ...
```

Соответствие словам человека: «согласен» — `agree`, «исправить на …» — `fix` + `correction`,
«не читается» — `unreadable`. `correction` — **весь фрагмент целиком**, а не одно слово: иначе
ответ нельзя сравнить с эталонным окном без догадок о месте замены.

### `ocr/gemini_verifier.py` (#86)

```python
"""ПРАВКА #86: Verifier на Gemini (REST, без SDK) + вырезки страниц по content_list."""

GEMINI_MODEL = "PLACEHOLDER"           # PLACEHOLDER: модель выбирает человек
API_URL = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
MIN_INTERVAL_SEC = 6.0                 # PLACEHOLDER: пауза между сетевыми вызовами (≈10 запросов/мин), не сверено
REQUEST_TIMEOUT_SEC = 120.0
RETRY_PAUSES_SEC = (1.0, 2.0, 4.0)     # как в mineru_provider: 5xx и обрывы связи
CACHE_ROOT = Path(__file__).resolve().parents[1] / ".cache" / "ocr" / "verify"
CROP_RESOLUTION = 200
CROP_MARGIN = 20                       # PLACEHOLDER: поле запаса, в единицах bbox (0–1000)
BBOX_SCALE = 1000
CONTEXT_TOKENS = 3

QUESTION = (...)                       # текст ниже, дословно


class VerifierError(RuntimeError): ...
class VerifierAuthError(VerifierError): ...     # нет ключа, HTTP 401 / 403
class VerifierQuotaError(VerifierError): ...    # HTTP 429


def verify_cache_key(image_png: bytes, fragment: str, question: str, model: str) -> str: ...


def parse_answer(raw: str) -> VerifyResult: ...


class GeminiVerifier:
    def __init__(self, api_key: str | None = None, *, model: str = GEMINI_MODEL,
                 cache_root: Path | None = CACHE_ROOT, live: bool | None = None,
                 min_interval_sec: float = MIN_INTERVAL_SEC,
                 session=None, sleep=time.sleep, clock=time.monotonic): ...

    def verify(self, image_png: bytes, fragment: str, question: str = QUESTION) -> VerifyResult: ...

    last_cache_hit: bool            # атрибут: попал ли последний verify в кэш


def locate_block(needle_tokens: list[str], context_before: list[str], context_after: list[str],
                 content_list: list) -> tuple[dict | None, bool]: ...       # (блок, неоднозначно)


def crop_block(pdf_bytes: bytes, page_idx: int, bbox: list, *,
               margin: int = CROP_MARGIN, resolution: int = CROP_RESOLUTION) -> bytes: ...   # PNG
```

**Ключ кэша.** `sha256(image_png + b"\0" + fragment.encode() + b"\0" + question.encode() + b"\0" + model.encode()).hexdigest()`.
Файл `CACHE_ROOT/<key>.json`, `utf-8`, `ensure_ascii=False`:
`{"key", "model", "image_sha256", "fragment", "question", "raw", "created_at"}`. Хранится **сырой ответ**,
а не вердикт: `VerifyResult` собирается `parse_answer(raw)` при каждом чтении — правка разбора не
требует квоты. `cache_root=None` — кэш выключен (тесты). Запись — только после успешного HTTP-ответа
с текстом; ошибки не кэшируются.

**Сеть и ключ.** Порядок в `verify`: (1) `model == "PLACEHOLDER"` → `VerifierError("модель не выбрана …")`
— раньше кэша, иначе в кэш лягут ответы под мусорным именем; (2) кэш: попадание → результат, ни ключ,
ни сеть не нужны, паузы нет; (3) `live` (`None` → `os.environ.get("GEMINI_LIVE") == "1"`) ложно →
`VerifierError("нет в кэше, сеть выключена: нужен GEMINI_LIVE=1")` — промах кэша обязан упасть, а не
уйти в облако (урок #71); (4) ключ `api_key or os.environ.get("GEMINI_API_KEY")`, нет →
`VerifierAuthError`; ключ в тексты исключений и в кэш не попадает; (5) пауза
`max(0, min_interval_sec - (clock() - время_прошлого_сетевого_вызова))` через `sleep`; (6) запрос.

Прокси — стандартные `HTTPS_PROXY` / `HTTP_PROXY` / `NO_PROXY`: `requests.Session` читает их сам
(`trust_env=True` по умолчанию). В коде про прокси **ничего**: ни параметра, ни константы.

**Запрос** (PLACEHOLDER 3 — форма сверяется на первом живом вызове):

```
POST API_URL.format(model=…)      headers: x-goog-api-key: <ключ>, Content-Type: application/json
{"contents": [{"parts": [{"inline_data": {"mime_type": "image/png", "data": "<base64>"}},
                         {"text": "<question>\n\nТекст OCR:\n<fragment>"}]}],
 "generationConfig": {"temperature": 0}}
```

Ответ: `raw = "".join(part["text"] for part in data["candidates"][0]["content"]["parts"] if "text" in part)`;
нет `candidates` или текста → `VerifierError` с первыми 200 символами тела. HTTP 401/403 → `VerifierAuthError`,
429 → `VerifierQuotaError` (без ретраев: квота паузой в секунды не лечится), 5xx и `ConnectionError`/`Timeout` —
ретраи по `RETRY_PAUSES_SEC`, затем `VerifierError`; прочие 4xx → `VerifierError`. Поля принудительного JSON
(`responseMimeType` / `response_format`) **не используются**: документация Google на 21.09.2026 в переходе
между двумя API, формат ответа задаётся вопросом и разбирается терпимо.

**`QUESTION`** (дословно; меняется — меняется ключ кэша, весь замер платится заново):

```
На картинке — фрагмент страницы документа на русском языке. Ниже — текст, который OCR распознал
в одном месте этой картинки. Найди это место и сравни текст с изображением посимвольно: буквы,
цифры, пробелы между словами, знаки, латиница и кириллица. Ничего не улучшай и не перефразируй:
важно только то, что напечатано на картинке.
Ответь одним JSON-объектом без пояснений:
{"verdict": "agree" | "fix" | "unreadable", "correction": "<весь текст OCR в исправленном виде>" | null,
 "confidence": <число от 0 до 1>}
agree — текст совпадает с картинкой; fix — не совпадает, в correction весь фрагмент, исправленный
по картинке; unreadable — место не найдено или не читается.
```

**`parse_answer`.** Снять обрамление ```` ```json … ``` ````, взять подстроку от первой `{` до последней `}`,
`json.loads`. Ключи обязательны — `d['verdict']`, `d['confidence']` (без `.get`); `verdict` не из `VERDICTS`,
`confidence` не число в [0, 1], `fix` без непустой `correction` → `VerifierError` с `raw[:200]`. При `agree` /
`unreadable` `correction` → `None`. Не разобрали — **ошибка, а не `unreadable`**: «модель не смогла
прочитать» и «мы не смогли разобрать ответ» — разные исходы замера.

**`locate_block`.** Привязка места в тексте тракта к блоку `content_list`. Текст блока — как в
`ocr/validate.py::page_of`: `block.get("text") or block.get("table_body") or ""`, теги сняты. Сравнение —
**без пробельных символов вообще** (половина остатка — потерянные пробелы, с пробелами место не найдётся).
Игла = `context_before[-k:] + needle_tokens + context_after[:k]`, `k` от `CONTEXT_TOKENS` до `0`; на первом `k`,
где нашёлся хоть один блок, — стоп: возвращается первый такой блок и флаг «неоднозначно», если блоков больше
одного. Пустая игла (вставка при `k = 0`) или блок без `bbox`/`page_idx` → `(None, False)`. Не нашли — не
угадывать: в замере это исход `unlocated`. Ловушка среза: `context_before[-0:]` — это **весь** список, при
`k = 0` контекст брать пустым явно.

Проба этого алгоритма на фикстурах (21.09.2026, одноразовый скрипт, в репозиторий не вошёл): из 54 опкодов
привязано **50**, неоднозначных 2 (оба `bakeoff2.pdf`), по типу блока — `table` 44, `text` 5, `header` 1.
Не привязаны 4: `bakeoff.pdf` — опкод подписей `A.II. Taipov P.P. Hypeeb A.P.` (пять токенов лежат в разных
блоках `content_list`); `textpdf1.pdf` — три вставки потерянного текста (`«ТЕНЗОСИЛА»;`, строка модификаций
`М0601 … М0808`, `Федеральное государственное унитарное`): a-сторона пуста, а контекст в тракте изменён.
Запасной поиск по одностороннему контексту спасает одну из четырёх — не вводится: лишняя ветка ради одного
случая. Исполнитель обязан получить те же числа; другие — стоп с распечаткой, алгоритм не подкручивать.

**`crop_block`.** `pdfplumber.open(BytesIO(pdf_bytes))`, `page = pdf.pages[page_idx]`;
`x = (bbox_x ∓ margin) * page.width / BBOX_SCALE`, `y` — так же от `page.height`; клэмп в `page.bbox`;
`page.crop(box).to_image(resolution=resolution).original` → PNG в `bytes`. Для текстовых PDF — то же самое:
растеризуется страница, текстовый слой не используется. `page_idx` вне документа, вырожденный прямоугольник →
`ValueError`. Вырезка таблицы — вся таблица страницы: ячеечных `bbox` нет ни в `content_list.json`, ни в
`layout.json` (проверено 21.09.2026) — PLACEHOLDER 5.

Обе функции лежат здесь, а не в `measure.py`: они понадобятся интеграции. Отдельный модуль — когда
появится второй `Verifier`.

### `ocr/measure.py` (#87)

```python
"""ПРАВКА #87: измерение vision-сверки на остатке эталонов. В тракт не подключено."""

MEASURE_SCHEMA_VERSION = 1
MEASURE_JSON = FIXTURES.parents[1] / "verify_measure.json"      # _test/verify_measure.json
CROPS_DIR = FIXTURES.parents[1] / "verify_crops"                # _test/verify_crops/<stem>/<id>.png — для глаз человека
ERROR_OUTCOMES = ("found", "not_found", "wrong_fix")
CONTROL_OUTCOMES = ("agree", "false_alarm", "unreadable")
SERVICE_OUTCOMES = ("unlocated", "no_image", "error")


@dataclass(frozen=True)
class Case:
    id: str                      # "<stem>-e07" / "<stem>-c07"
    fixture: str
    kind: str                    # "error" | "control"
    tag: str | None              # тег опкода difflib; у control — None
    fragment: str                # окно текста тракта, токены через пробел
    expected: str                # то же окно в эталоне; у control == fragment
    rules: tuple[str, ...]       # правила валидатора, чьи находки покрывают опкод; () — ни одно
    page: int | None             # 1-based
    block_type: str | None       # "text" | "table" | …
    ambiguous: bool
    image_png: bytes | None      # None -> исход unlocated / no_image без вызова модели


def build_cases(outputs: dict, *, fixtures: Path = FIXTURES) -> list[Case]: ...


def score(case: Case, result: "VerifyResult | None") -> str: ...


def run_measure(cases: list[Case], verifier: Verifier) -> dict: ...


def main(argv: "list[str] | None" = None) -> int: ...
```

**`build_cases`** — детерминирован, сети и модели не касается. `outputs` — второй элемент `build_board`
(оффлайн, спека 09). Для каждой пары `BOARD`:

- опкоды — тем же `difflib.SequenceMatcher(None, text_tokens(out), text_tokens(golden), autojunk=False)`,
  что в `count_diffs`; берутся все не-`equal`. Сверка: число опкодов фикстуры == `count_diffs(out, golden)`,
  иначе `AssertionError` — замер и табло обязаны считать одно и то же;
- окно: `a[i1 - CONTEXT_TOKENS : i2 + CONTEXT_TOKENS]` → `fragment`; `a[…:i1] + b[j1:j2] + a[i2:…]` → `expected`
  (контекст с обеих сторон — из `equal`-участков, он общий);
- `zip_name is None` (docx1, xlsx1) → `image_png = None`, исход потом `no_image`; иначе `content_list` — из
  `fixtures / zip_name` (единственный член архива с суффиксом `content_list.json`; `content_list_v2.json` не он),
  `locate_block` → `crop_block`; блок не найден → `image_png = None`, исход `unlocated`;
- `rules`: находка `report["findings"]` покрывает опкод, если её `snippet` и a-сторона опкода (оба без
  пробельных символов, a-сторона непуста) входят друг в друга в любую сторону. Это отвечает на вопрос
  будущей интеграции: дойдёт ли сверка «только по помеченным» до этих ошибок вообще;
- **контрольная выборка**: на каждый привязанный `error`-случай фикстуры — один `control` той же длины окна.
  Кандидаты — окна внутри `equal`-участков, отстоящие от любого не-`equal` опкода не меньше чем на
  `CONTEXT_TOKENS` токенов; берутся **равномерным шагом** по списку кандидатов (`random` не используется —
  замер воспроизводим без зерна); кандидат, не привязавшийся к блоку, пропускается, берётся следующий.
  Кандидатов не хватило — `control` меньше, число записывается в `totals`, не дотягивается.

**`score`**: `image_png is None` → `no_image` (фикстура без zip) или `unlocated`; `result is None`
(`VerifierError`) → `error`. Иначе сравнение — `text_tokens(correction) == text_tokens(expected)`:

| kind | `agree` | `fix`, совпало с `expected` | `fix`, не совпало | `unreadable` |
|---|---|---|---|---|
| `error` | `not_found` | `found` | `wrong_fix` | `not_found` |
| `control` | `agree` | — (невозможно: `expected == fragment`) → `false_alarm` | `false_alarm` | `unreadable` |

«Ложно исправлено» из задания = `wrong_fix` (ошибка была, исправление неверное) **плюс** `false_alarm`
(ошибки не было); в таблицах они врозь — у них разная цена для тракта.

**`run_measure`**: по случаям по порядку, `verifier.verify(case.image_png, case.fragment, QUESTION)`;
`VerifierError` → исход `error`, текст в случай, замер продолжается; `VerifierQuotaError` и
`VerifierAuthError` → замер **останавливается**, `complete = False` (повтор после сброса квоты
доберёт остальное: отвеченное уже в кэше). Возврат:

```json
{"schema_version": 1, "created_at": "…Z", "model": "…", "question_sha256": "…", "complete": true,
 "totals": {"opcodes": 63, "measured": 50, "no_image": 9, "unlocated": 4, "control": 50,
            "found": 0, "not_found": 0, "wrong_fix": 0, "agree": 0, "false_alarm": 0, "unreadable": 0, "error": 0},
 "by_fixture": [{"fixture", "opcodes", "found", "not_found", "wrong_fix", "unlocated", "no_image", "error",
                 "control", "false_alarm", "unreadable"}],
 "by_rule": [{"rule", "opcodes", "found", "not_found", "wrong_fix"}],
 "by_block_type": [{"block_type", "opcodes", "found", "not_found", "wrong_fix", "control", "false_alarm"}],
 "cases": [{"id", "fixture", "kind", "tag", "page", "block_type", "ambiguous", "rules", "fragment", "expected",
            "verdict", "correction", "confidence", "outcome", "cache_hit", "error"}]}
```

`by_rule`: опкод с несколькими правилами идёт в строку каждого; без правил — в строку `"—"`
(сумма строк поэтому может превышать `measured` — так и написать под таблицей). Картинка в JSON не пишется.

**`main`** (`python -X utf8 -m ocr.measure`): `build_board` во временной папке → `build_cases` →
`GeminiVerifier()` → `run_measure` → печать трёх таблиц в консоль, запись `MEASURE_JSON` и вырезок в
`CROPS_DIR` (перезаписываются). Коды: `0` — замер полный; `1` — исключение или `complete == False`
(JSON с тем, что успели, всё равно записан). Без `GEMINI_LIVE=1` команда работает только по кэшу:
холодный кэш → первый же случай даёт остановку с понятным текстом, а не поход в сеть.
Флагов нет: модель — константа, выборка — детерминирована.

## Приёмочные тесты

Сеть заглушена в обоих файлах (`monkeypatch.setattr(socket, "socket", <AssertionError>)`), `GEMINI_LIVE` и
`GEMINI_API_KEY` снимаются `monkeypatch.delenv`. Живого теста нет: живой прогон — команда человека.

### `tests/test_gemini_verifier.py`

```python
# разбор ответа
assert parse_answer('{"verdict":"agree","correction":null,"confidence":0.9}').verdict == "agree"
r = parse_answer('```json\n{"verdict":"fix","correction":"IP-камерами","confidence":0.7}\n```')
assert (r.verdict, r.correction, r.confidence) == ("fix", "IP-камерами", 0.7)
assert parse_answer('{"verdict":"agree","correction":"мусор","confidence":1}').correction is None
for bad in ('не json', '{"verdict":"maybe","confidence":0.5}', '{"verdict":"fix","correction":"","confidence":0.5}',
            '{"verdict":"agree","confidence":1.5}', '{"verdict":"agree"}'):
    with pytest.raises(VerifierError):
        parse_answer(bad)

# ключ кэша: зависит от всех четырёх частей
base = verify_cache_key(b"png", "ф", "в", "m")
assert len({base, verify_cache_key(b"png2", "ф", "в", "m"), verify_cache_key(b"png", "ф2", "в", "m"),
            verify_cache_key(b"png", "ф", "в2", "m"), verify_cache_key(b"png", "ф", "в", "m2")}) == 5

# сеть: фейковая session, журнал вызовов; clock/sleep инжектированы
v = GeminiVerifier("key", model="m", cache_root=tmp_path, live=True, session=fake, sleep=sleeps.append, clock=fake_clock)
first = v.verify(PNG, "фрагмент")
assert fake.calls[0].headers["x-goog-api-key"] == "key" and "key" not in fake.calls[0].url
assert fake.calls[0].json["contents"][0]["parts"][0]["inline_data"]["mime_type"] == "image/png"
assert v.verify(PNG, "фрагмент") == first and len(fake.calls) == 1 and v.last_cache_hit   # повтор — из кэша
v.verify(PNG, "другой"); assert sleeps == [pytest.approx(6.0)]                           # пауза только перед сетью
assert "key" not in (tmp_path / f"{verify_cache_key(PNG, 'фрагмент', QUESTION, 'm')}.json").read_text("utf-8")

# промах кэша без GEMINI_LIVE — ошибка, не сеть; попадание — без ключа
with pytest.raises(VerifierError, match="GEMINI_LIVE"):
    GeminiVerifier(model="m", cache_root=tmp_path, live=False).verify(PNG, "нового нет")
assert GeminiVerifier(model="m", cache_root=tmp_path, live=False).verify(PNG, "фрагмент") == first
with pytest.raises(VerifierAuthError):
    GeminiVerifier(model="m", cache_root=None, live=True, session=fake).verify(PNG, "ф")      # ключа нет
with pytest.raises(VerifierError, match="модель не выбрана"):
    GeminiVerifier("key", cache_root=tmp_path, live=True, session=fake).verify(PNG, "ф")      # GEMINI_MODEL = PLACEHOLDER
# HTTP: 429 -> VerifierQuotaError без ретраев; 403 -> VerifierAuthError; 500 x4 -> VerifierError после трёх пауз;
# ответ без candidates -> VerifierError; ошибки в кэш не пишутся (папка пуста)

# вырезка на настоящем скане
pdf = require_fixture("bakeoff.pdf").read_bytes()
_, content_list = read_raw("vlm_raw.zip")
block = content_list[0]                                     # «Утверждаю:», bbox в 0–1000
png = crop_block(pdf, block["page_idx"], block["bbox"])
assert png.startswith(b"\x89PNG\r\n\x1a\n")
w, h = PIL.Image.open(BytesIO(png)).size
assert abs(w - (block["bbox"][2] - block["bbox"][0] + 2 * CROP_MARGIN) / 1000 * 595 / 72 * 200) < 8   # масштаб 0–1000, не пункты
assert len(crop_block(pdf, 0, [0, 0, 1000, 1000])) > len(png)                                          # клэмп по краям страницы
with pytest.raises(ValueError):
    crop_block(pdf, 99, block["bbox"])

# привязка: пробелы не мешают, контекст сужается, неоднозначность помечается
cl = [{"type": "text", "text": "требованиям IP65не ниже", "bbox": [1, 1, 9, 9], "page_idx": 2},
      {"type": "table", "table_body": "<td>для IP-камерами</td>", "bbox": [1, 1, 9, 9], "page_idx": 3}]
assert locate_block(["IP65не"], ["требованиям"], ["ниже"], cl) == (cl[0], False)
assert locate_block(["IP-камерами"], ["чужой", "контекст"], [], cl)[0] is cl[1]        # k сузился до 0
assert locate_block(["нет"], [], [], cl) == (None, False) and locate_block([], [], [], cl) == (None, False)
assert locate_block(["IP"], [], [], cl) == (cl[0], True)                                # два блока — первый + флаг
```

### `tests/test_ocr_measure.py`

```python
board, outputs = build_board(tmp_path)                     # оффлайн, как в спеке 09
cases = build_cases(outputs)
errors = [c for c in cases if c.kind == "error"]
controls = [c for c in cases if c.kind == "control"]

assert len(errors) == sum(r["count_diffs"] for r in board["rows"])          # 63 на 21.09.2026 — замер и табло едины
assert sum(c.image_png is None and c.fixture in ("docx1.docx", "xlsx1.xlsx") for c in errors) == 9
assert all(c.image_png is None for c in cases if c.fixture in ("docx1.docx", "xlsx1.xlsx"))
located = [c for c in errors if c.image_png is not None]
unlocated = [c for c in errors if c.image_png is None and c.fixture.endswith(".pdf")]
assert len(located) + len(unlocated) == 54 and len(unlocated) == 4           # проба 21.09.2026: 50 + 4, без запаса
assert sum(c.ambiguous for c in located) == 2
assert len(controls) == len(located)
assert all(c.expected == c.fragment for c in controls) and all(c.expected != c.fragment for c in errors)
assert all(len(c.fragment.split()) == len(e.fragment.split()) for c, e in zip(controls_of(f), located_of(f)))
assert build_cases(outputs) == cases                                        # детерминизм, random не участвует
ir = next(c for c in errors if "IR-камерами" in c.fragment)
assert "IP-камерами" in ir.expected and ir.page is not None and ir.image_png.startswith(b"\x89PNG")

# «оракул»: знает эталон -> все привязанные найдены, ложных тревог нет
oracle = FakeVerifier({c.fragment: c.expected for c in cases})             # fix при expected != fragment, иначе agree
m = run_measure(cases, oracle)
assert m["complete"] and m["totals"]["found"] == len(located) and m["totals"]["false_alarm"] == 0
assert m["totals"]["no_image"] == 9 and m["totals"]["opcodes"] == len(errors)
assert sum(r["opcodes"] for r in m["by_fixture"]) == len(errors) and len(m["by_fixture"]) == 6
assert any(r["rule"] == "—" for r in m["by_rule"])                          # «сыручими» правилом не ловится

# «всегда согласен»: ничего не найдено, ложных тревог нет; «всегда правит мусором»: wrong_fix и false_alarm
# «квота на 10-м вызове»: complete False, первые девять исходов на месте, main -> 1, JSON записан
# разбор сломан на одном случае -> outcome "error", замер не остановился
# score: таблица 2x4 из спеки, поштучно
# main с подменёнными MEASURE_JSON / CROPS_DIR и GeminiVerifier -> oracle: код 0, PNG-вырезки лежат, в JSON нет ключа image_png
```

`FakeVerifier` — класс в тесте, три строки; это и есть «фейковый Verifier из фикстур»: его ответы строятся
из эталонов, а не из записанных ответов модели. Записанные ответы появятся в `.cache/ocr/verify/` после
живого прогона; в git они не идут (`.cache/` в `.gitignore`).

## Готово, когда

- `pytest -v` зелёный целиком, сеть не тронута; `git diff --stat` не содержит файлов из «Не трогать».
- `python -X utf8 -m ocr.measure` без `GEMINI_LIVE` завершается кодом `1` с текстом про `GEMINI_MODEL` /
  `GEMINI_LIVE` — это ожидаемо: исполнитель замер **не запускает**.
- `_test/verify_crops/` собран прогоном `main` с оракулом (или отдельной строкой в отчёте — как): человек
  глазами смотрит 5–10 вырезок до траты квоты. Вырезка мимо места — стоп раньше любого живого вызова.
- В docs_sync: `### ПРАВКА #86`, `### ПРАВКА #87`; таблица «фикстура → опкодов → привязано → `unlocated` →
  `no_image` → контрольных»; поимённый список `unlocated` с причиной; доля случаев `block_type == "table"`.
- В отчёте прогона — инструкция человеку: выбрать модель (константа `GEMINI_MODEL`, отдельный коммит
  человека или следующая правка), выставить `GEMINI_API_KEY`, при необходимости `HTTPS_PROXY`,
  `GEMINI_LIVE=1`, запустить; прислать `_test/verify_measure.json`. **После этого стоп**: интеграция не
  начинается, пока человек не посмотрел числа.

## PLACEHOLDER-ы

1. `GEMINI_MODEL` — выбирает человек. На 21.09.2026 в списке моделей Google есть линейки `gemini-3.x-flash`,
   `gemini-3.1-pro-preview`, `gemini-2.5-*` (прочитано агентом, не сверено руками). Смена модели = новый ключ
   кэша = новый платный замер; два замера разными моделями — два файла JSON, переименовывает человек.
2. `MIN_INTERVAL_SEC = 6.0` — не сверено: официальная страница лимитов per-model таблицы больше не публикует,
   фактический лимит виден только в AI Studio под аккаунтом. 100 запросов × 6 с = 10 минут. Дневной лимит
   может оказаться меньше 100 — тогда замер добирается за два дня, кэш это переживает (`complete: false` → повтор).
3. Форма REST-запроса: `generateContent` + `inline_data` + заголовок `x-goog-api-key`. По документации
   (21.09.2026) путь рабочий, но объявлен устаревшим в пользу Interactions API. Сверяется первым живым вызовом;
   не сработало — правится только тело `_post` в `gemini_verifier.py`, ключ кэша и протокол не зависят от формы.
4. `confidence` — **самооценка модели**, не вероятность: logprobs у текущих моделей Gemini ненадёжны. Замер
   заодно покажет, отделяет ли она верные вердикты от неверных (в `cases` есть и `confidence`, и `outcome`);
   до этого порогов по ней не строить.
5. Вырезка таблицы = вся таблица страницы: ячеечных координат MinerU не даёт. Большая часть остатка сидит в
   таблицах, модель ищет три слова на целой странице. `by_block_type` покажет цену; сужение вырезки (по строкам
   через распознавание линий) — отдельная спека, только если числа по `table` заметно хуже `text`.
6. Сравнение `correction` с эталоном — строгое, по `text_tokens`. Верное по сути исправление с другой кавычкой
   или «ё» уйдёт в `wrong_fix`. Поэтому `correction` пишется в `cases` целиком: человек просматривает `wrong_fix`
   глазами; мягкое сравнение вводится только после этого просмотра, не заранее.
7. Эталон собран от выхода тракта (PLACEHOLDER 3 спеки 10): незамеченная человеком ошибка считается в контроле
   «верным» текстом, и честное исправление модели там станет `false_alarm`. Каждый `false_alarm` человек сверяет
   со сканом; подтвердилось — новая строка в `errors.txt`, а не дефект модели.
8. `CROP_MARGIN = 20`, `CONTEXT_TOKENS = 3`, `CROP_RESOLUTION = 200` — стартовые. Менять до живого прогона
   бесплатно, после — `CONTEXT_TOKENS` и вырезка меняют ключ кэша.
9. Фрагменты документов заказчика и вырезки страниц уходят в облако Google. Для mineru.net это разрешено
   решением PLAN_OCR; для Google — тем, что человек заказал эту спеку. В интеграционной спеке записать отдельным решением.
10. `locate_block` берёт первый подходящий блок: повторяющаяся фраза на разных страницах даст вырезку не
    того места. Флаг `ambiguous` в `cases` — чтобы такие случаи исключались из выводов, а не портили их молча.

## Коммит

Два коммита, в этом порядке.

```
ПРАВКА #86: протокол Verifier и GeminiVerifier (REST, кэш ответов, вырезки)

pdf_core.py: Verifier — картинка вырезки + фрагмент + вопрос -> вердикт
(agree / fix / unreadable), уверенность 0–1, сырой ответ. ocr/gemini_verifier.py:
REST на requests без SDK, ключ GEMINI_API_KEY, модель — константа-PLACEHOLDER,
прокси — стандартные переменные окружения. Кэш сырых ответов в
.cache/ocr/verify/ по sha256(картинка + фрагмент + вопрос + модель); сеть только
при GEMINI_LIVE=1, промах кэша без флага — ошибка; пауза между сетевыми вызовами.
locate_block / crop_block: привязка места к блоку content_list и вырезка страницы
(pdfplumber, 200 dpi, bbox в шкале 0–1000). В тракт не подключено.
```

```
ПРАВКА #87: python -m ocr.measure — измерение vision-сверки на остатке эталонов

ocr/measure.py: по каждому опкоду остатка четырёх PDF-фикстур — вырезка, фрагмент,
вопрос, ответ модели; исходы found / not_found / wrong_fix. Контрольная выборка
верных фрагментов того же размера — false_alarm. DOCX/XLSX — no_image, в
знаменатель не входят. Таблицы по фикстурам, правилам валидатора и типу блока,
_test/verify_measure.json. Один опкод — один запрос. В pytest — фейковый Verifier,
сети нет; живой прогон делает человек. ingest, postprocess, validate, report.json
не тронуты: интеграция — после чисел.
```
