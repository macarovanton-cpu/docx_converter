# 02 — MineruProvider

**# ПРАВКА #61.** Зависит от: 01. Единственная спека, где нужна сеть (один live-тест).

## Цель

Провайдер за протоколом `OcrProvider` (спека 01): PDF уходит в облако MinerU,
возвращается zip, из него собирается `OcrResult`. Получение zip и сборка
результата разделены — кэш (03) хранит zip, CLI (07) собирает результат из кэша
без сети.

## Трогать

- `ocr/__init__.py` — создать, если его ещё нет (содержимое — дословно из README)
- `ocr/mineru_provider.py` — создать
- `tests/test_mineru_provider.py` — создать
- `requirements.txt` — добавить одну строку `requests`

## Не трогать

`pdf_core.py` (только импорт из него), всё из общего списка запретов.

## Порядок работы с API v4 (по mineru.net/apiManage/docs)

1. `POST {API_BASE}/file-urls/batch`, заголовки `Authorization: Bearer <key>`,
   `Content-Type: application/json`. Тело:

   ```json
   {"model_version": "vlm", "language": "east_slavic",
    "enable_table": true, "enable_formula": true,
    "files": [{"name": "<sha256[:16]>.pdf", "is_ocr": true,
               "data_id": "<sha256[:32]>", "page_ranges": "2,4-6"}]}
   ```

   Все параметры распознавания передаются **здесь**. `name` — хэш, а не имя файла
   заказчика: имя в облако не уходит. `page_ranges` — только если `page_range`
   задан. PLACEHOLDER: место `page_ranges` (в элементе `files` или на верхнем
   уровне) сверить с докой на live-прогоне. **Сверено по доке 21.09.2026 (спека 11):** `page_ranges` — поле элемента `files[]`, код верен; на верхнем уровне оно только у одиночного `extract/task`.
   Ответ: `{"code": 0, "data": {"batch_id": "…", "file_urls": ["https://…"]}, "msg": "ok"}`.
2. `PUT file_urls[0]`, тело — байты PDF, **без** `Authorization` и без `Content-Type`.
   Отдельного шага «создать задачу» нет: задача ставится сама после загрузки.
3. Опрос `GET {API_BASE}/extract-results/batch/{batch_id}` с `Authorization`.
   Ответ: `data.extract_result[0]` с полями `state`, `full_zip_url`, `err_msg`.
   `state`: `waiting-file | pending | running | converting` — ждать; `done` — готово;
   `failed` — `MineruError(err_msg)`. Неизвестное состояние — `MineruError`, не ожидание.
4. `GET full_zip_url` без `Authorization` → байты zip.

`code != 0` в любом ответе: `A0202`, `A0211` → `MineruAuthError`; `-60005`, `-60006` →
`MineruLimitError`; остальное → `MineruError(f"{code}: {msg}")`.
PLACEHOLDER: код «исчерпана квота» неизвестен — `MineruQuotaError` объявлен, на
него маппится код из константы `QUOTA_CODES: frozenset[str] = frozenset()`. **Сверено по доке 21.09.2026 (спека 11):** код квоты — `-60018` («Daily extract task limit reached»), на живом ответе не наблюдался; в коде `QUOTA_CODES` пуст до отдельной правки провайдера.

## Интерфейсы (дословно)

```python
"""ПРАВКА #61: провайдер MinerU cloud API v4 за протоколом pdf_core.OcrProvider."""

API_BASE = "https://mineru.net/api/v4"
MAX_BYTES = 200 * 1024 * 1024
MAX_PAGES = 200                       # по apiManage/docs; переопределяется в конструкторе
RETRY_PAUSES_SEC = (1.0, 2.0, 4.0)
QUOTA_CODES: frozenset[str] = frozenset()   # PLACEHOLDER: код «квота исчерпана» неизвестен


class MineruError(RuntimeError): ...
class MineruAuthError(MineruError): ...     # A0202 / A0211 / нет ключа
class MineruQuotaError(MineruError): ...
class MineruLimitError(MineruError): ...    # предпроверка и -60005 / -60006
class MineruTimeout(MineruError): ...


class MineruProvider:
    def __init__(self, api_key: str | None = None, *,
                 model_version: str = "vlm",
                 language: str = "east_slavic",
                 max_pages: int = MAX_PAGES,
                 poll_interval_sec: float = 5.0,
                 timeout_sec: float = 900.0,
                 raw_root: Path | None = None,
                 session=None,
                 sleep=time.sleep,
                 analyze_func=analyze_pdf_pages): ...

    def page_infos(self, pdf_bytes: bytes,
                   page_range: str | None = None) -> list[PageInfo]: ...

    def fetch_raw_zip(self, pdf_bytes: bytes,
                      page_range: str | None = None) -> bytes: ...

    def ocr_pdf(self, pdf_bytes: bytes,
                page_range: str | None = None) -> OcrResult: ...


def result_from_zip(zip_bytes: bytes, raw_dir: Path, *,
                    model_version: str | None,
                    pages: list[PageInfo]) -> OcrResult: ...
```

Поведение:

- `api_key=None` → `os.environ["MINERU_API_KEY"]`; нет и там → `MineruAuthError`
  **в конструкторе** с текстом, где назван `MINERU_API_KEY`. Ключ не логируется и
  не попадает в тексты исключений.
- `model_version` вне `("vlm", "pipeline")` → `ValueError`.
- `session=None` → `requests.Session()`. В тестах — фейк с методами `post/put/get`.
- `page_infos`: `analyze_func` по временному файлу (переиспользовать
  `pdf_core._write_temp_pdf` / `_unlink_quiet`) + `ocr_auto_mode.selected_pdf_pages`;
  `ocr_applied = not has_text_layer`.
- `fetch_raw_zip`: предпроверка **до любого запроса** — `len(pdf_bytes) > MAX_BYTES`
  или число страниц `> max_pages` → `MineruLimitError` с фактическим и предельным
  значением. `is_ocr = any(not p.has_text_layer for p in pages)`. Затем шаги 1–4.
- Ретраи: HTTP 5xx и `requests.ConnectionError`/`requests.Timeout` — повтор с паузами
  `RETRY_PAUSES_SEC` через `self._sleep`; после последней — `MineruError`. 4xx и
  `code != 0` не ретраятся.
- Опрос: `self._sleep(poll_interval_sec)` между запросами; суммарно больше
  `timeout_sec` → `MineruTimeout` с `batch_id` в тексте.
- `ocr_pdf` = `result_from_zip(fetch_raw_zip(...), raw_dir, model_version=…, pages=page_infos(...))`,
  где `raw_dir = (raw_root or Path(tempfile.gettempdir()) / "ocr_raw") / f"{sha256[:16]}-{model_version}"`.
- `result_from_zip`: распаковать в `raw_dir` (создать; существующую — очистить).
  Имя члена архива абсолютное или с `..` → `MineruError` (zip-slip), ничего не
  распаковывать. `markdown` — файл с именем `full.md` (в корне или глубже); нет —
  `MineruError`. `content_list` — файл, чьё имя оканчивается на `content_list.json`;
  нет — `None` (не ошибка). `provider="mineru"`. Возврат:
  `OcrResult(markdown, content_list, pages, "mineru", model_version, raw_dir)`.

## Приёмочные тесты (`tests/test_mineru_provider.py`)

Сеть не используется: `FakeSession` отдаёт заранее заданные ответы и пишет журнал
вызовов; `sleep` — список, в который складываются паузы. Ответы синтетические, по
форме из раздела «Порядок работы» (PLACEHOLDER: записанных настоящих нет).
Помощник `make_zip(full_md: str, content_list: list | None) -> bytes` собирает zip
в памяти.

```python
pdf = require_fixture("bakeoff.pdf").read_bytes()
vlm = read_fixture("vlm.md")
raw = make_zip(vlm, [{"type": "text", "text": "Утверждаю:", "page_idx": 0}])

# счастливый путь
provider = MineruProvider("k", session=session, sleep=pauses.append, raw_root=tmp_path)
result = provider.ocr_pdf(pdf)
assert result.markdown == vlm
assert (result.provider, result.model_version) == ("mineru", "vlm")
assert result.content_list[0]["page_idx"] == 0
assert [p.number for p in result.pages] == list(range(1, 10))
assert (result.raw_dir / "full.md").read_text(encoding="utf-8") == vlm

# тело первого запроса
body = session.calls[0].json
assert session.calls[0].url == API_BASE + "/file-urls/batch"
assert body["model_version"] == "vlm" and body["language"] == "east_slavic"
assert body["files"][0]["is_ocr"] is True                 # bakeoff.pdf — скан
assert body["files"][0]["name"] == hashlib.sha256(pdf).hexdigest()[:16] + ".pdf"
assert "bakeoff" not in json.dumps(body)                  # имя файла в облако не уходит
assert "page_ranges" not in body["files"][0]
# PUT без авторизации, байты те же
assert session.calls[1].method == "PUT" and session.calls[1].data == pdf
assert "Authorization" not in (session.calls[1].headers or {})
# опрос: pending, running, done → две паузы по 5 с
assert pauses == [5.0, 5.0]

# ошибки
with pytest.raises(MineruAuthError, match="MINERU_API_KEY"):      # без ключа и без env
    MineruProvider(None)
with pytest.raises(MineruAuthError):                              # code "A0211"
    ...
assert MAX_PAGES == 200
with pytest.raises(MineruLimitError) as excinfo:                  # max_pages=5, в файле 9
    MineruProvider("k", session=session, max_pages=5).fetch_raw_zip(pdf)
assert "9" in str(excinfo.value) and "5" in str(excinfo.value)    # факт и предел названы
assert session.calls == []                                        # до сети не дошло
with pytest.raises(MineruError, match="boom"):                    # state "failed", err_msg "boom"
    ...
with pytest.raises(MineruTimeout):                                # вечный "running", timeout_sec=12
    ...
with pytest.raises(ValueError):
    MineruProvider("k", model_version="best")

# ретраи: 502, 502, затем 200
assert pauses[:2] == [1.0, 2.0]
# 4 × 502 подряд → MineruError, ровно три паузы
assert pauses == [1.0, 2.0, 4.0]
# ключ не утекает
assert "k-secret" not in str(excinfo.value)

# result_from_zip
with pytest.raises(MineruError):
    result_from_zip(make_zip_with_member("../evil.txt"), tmp_path / "r", model_version="vlm", pages=[])
assert not (tmp_path / "evil.txt").exists()
no_cl = result_from_zip(make_zip(vlm, None), tmp_path / "r2", model_version="vlm", pages=[])
assert no_cl.content_list is None
with pytest.raises(MineruError, match="full.md"):
    result_from_zip(make_zip_without_full_md(), tmp_path / "r3", model_version="vlm", pages=[])
```

Живой тест:

```python
@pytest.mark.live
def test_live_bakeoff_close_to_vlm(tmp_path):
    pdf = require_fixture("bakeoff.pdf").read_bytes()
    provider = MineruProvider(raw_root=tmp_path)
    zip_bytes = provider.fetch_raw_zip(pdf)
    (FIXTURES / "vlm_raw.zip").write_bytes(zip_bytes)      # закрывает PLACEHOLDER про content_list
    result = result_from_zip(zip_bytes, tmp_path / "raw", model_version="vlm",
                             pages=provider.page_infos(pdf))
    a, b = text_tokens(result.markdown), text_tokens(read_fixture("vlm.md"))
    ratio = difflib.SequenceMatcher(None, a, b, autojunk=False).ratio()
    assert ratio >= 0.95                                    # PLACEHOLDER: порог стартовый
    assert result.content_list
```

## Готово, когда

- `pytest -v` зелёный без `MINERU_API_KEY` (live пропущен, остальное прошло).
- `python -c "import ocr.mineru_provider"` не тянет `streamlit`.
- В `requirements.txt` ровно одна новая строка: `requests`.
- Если ключ есть: live-тест прошёл, `vlm_raw.zip` лежит в фикстурах; в отчёте
  исполнителя — фактический `ratio` и место `page_ranges` в теле запроса.

## Коммит

```
ПРАВКА #61: MineruProvider — облачный OCR MinerU за протоколом OcrProvider

ocr/mineru_provider.py: загрузка через file-urls/batch, опрос по batch_id,
ретраи 5xx, предпроверка 200 МБ / 200 страниц, понятные ошибки по токену.
fetch_raw_zip и result_from_zip разделены под кэш. В облако уходит хэш вместо
имени файла. requirements.txt: + requests.
```
