# 03 — Кэш сырого ответа

**# ПРАВКА #62.** Зависит от: 01 (формат `OcrResult`), 02 (`result_from_zip`).

## Цель

Повторный прогон того же PDF не ходит в облако: сырой zip MinerU и `meta.json`
лежат в кэше под ключом от содержимого файла и параметров распознавания.

Объём этой спеки: **только локальный бэкенд**. Бэкенд `drive`, `drive_client.py`
и правка `app.py` перенесены в этап 8 (вне автономного прогона); здесь
`make_cache("drive")` честно отвечает `NotImplementedError`.

## Трогать

- `ocr/cache.py` — создать
- `tests/test_ocr_cache.py` — создать
- `.gitignore` — добавить строку `.cache/`

## Не трогать

`app.py` (никак), `ocr/mineru_provider.py`, `pdf_core.py`, всё из общего списка запретов.

## Интерфейсы (дословно)

```python
"""ПРАВКА #62: кэш сырого ответа OCR-провайдера (zip + meta.json)."""

CACHE_SCHEMA_VERSION = 1
DEFAULT_LOCAL_ROOT = Path(".cache") / "ocr"


def cache_key(pdf_bytes: bytes, provider: str, model_version: str | None,
              page_range: str | None = None) -> str: ...


def build_meta(key: str, pdf_bytes: bytes, zip_bytes: bytes, *, provider: str,
               model_version: str | None, page_range: str | None,
               pages: int) -> dict: ...


class CacheBackend(Protocol):
    def get(self, key: str) -> tuple[bytes, dict] | None: ...
    def put(self, key: str, zip_bytes: bytes, meta: dict) -> None: ...


class LocalCache:
    def __init__(self, root: Path = DEFAULT_LOCAL_ROOT): ...
    def get(self, key: str) -> tuple[bytes, dict] | None: ...
    def put(self, key: str, zip_bytes: bytes, meta: dict) -> None: ...


def make_cache(kind: str) -> CacheBackend: ...
```

### Ключ

```
{sha256(pdf_bytes).hexdigest()}-{provider}-{model_version or "none"}-p{range}-v{CACHE_SCHEMA_VERSION}
```

`range` — `page_range` без пробелов, либо `all`, если `page_range` пуст/`None`.
Пример: `3f…9a-mineru-vlm-pall-v1`. Символы ключа — только `[0-9a-z,.\-]`:
годится как имя файла на Windows и Linux.

`page_range` добавлен в ключ сверх PLAN: без него запрос другого диапазона молча
получил бы чужой результат. `CACHE_SCHEMA_VERSION` поднимается при любом изменении
того, что лежит в zip или в `meta.json`, — старые записи перестают находиться сами.

### `meta.json`

```json
{"key": "3f…9a-mineru-vlm-pall-v1",
 "sha256": "3f…9a",
 "provider": "mineru",
 "model_version": "vlm",
 "page_range": null,
 "created_at": "2026-09-18T12:00:00Z",
 "zip_size": 1234567,
 "pages": 9}
```

`created_at` — UTC, ISO-8601, секунды, суффикс `Z`. Набор ключей закрытый —
ровно эти восемь.

### `LocalCache`

- Файлы: `{root}/{key}.zip` и `{root}/{key}.meta.json` (`utf-8`, `ensure_ascii=False`, `indent=2`).
- `put`: создаёт `root`; пишет через временный файл в той же папке + `os.replace`,
  сначала zip, потом meta — прерванная запись не оставляет «полузаписи», которую
  `get` принял бы за валидную.
- `get`: нет любого из двух файлов → `None`. Оба есть, но
  `meta["key"] != key` или `meta["zip_size"] != len(zip_bytes)` → `RuntimeError`
  с путём к записи (повреждённый кэш не глотать и не чинить молча).
- `key`, содержащий что-либо вне `[0-9a-z,.\-]`, → `ValueError` (защита от выхода из `root`).

### `make_cache`

- `"local"` → `LocalCache()`.
- `"drive"` → `NotImplementedError("Кэш на Google Drive не реализован: появится на этапе 8 вместе с drive_client.py. Используйте --cache local.")`.
- иное → `KeyError(kind)` (как неизвестный `doc_style` в `convert.py` — без фолбэка).

## Приёмочные тесты (`tests/test_ocr_cache.py`)

Zip для тестов собирается в памяти из `vlm.md` (`full.md` внутри).

```python
pdf = require_fixture("bakeoff.pdf").read_bytes()
vlm = read_fixture("vlm.md")
sha = hashlib.sha256(pdf).hexdigest()

# ключ
assert cache_key(pdf, "mineru", "vlm") == f"{sha}-mineru-vlm-pall-v1"
assert cache_key(pdf, "mineru", "pipeline") != cache_key(pdf, "mineru", "vlm")
assert cache_key(pdf, "ocrmypdf", None) == f"{sha}-ocrmypdf-none-pall-v1"
assert cache_key(pdf, "mineru", "vlm", "1-3, 7") == f"{sha}-mineru-vlm-p1-3,7-v1"
assert cache_key(pdf, "mineru", "vlm", "") == cache_key(pdf, "mineru", "vlm", None)
assert cache_key(pdf + b"x", "mineru", "vlm") != cache_key(pdf, "mineru", "vlm")
assert re.fullmatch(r"[0-9a-z,.\-]+", cache_key(pdf, "mineru", "vlm", "1-3, 7"))

# раунд-трип
key = cache_key(pdf, "mineru", "vlm")
cache = LocalCache(tmp_path)
assert cache.get(key) is None
meta = build_meta(key, pdf, zip_bytes, provider="mineru", model_version="vlm",
                  page_range=None, pages=9)
cache.put(key, zip_bytes, meta)
got_zip, got_meta = cache.get(key)
assert got_zip == zip_bytes and got_meta == meta
assert set(got_meta) == {"key", "sha256", "provider", "model_version",
                         "page_range", "created_at", "zip_size", "pages"}
assert got_meta["sha256"] == sha and got_meta["zip_size"] == len(zip_bytes)
assert re.fullmatch(r"\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ", got_meta["created_at"])
assert sorted(p.name for p in tmp_path.iterdir()) == [f"{key}.meta.json", f"{key}.zip"]

# кэш стыкуется с провайдером: из кэша собирается тот же результат, сети нет
result = result_from_zip(got_zip, tmp_path / "raw", model_version="vlm", pages=[])
assert result.markdown == vlm

# повреждения не глотаются
(tmp_path / f"{key}.zip").write_bytes(b"short")
with pytest.raises(RuntimeError, match=key):
    cache.get(key)
(tmp_path / f"{key}.meta.json").unlink()
assert cache.get(key) is None                      # неполная запись = промах

# защита пути
with pytest.raises(ValueError):
    cache.get("../../etc/passwd")

# фабрика
assert isinstance(make_cache("local"), LocalCache)
with pytest.raises(NotImplementedError, match="этапе 8"):
    make_cache("drive")
with pytest.raises(KeyError):
    make_cache("s3")
```

## Готово, когда

- `pytest -v` зелёный.
- `git diff --stat`: `ocr/cache.py`, `tests/test_ocr_cache.py`, `.gitignore` (+1 строка).
- `git status` после прогона тестов не показывает `.cache/`.
- `app.py` не изменён: `git diff app.py` пуст.

## Коммит

```
ПРАВКА #62: локальный кэш сырого ответа OCR (zip + meta.json)

ocr/cache.py: cache_key (sha256 + provider + model_version + диапазон + версия
схемы), LocalCache с атомарной записью, make_cache. Повреждённая запись —
ошибка, не промах. Бэкенд drive — NotImplementedError до этапа 8.
.gitignore: + .cache/
```
