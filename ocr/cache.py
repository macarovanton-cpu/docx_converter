"""ПРАВКА #62: кэш сырого ответа OCR-провайдера (zip + meta.json)."""

import hashlib
import json
import os
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Protocol

CACHE_SCHEMA_VERSION = 1
DEFAULT_LOCAL_ROOT = Path(".cache") / "ocr"

# ключ годится как имя файла и на Windows, и на Linux
_SAFE_KEY_RE = re.compile(r"[0-9a-z,.\-]+")


def cache_key(pdf_bytes: bytes, provider: str, model_version: str | None,
              page_range: str | None = None) -> str:
    """sha256 содержимого + провайдер + модель + диапазон + версия схемы.

    Диапазон входит в ключ сверх PLAN: без него запрос других страниц молча
    получил бы чужой результат. CACHE_SCHEMA_VERSION поднимается при любом
    изменении содержимого zip или meta.json — старые записи перестают находиться.
    """
    pages = "".join(page_range.split()) if page_range else "all"
    return (f"{hashlib.sha256(pdf_bytes).hexdigest()}-{provider}"
            f"-{model_version or 'none'}-p{pages}-v{CACHE_SCHEMA_VERSION}")


def build_meta(key: str, pdf_bytes: bytes, zip_bytes: bytes, *, provider: str,
               model_version: str | None, page_range: str | None,
               pages: int) -> dict:
    """Спутник zip-а. Набор ключей закрытый — ровно эти восемь."""
    return {
        "key": key,
        "sha256": hashlib.sha256(pdf_bytes).hexdigest(),
        "provider": provider,
        "model_version": model_version,
        "page_range": page_range,
        "created_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "zip_size": len(zip_bytes),
        "pages": pages,
    }


class CacheBackend(Protocol):
    def get(self, key: str) -> tuple[bytes, dict] | None: ...

    def put(self, key: str, zip_bytes: bytes, meta: dict) -> None: ...


class LocalCache:
    """Кэш в папке: {root}/{key}.zip и {root}/{key}.meta.json."""

    def __init__(self, root: Path = DEFAULT_LOCAL_ROOT):
        self._root = Path(root)

    def get(self, key: str) -> tuple[bytes, dict] | None:
        zip_path, meta_path = self._paths(key)
        if not (zip_path.is_file() and meta_path.is_file()):
            return None                      # неполная запись — промах, не находка
        zip_bytes = zip_path.read_bytes()
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
        if meta["key"] != key or meta["zip_size"] != len(zip_bytes):
            raise RuntimeError(
                f"повреждённая запись кэша {zip_path}: meta.key={meta['key']!r}, "
                f"meta.zip_size={meta['zip_size']}, на диске {len(zip_bytes)} байт")
        return zip_bytes, meta

    def put(self, key: str, zip_bytes: bytes, meta: dict) -> None:
        zip_path, meta_path = self._paths(key)
        self._root.mkdir(parents=True, exist_ok=True)
        # сначала zip, потом meta: прерванная запись остаётся промахом, а не полузаписью
        _write_atomic(zip_path, zip_bytes)
        _write_atomic(meta_path,
                      json.dumps(meta, ensure_ascii=False, indent=2).encode("utf-8"))

    def _paths(self, key: str) -> tuple[Path, Path]:
        if not _SAFE_KEY_RE.fullmatch(key):
            raise ValueError(f"ключ кэша {key!r}: допустимы только [0-9a-z,.-]")
        return self._root / f"{key}.zip", self._root / f"{key}.meta.json"


def _write_atomic(path: Path, data: bytes) -> None:
    """Временный файл в той же папке + os.replace."""
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_bytes(data)
    os.replace(tmp, path)


def make_cache(kind: str) -> CacheBackend:
    if kind == "local":
        return LocalCache()
    if kind == "drive":
        raise NotImplementedError(
            "Кэш на Google Drive не реализован: появится на этапе 8 вместе с "
            "drive_client.py. Используйте --cache local.")
    raise KeyError(kind)                     # как неизвестный doc_style — без фолбэка
