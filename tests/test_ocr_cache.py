"""ПРАВКА #62: приёмочные тесты кэша сырого ответа OCR.

Сети нет: zip собирается в памяти из фикстуры vlm.md (full.md внутри).
"""

import hashlib
import io
import re
import zipfile

import pytest

from ocr_fixtures import read_fixture, require_fixture
from ocr.cache import LocalCache, build_meta, cache_key, make_cache
from ocr.mineru_provider import result_from_zip


def make_zip(full_md: str) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("full.md", full_md)
    return buffer.getvalue()


def test_cache_key_covers_content_provider_model_and_range():
    pdf = require_fixture("bakeoff.pdf").read_bytes()
    sha = hashlib.sha256(pdf).hexdigest()

    assert cache_key(pdf, "mineru", "vlm") == f"{sha}-mineru-vlm-pall-v1"
    assert cache_key(pdf, "mineru", "pipeline") != cache_key(pdf, "mineru", "vlm")
    assert cache_key(pdf, "ocrmypdf", None) == f"{sha}-ocrmypdf-none-pall-v1"
    assert cache_key(pdf, "mineru", "vlm", "1-3, 7") == f"{sha}-mineru-vlm-p1-3,7-v1"
    assert cache_key(pdf, "mineru", "vlm", "") == cache_key(pdf, "mineru", "vlm", None)
    assert cache_key(pdf + b"x", "mineru", "vlm") != cache_key(pdf, "mineru", "vlm")
    assert re.fullmatch(r"[0-9a-z,.\-]+", cache_key(pdf, "mineru", "vlm", "1-3, 7"))


def test_local_cache_round_trip(tmp_path):
    pdf = require_fixture("bakeoff.pdf").read_bytes()
    zip_bytes = make_zip(read_fixture("vlm.md"))
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
    assert got_meta["sha256"] == hashlib.sha256(pdf).hexdigest()
    assert got_meta["zip_size"] == len(zip_bytes)
    assert re.fullmatch(r"\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ", got_meta["created_at"])
    assert sorted(p.name for p in tmp_path.iterdir()) == [f"{key}.meta.json",
                                                          f"{key}.zip"]


def test_cached_zip_rebuilds_the_same_result(tmp_path):
    """Кэш стыкуется с провайдером: из записи собирается тот же результат, сети нет."""
    pdf = require_fixture("bakeoff.pdf").read_bytes()
    vlm = read_fixture("vlm.md")
    zip_bytes = make_zip(vlm)
    key = cache_key(pdf, "mineru", "vlm")
    cache = LocalCache(tmp_path)
    cache.put(key, zip_bytes, build_meta(key, pdf, zip_bytes, provider="mineru",
                                         model_version="vlm", page_range=None, pages=9))

    got_zip, _ = cache.get(key)
    result = result_from_zip(got_zip, tmp_path / "raw", model_version="vlm", pages=[])
    assert result.markdown == vlm


def test_damaged_entry_raises_and_incomplete_entry_is_a_miss(tmp_path):
    pdf = b"%PDF-1.4 fake"
    zip_bytes = make_zip("# md")
    key = cache_key(pdf, "mineru", "vlm")
    cache = LocalCache(tmp_path)
    cache.put(key, zip_bytes, build_meta(key, pdf, zip_bytes, provider="mineru",
                                         model_version="vlm", page_range=None, pages=1))

    (tmp_path / f"{key}.zip").write_bytes(b"short")
    with pytest.raises(RuntimeError, match=key):
        cache.get(key)

    (tmp_path / f"{key}.meta.json").unlink()
    assert cache.get(key) is None


def test_key_outside_charset_is_rejected(tmp_path):
    cache = LocalCache(tmp_path)
    with pytest.raises(ValueError):
        cache.get("../../etc/passwd")
    with pytest.raises(ValueError):
        cache.put("../../etc/passwd", b"zip", {})


def test_make_cache():
    assert isinstance(make_cache("local"), LocalCache)
    with pytest.raises(NotImplementedError, match="этапе 8"):
        make_cache("drive")
    with pytest.raises(KeyError):
        make_cache("s3")
