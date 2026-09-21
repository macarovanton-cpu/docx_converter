"""Общие помощники тестов OCR-тракта: пути к фикстурам, пропуски, метрика расхождений."""

import json
import zipfile
from pathlib import Path

import pytest

from ocr.board import count_diffs, text_tokens  # ПРАВКА #83

FIXTURES = Path(__file__).resolve().parents[1] / "_test" / "fixtures" / "ocr"


def require_fixture(name: str) -> Path:
    """Путь к фикстуре; pytest.skip, если файла нет (папка в .gitignore)."""
    path = FIXTURES / name
    if not path.is_file():
        pytest.skip(f"нет фикстуры {name}: {FIXTURES} (папка в .gitignore)")
    return path


def read_fixture(name: str) -> str:
    """require_fixture + read_text(encoding='utf-8')."""
    return require_fixture(name).read_text(encoding="utf-8")


def read_raw(name: str) -> tuple[str, list]:
    """ПРАВКА #74: (full.md, content_list) из сырого архива MinerU.

    content_list.json лежит под uuid-префиксом, рядом — content_list_v2.json,
    который под суффикс не подходит и берётся не он.
    """
    with zipfile.ZipFile(require_fixture(name)) as archive:
        markdown = archive.read("full.md").decode("utf-8")
        names = [n for n in archive.namelist() if n.endswith("content_list.json")]
        assert len(names) == 1, names
        content_list = json.loads(archive.read(names[0]).decode("utf-8"))
    return markdown, content_list
