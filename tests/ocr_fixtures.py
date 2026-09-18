"""Общие помощники тестов OCR-тракта: пути к фикстурам, пропуски, метрика расхождений."""

import difflib
import re
from pathlib import Path

import pytest

FIXTURES = Path(__file__).resolve().parents[1] / "_test" / "fixtures" / "ocr"

_SEPARATOR_RE = re.compile(r"^\s*\|?[\s:|-]+\|?\s*$")


def require_fixture(name: str) -> Path:
    """Путь к фикстуре; pytest.skip, если файла нет (папка в .gitignore)."""
    path = FIXTURES / name
    if not path.is_file():
        pytest.skip(f"нет фикстуры {name}: {FIXTURES} (папка в .gitignore)")
    return path


def read_fixture(name: str) -> str:
    """require_fixture + read_text(encoding='utf-8')."""
    return require_fixture(name).read_text(encoding="utf-8")


def text_tokens(md: str) -> list[str]:
    r"""Токены текста без разметки.

    1. строки-разделители pipe-таблиц (^\s*\|?[\s:|-]+\|?\s*$ с хотя бы одним '-') удалить;
    2. HTML-теги <[^>]+> заменить пробелом;
    3. символ '|' заменить пробелом;
    4. str.split().
    """
    kept = [line for line in md.split("\n")
            if not ("-" in line and _SEPARATOR_RE.match(line))]
    text = re.sub(r"<[^>]+>", " ", "\n".join(kept))
    return text.replace("|", " ").split()


def count_diffs(a: str, b: str) -> int:
    """Число не-'equal' опкодов
    difflib.SequenceMatcher(None, text_tokens(a), text_tokens(b), autojunk=False)."""
    matcher = difflib.SequenceMatcher(None, text_tokens(a), text_tokens(b), autojunk=False)
    return sum(1 for tag, *_ in matcher.get_opcodes() if tag != "equal")
