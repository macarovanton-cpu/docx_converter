"""ПРАВКА #83: эталоны по фикстурам — sha256, порог остатка, воспроизводимость из errors.txt."""

import hashlib
import shutil
import socket
from pathlib import Path

import pytest

from ocr import board as board_module
from ocr.board import BOARD, GOLDEN, THRESHOLDS, apply_errors, build_board, main, parse_errors, parse_errors_numbered
from ocr_fixtures import count_diffs, read_fixture, require_fixture

# эталоны вне git — sha ловит тихую подмену
GOLDEN_SHA256 = {
    "bakeoff2": "fe0d3f8886de482e7e2ffe23041ebf3ecd9a2a72e0c1f9b5342a70bfe6e461bc",
    "bakeoff3": "65d67c905b32661123418dfac5c539448f3ddcdb930a09b22de2da7dfb14d63e",
    "textpdf1": "195d349e1c17834f3b6fe07ec84a78642c7e44b5e5e62e654afb00f10ce9ab2d",
    "docx1": "317e9138abdbf8b44126025133eaf3927fc35d7e96efb99e55eafbd60eade9f6",
    "xlsx1": "1f019aef6a7745bb977fbeb85947533288f9133addee6cd91270a89b8bb7bb2e",
}
FIXTURE_BY_STEM = {Path(name).stem: name for name, _ in BOARD}


@pytest.fixture(autouse=True)
def no_network(monkeypatch):
    def refuse(*args, **kwargs):
        raise AssertionError("тест полез в сеть")
    monkeypatch.setattr(socket, "socket", refuse)


def read_exact(path: Path) -> str:
    return path.read_bytes().decode("utf-8")        # без перевода концов строк


def test_goldens(tmp_path):
    for pair in BOARD:
        for name in pair:
            if name is not None:
                require_fixture(name)
    for golden_name in GOLDEN.values():
        require_fixture(golden_name)
    _, outputs = build_board(tmp_path)

    for name, golden_name in GOLDEN.items():
        golden = read_fixture(golden_name)
        assert count_diffs(golden, golden) == 0
        assert count_diffs(outputs[name][0], golden) <= THRESHOLDS[name]

    for stem, sha in GOLDEN_SHA256.items():
        path = require_fixture(f"{stem}.golden.md")
        assert hashlib.sha256(path.read_bytes()).hexdigest() == sha
        # эталон воспроизводим из черновика и списка человека — руками его не правили
        errors = parse_errors(read_fixture(f"{stem}.errors.txt"))
        out = outputs[FIXTURE_BY_STEM[stem]][0]
        assert apply_errors(out, errors) == read_exact(path)
        if not errors:
            assert count_diffs(out, read_exact(path)) == 0      # якорь


def test_parse_and_apply_refuse_to_guess():
    assert apply_errors("текст", []) == "текст"                 # пустой список — якорь
    assert parse_errors("# шапка\n\n3 | а \\| б | в\n") == [("3", "а | б", "в")]
    assert parse_errors("- | лишнее |\n") == [("-", "лишнее", "")]            # пустое «надо» = удалить
    with pytest.raises(ValueError, match="строка 1"):
        parse_errors("3 | только два поля\n")
    with pytest.raises(ValueError, match="строка 2"):
        parse_errors("# шапка\n3 |  | надо\n")                  # пустое «было»
    with pytest.raises(ValueError, match="вхождений 2"):
        apply_errors("аа аа", [("1", "аа", "б")])
    with pytest.raises(ValueError, match="вхождений 0"):
        apply_errors("текст", [("1", "нет такого", "б")])
    with pytest.raises(ValueError, match="вхождений 2"):        # неуникальна после предыдущей правки
        apply_errors("аб вб", [("1", "в", "а"), ("1", "аб", "х")])

def test_apply_errors_names_the_line():
    """ПРАВКА #84: номер строки errors.txt в ошибке; без lines — поведение прежнее."""
    text = "# шапка\n\n3 | а | б\n4 | нет такого | в\n"
    numbered = parse_errors_numbered(text)
    assert numbered == [(3, ("3", "а", "б")), (4, ("4", "нет такого", "в"))]
    assert parse_errors(text) == [error for _, error in numbered]           # обёртка, поведение прежнее

    lines = [n for n, _ in numbered]
    errors = [e for _, e in numbered]
    with pytest.raises(ValueError, match=r"строка 4 errors\.txt.*вхождений 0"):
        apply_errors("а", errors, lines=lines)
    with pytest.raises(ValueError, match=r"строка 3 errors\.txt.*вхождений 2"):
        apply_errors("а а", errors, lines=lines)
    with pytest.raises(ValueError, match="lines"):
        apply_errors("а", errors, lines=[3])                                  # длины не совпали
    with pytest.raises(ValueError, match="вхождений 0") as info:
        apply_errors("текст", [("1", "нет такого", "б")])                     # без lines — как было
    assert "строка" not in str(info.value)


def test_golden_flag_keeps_manual_work(tmp_path, monkeypatch, capsys):
    fixtures = tmp_path / "fx"
    fixtures.mkdir()
    for pair in BOARD:
        for name in pair:
            if name is not None:
                shutil.copy(require_fixture(name), fixtures / name)
    shutil.copy(require_fixture("golden.md"), fixtures / "golden.md")
    for stem in GOLDEN_SHA256:
        shutil.copy(require_fixture(f"{stem}.errors.txt"), fixtures / f"{stem}.errors.txt")
    monkeypatch.setattr(board_module, "FIXTURES", fixtures)
    monkeypatch.setattr(board_module, "BOARD_JSON", tmp_path / "quality_board.json")
    monkeypatch.setattr(board_module, "DRAFTS", tmp_path / "board")
    bakeoff_golden_before = (fixtures / "golden.md").read_bytes()

    assert main(["--golden"]) == 0
    for stem, sha in GOLDEN_SHA256.items():
        assert hashlib.sha256((fixtures / f"{stem}.golden.md").read_bytes()).hexdigest() == sha
    assert main(["--golden"]) == 1                              # эталон уже есть
    assert "bakeoff2.golden.md" in capsys.readouterr().err
    assert main(["--golden", "--force"]) == 0
    assert (fixtures / "golden.md").read_bytes() == bakeoff_golden_before   # bakeoff неприкосновенен
    # ПРАВКА #84: отказ называет файл и номер строки, эталоны остаются прежними
    xlsx_errors = fixtures / "xlsx1.errors.txt"
    xlsx_before = xlsx_errors.read_bytes()
    text = xlsx_errors.read_text(encoding="utf-8")
    text += "" if text.endswith("\n") else "\n"
    xlsx_errors.write_text(text + "- | такого текста в черновике нет | x\n", encoding="utf-8")
    capsys.readouterr()
    assert main(["--golden", "--force"]) == 1
    err = capsys.readouterr().err
    assert "xlsx1.errors.txt" in err
    assert f"строка {len(text.splitlines()) + 1}" in err
    for stem, sha in GOLDEN_SHA256.items():                     # отказ не оставляет половину
        assert hashlib.sha256((fixtures / f"{stem}.golden.md").read_bytes()).hexdigest() == sha
    xlsx_errors.write_bytes(xlsx_before)

    (fixtures / "xlsx1.errors.txt").unlink()                    # пустой список подтверждается файлом
    assert main(["--golden", "--force"]) == 1
    assert "xlsx1.errors.txt" in capsys.readouterr().err
