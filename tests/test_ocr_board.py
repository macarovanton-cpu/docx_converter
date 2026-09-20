"""ПРАВКА #82: приёмочные тесты табло качества ocr.board. Сети нет вообще."""

import json
import shutil
import socket

import pytest

from ocr import SEVERITIES
from ocr import board as board_module
from ocr.board import BOARD, ERRORS_HEADER, build_board, main
from ocr_fixtures import require_fixture

ROW_KEYS = ["fixture", "route", "provider", "critical", "warning", "info", "by_rule",
            "tables", "table_rows", "table_broken", "chars", "tokens",
            "count_diffs", "threshold"]


@pytest.fixture(autouse=True)
def no_network(monkeypatch):
    def refuse(*args, **kwargs):
        raise AssertionError("тест полез в сеть")
    monkeypatch.setattr(socket, "socket", refuse)


def copy_fixtures(target, skip=()):
    target.mkdir(parents=True)
    for pair in BOARD:
        for name in pair:
            if name is not None and name not in skip:
                shutil.copy(require_fixture(name), target / name)
    return target


def test_build_board(tmp_path):
    for pair in BOARD:
        for name in pair:
            if name is not None:
                require_fixture(name)
    board, outputs = build_board(tmp_path)
    rows = board["rows"]
    assert board["schema_version"] == 1
    assert [r["fixture"] for r in rows] == [name for name, _ in BOARD] and len(rows) == 6
    assert all(list(r) == ROW_KEYS for r in rows)               # одна метрика на все форматы
    assert all(r["count_diffs"] is None and r["threshold"] is None for r in rows)
    for r in rows:
        md, report = outputs[r["fixture"]]
        assert (r["critical"], r["warning"], r["info"]) == tuple(
            report["summary"][s] for s in SEVERITIES)
        assert sum(r["by_rule"].values()) == len(report["findings"])
        assert list(r["by_rule"]) == sorted(r["by_rule"])
        assert r["chars"] == len(md) > 0
    assert next(r for r in rows if r["fixture"] == "bakeoff.pdf")["critical"] == 5   # как в test_ocr_cli
    assert next(r for r in rows if r["fixture"] == "xlsx1.xlsx")["tables"] >= 1


def test_cache_miss_fails_offline(tmp_path):
    fixtures = copy_fixtures(tmp_path / "fx", skip=("vlm_raw3.zip",))
    with pytest.raises(RuntimeError, match="оффлайн"):
        build_board(tmp_path / "x", fixtures=fixtures)


def test_missing_fixture_is_exit_code_1(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(board_module, "FIXTURES", tmp_path / "нет")
    monkeypatch.setattr(board_module, "BOARD_JSON", tmp_path / "quality_board.json")
    monkeypatch.setattr(board_module, "DRAFTS", tmp_path / "board")
    assert main([]) == 1
    assert "FileNotFoundError" in capsys.readouterr().err
    assert not (tmp_path / "quality_board.json").exists()


def test_main_writes_files_and_keeps_manual_work(tmp_path, monkeypatch, capsys):
    fixtures = copy_fixtures(tmp_path / "fx")
    board_json, drafts = tmp_path / "quality_board.json", tmp_path / "board"
    monkeypatch.setattr(board_module, "FIXTURES", fixtures)
    monkeypatch.setattr(board_module, "BOARD_JSON", board_json)
    monkeypatch.setattr(board_module, "DRAFTS", drafts)

    assert main([]) == 0 and json.loads(board_json.read_text(encoding="utf-8"))["rows"]
    assert len(capsys.readouterr().out.strip().splitlines()) == 7       # шапка + 6 строк
    assert (drafts / "docx1" / "out.md").exists() and (drafts / "docx1" / "report.json").exists()
    errors = fixtures / "docx1.errors.txt"
    assert errors.read_text(encoding="utf-8").startswith(ERRORS_HEADER)
    assert "_test/board/docx1/out.md" in errors.read_text(encoding="utf-8")
    assert len(list(fixtures.glob("*.errors.txt"))) == 5
    assert not (fixtures / "bakeoff.errors.txt").exists()
    errors.write_text(ERRORS_HEADER + "3 | а | б\n", encoding="utf-8")
    main([])
    assert errors.read_text(encoding="utf-8").endswith("3 | а | б\n")
