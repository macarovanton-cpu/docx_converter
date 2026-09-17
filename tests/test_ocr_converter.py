"""Тесты обёртки OCRmyPDF: таймауты и поиск бинарей (запускать из docx_converter/)."""
import subprocess
from pathlib import Path

import pytest

import ocr_converter
from ocr_converter import DEPENDENCY_TIMEOUT_SEC, OCR_TIMEOUT_SEC


class _FakeCompleted:
    """Минимальный стенд вместо CompletedProcess."""

    stdout = ""
    stderr = ""


def _timeout_run(command, **kwargs):
    """subprocess.run, который всегда упирается в таймаут."""
    raise subprocess.TimeoutExpired(cmd=command, timeout=kwargs["timeout"])


def test_ocr_timeout_raises_readable_error(tmp_path, monkeypatch):
    src = tmp_path / "scan.pdf"
    src.write_bytes(b"%PDF fake")
    monkeypatch.setattr(ocr_converter.subprocess, "run", _timeout_run)

    with pytest.raises(RuntimeError) as excinfo:
        ocr_converter.ocr_pdf_to_searchable_pdf(str(src), str(tmp_path / "out.pdf"))

    message = str(excinfo.value)
    assert str(OCR_TIMEOUT_SEC) in message
    assert "Traceback" not in message
    assert "TimeoutExpired" not in message


def test_ocr_passes_timeout_to_subprocess(tmp_path, monkeypatch):
    seen = {}

    def fake_run(command, **kwargs):
        seen.update(kwargs)
        Path(command[-1]).write_bytes(b"%PDF searchable")
        return _FakeCompleted()

    src = tmp_path / "scan.pdf"
    src.write_bytes(b"%PDF fake")
    monkeypatch.setattr(ocr_converter.subprocess, "run", fake_run)

    ocr_converter.ocr_pdf_to_searchable_pdf(str(src), str(tmp_path / "out.pdf"))

    assert seen["timeout"] == OCR_TIMEOUT_SEC


def test_dependency_timeout_reports_not_ok(monkeypatch):
    monkeypatch.setattr(ocr_converter.subprocess, "run", _timeout_run)

    statuses = ocr_converter.check_ocr_dependencies()

    assert set(statuses) == {"ocrmypdf", "tesseract", "ghostscript"}
    for name, status in statuses.items():
        assert status["ok"] is False, name
        assert str(DEPENDENCY_TIMEOUT_SEC) in status["error"], name


def test_dependency_check_passes_timeout_to_subprocess(monkeypatch):
    seen = []

    def fake_run(command, **kwargs):
        seen.append(kwargs["timeout"])
        return _FakeCompleted()

    monkeypatch.setattr(ocr_converter.subprocess, "run", fake_run)

    ocr_converter.check_ocr_dependencies()

    assert seen == [DEPENDENCY_TIMEOUT_SEC] * 3


def _which_only(*available):
    """shutil.which, который «находит» только перечисленные имена."""
    found = {name: f"/usr/bin/{name}" for name in available}
    asked = []

    def which(name):
        asked.append(name)
        return found.get(name)

    which.asked = asked
    return which


def test_ghostscript_windows_prefers_gswin64c(monkeypatch):
    monkeypatch.setattr(ocr_converter.sys, "platform", "win32")
    monkeypatch.setattr(ocr_converter.shutil, "which", _which_only("gswin64c", "gswin32c"))

    assert ocr_converter._ghostscript_command() == ["/usr/bin/gswin64c", "--version"]


def test_ghostscript_windows_falls_back_to_gswin32c(monkeypatch):
    monkeypatch.setattr(ocr_converter.sys, "platform", "win32")
    monkeypatch.setattr(ocr_converter.shutil, "which", _which_only("gswin32c"))

    assert ocr_converter._ghostscript_command() == ["/usr/bin/gswin32c", "--version"]


def test_ghostscript_linux_uses_gs(monkeypatch):
    which = _which_only("gs")
    monkeypatch.setattr(ocr_converter.sys, "platform", "linux")
    monkeypatch.setattr(ocr_converter.shutil, "which", which)

    assert ocr_converter._ghostscript_command() == ["/usr/bin/gs", "--version"]
    assert which.asked == ["gs"]


def test_ghostscript_missing_keeps_platform_name(monkeypatch):
    monkeypatch.setattr(ocr_converter.sys, "platform", "linux")
    monkeypatch.setattr(ocr_converter.shutil, "which", _which_only())

    assert ocr_converter._ghostscript_command() == ["gs", "--version"]


def test_ghostscript_ok_on_linux_when_gs_present(monkeypatch):
    monkeypatch.setattr(ocr_converter.sys, "platform", "linux")
    monkeypatch.setattr(ocr_converter.shutil, "which", _which_only("gs"))

    def fake_run(command, **kwargs):
        result = _FakeCompleted()
        result.stdout = "10.02.1\n"
        return result

    monkeypatch.setattr(ocr_converter.subprocess, "run", fake_run)

    status = ocr_converter.check_ocr_dependencies()["ghostscript"]

    assert status["ok"] is True
    assert status["version"] == "10.02.1"
    assert status["command"][0] == "/usr/bin/gs"
