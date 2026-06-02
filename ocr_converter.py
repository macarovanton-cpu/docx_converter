"""
Small OCR backend helper for converting image-only PDFs into searchable PDFs.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path
from typing import Any


def ocr_pdf_to_searchable_pdf(
    input_pdf_path: str,
    output_pdf_path: str,
    languages: str = "rus+eng",
    skip_text: bool = True,
    deskew: bool = True,
    rotate_pages: bool = True,
    force_ocr: bool = False,
) -> None:
    """
    Run OCRmyPDF and write a searchable PDF to output_pdf_path.

    OCRmyPDF is executed through the current Python interpreter so the package
    is loaded from the active virtual environment.
    """
    input_path = Path(input_pdf_path)
    output_path = Path(output_pdf_path)

    if not input_path.exists():
        raise FileNotFoundError(f"Input PDF not found: {input_pdf_path}")

    command = [sys.executable, "-m", "ocrmypdf"]
    if force_ocr:
        command.append("--force-ocr")
    elif skip_text:
        command.append("--skip-text")
    if deskew:
        command.append("--deskew")
    if rotate_pages:
        command.append("--rotate-pages")
    command.extend(["-l", languages, str(input_path), str(output_path)])

    try:
        result = subprocess.run(
            command,
            check=True,
            capture_output=True,
            text=True,
            errors="replace",
        )
    except subprocess.CalledProcessError as exc:
        raise RuntimeError(
            _format_subprocess_error("OCRmyPDF failed", exc)
        ) from exc
    except OSError as exc:
        raise RuntimeError(
            f"OCRmyPDF could not be started: {exc}\n"
            f"Command: {_format_command(command)}"
        ) from exc

    if not output_path.exists():
        raise RuntimeError(
            "OCRmyPDF finished successfully, but output PDF was not created.\n"
            f"Command: {_format_command(command)}\n"
            f"stdout:\n{result.stdout or '<empty>'}\n"
            f"stderr:\n{result.stderr or '<empty>'}"
        )


def check_ocr_dependencies() -> dict[str, dict[str, Any]]:
    """
    Check OCR-related command availability and return statuses with versions.

    The function reports every dependency independently and does not raise on
    the first failed check.
    """
    checks = {
        "ocrmypdf": [sys.executable, "-m", "ocrmypdf", "--version"],
        "tesseract": ["tesseract", "--version"],
        "ghostscript": ["gswin64c", "--version"],
    }

    return {
        name: _check_command_version(command)
        for name, command in checks.items()
    }


def _check_command_version(command: list[str]) -> dict[str, Any]:
    try:
        result = subprocess.run(
            command,
            check=True,
            capture_output=True,
            text=True,
            errors="replace",
        )
    except subprocess.CalledProcessError as exc:
        return {
            "ok": False,
            "version": None,
            "error": _format_subprocess_error("Dependency check failed", exc),
            "command": command,
        }
    except OSError as exc:
        return {
            "ok": False,
            "version": None,
            "error": str(exc),
            "command": command,
        }

    output = (result.stdout or result.stderr or "").strip()
    version = output.splitlines()[0] if output else ""
    return {
        "ok": True,
        "version": version,
        "error": None,
        "command": command,
    }


def _format_subprocess_error(message: str,
                            exc: subprocess.CalledProcessError) -> str:
    return (
        f"{message} with exit code {exc.returncode}.\n"
        f"Command: {_format_command(exc.cmd)}\n"
        f"stdout:\n{exc.stdout or '<empty>'}\n"
        f"stderr:\n{exc.stderr or '<empty>'}"
    )


def _format_command(command: Any) -> str:
    if isinstance(command, (list, tuple)):
        return " ".join(str(part) for part in command)
    return str(command)
