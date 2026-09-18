"""ПРАВКА #61: провайдер MinerU cloud API v4 за протоколом pdf_core.OcrProvider."""

import hashlib
import json
import os
import shutil
import tempfile
import time
import zipfile
from io import BytesIO
from pathlib import Path, PurePosixPath

import requests

from file_converter import analyze_pdf_pages
from ocr_auto_mode import selected_pdf_pages
from pdf_core import OcrResult, PageInfo, _unlink_quiet, _write_temp_pdf

API_BASE = "https://mineru.net/api/v4"
MAX_BYTES = 200 * 1024 * 1024
MAX_PAGES = 200                       # по apiManage/docs; переопределяется в конструкторе
RETRY_PAUSES_SEC = (1.0, 2.0, 4.0)
QUOTA_CODES: frozenset[str] = frozenset()   # PLACEHOLDER: код «квота исчерпана» неизвестен

AUTH_CODES = frozenset({"A0202", "A0211"})
LIMIT_CODES = frozenset({"-60005", "-60006"})
WAIT_STATES = frozenset({"waiting-file", "pending", "running", "converting"})
MODEL_VERSIONS = ("vlm", "pipeline")


class MineruError(RuntimeError): ...


class MineruAuthError(MineruError): ...     # A0202 / A0211 / нет ключа


class MineruQuotaError(MineruError): ...


class MineruLimitError(MineruError): ...    # предпроверка и -60005 / -60006


class MineruTimeout(MineruError): ...


class MineruProvider:
    """Облачный OCR MinerU: PDF -> zip -> OcrResult."""

    def __init__(self, api_key: str | None = None, *,
                 model_version: str = "vlm",
                 language: str = "east_slavic",
                 max_pages: int = MAX_PAGES,
                 poll_interval_sec: float = 5.0,
                 timeout_sec: float = 900.0,
                 raw_root: Path | None = None,
                 session=None,
                 sleep=time.sleep,
                 analyze_func=analyze_pdf_pages):
        if model_version not in MODEL_VERSIONS:
            raise ValueError(
                f"model_version={model_version!r}: допустимы {MODEL_VERSIONS}")
        key = api_key or os.environ.get("MINERU_API_KEY")
        if not key:
            raise MineruAuthError(
                "нет ключа MinerU: передайте api_key или задайте MINERU_API_KEY")
        self._api_key = key                  # в логи и тексты исключений не попадает
        self._model_version = model_version
        self._language = language
        self._max_pages = max_pages
        self._poll_interval_sec = poll_interval_sec
        self._timeout_sec = timeout_sec
        self._raw_root = raw_root
        self._session = session if session is not None else requests.Session()
        self._sleep = sleep
        self._analyze_func = analyze_func

    def page_infos(self, pdf_bytes: bytes,
                   page_range: str | None = None) -> list[PageInfo]:
        path = _write_temp_pdf(pdf_bytes)
        try:
            pages = selected_pdf_pages(self._analyze_func(path), page_range)
        finally:
            _unlink_quiet(path)
        return [PageInfo(number=page["page_number"],
                         has_text_layer=bool(page["has_text_layer"]),
                         ocr_applied=not page["has_text_layer"])
                for page in pages]

    def fetch_raw_zip(self, pdf_bytes: bytes,
                      page_range: str | None = None) -> bytes:
        pages = self.page_infos(pdf_bytes, page_range)
        # предпроверка до любого запроса: отказ облака пришёл бы после загрузки 200 МБ
        if len(pdf_bytes) > MAX_BYTES:
            raise MineruLimitError(
                f"PDF {len(pdf_bytes)} байт, предел MinerU {MAX_BYTES} байт")
        if len(pages) > self._max_pages:
            raise MineruLimitError(
                f"страниц {len(pages)}, предел MinerU {self._max_pages}")

        digest = hashlib.sha256(pdf_bytes).hexdigest()
        entry = {
            "name": f"{digest[:16]}.pdf",    # имя файла заказчика в облако не уходит
            "is_ocr": any(not page.has_text_layer for page in pages),
            "data_id": digest[:32],
        }
        if page_range:
            # PLACEHOLDER: место page_ranges (в элементе files или на верхнем
            # уровне) сверить с докой на live-прогоне
            entry["page_ranges"] = page_range
        data = self._json_data(self._send(
            "post", f"{API_BASE}/file-urls/batch",
            json={"model_version": self._model_version,
                  "language": self._language,
                  "enable_table": True,
                  "enable_formula": True,
                  "files": [entry]},
            headers={"Authorization": f"Bearer {self._api_key}",
                     "Content-Type": "application/json"}))

        # загрузка по ссылке: без Authorization и Content-Type, задача ставится сама
        self._expect_ok(self._send("put", data["file_urls"][0], data=pdf_bytes),
                        "загрузка PDF")
        zip_url = self._poll_zip_url(data["batch_id"])
        return self._expect_ok(self._send("get", zip_url), "скачивание архива").content

    def ocr_pdf(self, pdf_bytes: bytes,
                page_range: str | None = None) -> OcrResult:
        zip_bytes = self.fetch_raw_zip(pdf_bytes, page_range)
        root = self._raw_root or Path(tempfile.gettempdir()) / "ocr_raw"
        raw_dir = Path(root) / (
            f"{hashlib.sha256(pdf_bytes).hexdigest()[:16]}-{self._model_version}")
        return result_from_zip(zip_bytes, raw_dir,
                               model_version=self._model_version,
                               pages=self.page_infos(pdf_bytes, page_range))

    # --- HTTP ---------------------------------------------------------------

    def _send(self, method: str, url: str, **kwargs):
        """Запрос с ретраями на 5xx и обрывах связи. 4xx возвращается как есть."""
        attempts = len(RETRY_PAUSES_SEC) + 1
        for attempt in range(attempts):
            try:
                response = getattr(self._session, method)(url, **kwargs)
                if response.status_code < 500:
                    return response
                reason = f"HTTP {response.status_code}"
            except (requests.ConnectionError, requests.Timeout) as exc:
                reason = f"{type(exc).__name__}: {exc}"
            if attempt == attempts - 1:
                raise MineruError(
                    f"MinerU {method.upper()}: {reason}, попыток {attempts}")
            self._sleep(RETRY_PAUSES_SEC[attempt])

    def _json_data(self, response) -> dict:
        """data из ответа API; code != 0 — исключение по таблице кодов."""
        try:
            payload = response.json()
        except ValueError as exc:
            raise MineruError(
                f"MinerU: ответ не JSON (HTTP {response.status_code})") from exc
        if "code" not in payload:
            raise MineruError(f"MinerU: в ответе нет code: {str(payload)[:200]}")
        code = str(payload["code"])
        if code == "0":
            return payload["data"]
        message = f"{code}: {payload['msg'] if 'msg' in payload else ''}"
        if code in AUTH_CODES:
            raise MineruAuthError(message)
        if code in QUOTA_CODES:
            raise MineruQuotaError(message)
        if code in LIMIT_CODES:
            raise MineruLimitError(message)
        raise MineruError(message)

    @staticmethod
    def _expect_ok(response, what: str):
        if response.status_code >= 400:
            raise MineruError(f"MinerU: {what} — HTTP {response.status_code}")
        return response

    def _poll_zip_url(self, batch_id: str) -> str:
        url = f"{API_BASE}/extract-results/batch/{batch_id}"
        headers = {"Authorization": f"Bearer {self._api_key}"}
        waited = 0.0
        while True:
            item = self._json_data(
                self._send("get", url, headers=headers))["extract_result"][0]
            state = item["state"]
            if state == "done":
                return item["full_zip_url"]
            if state == "failed":
                raise MineruError(f"MinerU: распознавание не удалось: {item['err_msg']}")
            if state not in WAIT_STATES:
                raise MineruError(
                    f"MinerU: неизвестное состояние {state!r} (batch {batch_id})")
            if waited + self._poll_interval_sec > self._timeout_sec:
                raise MineruTimeout(
                    f"MinerU: batch {batch_id} не завершился за {self._timeout_sec} с "
                    f"(последнее состояние {state!r})")
            self._sleep(self._poll_interval_sec)
            waited += self._poll_interval_sec


def result_from_zip(zip_bytes: bytes, raw_dir: Path, *,
                    model_version: str | None,
                    pages: list[PageInfo]) -> OcrResult:
    """Сырой zip MinerU -> OcrResult. Сети не касается: тем же живёт кэш (спека 03)."""
    raw_dir = Path(raw_dir)
    with zipfile.ZipFile(BytesIO(zip_bytes)) as archive:
        names = archive.namelist()
        unsafe = [name for name in names if _is_unsafe_member(name)]
        if unsafe:
            raise MineruError(f"MinerU: небезопасные имена в архиве: {unsafe}")

        if raw_dir.exists():
            shutil.rmtree(raw_dir)
        raw_dir.mkdir(parents=True)
        archive.extractall(raw_dir)

        markdown_names = [n for n in names if PurePosixPath(n).name == "full.md"]
        if not markdown_names:
            raise MineruError(f"MinerU: в архиве нет full.md, есть {names}")
        markdown = archive.read(markdown_names[0]).decode("utf-8")

        content_names = [n for n in names if n.endswith("content_list.json")]
        content_list = (json.loads(archive.read(content_names[0]).decode("utf-8"))
                        if content_names else None)

    return OcrResult(markdown=markdown, content_list=content_list, pages=pages,
                     provider="mineru", model_version=model_version, raw_dir=raw_dir)


def _is_unsafe_member(name: str) -> bool:
    """Абсолютный путь, диск Windows или '..' в имени члена архива (zip-slip)."""
    parts = name.replace("\\", "/").split("/")
    return name.startswith("/") or ".." in parts or (len(name) > 1 and name[1] == ":")
