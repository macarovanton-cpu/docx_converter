"""ПРАВКА #86: Verifier на Gemini (REST, без SDK) + вырезки страниц по content_list."""

import base64
import hashlib
import json
import os
import re
import time
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path

import pdfplumber
import requests

from pdf_core import VERDICTS, VerifyResult

GEMINI_MODEL = "gemini-2.5-flash"      # выбор человека (21.09.2026); "PLACEHOLDER" -> verify падает
API_URL = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
MIN_INTERVAL_SEC = 6.0                 # PLACEHOLDER: пауза между сетевыми вызовами (≈10 запросов/мин), не сверено
REQUEST_TIMEOUT_SEC = 120.0
RETRY_PAUSES_SEC = (1.0, 2.0, 4.0)     # как в mineru_provider: 5xx и обрывы связи
CACHE_ROOT = Path(__file__).resolve().parents[1] / ".cache" / "ocr" / "verify"
CROP_RESOLUTION = 200
CROP_MARGIN = 20                       # PLACEHOLDER: поле запаса, в единицах bbox (0–1000)
BBOX_SCALE = 1000
CONTEXT_TOKENS = 3

QUESTION = (
    "На картинке — фрагмент страницы документа на русском языке. Ниже — текст, который OCR распознал\n"
    "в одном месте этой картинки. Найди это место и сравни текст с изображением посимвольно: буквы,\n"
    "цифры, пробелы между словами, знаки, латиница и кириллица. Ничего не улучшай и не перефразируй:\n"
    "важно только то, что напечатано на картинке.\n"
    "Ответь одним JSON-объектом без пояснений:\n"
    '{"verdict": "agree" | "fix" | "unreadable", "correction": "<весь текст OCR в исправленном виде>" | null,\n'
    ' "confidence": <число от 0 до 1>}\n'
    "agree — текст совпадает с картинкой; fix — не совпадает, в correction весь фрагмент, исправленный\n"
    "по картинке; unreadable — место не найдено или не читается."
)

_TAG_RE = re.compile(r"<[^>]+>")
_SPACE_RE = re.compile(r"\s+")


class VerifierError(RuntimeError): ...
class VerifierAuthError(VerifierError): ...     # нет ключа, HTTP 401 / 403
class VerifierQuotaError(VerifierError): ...    # HTTP 429
class VerifierConfigError(VerifierError): ...   # модель не выбрана / сеть выключена: замер останавливается, а не копит error


def verify_cache_key(image_png: bytes, fragment: str, question: str, model: str) -> str:
    return hashlib.sha256(image_png + b"\0" + fragment.encode() + b"\0"
                          + question.encode() + b"\0" + model.encode()).hexdigest()


def parse_answer(raw: str) -> VerifyResult:
    """Терпимый разбор: обрамление ```json снимается срезом от первой { до последней }."""
    try:
        d = json.loads(raw[raw.index("{"):raw.rindex("}") + 1])
        verdict, confidence = d['verdict'], d['confidence']
        if verdict not in VERDICTS:
            raise ValueError(f"verdict {verdict!r}")
        if isinstance(confidence, bool) or not isinstance(confidence, (int, float)) or not 0 <= confidence <= 1:
            raise ValueError(f"confidence {confidence!r}")
        correction = None
        if verdict == "fix":
            correction = d['correction']
            if not isinstance(correction, str) or not correction.strip():
                raise ValueError("fix без correction")
    except (ValueError, KeyError, TypeError) as exc:
        # не разобрали — ошибка, а не unreadable: это разные исходы замера
        raise VerifierError(f"ответ модели не разобран ({exc}): {raw[:200]!r}") from exc
    return VerifyResult(verdict=verdict, correction=correction, confidence=float(confidence), raw=raw)


class GeminiVerifier:
    def __init__(self, api_key: str | None = None, *, model: str = GEMINI_MODEL,
                 cache_root: Path | None = CACHE_ROOT, live: bool | None = None,
                 min_interval_sec: float = MIN_INTERVAL_SEC,
                 session=None, sleep=time.sleep, clock=time.monotonic):
        self._api_key = api_key
        self.model = model
        self._cache_root = None if cache_root is None else Path(cache_root)
        self._live = live
        self._min_interval_sec = min_interval_sec
        self._session = session
        self._sleep = sleep
        self._clock = clock
        self._last_call = None          # время прошлого сетевого вызова
        self.last_cache_hit = False
        self.network_calls = 0          # сколько verify ушло в сеть (для отчёта замера)

    def verify(self, image_png: bytes, fragment: str, question: str = QUESTION) -> VerifyResult:
        self.last_cache_hit = False
        if self.model == "PLACEHOLDER":     # раньше кэша: иначе в кэш лягут ответы под мусорным именем
            raise VerifierConfigError("модель не выбрана: GEMINI_MODEL = PLACEHOLDER")
        key = verify_cache_key(image_png, fragment, question, self.model)
        path = None if self._cache_root is None else self._cache_root / f"{key}.json"
        if path is not None and path.is_file():
            self.last_cache_hit = True
            return parse_answer(json.loads(path.read_text(encoding="utf-8"))['raw'])
        live = os.environ.get("GEMINI_LIVE") == "1" if self._live is None else self._live
        if not live:
            raise VerifierConfigError("нет в кэше, сеть выключена: нужен GEMINI_LIVE=1")
        api_key = self._api_key or os.environ.get("GEMINI_API_KEY")
        if not api_key:
            raise VerifierAuthError("нет ключа: GEMINI_API_KEY")
        if self._last_call is not None:
            pause = self._min_interval_sec - (self._clock() - self._last_call)
            if pause > 0:
                self._sleep(pause)
        try:
            raw = self._post(api_key, image_png, fragment, question)
        finally:
            self._last_call = self._clock()
            self.network_calls += 1
        if path is not None:        # только после успешного ответа с текстом; ошибки не кэшируются
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(json.dumps({
                "key": key, "model": self.model,
                "image_sha256": hashlib.sha256(image_png).hexdigest(),
                "fragment": fragment, "question": question, "raw": raw,
                "created_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            }, ensure_ascii=False, indent=2), encoding="utf-8")
        return parse_answer(raw)

    def _post(self, api_key: str, image_png: bytes, fragment: str, question: str) -> str:
        # PLACEHOLDER: форма REST-запроса сверяется первым живым вызовом (generateContent объявлен устаревшим)
        session = self._session or requests.Session()      # прокси — из окружения, trust_env по умолчанию
        body = {"contents": [{"parts": [
            {"inline_data": {"mime_type": "image/png", "data": base64.b64encode(image_png).decode("ascii")}},
            {"text": f"{question}\n\nТекст OCR:\n{fragment}"}]}],
            "generationConfig": {"temperature": 0}}
        attempts = len(RETRY_PAUSES_SEC) + 1
        for attempt in range(attempts):
            try:
                response = session.post(API_URL.format(model=self.model), json=body,
                                        headers={"x-goog-api-key": api_key, "Content-Type": "application/json"},
                                        timeout=REQUEST_TIMEOUT_SEC)
                if response.status_code < 500:
                    break
                reason = f"HTTP {response.status_code}"
            except (requests.ConnectionError, requests.Timeout) as exc:
                reason = type(exc).__name__       # без str(exc): в нём бывает URL и заголовки
            if attempt == attempts - 1:
                raise VerifierError(f"Gemini: {reason}, попыток {attempts}")
            self._sleep(RETRY_PAUSES_SEC[attempt])
        status, text = response.status_code, response.text[:200]
        if status in (401, 403):
            raise VerifierAuthError(f"Gemini: HTTP {status}: {text}")
        if status == 429:           # без ретраев: квота паузой в секунды не лечится
            raise VerifierQuotaError(f"Gemini: HTTP 429: {text}")
        if status >= 400:
            raise VerifierError(f"Gemini: HTTP {status}: {text}")
        try:
            parts = response.json()['candidates'][0]['content']['parts']
            raw = "".join(part["text"] for part in parts if "text" in part)
        except (ValueError, KeyError, IndexError, TypeError) as exc:
            raise VerifierError(f"Gemini: в ответе нет candidates: {text}") from exc
        if not raw:
            raise VerifierError(f"Gemini: в ответе нет текста: {text}")
        return raw


def _squash(text: str) -> str:
    return _SPACE_RE.sub("", text)


def locate_block(needle_tokens: list[str], context_before: list[str], context_after: list[str],
                 content_list: list) -> tuple[dict | None, bool]:
    """(блок, неоднозначно). Сравнение без пробельных символов; контекст сужается от CONTEXT_TOKENS до 0."""
    texts = [_squash(_TAG_RE.sub(" ", block.get("text") or block.get("table_body") or ""))
             for block in content_list]
    for k in range(CONTEXT_TOKENS, -1, -1):
        before = context_before[-k:] if k else []      # [-0:] — это весь список
        needle = _squash("".join(before + needle_tokens + context_after[:k]))
        if not needle:
            return None, False
        found = [block for block, text in zip(content_list, texts) if needle in text]
        if found:
            block = found[0]
            if "bbox" not in block or "page_idx" not in block:
                return None, False
            return block, len(found) > 1
    return None, False


def crop_block(pdf_bytes: bytes, page_idx: int, bbox: list, *,
               margin: int = CROP_MARGIN, resolution: int = CROP_RESOLUTION) -> bytes:
    """PNG вырезки блока; bbox — в шкале 0–BBOX_SCALE по обеим осям, не в пунктах PDF."""
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        if not 0 <= page_idx < len(pdf.pages):
            raise ValueError(f"page_idx {page_idx} вне документа ({len(pdf.pages)} стр.)")
        page = pdf.pages[page_idx]
        px0, py0, px1, py1 = page.bbox
        x0 = max(px0, px0 + (bbox[0] - margin) * page.width / BBOX_SCALE)
        y0 = max(py0, py0 + (bbox[1] - margin) * page.height / BBOX_SCALE)
        x1 = min(px1, px0 + (bbox[2] + margin) * page.width / BBOX_SCALE)
        y1 = min(py1, py0 + (bbox[3] + margin) * page.height / BBOX_SCALE)
        if x1 <= x0 or y1 <= y0:
            raise ValueError(f"вырожденный прямоугольник: {bbox}")
        buffer = BytesIO()
        page.crop((x0, y0, x1, y1)).to_image(resolution=resolution).original.save(buffer, "PNG")
    return buffer.getvalue()
