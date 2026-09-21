"""ПРАВКА #86: приёмочные тесты GeminiVerifier, привязки и вырезки. Сети нет вообще."""

import json
import socket
from io import BytesIO
from types import SimpleNamespace

import PIL.Image
import pytest
import requests

from ocr.gemini_verifier import (CROP_MARGIN, GEMINI_MODEL, QUESTION, GeminiVerifier, VerifierAuthError,
                                 VerifierError, VerifierQuotaError, crop_block, locate_block,
                                 parse_answer, verify_cache_key)
from ocr_fixtures import read_raw, require_fixture

PNG = b"\x89PNG\r\n\x1a\nfake"
SECRET = "s3cret-api"
AGREE = '{"verdict":"agree","correction":null,"confidence":0.9}'


@pytest.fixture(autouse=True)
def no_network(monkeypatch):
    def refuse(*args, **kwargs):
        raise AssertionError("тест полез в сеть")
    monkeypatch.setattr(socket, "socket", refuse)
    monkeypatch.delenv("GEMINI_LIVE", raising=False)
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)


class FakeResponse:
    def __init__(self, status_code=200, payload=None):
        self.status_code = status_code
        self._payload = ({"candidates": [{"content": {"parts": [{"text": AGREE}]}}]}
                         if payload is None else payload)
        self.text = json.dumps(self._payload)

    def json(self):
        return self._payload


class FakeSession:
    def __init__(self, *responses):
        self.responses = list(responses)        # пусто -> всегда 200 agree
        self.calls = []

    def post(self, url, *, headers, json, timeout):
        self.calls.append(SimpleNamespace(url=url, headers=headers, json=json))
        response = self.responses.pop(0) if self.responses else FakeResponse()
        if isinstance(response, Exception):
            raise response
        return response


def test_parse_answer():
    assert parse_answer(AGREE).verdict == "agree"
    r = parse_answer('```json\n{"verdict":"fix","correction":"IP-камерами","confidence":0.7}\n```')
    assert (r.verdict, r.correction, r.confidence) == ("fix", "IP-камерами", 0.7)
    assert parse_answer('{"verdict":"agree","correction":"мусор","confidence":1}').correction is None
    for bad in ('не json', '{"verdict":"maybe","confidence":0.5}',
                '{"verdict":"fix","correction":"","confidence":0.5}',
                '{"verdict":"agree","confidence":1.5}', '{"verdict":"agree"}', '{"verdict":"agree","confidence":true}'):
        with pytest.raises(VerifierError):
            parse_answer(bad)


def test_cache_key_depends_on_all_parts():
    base = verify_cache_key(b"png", "ф", "в", "m")
    assert len({base, verify_cache_key(b"png2", "ф", "в", "m"), verify_cache_key(b"png", "ф2", "в", "m"),
                verify_cache_key(b"png", "ф", "в2", "m"), verify_cache_key(b"png", "ф", "в", "m2")}) == 5


def test_network_cache_and_pause(tmp_path):
    fake, sleeps = FakeSession(), []
    v = GeminiVerifier(SECRET, model="m", cache_root=tmp_path, live=True, session=fake,
                       sleep=sleeps.append, clock=lambda: 100.0)
    first = v.verify(PNG, "фрагмент")
    assert not v.last_cache_hit and sleeps == []
    assert fake.calls[0].headers["x-goog-api-key"] == SECRET and SECRET not in fake.calls[0].url
    assert fake.calls[0].url.endswith("/models/m:generateContent")
    parts = fake.calls[0].json["contents"][0]["parts"]
    assert parts[0]["inline_data"]["mime_type"] == "image/png"
    assert parts[1]["text"] == f"{QUESTION}\n\nТекст OCR:\nфрагмент"
    assert v.verify(PNG, "фрагмент") == first and len(fake.calls) == 1 and v.last_cache_hit   # повтор — из кэша
    v.verify(PNG, "другой")
    assert sleeps == [pytest.approx(6.0)] and v.network_calls == 2                           # пауза только перед сетью
    cached = (tmp_path / f"{verify_cache_key(PNG, 'фрагмент', QUESTION, 'm')}.json").read_text("utf-8")
    assert SECRET not in cached
    assert sorted(json.loads(cached)) == sorted(
        ["key", "model", "image_sha256", "fragment", "question", "raw", "created_at"])

    # промах кэша без GEMINI_LIVE — ошибка, не сеть; попадание — без ключа
    with pytest.raises(VerifierError, match="GEMINI_LIVE"):
        GeminiVerifier(model="m", cache_root=tmp_path, live=False).verify(PNG, "нового нет")
    with pytest.raises(VerifierError, match="GEMINI_LIVE"):
        GeminiVerifier("key", model="m", cache_root=tmp_path, session=fake).verify(PNG, "нового нет")  # live=None, env пуст
    assert GeminiVerifier(model="m", cache_root=tmp_path, live=False).verify(PNG, "фрагмент") == first
    with pytest.raises(VerifierAuthError):
        GeminiVerifier(model="m", cache_root=None, live=True, session=fake).verify(PNG, "ф")      # ключа нет
    with pytest.raises(VerifierError, match="модель не выбрана"):
        GeminiVerifier("key", model="PLACEHOLDER", cache_root=tmp_path, live=True, session=fake).verify(PNG, "ф")
    assert len(fake.calls) == 2


def test_model_constant_is_chosen():
    assert GEMINI_MODEL != "PLACEHOLDER"        # выбор человека 21.09.2026; замер без модели не стартует


def test_http_errors_are_not_cached(tmp_path):
    def verifier(*responses):
        sleeps = []
        fake = FakeSession(*responses)
        return GeminiVerifier(SECRET, model="m", cache_root=tmp_path, live=True, session=fake,
                              sleep=sleeps.append, clock=lambda: 0.0), fake, sleeps

    v, fake, sleeps = verifier(FakeResponse(429))
    with pytest.raises(VerifierQuotaError):
        v.verify(PNG, "ф")
    assert len(fake.calls) == 1 and sleeps == []                # без ретраев
    v, fake, _ = verifier(FakeResponse(403))
    with pytest.raises(VerifierAuthError) as info:
        v.verify(PNG, "ф")
    assert SECRET not in str(info.value)
    v, fake, sleeps = verifier(*[FakeResponse(500)] * 4)
    with pytest.raises(VerifierError):
        v.verify(PNG, "ф")
    assert len(fake.calls) == 4 and sleeps == [1.0, 2.0, 4.0]
    v, fake, sleeps = verifier(requests.ConnectionError("обрыв"), FakeResponse())
    assert v.verify(PNG, "обрыв").verdict == "agree" and sleeps == [1.0]
    (tmp_path / f"{verify_cache_key(PNG, 'обрыв', QUESTION, 'm')}.json").unlink()
    v, _, _ = verifier(FakeResponse(400))
    with pytest.raises(VerifierError):
        v.verify(PNG, "ф")
    v, _, _ = verifier(FakeResponse(200, {"promptFeedback": {"blockReason": "OTHER"}}))
    with pytest.raises(VerifierError, match="candidates"):
        v.verify(PNG, "ф")
    v, _, _ = verifier(FakeResponse(200, {"candidates": [{"content": {"parts": [{"text": "не json"}]}}]}))
    with pytest.raises(VerifierError):
        v.verify(PNG, "разбор")
    # неразобранный ответ — успешный HTTP с текстом: он в кэше (правка разбора не требует квоты); ошибки HTTP — нет
    assert [p.name for p in tmp_path.iterdir()] == [f"{verify_cache_key(PNG, 'разбор', QUESTION, 'm')}.json"]


def test_crop_block_on_real_scan():
    pdf = require_fixture("bakeoff.pdf").read_bytes()
    _, content_list = read_raw("vlm_raw.zip")
    block = content_list[0]                                     # «Утверждаю:», bbox в 0–1000
    png = crop_block(pdf, block["page_idx"], block["bbox"])
    assert png.startswith(b"\x89PNG\r\n\x1a\n")
    w, h = PIL.Image.open(BytesIO(png)).size
    # масштаб 0–1000, не пункты
    assert abs(w - (block["bbox"][2] - block["bbox"][0] + 2 * CROP_MARGIN) / 1000 * 595 / 72 * 200) < 8
    assert len(crop_block(pdf, 0, [0, 0, 1000, 1000])) > len(png)       # клэмп по краям страницы
    with pytest.raises(ValueError):
        crop_block(pdf, 99, block["bbox"])
    with pytest.raises(ValueError):
        crop_block(pdf, 0, [500, 500, 400, 400])


def test_locate_block():
    cl = [{"type": "text", "text": "требованиям IP65не ниже", "bbox": [1, 1, 9, 9], "page_idx": 2},
          {"type": "table", "table_body": "<td>для IP-камерами</td>", "bbox": [1, 1, 9, 9], "page_idx": 3}]
    assert locate_block(["IP65не"], ["требованиям"], ["ниже"], cl) == (cl[0], False)
    assert locate_block(["IP-камерами"], ["чужой", "контекст"], [], cl)[0] is cl[1]        # k сузился до 0
    assert locate_block(["нет"], [], [], cl) == (None, False) and locate_block([], [], [], cl) == (None, False)
    assert locate_block(["IP"], [], [], cl) == (cl[0], True)                                # два блока — первый + флаг
    assert locate_block(["ниже"], [], [], [{"type": "text", "text": "не ниже"}]) == (None, False)   # блок без bbox
