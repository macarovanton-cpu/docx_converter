"""ПРАВКА #61: приёмочные тесты MineruProvider.

Сети нет: FakeSession отдаёт заранее заданные ответы и пишет журнал вызовов,
паузы складываются в список вместо time.sleep. Ответы синтетические, по форме
из доки (PLACEHOLDER: записанных настоящих ответов среди фикстур нет).
"""

import difflib
import hashlib
import io
import json
import subprocess
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import pytest
import requests

from ocr_fixtures import FIXTURES, read_fixture, require_fixture, text_tokens
from ocr.mineru_provider import (
    API_BASE,
    MAX_PAGES,
    MineruAuthError,
    MineruError,
    MineruLimitError,
    MineruProvider,
    MineruTimeout,
    result_from_zip,
)

REPO_DIR = Path(__file__).resolve().parents[1]
FAKE_PDF = b"%PDF-1.4 fake"


@dataclass
class Call:
    method: str
    url: str
    json: dict | None = None
    data: bytes | None = None
    headers: dict | None = None


@dataclass
class FakeResponse:
    status_code: int = 200
    payload: dict | None = None
    content: bytes = b""

    def json(self):
        if self.payload is None:
            raise ValueError("не JSON")
        return self.payload


@dataclass
class FakeSession:
    responses: list = field(default_factory=list)
    calls: list = field(default_factory=list)

    def post(self, url, json=None, headers=None):
        return self._next(Call("POST", url, json=json, headers=headers))

    def put(self, url, data=None, headers=None):
        return self._next(Call("PUT", url, data=data, headers=headers))

    def get(self, url, headers=None):
        return self._next(Call("GET", url, headers=headers))

    def _next(self, call):
        self.calls.append(call)
        item = self.responses.pop(0)
        if isinstance(item, Exception):
            raise item
        return item


def api_ok(data):
    return FakeResponse(payload={"code": 0, "data": data, "msg": "ok"})


def api_err(code, msg="ошибка"):
    return FakeResponse(status_code=401, payload={"code": code, "msg": msg})


def extract(state, **extra):
    return api_ok({"extract_result": [dict(state=state, **extra)]})


def zip_bytes(members: dict) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, text in members.items():
            archive.writestr(name, text)
    return buffer.getvalue()


def make_zip(full_md: str, content_list: list | None, *, prefix: str = "") -> bytes:
    members = {f"{prefix}full.md": full_md}
    if content_list is not None:
        members[f"{prefix}content_list.json"] = json.dumps(content_list,
                                                           ensure_ascii=False)
    return zip_bytes(members)


def make_zip_with_member(name: str) -> bytes:
    return zip_bytes({"full.md": "# md", name: "зло"})


def make_zip_without_full_md() -> bytes:
    return zip_bytes({"readme.txt": "нет markdown"})


def fake_pages(count: int, has_text_layer: bool = False):
    def analyze(path):
        return [{"page_number": number, "has_text_layer": has_text_layer}
                for number in range(1, count + 1)]
    return analyze


def make_provider(session, pauses, *, pages=1, **kwargs):
    """Провайдер на фейковой сессии: страницы не читаются из PDF, паузы копятся."""
    kwargs.setdefault("analyze_func", fake_pages(pages))
    return MineruProvider("k-secret", session=session, sleep=pauses.append, **kwargs)


def happy_responses(raw: bytes):
    return [
        api_ok({"batch_id": "b-1", "file_urls": ["https://upload.example/1"]}),
        FakeResponse(),                                             # PUT
        extract("pending"),
        extract("running"),
        extract("done", full_zip_url="https://zip.example/1.zip"),
        FakeResponse(content=raw),
    ]


# --- счастливый путь --------------------------------------------------------

def test_ocr_pdf_happy_path(tmp_path):
    pdf = require_fixture("bakeoff.pdf").read_bytes()
    vlm = read_fixture("vlm.md")
    raw = make_zip(vlm, [{"type": "text", "text": "Утверждаю:", "page_idx": 0}])
    session, pauses = FakeSession(happy_responses(raw)), []

    provider = MineruProvider("k-secret", session=session, sleep=pauses.append,
                              raw_root=tmp_path)
    result = provider.ocr_pdf(pdf)

    assert result.markdown == vlm
    assert (result.provider, result.model_version) == ("mineru", "vlm")
    assert result.content_list[0]["page_idx"] == 0
    assert [p.number for p in result.pages] == list(range(1, 10))
    assert all(p.ocr_applied and not p.has_text_layer for p in result.pages)
    assert (result.raw_dir / "full.md").read_text(encoding="utf-8") == vlm

    body = session.calls[0].json
    assert session.calls[0].url == API_BASE + "/file-urls/batch"
    assert session.calls[0].headers["Authorization"] == "Bearer k-secret"
    assert body["model_version"] == "vlm" and body["language"] == "east_slavic"
    assert body["enable_table"] is True and body["enable_formula"] is True
    assert body["files"][0]["is_ocr"] is True                 # bakeoff.pdf — скан
    assert body["files"][0]["name"] == hashlib.sha256(pdf).hexdigest()[:16] + ".pdf"
    assert body["files"][0]["data_id"] == hashlib.sha256(pdf).hexdigest()[:32]
    assert "bakeoff" not in json.dumps(body)      # имя файла в облако не уходит
    assert "page_ranges" not in body["files"][0]

    assert session.calls[1].method == "PUT" and session.calls[1].data == pdf
    assert "Authorization" not in (session.calls[1].headers or {})
    assert session.calls[2].url == API_BASE + "/extract-results/batch/b-1"
    assert "Authorization" not in (session.calls[5].headers or {})
    assert pauses == [5.0, 5.0]                   # pending, running, done


def test_page_range_and_text_layer_reach_request():
    session, pauses = FakeSession(happy_responses(make_zip("# md", None))), []
    provider = make_provider(session, pauses,
                             analyze_func=fake_pages(6, has_text_layer=True))
    provider.fetch_raw_zip(FAKE_PDF, "2,4-6")

    entry = session.calls[0].json["files"][0]
    assert entry["page_ranges"] == "2,4-6"
    assert entry["is_ocr"] is False               # текстовый слой на всех страницах


def test_result_from_zip_finds_nested_full_md(tmp_path):
    raw = make_zip("# вложенный", [{"page_idx": 1}], prefix="bakeoff/auto/")
    result = result_from_zip(raw, tmp_path / "raw", model_version="pipeline", pages=[])

    assert result.markdown == "# вложенный"
    assert result.content_list == [{"page_idx": 1}]
    assert (result.raw_dir / "bakeoff" / "auto" / "full.md").is_file()


def test_result_from_zip_clears_existing_raw_dir(tmp_path):
    raw_dir = tmp_path / "raw"
    raw_dir.mkdir()
    (raw_dir / "старое.md").write_text("хлам", encoding="utf-8")

    result_from_zip(make_zip("# новое", None), raw_dir, model_version="vlm", pages=[])

    assert not (raw_dir / "старое.md").exists()


def test_result_from_zip_rejects_zip_slip(tmp_path):
    with pytest.raises(MineruError):
        result_from_zip(make_zip_with_member("../evil.txt"), tmp_path / "r",
                        model_version="vlm", pages=[])
    assert not (tmp_path / "evil.txt").exists()
    assert not (tmp_path / "r").exists()          # не распаковано вовсе


def test_result_from_zip_without_content_list(tmp_path):
    no_cl = result_from_zip(make_zip("# md", None), tmp_path / "r2",
                            model_version="vlm", pages=[])
    assert no_cl.content_list is None             # не ошибка


def test_result_from_zip_without_full_md(tmp_path):
    with pytest.raises(MineruError, match="full.md"):
        result_from_zip(make_zip_without_full_md(), tmp_path / "r3",
                        model_version="vlm", pages=[])


# --- ошибки -----------------------------------------------------------------

def test_missing_key_is_constructor_error(monkeypatch):
    monkeypatch.delenv("MINERU_API_KEY", raising=False)
    with pytest.raises(MineruAuthError, match="MINERU_API_KEY"):
        MineruProvider(None)


def test_key_from_env(monkeypatch):
    monkeypatch.setenv("MINERU_API_KEY", "env-key")
    session, pauses = FakeSession(happy_responses(make_zip("# md", None))), []
    MineruProvider(None, session=session, sleep=pauses.append,
                   analyze_func=fake_pages(1)).fetch_raw_zip(FAKE_PDF)
    assert session.calls[0].headers["Authorization"] == "Bearer env-key"


def test_bad_model_version():
    with pytest.raises(ValueError):
        MineruProvider("k", model_version="best")


@pytest.mark.parametrize("code,expected", [
    ("A0202", MineruAuthError),
    ("A0211", MineruAuthError),
    ("-60005", MineruLimitError),
    ("-60006", MineruLimitError),
    ("B1234", MineruError),
])
def test_api_error_codes(code, expected):
    session, pauses = FakeSession([api_err(code, "текст от облака")]), []
    with pytest.raises(expected) as excinfo:
        make_provider(session, pauses).fetch_raw_zip(FAKE_PDF)
    assert code in str(excinfo.value) and "текст от облака" in str(excinfo.value)
    assert pauses == []                           # code != 0 не ретраится


def test_precheck_pages_before_any_request():
    pdf = require_fixture("bakeoff.pdf").read_bytes()
    session = FakeSession([])
    assert MAX_PAGES == 200
    with pytest.raises(MineruLimitError) as excinfo:
        MineruProvider("k", session=session, max_pages=5).fetch_raw_zip(pdf)
    assert "9" in str(excinfo.value) and "5" in str(excinfo.value)
    assert session.calls == []                    # до сети не дошло


def test_precheck_size_before_any_request(monkeypatch):
    monkeypatch.setattr("ocr.mineru_provider.MAX_BYTES", 5)
    session, pauses = FakeSession([]), []
    with pytest.raises(MineruLimitError) as excinfo:
        make_provider(session, pauses).fetch_raw_zip(FAKE_PDF)
    assert str(len(FAKE_PDF)) in str(excinfo.value) and "5" in str(excinfo.value)
    assert session.calls == []


def test_failed_state():
    session, pauses = FakeSession([
        api_ok({"batch_id": "b-1", "file_urls": ["https://upload.example/1"]}),
        FakeResponse(),
        extract("failed", err_msg="boom"),
    ]), []
    with pytest.raises(MineruError, match="boom"):
        make_provider(session, pauses).fetch_raw_zip(FAKE_PDF)


def test_unknown_state_is_error_not_waiting():
    session, pauses = FakeSession([
        api_ok({"batch_id": "b-1", "file_urls": ["https://upload.example/1"]}),
        FakeResponse(),
        extract("teleporting"),
    ]), []
    with pytest.raises(MineruError, match="teleporting"):
        make_provider(session, pauses).fetch_raw_zip(FAKE_PDF)
    assert pauses == []


def test_timeout_names_batch_id():
    session, pauses = FakeSession([
        api_ok({"batch_id": "b-42", "file_urls": ["https://upload.example/1"]}),
        FakeResponse(),
    ] + [extract("running")] * 10), []
    with pytest.raises(MineruTimeout, match="b-42"):
        make_provider(session, pauses, timeout_sec=12).fetch_raw_zip(FAKE_PDF)
    assert pauses == [5.0, 5.0]                   # 5 + 5 + 5 > 12 — дальше не ждём


# --- ретраи -----------------------------------------------------------------

def test_retries_5xx_then_succeeds():
    session, pauses = FakeSession(
        [FakeResponse(status_code=502), FakeResponse(status_code=502)]
        + happy_responses(make_zip("# md", None))), []
    make_provider(session, pauses).fetch_raw_zip(FAKE_PDF)
    assert pauses[:2] == [1.0, 2.0]
    assert pauses == [1.0, 2.0, 5.0, 5.0]


def test_retries_connection_error():
    session, pauses = FakeSession(
        [requests.ConnectionError("нет сети")]
        + happy_responses(make_zip("# md", None))), []
    make_provider(session, pauses).fetch_raw_zip(FAKE_PDF)
    assert pauses[0] == 1.0


def test_retries_exhausted_and_key_not_leaked():
    session, pauses = FakeSession([FakeResponse(status_code=502)] * 4), []
    with pytest.raises(MineruError) as excinfo:
        make_provider(session, pauses).fetch_raw_zip(FAKE_PDF)
    assert pauses == [1.0, 2.0, 4.0]              # ровно три паузы, четыре попытки
    assert len(session.calls) == 4
    assert "k-secret" not in str(excinfo.value)


def test_import_does_not_pull_streamlit():
    script = ("import sys; sys.modules['streamlit'] = None\n"
              "import ocr.mineru_provider as m\n"
              "assert m.API_BASE.startswith('https://')\n")
    result = subprocess.run([sys.executable, "-c", script], cwd=REPO_DIR,
                            capture_output=True, text=True)
    assert result.returncode == 0, result.stderr


# --- живой прогон -----------------------------------------------------------

@pytest.mark.live
def test_live_bakeoff_close_to_vlm(tmp_path):
    pdf = require_fixture("bakeoff.pdf").read_bytes()
    provider = MineruProvider(raw_root=tmp_path)
    raw = provider.fetch_raw_zip(pdf)
    (FIXTURES / "vlm_raw.zip").write_bytes(raw)   # закрывает PLACEHOLDER про content_list
    result = result_from_zip(raw, tmp_path / "raw", model_version="vlm",
                             pages=provider.page_infos(pdf))
    a, b = text_tokens(result.markdown), text_tokens(read_fixture("vlm.md"))
    ratio = difflib.SequenceMatcher(None, a, b, autojunk=False).ratio()
    assert ratio >= 0.95                          # PLACEHOLDER: порог стартовый
    assert result.content_list
