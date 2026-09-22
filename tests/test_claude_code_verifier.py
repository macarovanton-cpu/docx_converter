"""ПРАВКА #89: приёмочные тесты ClaudeCodeVerifier. Сети и реального claude нет: subprocess.run — фейковый."""

import json
import socket
import subprocess
from pathlib import Path
from types import SimpleNamespace

import pytest

from ocr import claude_code_verifier as ccv
from ocr.claude_code_verifier import (CLAUDE_MODEL, CLAUDE_TIMEOUT_SEC, FILES_HEADER, ClaudeCodeAuthError,
                                      ClaudeCodeLimitError, ClaudeCodeMissingError, ClaudeCodeTimeoutError,
                                      ClaudeCodeVerifier, claude_command, find_claude, parse_cli_output, parse_texts)
from ocr.gemini_verifier import VerifierConfigError, VerifierError
from ocr.measure import STOP_ERRORS

REPO_ROOT = Path(__file__).resolve().parents[1]
# PLACEHOLDER 3 спеки 14 закрыт для успешного ответа: живой вывод `claude -p --output-format json` человека
# (22.09.2026, записан через конвейер PowerShell: BOM и "ок" в cp866 — "╨╛╨║" — артефакт записи, не claude)
SAMPLE_TEXT = (REPO_ROOT / "tests" / "fixtures" / "claude_p_sample.json").read_text(encoding="utf-8-sig")
SAMPLE = json.loads(SAMPLE_TEXT)
PNG_A, PNG_B, PNG_C = (b"\x89PNG\r\n\x1a\n" + bytes([n]) * 8 for n in (1, 2, 3))
QUESTION = "Перепиши текст каждой картинки."
CACHE_FIELDS = ["key", "model", "image_sha256", "fragment", "question", "raw", "created_at", "file", "batch_files",
                "model_usage"]


@pytest.fixture(autouse=True)
def isolated(monkeypatch):
    def refuse(*args, **kwargs):
        raise AssertionError("тест полез в сеть или запустил процесс")
    monkeypatch.setattr(socket, "socket", refuse)
    monkeypatch.setattr(subprocess, "run", refuse)
    for name in ("CLAUDE_CODE_LIVE", "VERIFIER", "GEMINI_LIVE"):
        monkeypatch.delenv(name, raising=False)


def answer(model=CLAUDE_MODEL, text='{"01.png": "текст"}', **fields):
    """Успешный ответ в форме живого образца: своя модель в modelUsage и свой result."""
    return {**SAMPLE, "result": text, "modelUsage": {model: SAMPLE["modelUsage"]["claude-sonnet-5"]}, **fields}


class FakeRun:
    """subprocess.run: пишет argv, stdin, cwd, env и файлы в cwd на момент вызова; отвечает respond(files, model)."""
    def __init__(self):
        self.calls = []
        self.respond = lambda files, model: answer(model, json.dumps({n: f"текст {n}" for n in files},
                                                                     ensure_ascii=False))

    def __call__(self, argv, *, input, capture_output, text, encoding, cwd, env, timeout):
        assert capture_output and text and encoding == "utf-8"
        files = sorted(p.name for p in Path(cwd).iterdir())
        self.calls.append(SimpleNamespace(argv=argv, input=input, cwd=cwd, env=env, files=files, timeout=timeout))
        out = self.respond(files, argv[argv.index("--model") + 1])
        if isinstance(out, BaseException):
            raise out
        return SimpleNamespace(stdout=json.dumps(out, ensure_ascii=False), stderr="", returncode=0)


@pytest.fixture
def fake(monkeypatch):
    run = FakeRun()
    monkeypatch.setattr(subprocess, "run", run)
    monkeypatch.setattr(ccv, "find_claude", lambda: "claude-fake")
    return run


def test_parse_cli_output():
    out = parse_cli_output(SAMPLE_TEXT, "", "claude-sonnet-5")
    assert out["result"] == SAMPLE["result"] and out["modelUsage"]["claude-sonnet-5"]["provider"] == "firstParty"

    def failed(**fields):
        return json.dumps({**SAMPLE, "is_error": True, **fields})
    with pytest.raises(ClaudeCodeAuthError, match="/login"):
        parse_cli_output(failed(api_error_status=401, result="Invalid bearer token"), "", CLAUDE_MODEL)
    with pytest.raises(ClaudeCodeAuthError):
        parse_cli_output(failed(result="Not logged in · Please run /login"), "", CLAUDE_MODEL)
    with pytest.raises(ClaudeCodeLimitError):
        parse_cli_output(failed(api_error_status=429, result="Too many requests"), "", CLAUDE_MODEL)
    with pytest.raises(ClaudeCodeLimitError):
        parse_cli_output(failed(result="Claude AI usage limit reached|1758520800"), "", CLAUDE_MODEL)
    with pytest.raises(VerifierError, match="error_max_turns") as caught:
        parse_cli_output(failed(subtype="error_max_turns", result=None, errors=["Reached max turns (5)"]),
                         "", CLAUDE_MODEL)
    assert not isinstance(caught.value, STOP_ERRORS)                            # замер идёт дальше
    with pytest.raises(VerifierError, match="stderr 'boom'"):
        parse_cli_output("Error: something", "boom", CLAUDE_MODEL)
    with pytest.raises(VerifierConfigError, match="не та модель"):              # стоп, не тихий fallback
        parse_cli_output(SAMPLE_TEXT, "", "claude-opus-5")
    assert parse_cli_output(json.dumps(answer(extra=1) | {"modelUsage": {
        CLAUDE_MODEL: {}, "claude-haiku-4-5": {}}}), "", CLAUDE_MODEL)                # вспомогательная модель допустима


def test_parse_texts():
    names = ["01.png", "02.png"]
    assert parse_texts('```json\n{"01.png": "а б", "02.png": ""}\n```', names) == {"01.png": "а б", "02.png": ""}
    full = json.dumps({"C:\\Users\\x\\Temp\\claude-verify-1\\01.png": "а", "/tmp/claude-verify-2/02.png": "б"})
    assert parse_texts(full, names) == {"01.png": "а", "02.png": "б"}
    for bad in ('{"01.png": "а"}', '{"01.png": "а", "02.png": "б", "03.png": "в"}', '{"01.png": "а", "02.png": 5}',
                '{"01.png": "а", "C:\\\\x\\\\01.png": "б"}', "нет json", '["01.png", "02.png"]'):
        with pytest.raises(VerifierError):
            parse_texts(bad, names)
    # ПРАВКА #90: модель дописывает после JSON исправленную копию — берётся последний подходящий объект
    two = ('{"01.png": "утрежденные", "02.png": "б"}\n\nОдна оговорка: в 01.png я написал «утрежденные».\n\n'
           '{"01.png": "утвержденные", "02.png": "б"}')
    assert parse_texts(two, names) == {"01.png": "утвержденные", "02.png": "б"}           # последний, а не срез {…}
    assert parse_texts('{"01.png": "а", "02.png": "б"} и ещё {"03.png": "в"}', names) == {"01.png": "а", "02.png": "б"}
    assert parse_texts('{"01.png": "a {b} c", "02.png": ""}', names) == {"01.png": "a {b} c", "02.png": ""}
    with pytest.raises(VerifierError, match="ответ модели не JSON"):
        parse_texts("нет json", names)
    with pytest.raises(VerifierError, match="ответ не по файлам"):
        parse_texts('{"01.png": "а"} {"02.png": "б"}', names)


def test_find_claude(monkeypatch):
    found = {"claude.exe": "C:/claude/claude.exe", "claude.cmd": "C:/npm/claude.cmd", "claude": "/usr/bin/claude"}
    monkeypatch.setattr(ccv.shutil, "which", found.get)
    monkeypatch.setattr(ccv.sys, "platform", "win32")
    assert find_claude() == "C:/claude/claude.exe"
    del found["claude.exe"]
    assert find_claude() == "C:/npm/claude.cmd"
    monkeypatch.setattr(ccv.sys, "platform", "linux")
    assert find_claude() == "/usr/bin/claude"
    found.clear()
    with pytest.raises(ClaudeCodeMissingError, match="только локально"):
        find_claude()


def test_transcribe(fake, tmp_path, monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-ant-secret")
    monkeypatch.setenv("ANTHROPIC_AUTH_TOKEN", "token-secret")
    v = ClaudeCodeVerifier(cache_root=tmp_path, live=True)
    texts = v.transcribe([PNG_A, PNG_B, PNG_A], QUESTION)                 # две уникальные -> один вызов, два файла
    call = fake.calls[0]
    assert len(fake.calls) == 1 and call.files == ["01.png", "02.png"] and len(texts) == 3 and texts[0] == texts[2]
    assert texts == ["текст 01.png", "текст 02.png", "текст 01.png"] and v.last_cache_hits == [False] * 3
    assert call.argv[0] == "claude-fake" and call.argv[1:] == claude_command("x", CLAUDE_MODEL, 2)[1:]
    assert all(a.isascii() and not set(a) & set('\n"&|<>^%') for a in call.argv[1:])   # claude.cmd идёт через cmd.exe
    assert "ANTHROPIC_API_KEY" not in call.env and "ANTHROPIC_AUTH_TOKEN" not in call.env and call.env
    assert not Path(call.cwd).resolve().is_relative_to(REPO_ROOT)          # каталог вне репозитория
    assert not Path(call.cwd).exists()                                     # и убран после вызова
    assert all(str(Path(call.cwd, n)) in call.input for n in ("01.png", "02.png"))
    assert call.input.startswith(f"{QUESTION}\n\n{FILES_HEADER}\n") and call.timeout == CLAUDE_TIMEOUT_SEC
    cached = sorted(tmp_path.iterdir())
    assert len(cached) == 2 and sorted(json.loads(cached[0].read_text("utf-8"))) == sorted(CACHE_FIELDS)
    record = json.loads(cached[0].read_text("utf-8"))
    assert record["batch_files"] == ["01.png", "02.png"] and record["fragment"] == "" and record["model"] == CLAUDE_MODEL
    assert record["model_usage"] == SAMPLE["modelUsage"]["claude-sonnet-5"]
    assert "secret" not in json.dumps(record)
    # calls_log: input_tokens = inputTokens + cacheReadInputTokens + cacheCreationInputTokens образца (2 + 15540 + 17769)
    assert v.network_calls == 1 and v.calls_log == [
        {"images": 2, "input_tokens": 33311, "output_tokens": 4, "duration_ms": 1443, "cost_usd": 0.074228}]

    assert v.transcribe([PNG_B], QUESTION) == [texts[1]] and len(fake.calls) == 1 and v.last_cache_hits == [True]
    assert ClaudeCodeVerifier(cache_root=tmp_path, live=False).transcribe([PNG_A, PNG_B], QUESTION) == texts[:2]
    v.transcribe([PNG_A, PNG_C], QUESTION)
    assert fake.calls[1].files == ["01.png"] and v.last_cache_hits == [True, False]   # уходит только недостающая

    # ключ зависит от модели и вопроса
    opus = ClaudeCodeVerifier(model="claude-opus-5", cache_root=tmp_path, live=True)
    opus.transcribe([PNG_A], QUESTION)
    assert fake.calls[2].argv[fake.calls[2].argv.index("--model") + 1] == "claude-opus-5"
    v.transcribe([PNG_A], QUESTION + " ")
    assert len(fake.calls) == 4 and len(list(tmp_path.iterdir())) == 5

    # cache_root=None: записи живут только в памяти вызова
    memory = ClaudeCodeVerifier(cache_root=None, live=True)
    assert memory.transcribe([PNG_A, PNG_A], QUESTION) == ["текст 01.png"] * 2
    memory.transcribe([PNG_A], QUESTION)
    assert memory.network_calls == 2 and len(list(tmp_path.iterdir())) == 5


def test_miss_without_live(tmp_path, monkeypatch):
    monkeypatch.setattr(ccv, "find_claude", lambda: pytest.fail("find_claude без CLAUDE_CODE_LIVE"))
    v = ClaudeCodeVerifier(cache_root=tmp_path)                           # live=None, окружение пустое
    with pytest.raises(VerifierConfigError, match="CLAUDE_CODE_LIVE=1"):
        v.transcribe([PNG_A], QUESTION)
    monkeypatch.setenv("CLAUDE_CODE_LIVE", "true")                        # только "1"
    with pytest.raises(VerifierConfigError):
        v.transcribe([PNG_A], QUESTION)
    assert v.network_calls == 0 and not list(tmp_path.iterdir())


def test_failures_are_not_cached(fake, tmp_path):
    v = ClaudeCodeVerifier(cache_root=tmp_path, live=True, timeout_sec=5)
    fake.respond = lambda files, model: subprocess.TimeoutExpired(["claude"], 5)
    with pytest.raises(ClaudeCodeTimeoutError, match="5 с") as caught:
        v.transcribe([PNG_A, PNG_B], QUESTION)
    assert not isinstance(caught.value, STOP_ERRORS) and not Path(fake.calls[0].cwd).exists()
    fake.respond = lambda files, model: {**SAMPLE, "is_error": True, "api_error_status": 401, "result": "Invalid"}
    with pytest.raises(ClaudeCodeAuthError) as caught:
        v.transcribe([PNG_A], QUESTION)
    assert isinstance(caught.value, STOP_ERRORS)
    fake.respond = lambda files, model: {**SAMPLE, "is_error": True, "result": "Claude AI usage limit reached|1"}
    with pytest.raises(ClaudeCodeLimitError) as caught:
        v.transcribe([PNG_A], QUESTION)
    assert isinstance(caught.value, STOP_ERRORS)
    fake.respond = lambda files, model: answer("claude-sonnet-4-5")
    with pytest.raises(VerifierConfigError, match="не та модель"):
        v.transcribe([PNG_A], QUESTION)
    assert not list(tmp_path.iterdir()) and v.network_calls == 4 and v.calls_log == []

    # неразобранный успешный ответ: VerifierError, но файл кэша записан (как в #86); повтор — из кэша, без вызова
    fake.respond = lambda files, model: answer(model, "Не могу прочитать картинку.")
    with pytest.raises(VerifierError, match="не JSON"):
        v.transcribe([PNG_A], QUESTION)
    assert len(list(tmp_path.iterdir())) == 1 and len(v.calls_log) == 1
    with pytest.raises(VerifierError, match="не JSON"):
        ClaudeCodeVerifier(cache_root=tmp_path, live=False).transcribe([PNG_A], QUESTION)
    assert len(fake.calls) == 5
