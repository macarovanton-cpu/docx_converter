"""ПРАВКА #89: Verifier на Claude Code — подпроцесс `claude -p` под авторизацией Claude Code на этой машине.
Только локально: на Streamlit Cloud бинаря claude нет и не будет. Ключей, токенов и HTTP в этом коде нет."""

import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from ocr.gemini_verifier import (CACHE_ROOT, VerifierAuthError, VerifierConfigError, VerifierError,
                                 VerifierQuotaError, verify_cache_key)

CLAUDE_MODEL = "claude-sonnet-5"         # полное имя, не алиас: ключ кэша не должен уехать вместе с алиасом
CLAUDE_TIMEOUT_SEC = 600                 # PLACEHOLDER: на один вызов (страница, до 6 полос)
SYSTEM_PROMPT = "You transcribe document images verbatim. Follow the user instructions exactly."
FILES_HEADER = "Картинки (прочитай каждую инструментом чтения файлов; ключ ответа — имя файла):"
STRIPPED_ENV = ("ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN")        # только авторизация подписки Claude Code
AUTH_MARKERS = ("/login", "not logged in", "invalid api key")       # PLACEHOLDER: по живому выводу
LIMIT_MARKERS = ("usage limit", "limit reached", "hit your limit")  # PLACEHOLDER: по живому выводу


class ClaudeCodeMissingError(VerifierConfigError): ...   # нет бинаря claude
class ClaudeCodeAuthError(VerifierAuthError): ...        # Claude Code не залогинен
class ClaudeCodeLimitError(VerifierQuotaError): ...      # упёрлись в лимит подписки
class ClaudeCodeTimeoutError(VerifierError): ...         # вызов не уложился в таймаут


def find_claude() -> str:
    candidates = ("claude.exe", "claude.cmd") if sys.platform.startswith("win") else ("claude",)
    for name in candidates:
        found = shutil.which(name)
        if found:
            return found
    raise ClaudeCodeMissingError("не найден claude (Claude Code) в PATH: этот бэкенд работает только локально, "
                                 "где Claude Code установлен и в нём выполнен вход")


def claude_command(binary: str, model: str, n_images: int) -> list[str]:
    # claude.cmd идёт через cmd.exe: в argv только ASCII без перевода строки, кавычек и & | < > ^ %
    return [binary, "-p", "--output-format", "json", "--model", model,
            "--tools", "Read", "--permission-mode", "dontAsk", "--safe-mode", "--no-session-persistence",
            "--max-turns", str(n_images + 3), "--system-prompt", SYSTEM_PROMPT]


def parse_cli_output(stdout: str, stderr: str, model: str) -> dict:
    """Итоговое сообщение claude -p --output-format json (SDKResultMessage); ошибки — по классам остановки."""
    try:
        out = json.loads(stdout)
    except ValueError:
        raise VerifierError(f"claude -p: ответ не JSON: stdout {stdout[:200]!r}, stderr {stderr[:200]!r}") from None
    if out['is_error'] or out['subtype'] != "success":
        text = out.get("result") or "; ".join(out.get("errors") or [])    # поля ветки ошибок необязательны
        status, lowered = out.get("api_error_status"), text.lower()
        if status in (401, 403) or any(marker in lowered for marker in AUTH_MARKERS):
            raise ClaudeCodeAuthError(f"Claude Code не авторизован: {text[:200]} — запустите claude и выполните /login")
        if status == 429 or any(marker in lowered for marker in LIMIT_MARKERS):
            raise ClaudeCodeLimitError(f"лимит подписки Claude Code: {text[:200]}")
        raise VerifierError(f"claude -p: {out['subtype']}: {text[:200]}")
    if model not in out['modelUsage']:      # стоп, а не тихая подмена алиасом или fallback-моделью
        raise VerifierConfigError(f"ответила не та модель: {sorted(out['modelUsage'])}, ждали {model}")
    return out


def parse_texts(raw: str, names: list[str]) -> dict[str, str]:
    try:    # срез от { до } снимает обёртку ```json и даёт dict или ошибку разбора
        answer = json.loads(raw[raw.index("{"):raw.rindex("}") + 1])
    except ValueError as exc:
        raise VerifierError(f"ответ модели не JSON ({exc}): {raw[:200]!r}") from exc
    texts = {key.replace("\\", "/").rsplit("/", 1)[-1]: value for key, value in answer.items()}  # полный путь -> имя
    if sorted(texts) != sorted(names) or len(texts) != len(answer) or not all(isinstance(v, str) for v in texts.values()):
        raise VerifierError(f"ответ не по файлам: ждали {names}, пришло {list(answer)}: {raw[:200]!r}")
    return texts


class ClaudeCodeVerifier:
    name = "claude-code"

    def __init__(self, *, model: str = CLAUDE_MODEL, cache_root: Path | None = CACHE_ROOT,
                 live: bool | None = None, timeout_sec: float = CLAUDE_TIMEOUT_SEC):
        self.model = model
        self._cache_root = None if cache_root is None else Path(cache_root)
        self._live = live
        self._timeout_sec = timeout_sec
        self.last_cache_hits: list[bool] = []    # по картинкам последнего transcribe
        self.network_calls = 0                   # сколько раз запускался claude
        self.calls_log: list[dict] = []          # {"images", "input_tokens", "output_tokens", "duration_ms", "cost_usd"}

    def transcribe(self, images: list[bytes], question: str) -> list[str]:
        head = f"{question}\n\n{FILES_HEADER}"
        # фрагмента в промте нет — отсюда ""; одинаковые картинки дают один ключ и делят ответ
        keys = [verify_cache_key(image, "", head, f"claude-code:{self.model}") for image in images]
        records = {}
        for key in dict.fromkeys(keys):
            path = None if self._cache_root is None else self._cache_root / f"{key}.json"
            if path is not None and path.is_file():
                records[key] = json.loads(path.read_text(encoding="utf-8"))
        hits = [key in records for key in keys]
        misses = {key: image for key, image in zip(keys, images) if key not in records}
        if misses:
            live = os.environ.get("CLAUDE_CODE_LIVE") == "1" if self._live is None else self._live
            if not live:
                raise VerifierConfigError("нет в кэше, вызов выключен: нужен CLAUDE_CODE_LIVE=1")
            records.update(self._call(find_claude(), misses, head))
        self.last_cache_hits = hits
        # разбор — при каждом чтении записи: правка parse_texts не тратит лимит
        return [parse_texts(records[key]['raw'], records[key]['batch_files'])[records[key]['file']] for key in keys]

    def _call(self, binary: str, misses: dict[str, bytes], head: str) -> dict[str, dict]:
        files = [f"{number:02d}.png" for number in range(1, len(misses) + 1)]
        env = {key: value for key, value in os.environ.items() if key.upper() not in STRIPPED_ENV}
        # пустой каталог вне репозитория: выше него нет CLAUDE.md и .claude/settings.json проекта
        with tempfile.TemporaryDirectory(prefix="claude-verify-", ignore_cleanup_errors=True) as work:
            paths = []
            for name, image in zip(files, misses.values()):
                path = Path(work, name)
                path.write_bytes(image)
                paths.append(str(path))
            prompt = head + "\n" + "\n".join(paths)       # русский текст — только через stdin
            try:
                # ponytail: по таймауту subprocess.run убивает cmd.exe, дочерний node claude.cmd может остаться;
                # случится — Popen + taskkill /T
                proc = subprocess.run(claude_command(binary, self.model, len(files)), input=prompt,
                                      capture_output=True, text=True, encoding="utf-8", cwd=work, env=env,
                                      timeout=self._timeout_sec)
            except subprocess.TimeoutExpired as exc:
                raise ClaudeCodeTimeoutError(
                    f"claude -p: нет ответа за {self._timeout_sec} с ({len(files)} картинок)") from exc
            finally:
                self.network_calls += 1
        out = parse_cli_output(proc.stdout, proc.stderr, self.model)     # ошибка — ничего не кэшируется
        usage = out['modelUsage'][self.model]
        self.calls_log.append({
            "images": len(files),
            "input_tokens": usage['inputTokens'] + usage['cacheReadInputTokens'] + usage['cacheCreationInputTokens'],
            "output_tokens": usage['outputTokens'], "duration_ms": out['duration_ms'],
            "cost_usd": out['total_cost_usd']})    # расчётная по прайсу, реально платит подписка
        created = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        records = {}
        for name, (key, image) in zip(files, misses.items()):
            # весь ответ пакета: неразобранный успешный ответ тоже кэшируется, как в #86
            records[key] = {"key": key, "model": self.model, "image_sha256": hashlib.sha256(image).hexdigest(),
                            "fragment": "", "question": head, "raw": out['result'], "created_at": created,
                            "file": name, "batch_files": files, "model_usage": usage}
            if self._cache_root is not None:
                self._cache_root.mkdir(parents=True, exist_ok=True)
                (self._cache_root / f"{key}.json").write_text(
                    json.dumps(records[key], ensure_ascii=False, indent=2), encoding="utf-8")
        return records
