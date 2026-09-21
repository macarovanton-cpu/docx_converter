"""ПРАВКА #87: измерение vision-сверки на остатке эталонов. В тракт не подключено."""

import difflib
import hashlib
import json
import shutil
import sys
import tempfile
import traceback
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from ocr.board import BOARD, FIXTURES, GOLDEN, build_board, count_diffs, text_tokens
from ocr.gemini_verifier import (CONTEXT_TOKENS, QUESTION, GeminiVerifier, VerifierAuthError,
                                 VerifierConfigError, VerifierError, VerifierQuotaError,
                                 crop_block, locate_block)
from pdf_core import Verifier, VerifyResult

MEASURE_SCHEMA_VERSION = 1
MEASURE_JSON = FIXTURES.parents[1] / "verify_measure.json"      # _test/verify_measure.json
CROPS_DIR = FIXTURES.parents[1] / "verify_crops"                # _test/verify_crops/<stem>/<id>.png — для глаз человека
ERROR_OUTCOMES = ("found", "not_found", "wrong_fix")
CONTROL_OUTCOMES = ("agree", "false_alarm", "unreadable")
SERVICE_OUTCOMES = ("unlocated", "no_image", "error")
STOP_ERRORS = (VerifierQuotaError, VerifierAuthError, VerifierConfigError)
NO_RULE = "—"


@dataclass(frozen=True)
class Case:
    id: str                      # "<stem>-e07" / "<stem>-c07"
    fixture: str
    kind: str                    # "error" | "control"
    tag: str | None              # тег опкода difflib; у control — None
    fragment: str                # окно текста тракта, токены через пробел
    expected: str                # то же окно в эталоне; у control == fragment
    rules: tuple[str, ...]       # правила валидатора, чьи находки покрывают опкод; () — ни одно
    page: int | None             # 1-based
    block_type: str | None       # "text" | "table" | …
    ambiguous: bool
    image_png: bytes | None      # None -> исход unlocated / no_image без вызова модели


def _content_list(zip_path: Path) -> list:
    with zipfile.ZipFile(zip_path) as archive:
        names = [n for n in archive.namelist() if n.endswith("content_list.json")]   # _v2 под суффикс не подходит
        if len(names) != 1:
            raise ValueError(f"{zip_path.name}: content_list.json — {names}")
        return json.loads(archive.read(names[0]).decode("utf-8"))


def _squash(text: str) -> str:
    return "".join(text.split())


def _rules(a_side: list[str], findings: list[dict]) -> tuple[str, ...]:
    needle = _squash("".join(a_side))
    if not needle:
        return ()
    hit = set()
    for finding in findings:
        snippet = _squash(finding['snippet'])
        if snippet and (snippet in needle or needle in snippet):
            hit.add(finding['rule'])
    return tuple(sorted(hit))


def build_cases(outputs: dict, *, fixtures: Path = FIXTURES) -> list[Case]:
    cases = []
    for name, zip_name in BOARD:
        md, report = outputs[name]
        golden = (fixtures / GOLDEN[name]).read_text(encoding="utf-8")
        a, b = text_tokens(md), text_tokens(golden)
        opcodes = difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes()
        diffs = [op for op in opcodes if op[0] != "equal"]
        assert len(diffs) == count_diffs(md, golden), (name, len(diffs))    # замер и табло считают одно и то же
        stem = Path(name).stem
        pdf = content_list = None
        crops = {}
        if zip_name is not None:
            pdf = (fixtures / name).read_bytes()
            content_list = _content_list(fixtures / zip_name)

        def place(needle, before, after):
            """(page, block_type, ambiguous, png) или четыре пустышки."""
            if content_list is None:
                return None, None, False, None
            block, ambiguous = locate_block(needle, before, after, content_list)
            if block is None:
                return None, None, False, None
            key = (block['page_idx'], tuple(block['bbox']))     # большинство опкодов сидит в одних и тех же таблицах
            if key not in crops:
                crops[key] = crop_block(pdf, block['page_idx'], block['bbox'])
            return block['page_idx'] + 1, block['type'], ambiguous, crops[key]

        located = []
        for number, (tag, i1, i2, j1, j2) in enumerate(diffs, 1):
            lo, hi = max(0, i1 - CONTEXT_TOKENS), i2 + CONTEXT_TOKENS
            page, block_type, ambiguous, png = place(a[i1:i2], a[lo:i1], a[i2:hi])
            case = Case(id=f"{stem}-e{number:02d}", fixture=name, kind="error", tag=tag,
                        fragment=" ".join(a[lo:hi]), expected=" ".join(a[lo:i1] + b[j1:j2] + a[i2:hi]),
                        rules=_rules(a[i1:i2], report['findings']), page=page, block_type=block_type,
                        ambiguous=ambiguous, image_png=png)
            cases.append(case)
            if png is not None:
                located.append(case)

        # контроль: окна внутри equal-участков, не ближе CONTEXT_TOKENS к любому опкоду; шаг равномерный, без random
        used = set()
        for index, error in enumerate(located):
            length = len(error.fragment.split())
            starts = [s for tag, i1, i2, _, _ in opcodes if tag == "equal"
                      for s in range(i1 + (CONTEXT_TOKENS if i1 else 0),
                                     i2 - (CONTEXT_TOKENS if i2 < len(a) else 0) - length + 1)]
            if not starts:
                continue
            step = len(starts) / len(located)
            for start in starts[int(index * step + step / 2):]:     # не привязался — следующий кандидат
                if used & set(range(start, start + length)):
                    continue
                window = a[start:start + length]
                page, block_type, ambiguous, png = place(window, [], [])
                if png is None:
                    continue
                used.update(range(start, start + length))
                text = " ".join(window)
                cases.append(Case(id=f"{stem}-c{error.id[-2:]}", fixture=name, kind="control", tag=None,
                                  fragment=text, expected=text, rules=(), page=page, block_type=block_type,
                                  ambiguous=ambiguous, image_png=png))
                break
    return cases


def score(case: Case, result: "VerifyResult | None") -> str:
    if case.image_png is None:
        return "no_image" if dict(BOARD)[case.fixture] is None else "unlocated"
    if result is None:
        return "error"
    if case.kind == "control":
        return {"agree": "agree", "fix": "false_alarm", "unreadable": "unreadable"}[result.verdict]
    # PLACEHOLDER: сравнение строгое, по text_tokens; мягкое — только после просмотра wrong_fix человеком
    if result.verdict == "fix":
        return "found" if text_tokens(result.correction) == text_tokens(case.expected) else "wrong_fix"
    return "not_found"


def _count(cases: list[dict], **where) -> int:
    return sum(all(case[key] == value for key, value in where.items()) for case in cases)


def run_measure(cases: list[Case], verifier: Verifier) -> dict:
    rows, complete = [], True
    for case in cases:
        row = {"id": case.id, "fixture": case.fixture, "kind": case.kind, "tag": case.tag, "page": case.page,
               "block_type": case.block_type, "ambiguous": case.ambiguous, "rules": list(case.rules),
               "fragment": case.fragment, "expected": case.expected, "verdict": None, "correction": None,
               "confidence": None, "outcome": None, "cache_hit": False, "error": None}
        rows.append(row)
        if not complete and case.image_png is not None:
            continue                    # после остановки: случай в списке, исхода нет (служебные исходы — есть)
        result = None
        if case.image_png is not None:
            try:
                result = verifier.verify(case.image_png, case.fragment, QUESTION)
            except STOP_ERRORS as exc:  # повтор доберёт остальное: отвеченное уже в кэше
                row["error"] = f"{type(exc).__name__}: {exc}"
                complete = False
                continue
            except VerifierError as exc:
                row["error"] = f"{type(exc).__name__}: {exc}"
            row["cache_hit"] = bool(getattr(verifier, "last_cache_hit", False))
        if result is not None:
            row.update(verdict=result.verdict, correction=result.correction, confidence=result.confidence)
        row["outcome"] = score(case, result)

    errors = [row for row in rows if row["kind"] == "error"]
    controls = [row for row in rows if row["kind"] == "control"]
    measured = [row for row in errors if row["page"] is not None]      # привязан к блоку <=> есть страница
    totals = {"opcodes": len(errors), "measured": len(measured),
              "no_image": _count(errors, outcome="no_image"), "unlocated": _count(errors, outcome="unlocated"),
              "control": len(controls)}
    totals.update({outcome: _count(rows, outcome=outcome)
                   for outcome in ERROR_OUTCOMES + CONTROL_OUTCOMES + ("error",)})

    def error_columns(subset):
        return {outcome: _count(subset, outcome=outcome) for outcome in ERROR_OUTCOMES}

    by_fixture = []
    for name, _ in BOARD:
        own = [row for row in rows if row["fixture"] == name]
        own_errors = [row for row in own if row["kind"] == "error"]
        by_fixture.append({"fixture": name, "opcodes": len(own_errors), **error_columns(own_errors),
                           "unlocated": _count(own_errors, outcome="unlocated"),
                           "no_image": _count(own_errors, outcome="no_image"),
                           "error": _count(own, outcome="error"),
                           "control": _count(own, kind="control"),
                           "false_alarm": _count(own, outcome="false_alarm"),
                           "unreadable": _count(own, outcome="unreadable")})
    rules = sorted({rule for row in measured for rule in row["rules"] or [NO_RULE]})
    by_rule = []
    for rule in rules:
        own = [row for row in measured if rule in (row["rules"] or [NO_RULE])]
        by_rule.append({"rule": rule, "opcodes": len(own), **error_columns(own)})
    by_block_type = []
    for block_type in sorted({row["block_type"] for row in measured + controls}):
        own = [row for row in measured if row["block_type"] == block_type]
        own_controls = [row for row in controls if row["block_type"] == block_type]
        by_block_type.append({"block_type": block_type, "opcodes": len(own), **error_columns(own),
                              "control": len(own_controls),
                              "false_alarm": _count(own_controls, outcome="false_alarm")})
    return {"schema_version": MEASURE_SCHEMA_VERSION,
            "created_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "model": getattr(verifier, "model", type(verifier).__name__),
            "question_sha256": hashlib.sha256(QUESTION.encode()).hexdigest(),
            "complete": complete, "totals": totals, "by_fixture": by_fixture, "by_rule": by_rule,
            "by_block_type": by_block_type, "cases": rows}


def _print_table(rows: list[dict]) -> None:
    if not rows:
        return
    lines = [tuple(rows[0])] + [tuple(str(value) for value in row.values()) for row in rows]
    widths = [max(len(line[i]) for line in lines) for i in range(len(lines[0]))]
    for line in lines:
        print("  ".join(cell.ljust(width) for cell, width in zip(line, widths)).rstrip())
    print()


def main(argv: "list[str] | None" = None) -> int:
    try:
        # глобалы читаются в момент вызова: тесты подменяют пути и GeminiVerifier
        with tempfile.TemporaryDirectory() as work:
            _, outputs = build_board(Path(work), fixtures=FIXTURES)
        cases = build_cases(outputs, fixtures=FIXTURES)
        shutil.rmtree(CROPS_DIR, ignore_errors=True)        # вырезки перезаписываются — до замера: человек смотрит их раньше квоты
        for case in cases:
            if case.image_png is not None:
                path = CROPS_DIR / Path(case.fixture).stem / f"{case.id}.png"
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_bytes(case.image_png)
        verifier = GeminiVerifier()
        measure = run_measure(cases, verifier)
        for table in ("by_fixture", "by_rule", "by_block_type"):
            _print_table(measure[table])
        print("by_rule: опкод с несколькими правилами идёт в строку каждого — сумма строк может превышать measured")
        print("totals:", json.dumps(measure["totals"], ensure_ascii=False))
        print("сетевых вызовов:", getattr(verifier, "network_calls", "-"))
        MEASURE_JSON.parent.mkdir(parents=True, exist_ok=True)
        MEASURE_JSON.write_text(json.dumps(measure, ensure_ascii=False, indent=2), encoding="utf-8")
        if not measure["complete"]:
            stopped = next(row for row in measure["cases"] if row["outcome"] is None and row["error"])
            print(f"замер остановлен на {stopped['id']}: {stopped['error']}", file=sys.stderr)
            return 1
        return 0
    except Exception:
        traceback.print_exc(file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
