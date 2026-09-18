# 07 — CLI для агента

**# ПРАВКА #66** (+ #67, #70). Зависит от: 01–06. Последняя спека автономного прогона.

## Цель

Одна точка подключения для Hermes Agent и человека: PDF на входе — `out.md`,
`report.json` и одна строка JSON в stdout на выходе. Вся сборка тракта живёт в
функции `run_pipeline`, которую на этапе 8 без изменений вызовет UI.

## Трогать

- `ocr/cli.py` — создать
- `tests/test_ocr_cli.py` — создать
- `CLAUDE.md`, `README.md` — описание инструмента и список модулей (см. ниже)

## Не трогать

`app.py`, все модули `ocr/*` предыдущих спек, `pdf_core.py`, всё из общего списка запретов.

## Интерфейсы (дословно)

```
python -m ocr.cli INPUT.pdf --out DIR
                  [--engine mineru|ocrmypdf]   (по умолчанию mineru)
                  [--mode vlm|pipeline]        (по умолчанию vlm; для ocrmypdf игнорируется)
                  [--verify] [--annotate] [--annotate-all]
                  [--cache local|drive]        (по умолчанию local)
```

Запускать из корня репозитория: `pdf_core` лежит в корне, локальный кэш —
`.cache/ocr/` относительно текущей папки.

```python
"""ПРАВКА #66: CLI OCR-тракта: PDF -> out.md + report.json + JSON-итог в stdout."""

EXIT_OK, EXIT_ERROR, EXIT_CRITICAL = 0, 1, 2


def run_pipeline(pdf_bytes: bytes, *, source_name: str, work_dir: Path,
                 engine: str = "mineru", mode: str = "vlm",
                 verify: bool = False, annotate: bool = False,
                 annotate_all: bool = False,             # ПРАВКА #70
                 cache: CacheBackend | None = None,
                 provider_factory=None) -> tuple[str, dict]: ...


def main(argv: list[str] | None = None, *, provider_factory=None) -> int: ...


if __name__ == "__main__":
    raise SystemExit(main())
```

`provider_factory(model_version: str)` → объект с `fetch_raw_zip` и `page_infos`
(как у `MineruProvider`). `None` → `lambda mv: MineruProvider(model_version=mv)`.
Фабрика, а не готовый провайдер: `--verify` нужен второй, с другим `model_version`.

### `run_pipeline`, engine `mineru`

Для `model_version` (основной — `mode`; при `verify` ещё и второй: для `vlm` это
`pipeline`, для `pipeline` — `vlm`):

1. `key = cache_key(pdf_bytes, "mineru", model_version)`.
2. `cache` задан и `cache.get(key)` не `None` → zip из кэша, **провайдер не
   создаётся и сеть не трогается**. Иначе `provider_factory(mv).fetch_raw_zip(pdf_bytes)`,
   затем (если `cache` задан) `cache.put(key, zip, build_meta(...))`.
3. `result_from_zip(zip, work_dir / "raw" / model_version, model_version=mv, pages=…)`.
   `pages` при попадании в кэш — `[]` (провайдера нет; в отчёт страницы идут из
   `content_list`).

Дальше:

4. `md, f_post = postprocess(result.markdown)`; `f_val = validate(md, result.content_list)`.
5. `verify`: второй прогон → `postprocess(...)[0]` → `f_diff = diff_findings(md, md2, result.content_list)`.
6. `report = build_report(source=source_name, sha256=…, provider="mineru",
   model_version=mode, cache_hit=<попадание основного прогона>, verified=verify,
   findings=f_post + f_val + f_diff, content_list=result.content_list)`.
7. `annotate` или `annotate_all` → `md = annotate(md, report,
   include_low_confidence=annotate_all)`. Возврат `(md, report)`. `--annotate-all`
   включает пометки сам: отдельно писать `--annotate` не нужно (ПРАВКА #70).

### engine `ocrmypdf`

`OcrmypdfProvider().ocr_pdf(pdf_bytes)`; кэш не используется (zip нет),
`provider="ocrmypdf"`, `model_version=None`, `cache_hit=False`. `verify=True` →
`ValueError("--verify доступен только для --engine mineru")`. Шаги 4, 6, 7 — те же.

Неизвестные `engine` / `mode` → `ValueError`.

### `main`

- Пишет `{out}/out.md` и `{out}/report.json` (`utf-8`, `ensure_ascii=False`, `indent=2`);
  папку создаёт. `work_dir = out`. Кэш — `make_cache(args.cache)`.
- В **stdout ровно одна строка** JSON, всё остальное (прогресс, предупреждения) — в stderr:

```json
{"status": "findings", "out_md": "/abs/out/out.md", "report": "/abs/out/report.json",
 "cache_hit": false, "findings": {"critical": 5, "warning": 4, "info": 1}, "error": null}
```

| Итог | `status` | Код |
|---|---|---|
| нет находок `critical` | `"ok"` | 0 |
| есть хотя бы одна `critical` | `"findings"` | 2 |
| любое исключение, нет файла, плохие аргументы | `"error"` | 1 |

- При ошибке: `out_md`, `report` — `null`, `findings` — `null`, `error` — текст
  исключения (у `MineruError` он уже человекочитаемый); traceback — в stderr.
  Файлы при ошибке не пишутся (недописанный `out.md` хуже отсутствующего).
- **argparse по умолчанию выходит с кодом 2** — это совпало бы с «есть критичные
  находки». `ArgumentParser.error` переопределить: ошибка аргументов → тот же
  JSON со `status="error"`, `main` **возвращает 1**, `SystemExit` наружу не летит.
- `--cache drive` → `NotImplementedError` из `make_cache` → `status="error"`, код 1,
  текст про этап 8 доходит до пользователя как есть.
- Ключ API в stdout/stderr не попадает.

### Документация

`CLAUDE.md`:
- «Seven Python modules» → по факту, добавить пакет `ocr/` (шесть модулей, по строке на каждый);
- в «Numbered edits convention» обновить перечень: `#60` в `pdf_core.py`, `#61–#66` в `ocr/`;
- новый раздел `## OCR CLI (инструмент для агента)` — текст ниже дословно.

`README.md`: тот же раздел + строки в дереве файлов.

```markdown
## OCR CLI (инструмент для агента)

**Назначение:** тендерный PDF (в т.ч. скан) → Markdown со структурой + отчёт о сомнительных местах.
Смысл исходника не меняется: чинятся только известные артефакты OCR, остальное помечается.

**Вызов (из корня репозитория):**
`python -m ocr.cli ВХОД.pdf --out ПАПКА [--engine mineru|ocrmypdf] [--mode vlm|pipeline] [--verify] [--annotate] [--annotate-all] [--cache local|drive]`

**Нужно:** переменная окружения `MINERU_API_KEY` (для `--engine mineru`). PDF до 200 МБ и 200 страниц.
Документ уходит в облако mineru.net; повторный прогон того же файла берётся из `.cache/ocr/`.

**Выход:** `ПАПКА/out.md`, `ПАПКА/report.json`, в stdout — одна строка JSON:
`{"status","out_md","report","cache_hit","findings":{"critical","warning","info"},"error"}`.

**Коды выхода:** `0` — критичных находок нет; `2` — есть критичные находки, `out.md` написан,
читать `report.json`; `1` — тракт не отработал, читать `error`.

**Флаги:** `--verify` — второй прогон другим движком MinerU, расхождения → находки `low_confidence`
(дольше, вдвое больше квоты). `--annotate` — находки вставлены в `out.md` как `!! ПРОВЕРИТЬ: … !!`;
перед конвертацией в DOCX снять через `ocr.validate.strip_annotations`. `low_confidence` в текст
не вставляются (их десятки, в `report.json` они есть все) — для полной картины `--annotate-all`.
Находки, чей фрагмент в тексте не нашёлся, уходят в конец файла, в раздел «Не привязанные находки».
`--cache drive` пока не реализован.

**report.json:** `findings[]` = `{id, rule, severity, page, snippet, suggestion}`.
`snippet` — дословный фрагмент `out.md`; `suggestion` — что предлагает тракт или что увидел второй прогон;
`page` — страница PDF или `null`.
```

## Приёмочные тесты (`tests/test_ocr_cli.py`)

Сети нет. `make_zip(full_md)` собирает zip в памяти; фейковая фабрика пишет журнал:

```python
pdf = require_fixture("bakeoff.pdf").read_bytes()
vlm, pipeline, golden = (read_fixture(n) for n in ("vlm.md", "pipeline.md", "golden.md"))
calls = []

class FakeProvider:
    texts = {"vlm": vlm, "pipeline": pipeline}
    def __init__(self, mv): self.mv = mv
    def page_infos(self, pdf_bytes, page_range=None):
        return [PageInfo(n, False, True) for n in range(1, 10)]
    def fetch_raw_zip(self, pdf_bytes, page_range=None):
        calls.append(self.mv)
        return make_zip(self.texts[self.mv])

cache = LocalCache(tmp_path / "cache")
kw = dict(source_name="bakeoff.pdf", work_dir=tmp_path / "w", cache=cache, provider_factory=FakeProvider)

# первый прогон: сеть, второй: кэш
md, report = run_pipeline(pdf, **kw)
assert calls == ["vlm"] and report["cache_hit"] is False
md2, report2 = run_pipeline(pdf, **kw)
assert calls == ["vlm"] and report2["cache_hit"] is True          # провайдер не вызван
assert md2 == md == postprocess(vlm)[0]
assert report2["findings"] == report["findings"]                  # кэш не меняет результат
assert (report["provider"], report["model_version"], report["verified"]) == ("mineru", "vlm", False)
assert report["sha256"] == hashlib.sha256(pdf).hexdigest()
assert report["summary"]["critical"] == 5
assert (tmp_path / "w" / "raw" / "vlm" / "full.md").exists()

# результат годится конвертеру без правок convert.py
assert "<table" not in md and "!! ПРОВЕРИТЬ" not in md

# --verify: второй движок, свой ключ кэша
md_v, report_v = run_pipeline(pdf, verify=True, **kw)
assert calls == ["vlm", "pipeline"] and report_v["verified"] is True
rules = [f["rule"] for f in report_v["findings"]]
assert "low_confidence" in rules
assert any("сыручими" in f["snippet"] for f in report_v["findings"])
assert any("IR-камерами" in f["snippet"] for f in report_v["findings"])
assert md_v == md                                                 # сверка текст не меняет
run_pipeline(pdf, verify=True, **kw)
assert calls == ["vlm", "pipeline"]                               # оба уже в кэше

# --annotate
md_a, report_a = run_pipeline(pdf, annotate=True, **kw)
assert strip_annotations(md_a) == md
assert md_a.count("!! ПРОВЕРИТЬ: ") == len(report_a["findings"])

# --annotate-all: low_confidence в текст только по нему (ПРАВКА #70)
md_v, report_v = run_pipeline(pdf, verify=True, annotate=True, **kw)
md_all, _ = run_pipeline(pdf, verify=True, annotate_all=True, **kw)
low = [f for f in report_v["findings"] if f["rule"] == "low_confidence"]
assert low and md_v.count("!! ПРОВЕРИТЬ: ") == len(report_v["findings"]) - len(low)
assert md_all.count("!! ПРОВЕРИТЬ: ") == len(report_v["findings"])

# без кэша
calls.clear()
run_pipeline(pdf, source_name="b.pdf", work_dir=tmp_path / "n", cache=None, provider_factory=FakeProvider)
run_pipeline(pdf, source_name="b.pdf", work_dir=tmp_path / "n", cache=None, provider_factory=FakeProvider)
assert calls == ["vlm", "vlm"]

# аргументы
with pytest.raises(ValueError, match="verify"):
    run_pipeline(pdf, source_name="b.pdf", work_dir=tmp_path, engine="ocrmypdf", verify=True)
with pytest.raises(ValueError):
    run_pipeline(pdf, source_name="b.pdf", work_dir=tmp_path, engine="tesseract")
```

`main` (через `capsys`, `monkeypatch.chdir(tmp_path)` — чтобы `.cache/` не упал в репозиторий):

```python
src = require_fixture("bakeoff.pdf")
out = tmp_path / "out"

# критичные находки → 2
code = main([str(src), "--out", str(out)], provider_factory=FakeProvider)
stdout = capsys.readouterr().out
assert code == 2 and stdout.count("\n") == 1
summary = json.loads(stdout)
assert list(summary) == ["status", "out_md", "report", "cache_hit", "findings", "error"]
assert summary["status"] == "findings" and summary["error"] is None
assert summary["findings"]["critical"] == 5
assert Path(summary["out_md"]).is_absolute()
assert Path(summary["out_md"]).read_text(encoding="utf-8") == postprocess(vlm)[0]
on_disk = json.loads(Path(summary["report"]).read_text(encoding="utf-8"))
assert on_disk["schema_version"] == 1 and on_disk["source"] == "bakeoff.pdf"
assert "Сиражитдинов" in Path(summary["report"]).read_text(encoding="utf-8")   # не \uXXXX

# чистый документ → 0 (эталон: латиницы в ФИО нет, остались только warning/info).
# Отдельный тест со своим tmp_path: иначе основной прогон честно возьмётся из кэша шага выше.
FakeProvider.texts = {"vlm": golden, "pipeline": pipeline}
code = main([str(src), "--out", str(tmp_path / "g")], provider_factory=FakeProvider)
payload = json.loads(capsys.readouterr().out)
assert code == 0 and payload["status"] == "ok" and payload["findings"]["critical"] == 0

# ошибки → 1, всегда JSON, файлов нет
errors = {}
for name, argv in {
        "no_file":  [str(tmp_path / "нет.pdf"), "--out", str(tmp_path / "e1")],
        "drive":    [str(src), "--out", str(tmp_path / "e2"), "--cache", "drive"],
        "verify":   [str(src), "--out", str(tmp_path / "e3"), "--engine", "ocrmypdf", "--verify"],
        "no_out":   [str(src)],                                    # ошибка argparse
        "bad_mode": [str(src), "--out", str(tmp_path / "e5"), "--mode", "best"]}.items():
    code = main(argv, provider_factory=FakeProvider)               # SystemExit наружу не летит
    payload = json.loads(capsys.readouterr().out)
    assert code == 1 and payload["status"] == "error" and payload["error"]
    assert payload["out_md"] is None and payload["findings"] is None
    errors[name] = payload["error"]
assert "этапе 8" in errors["drive"]
assert not (tmp_path / "e2" / "out.md").exists()

# ключ не утекает
monkeypatch.setenv("MINERU_API_KEY", "k-secret-123")
main([str(src), "--out", str(tmp_path / "s")], provider_factory=FakeProvider)
captured = capsys.readouterr()
assert "k-secret-123" not in captured.out + captured.err
```

Сквозной прогон `out.md` → DOCX, без правок `convert.py`:

```python
convert_md_to_docx(md, str(tmp_path / "t.docx"), template_path=None)
doc = docx.Document(str(tmp_path / "t.docx"))
assert max(len(t.rows) for t in doc.tables) >= 50                 # таблица ТЗ доехала целиком
```

Живой (ПРАВКА #67 — метрика относительная):

```python
@pytest.mark.live
def test_live_cli_bakeoff(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)
    out_dir = tmp_path / "o"
    code = main([str(require_fixture("bakeoff.pdf")), "--out", str(out_dir), "--verify"])
    payload = json.loads(capsys.readouterr().out)
    assert code in (0, 2) and payload["status"] in ("ok", "findings")

    out = Path(payload["out_md"]).read_text(encoding="utf-8")
    golden = read_fixture("golden.md")
    raw = (out_dir / "raw" / "vlm" / "full.md").read_text(encoding="utf-8")
    assert count_diffs(out, golden) < count_diffs(raw, golden)
    assert not re.search(r"\bNo\.?[ \t]*(?=\d|п/п)", out)
    assert "$C^" not in out
    assert "РоЕ" not in out                       # кириллические Р, о, Е
    assert len(parse_pipe_tables(out)) == 1
```

Почему не абсолютный порог. `2 * VLM_TO_GOLDEN_DIFFS` мерил не работу тракта, а
совпадение свежего прогона vlm с тем прогоном, из которого собран `golden.md`.
Со вторым живым прогоном (18.09) он дал 29 при пороге 24 — без регрессии тракта:
24 из 29 опкодов — пробелы, потерянные на переносах внутри ячеек, остальные 5 — известный
остаток спек 04–05 («сыручими», «IR-камерами», транслит в подписях). Сырой markdown
берётся из `<out>/raw/vlm/full.md` — туда `result_from_zip` распаковывает ответ провайдера.
Четыре последние проверки от прогона не зависят: это ровно то, что постпроцессор обязан
починить по спеке 04 (правила 1, 3, 4, 5).

## Готово, когда

- `pytest -v` зелёный, фикстурные тесты не пропущены. Живые тесты (ПРАВКА #71)
  идут только при `MINERU_LIVE=1`; голый `pytest` в сеть не ходит, даже когда
  `MINERU_API_KEY` выставлен. Прогон без сети — `pytest -m "not live"`.
- Из корня репозитория `python -m ocr.cli --help` печатает справку и выходит с 0,
  не импортируя `streamlit`.
- Ручная проверка на Windows и на Linux/WSL2 (у кого есть ключ): команда из
  раздела «Документация» на `bakeoff.pdf`, в отчёте исполнителя — stdout, код
  выхода и время обоих запусков (второй — из кэша). Нет ключа или второй
  платформы — так и написать в отчёте, не выдавать непроверенное за проверенное.
- `git diff --stat`: `ocr/cli.py`, `tests/test_ocr_cli.py`, `CLAUDE.md`, `README.md`.
  `app.py` не изменён. Правка #67 трогает только `tests/test_ocr_cli.py` и эту спеку.

## Коммит

```
ПРАВКА #66: CLI OCR-тракта для агента — python -m ocr.cli

run_pipeline: кэш → MinerU → постпроцессор → валидатор → (--verify) сверка →
report.json; main: out.md + report.json + одна строка JSON в stdout, коды 0/2/1
(ошибка аргументов — 1, не argparse-овская 2). CLAUDE.md и README.md: описание
инструмента для агента, пакет ocr/ в списке модулей.
```
