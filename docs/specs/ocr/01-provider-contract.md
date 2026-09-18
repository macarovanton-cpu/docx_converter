# 01 — Расширение контракта провайдера

**# ПРАВКА #60.** Зависит от: 00. Независима от 04.

## Цель

Провайдер OCR возвращает не строку, а `OcrResult`: markdown плюс то, что нужно
кэшу, валидатору и отчёту (блоки с номерами страниц, сведения о страницах, имя
движка, папка с сырьём). Поведение UI не меняется ни на байт.

## Трогать

- `pdf_core.py`
- `tests/test_pdf_core.py`

## Не трогать

- `ocr_auto_mode.py` — протокол провайдера он не использует (работает с
  `ocr_func`/`convert_func`), менять нечего. PLAN тут неточен.
- `app.py`, `ocr_converter.py`, `file_converter.py` и всё из общего списка запретов.

## Интерфейсы (дословно)

```python
from dataclasses import dataclass
from pathlib import Path


# ПРАВКА #60: провайдер отдаёт не голый markdown, а результат со сведениями о страницах.
@dataclass(frozen=True)
class PageInfo:
    number: int                      # 1-based
    has_text_layer: bool
    ocr_applied: bool
    warnings: tuple[str, ...] = ()


@dataclass(frozen=True)
class OcrResult:
    markdown: str
    content_list: list | None        # блоки MinerU; None — провайдер блоков не даёт
    pages: list[PageInfo]
    provider: str                    # "ocrmypdf" | "mineru"
    model_version: str | None        # "vlm" | "pipeline" | None
    raw_dir: Path | None             # папка с распакованным сырьём; None — сырья нет


class OcrProvider(Protocol):
    """Интерфейс OCR-провайдера: скан-PDF (bytes) -> OcrResult."""

    def ocr_pdf(self, pdf_bytes: bytes,
                page_range: str | None = None) -> OcrResult: ...
```

Метод переименован `ocr_pdf_to_markdown` → `ocr_pdf`: старое имя с новым типом
возврата врало бы. Старого имени не остаётся (алиасов не заводить — потребителей,
кроме тестов, нет).

```python
class OcrmypdfProvider:
    def __init__(self, ocr_func=ocr_pdf_to_searchable_pdf,
                 convert_func=convert_with_markitdown,
                 analyze_func=analyze_pdf_pages): ...

    def ocr_pdf(self, pdf_bytes: bytes,
                page_range: str | None = None) -> OcrResult: ...
```

`OcrmypdfProvider.ocr_pdf` возвращает:

- `markdown` — как раньше (`convert_func(ocr_path, page_range=page_range)`);
- `content_list=None`, `provider="ocrmypdf"`, `model_version=None`, `raw_dir=None`;
- `pages` — по `analyze_func(src_path)`, отфильтрованным существующей
  `ocr_auto_mode.selected_pdf_pages(pages, page_range)` (переиспользовать, не
  писать заново): `PageInfo(number=p["page_number"], has_text_layer=bool(p["has_text_layer"]),
  ocr_applied=not p["has_text_layer"])` — движок запускается с `--skip-text`.
- Временные файлы чистятся в `finally`, как сейчас.

**Не меняются** (от них зависит `app.py`):

```python
def pdf_to_markdown_with_status(pdf_bytes: bytes, *, page_range: str | None = None,
                                mode: str = "auto",
                                provider: OcrProvider | None = None
                                ) -> tuple[str, dict[str, Any] | None]
def pdf_to_markdown(pdf_bytes: bytes, *, page_range: str | None = None,
                    mode: str = "auto", provider: OcrProvider | None = None) -> str
```

Единственная правка внутри: `markdown = provider.ocr_pdf(pdf_bytes, page_range).markdown`.
Словарь статуса, его ключи и тексты сообщений — без изменений. Ветка
`provider is None` — без изменений.

## Приёмочные тесты (`tests/test_pdf_core.py`)

Существующие тесты адаптируются минимально: `FakeProvider.ocr_pdf_to_markdown` →
`ocr_pdf`, возвращает `OcrResult(markdown="provider markdown", content_list=None,
pages=[], provider="fake", model_version=None, raw_dir=None)`; остальные assert-ы
этих тестов остаются как есть и обязаны проходить.

В `test_ocrmypdf_provider_pipeline_and_temp_cleanup` добавить
`analyze_func=lambda path: [{"page_number": 1, "has_text_layer": False}]`
и заменить проверку результата:

```python
result = provider.ocr_pdf(b"%PDF fake scan", page_range="3")
assert isinstance(result, OcrResult)
assert result.markdown == "ocr markdown"
assert (result.provider, result.model_version) == ("ocrmypdf", None)
assert result.content_list is None and result.raw_dir is None
```

Новый тест на фикстуре (реальный `analyze_pdf_pages`, фейковые OCR и конвертация):

```python
pdf = require_fixture("bakeoff.pdf").read_bytes()
vlm = read_fixture("vlm.md")

def ocr_func(src, dst): Path(dst).write_bytes(b"%PDF searchable")
def convert_func(path, page_range=None): return vlm

result = OcrmypdfProvider(ocr_func=ocr_func, convert_func=convert_func).ocr_pdf(pdf)
assert result.markdown == vlm
assert [p.number for p in result.pages] == list(range(1, 10))
assert all(not p.has_text_layer and p.ocr_applied for p in result.pages)
assert all(p.warnings == () for p in result.pages)

ranged = OcrmypdfProvider(ocr_func=ocr_func, convert_func=convert_func).ocr_pdf(pdf, "2-3")
assert [p.number for p in ranged.pages] == [2, 3]
```

Контракт датаклассов:

```python
with pytest.raises(dataclasses.FrozenInstanceError):
    result.markdown = "x"
assert [f.name for f in dataclasses.fields(OcrResult)] == [
    "markdown", "content_list", "pages", "provider", "model_version", "raw_dir"]
assert [f.name for f in dataclasses.fields(PageInfo)] == [
    "number", "has_text_layer", "ocr_applied", "warnings"]
```

Совместимость с UI — статус не изменился:

```python
md, status = pdf_to_markdown_with_status(SCAN_PDF.read_bytes(), page_range="1-2",
                                         provider=FakeProvider())
assert md == "provider markdown"
assert set(status) == {"mode", "status", "message", "pages_without_text_layer"}
```

## Готово, когда

- `pytest -v` зелёный, включая `tests/test_app_fixes.py` и `tests/test_ocr_auto_mode.py`
  без единой правки в них.
- `grep -rn "ocr_pdf_to_markdown" --include=*.py .` — пусто.
- `git diff --stat`: только `pdf_core.py` и `tests/test_pdf_core.py`.

## Коммит

```
ПРАВКА #60: OcrProvider возвращает OcrResult (markdown + страницы + сырьё)

Датаклассы PageInfo и OcrResult в pdf_core.py, метод протокола ocr_pdf.
OcrmypdfProvider отдаёт сведения о страницах через analyze_pdf_pages.
pdf_to_markdown_with_status и словарь статуса не изменились — app.py не тронут.
```
