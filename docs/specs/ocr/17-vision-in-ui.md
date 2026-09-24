# 17 — Сверка по картинке в UI

**# ПРАВКА #92.** Зависит от: 16 (`ingest(..., vision=…, vision_progress=…)`, `report.json` v2,
`ocr.vision.VISION_MODEL`), 12 (режим MinerU на `ingest`). `app.py` **разрешён** этой спекой. Сеть не нужна,
`claude` в тестах не запускается.

## Цель

Сверка по картинке (спека 16) доступна из режима «Файлы → Markdown» → MinerU.

Что появляется в UI:
- одна галочка;
- прогресс по страницам в `st.status`;
- блок подсказок в карточке результата: страница, было, по скану, стало.

Работает только там, где на машине есть `claude`. На Streamlit Cloud галочки нет, и больше ничего не
меняется.

**Стиль UI не меняется**: те же виджеты (`st.checkbox`, `st.status`, `st.expander`, `st.dataframe`,
`st.caption`), та же раскладка, новых стилей и CSS нет.

### Что изменится для человека (записать в docs_sync дословно)

1. В режиме MinerU, между «Сверка вторым прогоном» и «Пометки в тексте», — галочка «Сверка по картинке
   (Claude Code, локально)».
   - Есть только там, где установлен `claude`. На Streamlit Cloud её нет.
   - Модель — `claude-opus-5`, выбора в интерфейсе нет.
2. С галочкой сканы и текстовые PDF с таблицами после распознавания сверяются по картинке. В статусе видно
   «Сверка по картинке: стр. N (k из M)…». Одна страница — один вызов подписки Claude Code; повтор того же
   файла берётся из кэша.
3. В карточке результата — блок «Сверка по картинке — подсказки (N)»: страница, было, по скану, стало.
   - Текст подсказки не меняют.
   - С галочкой «Пометки в тексте» подсказки вставляются в Markdown как «!! ПРОВЕРИТЬ: … !!», как прочие
     находки.
4. Если сверка не удалась (нет входа в Claude Code, лимит, таймаут), конвертация всё равно готова. Причина —
   находка `vision_skipped` в общем списке находок.
5. DOCX/XLSX и текстовые PDF без таблиц по картинке не сверяются. При галочке — подпись «Сверка по картинке
   не применялась: файл не шёл через MinerU.»

## Решения человека (22.09.2026, не пересматриваются)

- Галочка «Сверка по картинке (Claude Code, локально)» — только в режиме MinerU и только при найденном
  `claude`.
- Модель — Opus (`VISION_MODEL`), выбора в UI нет.
- Прогресс — через `st.status` по страницам.
- Пометки в тексте — через существующую галочку «Пометки в тексте», своей нет.
- Тесты — с фейковым `Verifier` или фейковой сверкой, без сети.
- Стиль UI не менять.

## Шаг 0 — условия входа (иначе стоп)

1. ПРАВКА #91 закоммичена; `grep -rn "ПРАВКА #92" --include=*.py .` пуст; `pytest -q` зелёный.
2. Сигнатура `ingest` — как в спеке 16: `vision: str | None = None, vision_progress=None`; в `ocr/vision.py`
   есть `VISION_MODEL`; `report["schema_version"] == 2`. Расхождение — стоп.
3. `grep -n "ingest(" app.py` — ровно один вызов, в `_convert_uploaded_file`. `grep -n "st.status" app.py` —
   один, в цикле конвертации. Другие — стоп: спека писалась по коду, где они одни.

## Трогать

- `app.py` (#92):
  - импорты;
  - `_claude_available`;
  - `_vision_progress`;
  - `_vision_rows`;
  - `_render_ocr_report`;
  - `_convert_uploaded_file`;
  - галочка и передача `vision` в `render_files_to_markdown_mode`;
  - подпись «не применялась» в карточке.
- `tests/test_app_fixes.py` — новые тесты блока #92.
- `CLAUDE.md`:
  - строка `app.py` в «Architecture» — «+ сверка по картинке (#92, только при найденном `claude`)»;
  - «Known production limitation» — фраза «The Claude Code verification backend (#89) is local only…»
    дополняется: UI предлагает сверку, только если `claude` найден (#92);
  - нумерация — `#73, #85 и #92 — app.py`.
- `docs/specs/ocr/README.md` — строка спеки 17.
- `docs/PLAN_OCR.md` — «Этап B, шаг 5 (#92)».
- `docs/docx_converter_docs_sync.md` — `### ПРАВКА #92` («Что изменится для человека» дословно) и снимок
  «следующий — #93».

## Не трогать

- Весь `ocr/` (спека 16 закрыта), `pdf_core.py`, `convert.py`, `file_converter.py`, `requirements.txt`,
  `conftest.py`, `.devcontainer/*`.
- В `app.py`:
  - режим «Markdown → DOCX»;
  - `DOC_TYPES`;
  - `_mineru_provider_factory`;
  - `_pdf_page_subset`;
  - `_split_findings` и `_findings_table` — поведение и подписи колонок;
  - ключи результата `_convert_uploaded_file` (`RESULT_KEYS` закрыт тестом #85);
  - CSS в `st.markdown("<style>…")`, `st.set_page_config`.
- Существующие тесты `tests/test_app_fixes.py` — не править. Они обязаны пройти как есть.

## Интерфейсы (дословно)

### Импорты

```python
from ocr.claude_code_verifier import ClaudeCodeMissingError, find_claude   # ПРАВКА #92
from ocr.vision import VISION_MODEL                                        # ПРАВКА #92
```

После `from ocr.ingest import …`: так `ocr.ingest` уже загружен, когда `ocr.vision` тянет `ocr.measure` →
`ocr.board` → `ocr.ingest`. Порядок проверяет обычный `import app` в тестах.

### Константы и помощники

```python
_VISION_RULE = "vision_diff"                # ПРАВКА #92: подсказки — своим блоком, не в общей таблице


@st.cache_resource(show_spinner=False)
def _claude_available() -> bool:
    """ПРАВКА #92: сверка по картинке — только где есть claude; на Streamlit Cloud его нет."""
    try:
        find_claude()
    except ClaudeCodeMissingError:
        return False
    return True


def _vision_progress(status):
    """ПРАВКА #92: vision_progress для ingest — метка st.status по страницам; без status — None."""
    if status is None:
        return None

    def update(page: int, done: int, total: int) -> None:
        status.update(label=f"Сверка по картинке: стр. {page} ({done} из {total})…")

    return update


def _vision_rows(findings: list[dict]) -> list[dict]:
    """ПРАВКА #92: подсказки сверки по картинке — страница, было, по скану, стало."""
    return [{"Страница": f["page"], "Было": f["snippet"], "По скану": f["reading"],
             "Стало": f["suggestion"] or "∅"} for f in findings]
```

- `_claude_available` кэшируется на процесс, как `_ocrmypdf_available`: `find_claude` — это `shutil.which`,
  но звать его на каждый rerun незачем. Поставленный на ходу `claude` появится после перезапуска Streamlit.
- «Стало» у удаления («у нас лишнее», `suggestion == ""`) — `∅` (PLACEHOLDER 5 спеки 16).

### `_render_ocr_report(report)` (#92)

- Сводка и `low_confidence` — как были.
- `main, low_conf = _split_findings(report)`, затем `hints = [f for f in main if f["rule"] == _VISION_RULE]`,
  `main = [f for f in main if f["rule"] != _VISION_RULE]`. `vision_skipped` остаётся в `main`: это причина
  «сверка не выполнена», её видно в общем списке.
- После блока `low_confidence`, если `hints` непусты:

  ```python
  with st.expander(f"Сверка по картинке — подсказки ({len(hints)})", expanded=False):
      st.caption(f"Модель {hints[0]['model']} переписала страницы скана; здесь — места, где прочитанное "
                 "расходится с текстом. Текст не изменён. Гомоглифы, пунктуация и тире не подсказываются.")
      st.dataframe(_vision_rows(hints), use_container_width=True, hide_index=True)
  ```

- Порядок подсказок — как в отчёте.

### `_convert_uploaded_file` (#92)

```python
def _convert_uploaded_file(uploaded_file, page_range: str | None,
                           ocr_mode: str = "off", *, verify: bool = False,
                           vision: bool = False,                 # ПРАВКА #92
                           annotate: bool = True, status=None) -> dict:
```

В ветке `ocr_mode == "mineru" and ext in _INGEST_EXTS` вызов `ingest` получает ещё два аргумента:

```python
                    # ПРАВКА #92: сверка по картинке — только на маршрутах MinerU (на text/office ingest бросил бы ValueError)
                    vision=VISION_MODEL if vision and route in _MINERU_ROUTES else None,
                    vision_progress=_vision_progress(status),
```

Остальное без изменений: ключи результата, обработка ошибок, `MineruAuthError`.

### `render_files_to_markdown_mode` (#92)

- `vision = False` — рядом с `verify = False`.
- В ветке MinerU, **между** галочками «Сверка вторым прогоном» и «Пометки в тексте»:

  ```python
        if _claude_available():     # ПРАВКА #92: на Streamlit Cloud claude нет — галочки нет
            vision = st.checkbox(
                "Сверка по картинке (Claude Code, локально)",
                value=False,
                key="files_to_md_vision",
                help="Каждая страница скана или PDF с таблицами переписывается моделью "
                     f"{VISION_MODEL} через claude -p; расхождения с текстом — подсказки в находках, "
                     "текст не меняется. Одна страница — один вызов подписки Claude Code; повтор — из кэша.",
            )
  ```

- В цикле конвертации (ветка `st.status`) — `_convert_uploaded_file(..., verify=verify, vision=vision,
  annotate=annotate, status=status)`.
- Метка `st.status` после сверки возвращается к «готово» / «ошибка» существующим кодом — он её и так
  перезаписывает.
- В карточке результата, сразу после подписи «Сверка вторым прогоном не применялась…»:

  ```python
                if (st.session_state.get("files_to_md_vision")
                        and route not in _MINERU_ROUTES):
                    st.caption("Сверка по картинке не применялась: файл не шёл через MinerU.")
  ```

- Подпись о пометках (`ANNOTATION_PREFIX in markdown`) уже есть и покрывает пометки `vision_diff`: они той же
  формы.

## Приёмочные тесты

В `tests/test_app_fixes.py`, блок `# --- ПРАВКА #92: сверка по картинке ---`. Все — с `_offline` (фикстура
`no_network`). Дополнительно `subprocess.run` — заглушка с `AssertionError`: реальный `claude` не
запускается ни в одном тесте.

```python
class _FakeStatus:
    def __init__(self):
        self.labels = []

    def update(self, label=None, state=None):
        self.labels.append(label)


def test_claude_available_follows_find_claude(monkeypatch):
    def missing():
        raise app.ClaudeCodeMissingError("нет")
    monkeypatch.setattr(app, "find_claude", missing)
    app._claude_available.clear()
    assert app._claude_available() is False
    monkeypatch.setattr(app, "find_claude", lambda: "C:/x/claude.cmd")
    app._claude_available.clear()
    assert app._claude_available() is True
    app._claude_available.clear()                       # не оставлять состояние другим тестам


def test_vision_rows():
    f = {"page": 3, "snippet": "IP65не", "reading": "… IP65 не ниже …", "suggestion": "IP65 не", "model": "m"}
    d = dict(f, snippet="лишнее", suggestion="")
    assert app._vision_rows([f, d]) == [
        {"Страница": 3, "Было": "IP65не", "По скану": "… IP65 не ниже …", "Стало": "IP65 не"},
        {"Страница": 3, "Было": "лишнее", "По скану": "… IP65 не ниже …", "Стало": "∅"}]


# сверка подменена: ocr.vision.vision_findings зовёт progress(1, 1, 2), progress(4, 2, 2)
#   и отдаёт одну vision_diff (reading, model = VISION_MODEL); записывает свои аргументы
def test_mineru_mode_vision_passes_model_and_progress(monkeypatch, tmp_path): ...
#   _fake_mineru(vlm="# Договор\n\nТекст договора."), скан _blank_pdf(1), vision=True, status=_FakeStatus()
#   -> error None; report["schema_version"] == 2; одна находка vision_diff с reading и model == VISION_MODEL;
#   подменённая сверка вызвана с model == VISION_MODEL;
#   status.labels содержит "Сверка по картинке: стр. 1 (1 из 2)…" и "Сверка по картинке: стр. 4 (2 из 2)…"
#   (лишних «из 2» нет); set(result) == RESULT_KEYS

def test_mineru_mode_without_vision_does_not_call_it(monkeypatch, tmp_path): ...
#   то же, vision=False -> подменённая сверка не вызывалась; находок vision_* нет

@pytest.mark.parametrize("name, data", [("док.docx", _docx_bytes()), ("табл.xlsx", _xlsx_bytes())],
                         ids=["docx", "xlsx"])
def test_vision_skipped_outside_mineru_routes(monkeypatch, tmp_path, name, data): ...
#   vision=True -> error None (не ValueError), route "office", сверка не вызывалась
#   (та же проверка — текстовый PDF: require_fixture("textpdf1.pdf"), диапазон "1-2", route "text")

def test_vision_failure_keeps_conversion(monkeypatch, tmp_path): ...
#   настоящая ocr.vision.vision_findings; _fake_mineru кладёт zip только с full.md (нет content_list) ->
#   error None, markdown есть, в report одна vision_skipped «сверка не выполнена: ValueError: …»;
#   ocr.vision.make_verifier подменён на фейк, который падает при создании — он не вызывался
```

Подмена — `monkeypatch.setattr(ocr.vision, "vision_findings", fake)`: `ingest` импортирует имя внутри
функции, в момент вызова (спека 16). Галочку и expander тесты не рендерят: Streamlit-рендер в тестах
проекта не проверяется (как в #73 и #85). Проверяются помощники и `_convert_uploaded_file`.

### Руками (в отчёт прогона, не в тест)

Делает человек: живая сверка тратит подписку.
1. Запустить `streamlit run app.py` (локально, где есть `claude`). Режим «Файлы → Markdown» → MinerU.
   Галочка «Сверка по картинке (Claude Code, локально)» есть, между «Сверка вторым прогоном» и «Пометки в
   тексте».
2. Загрузить `_test/fixtures/ocr/bakeoff.pdf`, поставить галочку, нажать «Конвертировать»:
   - в статусе идут «Сверка по картинке: стр. N (k из 9)…»;
   - в карточке — блок подсказок с колонками Страница / Было / По скану / Стало;
   - с «Пометками в тексте» в превью есть `!! ПРОВЕРИТЬ: … vision_diff: … !!`.
3. Повторить ту же конвертацию: сверка из кэша, статус пробегает за секунды.
4. Без «Пометок в тексте» превью совпадает с конвертацией без сверки.
5. DOCX с галочкой → подпись «Сверка по картинке не применялась…».
6. `claude` вне PATH (например, запуск из окружения без него) → галочки нет, остальной режим MinerU как
   раньше.

## Готово, когда

- `pytest -v` зелёный целиком; существующие тесты `test_app_fixes.py` не правились; сеть и `claude` не
  тронуты.
- `git diff --stat` — только `app.py`, `tests/test_app_fixes.py` и документы из «Трогать».
- `git diff app.py`:
  - нет изменений CSS, `set_page_config` и режима «Markdown → DOCX»;
  - нет новых видов виджетов — только `st.checkbox`, `st.expander`, `st.dataframe`, `st.caption`.
- В docs_sync `### ПРАВКА #92`:
  - «Что изменится для человека» — дословно;
  - отклонения от буквы спеки, если были;
  - строка «Руками — не проверялось» или результат проверки человека;
  - снимок нумерации: «следующий — #93».

## PLACEHOLDER-ы

1. **Длинный документ.** Одна страница — один вызов (~15 с по v3) плюс ~3 с вердикта. 100 страниц — около
   получаса в одном `st.status`, rerun Streamlit всё это время ждёт. Ограничитель уже есть — диапазон
   страниц. Отдельный лимит страниц для сверки — решение человека после живой проверки.
2. **`_claude_available` кэшируется на процесс.** Установленный или удалённый на ходу `claude` учитывается
   после перезапуска Streamlit.
3. **Вход не проверяется заранее.** `claude` найден, но не залогинен → первая страница даёт «сверка
   остановлена на стр. 1: …/login…» в находках. Проверять вход до конвертации — это лишний вызов `claude` на
   каждый rerun; не делаем.
4. **Номера страниц в подсказках** при заданном диапазоне считаются от вырезки, как у всех находок режима
   MinerU (#73).

## Коммит

Один коммит кода и документов. Push делает человек.

```
ПРАВКА #92: сверка по картинке в UI — галочка, прогресс по страницам, блок подсказок

app.py: в режиме MinerU галочка «Сверка по картинке (Claude Code, локально)» —
только если claude найден на машине (на Streamlit Cloud его нет, галочки нет).
Модель — claude-opus-5 (ocr.vision.VISION_MODEL), выбора в интерфейсе нет.
Сверка передаётся в ingest только на маршрутах MinerU; прогресс — метка
st.status по страницам. В карточке результата подсказки vision_diff — своим
блоком: страница, было, по скану, стало; текст они не меняют, пометки в
Markdown — через существующую галочку. Неудача сверки — находка vision_skipped,
конвертация цела. Стиль UI не менялся.
```
