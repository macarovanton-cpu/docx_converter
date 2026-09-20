# 05 — Валидатор и отчёт

**# ПРАВКА #64** (+ #70, #72, #77, #79, #80). Зависит от: 04 (`Finding`, `parse_pipe_tables`, `fold_to_cyrillic`).

## Цель

Проверки по правилам над уже постобработанным markdown; сборка `report.json` —
источника истины о сомнительных местах; пометки в markdown как производная от отчёта.

Граница возможностей зафиксирована честно: «сыручими» и «IR-камерами» правилами
**не ловятся** — это грамматически и форматно допустимый текст. Их ловит только
сверка прогонов (спека 06). «Валидатор ловит все известные дефекты бейкоффа» из
PLAN читается как «объединение находок 04 + 05 + 06».

## Трогать

- `ocr/validate.py` — создать
- `tests/test_ocr_validate.py` — создать

## Не трогать

`ocr/postprocess.py`, `ocr/__init__.py`, всё из общего списка запретов.

## Интерфейсы (дословно)

```python
"""ПРАВКА #64: валидатор OCR-markdown, report.json, пометки в markdown."""

REPORT_SCHEMA_VERSION = 1
ANNOTATION_PREFIX = "!! ПРОВЕРИТЬ: "
ANNOTATION_SUFFIX = " !!"
LOST_HEADING = "## Не привязанные находки"       # ПРАВКА #70

UNITS = ("мм", "см", "м", "км", "г", "кг", "т", "л", "шт", "В", "А", "Вт", "кВт",
         "кВА", "Гц", "лм", "ч", "с", "мин", "сут")
SCALE_FAMILIES = ("ВЕСТА-С", "РЕУС-Д", "ВЕКТОР-ДВ")
GLUED_CODE_SUGGESTION = "проверить границу с следующим пунктом"   # ПРАВКА #79
ORG_NAME_MIN_LEN = 5                                             # ПРАВКА #80


def validate(md: str, content_list: list | None = None) -> list[Finding]: ...

def page_of(snippet: str, content_list: list | None) -> int | None: ...

def build_report(*, source: str, sha256: str, provider: str,
                 model_version: str | None, cache_hit: bool, verified: bool,
                 findings: list[Finding],
                 content_list: list | None = None) -> dict: ...

def annotate(md: str, report: dict, *,
             include_low_confidence: bool = False) -> str: ...   # ПРАВКА #70
def strip_annotations(md: str) -> str: ...
```

### Схема `report.json` (закрытая)

```json
{
  "schema_version": 1,
  "source": "bakeoff.pdf",
  "sha256": "3f…9a",
  "provider": "mineru",
  "model_version": "vlm",
  "cache_hit": false,
  "verified": false,
  "created_at": "2026-09-18T12:00:00Z",
  "summary": {"critical": 5, "warning": 3, "info": 1},
  "findings": [
    {"id": 1, "rule": "translit_suspect", "severity": "critical", "page": 9,
     "snippet": "A.P. Сиражитдинов", "suggestion": "А.Р. Сиражитдинов"}
  ]
}
```

- Ключи верхнего уровня — ровно эти десять; ключи находки — ровно эти шесть.
- `id` — с 1, по порядку списка. `summary` всегда содержит все три ключа.
- `page`: если у находки `None` и дан `content_list` — `page_of(snippet, content_list)`.
- `created_at` — UTC, ISO-8601, секунды, `Z`. `verified` — был ли прогон сверки (06).
- Словарь сериализуется `json.dumps(..., ensure_ascii=False)` без доп. кодеков.

### `page_of`

`content_list` MinerU — список блоков `{"type", "text" | "table_body", "page_idx", …}`,
`page_idx` с нуля. Возвращает `page_idx + 1` первого блока, у которого в `text` или в
`table_body` (после снятия тегов и схлопывания пробелов) содержится `snippet`
(тоже со схлопнутыми пробелами). Нет `content_list`, не нашлось, у блока нет
`page_idx` → `None`.

**ПРАВКА #72: PLACEHOLDER 3 закрыт.** Форма блока сверена с настоящим
`content_list.json` из `vlm_raw.zip` (живой прогон `vlm`, 35 блоков). В архиве
файл лежит под именем `<uuid>_content_list.json` — так его и ищет
`result_from_zip` (`n.endswith("content_list.json")`), `content_list_v2.json`
под этот хвост не попадает. Дока подтвердилась: у каждого блока есть `type` и
`page_idx` (с нуля), текст лежит во взаимоисключающих `text` / `table_body`;
сверх них встречаются `bbox`, `img_path`, `text_level`, `table_caption`,
`table_footnote` — `page_of` их не читает. Тест — на настоящем файле, синтетика
рядом остаётся (она проверяет краевые случаи, которых в архиве нет).

### Правила

Серьёзность — по таблице в README. `snippet` — дословный фрагмент из `md`.

| rule | Когда срабатывает |
|---|---|
| `inn_checksum` | `ИНН` + до 3 нецифровых символов + цифры. Длина не 10 и не 12 → находка. Иначе контрольная сумма: 10 знаков — веса `2,4,10,3,5,9,4,6,8`, `sum % 11 % 10 == d[9]`; 12 знаков — `d[10]` по весам `7,2,4,10,3,5,9,4,6,8`, `d[11]` по весам `3,7,2,4,10,3,5,9,4,6,8` |
| `ogrn_checksum` | `ОГРН`/`ОГРНИП` + цифры. 13 знаков: `int(d[:12]) % 11 % 10 == d[12]`; 15 знаков: `int(d[:14]) % 13 % 10 == d[14]`; иная длина → находка |
| `kpp_format` | `КПП` + значение не по `\d{4}[\dA-Z]{2}\d{3}` |
| `gost_format` | токен, у которого `fold_to_cyrillic(t).upper() == "ГОСТ"`: (а) написан не ровно `ГОСТ` → находка, `suggestion="ГОСТ"`; (б) дальше не идёт `( Р)? \d+(\.\d+)*-\d{2,4}` → находка. `ТУ` проверяется только если сразу за ним цифра: шаблон `\d+(\.\d+)*-\d+-\d+-\d{2,4}` |
| `unit_unknown` | после цифры (через необязательный пробел) токен из 1–3 букв, в котором есть латиница, `fold_to_cyrillic(t).lower()` ∈ `UNITS` (без учёта регистра), а сам `t` ∉ `UNITS`. `suggestion` — каноническая единица |
| `scale_model` | токен (по пробелам, без краевой пунктуации), у которого `fold_to_cyrillic(t).upper()` начинается с элемента `SCALE_FAMILIES`, а сам `t` с него не начинается. `suggestion` — каноническое написание + остаток токена |
| `code_digits_glued` | **ПРАВКА #79.** Обозначение `СП`/`СНиП`/`ГОСТ`/`ГОСТ Р`/`ТУ` (граница слова слева) плюс номер `\d+((\.|[-–—])\d+)*`, в котором числовых групп больше одной, а последняя длиннее 4 цифр → находка, `suggestion` — `GLUED_CODE_SUGGESTION`. Последняя группа составного номера — год; длиннее — к нему прилип номер следующего пункта (`СП 76.13330.20163.` = `СП 76.13330.2016` и пункт `3.`). Номер из одной группы (`ГОСТ Р53228`) — не год, под правило не попадает. Текст не правится: границу видно только по скану |
| `org_name_variant` | **ПРАВКА #80.** Содержимое «ёлочек» без вложенных (`«([^«»]+)»`) длиной ≥ `ORG_NAME_MIN_LEN`. Два варианта одной длины, различающиеся ровно в одной позиции, **и оба различающихся символа — буквы** → находка на более редкий: `snippet` — редкий вариант в кавычках, `suggestion` — частый. Равная частота находки не даёт — кто из двоих опечатка, не сказать. Требование «оба символа — буквы» снимает перенос строки вместо пробела внутри одного и того же названия: это не вариант написания |
| `table_total_mismatch` | строка таблицы, первая непустая ячейка которой начинается с `Итого` (без учёта регистра). Для каждой колонки, где все ячейки строк выше (кроме шапки) и ячейка итога — числа (`,`→`.`, пробелы внутри числа убрать): `abs(sum - total) > 0.01` → находка, `suggestion` — посчитанная сумма |
| `table_row_cells` | число ячеек строки ≠ числу ячеек шапки |
| `table_empty_number` | первая колонка — нумерационная, если не меньше половины строк таблицы (кроме первой) имеют в первой ячейке `^\d+(\.\d+)*\.?$`. В такой таблице любая строка, **включая первую**, с пустой первой ячейкой и хотя бы одной непустой другой → находка. Первая строка проверяется нарочно: у продолжения, разорванного страницей, «шапкой» становится именно строка с потерянным номером. На шапку `№` не опираемся — у продолжения её нет |

Таблицы берутся через `ocr.postprocess.parse_pipe_tables`. Сырая `<table` в тексте
валидатором не разбирается — о ней уже сказал `html_table_unparsed`.

### `annotate` / `strip_annotations`

- Пометка — **отдельный блок** (пустая строка до и после):
  `!! ПРОВЕРИТЬ: [{id}] {rule}: {suggestion or snippet} !!`
  Это существующий синтаксис callout в `convert.py` (`text.startswith('!!') and text.endswith('!!')`);
  `convert.py` не трогать.
- **ПРАВКА #70:** находки `low_confidence` по умолчанию в текст не идут — их на
  документ десятки, и за ними не видно находок правил; в `report.json` они есть
  целиком, а `include_low_confidence=True` (`--annotate-all`) возвращает их в текст.
- Место: сразу после блока (разбиение по `\n\n`), в котором впервые встретился
  `snippet`. Блок — таблица → пометка после всей таблицы (внутрь таблицы callout
  вставить нельзя). `snippet` не найден или пуст → пометка **в конец документа**,
  в отдельный раздел `LOST_HEADING` (ПРАВКА #70: пустой `snippet` находится в
  любом блоке, и пометка садилась первой строкой документа, в шапку). Находка
  при этом не теряется.
- Несколько находок на один блок — пометки подряд в порядке `id`.
- Переводы строк внутри текста пометки → пробел; текст обрезается до 200 символов с `…`.
- `strip_annotations` удаляет блоки, целиком совпадающие с
  `^!! ПРОВЕРИТЬ: .* !!$`, и восстанавливает разделители. Прочие callout-ы
  (`!! формула !!`) не трогает. Блок `LOST_HEADING` снимается только когда после
  него до конца документа одни пометки — то есть когда его поставил `annotate`;
  такой же заголовок в тексте самого документа остаётся на месте.
- Инвариант: `strip_annotations(annotate(md, report)) == md` для любого `md`,
  нормализованного `cleanup_ocr_markdown` (без тройных переводов строк).

## Приёмочные тесты (`tests/test_ocr_validate.py`)

```python
vlm, golden, pipeline = (read_fixture(n) for n in ("vlm.md", "golden.md", "pipeline.md"))
md, post = postprocess(vlm)
val = validate(md)

# на vlm: реквизиты и ГОСТы в порядке — валидатор молчит по ним
assert "ОГРН 1020202283287" in md                                # сумма сходится: 102020228328 % 11 % 10 == 7
assert [f for f in val if f.rule in ("ogrn_checksum", "inn_checksum", "gost_format",
                                     "unit_unknown", "table_row_cells",
                                     "table_empty_number", "table_total_mismatch")] == []
assert validate(golden) == []

# без склейки валидатор сам видит потерянный номер пункта
unmerged = html_tables_to_pipe(vlm)[0]
v = validate(unmerged)
assert [f.rule for f in v].count("table_empty_number") == 1      # первая строка табл. 3: '', '', текст
# таблица 2 (одна строка '', текст) — первая колонка не нумерационная, под правило не попадает

# на шумном pipeline: ГОСТ гомоглифами/регистром и единица латиницей
pv = validate(postprocess(pipeline)[0])
assert [f.rule for f in pv].count("gost_format") >= 3            # «ГОст», «ГОСт», «гОСт», «ГОСТ Р53228»
assert any(f.rule == "unit_unknown" and f.suggestion == "мм" for f in pv)   # «10MM.»

# отчёт
report = build_report(source="bakeoff.pdf", sha256="ab" * 32, provider="mineru",
                      model_version="vlm", cache_hit=False, verified=False,
                      findings=post + val)
assert list(report) == ["schema_version", "source", "sha256", "provider", "model_version",
                        "cache_hit", "verified", "created_at", "summary", "findings"]
assert report["schema_version"] == 1
assert [f["id"] for f in report["findings"]] == list(range(1, len(post + val) + 1))
assert all(list(f) == ["id", "rule", "severity", "page", "snippet", "suggestion"]
           for f in report["findings"])
assert report["summary"]["critical"] == 5                        # пять ФИО с латиницей
assert sum(report["summary"].values()) == len(report["findings"])
assert set(report["summary"]) == {"critical", "warning", "info"}
assert json.loads(json.dumps(report, ensure_ascii=False)) == report

# страницы: настоящий content_list живого прогона (ПРАВКА #72, PLACEHOLDER 3)
names = [n for n in zipfile.ZipFile(require_fixture("vlm_raw.zip")).namelist()
         if n.endswith("content_list.json")]
assert len(names) == 1
real = json.loads(zipfile.ZipFile(require_fixture("vlm_raw.zip")).read(names[0]).decode("utf-8"))
assert all("page_idx" in b and "type" in b for b in real)
assert page_of("сыручими", real) == 2 and page_of("A.II. Taipov", real) == 9
assert page_of("такого в документе нет", real) is None

# страницы: синтетика на краевые случаи
cl = [{"type": "text", "text": "P.P. Hypeeb", "page_idx": 8},
      {"type": "table", "table_body": "<table><tr><td>сыручими  и</td></tr></table>", "page_idx": 1}]
assert page_of("P.P. Hypeeb", cl) == 9 and page_of("сыручими и", cl) == 2
assert page_of("нет такого", cl) is None and page_of("x", None) is None
r2 = build_report(source="s", sha256="0" * 64, provider="mineru", model_version="vlm",
                  cache_hit=True, verified=False, findings=post, content_list=cl)
assert next(f for f in r2["findings"] if f["snippet"] == "P.P. Hypeeb")["page"] == 9

# пометки
ann = annotate(md, report)
assert strip_annotations(ann) == md
assert ann.count(ANNOTATION_PREFIX) == len(report["findings"])
assert parse_pipe_tables(ann) == parse_pipe_tables(md)           # таблица не разорвана пометками
block = next(b for b in ann.split("\n\n") if "translit_suspect" in b and "Hypeeb" in b)
assert block.startswith("!!") and block.endswith("!!") and "\n" not in block
assert ann.split("\n\n").index(block) == ann.split("\n\n").index("P.P. Hypeeb") + 1
assert strip_annotations("!! E = mc2 !!\n\nТекст") == "!! E = mc2 !!\n\nТекст"
low = {"id": 1, "rule": "low_confidence", "severity": "warning",
       "page": None, "snippet": "нет в тексте", "suggestion": None}
lost = annotate("Абзац.", {"findings": [low]}, include_low_confidence=True)
assert lost.endswith("!! ПРОВЕРИТЬ: [1] low_confidence: нет в тексте !!")   # не потеряна

# ПРАВКА #70: по умолчанию low_confidence в текст не идёт, непривязанное — в хвост
assert annotate("Абзац.", {"findings": [low]}) == "Абзац."
md2 = "# Шапка\n\nТело."
ann2 = annotate(md2, {"findings": [low, dict(low, id=2, snippet="")]},
                include_low_confidence=True)
assert ann2.startswith(md2) and ann2.split("\n\n")[2] == LOST_HEADING
assert strip_annotations(ann2) == md2
own = "## Не привязанные находки\n\nТекст."
assert strip_annotations(own) == own            # свой такой же заголовок не трогаем
```

Синтетика — по тесту на правило:

```python
def rules(text): return [f.rule for f in validate(text)]

assert rules("ИНН 7707083893") == [] and rules("ИНН 7707083894") == ["inn_checksum"]
assert rules("ИНН: 500100732259") == [] and rules("ИНН 77070838") == ["inn_checksum"]
assert rules("ОГРН 1027700132195") == [] and rules("ОГРН 1027700132196") == ["ogrn_checksum"]
assert rules("ОГРНИП 304500116000157") == []
assert rules("КПП 773601001") == [] and rules("КПП 77360100") == ["kpp_format"]
assert rules("ГОСТ Р 53228-2008, ГОСТ 8.726-2010") == []
assert rules("ГОСт 380-2005") == ["gost_format"] and rules("ГОСТ 380") == ["gost_format"]
assert rules("(ТУ) на подключение") == [] and rules("ТУ 4274-001-12345678-2015") == []
# ПРАВКА #79 / #80
assert rules("по СП 76.13330.20163. Подрядчик") == ["code_digits_glued"]
for clean in ("СП 131.13330.2020", "ГОСТ Р 58760—2019", "ТУ 4274-001-12345678-2015",
              "ГОСТ Р 53228"):                       # одна группа — номер, не год
    assert "code_digits_glued" not in rules(clean)
pair = "«ГПИ имени Д.С. Косьяна» «ГПП имени Д.С. Косьяна» «ГПП имени Д.С. Косьяна»"
assert [(f.snippet, f.suggestion) for f in validate(pair)] == [
    ("«ГПИ имени Д.С. Косьяна»", "«ГПП имени Д.С. Косьяна»")]
assert rules("«ГПИ имени Косьяна» и «ГПП имени Косьяна»") == []          # частота равна
assert rules("«Башкирская\nкомпания» «Башкирская компания» «Башкирская компания»") == []
assert rules("не менее 10 MM") == ["unit_unknown"] and rules("не менее 10 мм, 12В") == []
assert rules("весы BЕСТА-С60") == ["scale_model"] and rules("весы ВЕСТА-С60") == []
t = "| № | Сумма |\n|---|---|\n| 1 | 10,5 |\n| 2 | 1 000 |\n| Итого | 1 010,5 |"
assert rules(t) == [] and rules(t.replace("1 010,5", "1 011")) == ["table_total_mismatch"]
assert rules("| № | A |\n|---|---|\n| 1 | x | y |") == ["table_row_cells"]
assert rules("| № | A |\n|---|---|\n| 1 | a |\n|  | x |\n| 2 | b |") == ["table_empty_number"]
assert rules("|  | x |\n|---|---|\n| 26 | a |\n| 27 | b |") == ["table_empty_number"]   # потерян номер в «шапке»
assert rules("|  | 2024 |\n|---|---|\n| план | 1 |\n| факт | 2 |") == []               # матрица, не нумерация
assert rules("| № | A | B |\n|---|---|---|\n| 1 | a | b |\n| Раздел |  |  |") == []
```

## Готово, когда

- `pytest -v` зелёный, фикстурные тесты не пропущены.
- Схема отчёта в коде совпадает со спекой ключ в ключ, порядок ключей тот же.
- `git diff --stat`: только `ocr/validate.py`, `tests/test_ocr_validate.py`.
- В отчёте исполнителя — таблица «известный дефект → каким правилом пойман»;
  для «сыручими» и «IR-камерами» в ней стоит «спека 06».

## Коммит

```
ПРАВКА #64: валидатор OCR-markdown и report.json

ocr/validate.py: ИНН/ОГРН с контрольными суммами, КПП, ГОСТ/ТУ, единицы и марки
весов гомоглифами, «Итого», целостность таблиц. build_report — схема v1,
страницы из content_list. annotate/strip_annotations — пометки
«!! ПРОВЕРИТЬ: … !!» как производная отчёта, снимаются без следа.
```
