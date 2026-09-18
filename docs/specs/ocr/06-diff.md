# 06 — Детектор расхождений между прогонами

**# ПРАВКА #65** (+ #70). Зависит от: 04 (`HOMOGLYPHS_CYR_TO_LAT`), 05 (`page_of`).

## Цель

Два прогона MinerU (`vlm` и `pipeline`) ошибаются в разных местах. Место, где
они разошлись, — кандидат на ошибку OCR. Модуль сравнивает два уже
постобработанных markdown и выдаёт находки `low_confidence`. Это единственный
механизм тракта, который ловит «сыручими» и «IR-камерами».

Модуль чистый: сеть, кэш и второй прогон — забота CLI (спека 07). Текст не
меняется, выбор «кто прав» не делается.

## Трогать

- `ocr/diff.py` — создать
- `tests/test_ocr_diff.py` — создать

## Не трогать

`ocr/postprocess.py`, `ocr/validate.py`, всё из общего списка запретов.

## Интерфейсы (дословно)

```python
"""ПРАВКА #65: сверка двух OCR-прогонов, расхождения -> находки low_confidence."""

SNIPPET_MAX = 120
MIN_RUN = 3        # ПРАВКА #70: расхождение короче — джиттер OCR, не находка


def normalize(text: str) -> str: ...

def diff_findings(primary_md: str, secondary_md: str,
                  content_list: list | None = None) -> list[Finding]: ...
```

### `normalize`

Строго в этом порядке:

1. `unicodedata.normalize("NFKC", text)`;
2. `\bNo\.?\s*(?=\d|п/п)` → `№` и `№\s*` → `№`;
3. свёртка гомоглифов **латиница → кириллица**, посимвольно и безусловно, по
   обращённой `ocr.postprocess.HOMOGLYPHS_CYR_TO_LAT` (до приведения регистра —
   в таблице есть пары только для заглавных);
4. `lower()`, `ё` → `е`;
5. всё, что не буква, не цифра и не `№`, → пробел;
6. **ПРАВКА #70: все пробелы и переносы удаляются целиком** — на выходе
   сплошная последовательность символов.

Свёртка безусловная нарочно: обе стороны сворачиваются одинаково, а «IР54» и
«IP54» обязаны совпасть. Цена — `РоЕ`/`PoE` неразличимы; это уже починено в 04.

Пробелы убраны потому, что три четверти находок первого живого прогона были не
расхождениями чтения, а разной разбивкой на слова: один прогон отдаёт
«твердыми, сыпучими», другой «твердыми,сыпучими» — MinerU теряет пробел на
переносе строки внутри ячейки. Толку в такой находке нет, а человек, читая
отчёт, тонет в ней. `normalize("твердыми, сыпучими") == normalize("твердыми,сыпучими")`.

### `diff_findings`

Единица сравнения — **символ**, не слово, не абзац и не строка (ПРАВКА #70;
до неё было слово, до PLAN-а — абзац; таблица MinerU — один абзац на 30 КБ, а
таблицы двух прогонов порезаны на строки по-разному, так что ни абзацы, ни
строки не выравниваются).

Проход двухуровневый — так же, как выглядит результат глобальной посимвольной
сверки, но за 0,07 с вместо 13 с (`difflib` на 17 тыс. символов против 3,5 тыс.
токенов). Ускорение здесь не «наугад»: посимвольный проход в лоб не укладывается
в бюджет 10 с уже на девяти страницах.

1. Токены стороны: `re.finditer(r"[^\s|]+", md)`, строки-разделители таблиц
   (`|---|`) пропускаются. Токен помнит **свой срез** `(start, end)` в `md` и
   свою нормализованную запись `normalize(токен)` (пустая — токен был чистой
   разметкой, он выбрасывается).
2. Грубое выравнивание: `difflib.SequenceMatcher(None, токены_primary,
   токены_secondary, autojunk=False)` по нормализованным записям.
3. Каждый опкод, кроме `equal`, — **кусок**. Символы куска склеиваются подряд
   (каждый помнит свой токен). Строки куска совпали — расхождение было только
   в разбивке на слова, находок нет.
4. Иначе внутри куска — посимвольный `SequenceMatcher`. Каждый не-`equal` опкод:
   - в расходящихся символах нет ни буквы, ни цифры (разошёлся один `№`) —
     пропустить;
   - попал в тот же токен primary, что и предыдущий, — это одна находка
     («Hypeeb» против «Hyreev» расходится в двух местах, место одно);
   - у primary символов нет (вставка) — цепляемся к соседнему токену слева,
     нет слева — справа.
5. Находка: `rule="low_confidence"`, `severity="warning"`;
   - `snippet` — **дословный срез** `primary_md` от начала первого до конца
     последнего затронутого токена. Длиннее `SNIPPET_MAX` → обрезать по границе
     токена. Короче `MIN_RUN` — **не находка**: одна-две буквы в коротком токене
     («и» против «в», «5» против «6») — джиттер OCR, а не разное чтение;
   - `suggestion` — то же со стороны secondary (что увидел второй прогон),
     обрезка та же; ничего не затронуто (`delete`) — `None`;
   - `page` — `ocr.validate.page_of(snippet, content_list)`.
6. Порядок находок — по позиции в `primary_md`. Одинаковые `(snippet, suggestion)`
   подряд не схлопывать: каждая — своё место в документе.

Инварианты: `diff_findings(x, x) == []`; у каждой находки `snippet in primary_md`
(по нему `annotate` ищет место).

## Приёмочные тесты (`tests/test_ocr_diff.py`)

```python
vlm, pipeline = read_fixture("vlm.md"), read_fixture("pipeline.md")
a, b = postprocess(vlm)[0], postprocess(pipeline)[0]
fs = diff_findings(a, b)

# три известных дефекта подсвечены
syr = [f for f in fs if "сыручими" in f.snippet]
assert len(syr) == 1 and "сыпучими" in syr[0].suggestion         # второй прогон видел правильно
ir = [f for f in fs if "IR-камерами" in f.snippet]
assert len(ir) == 1 and normalize("IP-камерами") in normalize(ir[0].suggestion)
assert any("Taipov" in f.snippet for f in fs)                    # блок подписей
assert any("Hypeeb" in f.snippet for f in fs)

# контракт находок
assert all((f.rule, f.severity) == ("low_confidence", "warning") for f in fs)
assert all(f.page is None for f in fs)                           # content_list в фикстурах нет
assert all(f.snippet and f.snippet in a for f in fs)
assert all(len(f.snippet) <= SNIPPET_MAX for f in fs)
assert all(f.suggestion is None or len(f.suggestion) <= SNIPPET_MAX for f in fs)
assert diff_findings(a, a) == [] and diff_findings(b, b) == []

# шум: вход фиксированный, поэтому число точное, а не граница (ПРАВКА #70)
assert len(fs) == FIXTURE_FINDINGS      # 65; было 212 при пословном выравнивании
# совпадающее не шумит: ОГРН оба прогона прочитали одинаково
assert not [f for f in fs if "1020202283287" in f.snippet]

# находки стыкуются с отчётом и пометками
report = build_report(source="bakeoff.pdf", sha256="0" * 64, provider="mineru",
                      model_version="vlm", cache_hit=True, verified=True, findings=fs)
assert strip_annotations(annotate(a, report)) == a
cl = [{"type": "table", "table_body": "<td>твердыми, сыручими и жидкими</td>", "page_idx": 1}]
assert [f.page for f in diff_findings(a, b, cl) if "сыручими" in f.snippet] == [2]
```

Синтетика:

```python
assert normalize("No384-ФЗ") == normalize("№ 384-ФЗ") == "№384фз"
assert normalize("Noп/п") == normalize("№ п/п")
assert normalize("IР54") == normalize("IP54")                    # кир. Р против лат. P
assert normalize("«НПФ»  БЗК") == normalize("НПф БЗК")
assert normalize("Ёлка,  ёж.") == "елкаеж"
assert normalize("твердыми, сыпучими") == normalize("твердыми,сыпучими")   # ради этого всё
assert normalize("№384") == normalize("№ 384")

f = diff_findings("Груз сыручими и жидкими.", "Груз твердыми,сыпучими и жидкими.")
assert [(x.snippet, x.suggestion) for x in f] == [("сыручими", "твердыми,сыпучими")]
assert diff_findings("| 1 | а |\n|---|---|\n| 2 | б |", "1 а 2 б") == []       # разметка не шум
ins = diff_findings("один три", "один два три")
assert [(x.snippet, x.suggestion) for x in ins] == [("один", "два")]          # insert → сосед слева
dele = diff_findings("один два три", "один три")
assert [(x.snippet, x.suggestion) for x in dele] == [("два", None)]
long = diff_findings(" ".join(f"а{i}" for i in range(100)), "совсем другое")
assert len(long[0].snippet) <= SNIPPET_MAX and not long[0].snippet.endswith(" ")

# короткое расхождение — не находка (ПРАВКА #70)
assert diff_findings("код и цифра 5", "код в цифра 6") == []
assert [f.snippet for f in diff_findings("код или цифра", "код либо цифра")] == ["или"]
assert diff_findings("пункт № 12", "пункт 12") == []            # ни буквы, ни цифры
```

Замечание к тесту `сыручими`: `suggestion` — срез исходного токена второго
прогона, а он там склеен (`твердыми,сыпучими`), поэтому проверка — `in`, не `==`.

## Готово, когда

- `pytest -v` зелёный, фикстурные тесты не пропущены.
- Прогон на паре фикстур укладывается в 10 с (иначе — стоп и отчёт, не оптимизация наугад).
  Факт после ПРАВКИ #70 — 0,07 с.
- `git diff --stat`: только `ocr/diff.py`, `tests/test_ocr_diff.py`.
- В отчёте исполнителя: фактическое `len(fs)` на паре фикстур и 10 первых находок —
  человек по ним решает, годится ли уровень шума для `--verify`. Факт: 212 до
  ПРАВКИ #70, 65 после (PLACEHOLDER 6 закрыт).

## Коммит

```
ПРАВКА #65: сверка прогонов vlm/pipeline — находки low_confidence

ocr/diff.py: normalize (№, гомоглифы, регистр, пунктуация) и пословное
выравнивание difflib. Расхождение = находка с дословным фрагментом основного
прогона и вариантом второго. Текст не меняется, правый не назначается.
```
