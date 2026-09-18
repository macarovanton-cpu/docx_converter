# 04 — Детерминированный постпроцессор

**# ПРАВКА #63** (+ #68, #69, #72). Зависит от: 00. Независима от 01–03. Сеть и кэш не нужны.

## Цель

Цепочка чистых функций `markdown → markdown (+ находки)` без моделей. Чинит
только известные систематические артефакты vlm. Всё, что нельзя починить без
взгляда на изображение, остаётся в тексте как есть и уходит в находки.

## Трогать

- `ocr/__init__.py` — создать, если его ещё нет (содержимое — дословно из README)
- `ocr/postprocess.py` — создать
- `tests/test_ocr_postprocess.py` — создать

## Не трогать

`markdown_cleanup.py` (только вызывать), `convert.py`, всё из общего списка запретов.

## Интерфейсы (дословно)

```python
"""ПРАВКА #63: детерминированный постпроцессор markdown после OCR MinerU."""

from markdown_cleanup import cleanup_ocr_markdown
from ocr import Finding

# канонические написания; сравнение — по свёртке в латиницу и верхнему регистру
KNOWN_LATIN = ("PoE", "IP", "SNR", "ITV", "VIDEOMAX", "RS", "SFP", "SQL",
               "USB", "PDF", "DWG", "Ethernet", "Parsec")

HOMOGLYPHS_CYR_TO_LAT = {
    "А": "A", "В": "B", "Е": "E", "К": "K", "М": "M", "Н": "H", "О": "O",
    "Р": "P", "С": "C", "Т": "T", "Х": "X", "І": "I",
    "а": "a", "е": "e", "о": "o", "р": "p", "с": "c", "у": "y", "х": "x", "і": "i",
}


def fold_to_latin(token: str) -> str | None: ...
def fold_to_cyrillic(token: str) -> str | None: ...
def parse_pipe_tables(md: str) -> list[list[list[str]]]: ...

def html_tables_to_pipe(md: str) -> tuple[str, list[Finding]]: ...
def merge_split_tables(md: str) -> tuple[str, list[Finding]]: ...
def fix_numero(md: str) -> str: ...
def fix_degree(md: str) -> str: ...
def fix_list_glue(md: str) -> str: ...        # ПРАВКА #68
def fix_sentence_glue(md: str) -> str: ...    # ПРАВКА #72
def fix_mixed_alphabet(md: str) -> tuple[str, list[Finding]]: ...
def flag_translit(md: str) -> list[Finding]: ...
def flag_signature_block(md: str) -> list[Finding]: ...
def postprocess(md: str) -> tuple[str, list[Finding]]: ...
```

Все функции чистые и идемпотентные: `f(f(x)) == f(x)`. `page` во всех находках
этой спеки — `None` (страницу проставляет `build_report`, спека 05).
Серьёзность — строго по таблице в README.

### Помощники

- `fold_to_latin(token)`: каждую кириллическую букву заменить по
  `HOMOGLYPHS_CYR_TO_LAT`; встретилась кириллическая буква без пары → `None`.
  Латиница, цифры и прочее — без изменений.
- `fold_to_cyrillic(token)`: обратная таблица (строится из той же константы; для
  неоднозначности `I`/`i` берётся `І`/`і` → затем не используется для сравнения
  с русскими словами, только для инициалов и единиц). Латинская буква без пары → `None`.
- `parse_pipe_tables(md)`: таблицы → строки → ячейки (`strip`), строка-разделитель
  исключена, `\|` внутри ячейки раскрыт в `|`. Блок таблицы — подряд идущие строки,
  начинающиеся с `|`.

### 1. `html_tables_to_pipe`

- Ищет `<table …>…</table>` (без учёта регистра, многострочно). Разбор —
  `html.parser.HTMLParser` из стандартной библиотеки.
- Ячейка: `<td>`/`<th>`; текст — `html.unescape`, `<br>` → пробел, пробельные
  последовательности → один пробел, `strip`; `|` → `\|`.
- `colspan=N`: текст в первой ячейке, следующие `N-1` — пустые.
  `rowspan=N`: текст **повторяется** в той же колонке следующих `N-1` строк.
  На таблицу, где встретился хоть один span > 1, — одна находка `table_span`
  (`snippet` — текст первой такой ячейки, обрезанный до 80 символов).
- Выход: строки `| a | b | c |`, после первой строки — `|---|---|---|` по её ширине;
  пустая строка до и после таблицы. Короткие строки **не дополняются** пустыми
  ячейками — несовпадение ширины увидит валидатор (`table_row_cells`), а не
  спрячет постпроцессор.
- Не разобралось (вложенная `<table>`, нет ни одной строки, незакрытый тег) →
  таблица остаётся в тексте **как есть** + `html_table_unparsed`
  (`snippet` — первые 80 символов таблицы).

### 2. `merge_split_tables`

Правило переписано относительно PLAN: на фикстуре продолжение пункта приходит
таблицей в **2** колонки при основной в 3, «одинаковое число колонок» не работает.

Таблица B — продолжение предыдущей таблицы A, если между ними только пустые
строки **и** в первой строке B первая ячейка пуста и непустая ячейка ровно одна.
Тогда:

1. Сверху B снимаются все подряд идущие строки такого вида; их текст по порядку
   дописывается через пробел в **последнюю ячейку последней строки A**.
2. Остаток B: пуст → B исчезла. Ширина всех оставшихся строк равна ширине A →
   строки дописываются в A. Иначе → слияние **отменяется целиком**, обе таблицы
   остаются как были, находка `table_merge_failed` (`snippet` — первые 80 символов
   первой непустой ячейки B).
3. Успех → находка `table_merged`: `snippet` — первые 80 символов дописанного
   текста, `suggestion` — `f"Продолжение дописано в строку «{первая ячейка строки A}»; сверить со сканом"`.
4. Повторять, пока есть что сливать (на фикстуре — два слияния подряд).

Первая строка B с пустой первой ячейкой и **двумя и более** непустыми — это
обычная шапка матричной таблицы (`| | 2024 | 2025 |`): не сливать, находок не давать.

**ПРАВКА #69, хвосты строк внутри таблицы.** Вторым проходом, уже по готовым
таблицам: строка, у которой первая ячейка пуста, непустых **две и более**, а
ширина равна ширине предыдущей строки, — хвост этой предыдущей строки. Ячейки
дописываются **поячеечно** через пробел (`| | компьютер, | Для передачи |` →
первая ячейка к первой, третья к третьей), строка исчезает, находка
`table_merged` с тем же `suggestion`. Первая строка таблицы не сливается
никогда — дописывать некуда. Ровно одна непустая ячейка — **не хвост**: это
разделитель (`| I. Общие данные | | |`) или продолжение с прошлой страницы,
которое уже разобрал шаг выше. Другая ширина — тоже не хвост: лишние ячейки
пришлось бы выбросить, а тракт не глотает данные.

Правило нужно живому прогону: пункт 13 бейкоффа приезжает из MinerU двумя
строками (`13 | Программно-технический комплекс (персональный | В качестве…` и
`| компьютер, общесистемного… | Для передачи данных…`), и без склейки вторая
строка теряет номер — валидатор честно даёт на неё `table_empty_number`.

### 3. `fix_numero`

`No` / `No.` перед цифрой или перед `п/п` → `№ ` (с одним пробелом после):
`No1`→`№ 1`, `No 12`→`№ 12`, `No384-ФЗ`→`№ 384-ФЗ`, `Noп/п`→`№ п/п`.
Граница слова слева обязательна: `Nokia`, `Note`, `ПNo1` не трогать.

### 4. `fix_degree`

`$C^{\circ}$` и `$^{\circ}C$` (с любыми пробелами внутри `$…$`) → `°C`. Пробелы
вокруг: слева ровно один, между `°C` и следующим знаком `; . , )` — ни одного.
`+50  $C^{\circ}$ ;` → `+50 °C;`. Другие формулы `$…$` не трогать.

### 5. `fix_list_glue` (ПРАВКА #68)

`;-` и `:-` → `; -` и `: -`: OCR прилепляет маркер перечисления к концу
предыдущей фразы (`Разграничить права пользователей на:- администратор`).
Только эти два сочетания, ничего шире: `+-30кг` в допусках и `--` не трогаются,
разделитель pipe-таблицы `|:---|` остаётся собой. То же правило — правка 8
сборки `golden.md` (спека 00), поэтому в остаток метрики склейки не попадают.

### 5b. `fix_sentence_glue` (ПРАВКА #72)

Точка между предложениями, у которой OCR съел пробел:
`проектом.Предусмотреть` → `проектом. Предусмотреть`. Шаблон
`(?<=[^\W\d_]{2})\.(?=[А-ЯЁ])` — слева от точки **две буквы подряд** (алфавит
любой, цифры и `_` не считаются), справа — **заглавная кириллическая**.

Всё остальное не трогается, и это намеренно:

| Осталось как было | Почему |
|---|---|
| `И.М. Халиуллин`, `т.е.Х` | слева одна буква — инициал или сокращение |
| `1.3.4.Требования` | слева цифра — нумерация пункта |
| `Windows 8.1` | справа цифра |
| `в т.ч.дистрибутивы` | справа строчная — сокращение, а не новое предложение |
| `компания».Юридический` | слева кавычка, а не буква |

На `vlm.md` правило чинит два места (`Ethernet.Для`, `ПО.Работы`), на сыром
`vlm_raw.zip/full.md` — десять. То же правило — правка 9 сборки `golden.md`
(спека 00), поэтому в остаток метрики оно не попадает.

### 6. `fix_mixed_alphabet`

Токен — максимальная последовательность букв и цифр (дефис делит токены).

| Токен | Условие | Действие |
|---|---|---|
| есть и латиница, и кириллица | `fold_to_latin(t).upper()` совпал с элементом `KNOWN_LATIN` (в верхнем регистре) или с `IP\d{2}` | заменить на каноническое написание, без находки |
| есть и латиница, и кириллица | иначе | не менять, `mixed_alphabet_unknown`, `snippet` = токен, `suggestion` = `fold_to_latin(t)` или `fold_to_cyrillic(t)` — какой не `None` |
| только кириллица, длина ≥ 3 | `fold_to_latin(t).upper()` ∈ `KNOWN_LATIN` | заменить (`РоЕ` → `PoE`) |
| прочее | | не трогать |

Цифры алфавитом не считаются: `12В`, `2шт`, `Ст3` — не смешанные.

### 7. `flag_translit` — только пометка

Строка вида «инициалы + фамилия»:
`^\s*[A-Za-zА-ЯЁ]\.\s?[A-Za-zА-ЯЁ]{1,3}\.\s+[A-Za-zА-Яа-яЁё-]+\s*$`,
в которой есть хотя бы одна латинская буква → `translit_suspect`. `snippet` —
строка без краевых пробелов; `suggestion` — `fold_to_cyrillic(строка)`, если все
латинские буквы свернулись, иначе `None`. **ПРАВКА #70:** свёртка дала `І`/`і`
(украинская буква из латинской `I`) — `suggestion=None`. `A.II.` и `A.III.` —
это разобранная на палки `Ш`, а «А.ІІІ.» в инициалах заведомо не то; гадать
тракт не имеет права. Текст не меняется. Детектор намеренно
узкий — только ФИО с инициалами; общий поиск транслита по тексту без словаря
даёт шум на `Ethernet`, `Windows`, `Parsec`.

### 8. `flag_signature_block` — только пометка

Три и более **подряд идущих абзаца** (блоки через пустую строку), каждый из
которых — «инициалы + фамилия» по шаблону из п. 7 (в любом алфавите) → одна
находка `signature_block` на серию. `snippet` — первый абзац серии;
`suggestion` — `"Должности и ФИО идут отдельными списками — сопоставить по скану"`.

### 9. `postprocess`

Порядок: 1 → 2 → 3 → 4 → 5 → 5b → 6 → 7 → 8 → `cleanup_ocr_markdown`. Находки — в
порядке получения. `cleanup_ocr_markdown` вызывается последним и как есть.

## Приёмочные тесты (`tests/test_ocr_postprocess.py`)

```python
vlm, golden, pipeline = (read_fixture(n) for n in ("vlm.md", "golden.md", "pipeline.md"))
out, findings = postprocess(vlm)
rules = [f.rule for f in findings]

# метрика: не хуже эталонной и детерминированное ушло
assert count_diffs(out, golden) <= VLM_TO_GOLDEN_DIFFS          # из tests/test_ocr_fixtures.py
assert count_diffs(out, golden) <= 6     # остаток: сыручими, IR-, 3 опкода на 5 подписях,
                                         # «(персональныйкомпьютер,» (правка 10, ПРАВКА #72)
assert postprocess(out)[0] == out                               # идемпотентность

# починено
assert not [t for t in text_tokens(out) if t.startswith("No")]
assert out.count("№ ") == 4 and "№ п/п" in out and "№ 384-ФЗ" in out and "№ 7-ФЗ" in out
assert "$" not in out and "+50 °C;" in out and "+35 °C." in out
assert "РоЕ" not in out and out.count("PoE") == 2

# НЕ починено (нельзя без скана) — осталось дословно
assert "сыручими" in out and "сыпучими" not in out
assert "IR-камерами" in out
assert "A.II. Taipov" in out and "P.P. Hypeeb" in out
assert "(персональныйкомпьютер," in out      # ПРАВКА #72: тракт слова не режет

# починено правилом 5b (ПРАВКА #72)
assert "RS-485,Ethernet. Для" in out and "и ПО. Работы," in out
assert "в т.ч.дистрибутивы" in out and "Приложение 1.План" in out

# таблица
assert "<table" not in out and "<td" not in out
tables = parse_pipe_tables(out)
assert len(tables) == 1
assert len(tables[0]) == 53                                     # 55 <tr> минус 2 строки-продолжения
assert all(len(row) == 3 for row in tables[0])
assert tables[0][0][0] == "№ п/п"
assert tables[0][1] == ["I. Общие данные", "", ""]
assert "" not in [row[0] for row in tables[0]]
row25 = next(r for r in tables[0] if r[0] == "25")
assert "пуско-наладочные работы; - первичная поверка" in row25[2]
assert row25[2].rstrip().endswith("до сети ИТСО.")
assert tables[0][-1][0] == "43"

# находки
assert rules.count("table_merged") == 2
assert rules.count("table_span") == 1
assert rules.count("signature_block") == 1
assert rules.count("translit_suspect") == 5
# единственный смешанный токен vlm — склейка «bzdk@list.ruОГРН»; «Noп» уже снят шагом 3
assert [f.snippet for f in findings if f.rule == "mixed_alphabet_unknown"] == ["ruОГРН"]
assert "html_table_unparsed" not in rules and "table_merge_failed" not in rules
assert all(f.page is None for f in findings)
assert all(f.severity in SEVERITIES for f in findings)
tr = {f.snippet: f for f in findings if f.rule == "translit_suspect"}
assert tr["A.P. Сиражитдинов"].suggestion == "А.Р. Сиражитдинов"
assert tr["E.K. Кустова"].suggestion == "Е.К. Кустова"
assert tr["P.P. Hypeeb"].suggestion is None                     # 'b', 'y' без пары — не гадать
assert tr["A.II. Taipov"].suggestion is None                    # ПРАВКА #70: «І» в инициалах
assert tr["A.III. Ямалов"].suggestion is None
assert all(f.severity == "critical" for f in tr.values())
for f in findings:
    assert f.snippet in out or f.rule in ("table_merged", "table_span")

# живой сырой прогон (vlm_raw.zip, ПРАВКА #69): пункт 13 собран в одну строку
raw = zipfile.ZipFile(require_fixture("vlm_raw.zip")).read("full.md").decode("utf-8")
r_out, r_findings = postprocess(raw)
table = parse_pipe_tables(r_out)[0]
assert len([row for row in table if row[0] == "13"]) == 1
assert "(персональный компьютер," in [r for r in table if r[0] == "13"][0][1]
assert "" not in [row[0] for row in table]
assert [f.rule for f in r_findings].count("table_merged") == 3
assert "table_empty_number" not in [f.rule for f in validate(r_out)]

# шумный прогон не роняет цепочку и ничего не глотает молча
p_out, p_findings = postprocess(pipeline)
assert "mixed_alphabet_unknown" in [f.rule for f in p_findings]
assert len(text_tokens(p_out)) >= len(text_tokens(pipeline)) - 5 # текст не потерян
```

Синтетика — по одному тесту на функцию:

```python
# html_tables_to_pipe
md, f = html_tables_to_pipe('<table><tr><td rowspan="2">A</td><td>1</td></tr><tr><td>2</td></tr></table>')
assert parse_pipe_tables(md) == [[["A", "1"], ["A", "2"]]] and f[0].rule == "table_span"
md, f = html_tables_to_pipe("<table><tr><td>a|b</td><td>x<br>y &amp; z</td></tr></table>")
assert "| a\\|b | x y & z |" in md
bad = "<table><tr><td><table><tr><td>in</td></tr></table></td></tr></table>"
md, f = html_tables_to_pipe(bad)
assert md == bad and [x.rule for x in f] == ["html_table_unparsed"]
md, f = html_tables_to_pipe("<table><tr><td>1</td><td>2</td></tr><tr><td>3</td></tr></table>")
assert [len(r) for r in parse_pipe_tables(md)[0]] == [2, 1]      # не дополняем

# merge_split_tables
a = "| № | Что |\n|---|---|\n| 1 | начало |"
assert parse_pipe_tables(merge_split_tables(a + "\n\n|  | конец |\n|---|---|")[0]) == \
    [[["№", "Что"], ["1", "начало конец"]]]
matrix = a + "\n\n|  | 2024 | 2025 |\n|---|---|---|\n| план | 1 | 2 |"
assert merge_split_tables(matrix) == (matrix, [])               # шапка матрицы — не продолжение
wide = a + "\n\n|  | хвост |\n|---|---|\n| 2 | x | лишняя |"
md, f = merge_split_tables(wide)
assert md == wide and [x.rule for x in f] == ["table_merge_failed"]
assert merge_split_tables(a + "\n\nАбзац.\n\n|  | конец |\n|---|---|")[1] == []   # между ними текст

# хвосты строк внутри таблицы (ПРАВКА #69)
head = "| № | A | B |\n|---|---|---|\n"
md, f = merge_split_tables(head + "| 1 | a | b |\n|  | хвост | ещё |")
assert parse_pipe_tables(md) == [[["№", "A", "B"], ["1", "a хвост", "b ещё"]]]
assert [x.rule for x in f] == ["table_merged"]
one = head + "| 1 | a | b |\n| I. Общие данные |  |  |\n|  |  | одна |"
assert merge_split_tables(one) == (one, [])          # одна непустая — не хвост
first = "|  | x | y |\n|---|---|---|\n| 1 | a | b |"
assert merge_split_tables(first) == (first, [])      # первой строке некуда дописывать
wide2 = head + "| 1 | a | b |\n|  | x | y | лишняя |"
assert merge_split_tables(wide2) == (wide2, [])      # ширина не та — ячейку не выбросим

# fix_list_glue
assert fix_list_glue("на:- один;- два") == "на: - один; - два"
assert fix_list_glue("+-30кг, 60 т. +- 50кг") == "+-30кг, 60 т. +- 50кг"
assert fix_list_glue("|:---|---:|") == "|:---|---:|"

# fix_sentence_glue (ПРАВКА #72)
assert fix_sentence_glue("проектом.Предусмотреть") == "проектом. Предусмотреть"
assert fix_sentence_glue("Ethernet.Для и ПО.Работы") == "Ethernet. Для и ПО. Работы"
for same in ("И.М. Халиуллин", "т.е.Х", "1.3.4.Требования", "Windows 8.1",
             "в т.ч.дистрибутивы", "компания».Юридический", "A.II. Taipov"):
    assert fix_sentence_glue(same) == same

# fix_numero / fix_degree
assert fix_numero("No1, No 12, No.7, Noп/п") == "№ 1, № 12, № 7, № п/п"
assert fix_numero("Nokia Note ПNo1") == "Nokia Note ПNo1"
assert fix_degree("от +5 до +35  $C^{\\circ}$ .") == "от +5 до +35 °C."
assert fix_degree("$x^2$") == "$x^2$"

# fix_mixed_alphabet
assert fix_mixed_alphabet("функции РоЕ; IР68; 12В; Ст3")[0] == "функции PoE; IP68; 12В; Ст3"
md, f = fix_mixed_alphabet("ЦСмц и Sм")
assert md == "ЦСмц и Sм" and [x.snippet for x in f] == ["Sм"]
assert fix_mixed_alphabet("РОЕ")[0] == "PoE" and fix_mixed_alphabet("Ре")[0] == "Ре"

# подписи
assert flag_translit("Текст A.P. Сиражитдинов в строке") == []   # не отдельной строкой
assert flag_translit("А.Р. Сиражитдинов") == []                  # латиницы нет
assert flag_signature_block("А.А. Иванов\n\nБ.Б. Петров") == []  # двух мало
assert len(flag_signature_block("А.А. Иванов\n\nБ.Б. Петров\n\nВ.В. Сидоров")) == 1
```

## Готово, когда

- `pytest -v` зелёный, фикстурные тесты не пропущены.
- `python -c "import ocr.postprocess"` не тянет `streamlit`, `requests`, `pdf_core`.
- `git diff --stat`: только `ocr/__init__.py` (если создан здесь), `ocr/postprocess.py`,
  `tests/test_ocr_postprocess.py`.
- В отчёте исполнителя: фактическое `count_diffs(out, golden)` и список оставшихся опкодов.

## Коммит

```
ПРАВКА #63: детерминированный постпроцессор OCR-markdown

ocr/postprocess.py: HTML-таблицы → pipe (rowspan/colspan), склейка таблиц через
разрыв страницы, No→№, LaTeX-градус→°C, гомоглифы по словарю (РоЕ→PoE).
Транслит в ФИО и разорванный блок подписей только помечаются. Неразобранное
остаётся в тексте и уходит в находки. В конце — cleanup_ocr_markdown.
```
