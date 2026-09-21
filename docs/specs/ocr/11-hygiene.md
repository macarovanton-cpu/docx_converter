# 11 — Гигиена после этапа A

**# ПРАВКА #84** (одна). Зависит от: 10. Сеть не нужна: факты из документации MinerU уже
собраны и записаны ниже (сверка 21.09.2026), исполнитель в облако не ходит.

## Цель

Закрыть долги, накопившиеся за #73–#83, до начала этапа B:

1. Ошибка `apply_errors` при 0 или 2 вхождениях называет **номер строки `errors.txt`** —
   сейчас человек ищет строку по тексту «было» глазами.
2. Ретроспективная спека **08** про MinerU в UI (ПРАВКА #73) — написать по коду `app.py`
   задним числом. Номер 08 за ней зарезервирован спекой 09.
3. `CLAUDE.md` приведён к коду по правкам #73–#83 (закрытый список расхождений ниже).
4. `requirements.txt`: `streamlit` закреплён на `1.58.0`.
5. PLACEHOLDER-ы MinerU 1 и 2 закрыты **в документах** по документации API.
6. `docs/PLAN_OCR.md` обновлён по факту этапа A.

Единственный продуктовый файл с изменением кода — `ocr/board.py`. Тракт не меняется:
выход `ingest` на шести фикстурах, эталоны, sha и пороги остаются байт в байт.

## Шаг 0 — условия входа (иначе стоп)

1. Спека 10 закоммичена, `pytest -q` зелёный, `python -X utf8 -m ocr.board` отрабатывает оффлайн.
2. `grep -rn "ПРАВКА #8[4-9]" --include=*.py .` (без `.venv`) пуст: #84 свободен.
   Иначе — стоп, номера в спеках 11–13 сдвигает человек.
3. `.venv/Scripts/python.exe -c "import streamlit; print(streamlit.__version__)"` даёт `1.58.0`.
   Другая версия — стоп: закрепляется то, на чём приложение проверено, а не то, что написано в спеке.

## Трогать

- `ocr/board.py` — `parse_errors_numbered`, параметр `lines` у `apply_errors`, вызов в `write_goldens` (#84)
- `tests/test_ocr_golden.py` — новые assert-ы (ниже)
- `requirements.txt` — **одна** строка: `streamlit` → `streamlit==1.58.0`
- `docs/specs/ocr/08-ui-mineru.md` — создать (ретроспектива, оглавление ниже)
- `docs/specs/ocr/README.md` — таблица спек, список правил, PLACEHOLDER-ы 1 и 2, оговорка про `app.py`
- `docs/specs/ocr/02-mineru-provider.md` — только два PLACEHOLDER-абзаца (квота, `page_ranges`): приписать результат сверки
- `CLAUDE.md` — закрытый список ниже
- `docs/PLAN_OCR.md` — закрытый список ниже
- `docs/docx_converter_docs_sync.md` — `### ПРАВКА #84`, пункты 1–3 раздела «Расхождения → CLAUDE.md» пометить исправленными, снимок «следующий свободный номер — #85»

## Не трогать

`app.py`, `ocr/mineru_provider.py` (**решение человека 21.09.2026**: PLACEHOLDER-ы закрываются
только в документах; `QUOTA_CODES` и комментарий у `page_ranges` в коде остаются как есть до
отдельной правки провайдера), остальные модули `ocr/*`, `pdf_core.py`, `file_converter.py`,
`<stem>.errors.txt`, `*.golden.md`, `GOLDEN_SHA256`, `THRESHOLDS`, `tests/test_app_fixes.py`,
`README.md` корня, всё из общего списка запретов. В `requirements.txt` — ничего, кроме
названной строки: остальные зависимости не закрепляются, новые не добавляются.

Разрешение на `requirements.txt` (дословно, человек, 21.09.2026): «requirements.txt: закрепить
streamlit==1.58.0 — разрешено человеком 21.09.2026». `CLAUDE.md` требует на этот файл явного
запроса — это он.

## Интерфейсы (дословно)

### `ocr/board.py` (#84)

Сигнатуры спеки 10 не ломаются: `parse_errors` возвращает то же, `apply_errors` получает
только необязательный именованный параметр.

```python
def parse_errors_numbered(text: str) -> list[tuple[int, tuple[str, str, str]]]: ...
    # ПРАВКА #84: (номер строки файла, (страница, было, надо)); тело — прежнее тело parse_errors


def parse_errors(text: str) -> list[tuple[str, str, str]]:
    return [error for _, error in parse_errors_numbered(text)]


def apply_errors(md: str, errors: list[tuple[str, str, str]], *,
                 lines: "list[int] | None" = None) -> str: ...
```

- `apply_errors`: `lines is None` — поведение и текст ошибки прежние
  (`«{было}»: вхождений {n}, нужно ровно 1`). `lines` задан — обязан быть той же длины, что
  `errors` (иначе `ValueError("lines: … номеров на … правок")`), и сообщение становится
  `f"строка {номер} errors.txt: «{было}»: вхождений {n}, нужно ровно 1"`.
  Подстрока `вхождений {n}` сохраняется — существующие `match=` зелёные без правок.
- `write_goldens`: читает `parse_errors_numbered`, зовёт `apply_errors(…, lines=…)`;
  `ValueError` перехватывает и поднимает заново с именем файла впереди:
  `ValueError(f"{errors.name}: {exc}") from exc`. Остальное — без изменений: сначала собрать
  все пять, потом писать.
- Нумерация строк — как у `parse_errors` сейчас: `enumerate(text.splitlines(), 1)`, с учётом
  шапки и комментариев, то есть номер совпадает с тем, что человек видит в редакторе.

### `08-ui-mineru.md` — обязательное оглавление

Шапка: `**# ПРАВКА #73.** Ретроспектива: написана 2026-09 по коду app.py, после исполнения.
Этап 8 выполнен по явному запросу человека без спеки; этот файл фиксирует факт, а не задаёт работу.`
Разделы — тот же скелет, в прошедшем времени; «Коммит» — хэш и заголовок из `git log`, не выдуманный текст.

Факты, которые обязаны в неё попасть (каждый перепроверить по `app.py`, номера строк — на день написания):

| Что | Где |
|---|---|
| `ocr_mode`: `off` / `auto` / `mineru`; `auto` предлагается только при найденных бинарниках | `_ocr_mode_options`, `_ocrmypdf_available` (`st.cache_resource`), `_OCR_MODE_LABELS` |
| ключ: `st.secrets["MINERU_API_KEY"]`, затем окружение; `FileNotFoundError`/`KeyError` → окружение | `_mineru_api_key` |
| диапазон страниц: вырезка `pypdf` до отправки, страница вне PDF → `ValueError`; номера страниц в находках — от вырезки | `_pdf_page_subset` |
| прогресс без колбэка в тракте: подкласс `MineruProvider` метит `fetch_raw_zip`, инжектируемый `sleep` — «ожидание»; `status=None` допустим | `_mineru_provider_factory` |
| вызов: `run_pipeline(pdf_bytes, source_name, work_dir=<tmp>, verify, annotate, cache=LocalCache(_OCR_CACHE_ROOT), provider_factory)` — только для `ext == "pdf"` | `_convert_uploaded_file` |
| ошибки: текст исключения в `result["error"]`, у `MineruAuthError` — плюс `_MINERU_KEY_HINT`; трейсбека нет | там же |
| результат: ключ `report` (`None` вне `mineru` и при ошибке) | там же |
| UI: чекбоксы «Сверка вторым прогоном» (по умолчанию выкл.) и «Пометки в тексте» (вкл.), `st.status` на файл, предупреждение об отсутствии ключа | `render_files_to_markdown_mode` |
| находки: сводка, таблица, `low_confidence` отдельным свёрнутым блоком; подпись про пометки `!! ПРОВЕРИТЬ` | `_split_findings`, `_findings_table`, `_render_ocr_report` |
| `report.json` — отдельная кнопка, в общий ZIP не входит | цикл результатов |
| кэш `.cache/ocr` рядом с `app.py`, на Streamlit Cloud эфемерный | `_OCR_CACHE_ROOT` |
| тесты: 11 функций блока «ПРАВКА #73» в `tests/test_app_fixes.py`, фейковый провайдер, без сети | — |

Раздел «Известные ограничения» ретроспективы: DOCX/XLSX в режиме `mineru` шли голым
MarkItDown без отчёта (чинит спека 12); живой прогон MinerU из UI правкой #73 не проверялся;
`sleep` зовётся и в ретраях загрузки — метка «ожидание» может мигнуть раньше времени (`# ponytail:` в коде).

### `CLAUDE.md` — закрытый список правок

Сверено с кодом 21.09.2026. Править только перечисленное; файл остаётся в пределах ~250 строк.

1. Architecture, `convert.py`: «1316 lines» → фактическое `wc -l` на день исполнения (21.09.2026 — 1630);
   сигнатура — с `doc_style='pz'`, как в «Forbidden patterns».
2. Architecture, `app.py`: добавить режим «Файлы → Markdown» и три значения `ocr_mode`; MinerU —
   через `ocr.cli.run_pipeline` (#73). **Эту фразу потом меняет спека 12, здесь — по коду как есть.**
3. Architecture, `pdf_core.py`: протокол `OcrProvider.ocr_pdf(pdf_bytes, page_range) -> OcrResult` (#60),
   датаклассы `PageInfo` / `OcrResult`; две реализации — `OcrmypdfProvider` и `ocr.mineru_provider.MineruProvider`.
   Фразу «A second (cloud vision) provider is a planned separate PR» убрать.
4. `ocr/cache.py`: «until stage 8» → «не реализован» (этап 8 прошёл без Drive-кэша).
5. Абзац «OCR pipeline»: рядом с цепочкой `auto` — цепочка `mineru` (`_pdf_page_subset` → `run_pipeline`).
6. «Numbered edits convention»: после «#61–#66» дописать — #67–#70, #72, #74–#80 — исправления тракта в
   `ocr/*` (по файлам не расписывать: маркеры стоят в коде), #71 — `conftest.py`, #73 — `app.py`, #84 — `ocr/board.py`.
7. «Known production limitation»: `auto` на Streamlit Cloud **не предлагается** (бинарников нет,
   `_ocr_mode_options`), рабочий OCR-путь прода — MinerU. Запрет на `ocrmypdf` в `requirements.txt` и `packages.txt` остаётся.
8. «Secrets»: `MINERU_API_KEY` — верхнеуровневый ключ `secrets.toml` либо переменная окружения.
9. «OCR CLI»: одной строкой — ошибка `--golden` называет файл и номер строки `errors.txt` (#84).

Расхождение, найденное сверх списка, — не чинить молча: дописать в docs_sync, раздел «Расхождения».

### `docs/specs/ocr/README.md`

- Таблица спек: строки `08-ui-mineru.md` (`app.py`, **#73**, ретроспектива) и `11-hygiene.md` (**#84**).
  Абзац «Этапы 8 и 9 — вне автономного прогона, спек нет» → этап 8 выполнен (#73, спека 08 задним числом).
- «Запрещено», пункт `app.py`: оговорка «кроме спек, где он назван в „Трогать“ (08 — факт, 12)».
- Список правил — три строки, которые есть в коде и которых нет в таблице (серьёзность взять из кода, не из спеки):
  `recovered_block` — info, 04 (#74); `code_digits_glued` — warning; `org_name_variant` — warning.
  «Кто выдаёт» и «смысл» — по докстрингу места, где правило рождается.
- PLACEHOLDER 1: зачеркнуть, «закрыт по доке 21.09.2026: код `-60018`; в коде `QUOTA_CODES` пуст до
  отдельной правки провайдера». PLACEHOLDER 2: зачеркнуть, «закрыт по доке: `page_ranges` — поле элемента `files[]`, код верен».

### Факты сверки с документацией MinerU (21.09.2026)

Источник — `https://mineru.net/apiManage/docs`, прочитано агентом-исследователем, руками не перепроверено.

| Вопрос | Ответ доки | Статус |
|---|---|---|
| место `page_ranges` в `POST /api/v4/file-urls/batch` | поле элемента `files[]` (`{"name", "data_id", "is_ocr", "page_ranges"}`); на верхнем уровне оно только у одиночного `extract/task` | **сверено**, совпадает с `mineru_provider.py:106-109` |
| код «квота исчерпана» | `-60018` — «Daily extract task limit reached»; `-60019` — квота HTML-извлечения, к PDF не относится | **сверено по доке**, на живом ответе не наблюдалось |
| лимит страниц | текст лимитов — 600, строка ошибки `-60006` — 200 | **не сверено**: дока противоречит себе; `MAX_PAGES = 200` остаётся как консервативное |
| код «слишком частые запросы» | для API с токеном не документирован | **не сверено** |

Вход для будущей правки провайдера (не эта спека): `QUOTA_CODES = frozenset({"-60018"})`, снять
два комментария `# PLACEHOLDER`, тест на `MineruQuotaError`.

### `docs/PLAN_OCR.md` — закрытый список

1. Строка статуса: «этапы 0–8 и этап A выполнены (#60–#83); этап B — в работе, спека 13».
2. «600 страниц» (этап 2) → «200; дока MinerU противоречива (600 в тексте лимитов, 200 в `-60006`), не сверено».
3. Этап 8: «выполнен, ПРАВКА #73, спека 08 задним числом; Drive-кэш не сделан».
4. Новый раздел «Этап A — единый вход и эталоны (факт)»: `ocr.ingest` и четыре маршрута; `ocr.board`;
   шесть фикстур, пять `errors.txt`, эталоны; остаток **63** опкода (13 + 20 + 8 + 13 + 8 + 1) — таблицей из docs_sync, без пересказа.
5. Этап 9 → «Этап B»: первый шаг — измерение (спека 13), в тракт не подключается до чисел.
6. «Ограничения проекта», пункт про `app.py`: вынос Drive-кода не состоялся; `app.py` правится только спеками, где назван.

Исторический текст этапов 0–7 не переписывать.

## Приёмочные тесты

### `tests/test_ocr_golden.py` — дополнение

```python
from ocr.board import parse_errors_numbered

text = "# шапка\n\n3 | а | б\n4 | нет такого | в\n"
numbered = parse_errors_numbered(text)
assert numbered == [(3, ("3", "а", "б")), (4, ("4", "нет такого", "в"))]
assert parse_errors(text) == [error for _, error in numbered]           # обёртка, поведение прежнее

lines = [n for n, _ in numbered]
errors = [e for _, e in numbered]
with pytest.raises(ValueError, match=r"строка 4 errors\.txt.*вхождений 0"):
    apply_errors("а", errors, lines=lines)
with pytest.raises(ValueError, match=r"строка 3 errors\.txt.*вхождений 2"):
    apply_errors("а а", errors, lines=lines)
with pytest.raises(ValueError, match="lines"):
    apply_errors("а", errors, lines=[3])                                  # длины не совпали
with pytest.raises(ValueError, match="вхождений 0") as info:
    apply_errors("текст", [("1", "нет такого", "б")])                     # без lines — как было
assert "строка" not in str(info.value)
```

В `test_golden_flag_keeps_manual_work` — хвост: в копию `xlsx1.errors.txt` дописать строку
`- | такого текста в черновике нет | x`, `main(["--golden", "--force"]) == 1`, в `stderr` есть и
`xlsx1.errors.txt`, и `строка <номер дописанной строки>`; эталоны в `fixtures` при этом не изменились
(sha всех пяти прежние — отказ не оставляет половину).

`test_goldens` зелёный без правок — доказательство, что тракт и эталоны не тронуты.

## Готово, когда

- `pytest -v` зелёный целиком; `python -X utf8 -m ocr.board` — те же шесть строк, те же `diffs`/`thr`.
- `git diff --stat`: из продуктового кода только `ocr/board.py`; `requirements.txt` — одна строка;
  нет `app.py`, `ocr/mineru_provider.py`, `tests/test_app_fixes.py`.
- `docs/specs/ocr/08-ui-mineru.md` существует, каждая строка таблицы фактов перепроверена по `app.py`.
- `grep -n "1316\|ocr_pdf_to_markdown\|planned separate PR" CLAUDE.md` пуст.
- В docs_sync — `### ПРАВКА #84` и снимок «следующий свободный номер — **#85**».

## PLACEHOLDER-ы

1. Лимит страниц MinerU (200 / 600) — не сверено, см. таблицу. Снимается живым прогоном файла на 201+ страниц; не этой спекой.
2. Факты доки MinerU прочитаны агентом, не человеком: при первом живом отказе по квоте сверить код ответа с `-60018`.
3. `streamlit==1.58.0` закрепляет только прямую зависимость; транзитивные плавают. Полный lock-файл — отдельное решение человека.
4. Номера строк `app.py` в спеке 08 — на день написания; якорь — имя функции.

## Коммит

```
ПРАВКА #84: гигиена после этапа A

ocr/board.py: ошибка apply_errors при 0/2 вхождениях называет номер строки
errors.txt (parse_errors_numbered, параметр lines; сигнатуры спеки 10 целы),
--golden добавляет имя файла. requirements.txt: streamlit==1.58.0 (разрешено
человеком 21.09.2026). docs: ретроспективная спека 08 про MinerU в UI (#73);
CLAUDE.md сверен с кодом по #73–#83; PLACEHOLDER-ы MinerU 1 и 2 закрыты по
документации API (код квоты -60018, page_ranges в элементе files) — в коде
провайдера константа до отдельной правки; PLAN_OCR.md обновлён по факту этапа A.
Тракт, эталоны, sha и пороги не менялись.
```
