# docx_converter

Веб-приложение на Streamlit для двух рабочих сценариев:

- конвертация Markdown в DOCX-документы в фирменном стиле ООО «ТПК «Тензосила»;
- пакетная конвертация PDF/DOCX/XLSX/PPTX в Markdown через Microsoft MarkItDown.

Используется для коммерческих предложений, пояснительных записок, сопроводительных писем и извлечения Markdown из типовых документов для дальнейшего анализа.

🌐 **Прод:** https://docxconverter-5n8ntdurz5scqneuvithfb.streamlit.app/

## Возможности

### Markdown -> DOCX

- Три способа ввода: вставить Markdown текст, загрузить `.md` файл, загрузить готовый DOCX/PDF/TXT с авто-распознаванием структуры
- Два типа документа на выходе: письмо/коммерческое предложение, пояснительная записка
- Шаблон с фирменным хедером (логотип, реквизиты компании, ОГРН, ИНН) подгружается с Google Drive
- Кастомный Markdown-синтаксис: callout-врезки `!! ... !!`, блок реквизитов `**Кому:**`, цитаты, плейсхолдеры фото, заголовки `####`/`#####`/`######` (H4–H6), автоссылки `<https://...>` в угловых скобках и голые URL прямо в тексте
- Автоматические ✓/✗ перед ячейками, начинающимися с «Да»/«Нет»/«Отсутствует» — иконка добавляется перед текстом, сам текст ячейки не обрезается (например, «Да — 36 месяцев» → «✓ Да — 36 месяцев»)

### Файлы -> Markdown

- Пакетная загрузка нескольких файлов
- Поддерживаемые форматы: `.pdf`, `.docx`, `.xlsx`, `.pptx`
- Конвертация в Markdown через Microsoft MarkItDown
- Для PDF можно указать диапазон страниц
- Для PDF показывается диагностика: количество страниц, наличие текстового слоя, предупреждение по image-only страницам
- Три режима OCR (переключатель «OCR mode»):
  - **Без OCR** (`off`) — конвертация через MarkItDown как есть.
  - **OCRmyPDF (локально)** (`auto`) — PDF без текстового слоя (с учётом выбранного диапазона страниц) прогоняется через `ocrmypdf`, после чего конвертируется в Markdown. Вариант показывается, только если на машине найдены `ocrmypdf`, Tesseract и Ghostscript; на Streamlit Cloud их нет, и он там не предлагается.
  - **MinerU (облако)** (`mineru`) — PDF уходит в облако mineru.net и проходит весь OCR-тракт (`ocr.cli.run_pipeline`: распознавание → постобработка → проверки). Работает и на Streamlit Cloud. Нужен ключ `MINERU_API_KEY` — в `st.secrets` (на проде — Secrets приложения) или в переменной окружения; без ключа вариант в списке остаётся, но при выборе показывается предупреждение. При заданном диапазоне в облако уходят только выбранные страницы, и номера страниц в находках считаются от этой вырезки. Лимиты MinerU: 200 МБ и 200 страниц.
- Настройки режима MinerU: «Сверка вторым прогоном» (выключена по умолчанию; второй прогон другой моделью, расхождения → находки `low_confidence`; удваивает время и расход квоты) и «Пометки в тексте» (включены по умолчанию; находки вставляются в Markdown как `!! ПРОВЕРИТЬ: … !!` — перед конвертацией в DOCX их нужно снять).
- Под результатом MinerU — блок находок: сводка critical / warning / info, список (правило, страница, фрагмент, предложение), `low_confidence` — отдельным свёрнутым списком. `report.json` скачивается кнопкой рядом с `.md`.
- Сырые ответы MinerU кэшируются в `.cache/ocr`: повторный прогон того же файла не тратит квоту. На Streamlit Cloud кэш эфемерный — живёт до перезапуска контейнера.
- Для каждого успешного результата доступно скачивание отдельного `.md`
- Все успешные `.md` можно скачать одним ZIP-архивом

Примеры диапазонов страниц для PDF:

```text
all
1-3
1-3, 7
1-3, 7, 10-12
```

## Запуск локально

```bash
pip install -r requirements.txt
streamlit run app.py
```

Или через Codespaces — `.devcontainer/devcontainer.json` запускает приложение автоматически после attach.

Для работы с шаблоном с Google Drive нужен `.streamlit/secrets.toml` с секцией `[gcp_service_account]`. Без него приложение использует локальный шаблон (путь задаётся в `DOC_TYPES` в `app.py`).

Важно: режим `Файлы -> Markdown` не требует доступа к Google Drive. Доступ к шаблону нужен только для генерации DOCX в режиме `Markdown -> DOCX`.

## Ограничения

- Выравнивание колонок таблиц из Markdown-сепаратора (`:----`, `:---:`, `----:`) не реализовано — строка-сепаратор просто отфильтровывается, все колонки рендерятся с выравниванием по умолчанию.
- OCR подключён к UI режима `Файлы -> Markdown` двумя вариантами. `auto`: для PDF без текстового слоя (с учётом выбранного диапазона страниц) выполняется `ocrmypdf --skip-text --deskew --rotate-pages -l rus+eng`, после чего MarkItDown извлекает текст из OCR-слоя; доступен только локально, см. «Известное ограничение прода» ниже. `mineru`: облачный OCR-тракт с отчётом о находках; документ при этом покидает машину.
- В режиме MinerU `report.json` в общий ZIP и в объединённый Markdown не входит — скачивается отдельной кнопкой у каждого файла.
- Диапазоны страниц сейчас поддержаны только для PDF.
- Для DOCX/XLSX/PPTX выполняется конвертация всего файла. Если для этих форматов указан page range, приложение покажет ограничение.
- ZIP-архив в режиме `Файлы -> Markdown` содержит только успешно сконвертированные `.md`; результаты с ошибками не включаются.

## Известное ограничение прода

OCR-режим `auto` (OCRmyPDF) реализован, но `ocrmypdf` **не** добавлен в `requirements.txt`, и файла `packages.txt` нет. На Streamlit Community Cloud из-за этого недоступны системные Tesseract/Ghostscript. С ПРАВКИ #73 приложение проверяет бинарники (один раз на процесс) и на проде вариант OCRmyPDF просто не показывает (раньше он был в списке и падал). Рабочий OCR-путь прода — **MinerU (облако)**: это чистый HTTP, системных пакетов не требует, нужен только `MINERU_API_KEY` в Secrets. Пакетирование `ocrmypdf` остаётся отдельной задачей; в `requirements.txt` / `packages.txt` ничего не добавляем.

## OCR CLI (инструмент для агента)

**Назначение:** тендерный PDF (в т.ч. скан) → Markdown со структурой + отчёт о сомнительных местах.
Смысл исходника не меняется: чинятся только известные артефакты OCR, остальное помечается.

**Вызов (из корня репозитория):**
`python -m ocr.cli ВХОД.(pdf|docx|xlsx) --out ПАПКА [--engine mineru|ocrmypdf] [--mode vlm|pipeline] [--verify] [--annotate] [--annotate-all] [--cache local|drive]`

**Вход (ПРАВКА #81):** `.pdf`, `.docx`, `.xlsx` — через `ocr.ingest.ingest`. PDF без текстового слоя хотя бы
на одной странице и текстовый PDF с таблицами идут в MinerU; текстовый PDF без таблиц, DOCX и XLSX —
в MarkItDown (без сети и ключа; `--verify` там — ошибка). Любой маршрут проходит `postprocess` + `validate`
и даёт тот же `report.json`.

**Табло качества (ПРАВКА #82):** `python -X utf8 -m ocr.board` — строго оффлайн, шесть фикстур из
`_test/fixtures/ocr/` (PDF — из `vlm_raw*.zip` через временный кэш, промах — ошибка, не облако).
Пишет `_test/quality_board.json`, черновики `_test/board/<имя>/out.md` + `report.json` и заготовки
`_test/fixtures/ocr/<имя>.errors.txt` (существующие не перезаписывает — в них ручной труд).
Код `0` — табло построено, `1` — исключение; критичные находки кода не меняют.

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

## Архитектура

Семь Python-модулей и пакет `ocr/`:

- **`app.py`** — Streamlit UI. Содержит режимы `Markdown -> DOCX` и `Файлы -> Markdown`. В первом режиме загружает `.docx` шаблон с Google Drive через service account, вызывает `convert_md_to_docx`, отдаёт результат на скачивание. Во втором режиме принимает несколько файлов, вызывает `pdf_core`/MarkItDown-слой и отдаёт `.md`/ZIP на скачивание.
- **`convert.py`** — ядро Markdown → DOCX (1316 строк). Единственная публичная функция: `convert_md_to_docx(md_text, output_filename, template_path=None, images=None)`.
- **`file_converter.py`** — обратная сторона: DOCX/PDF/TXT → Markdown для предзаполнения редактора, а также отдельный MarkItDown-слой (`convert_with_markitdown`) для PDF/DOCX/XLSX/PPTX → Markdown и диагностика PDF (`analyze_pdf_pages` через pypdf).
- **`markdown_cleanup.py`** — детерминированная OCR-чистка Markdown (`cleanup_ocr_markdown`). Покрыта тестами, но **не подключена к тракту/UI** (backend-only).
- **`ocr_auto_mode.py`** — оркестратор «OCR или нет» (`convert_pdf_with_optional_ocr`): по диагностике страниц решает, гнать ли PDF через OCR, с учётом выбранного диапазона страниц.
- **`ocr_converter.py`** — обёртка OCRmyPDF через `subprocess` (`ocrmypdf --skip-text --deskew --rotate-pages -l rus+eng`).
- **`pdf_core.py`** — провайдеро-независимое ядро PDF → Markdown. Публичные функции: `pdf_to_markdown(pdf_bytes, *, page_range, mode, provider)` и `pdf_to_markdown_with_status(...)` (последнюю использует `app.py` — UI показывает `ocr_status`). Берёт на себя работу с bytes/tempfile, не зависит от Streamlit. Определяет протокол `OcrProvider` с одной реализацией — `OcrmypdfProvider`; при `provider=None` маршрутизирует через `ocr_auto_mode.convert_pdf_with_optional_ocr` без изменений в поведении.
- **`ocr/`** — тракт MinerU (общие типы — `ocr/__init__.py`): `mineru_provider.py` (#61, облачный API v4 за протоколом `OcrProvider`, `result_from_zip` без сети), `cache.py` (#62, кэш сырого zip + `meta.json`), `postprocess.py` (#63, детерминированная чистка markdown), `validate.py` (#64, проверки + `report.json` + пометки), `diff.py` (#65, сверка прогонов `vlm`/`pipeline`), `cli.py` (#66, `python -m ocr.cli`: `run_pipeline` + `main`), `ingest.py` (#81, единый вход `.pdf/.docx/.xlsx`: выбор маршрута, один `report.json`), `board.py` (#82, `python -m ocr.board`: оффлайн-табло качества по фикстурам).

OCR-тракт (режим `auto`): `pdf_core.pdf_to_markdown_with_status` → `analyze_pdf_pages` (pypdf) → `ocr_auto_mode.convert_pdf_with_optional_ocr` → `ocr_converter` (subprocess `ocrmypdf --skip-text --deskew --rotate-pages -l rus+eng`) → `convert_with_markitdown` по OCR-слою.

Шаблон `template.docx` хранится на Google Drive (file_id `1FdPo8Ddo317ZYoPzraCTy5R4E72Ieqba`), в репозитории его нет.

## Брендбук

```
BRAND_BLUE   = #015198   заголовки H1, синий блок «Кому/От кого»
BRAND_RED    = #D04514   заголовки H2, декоративные линии
BRAND_ORANGE = #EF7F1A   цитаты, плейсхолдеры фото
TEXT_DARK    = #1A1A1A   тело документа
```

Шрифты: PT Sans (тело, 12pt), PT Sans Narrow (заголовки).

Поля страницы: left 2 cm, right 1.5 cm. Рабочая ширина — 17.5 cm.

## Соглашения

В `convert.py` каждое нетривиальное изменение помечено комментарием `# ПРАВКА #N: ...`. Новые правки нумеруются по возрастанию. Это плоская структура, не группируется в категории.

Сигнатура `convert_md_to_docx(md_text, output_filename, template_path=None, images=None)` зафиксирована и не меняется — её зовёт `app.py`.

## Деплой

Push в `main` → Streamlit Community Cloud автоматически передеплоит приложение в течение 1–2 минут. Никаких ручных действий не требуется. Если деплой упал — смотреть логи в дашборде Streamlit Cloud.

## Связанные проекты

- [`tenzosila-kp-dogovor`](https://github.com/macarovanton-cpu/tenzosila-kp-dogovor) — конструктор КП и договоров с переменными подстановками

## Структура репозитория

```
docx_converter/
├── app.py                  # Streamlit UI
├── convert.py              # ядро Markdown → DOCX
├── file_converter.py       # обратное направление + MarkItDown-слой
├── markdown_cleanup.py     # детерминированная OCR-чистка (backend-only)
├── ocr_auto_mode.py        # оркестратор «OCR или нет»
├── ocr_converter.py        # обёртка OCRmyPDF через subprocess
├── pdf_core.py             # провайдеро-независимое ядро PDF → Markdown
├── ocr/                    # тракт MinerU (OCR CLI для агента)
│   ├── __init__.py         # общие типы: Finding, SEVERITIES
│   ├── mineru_provider.py  # облачный MinerU API v4 + result_from_zip
│   ├── cache.py            # кэш сырого ответа (zip + meta.json)
│   ├── postprocess.py      # детерминированная чистка markdown после OCR
│   ├── validate.py         # проверки, report.json, пометки «!! ПРОВЕРИТЬ: … !!»
│   ├── diff.py             # сверка прогонов vlm/pipeline
│   ├── cli.py              # python -m ocr.cli
│   ├── ingest.py           # единый вход PDF/DOCX/XLSX (#81)
│   └── board.py            # python -m ocr.board — табло качества (#82)
├── requirements.txt
├── conftest.py             # пустой, нужен pytest для корневого rootdir
├── tests/
│   ├── test_convert.py         # регрессионные тесты convert.py
│   ├── test_markdown_cleanup.py
│   ├── test_ocr_auto_mode.py
│   ├── test_ocr_cli.py         # приёмочные тесты python -m ocr.cli
│   └── test_pdf_core.py
├── test_formatting.md      # md-фикстура для test_convert.py (полный прогон форматирования)
├── test_formatting_bom.md  # md-фикстура: файл с BOM в начале
├── test_image.png          # png-фикстура для инлайн-картинки в test_formatting.md
├── CLAUDE.md                # шпаргалка для Claude Code
├── AGENTS.md                # шпаргалка для Codex/других агентов
├── PROJECT_PLAN.md          # роадмап OCR-фичи
├── PROJECT_STATUS.md        # текущий статус проекта
├── PROJECT_PROGRESS.md      # архивный лог фичи MarkItDown-импорта
├── README.md                # этот файл
├── .devcontainer/
│   └── devcontainer.json    # авто-запуск Streamlit в Codespaces
└── .gitignore
```
