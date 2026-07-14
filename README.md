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
- Кастомный Markdown-синтаксис: callout-врезки `!! ... !!`, блок реквизитов `**Кому:**`, цитаты, плейсхолдеры фото
- Автоматические ✓/✗ в сравнительных таблицах для ячеек «Да»/«Нет»/«Отсутствует»

### Файлы -> Markdown

- Пакетная загрузка нескольких файлов
- Поддерживаемые форматы: `.pdf`, `.docx`, `.xlsx`, `.pptx`
- Конвертация в Markdown через Microsoft MarkItDown
- Для PDF можно указать диапазон страниц
- Для PDF показывается диагностика: количество страниц, наличие текстового слоя, предупреждение по image-only страницам
- Режим OCR `off` / `auto`: в `auto` PDF без текстового слоя (с учётом выбранного диапазона страниц) прогоняется через OCR, после чего конвертируется в Markdown
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
- OCR реализован и подключён к UI режима `Файлы -> Markdown` (режим `auto`). Для PDF без текстового слоя (с учётом выбранного диапазона страниц) выполняется `ocrmypdf --skip-text --deskew --rotate-pages -l rus+eng`, после чего MarkItDown извлекает текст из OCR-слоя. Важное прод-ограничение см. в разделе «Известное ограничение прода» ниже.
- Диапазоны страниц сейчас поддержаны только для PDF.
- Для DOCX/XLSX/PPTX выполняется конвертация всего файла. Если для этих форматов указан page range, приложение покажет ограничение.
- ZIP-архив в режиме `Файлы -> Markdown` содержит только успешно сконвертированные `.md`; результаты с ошибками не включаются.

## Известное ограничение прода

OCR-режим `auto` реализован и подключён к UI, но `ocrmypdf` **не** добавлен в `requirements.txt`, и файла `packages.txt` нет. На Streamlit Community Cloud из-за этого недоступны системные Tesseract/Ghostscript, поэтому OCR-режим `auto` на проде сейчас падает. Это открытый вопрос, который решается отдельной задачей о пакетировании. В рамках текущей документационной правки в `requirements.txt` / `packages.txt` ничего не добавляем.

## Архитектура

Шесть Python-модулей:

- **`app.py`** — Streamlit UI. Содержит режимы `Markdown -> DOCX` и `Файлы -> Markdown`. В первом режиме загружает `.docx` шаблон с Google Drive через service account, вызывает `convert_md_to_docx`, отдаёт результат на скачивание. Во втором режиме принимает несколько файлов, вызывает MarkItDown-слой и отдаёт `.md`/ZIP на скачивание.
- **`convert.py`** — ядро Markdown → DOCX (~1150 строк). Единственная публичная функция: `convert_md_to_docx(md_text, output_filename, template_path=None, images=None)`.
- **`file_converter.py`** — обратная сторона: DOCX/PDF/TXT → Markdown для предзаполнения редактора, а также отдельный MarkItDown-слой (`convert_with_markitdown`) для PDF/DOCX/XLSX/PPTX → Markdown и диагностика PDF (`analyze_pdf_pages` через pypdf).
- **`markdown_cleanup.py`** — детерминированная OCR-чистка Markdown (`cleanup_ocr_markdown`). Покрыта тестами, но **не подключена к тракту/UI** (backend-only).
- **`ocr_auto_mode.py`** — оркестратор «OCR или нет» (`convert_pdf_with_optional_ocr`): по диагностике страниц решает, гнать ли PDF через OCR, с учётом выбранного диапазона страниц.
- **`ocr_converter.py`** — обёртка OCRmyPDF через `subprocess` (`ocrmypdf --skip-text --deskew --rotate-pages -l rus+eng`).

OCR-тракт (режим `auto`): `analyze_pdf_pages` (pypdf) → `ocr_auto_mode.convert_pdf_with_optional_ocr` → `ocr_converter` (subprocess `ocrmypdf --skip-text --deskew --rotate-pages -l rus+eng`) → `convert_with_markitdown` по OCR-слою.

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
├── app.py                 # Streamlit UI
├── convert.py             # ядро Markdown → DOCX
├── file_converter.py      # обратное направление + MarkItDown-слой
├── markdown_cleanup.py    # детерминированная OCR-чистка (backend-only)
├── ocr_auto_mode.py       # оркестратор «OCR или нет»
├── ocr_converter.py       # обёртка OCRmyPDF через subprocess
├── requirements.txt
├── CLAUDE.md              # шпаргалка для Claude Code
├── README.md              # этот файл
├── .devcontainer/
│   └── devcontainer.json  # авто-запуск Streamlit в Codespaces
└── .gitignore
```
