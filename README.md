# docx_converter

Веб-приложение на Streamlit, конвертирующее Markdown в DOCX-документы в фирменном стиле ООО «ТПК «Тензосила». Используется для коммерческих предложений, пояснительных записок и сопроводительных писем.

🌐 **Прод:** https://docxconverter-5n8ntdurz5scqneuvithfb.streamlit.app/

## Возможности

- Три способа ввода: вставить Markdown текст, загрузить `.md` файл, загрузить готовый DOCX/PDF/TXT с авто-распознаванием структуры
- Два типа документа на выходе: письмо/коммерческое предложение, пояснительная записка
- Шаблон с фирменным хедером (логотип, реквизиты компании, ОГРН, ИНН) подгружается с Google Drive
- Кастомный Markdown-синтаксис: callout-врезки `!! ... !!`, блок реквизитов `**Кому:**`, цитаты, плейсхолдеры фото
- Автоматические ✓/✗ в сравнительных таблицах для ячеек «Да»/«Нет»/«Отсутствует»

## Запуск локально

```bash
pip install -r requirements.txt
streamlit run app.py
```

Или через Codespaces — `.devcontainer/devcontainer.json` запускает приложение автоматически после attach.

Для работы с шаблоном с Google Drive нужен `.streamlit/secrets.toml` с секцией `[gcp_service_account]`. Без него приложение использует локальный шаблон (путь задаётся в `DOC_TYPES` в `app.py`).

## Архитектура

Три Python-модуля:

- **`app.py`** — Streamlit UI. Загружает `.docx` шаблон с Google Drive через service account, вызывает `convert_md_to_docx`, отдаёт результат на скачивание.
- **`convert.py`** — ядро Markdown → DOCX (~705 строк). Единственная публичная функция: `convert_md_to_docx(md_text, output_filename, template_path=None)`.
- **`file_converter.py`** — обратная сторона: DOCX/PDF/TXT → Markdown. Используется для предзаполнения редактора при загрузке готового документа.

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

Сигнатура `convert_md_to_docx(md_text, output_filename, template_path=None)` зафиксирована и не меняется — её зовёт `app.py`.

## Деплой

Push в `main` → Streamlit Community Cloud автоматически передеплоит приложение в течение 1–2 минут. Никаких ручных действий не требуется. Если деплой упал — смотреть логи в дашборде Streamlit Cloud.

## Связанные проекты

- [`tenzosila-kp-dogovor`](https://github.com/macarovanton-cpu/tenzosila-kp-dogovor) — конструктор КП и договоров с переменными подстановками

## Структура репозитория

```
docx_converter/
├── app.py                 # Streamlit UI
├── convert.py             # ядро Markdown → DOCX
├── file_converter.py      # обратное направление
├── requirements.txt
├── CLAUDE.md              # шпаргалка для Claude Code
├── README.md              # этот файл
├── .devcontainer/
│   └── devcontainer.json  # авто-запуск Streamlit в Codespaces
└── .gitignore
```
