import streamlit as st
import io
import os
import tempfile
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
# ИМПОРТ ЯДРА — подключаем convert.py как модуль
# ─────────────────────────────────────────────────────────────────────────────
from convert import convert_md_to_docx

# ─────────────────────────────────────────────────────────────────────────────
# GOOGLE DRIVE — скачивание шаблона
# ─────────────────────────────────────────────────────────────────────────────

def download_template_from_drive(file_id: str) -> str:
    """
    Скачивает файл с Google Drive по ID в временную папку.
    Возвращает путь к скачанному файлу.
    Использует сервисный аккаунт из st.secrets.
    """
    import json
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    from google.oauth2 import service_account

    # Читаем ключи из Streamlit secrets
    # Локально: .streamlit/secrets.toml
    # На сервере: вставляешь в настройках Streamlit Cloud
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )

    service = build("drive", "v3", credentials=creds)

    # Скачиваем файл во временный файл
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    request = service.files().get_media(fileId=file_id)
    downloader = MediaIoBaseDownload(tmp, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    tmp.close()
    return tmp.name


def get_template(use_drive: bool, drive_file_id: str, local_path: str) -> str | None:
    """
    Возвращает путь к шаблону.
    Если use_drive=True — скачивает с Google Drive.
    Если False — ищет локально.
    """
    if use_drive:
        try:
            with st.spinner("Загружаю шаблон с Google Drive..."):
                path = download_template_from_drive(drive_file_id)
            return path
        except Exception as e:
            st.warning(f"⚠️ Не удалось загрузить шаблон с Drive: {e}\n"
                       f"Попробую локальный шаблон.")

    if local_path and os.path.exists(local_path):
        return local_path

    return None


# ─────────────────────────────────────────────────────────────────────────────
# КОНФИГУРАЦИЯ ТИПОВ ДОКУМЕНТОВ
# ─────────────────────────────────────────────────────────────────────────────

# Для каждого типа — название, ID шаблона на Google Drive, локальный путь
# Замени DRIVE_FILE_ID на реальные ID файлов из адресной строки Google Drive
DOC_TYPES = {
    "📄 Письмо / Сопроводительное письмо": {
        "drive_id":    "ВСТАВЬ_ID_ФАЙЛА_letter_template",   # ← заменить
        "local_path":  r"C:\Users\tonik\Desktop\docx_converter\template.docx",
        "output_name": "letter",
        "hint": "Структура: заголовок `# Название`, блок `**Кому:**`, разделы `## ...`, подпись `С уважением,`"
    },
    "📋 Пояснительная записка": {
        "drive_id":    "ВСТАВЬ_ID_ФАЙЛА_pz_template",       # ← заменить
        "local_path":  r"C:\Users\tonik\Desktop\docx_converter\template.docx",
        "output_name": "pz",
        "hint": "Структура: заголовок `# Название`, разделы `## 1. ...`, подразделы `### 1.1. ...`, callout `!! формула !!`"
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# НАСТРОЙКИ СТРАНИЦЫ
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Тензосила — Конструктор документов",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Базовые стили — убираем лишние отступы
st.markdown("""
<style>
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    .stTextArea textarea { font-family: monospace; font-size: 13px; }
    h1 { color: #015198; }
    .status-box {
        padding: 12px 16px;
        border-radius: 6px;
        margin: 8px 0;
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# ШАПКА
# ─────────────────────────────────────────────────────────────────────────────

col_logo, col_title = st.columns([1, 5])
with col_logo:
    st.markdown("## ⚖️")
with col_title:
    st.markdown("## Конструктор документов")
    st.caption("ООО «ТПК «Тензосила» · Фирменное оформление по брендбуку")

st.divider()

# ─────────────────────────────────────────────────────────────────────────────
# ОСНОВНОЙ ИНТЕРФЕЙС — две колонки
# ─────────────────────────────────────────────────────────────────────────────

col_left, col_right = st.columns([3, 2], gap="large")

with col_left:

    # ── Выбор типа документа ─────────────────────────────────────────────────
    st.markdown("#### 1. Тип документа")
    doc_type = st.selectbox(
        label="Выбери тип",
        options=list(DOC_TYPES.keys()),
        label_visibility="collapsed"
    )
    config = DOC_TYPES[doc_type]

    # Подсказка по структуре
    with st.expander("📌 Как оформить текст для этого типа", expanded=False):
        st.code(config["hint"], language=None)

    st.markdown("#### 2. Текст документа")

    # ── Вкладки: вставить текст или загрузить файл ───────────────────────────
    tab_paste, tab_upload = st.tabs(["✏️ Вставить текст", "📂 Загрузить .md файл"])

    md_text = ""

    with tab_paste:
        md_text_input = st.text_area(
            label="Вставь текст в формате Markdown",
            height=400,
            placeholder="# Заголовок документа\n\n**Кому:** Название компании\n\nТекст...",
            label_visibility="collapsed"
        )
        if md_text_input:
            md_text = md_text_input

    with tab_upload:
        uploaded_file = st.file_uploader(
            label="Выбери .md файл",
            type=["md", "txt"],
            label_visibility="collapsed"
        )
        if uploaded_file is not None:
            md_text = uploaded_file.read().decode("utf-8")
            st.success(f"Файл загружен: **{uploaded_file.name}** · {len(md_text)} символов")

            # Показываем превью
            with st.expander("👁 Превью текста", expanded=False):
                st.text(md_text[:1500] + ("..." if len(md_text) > 1500 else ""))


with col_right:

    st.markdown("#### 3. Сформировать документ")

    # ── Имя выходного файла ───────────────────────────────────────────────────
    default_name = f"{config['output_name']}_{datetime.now().strftime('%d%m%Y')}"
    output_name = st.text_input(
        "Имя файла (без расширения)",
        value=default_name
    )

    st.markdown("---")

    # ── Кнопка запуска ────────────────────────────────────────────────────────
    generate_btn = st.button(
        "⚙️ Сформировать документ",
        type="primary",
        use_container_width=True,
        disabled=(not md_text.strip())
    )

    if not md_text.strip():
        st.caption("Вставь текст или загрузи файл чтобы активировать кнопку")

    # ── Генерация ─────────────────────────────────────────────────────────────
    if generate_btn and md_text.strip():

        with st.spinner("Формирую документ..."):

            # Определяем использовать ли Drive
            # Если в secrets есть ключ gcp_service_account — используем Drive
            use_drive = "gcp_service_account" in st.secrets

            # Получаем шаблон
            template_path = get_template(
                use_drive=use_drive,
                drive_file_id=config["drive_id"],
                local_path=config["local_path"]
            )

            if template_path is None:
                st.error("❌ Шаблон не найден. Проверь настройки Google Drive или локальный путь в app.py")
                st.stop()

            # Генерируем документ во временный файл
            try:
                tmp_out = tempfile.NamedTemporaryFile(
                    delete=False, suffix=".docx"
                )
                tmp_out.close()

                convert_md_to_docx(
                    md_text=md_text,
                    output_filename=tmp_out.name,
                    template_path=template_path
                )

                # Читаем результат в байты для скачивания
                with open(tmp_out.name, "rb") as f:
                    docx_bytes = f.read()

                # Убираем временные файлы
                os.unlink(tmp_out.name)
                if use_drive and template_path != config["local_path"]:
                    try:
                        os.unlink(template_path)
                    except Exception:
                        pass

                st.success("✅ Документ готов!")

                # ── Кнопка скачивания ─────────────────────────────────────
                st.download_button(
                    label="⬇️ Скачать .docx",
                    data=docx_bytes,
                    file_name=f"{output_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

                # Статистика
                size_kb = len(docx_bytes) / 1024
                st.caption(f"Размер файла: {size_kb:.1f} КБ")

            except Exception as e:
                st.error(f"❌ Ошибка при генерации документа:\n\n```\n{e}\n```")
                st.info("Проверь корректность Markdown-разметки в тексте")

    # ── Справка по Markdown ───────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("#### 📖 Шпаргалка по разметке")
    st.markdown("""
| Что | Как писать |
|-----|------------|
| Заголовок документа | `# Текст` |
| Раздел | `## Текст` |
| Подраздел | `### Текст` |
| **Жирный** | `**текст**` |
| *Курсив* | `*текст*` |
| Таблица | `\\| А \\| Б \\|` |
| Врезка-акцент | `!! текст !!` |
| Блок «Кому» | `**Кому:** ...` |
| Подпись | `С уважением,` |
| Разделитель | `---` |
""")


# ─────────────────────────────────────────────────────────────────────────────
# ПОДВАЛ
# ─────────────────────────────────────────────────────────────────────────────

st.divider()
st.caption("ООО «ТПК «Тензосила» · Внутренний инструмент · Документы формируются по фирменному брендбуку")
