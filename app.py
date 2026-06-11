import hashlib
import io
import os
import re
import tempfile
import zipfile
from datetime import datetime

import streamlit as st

from convert import convert_md_to_docx
from file_converter import (
    analyze_pdf_pages,
    convert_file_to_md,
    convert_with_markitdown,
    parse_page_range,
)


def download_template_from_drive(file_id: str) -> str:
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    from google.oauth2 import service_account
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = service_account.Credentials.from_service_account_info(
        creds_dict, scopes=["https://www.googleapis.com/auth/drive.readonly"])
    service = build("drive", "v3", credentials=creds)
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    request = service.files().get_media(fileId=file_id)
    downloader = MediaIoBaseDownload(tmp, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    tmp.close()
    return tmp.name


def get_template(use_drive, drive_file_id, local_path):
    if use_drive:
        try:
            with st.spinner("Загружаю шаблон с Google Drive..."):
                path = download_template_from_drive(drive_file_id)
            return path
        except Exception as e:
            st.warning(f"⚠️ Не удалось загрузить шаблон с Drive: {e}")
    if local_path and os.path.exists(local_path):
        return local_path
    return None


def _file_ext(filename: str) -> str:
    return filename.lower().rsplit('.', 1)[-1] if '.' in filename else ''


def _normalize_page_range(range_text: str | None) -> str | None:
    if not range_text:
        return None
    value = range_text.strip()
    if not value or value.lower() in ("all", "все"):
        return None
    return value


def _safe_md_filename(filename: str) -> str:
    stem = filename.rsplit('.', 1)[0]
    stem = re.sub(r'[^\w\-а-яА-ЯёЁ]+', '_', stem, flags=re.UNICODE).strip('_')
    return f"{stem or 'converted'}.md"


def _unique_md_filename(filename: str, used_names: set[str]) -> str:
    safe_name = _safe_md_filename(filename)
    stem = safe_name[:-3]
    candidate = safe_name
    counter = 2
    while candidate in used_names:
        candidate = f"{stem}_{counter}.md"
        counter += 1
    used_names.add(candidate)
    return candidate


def _build_markdown_zip(results: list[dict]) -> tuple[bytes, int, int]:
    buffer = io.BytesIO()
    used_names = set()
    included = 0
    skipped = 0

    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for result in results:
            if result.get("error") or not result.get("markdown"):
                skipped += 1
                continue
            archive_name = _unique_md_filename(result["filename"], used_names)
            zf.writestr(archive_name, result["markdown"])
            included += 1

    return buffer.getvalue(), included, skipped


def _build_combined_markdown(results: list[dict]) -> tuple[str, int, list[str]]:
    chunks = []
    skipped_files = []

    for result in results:
        if result.get("error") or not result.get("markdown"):
            skipped_files.append(result["filename"])
            continue

        file_number = len(chunks) + 1
        filename = result["filename"]
        page_range = result.get("page_range") or "all"
        file_type = result.get("file_type") or _file_ext(filename).upper()
        markdown = result["markdown"].strip()
        chunks.append(
            f"# Файл {file_number}: {filename}\n\n"
            f"Источник: {filename}  \n"
            f"Диапазон страниц: {page_range}  \n"
            f"Тип файла: {file_type}  \n\n"
            "---\n\n"
            f"{markdown}\n\n"
            "---"
        )

    combined = "\n\n".join(chunks)
    if combined:
        combined += "\n"
    return combined, len(chunks), skipped_files


def _save_uploaded_to_temp(uploaded_file, ext: str) -> str:
    suffix = f".{ext}" if ext else ""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    try:
        tmp.write(uploaded_file.getvalue())
        return tmp.name
    finally:
        tmp.close()


@st.cache_data(show_spinner=False)
def _analyze_pdf_pages_cached(file_bytes: bytes, file_hash: str) -> list[dict]:
    _ = file_hash
    tmp_path = None
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    try:
        tmp_path = tmp.name
        try:
            tmp.write(file_bytes)
        finally:
            tmp.close()
        return analyze_pdf_pages(tmp_path)
    finally:
        if tmp_path:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass


def _display_pdf_diagnostics(uploaded_file, ext: str):
    if ext != "pdf":
        return

    try:
        file_bytes = uploaded_file.getvalue()
        file_hash = hashlib.sha256(file_bytes).hexdigest()
        pages = _analyze_pdf_pages_cached(file_bytes, file_hash)
    except Exception as e:
        st.warning(f"Не удалось прочитать диагностику PDF: {e}")
        return

    image_only = [
        page["page_number"] for page in pages
        if not page.get("has_text_layer")
    ]
    st.caption(f"PDF: {len(pages)} стр.")
    if image_only:
        st.warning(
            "Страницы без текстового слоя: "
            f"{', '.join(map(str, image_only))}. "
            "Для image-only страниц потребуется OCR."
        )


def _display_ocr_candidate_status(uploaded_file, ext: str, ocr_mode: str):
    if ext != "pdf" or ocr_mode != "auto":
        return

    try:
        file_bytes = uploaded_file.getvalue()
        file_hash = hashlib.sha256(file_bytes).hexdigest()
        pages = _analyze_pdf_pages_cached(file_bytes, file_hash)
    except Exception as e:
        st.warning(f"OCR auto: не удалось прочитать диагностику PDF: {e}")
        return

    image_only = [
        page["page_number"] for page in pages
        if not page.get("has_text_layer")
    ]
    if image_only:
        st.info(
            "OCR auto: кандидат на OCR "
            f"(страницы без текстового слоя: {', '.join(map(str, image_only))}). "
            "OCR пока не запускается."
        )
    else:
        st.success("OCR auto: текстовый слой найден, OCR не нужен.")


def _convert_uploaded_file(uploaded_file, page_range: str | None) -> dict:
    ext = _file_ext(uploaded_file.name)
    display_range = page_range or "all"
    tmp_path = _save_uploaded_to_temp(uploaded_file, ext)
    try:
        markdown = convert_with_markitdown(tmp_path, page_range=page_range)
        return {
            "filename": uploaded_file.name,
            "download_name": _safe_md_filename(uploaded_file.name),
            "file_type": ext.upper() or "UNKNOWN",
            "page_range": display_range,
            "markdown": markdown,
            "error": None,
        }
    except Exception as e:
        return {
            "filename": uploaded_file.name,
            "download_name": _safe_md_filename(uploaded_file.name),
            "file_type": ext.upper() or "UNKNOWN",
            "page_range": display_range,
            "markdown": "",
            "error": str(e),
        }
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


DOC_TYPES = {
    "📄 Письмо / Сопроводительное письмо": {
        "drive_id":    "1FdPo8Ddo317ZYoPzraCTy5R4E72Ieqba",
        "local_path":  r"C:\Users\tonik\Desktop\docx_converter\template.docx",
        "output_name": "letter",
        "hint": "Структура: заголовок `# Название`, блок `**Кому:**`, разделы `## ...`, подпись `С уважением,`"
    },
    "📋 Пояснительная записка": {
        "drive_id":    "1FdPo8Ddo317ZYoPzraCTy5R4E72Ieqba",
        "local_path":  r"C:\Users\tonik\Desktop\docx_converter\template.docx",
        "output_name": "pz",
        "hint": "Структура: заголовок `# Название`, разделы `## 1. ...`, подразделы `### 1.1. ...`, callout `!! формула !!`"
    },
}


def render_md_to_docx_mode():
    col_left, col_right = st.columns([3, 2], gap="large")

    with col_left:
        st.markdown("#### 1. Тип документа")
        doc_type = st.selectbox("Тип", options=list(DOC_TYPES.keys()),
                                label_visibility="collapsed")
        config = DOC_TYPES[doc_type]

        with st.expander("📌 Как оформить текст для этого типа", expanded=False):
            st.code(config["hint"], language=None)

        st.markdown("#### 2. Текст документа")

        tab_paste, tab_md, tab_file = st.tabs([
            "✏️ Вставить текст",
            "📂 Загрузить .md",
            "📄 Загрузить файл (docx / pdf / txt)"
        ])

        md_text = ""
        source_images = []

        with tab_paste:
            md_input = st.text_area(
                "Markdown текст", height=400,
                placeholder="# Заголовок\n\nТекст...",
                label_visibility="collapsed")
            if md_input:
                md_text = md_input

        with tab_md:
            upl_md = st.file_uploader("MD файл", type=["md", "txt"],
                                      label_visibility="collapsed", key="upl_md")
            if upl_md:
                md_text = upl_md.read().decode("utf-8")
                st.success(
                    f"Загружен: **{upl_md.name}** · {len(md_text)} символов")
                with st.expander("👁 Превью", expanded=False):
                    st.text(md_text[:1500] + (
                        "..." if len(md_text) > 1500 else ""))

        with tab_file:
            st.caption(
                "Загрузи DOCX, PDF или TXT — скрипт определит заголовки "
                "по размеру и жирности шрифта и переоформит в фирменный стиль."
            )
            upl_file = st.file_uploader(
                "Файл", type=["docx", "pdf", "txt"],
                label_visibility="collapsed", key="upl_file")

            if upl_file:
                file_bytes = upl_file.read()
                with st.spinner(f"Извлекаю текст из {upl_file.name}..."):
                    try:
                        md_text, source_images = convert_file_to_md(
                            file_bytes, upl_file.name)

                        img_info = (f" · {len(source_images)} изображений"
                                    if source_images else "")
                        st.success(
                            f"✅ Обработан: **{upl_file.name}** · "
                            f"{len(md_text)} символов{img_info}")

                        with st.expander(
                                "👁 Распознанная структура", expanded=False):
                            headers = [l for l in md_text.split('\n')
                                       if l.startswith('#')]
                            st.text('\n'.join(headers[:30]) if headers
                                    else md_text[:800])

                        edited = st.text_area(
                            "✏️ Отредактируй при необходимости:",
                            value=md_text, height=300, key="edit_md")
                        if edited:
                            md_text = edited

                    except Exception as e:
                        st.error(f"❌ Ошибка: {e}")
                        md_text = ""

    with col_right:
        st.markdown("#### 3. Сформировать документ")

        default_name = (
            f"{config['output_name']}_{datetime.now().strftime('%d%m%Y')}")
        output_name = st.text_input("Имя файла (без расширения)",
                                    value=default_name)
        st.markdown("---")

        btn = st.button("⚙️ Сформировать документ", type="primary",
                        use_container_width=True,
                        disabled=(not md_text.strip()))

        if not md_text.strip():
            st.caption("Вставь текст или загрузи файл")

        if btn and md_text.strip():
            with st.spinner("Формирую документ..."):
                use_drive = "gcp_service_account" in st.secrets
                template_path = get_template(
                    use_drive=use_drive,
                    drive_file_id=config["drive_id"],
                    local_path=config["local_path"])

                if template_path is None:
                    st.error("❌ Шаблон не найден.")
                    st.stop()

                try:
                    tmp_out = tempfile.NamedTemporaryFile(
                        delete=False, suffix=".docx")
                    tmp_out_path = tmp_out.name
                    try:
                        tmp_out.close()

                        convert_md_to_docx(md_text=md_text,
                                           output_filename=tmp_out_path,
                                           template_path=template_path,
                                           images=source_images)

                        with open(tmp_out_path, 'rb') as f:
                            docx_bytes = f.read()
                    finally:
                        try:
                            os.unlink(tmp_out_path)
                        except OSError:
                            pass

                    if use_drive and template_path != config["local_path"]:
                        try:
                            os.unlink(template_path)
                        except Exception:
                            pass

                    st.success("✅ Документ готов!")
                    st.download_button(
                        "⬇️ Скачать .docx", data=docx_bytes,
                        file_name=f"{output_name}.docx",
                        mime="application/vnd.openxmlformats-officedocument"
                             ".wordprocessingml.document",
                        use_container_width=True)
                    st.caption(f"Размер: {len(docx_bytes)/1024:.1f} КБ")

                except Exception as e:
                    st.error(f"❌ Ошибка:\n\n```\n{e}\n```")

        st.markdown("---")
        st.markdown("#### 📖 Шпаргалка")
        st.markdown("""
| Что | Разметка |
|-----|----------|
| Заголовок | `# Текст` |
| Раздел | `## Текст` |
| Подраздел | `### Текст` |
| **Жирный** | `**текст**` |
| Таблица | `\\| А \\| Б \\|` |
| Врезка | `!! текст !!` |
| Блок «Кому» | `**Кому:** ...` |
| Подпись | `С уважением,` |
""")


def render_files_to_markdown_mode():
    st.session_state.setdefault("files_to_md_results", [])

    st.markdown("#### Файлы -> Markdown")
    uploaded_files = st.file_uploader(
        "Загрузите PDF, DOCX, XLSX или PPTX",
        type=["pdf", "docx", "xlsx", "pptx"],
        accept_multiple_files=True,
        key="files_to_md_uploader",
    )

    if not uploaded_files:
        st.caption("Загрузите один или несколько файлов для конвертации.")
        return

    ocr_mode = st.radio(
        "OCR mode",
        options=["off", "auto"],
        index=0,
        horizontal=True,
        key="files_to_md_ocr_mode",
        help=(
            "off: текущая конвертация через MarkItDown без OCR. "
            "auto: сейчас только показывает PDF-кандидаты на OCR."
        ),
    )
    if ocr_mode == "off":
        st.caption("OCR выключен: используется текущий MarkItDown flow.")
    else:
        st.caption("OCR auto: на этом шаге OCR не запускается, показывается только диагностика.")

    range_keys = {
        idx: f"page_range_{idx}_{_safe_md_filename(uploaded_file.name)}"
        for idx, uploaded_file in enumerate(uploaded_files)
    }

    st.markdown("#### Общий диапазон для PDF")
    common_col, apply_col = st.columns([3, 1], vertical_alignment="bottom")
    with common_col:
        common_pdf_range = st.text_input(
            "Общий диапазон страниц для PDF",
            key="common_pdf_page_range",
            placeholder="Например: 1-3, 7, 10-12",
        )
    with apply_col:
        apply_common_range = st.button(
            "Применить диапазон ко всем PDF",
            use_container_width=True,
        )

    if apply_common_range:
        normalized_common_range = _normalize_page_range(common_pdf_range)
        pdf_count = 0
        non_pdf_count = 0
        if not normalized_common_range:
            st.warning("Введите диапазон страниц для PDF.")
        else:
            try:
                parse_page_range(normalized_common_range)
            except ValueError as e:
                normalized_common_range = None
                st.error(f"Некорректный диапазон страниц: {e}")
        if normalized_common_range:
            for idx, uploaded_file in enumerate(uploaded_files):
                if _file_ext(uploaded_file.name) == "pdf":
                    st.session_state[range_keys[idx]] = normalized_common_range
                    pdf_count += 1
                else:
                    non_pdf_count += 1
            if pdf_count:
                st.success(
                    f"Диапазон {normalized_common_range} применён к PDF: "
                    f"{pdf_count}."
                )
            else:
                st.warning("Среди загруженных файлов нет PDF.")
            if non_pdf_count:
                st.info(
                    "Для DOCX/XLSX/PPTX диапазоны страниц пока не "
                    "поддержаны, эти файлы будут конвертированы целиком."
                )

    st.markdown("#### Настройки файлов")
    for idx, uploaded_file in enumerate(uploaded_files):
        ext = _file_ext(uploaded_file.name)
        with st.container(border=True):
            meta_col, range_col = st.columns([2, 1], vertical_alignment="top")
            with meta_col:
                st.markdown(f"**{uploaded_file.name}**")
                st.caption(f"Тип: .{ext or 'unknown'}")
                _display_pdf_diagnostics(uploaded_file, ext)
                _display_ocr_candidate_status(uploaded_file, ext, ocr_mode)
            with range_col:
                key = range_keys[idx]
                if key not in st.session_state:
                    st.session_state[key] = "all"
                page_range = st.text_input(
                    "Диапазон страниц",
                    key=key,
                    help="Для PDF: 1-3, 7, 10-12. Для остальных форматов используйте all.",
                )
                normalized = _normalize_page_range(page_range)
                if ext != "pdf" and normalized:
                    st.warning(
                        "Page range пока поддержан только для PDF. "
                        "Для DOCX/XLSX/PPTX конвертируется весь файл."
                    )
                elif ext == "pdf" and normalized:
                    try:
                        parse_page_range(normalized)
                    except ValueError as e:
                        st.error(str(e))

    if st.button("Конвертировать в Markdown", type="primary"):
        invalid_ranges = []
        for idx, uploaded_file in enumerate(uploaded_files):
            if _file_ext(uploaded_file.name) != "pdf":
                continue
            page_range = _normalize_page_range(st.session_state.get(range_keys[idx]))
            if not page_range:
                continue
            try:
                parse_page_range(page_range)
            except ValueError as e:
                invalid_ranges.append(f"{uploaded_file.name}: {e}")

        if invalid_ranges:
            st.error(
                "Конвертация не запущена — исправьте диапазоны страниц:\n\n"
                + "\n\n".join(invalid_ranges)
            )
        else:
            results = []
            progress = st.progress(0)
            for idx, uploaded_file in enumerate(uploaded_files):
                key = range_keys[idx]
                page_range = _normalize_page_range(st.session_state.get(key))
                with st.spinner(f"Конвертирую {uploaded_file.name}..."):
                    result = _convert_uploaded_file(uploaded_file, page_range)
                results.append(result)
                progress.progress((idx + 1) / len(uploaded_files))
            progress.empty()
            st.session_state.files_to_md_results = results

    results = st.session_state.get("files_to_md_results", [])
    if not results:
        return

    st.markdown("#### Результаты")
    zip_bytes, included_count, skipped_count = _build_markdown_zip(results)
    if skipped_count:
        st.warning(
            f"В ZIP не включено файлов с ошибками: {skipped_count}."
        )
    if included_count:
        st.download_button(
            "Скачать все .md в ZIP",
            data=zip_bytes,
            file_name="markdown_results.zip",
            mime="application/zip",
            key="download_all_md_zip",
            use_container_width=True,
        )

    combined_md, combined_count, skipped_files = _build_combined_markdown(results)
    if skipped_files:
        st.warning(
            "В объединенный Markdown не включены файлы с ошибками: "
            f"{', '.join(skipped_files)}."
        )
    if combined_count:
        st.download_button(
            "Скачать объединенный Markdown",
            data=combined_md.encode("utf-8"),
            file_name="combined.md",
            mime="text/markdown",
            key="download_combined_md",
            use_container_width=True,
        )

    for idx, result in enumerate(results):
        with st.container(border=True):
            st.markdown(f"**{result['filename']}**")
            if result["error"]:
                st.error(result["error"])
                continue

            markdown = result["markdown"]
            st.text_area(
                "Markdown preview",
                value=markdown[:5000],
                height=260,
                key=f"md_preview_{idx}_{result['download_name']}",
            )
            if len(markdown) > 5000:
                st.caption(
                    f"Показаны первые 5000 символов из {len(markdown)}.")
            st.download_button(
                "Скачать .md",
                data=markdown.encode("utf-8"),
                file_name=result["download_name"],
                mime="text/markdown",
                key=f"download_md_{idx}_{result['download_name']}",
                use_container_width=True,
            )


st.set_page_config(
    page_title="Тензосила — Конструктор документов",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    .stTextArea textarea { font-family: monospace; font-size: 13px; }
</style>
""", unsafe_allow_html=True)

col_logo, col_title = st.columns([1, 5])
with col_logo:
    st.markdown("## ⚖️")
with col_title:
    st.markdown("## Конструктор документов")
    st.caption("ООО «ТПК «Тензосила» · Фирменное оформление по брендбуку")

st.divider()

mode = st.segmented_control(
    "Режим",
    ["Markdown -> DOCX", "Файлы -> Markdown"],
    default="Markdown -> DOCX",
)

st.divider()

if mode == "Файлы -> Markdown":
    render_files_to_markdown_mode()
else:
    render_md_to_docx_mode()

st.divider()
st.caption("ООО «ТПК «Тензосила» · Внутренний инструмент")
