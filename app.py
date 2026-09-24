import hashlib
import io
import json
import os
import re
import tempfile
import time
import zipfile
from datetime import datetime
from pathlib import Path

import streamlit as st

from convert import convert_md_to_docx
from file_converter import (
    analyze_pdf_pages,
    convert_file_to_md,
    convert_with_markitdown,
    parse_page_range,
)
from ocr.cache import LocalCache
from ocr.ingest import detect_route, ingest  # ПРАВКА #85: UI зовёт единый вход, не run_pipeline
from ocr.mineru_provider import MineruAuthError, MineruProvider
from ocr.validate import ANNOTATION_PREFIX
from ocr.claude_code_verifier import ClaudeCodeMissingError, find_claude   # ПРАВКА #92
from ocr.vision import VISION_MODEL                                        # ПРАВКА #92
from ocr_auto_mode import pdf_pages_without_text_layer
from ocr_converter import check_ocr_dependencies
from pdf_core import pdf_to_markdown_with_status


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


def _drive_secrets_available() -> bool:
    """Нет secrets.toml — работаем без Google Drive, а не падаем.

    StreamlitSecretNotFoundError — подкласс FileNotFoundError.
    """
    try:
        return "gcp_service_account" in st.secrets
    except FileNotFoundError:
        return False


def _file_ext(filename: str) -> str:
    return filename.lower().rsplit('.', 1)[-1] if '.' in filename else ''


def _normalize_page_range(range_text: str | None) -> str | None:
    if not range_text:
        return None
    value = range_text.strip()
    if not value or value.lower() in ("all", "все"):
        return None
    return value


def _decode_md_upload(data: bytes) -> str:
    """utf-8-sig срезает BOM: без этого первый '# Title' не распознаётся
    как H1, cp1251 — откат для файлов из Windows-редакторов."""
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return data.decode("cp1251", errors="replace")


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

    image_only = pdf_pages_without_text_layer(pages)
    st.caption(f"PDF: {len(pages)} стр.")
    if image_only:
        st.warning(
            "Страницы без текстового слоя: "
            f"{', '.join(map(str, image_only))}. "
            "Для image-only страниц потребуется OCR."
        )


def _display_ocr_candidate_status(uploaded_file, ext: str, ocr_mode: str,
                                  page_range: str | None):
    if ext != "pdf" or ocr_mode != "auto":
        return

    try:
        file_bytes = uploaded_file.getvalue()
        file_hash = hashlib.sha256(file_bytes).hexdigest()
        pages = _analyze_pdf_pages_cached(file_bytes, file_hash)
    except Exception as e:
        st.warning(f"OCR auto: не удалось прочитать диагностику PDF: {e}")
        return

    try:
        image_only = pdf_pages_without_text_layer(pages, page_range)
    except ValueError as e:
        st.warning(f"OCR auto: исправьте диапазон страниц: {e}")
        return

    if image_only:
        st.info(
            "OCR auto: кандидат на OCR "
            f"(страницы без текстового слоя: {', '.join(map(str, image_only))}). "
            "OCR будет применён при конвертации."
        )
    else:
        st.success("OCR auto: текстовый слой найден, OCR не нужен.")


# ПРАВКА #73: MinerU (этап 8 OCR-тракта) в режиме «Файлы -> Markdown».
# Кэш рядом с app.py, а не от cwd; на Streamlit Cloud он эфемерный — это осознанно.
_OCR_CACHE_ROOT = Path(__file__).resolve().parent / ".cache" / "ocr"

_OCR_MODE_LABELS = {
    "off": "Без OCR",
    "auto": "OCRmyPDF (локально)",
    "mineru": "MinerU (облако)",
}

# ПРАВКА #85: в режиме mineru PDF/DOCX/XLSX идут через ocr.ingest.
_INGEST_EXTS = ("pdf", "docx", "xlsx")      # ПРАВКА #85: что в режиме mineru идёт через ocr.ingest
_MINERU_ROUTES = ("scan", "text_tables")    # маршруты, на которых есть облако и второй прогон

_ROUTE_LABELS = {
    "scan": "скан → MinerU (облако)",
    "text_tables": "текстовый PDF с таблицами → MinerU (облако)",
    "text": "текстовый PDF без таблиц → MarkItDown (без облака)",
    "office": "DOCX/XLSX → MarkItDown (без облака)",
}

_MINERU_KEY_HINT = (
    "Проверьте MINERU_API_KEY: .streamlit/secrets.toml локально, Secrets на "
    "Streamlit Cloud или переменная окружения."
)


def _ocr_mode_options(ocrmypdf_ok: bool) -> list[str]:
    """OCRmyPDF предлагается, только если его бинарники есть (на Streamlit Cloud их нет)."""
    return ["off"] + (["auto"] if ocrmypdf_ok else []) + ["mineru"]


@st.cache_resource(show_spinner=False)
def _ocrmypdf_available() -> bool:
    # три subprocess --version: один раз на процесс, а не на каждый rerun
    return all(check["ok"] for check in check_ocr_dependencies().values())


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


def _mineru_api_key() -> str | None:
    """st.secrets, затем окружение. Нет secrets.toml — FileNotFoundError, нет ключа — KeyError."""
    try:
        return st.secrets["MINERU_API_KEY"]
    except (FileNotFoundError, KeyError):
        return os.environ.get("MINERU_API_KEY")


def _pdf_page_subset(pdf_bytes: bytes, page_range: str | None) -> bytes:
    """run_pipeline диапазона не принимает: выбранные страницы вырезаются до отправки."""
    page_indexes = parse_page_range(page_range)
    if page_indexes is None:
        return pdf_bytes

    from pypdf import PdfReader, PdfWriter

    reader = PdfReader(io.BytesIO(pdf_bytes))
    total_pages = len(reader.pages)
    writer = PdfWriter()
    for page_index in page_indexes:
        if page_index >= total_pages:
            raise ValueError(
                f"Страница {page_index + 1} вне диапазона PDF: "
                f"в файле всего {total_pages} стр."
            )
        writer.add_page(reader.pages[page_index])
    buffer = io.BytesIO()
    writer.write(buffer)
    return buffer.getvalue()


def _mineru_provider_factory(api_key: str | None, status):
    """provider_factory для run_pipeline: ключ из UI + метки этапов в st.status.

    У run_pipeline колбэка прогресса нет (и сигнатура заморожена спекой), поэтому
    этапы ловятся на швах провайдера: fetch_raw_zip и инжектируемый sleep.
    """
    def set_label(text: str) -> None:
        if status is not None:
            status.update(label=text)

    def build(model_version: str):
        # ponytail: sleep зовётся и в ретраях загрузки — «ожидание» может мигнуть
        # раньше времени. Косметика; точнее — только колбэком внутри ocr/.
        def sleep(seconds: float) -> None:
            set_label(f"MinerU ({model_version}): ожидание результата…")
            time.sleep(seconds)

        class _StatusProvider(MineruProvider):
            def fetch_raw_zip(self, pdf_bytes, page_range=None):
                set_label(f"MinerU ({model_version}): загрузка файла…")
                zip_bytes = super().fetch_raw_zip(pdf_bytes, page_range)
                set_label("Постобработка и проверка…")
                return zip_bytes

        return _StatusProvider(api_key, model_version=model_version, sleep=sleep)

    return build


def _split_findings(report: dict) -> tuple[list[dict], list[dict]]:
    """(обычные находки, low_confidence): вторых десятки, в UI они отдельно."""
    findings = report["findings"]
    return ([f for f in findings if f["rule"] != "low_confidence"],
            [f for f in findings if f["rule"] == "low_confidence"])


def _findings_table(findings: list[dict]) -> list[dict]:
    return [{"Правило": f["rule"], "Серьёзность": f["severity"],
             "Страница": f["page"], "Фрагмент": f["snippet"],
             "Предложение": f["suggestion"]} for f in findings]


def _render_ocr_report(report: dict) -> None:
    """Блок находок под результатом: сводка, список, low_confidence отдельно.

    ПРАВКА #92: подсказки vision_diff — отдельным блоком; vision_skipped остаётся в общем списке.
    """
    summary = report["summary"]
    main, low_conf = _split_findings(report)
    hints = [f for f in main if f["rule"] == _VISION_RULE]      # ПРАВКА #92
    main = [f for f in main if f["rule"] != _VISION_RULE]
    line = (f"Находки: critical — {summary['critical']}, "
            f"warning — {summary['warning']}, info — {summary['info']}"
            + (" (результат из кэша)" if report["cache_hit"] else ""))
    if summary["critical"]:
        st.error(line)
    else:
        st.info(line)
    if main:
        with st.expander(f"Находки ({len(main)})",
                         expanded=bool(summary["critical"])):
            st.dataframe(_findings_table(main), use_container_width=True,
                         hide_index=True)
    if low_conf:
        with st.expander(f"low_confidence — расхождения двух прогонов ({len(low_conf)})",
                         expanded=False):
            st.dataframe(_findings_table(low_conf), use_container_width=True,
                         hide_index=True)
    if hints:       # ПРАВКА #92
        with st.expander(f"Сверка по картинке — подсказки ({len(hints)})", expanded=False):
            st.caption(f"Модель {hints[0]['model']} переписала страницы скана; здесь — места, где прочитанное "
                       "расходится с текстом. Текст не изменён. Гомоглифы, пунктуация и тире не подсказываются.")
            st.dataframe(_vision_rows(hints), use_container_width=True, hide_index=True)


def _convert_uploaded_file(uploaded_file, page_range: str | None,
                           ocr_mode: str = "off", *, verify: bool = False,
                           vision: bool = False,                 # ПРАВКА #92
                           annotate: bool = True, status=None) -> dict:
    ext = _file_ext(uploaded_file.name)
    if ext != "pdf":
        # Диапазон страниц поддержан только для PDF: convert_with_markitdown
        # на непустом page_range для DOCX/XLSX/PPTX бросает ValueError.
        page_range = None
    display_range = page_range or "all"
    ocr_status = None
    report = None
    route = None        # ПРАВКА #85: маршрут ingest; None вне ветки mineru и при ранней ошибке
    tmp_path = _save_uploaded_to_temp(uploaded_file, ext)
    try:
        if ocr_mode == "mineru" and ext in _INGEST_EXTS:
            # ПРАВКА #73: весь тракт — в ocr/, здесь только сборка входа.
            # ПРАВКА #85: вход — ocr.ingest.ingest (раньше run_pipeline и только PDF):
            # DOCX/XLSX получают ту же проверку и report.json.
            data = uploaded_file.getvalue()
            if ext == "pdf":
                data = _pdf_page_subset(data, page_range)
            # ponytail: detect_route зовётся дважды (здесь и внутри ingest) — второй проход
            # pdfplumber по страницам. Убирается только параметром route у ingest, а он заморожен спекой 09.
            route = detect_route(data, uploaded_file.name)
            with tempfile.TemporaryDirectory() as work_dir:
                # verify на text/office не передаётся: ingest бросил бы ValueError на пачке файлов;
                # карточка результата пишет, что сверка не применялась.
                markdown, report = ingest(
                    data, source_name=uploaded_file.name, work_dir=Path(work_dir),
                    verify=verify and route in _MINERU_ROUTES, annotate=annotate,
                    cache=LocalCache(_OCR_CACHE_ROOT),
                    provider_factory=_mineru_provider_factory(_mineru_api_key(), status),
                    # ПРАВКА #92: сверка по картинке — только на маршрутах MinerU (на text/office ingest бросил бы ValueError)
                    vision=VISION_MODEL if vision and route in _MINERU_ROUTES else None,
                    vision_progress=_vision_progress(status))
        elif ext == "pdf":
            markdown, ocr_status = pdf_to_markdown_with_status(
                uploaded_file.getvalue(), page_range=page_range, mode=ocr_mode)
        else:
            markdown = convert_with_markitdown(tmp_path, page_range=page_range)
        return {
            "filename": uploaded_file.name,
            "download_name": _safe_md_filename(uploaded_file.name),
            "file_type": ext.upper() or "UNKNOWN",
            "page_range": display_range,
            "ocr_status": ocr_status,
            "markdown": markdown,
            "report": report,
            "route": route,     # ПРАВКА #85
            "error": None,
        }
    except Exception as e:
        error = str(e) or type(e).__name__
        if isinstance(e, MineruAuthError):
            error = f"{error}. {_MINERU_KEY_HINT}"
        return {
            "filename": uploaded_file.name,
            "download_name": _safe_md_filename(uploaded_file.name),
            "file_type": ext.upper() or "UNKNOWN",
            "page_range": display_range,
            "ocr_status": ocr_status,
            "markdown": "",
            "report": None,
            "route": route,     # ПРАВКА #85
            "error": error,
        }
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


# ПРАВКА #54: локальный фоллбэк шаблона письма лежит рядом с app.py,
# а не по абсолютному пути с чужого Desktop.
_LETTER_TEMPLATE_LOCAL = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "template_letter.docx")

DOC_TYPES = {
    # ПРАВКА #54: у каждого типа свой шаблон на Drive — раньше оба типа грузили
    # один file_id, и выбор типа менял только имя выходного файла и подсказку.
    # drive_id отсюда уходит в get_template -> download_template_from_drive.
    # ПРАВКА #59: "style" выбирает профиль оформления в convert.py (#55).
    # Ключ обязателен: отсутствие или опечатка роняет конвертацию, а не отдаёт
    # клиенту документ, оформленный не тем стилем.
    "📄 Письмо / Сопроводительное письмо": {
        "drive_id":    "1_E7eI5PgMD50MEI8RNl8xoiWmhUsOUap",
        "local_path":  _LETTER_TEMPLATE_LOCAL,
        "output_name": "letter",
        "style":       "letter",
        "hint": "Структура: `**Дата:**` и `**Исх.:**` — шапка слева, `**Кому:**` — адресат справа, `# ПИСЬМО`, тема курсивом `*О чём письмо*`, разделы `## 1. ...`, подпись `С уважением,` + должность + фамилия отдельными строками"
    },
    "📋 Пояснительная записка": {
        "drive_id":    "1FdPo8Ddo317ZYoPzraCTy5R4E72Ieqba",
        "local_path":  r"C:\Users\tonik\Desktop\docx_converter\template.docx",
        "output_name": "pz",
        "style":       "pz",
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
                md_text = _decode_md_upload(upl_md.read())
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
                use_drive = _drive_secrets_available()
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
                                           images=source_images,
                                           doc_style=config["style"])

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

    # ПРАВКА #73: вариант MinerU; OCRmyPDF — только при найденных бинарниках.
    ocr_mode = st.radio(
        "OCR mode",
        options=_ocr_mode_options(_ocrmypdf_available()),
        format_func=_OCR_MODE_LABELS.__getitem__,
        index=0,
        horizontal=True,
        key="files_to_md_ocr_mode",
        help=(
            "Без OCR: конвертация через MarkItDown. "
            "OCRmyPDF: OCR только для PDF без текстового слоя (нужны локальные "
            "Tesseract и Ghostscript). "
            # ПРАВКА #85: через тракт идут и DOCX/XLSX
            "MinerU: распознавание и проверка PDF/DOCX/XLSX с отчётом о сомнительных местах."
        ),
    )
    verify = False
    vision = False      # ПРАВКА #92
    annotate = True
    if ocr_mode == "off":
        st.caption("OCR выключен: используется текущий MarkItDown flow.")
    elif ocr_mode == "auto":
        st.caption("OCR auto включен: OCR применяется только к PDF-кандидатам.")
    else:
        # ПРАВКА #85: в облако идут не все PDF, а DOCX/XLSX получают отчёт
        st.caption(
            "В облако mineru.net уходят сканы и текстовые PDF с таблицами (до 200 МБ и 200 стр.). "
            "Текстовые PDF без таблиц, DOCX и XLSX конвертируются локально, но проходят ту же "
            "проверку и получают отчёт. PPTX — как обычно, без отчёта. При заданном диапазоне "
            "отправляются только выбранные страницы, и номера страниц в находках "
            "считаются от этой вырезки."
        )
        if not _mineru_api_key():
            st.warning(f"Ключ MinerU не найден. {_MINERU_KEY_HINT}")
        verify = st.checkbox(
            "Сверка вторым прогоном",
            value=False,
            key="files_to_md_mineru_verify",
            help="Второй прогон другой моделью MinerU, расхождения попадают в "
                 "находки. Удваивает время и расход квоты.",
        )
        if _claude_available():     # ПРАВКА #92: на Streamlit Cloud claude нет — галочки нет
            vision = st.checkbox(
                "Сверка по картинке (Claude Code, локально)",
                value=False,
                key="files_to_md_vision",
                help="Каждая страница скана или PDF с таблицами переписывается моделью "
                     f"{VISION_MODEL} через claude -p; расхождения с текстом — подсказки в находках, "
                     "текст не меняется. Одна страница — один вызов подписки Claude Code; повтор — из кэша.",
            )
        annotate = st.checkbox(
            "Пометки в тексте",
            value=True,
            key="files_to_md_mineru_annotate",
            help="Находки вставляются в Markdown как «!! ПРОВЕРИТЬ: … !!».",
        )

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
        key = range_keys[idx]
        if key not in st.session_state:
            st.session_state[key] = "all"
        selected_page_range = _normalize_page_range(st.session_state.get(key))
        with st.container(border=True):
            meta_col, range_col = st.columns([2, 1], vertical_alignment="top")
            with meta_col:
                st.markdown(f"**{uploaded_file.name}**")
                st.caption(f"Тип: .{ext or 'unknown'}")
                _display_pdf_diagnostics(uploaded_file, ext)
                _display_ocr_candidate_status(
                    uploaded_file,
                    ext,
                    ocr_mode,
                    selected_page_range,
                )
            with range_col:
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
                if ocr_mode == "mineru" and _file_ext(uploaded_file.name) in _INGEST_EXTS:
                    # ПРАВКА #73: минутное ожидание облака не должно выглядеть зависанием
                    # ПРАВКА #85: тот же статус для DOCX/XLSX, поэтому «OCR-тракт», не «MinerU»
                    with st.status(f"OCR-тракт: {uploaded_file.name}…") as status:
                        result = _convert_uploaded_file(
                            uploaded_file, page_range, ocr_mode=ocr_mode,
                            verify=verify, vision=vision,   # ПРАВКА #92
                            annotate=annotate, status=status)
                        if result["error"]:
                            status.update(
                                label=f"OCR-тракт: {uploaded_file.name} — ошибка",
                                state="error")
                        else:
                            status.update(
                                label=f"OCR-тракт: {uploaded_file.name} — готово",
                                state="complete")
                else:
                    with st.spinner(f"Конвертирую {uploaded_file.name}..."):
                        result = _convert_uploaded_file(
                            uploaded_file,
                            page_range,
                            ocr_mode=ocr_mode,
                        )
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
            ocr_status = result.get("ocr_status")
            if ocr_status:
                if ocr_status.get("status") == "applied":
                    st.info(ocr_status["message"])
                elif ocr_status.get("status") == "not_needed":
                    st.success(ocr_status["message"])
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
            # ПРАВКА #85: .get — в session_state могут лежать результаты без ключа route
            route = result.get("route")
            if route:
                st.caption(f"Маршрут: {_ROUTE_LABELS[route]}")
                if (st.session_state.get("files_to_md_mineru_verify")
                        and route not in _MINERU_ROUTES):
                    st.caption("Сверка вторым прогоном не применялась: файл не шёл через MinerU.")
                if (st.session_state.get("files_to_md_vision")      # ПРАВКА #92
                        and route not in _MINERU_ROUTES):
                    st.caption("Сверка по картинке не применялась: файл не шёл через MinerU.")
            report = result.get("report")
            if report:
                _render_ocr_report(report)
                if ANNOTATION_PREFIX in markdown:
                    st.caption("В тексте есть пометки «!! ПРОВЕРИТЬ: … !!» — "
                               "перед конвертацией в DOCX их нужно снять.")
            md_col, report_col = st.columns(2) if report else (st.container(), None)
            with md_col:
                st.download_button(
                    "Скачать .md",
                    data=markdown.encode("utf-8"),
                    file_name=result["download_name"],
                    mime="text/markdown",
                    key=f"download_md_{idx}_{result['download_name']}",
                    use_container_width=True,
                )
            if report:
                with report_col:
                    st.download_button(
                        "Скачать report.json",
                        data=json.dumps(report, ensure_ascii=False,
                                        indent=2).encode("utf-8"),
                        file_name=f"{Path(result['download_name']).stem}.report.json",
                        mime="application/json",
                        key=f"download_report_{idx}_{result['download_name']}",
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
