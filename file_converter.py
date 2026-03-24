"""
file_converter.py
Извлекает текст и структуру из DOCX / PDF / TXT
и возвращает:
  - md_text: строка в формате Markdown
  - images: список (filename, bytes) для вставки в итоговый документ
"""

import io
import re
from typing import Optional


# =============================================================================
# DOCX → MD
# =============================================================================

def docx_to_md(file_bytes: bytes) -> tuple[str, list]:
    """
    Конвертирует DOCX в Markdown.
    Заголовки определяются по:
      1. Стилю Word (Heading 1 / 2 / 3)
      2. Жирный + крупный шрифт (если стилей нет)
    Возвращает (md_text, images).
    images = список (image_filename, image_bytes)
    """
    from docx import Document
    from docx.shared import Pt

    doc = Document(io.BytesIO(file_bytes))
    lines = []
    images = []
    image_counter = [0]

    # Собираем изображения из документа
    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            try:
                img_bytes = rel.target_part.blob
                ext = rel.target_part.content_type.split('/')[-1]
                if ext == 'jpeg':
                    ext = 'jpg'
                fname = f"image_{image_counter[0]}.{ext}"
                images.append((fname, img_bytes))
                image_counter[0] += 1
            except Exception:
                pass

    # Среднее значение размера шрифта для эвристики
    font_sizes = []
    for para in doc.paragraphs:
        for run in para.runs:
            if run.font.size:
                font_sizes.append(run.font.size.pt)
    avg_size = sum(font_sizes) / len(font_sizes) if font_sizes else 12.0

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            lines.append('')
            continue

        style_name = para.style.name if para.style else ''

        # --- Определяем уровень заголовка ---
        heading_level = 0

        # Способ 1: по стилю Word
        if 'Heading 1' in style_name or style_name == 'Title':
            heading_level = 1
        elif 'Heading 2' in style_name:
            heading_level = 2
        elif 'Heading 3' in style_name:
            heading_level = 3

        # Способ 2: эвристика — жирный + крупный шрифт
        if heading_level == 0:
            run_sizes = [r.font.size.pt for r in para.runs if r.font.size]
            run_bolds = [r.bold for r in para.runs if r.text.strip()]
            is_bold = all(run_bolds) and len(run_bolds) > 0
            max_size = max(run_sizes) if run_sizes else avg_size

            if is_bold and max_size >= avg_size * 1.5:
                heading_level = 1
            elif is_bold and max_size >= avg_size * 1.25:
                heading_level = 2
            elif is_bold and max_size >= avg_size * 1.1:
                heading_level = 3

        # --- Формируем MD строку ---
        if heading_level == 1:
            lines.append(f'# {text}')
        elif heading_level == 2:
            lines.append(f'## {text}')
        elif heading_level == 3:
            lines.append(f'### {text}')
        else:
            # Обычный текст — сохраняем жирный/курсив
            md_text = _runs_to_md(para.runs)
            if md_text.strip():
                lines.append(md_text)

        lines.append('')  # пустая строка между абзацами

    return '\n'.join(lines), images


def _runs_to_md(runs) -> str:
    """Переводит runs параграфа в MD с сохранением bold/italic."""
    result = ''
    for run in runs:
        text = run.text
        if not text:
            continue
        if run.bold and run.italic:
            result += f'***{text}***'
        elif run.bold:
            result += f'**{text}**'
        elif run.italic:
            result += f'*{text}*'
        else:
            result += text
    return result


# =============================================================================
# PDF → MD
# =============================================================================

def pdf_to_md(file_bytes: bytes) -> tuple[str, list]:
    """
    Конвертирует PDF в Markdown.
    Заголовки определяются по размеру шрифта относительно среднего.
    Изображения извлекаются через pdfplumber (если возможно).
    """
    try:
        import pdfplumber
    except ImportError:
        return _pdf_fallback(file_bytes), []

    lines = []
    images = []
    image_counter = [0]

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        # Шаг 1: определяем средний размер шрифта по всему документу
        all_sizes = []
        for page in pdf.pages:
            for word in page.extract_words(extra_attrs=['size', 'fontname']):
                try:
                    all_sizes.append(float(word.get('size', 12)))
                except (ValueError, TypeError):
                    pass
        avg_size = sum(all_sizes) / len(all_sizes) if all_sizes else 12.0

        # Шаг 2: извлекаем текст постранично
        prev_text = ''
        for page_num, page in enumerate(pdf.pages):
            words = page.extract_words(
                extra_attrs=['size', 'fontname'],
                x_tolerance=3,
                y_tolerance=3
            )

            # Группируем слова в строки по Y-координате
            lines_on_page = {}
            for word in words:
                y_key = round(float(word.get('top', 0)), 0)
                if y_key not in lines_on_page:
                    lines_on_page[y_key] = []
                lines_on_page[y_key].append(word)

            for y_key in sorted(lines_on_page.keys()):
                line_words = lines_on_page[y_key]
                text = ' '.join(w['text'] for w in line_words).strip()
                if not text or text == prev_text:
                    continue
                prev_text = text

                # Определяем размер шрифта строки
                sizes = []
                for w in line_words:
                    try:
                        sizes.append(float(w.get('size', avg_size)))
                    except (ValueError, TypeError):
                        sizes.append(avg_size)
                line_size = max(sizes) if sizes else avg_size

                # Определяем жирность по имени шрифта
                fontnames = [w.get('fontname', '') for w in line_words]
                is_bold = any('Bold' in f or 'bold' in f or 'Heavy' in f
                              for f in fontnames)

                # Определяем уровень заголовка
                ratio = line_size / avg_size if avg_size > 0 else 1.0

                if ratio >= 1.5 or (ratio >= 1.3 and is_bold):
                    lines.append(f'# {text}')
                elif ratio >= 1.25 or (ratio >= 1.1 and is_bold):
                    lines.append(f'## {text}')
                elif ratio >= 1.1 and is_bold:
                    lines.append(f'### {text}')
                else:
                    lines.append(text)

                lines.append('')

            # Извлекаем изображения со страницы
            try:
                for img in page.images:
                    img_bytes = _extract_pdf_image(page, img)
                    if img_bytes:
                        fname = f"image_p{page_num}_{image_counter[0]}.png"
                        images.append((fname, img_bytes))
                        image_counter[0] += 1
            except Exception:
                pass

    return '\n'.join(lines), images


def _extract_pdf_image(page, img_dict) -> Optional[bytes]:
    """Пытается извлечь изображение из PDF страницы."""
    try:
        x0 = img_dict['x0']
        y0 = img_dict['y0']
        x1 = img_dict['x1']
        y1 = img_dict['y1']
        cropped = page.crop((x0, y0, x1, y1))
        img_obj = cropped.to_image(resolution=150)
        buf = io.BytesIO()
        img_obj.save(buf, format='PNG')
        return buf.getvalue()
    except Exception:
        return None


def _pdf_fallback(file_bytes: bytes) -> str:
    """Запасной вариант если pdfplumber не установлен."""
    try:
        import pypdf
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        lines = []
        for page in reader.pages:
            text = page.extract_text()
            if text:
                lines.append(text)
                lines.append('')
        return '\n'.join(lines)
    except Exception:
        return "❌ Не удалось извлечь текст из PDF. Установите pdfplumber."


# =============================================================================
# TXT → MD
# =============================================================================

def txt_to_md(file_bytes: bytes) -> tuple[str, list]:
    """
    Конвертирует TXT в Markdown.
    В TXT нет форматирования — просто разбиваем на абзацы.
    Пустые строки = разделители абзацев.
    """
    try:
        text = file_bytes.decode('utf-8')
    except UnicodeDecodeError:
        text = file_bytes.decode('cp1251', errors='replace')

    # Нормализуем переносы строк
    text = text.replace('\r\n', '\n').replace('\r', '\n')

    # Разбиваем на абзацы по двойному переносу
    paragraphs = re.split(r'\n{2,}', text)
    lines = []
    for para in paragraphs:
        para = para.strip()
        if para:
            # Объединяем одиночные переносы внутри абзаца в пробел
            para = re.sub(r'\n', ' ', para)
            lines.append(para)
            lines.append('')

    return '\n'.join(lines), []


# =============================================================================
# УНИВЕРСАЛЬНЫЙ ВХОД
# =============================================================================

def convert_file_to_md(file_bytes: bytes, filename: str) -> tuple[str, list]:
    """
    Главная функция — определяет формат по расширению и вызывает нужный конвертер.
    Возвращает (md_text, images).
    """
    ext = filename.lower().split('.')[-1]

    if ext == 'docx':
        return docx_to_md(file_bytes)
    elif ext == 'pdf':
        return pdf_to_md(file_bytes)
    elif ext in ('txt', 'text'):
        return txt_to_md(file_bytes)
    else:
        raise ValueError(f"Неподдерживаемый формат: .{ext}")
