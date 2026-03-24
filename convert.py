import os
import re
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# =============================================================================
# НАСТРОЙКИ — МЕНЯЙ ТОЛЬКО ЭТИ ТРИ СТРОКИ
# =============================================================================
INPUT_FILE    = r"C:\Users\tonik\Desktop\docx_converter\Cheremhovo.md"       # ← твой .md файл
OUTPUT_FILE   = r"C:\Users\tonik\Desktop\docx_converter\Cheremhovo.docx"     # ← куда сохранить
TEMPLATE_FILE = r"C:\Users\tonik\Desktop\docx_converter\template.docx"   # ← шаблон с хедером
# =============================================================================

# === ЦВЕТОВАЯ ПАЛИТРА ===
BRAND_BLUE      = "015198"
BRAND_RED       = "D04514"
BRAND_ORANGE    = "EF7F1A"
BRAND_WHITE     = "FFFFFF"
TEXT_DARK       = "1A1A1A"
BG_LIGHT_BLUE   = "EBF3FB"
BG_LIGHT_ORANGE = "FFF8F0"
BG_TABLE_ROW    = "F0F5FA"
BORDER_LIGHT    = "CCCCCC"
TEXT_MUTED      = "888888"
COLOR_YES       = "1E7A34"   # зелёный для ✓
COLOR_NO        = "C0392B"   # красный для ✗

# Поля шаблона: left=2cm, right=1.5cm → рабочая ширина 17.5cm
CONTENT_WIDTH_CM = 17.5


# =============================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =============================================================================

def clear_body(doc):
    """Удаляет всё содержимое тела, сохраняя финальный sectPr."""
    body = doc.element.body
    to_remove = [c for c in body
                 if (c.tag.split('}')[-1] if '}' in c.tag else c.tag) != 'sectPr']
    for el in to_remove:
        body.remove(el)
    body.insert(0, OxmlElement('w:p'))


def set_cell_shading(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), hex_color)
    shd.set(qn('w:val'), 'clear')
    tcPr.append(shd)


def set_cell_margins_and_borders(cell, hex_color, sz):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for side, w in [('top','100'),('bottom','100'),('left','160'),('right','160')]:
        node = OxmlElement(f'w:{side}')
        node.set(qn('w:w'), w)
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top','left','bottom','right']:
        bdr = OxmlElement(f'w:{side}')
        bdr.set(qn('w:val'), 'single')
        bdr.set(qn('w:sz'), str(sz))
        bdr.set(qn('w:color'), hex_color)
        tcBorders.append(bdr)
    tcPr.append(tcBorders)


def set_cell_no_borders(cell):
    """Убирает все видимые границы ячейки (для таблиц-обёрток)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top','left','bottom','right','insideH','insideV']:
        bdr = OxmlElement(f'w:{side}')
        bdr.set(qn('w:val'), 'none')
        bdr.set(qn('w:sz'), '0')
        bdr.set(qn('w:color'), 'auto')
        tcBorders.append(bdr)
    tcPr.append(tcBorders)


def add_paragraph_border(paragraph, side, color, size, space="0"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = pPr.find(qn('w:pBdr'))
    if pBdr is None:
        pBdr = OxmlElement('w:pBdr')
        pPr.append(pBdr)
    border = OxmlElement(f'w:{side}')
    border.set(qn('w:val'), 'single')
    border.set(qn('w:sz'), str(size))
    border.set(qn('w:space'), str(space))
    border.set(qn('w:color'), color)
    pBdr.append(border)


def add_paragraph_shading(paragraph, hex_color):
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), hex_color)
    shd.set(qn('w:val'), 'clear')
    pPr.append(shd)


def set_table_width_dxa(table, width_cm):
    """Фиксирует ширину таблицы. 1cm = 567 DXA."""
    dxa = int(width_cm * 567)
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    existing = tblPr.find(qn('w:tblW'))
    if existing is not None:
        tblPr.remove(existing)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), str(dxa))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)


def set_table_no_spacing(table):
    """Убирает межтабличные отступы (для таблиц-обёрток)."""
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    spacing = OxmlElement('w:tblCellSpacing')
    spacing.set(qn('w:w'), '0')
    spacing.set(qn('w:type'), 'dxa')
    tblPr.append(spacing)


def set_row_height(row, height_dxa):
    """Устанавливает минимальную высоту строки таблицы."""
    trPr = row._tr.find(qn('w:trPr'))
    if trPr is None:
        trPr = OxmlElement('w:trPr')
        row._tr.insert(0, trPr)
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), str(height_dxa))
    trHeight.set(qn('w:hRule'), 'atLeast')
    trPr.append(trHeight)


def set_keep_with_next(paragraph):
    """Параграф остаётся на той же странице что и следующий."""
    pPr = paragraph._p.get_or_add_pPr()
    kwn = OxmlElement('w:keepNext')
    pPr.append(kwn)


def set_keep_together(paragraph):
    """Не разрывает параграф по страницам."""
    pPr = paragraph._p.get_or_add_pPr()
    kt = OxmlElement('w:keepLines')
    pPr.append(kt)


def set_run_font(run, name, size_pt, color_hex, bold=False, italic=False):
    run.font.name = name
    run.font.size = Pt(size_pt)
    run.font.color.rgb = RGBColor.from_string(color_hex)
    if bold:
        run.bold = True
    if italic:
        run.italic = True


def add_page_number_field(run):
    for ftype in ['begin', None, 'separate', 'end']:
        if ftype is None:
            el = OxmlElement('w:instrText')
            el.set(qn('xml:space'), 'preserve')
            el.text = 'PAGE'
        else:
            el = OxmlElement('w:fldChar')
            el.set(qn('w:fldCharType'), ftype)
        run._r.append(el)


# =============================================================================
# ДЕТЕКТОРЫ ТИПОВ БЛОКОВ
# =============================================================================

def is_stage_paragraph(text):
    return bool(re.match(r'^\*\*(Стадия|Фаза|Шаг|Этап|ВАЖНО)', text, re.IGNORECASE))

def is_photo_placeholder(text):
    return '📷' in text or '[Место для фото' in text

def is_requisites_block(text):
    return bool(re.match(r'^\*\*(Кому|От кого|Кому:|От кого:)', text))

def is_signature_block(text):
    return bool(re.match(r'^\*?С уважением', text))

def is_callout_block(text):
    """Блок !! текст !! — callout-врезка."""
    return text.startswith('!!') and text.endswith('!!')


# =============================================================================
# INLINE MARKDOWN ПАРСЕР
# =============================================================================

def parse_inline_markdown(paragraph, text, font_name='PT Sans', font_size=12,
                          font_color=TEXT_DARK, is_italic_base=False):
    """Обрабатывает **жирный** и *курсив*."""
    pattern = re.compile(r'(\*\*[^*\n]+?\*\*|\*(?!\*)[^*\n]+?\*(?!\*))')
    parts = pattern.split(text)
    for part in parts:
        if not part:
            continue
        run = paragraph.add_run()
        is_bold   = False
        is_italic = is_italic_base
        clean_text = part
        if part.startswith('**') and part.endswith('**') and len(part) >= 5:
            clean_text = part[2:-2]
            is_bold = True
        elif (part.startswith('*') and part.endswith('*')
              and len(part) >= 3 and not part.startswith('**')):
            clean_text = part[1:-1]
            is_italic = True
        set_run_font(run, font_name, font_size, font_color,
                     bold=is_bold, italic=is_italic)
        run.text = clean_text


# =============================================================================
# СПЕЦИАЛЬНЫЕ БЛОКИ-КОНСТРУКТОРЫ
# =============================================================================

def add_intro_paragraph(doc, block, content_width_cm):
    """
    ПРАВКА #4: Вводный абзац после H1 — таблица-обёртка с цветной левой полосой.
    Выглядит как акцентный callout для главной мысли документа.
    """
    table = doc.add_table(rows=1, cols=1)
    table.autofit = table.allow_autofit = False
    set_table_width_dxa(table, content_width_cm)
    set_table_no_spacing(table)

    cell = table.rows[0].cells[0]

    # Только левая граница — фирменный синий, 2pt
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top', 'bottom', 'right']:
        bdr = OxmlElement(f'w:{side}')
        bdr.set(qn('w:val'), 'none')
        bdr.set(qn('w:sz'), '0')
        bdr.set(qn('w:color'), 'auto')
        tcBorders.append(bdr)
    left_bdr = OxmlElement('w:left')
    left_bdr.set(qn('w:val'), 'single')
    left_bdr.set(qn('w:sz'), '18')   # ~2.25pt
    left_bdr.set(qn('w:space'), '4')
    left_bdr.set(qn('w:color'), BRAND_BLUE)
    tcBorders.append(left_bdr)
    tcPr.append(tcBorders)

    # Внутренние отступы ячейки
    tcMar = OxmlElement('w:tcMar')
    for side, w in [('top','80'),('bottom','80'),('left','220'),('right','0')]:
        node = OxmlElement(f'w:{side}')
        node.set(qn('w:w'), w)
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

    p = cell.paragraphs[0]
    p.paragraph_format.alignment  = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(0)
    # ПРАВКА #1: межстрочный 1.3
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    p.paragraph_format.line_spacing      = 1.3

    lines = block.split('\n')
    for i, line in enumerate(lines):
        line = line.strip()
        if not line: continue
        if i > 0: p.add_run().add_break()
        parse_inline_markdown(p, line)


def add_callout_box(doc, text, content_width_cm):
    """
    ПРАВКА #6: Callout-врезка !! текст !! — таблица с заливкой и бордером.
    Используется для формул, ключевых выводов, важных цифр.
    """
    clean = text.strip('!').strip()
    table = doc.add_table(rows=1, cols=1)
    table.autofit = table.allow_autofit = False
    set_table_width_dxa(table, content_width_cm)

    cell = table.rows[0].cells[0]
    set_cell_shading(cell, 'F2F6FA')   # очень лёгкий голубой

    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top','left','bottom','right']:
        bdr = OxmlElement(f'w:{side}')
        bdr.set(qn('w:val'), 'single')
        bdr.set(qn('w:sz'), '4')     # 0.5pt
        bdr.set(qn('w:color'), 'C5D8EC')
        tcBorders.append(bdr)
    # Акцент — левая граница чуть толще
    tcPr.append(tcBorders)

    tcMar = OxmlElement('w:tcMar')
    for side, w in [('top','140'),('bottom','140'),('left','220'),('right','220')]:
        node = OxmlElement(f'w:{side}')
        node.set(qn('w:w'), w)
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

    p = cell.paragraphs[0]
    p.paragraph_format.alignment  = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(0)
    parse_inline_markdown(p, clean, 'PT Sans', 12, BRAND_BLUE)
    for r in p.runs:
        r.bold = True


def add_table_cell_content(p, text, font_size=10):
    """
    ПРАВКА #8: Добавляет ✓/✗ перед значениями «Да»/«Нет» в ячейках таблицы.
    Применяется для любой сравнительной таблицы автоматически.
    """
    stripped = text.strip()

    # Проверяем начало ячейки на Да/Нет
    if re.match(r'^Да\b', stripped, re.IGNORECASE):
        icon_run = p.add_run('✓ ')
        set_run_font(icon_run, 'PT Sans', font_size, COLOR_YES, bold=True)
        rest = re.sub(r'^Да\b\s*', '', stripped, flags=re.IGNORECASE)
        if rest:
            parse_inline_markdown(p, rest, 'PT Sans', font_size, TEXT_DARK)
        else:
            run = p.add_run('Да')
            set_run_font(run, 'PT Sans', font_size, TEXT_DARK)
    elif re.match(r'^Нет\b', stripped, re.IGNORECASE):
        icon_run = p.add_run('✗ ')
        set_run_font(icon_run, 'PT Sans', font_size, COLOR_NO, bold=True)
        rest = re.sub(r'^Нет\b\s*', '', stripped, flags=re.IGNORECASE)
        if rest:
            parse_inline_markdown(p, rest, 'PT Sans', font_size, TEXT_DARK)
        else:
            run = p.add_run('Нет')
            set_run_font(run, 'PT Sans', font_size, TEXT_DARK)
    elif re.match(r'^Отсутствует\b', stripped, re.IGNORECASE):
        icon_run = p.add_run('✗ ')
        set_run_font(icon_run, 'PT Sans', font_size, COLOR_NO, bold=True)
        rest = re.sub(r'^Отсутствует\b\s*', '', stripped, flags=re.IGNORECASE)
        parse_inline_markdown(p, ('Отсутствует ' + rest).strip(), 'PT Sans', font_size, TEXT_DARK)
    else:
        parse_inline_markdown(p, stripped, 'PT Sans', font_size, TEXT_DARK)


# =============================================================================
# ОСНОВНАЯ ЛОГИКА КОНВЕРТАЦИИ
# =============================================================================

def convert_md_to_docx(md_text, output_filename, template_path=None):

    # --- Открываем шаблон или создаём чистый документ ---
    if template_path and os.path.exists(template_path):
        doc = Document(template_path)
        clear_body(doc)
        content_width_cm = CONTENT_WIDTH_CM
        print(f"  Шаблон: {template_path}")
    else:
        doc = Document()
        section = doc.sections[0]
        section.page_width    = Cm(21.0)
        section.page_height   = Cm(29.7)
        section.left_margin   = Cm(2.54)
        section.right_margin  = Cm(2.54)
        section.top_margin    = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        content_width_cm = 21.0 - 2.54 * 2
        print("  ⚠️  Шаблон не найден — хедер не будет добавлен")

        footer   = doc.sections[0].footer
        footer_p = footer.paragraphs[0]
        footer_p.paragraph_format.tab_stops.add_tab_stop(
            Cm(content_width_cm), WD_TAB_ALIGNMENT.RIGHT)
        rl = footer_p.add_run("ООО «ТПК «Тензосила»")
        set_run_font(rl, 'PT Sans', 9, TEXT_MUTED)
        footer_p.add_run("\t")
        rr = footer_p.add_run()
        set_run_font(rr, 'PT Sans', 9, TEXT_MUTED)
        add_page_number_field(rr)

    # --- Базовый стиль Normal ---
    sn = doc.styles['Normal']
    sn.font.name      = 'PT Sans'
    sn.font.size      = Pt(12)
    sn.font.color.rgb = RGBColor.from_string(TEXT_DARK)
    sn.paragraph_format.alignment         = WD_ALIGN_PARAGRAPH.JUSTIFY
    sn.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    sn.paragraph_format.line_spacing      = 1.3   # ПРАВКА #1: было 1.2
    sn.paragraph_format.space_after       = Pt(8)  # ПРАВКА #2: было Pt(6)

    # --- Парсинг блоков Markdown ---
    blocks = md_text.split('\n\n')

    after_heading         = False
    intro_done            = False   # ПРАВКА #4: флаг первого абзаца после H1
    last_list_paragraph   = None
    last_regular_paragraph = None  # ПРАВКА #10: для keep_with_next перед подписью

    for block in blocks:
        block = block.strip()
        if not block:
            continue

        is_list_item = (block.startswith('- ')
                        or block.startswith('* ')
                        or bool(re.match(r'^\d+\.', block)))
        if not is_list_item and last_list_paragraph:
            last_list_paragraph.paragraph_format.space_after = Pt(10)
            last_list_paragraph = None

        # ── H1 ───────────────────────────────────────────────────────────────
        if block.startswith('# '):
            p = doc.add_paragraph()
            p.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(24)
            p.paragraph_format.space_after  = Pt(12)
            parse_inline_markdown(p, block[2:], 'PT Sans Narrow', 18, BRAND_BLUE)
            for r in p.runs: r.bold = True
            dec = doc.add_paragraph()
            dec.paragraph_format.space_before = Pt(0)
            dec.paragraph_format.space_after  = Pt(10)
            add_paragraph_border(dec, 'bottom', BRAND_RED, 12)
            after_heading = True
            intro_done = False   # сбрасываем при каждом H1

        # ── H2 ───────────────────────────────────────────────────────────────
        elif block.startswith('## '):
            p = doc.add_paragraph()
            p.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(20)
            p.paragraph_format.space_after  = Pt(8)
            # ПРАВКА #3: H2 остаётся цветным 14pt
            parse_inline_markdown(p, block[3:], 'PT Sans Narrow', 14, BRAND_RED)
            for r in p.runs: r.bold = True
            after_heading = True

        # ── H3 ───────────────────────────────────────────────────────────────
        elif block.startswith('### '):
            p = doc.add_paragraph()
            p.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after  = Pt(6)
            # ПРАВКА #3: H3 — чёрный bold 11.5pt (отличается от H2 цветом и размером)
            parse_inline_markdown(p, block[4:], 'PT Sans Narrow', 11, TEXT_DARK)
            for r in p.runs: r.bold = True
            after_heading = True

        # ── Цитаты > ─────────────────────────────────────────────────────────
        elif block.startswith('>'):
            clean = '\n'.join([l.lstrip('> ') for l in block.split('\n')])
            p = doc.add_paragraph()
            # ПРАВКА #5: увеличен отступ, лёгкая заливка
            p.paragraph_format.left_indent  = Cm(1.8)
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after  = Pt(10)
            add_paragraph_border(p, 'left', BRAND_ORANGE, 18, space=4)
            add_paragraph_shading(p, 'F7F7F7')
            parse_inline_markdown(p, clean, 'PT Sans', 12, "555555", is_italic_base=True)
            last_regular_paragraph = p
            after_heading = False

        # ── Callout-врезка !! текст !! ────────────────────────────────────────
        elif is_callout_block(block):
            # ПРАВКА #6: новый тип блока — оформляется как акцентная таблица
            add_callout_box(doc, block, content_width_cm)
            after_heading = False

        # ── Плейсхолдеры фото ────────────────────────────────────────────────
        elif is_photo_placeholder(block):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent  = Cm(1.0)
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after  = Pt(10)
            add_paragraph_border(p, 'left', BRAND_ORANGE, 18)
            add_paragraph_shading(p, BG_LIGHT_ORANGE)
            parse_inline_markdown(p, block.replace('>', '').strip(),
                                  'PT Sans', 11, "999999", is_italic_base=True)
            after_heading = False

        # ── Стадии / ВАЖНО ───────────────────────────────────────────────────
        elif is_stage_paragraph(block):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent  = Cm(0.75)
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after  = Pt(6)
            add_paragraph_border(p, 'left', BRAND_BLUE, 12)
            add_paragraph_shading(p, BG_LIGHT_BLUE)
            parse_inline_markdown(p, block)
            after_heading = False

        # ── Списки ───────────────────────────────────────────────────────────
        elif is_list_item:
            for line in block.split('\n'):
                line = line.strip()
                if not line: continue
                is_num = bool(re.match(r'^\d+\.', line))
                style_name = 'List Number' if is_num else 'List Bullet'
                # Создаём стиль если его нет в шаблоне
                if style_name not in [s.name for s in doc.styles]:
                    from docx.enum.style import WD_STYLE_TYPE
                    new_style = doc.styles.add_style(
                        style_name, WD_STYLE_TYPE.PARAGRAPH
                    )
                    new_style.base_style = doc.styles['Normal']
                p = doc.add_paragraph(style=style_name)
                p.paragraph_format.left_indent       = Cm(1.5)
                p.paragraph_format.first_line_indent = Cm(-0.75)
                p.paragraph_format.space_after       = Pt(4)
                parse_inline_markdown(p, re.sub(r'^(\- |\* |\d+\. )', '', line))
                last_list_paragraph = p
                last_regular_paragraph = p
            after_heading = False

        # ── Таблицы ──────────────────────────────────────────────────────────
        elif block.startswith('|') and '\n|' in block:
            lines = [l.strip() for l in block.split('\n')
                     if l.strip() and not re.match(r'^\|[-| ]+\|$', l.strip())]
            if not lines: continue

            headers  = [c.strip() for c in lines[0].strip('|').split('|')]
            n_cols   = len(headers)
            table    = doc.add_table(rows=1, cols=n_cols)
            table.autofit = table.allow_autofit = False
            set_table_width_dxa(table, content_width_cm)

            # Пропорции колонок: 35/30/35 для 3-колоночных, равные для остальных
            ratios     = [0.35, 0.30, 0.35] if n_cols == 3 else [1/n_cols]*n_cols
            total_dxa  = int(content_width_cm * 567)
            widths_dxa = [int(content_width_cm * 567 * r) for r in ratios]
            widths_dxa[-1] = total_dxa - sum(widths_dxa[:-1])
            for i, col in enumerate(table.columns):
                col.width = widths_dxa[i]

            # ПРАВКА #7: заголовочная строка — увеличена высота и шрифт
            set_row_height(table.rows[0], 560)   # ~1cm минимальная высота

            for i, h in enumerate(headers):
                if i < len(table.rows[0].cells):
                    cell = table.rows[0].cells[i]
                    set_cell_shading(cell, BRAND_BLUE)
                    set_cell_margins_and_borders(cell, BORDER_LIGHT, 4)
                    p = cell.paragraphs[0]
                    p.paragraph_format.alignment   = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_after = Pt(0)
                    # ПРАВКА #7: шрифт 10pt → 11pt в заголовке таблицы
                    parse_inline_markdown(p, h, 'PT Sans Narrow', 11, BRAND_WHITE)
                    for r in p.runs: r.bold = True

            for row_idx, line in enumerate(lines[1:]):
                cols      = [c.strip() for c in line.strip('|').split('|')]
                row_cells = table.add_row().cells
                is_even   = (row_idx % 2 == 1)

                for i, c in enumerate(cols):
                    if i < len(row_cells):
                        cell = row_cells[i]

                        # ПРАВКА #9: последняя колонка 3-колоночной таблицы
                        # получает фирменный голубой фон (выделяем «наш» столбец)
                        if n_cols == 3 and i == n_cols - 1:
                            bg = BG_LIGHT_BLUE
                        else:
                            bg = BG_TABLE_ROW if is_even else BRAND_WHITE

                        set_cell_shading(cell, bg)
                        set_cell_margins_and_borders(cell, BORDER_LIGHT, 4)
                        p = cell.paragraphs[0]
                        p.paragraph_format.space_after = Pt(0)
                        # ПРАВКА #8: автоматические ✓/✗ для Да/Нет/Отсутствует
                        add_table_cell_content(p, c, font_size=10)

            # Компактный отступ после таблицы
            sp = doc.add_paragraph()
            pPr = sp._p.get_or_add_pPr()
            s = OxmlElement('w:spacing')
            s.set(qn('w:before'), '0')
            s.set(qn('w:after'), '120')
            s.set(qn('w:line'), '120')
            s.set(qn('w:lineRule'), 'exact')
            pPr.append(s)
            after_heading = False

        # ── Разделители --- ───────────────────────────────────────────────────
        elif block.startswith('---'):
            continue

        # ── Реквизиты «Кому / От кого» ───────────────────────────────────────
        elif is_requisites_block(block):
            p = doc.add_paragraph()
            p.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after  = Pt(14)
            p.paragraph_format.left_indent  = Cm(0.4)
            p.paragraph_format.right_indent = Cm(0.4)
            add_paragraph_shading(p, BG_LIGHT_BLUE)
            for i, line in enumerate(block.split('\n')):
                line = line.strip()
                if not line: continue
                if i > 0: p.add_run().add_break()
                parse_inline_markdown(p, line)
            last_regular_paragraph = p
            after_heading = False

        # ── Подпись «С уважением» ─────────────────────────────────────────────
        elif is_signature_block(block):
            # ПРАВКА #10: последний абзац перед подписью держим вместе с ней
            if last_regular_paragraph is not None:
                set_keep_with_next(last_regular_paragraph)

            p = doc.add_paragraph()
            p.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(28)
            p.paragraph_format.space_after  = Pt(0)
            add_paragraph_border(p, 'top', BRAND_RED, 4, space=8)
            # ПРАВКА #11: весь блок подписи не разрывается по страницам
            set_keep_together(p)
            for i, line in enumerate(block.split('\n')):
                line = line.strip()
                if not line: continue
                if i > 0: p.add_run().add_break()
                parse_inline_markdown(p, line)
            after_heading = False

        # ── Обычные абзацы ────────────────────────────────────────────────────
        else:
            # ПРАВКА #4: первый абзац после H1 — вводный с левой полосой
            if after_heading and not intro_done:
                add_intro_paragraph(doc, block, content_width_cm)
                intro_done = True
                after_heading = False
                # Добавляем пустой параграф-отступ после врезки
                sp = doc.add_paragraph()
                pPr = sp._p.get_or_add_pPr()
                s = OxmlElement('w:spacing')
                s.set(qn('w:before'), '0')
                s.set(qn('w:after'), '120')
                s.set(qn('w:line'), '120')
                s.set(qn('w:lineRule'), 'exact')
                pPr.append(s)
                continue

            p = doc.add_paragraph()
            if not after_heading:
                p.paragraph_format.first_line_indent = Cm(0.75)
            for i, line in enumerate(block.split('\n')):
                line = line.strip()
                if not line: continue
                if i > 0: p.add_run().add_break()
                parse_inline_markdown(p, line)
            last_regular_paragraph = p
            after_heading = False

    doc.save(output_filename)
    print(f"✅ Готово! Файл сохранён: {output_filename}")


# =============================================================================
# ЗАПУСК
# =============================================================================
if __name__ == '__main__':
    if not os.path.exists(INPUT_FILE):
        print(f"❌ .md файл не найден: {INPUT_FILE}")
        print("   Проверь путь INPUT_FILE в начале скрипта.")
        exit(1)

    if not os.path.exists(TEMPLATE_FILE):
        print(f"⚠️  Шаблон не найден: {TEMPLATE_FILE}")
        print("   Документ будет создан без фирменного хедера.\n")

    print(f"Конвертирую: {os.path.basename(INPUT_FILE)}")
    with open(INPUT_FILE, 'r', encoding='utf-8') as f:
        md_text = f.read()

    template = TEMPLATE_FILE if os.path.exists(TEMPLATE_FILE) else None
    convert_md_to_docx(md_text, OUTPUT_FILE, template_path=template)
