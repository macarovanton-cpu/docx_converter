import io
import os
import re
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE as RT

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

# ПРАВКА #17: авто-замена ячеек «Да» → «✓ Да», «Нет»/«Отсутствует» → «✗ ...»
# Установи False, если хочешь сохранять текст ячеек как есть.
ENABLE_TABLE_SYMBOLS = True

# Поля шаблона: left=2cm, right=1.5cm → рабочая ширина 17.5cm
CONTENT_WIDTH_CM = 17.5


# =============================================================================
# ПРАВКА #55: ПРОФИЛИ ОФОРМЛЕНИЯ
# =============================================================================
# Письмо и пояснительная записка — разные документы с разными шаблонами (#54),
# а рисовались одинаково. Профиль — плоский словарь: значения (цвет, кегль,
# шрифт, заливка) и флаги, меняющие структуру вывода. Выбирается аргументом
# doc_style и доезжает до функций явным параметром style, а не глобалом:
# Streamlit обслуживает сессии потоками одного процесса, и модульная переменная
# при двух одновременных конвертациях разных типов молча дала бы клиенту письмо
# в стиле ПЗ.
# Правило доступа: только style['key'], никогда style.get() — .get вернёт None,
# и Pt(None) уронит рендер далеко от места опечатки.

STYLE_PZ = {
    # --- шрифты и кегли ---
    'body_font':        'PT Sans',
    'body_size':        12,
    'body_line':        1.3,             # интерлиньяж
    'body_after':       8,               # Pt, space_after в стиле Normal
    'head_font':        'PT Sans Narrow',
    'h1_size':          18,
    'h2_size':          14,
    'h3_size':          13,
    'h4_size':          12,              # H4-H6 рисуются одинаково (#40)
    'photo_size':       11,
    'table_head_size':  11,
    'table_cell_size':  10,

    # --- цвета ---
    'text_color':       TEXT_DARK,
    'h1_color':         BRAND_BLUE,
    'h2_color':         BRAND_RED,
    'link_color':       BRAND_BLUE,
    'accent_color':     BRAND_BLUE,      # левые полосы intro и стадий
    'rule_color':       BRAND_RED,       # линия под H1, над подписью, под шапкой письма
    'quote_color':      BRAND_ORANGE,
    'quote_text':       '555555',
    'quote_fill':       'F7F7F7',
    'photo_color':      BRAND_ORANGE,
    'photo_text':       '999999',
    'photo_fill':       BG_LIGHT_ORANGE,
    'block_fill':       BG_LIGHT_BLUE,   # реквизиты и стадии
    'callout_fill':     'F2F6FA',
    'callout_border':   'C5D8EC',
    'callout_text':     BRAND_BLUE,
    'table_head_fill':  BRAND_BLUE,
    'table_head_text':  BRAND_WHITE,
    'table_row_fill':   BG_TABLE_ROW,
    'table_alt_fill':   BG_LIGHT_BLUE,   # последняя колонка 3-колоночной таблицы
    'icon_yes':         COLOR_YES,
    'icon_no':          COLOR_NO,

    # --- декор и раскладка ---
    'h1_center':        False,
    'h1_rule':          True,            # декоративная линия под H1
    'intro_band':       True,            # False → тема курсивом по центру
    'para_indent':      0.75,            # красная строка, см
    'requisites_fill':  True,
    'requisites_right': False,
    'signature_rule':   True,
    'signature_tab':    False,           # True → должность/фамилия в одну строку
    'header_table':     False,           # True → шапка «Дата/Исх. + Кому» таблицей
    'header_left_cm':   6.0,             # ширина левой колонки шапки
    'bullet_char':      '•',
    'bullet_font':      'Symbol',
}

STYLE_LETTER = {
    # --- шрифты и кегли: тело плотнее и мельче, чем в ПЗ ---
    'body_font':        'PT Sans',
    'body_size':        10.5,
    'body_line':        1.15,
    'body_after':       6,
    'head_font':        'PT Sans',       # бланк без контрастных шрифтов
    'h1_size':          12,              # «ПИСЬМО» чуть крупнее тела
    'h2_size':          10.5,            # разделы — кегль тела, только полужирный
    'h3_size':          10.5,
    'h4_size':          10.5,
    'photo_size':       10,
    'table_head_size':  10,
    'table_cell_size':  9.5,

    # --- цвета: строгий бланк, ни одного фирменного цвета ---
    'text_color':       '000000',
    'h1_color':         '000000',
    'h2_color':         '000000',
    'link_color':       '000000',        # признак ссылки — подчёркивание
    'accent_color':     '000000',
    'rule_color':       '999999',
    'quote_color':      'BBBBBB',
    'quote_text':       '000000',
    'quote_fill':       BRAND_WHITE,
    'photo_color':      'BBBBBB',
    'photo_text':       '666666',
    'photo_fill':       BRAND_WHITE,
    'block_fill':       BRAND_WHITE,
    'callout_fill':     BRAND_WHITE,
    'callout_border':   '999999',
    'callout_text':     '000000',
    'table_head_fill':  BRAND_WHITE,
    'table_head_text':  '000000',
    'table_row_fill':   BRAND_WHITE,
    'table_alt_fill':   BRAND_WHITE,
    'icon_yes':         '000000',
    'icon_no':          '000000',

    # --- декор и раскладка ---
    'h1_center':        True,
    'h1_rule':          False,
    'intro_band':       False,
    'para_indent':      0,               # в деловом бланке красной строки нет
    'requisites_fill':  False,
    'requisites_right': True,
    'signature_rule':   False,
    'signature_tab':    True,
    'header_table':     True,
    'header_left_cm':   6.0,
    'bullet_char':      '—',
    'bullet_font':      'PT Sans',
}

STYLES = {'pz': STYLE_PZ, 'letter': STYLE_LETTER}

# ПРАВКА #55: профили обязаны иметь одинаковый набор ключей. Не assert —
# assert вырезается под python -O, а профиль с дырой роняет рендер в
# неочевидном месте.
if STYLE_PZ.keys() != STYLE_LETTER.keys():
    raise RuntimeError('ПРАВКА #55: наборы ключей профилей разошлись: '
                       f'{sorted(STYLE_PZ.keys() ^ STYLE_LETTER.keys())}')


# =============================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =============================================================================

# ПРАВКА #50: порядок дочерних элементов в OXML задаёт схема, а не порядок
# вызовов: append ставил shd/tcMar/tcBorders после tcW, keepNext после jc и так
# далее. python-docx знает эти последовательности, но удаляет их из своих
# классов (`del _tag_seq`), поэтому нужные держим здесь.
# Для settings перечислен только хвост от autoHyphenation — других элементов
# этот файл в settings.xml не вставляет.
_CHILD_ORDER = {
    'pPr': (
        'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr '
        'suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap '
        'overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd '
        'snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap '
        'jc textDirection textAlignment textboxTightWrap outlineLvl divId '
        'cnfStyle rPr sectPr pPrChange').split(),
    'tcPr': (
        'cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar '
        'textDirection tcFitText vAlign hideMark headers cellIns cellDel '
        'cellMerge tcPrChange').split(),
    'tblPr': (
        'tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize '
        'tblStyleColBandSize tblW jc tblCellSpacing tblInd tblBorders shd '
        'tblLayout tblCellMar tblLook tblCaption tblDescription '
        'tblPrChange').split(),
    'numbering': 'numPicBullet abstractNum num numIdMacAtCleanup'.split(),
    'settings': (
        'autoHyphenation consecutiveHyphenLimit hyphenationZone '
        'doNotHyphenateCaps showEnvelope summaryLength clickAndTypeStyle '
        'defaultTableStyle evenAndOddHeaders bookFoldRevPrinting '
        'bookFoldPrinting bookFoldPrintingSheets drawingGridHorizontalSpacing '
        'drawingGridVerticalSpacing displayHorizontalDrawingGridEvery '
        'displayVerticalDrawingGridEvery doNotUseMarginsForDrawingGridOrigin '
        'drawingGridHorizontalOrigin drawingGridVerticalOrigin '
        'doNotShadeFormData noPunctuationKerning characterSpacingControl '
        'printTwoOnOne strictFirstAndLastChars noLineBreaksAfter '
        'noLineBreaksBefore savePreviewPicture doNotValidateAgainstSchema '
        'saveInvalidXml ignoreMixedContent alwaysShowPlaceholderText '
        'doNotDemarcateInvalidXml saveXmlDataOnly useXSLTWhenSaving '
        'saveThroughXslt showXMLTags alwaysMergeEmptyNamespace updateFields '
        'hdrShapeDefaults footnotePr endnotePr compat docVars rsids mathPr '
        'attachedSchema themeFontLang clrSchemeMapping '
        'doNotIncludeSubdocsInStats doNotAutoCompressPictures forceUpgrade '
        'captions readModeInkLockDown smartTagType shapeDefaults '
        'doNotEmbedSmartTags decimalSymbol listSeparator').split(),
}


def insert_in_order(parent, child):
    """ПРАВКА #50: ставит элемент туда, где его ждёт схема OOXML.
    Родитель или тег без известного порядка — фолбэк на append (как до #50),
    чтобы будущие вызовы с непокрытыми тегами не падали."""
    order = _CHILD_ORDER.get(parent.tag.split('}')[-1])
    tag = child.tag.split('}')[-1]
    if order is None or tag not in order:
        parent.append(child)
        return child
    successors = order[order.index(tag) + 1:]
    return parent.insert_element_before(child, *('w:' + s for s in successors))


def clear_body(doc):
    """Удаляет всё содержимое тела, сохраняя финальный sectPr."""
    body = doc.element.body
    to_remove = [c for c in body
                 if (c.tag.split('}')[-1] if '}' in c.tag else c.tag) != 'sectPr']
    # ПРАВКА #43: пустой w:p здесь не нужен — python-docx сам вставляет
    # абзацы перед sectPr, а лишний давал провал над первым заголовком
    # в дополнение к его собственному отступу 24pt.
    for el in to_remove:
        body.remove(el)


def set_cell_shading(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), hex_color)
    shd.set(qn('w:val'), 'clear')
    insert_in_order(tcPr, shd)


def set_cell_margins_and_borders(cell, hex_color, sz):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    # ПРАВКА #51: порядок сторон задан схемой — top, left, bottom, right
    for side, w in [('top','100'),('left','160'),('bottom','100'),('right','160')]:
        node = OxmlElement(f'w:{side}')
        node.set(qn('w:w'), w)
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    insert_in_order(tcPr, tcMar)
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top','left','bottom','right']:
        bdr = OxmlElement(f'w:{side}')
        bdr.set(qn('w:val'), 'single')
        bdr.set(qn('w:sz'), str(sz))
        bdr.set(qn('w:color'), hex_color)
        tcBorders.append(bdr)
    insert_in_order(tcPr, tcBorders)


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
    insert_in_order(tcPr, tcBorders)


def add_paragraph_border(paragraph, side, color, size, space="0"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = pPr.find(qn('w:pBdr'))
    if pBdr is None:
        pBdr = OxmlElement('w:pBdr')
        insert_in_order(pPr, pBdr)
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
    insert_in_order(pPr, shd)


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
    insert_in_order(tblPr, tblW)


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
    insert_in_order(tblPr, spacing)


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


def set_row_flag(row, tag):
    """ПРАВКА #35: булев флаг строки таблицы: 'w:tblHeader' | 'w:cantSplit'."""
    trPr = row._tr.find(qn('w:trPr'))
    if trPr is None:
        trPr = OxmlElement('w:trPr')
        row._tr.insert(0, trPr)
    trPr.append(OxmlElement(tag))


def set_keep_with_next(paragraph):
    """Параграф остаётся на той же странице что и следующий."""
    pPr = paragraph._p.get_or_add_pPr()
    kwn = OxmlElement('w:keepNext')
    insert_in_order(pPr, kwn)


def set_keep_together(paragraph):
    """Не разрывает параграф по страницам."""
    pPr = paragraph._p.get_or_add_pPr()
    kt = OxmlElement('w:keepLines')
    insert_in_order(pPr, kt)


def set_run_font(run, name, size_pt, color_hex, bold=False, italic=False):
    run.font.name = name
    run.font.size = Pt(size_pt)
    run.font.color.rgb = RGBColor.from_string(color_hex)
    if bold:
        run.bold = True
    if italic:
        run.italic = True


# ПРАВКА #21: обработка markdown-ссылок [text](url)
def add_hyperlink_run(paragraph, url, text, font_name='PT Sans', font_size=12,
                      bold=False, italic=False, link_color=None):
    # ПРАВКА #55: цвет ссылки приезжает из профиля; дефолт сохраняет поведение
    # внешних вызовов и блока __main__
    link_color = link_color or BRAND_BLUE
    if not url or not url.strip():
        return None
    url = _unshield_escapes(url.strip())            # ПРАВКА #37
    url = re.sub(r'\\([._\-+~#?&=/])', r'\1', url)      # ПРАВКА #24: раскрытие markdown-экранирования в URL
    # ПРАВКА #38: префикс добавляется только если схемы нет вообще — иначе
    # mailto:/tel:/ftp: превращались в мёртвый https://mailto:sales@tenzosila.ru
    if not re.match(r'^[a-z][a-z0-9+.\-]*:', url, re.I):
        url = 'https://' + url
    r_id = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    run_el = OxmlElement('w:r')
    # ПРАВКА #51: порядок детей rPr задан CT_RPr —
    # rFonts, b, i, color, sz, szCs, u. Раньше u стоял перед sz и szCs.
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rPr.append(rFonts)
    if bold:
        rPr.append(OxmlElement('w:b'))
    if italic:
        rPr.append(OxmlElement('w:i'))
    color_el = OxmlElement('w:color')
    color_el.set(qn('w:val'), link_color)
    rPr.append(color_el)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(font_size * 2))
    rPr.append(sz)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), str(font_size * 2))
    rPr.append(szCs)
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)
    run_el.append(rPr)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = _unshield_escapes(text)                # ПРАВКА #37
    run_el.append(t)
    hyperlink.append(run_el)
    paragraph._p.append(hyperlink)
    return hyperlink


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
    # ПРАВКА #41: эмодзи — метка плейсхолдера только в начале абзаца. Проверка
    # «📷 in text» по всему абзацу красила оранжевой полосой любой обычный
    # абзац, где эмодзи стоит в середине предложения.
    return text.lstrip().startswith('📷') or '[Место для фото' in text

# ПРАВКА #23: блок-картинка ![alt](src)
_IMG_BLOCK_RE = re.compile(r'^!\[([^\]]*)\]\(([^)]+)\)$')

# ПРАВКА #33: сплит строки таблицы по неэкранированному |, unescape \| в ячейках
def split_table_row(line):
    parts = re.split(r'(?<!\\)\|', line.strip())
    if parts and parts[0] == '':
        parts = parts[1:]
    if parts and parts[-1] == '':
        parts = parts[:-1]
    return [p.strip().replace('\\|', '|') for p in parts]


def is_requisites_block(text):
    return bool(re.match(r'^\*\*(Кому|От кого|Кому:|От кого:)', text))

# ПРАВКА #57: дата и исходящий номер письма. Регистрозависимо — как у
# is_requisites_block, иначе «дата» в начале обычного абзаца станет шапкой.
_KOMU_LABEL_RE = re.compile(r'^\*\*(?:Кому|От кого)\s*:?\*\*:?\s*')

def is_letter_meta_block(text):
    return bool(re.match(r'^\*\*(?:Дата|Исх\.?)\s*:?\*\*', text))

def is_signature_block(text):
    # ПРАВКА #42: ^\** вместо ^\*? — одна звёздочка не покрывала «**С уважением,**»
    return bool(re.match(r'^\**С уважением', text))

def is_callout_block(text):
    """Блок !! текст !! — callout-врезка."""
    return text.startswith('!!') and text.endswith('!!')


# =============================================================================
# INLINE MARKDOWN ПАРСЕР
# =============================================================================

# ПРАВКА #37: экранированные \* \[ \] прячутся в private use area до парсинга —
# иначе звёздочка снова попадёт под правила курсива, а скобки соберутся в ложную
# ссылку. Символ возвращается на месте записи в run: X → chr(0xE000 + ord(X)).
# Диапазон E020–E07E покрывает печатный ASCII и в реальных документах не встречается.
_SHIELD_BASE = 0xE000
_SHIELD_RE   = re.compile(r'\\([*\[\]])')
_UNSHIELD_RE = re.compile('[%s-%s]' % (chr(_SHIELD_BASE + 0x20),
                                       chr(_SHIELD_BASE + 0x7E)))


def _shield_escapes(text):
    return _SHIELD_RE.sub(lambda m: chr(_SHIELD_BASE + ord(m.group(1))), text)


def _unshield_escapes(text):
    return _UNSHIELD_RE.sub(lambda m: chr(ord(m.group()) - _SHIELD_BASE), text)


def _parse_bold_italic(paragraph, text, font_name, font_size,
                       font_color, is_italic_base):
    # ПРАВКА #22: поддержка ***bold-italic*** в inline-парсере
    # ПРАВКА #36: правила флангов CommonMark — у открывающей звёздочки не может
    # быть пробела после себя, у закрывающей — пробела перед собой; звёздочка
    # между цифрами курсив не открывает. Иначе «2 * 3 * 4» и «2*3*4» съедались
    # как эмфаза и цифры габаритов слипались в «234». Квантификатор +? (не *?)
    # у содержимого — иначе «****» матчится с пустым содержимым между **…** и
    # молча пропадает из документа вместо буквального вывода.
    pattern = re.compile(
        r'(\*\*\*(?!\s)[^*\n]+?(?<!\s)\*\*\*'
        # ПРАВКА #47: ограничение на цифры из #36 было только у одиночной
        # звёздочки, поэтому «2**3**4» давало жирную тройку.
        r'|(?:(?<!\d)\*\*|\*\*(?!\d))(?!\s)[^*\n]+?(?<!\s)\*\*'
        r'|(?:(?<!\d)\*|\*(?!\d))(?!\*)(?!\s)[^*\n]+?(?<![\s*])\*(?!\*))')
    parts = pattern.split(text)
    for part in parts:
        if not part:
            continue
        run = paragraph.add_run()
        is_bold   = False
        is_italic = is_italic_base
        clean_text = part
        if part.startswith('***') and part.endswith('***') and len(part) >= 7:
            clean_text = part[3:-3]
            is_bold = True
            is_italic = True
        elif part.startswith('**') and part.endswith('**') and len(part) >= 5:
            clean_text = part[2:-2]
            is_bold = True
        elif (part.startswith('*') and part.endswith('*')
              and len(part) >= 3 and not part.startswith('**')):
            clean_text = part[1:-1]
            is_italic = True
        set_run_font(run, font_name, font_size, font_color,
                     bold=is_bold, italic=is_italic)
        run.text = _unshield_escapes(clean_text)   # ПРАВКА #37


# ПРАВКА #39: одна формулировка заглушки на блочный и инлайн путь
def _missing_image_text(alt, src):
    return f'{alt} (изображение не найдено: {src})'.strip()


def parse_inline_markdown(paragraph, text, font_name='PT Sans', font_size=12,
                          font_color=TEXT_DARK, is_italic_base=False,
                          images=None, content_width_cm=None, style=None):
    """Обрабатывает ***жирный-курсив***, **жирный**, *курсив*, [ссылки](url)
    и инлайн-картинки ![alt](src) (ПРАВКА #32).

    ПРАВКА #55: style нужен здесь ровно для одного — цвета ссылки, который
    иначе не доехал бы до add_hyperlink_run. Шрифт, кегль и цвет текста
    функция и так принимает аргументами, профиль для них не нужен."""
    link_color = (style or STYLE_PZ)['link_color']
    text = _shield_escapes(text)                    # ПРАВКА #37: \* \[ \] → PUA, до разбиения по ссылкам
    text = re.sub(r'\\([.\-+_)(:!=#|>`])', r'\1', text)  # ПРАВКА #24 + #37: раскрытие markdown-экранирования \X → X
    # ПРАВКА #21: сначала разбиваем по ссылкам [text](url)
    # ПРАВКА #32: ![alt](src) распознаётся ДО ссылок — раньше «!» оставался
    # литералом, а src превращался в мусорную гиперссылку https://image_1.png
    # ПРАВКА #38: <схема:адрес> и голый http(s)-URL — альтернативы дописаны
    # ПОСЛЕ markdown-ссылки, поэтому внутри [text](url) и ![alt](src) не
    # срабатывают: re.split на каждой позиции берёт первую подошедшую.
    # «(» в lookbehind — адрес сразу после скобки это markdown-адресат, а не
    # проза: у экранированных \[…\] скобки уже спрятаны в PUA и первая
    # альтернатива их не ловит, иначе (url) стал бы ложной автоссылкой.
    link_re = re.compile(
        r'(!?\[[^\]]*?\]\([^)]*?\)'
        r'|<[a-z][a-z0-9+.\-]*:[^<>\s]+>'
        r'|(?<![\w(])https?://[^\s<>()]*[^\s<>()\.,;:!?])', re.I)
    img_detail = re.compile(r'^!\[([^\]]*?)\]\(([^)]*?)\)$')
    link_detail = re.compile(r'^\[([^\]]+?)\]\(([^)]*?)\)$')
    auto_detail = re.compile(r'^<([a-z][a-z0-9+.\-]*:[^<>\s]+)>$'
                             r'|^(https?://[^\s<>]+)$', re.I)
    for segment in link_re.split(text):
        if not segment:
            continue
        mi = img_detail.match(segment)
        if mi:
            alt = _unshield_escapes(mi.group(1).strip())    # ПРАВКА #37
            src = _unshield_escapes(mi.group(2).strip())   # ПРАВКА #37
            if images and src in images:
                run = paragraph.add_run()
                run.add_picture(io.BytesIO(images[src]),
                                width=_image_width(images[src], content_width_cm))
            else:
                # ПРАВКА #39: та же заглушка, что и у блочной картинки —
                # раньше пустой alt давал молчаливое исчезновение
                _parse_bold_italic(paragraph, _missing_image_text(alt, src),
                                   font_name, font_size, font_color,
                                   is_italic_base)
            continue
        m = link_detail.match(segment)
        if m:
            link_text = m.group(1)
            link_url = _unshield_escapes(m.group(2).strip())   # ПРАВКА #37
            if link_url:
                add_hyperlink_run(paragraph, link_url, link_text,
                                  font_name, font_size, link_color=link_color)
            else:
                _parse_bold_italic(paragraph, link_text, font_name,
                                   font_size, font_color, is_italic_base)
            continue
        ma = auto_detail.match(segment)
        if ma:
            # ПРАВКА #38: текст автоссылки — адрес без угловых скобок
            auto_url = _unshield_escapes(ma.group(1) or ma.group(2))   # ПРАВКА #37
            add_hyperlink_run(paragraph, auto_url, auto_url,
                              font_name, font_size, link_color=link_color)
        else:
            _parse_bold_italic(paragraph, segment, font_name,
                               font_size, font_color, is_italic_base)


# ПРАВКА #32: расчёт ширины картинки вынесен из _add_inline_image,
# используется и блочной, и инлайн-вставкой
def _image_width(img_bytes, content_width_cm):
    if content_width_cm is None:
        content_width_cm = CONTENT_WIDTH_CM
    try:
        from PIL import Image
        img = Image.open(io.BytesIO(img_bytes))
        w_px, _ = img.size
        img.close()
        w_cm = w_px / 96 * 2.54
        return Cm(min(w_cm, content_width_cm))
    except Exception:
        return Cm(content_width_cm)


def _add_inline_image(doc, img_bytes, content_width_cm):
    import tempfile
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
    tmp.write(img_bytes)
    tmp.close()
    width = _image_width(img_bytes, content_width_cm)   # ПРАВКА #32
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(tmp.name, width=width)
    os.unlink(tmp.name)


# =============================================================================
# СПЕЦИАЛЬНЫЕ БЛОКИ-КОНСТРУКТОРЫ
# =============================================================================

def add_intro_paragraph(doc, block, content_width_cm, style=None):
    """
    ПРАВКА #4: Вводный абзац после H1 — таблица-обёртка с цветной левой полосой.
    Выглядит как акцентный callout для главной мысли документа.
    ПРАВКА #56: в письме врезки нет — тот же блок это тема письма, курсивом
    по центру под словом «ПИСЬМО».
    """
    style = style or STYLE_PZ

    if not style['intro_band']:
        p = doc.add_paragraph()
        p.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(16)
        set_keep_with_next(p)
        for i, line in enumerate(block.split('\n')):
            line = line.strip()
            if not line: continue
            if i > 0: p.add_run().add_break()
            parse_inline_markdown(p, line, style['body_font'],
                                  style['body_size'], style['text_color'],
                                  is_italic_base=True, style=style)
        return

    table = doc.add_table(rows=1, cols=1)
    table.autofit = False   # ПРАВКА #46: allow_autofit в python-docx нет
    set_table_width_dxa(table, content_width_cm)
    set_table_no_spacing(table)

    cell = table.rows[0].cells[0]

    # Только левая граница — фирменный синий, 2pt
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    # ПРАВКА #51: порядок сторон задан схемой — левая граница не может
    # дописываться последней, её место между top и bottom
    for side in ['top', 'left', 'bottom', 'right']:
        bdr = OxmlElement(f'w:{side}')
        if side == 'left':
            bdr.set(qn('w:val'), 'single')
            bdr.set(qn('w:sz'), '18')   # ~2.25pt
            bdr.set(qn('w:space'), '4')
            bdr.set(qn('w:color'), style['accent_color'])
        else:
            bdr.set(qn('w:val'), 'none')
            bdr.set(qn('w:sz'), '0')
            bdr.set(qn('w:color'), 'auto')
        tcBorders.append(bdr)
    insert_in_order(tcPr, tcBorders)

    # Внутренние отступы ячейки
    tcMar = OxmlElement('w:tcMar')
    # ПРАВКА #51: порядок сторон задан схемой — top, left, bottom, right
    for side, w in [('top','80'),('left','220'),('bottom','80'),('right','0')]:
        node = OxmlElement(f'w:{side}')
        node.set(qn('w:w'), w)
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    insert_in_order(tcPr, tcMar)

    p = cell.paragraphs[0]
    p.paragraph_format.alignment  = WD_ALIGN_PARAGRAPH.LEFT  # ПРАВКА #27
    p.paragraph_format.space_after = Pt(0)
    # ПРАВКА #1: межстрочный 1.3
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    p.paragraph_format.line_spacing      = style['body_line']

    lines = block.split('\n')
    for i, line in enumerate(lines):
        line = line.strip()
        if not line: continue
        if i > 0: p.add_run().add_break()
        parse_inline_markdown(p, line, style['body_font'], style['body_size'],
                              style['text_color'], style=style)


def add_callout_box(doc, text, content_width_cm, style=None):
    """
    ПРАВКА #6: Callout-врезка !! текст !! — таблица с заливкой и бордером.
    Используется для формул, ключевых выводов, важных цифр.
    """
    style = style or STYLE_PZ
    clean = text.strip('!').strip()
    table = doc.add_table(rows=1, cols=1)
    table.autofit = False   # ПРАВКА #46: allow_autofit в python-docx нет
    set_table_width_dxa(table, content_width_cm)

    cell = table.rows[0].cells[0]
    set_cell_shading(cell, style['callout_fill'])

    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top','left','bottom','right']:
        bdr = OxmlElement(f'w:{side}')
        bdr.set(qn('w:val'), 'single')
        bdr.set(qn('w:sz'), '4')     # 0.5pt
        bdr.set(qn('w:color'), style['callout_border'])
        tcBorders.append(bdr)
    # Акцент — левая граница чуть толще
    insert_in_order(tcPr, tcBorders)

    tcMar = OxmlElement('w:tcMar')
    # ПРАВКА #51: порядок сторон задан схемой — top, left, bottom, right
    for side, w in [('top','140'),('left','220'),('bottom','140'),('right','220')]:
        node = OxmlElement(f'w:{side}')
        node.set(qn('w:w'), w)
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    insert_in_order(tcPr, tcMar)

    p = cell.paragraphs[0]
    p.paragraph_format.alignment  = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(0)
    parse_inline_markdown(p, clean, style['body_font'], style['body_size'],
                          style['callout_text'], style=style)
    for r in p.runs:
        r.bold = True


def add_letter_header(doc, meta, komu, content_width_cm, style):
    """
    ПРАВКА #57: шапка письма — безрамочная таблица 1x2. Слева дата и исходящий
    номер с тонкой линией под ними, справа адресат; оба блока на одном уровне
    по вертикали. Линия рисуется границей абзаца, а не ячейки, поэтому идёт
    ровно по ширине левой колонки, а не через всю страницу.
    Любая из половин может отсутствовать — ячейка останется пустой.
    """
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False   # иначе Word разложит колонки по содержимому
    set_table_width_dxa(table, content_width_cm)
    set_table_no_spacing(table)

    widths = [style['header_left_cm'], content_width_cm - style['header_left_cm']]
    for cell, width_cm in zip(table.rows[0].cells, widths):
        set_cell_no_borders(cell)
        tcPr = cell._tc.get_or_add_tcPr()
        existing_w = tcPr.find(qn('w:tcW'))
        if existing_w is not None:
            tcPr.remove(existing_w)
        tcW = OxmlElement('w:tcW')
        tcW.set(qn('w:w'), str(int(width_cm * 567)))
        tcW.set(qn('w:type'), 'dxa')
        insert_in_order(tcPr, tcW)
        # нулевые поля ячейки: иначе адресат встанет на 2 мм внутрь от правого
        # поля страницы и разойдётся с краем текста
        tcMar = OxmlElement('w:tcMar')
        # ПРАВКА #51: порядок сторон задан схемой — top, left, bottom, right
        for side in ['top', 'left', 'bottom', 'right']:
            node = OxmlElement(f'w:{side}')
            node.set(qn('w:w'), '0')
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
        insert_in_order(tcPr, tcMar)

    for cell, text, align, rule in [
            (table.rows[0].cells[0], meta, WD_ALIGN_PARAGRAPH.LEFT, True),
            (table.rows[0].cells[1], komu, WD_ALIGN_PARAGRAPH.RIGHT, False)]:
        if not text:
            continue
        p = cell.paragraphs[0]
        p.paragraph_format.alignment   = align
        p.paragraph_format.space_after = Pt(0)
        if rule:
            add_paragraph_border(p, 'bottom', style['rule_color'], 4, space=4)
        for i, line in enumerate(text.split('\n')):
            line = line.strip()
            if not line: continue
            if i > 0: p.add_run().add_break()
            parse_inline_markdown(p, line, style['body_font'],
                                  style['body_size'], style['text_color'],
                                  style=style)


def _pop_block(blocks, match):
    """ПРАВКА #57: вынимает первый подходящий блок из списка.

    Шапка собирается из двух markdown-блоков, поэтому они выбираются до
    основного цикла — так в цикле не появляется состояния «мету видели, ждём
    Кому», которое пришлось бы тащить через все ветки."""
    for i, block in enumerate(blocks):
        if match(block.strip()):
            return blocks.pop(i).strip()
    return None


# ПРАВКА #30: компактный пустой параграф-спейсер после таблиц-блоков,
# чтобы соседние w:tbl не склеивались Word'ом в одну таблицу
def add_compact_spacer(doc):
    sp = doc.add_paragraph()
    pPr = sp._p.get_or_add_pPr()
    s = OxmlElement('w:spacing')
    s.set(qn('w:before'), '0')
    s.set(qn('w:after'), '120')
    s.set(qn('w:line'), '120')
    s.set(qn('w:lineRule'), 'exact')
    insert_in_order(pPr, s)


def add_table_cell_content(p, text, font_size=10, style=None):
    """
    ПРАВКА #8: Добавляет ✓/✗ перед значениями «Да»/«Нет» в ячейках таблицы.
    Применяется для любой сравнительной таблицы автоматически.
    ПРАВКА #17: поведение управляется флагом ENABLE_TABLE_SYMBOLS.
    ПРАВКА #34: иконка ставится ПЕРЕД текстом, сам текст ячейки не вырезается —
    «Да — 36 месяцев» → «✓ Да — 36 месяцев», «Нет данных» → «✗ Нет данных».
    """
    style = style or STYLE_PZ
    stripped = text.strip()

    if not ENABLE_TABLE_SYMBOLS:
        parse_inline_markdown(p, stripped, style['body_font'], font_size,
                              style['text_color'], style=style)
        return

    # Проверяем начало ячейки на Да/Нет/Отсутствует
    m = re.match(r'^(Да|Нет|Отсутствует)\b', stripped, re.IGNORECASE)
    if m:
        is_yes = m.group(1).lower() == 'да'
        icon_run = p.add_run('✓ ' if is_yes else '✗ ')
        set_run_font(icon_run, style['body_font'], font_size,
                     style['icon_yes'] if is_yes else style['icon_no'], bold=True)
    parse_inline_markdown(p, stripped, style['body_font'], font_size,
                          style['text_color'], style=style)


# =============================================================================
# ПРАВКА #18: АВТОМАТИЧЕСКИЕ ПЕРЕНОСЫ СЛОВ
# =============================================================================

def enable_auto_hyphenation(doc):
    """
    ПРАВКА #18: включает автоматические переносы слов на уровне документа.
    Без этого justify создаёт большие пробелы между словами в коротких строках.
    """
    settings = doc.settings.element
    existing = settings.find(qn('w:autoHyphenation'))
    if existing is not None:
        settings.remove(existing)
    insert_in_order(settings, OxmlElement('w:autoHyphenation'))
    insert_in_order(settings, OxmlElement('w:doNotHyphenateCaps'))


# =============================================================================
# ПРАВКА #13: НУМЕРАЦИЯ СПИСКОВ
# =============================================================================

def ensure_list_numbering(doc, style=None):
    """
    ПРАВКА #13: гарантирует наличие в numbering.xml кастомных определений
    для bullet (•) и numbered (1. 2. 3.) списков.
    Возвращает (bullet_num_id, numbered_num_id).
    ПРАВКА #55: глиф маркера и его шрифт берутся из профиля.
    """
    style = style or STYLE_PZ

    BULLET_ABSTRACT_ID = 100
    NUMBERED_ABSTRACT_ID = 101
    BULLET_NUM_ID = 100
    NUMBERED_NUM_ID = 101

    # Получаем или создаём numbering part
    try:
        numbering_part = doc.part.numbering_part
    except (KeyError, AttributeError):
        # Нет numbering part — создаём заглушку и регистрируем
        from docx.opc.part import Part
        from docx.opc.packuri import PackURI
        from lxml import etree

        numbering_uri = PackURI('/word/numbering.xml')
        nsmap = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        }
        numbering_xml = etree.Element(qn('w:numbering'), nsmap=nsmap)
        content_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'
        part = Part(
            numbering_uri,
            content_type,
            etree.tostring(numbering_xml, xml_declaration=True, encoding='UTF-8', standalone=True),
            doc.part.package
        )
        doc.part.relate_to(part, RT.NUMBERING)
        numbering_part = doc.part.numbering_part

    numbering_elem = numbering_part.element

    # ПРАВКА #58: если определение с нашим id уже пришло из numbering.xml
    # шаблона, снимаем его и создаём заново. Раньше функция в этом случае
    # молча возвращалась, и письмо получило бы маркер шаблона вместо длинного
    # тире из профиля — поймать такое можно было бы только глазами в Word.
    # У template.docx id идут 0-28, так что сегодня это холостой проход.
    for existing_el in numbering_elem.findall(qn('w:abstractNum')):
        if int(existing_el.get(qn('w:abstractNumId'))) in (BULLET_ABSTRACT_ID,
                                                           NUMBERED_ABSTRACT_ID):
            numbering_elem.remove(existing_el)

    existing_abstract = set()

    # --- Bullet abstractNum ---
    if BULLET_ABSTRACT_ID not in existing_abstract:
        abstract_bullet = OxmlElement('w:abstractNum')
        abstract_bullet.set(qn('w:abstractNumId'), str(BULLET_ABSTRACT_ID))
        lvl = OxmlElement('w:lvl')
        lvl.set(qn('w:ilvl'), '0')
        start = OxmlElement('w:start')
        start.set(qn('w:val'), '1')
        lvl.append(start)
        numFmt = OxmlElement('w:numFmt')
        numFmt.set(qn('w:val'), 'bullet')
        lvl.append(numFmt)
        lvlText = OxmlElement('w:lvlText')
        lvlText.set(qn('w:val'), style['bullet_char'])
        lvl.append(lvlText)
        lvlJc = OxmlElement('w:lvlJc')
        lvlJc.set(qn('w:val'), 'left')
        lvl.append(lvlJc)
        pPr = OxmlElement('w:pPr')
        ind = OxmlElement('w:ind')
        ind.set(qn('w:left'), '720')
        ind.set(qn('w:hanging'), '360')
        pPr.append(ind)
        lvl.append(pPr)
        rPr = OxmlElement('w:rPr')
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), style['bullet_font'])
        rFonts.set(qn('w:hAnsi'), style['bullet_font'])
        rFonts.set(qn('w:hint'), 'default')
        rPr.append(rFonts)
        lvl.append(rPr)
        abstract_bullet.append(lvl)
        insert_in_order(numbering_elem, abstract_bullet)

    # --- Numbered abstractNum ---
    if NUMBERED_ABSTRACT_ID not in existing_abstract:
        abstract_num = OxmlElement('w:abstractNum')
        abstract_num.set(qn('w:abstractNumId'), str(NUMBERED_ABSTRACT_ID))
        lvl = OxmlElement('w:lvl')
        lvl.set(qn('w:ilvl'), '0')
        start = OxmlElement('w:start')
        start.set(qn('w:val'), '1')
        lvl.append(start)
        numFmt = OxmlElement('w:numFmt')
        numFmt.set(qn('w:val'), 'decimal')
        lvl.append(numFmt)
        lvlText = OxmlElement('w:lvlText')
        lvlText.set(qn('w:val'), '%1.')
        lvl.append(lvlText)
        lvlJc = OxmlElement('w:lvlJc')
        lvlJc.set(qn('w:val'), 'left')
        lvl.append(lvlJc)
        pPr = OxmlElement('w:pPr')
        ind = OxmlElement('w:ind')
        ind.set(qn('w:left'), '720')
        ind.set(qn('w:hanging'), '360')
        pPr.append(ind)
        lvl.append(pPr)
        abstract_num.append(lvl)
        insert_in_order(numbering_elem, abstract_num)

    # --- num элементы (ссылки на abstractNum) ---
    existing_num = {
        int(n.get(qn('w:numId')))
        for n in numbering_elem.findall(qn('w:num'))
    }
    if BULLET_NUM_ID not in existing_num:
        num_bullet = OxmlElement('w:num')
        num_bullet.set(qn('w:numId'), str(BULLET_NUM_ID))
        abstract_ref = OxmlElement('w:abstractNumId')
        abstract_ref.set(qn('w:val'), str(BULLET_ABSTRACT_ID))
        num_bullet.append(abstract_ref)
        insert_in_order(numbering_elem, num_bullet)

    if NUMBERED_NUM_ID not in existing_num:
        num_numbered = OxmlElement('w:num')
        num_numbered.set(qn('w:numId'), str(NUMBERED_NUM_ID))
        abstract_ref = OxmlElement('w:abstractNumId')
        abstract_ref.set(qn('w:val'), str(NUMBERED_ABSTRACT_ID))
        num_numbered.append(abstract_ref)
        insert_in_order(numbering_elem, num_numbered)

    return BULLET_NUM_ID, NUMBERED_NUM_ID


def new_numbered_num_id(doc):
    """
    ПРАВКА #31: свежий w:num со startOverride=1 для каждого нового списка.
    Без этого все нумерованные списки документа делили один счётчик и
    второй список продолжался с 4 вместо 1.
    """
    numbering_elem = doc.part.numbering_part.element
    existing = [int(n.get(qn('w:numId')))
                for n in numbering_elem.findall(qn('w:num'))]
    num_id = max(existing) + 1
    num = OxmlElement('w:num')
    num.set(qn('w:numId'), str(num_id))
    abstract_ref = OxmlElement('w:abstractNumId')
    abstract_ref.set(qn('w:val'), '101')   # NUMBERED_ABSTRACT_ID из ensure_list_numbering
    num.append(abstract_ref)
    override = OxmlElement('w:lvlOverride')
    override.set(qn('w:ilvl'), '0')
    start_override = OxmlElement('w:startOverride')
    start_override.set(qn('w:val'), '1')
    override.append(start_override)
    num.append(override)
    insert_in_order(numbering_elem, num)
    return num_id


def set_paragraph_numbering(paragraph, num_id, ilvl=0):
    """ПРАВКА #13: привязывает абзац к numbering definition через OXML."""
    pPr = paragraph._p.get_or_add_pPr()
    numPr = OxmlElement('w:numPr')
    ilvl_el = OxmlElement('w:ilvl')
    ilvl_el.set(qn('w:val'), str(ilvl))
    numId_el = OxmlElement('w:numId')
    numId_el.set(qn('w:val'), str(num_id))
    numPr.append(ilvl_el)
    numPr.append(numId_el)
    insert_in_order(pPr, numPr)


# =============================================================================
# ОСНОВНАЯ ЛОГИКА КОНВЕРТАЦИИ
# =============================================================================

def convert_md_to_docx(md_text, output_filename, template_path=None, images=None,
                       doc_style='pz'):

    # ПРАВКА #55: профиль оформления выбирается типом документа. Неизвестный
    # ключ роняет конвертацию громко — молча отрисованное не тем стилем
    # клиентское письмо хуже ошибки в UI.
    style = STYLES[doc_style]

    # ПРАВКА #28: нормализация переводов строк — CRLF/CR ломали split('\n\n') и regex #26
    md_text = md_text.replace('\r\n', '\n').replace('\r', '\n')

    # --- Открываем шаблон или создаём чистый документ ---
    if template_path and os.path.exists(template_path):
        doc = Document(template_path)
        clear_body(doc)
        content_width_cm = CONTENT_WIDTH_CM
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
    sn.font.name      = style['body_font']
    sn.font.size      = Pt(style['body_size'])
    sn.font.color.rgb = RGBColor.from_string(style['text_color'])
    sn.paragraph_format.alignment         = WD_ALIGN_PARAGRAPH.LEFT  # ПРАВКА #27: justify → left для читаемости с латинской терминологией
    sn.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    sn.paragraph_format.line_spacing      = style['body_line']   # ПРАВКА #1: было 1.2
    sn.paragraph_format.space_after       = Pt(style['body_after'])  # ПРАВКА #2: было Pt(6)

    # ПРАВКА #18: включаем автоматические переносы слов
    enable_auto_hyphenation(doc)

    # ПРАВКА #13: регистрируем numbering для bullet/numbered списков
    bullet_num_id, numbered_num_id = ensure_list_numbering(doc, style)

    # ПРАВКА #23: позиционирование картинок по тексту
    _images_dict = {}
    if images:
        for fname, img_bytes in images:
            _images_dict[fname] = img_bytes

    # ПРАВКА #26: bold prefix + hard break → отдельные параграфы
    # ПРАВКА #29: не резать блоки реквизитов/стадий — их префиксы ловят
    # is_requisites_block (case-sensitive) и is_stage_paragraph (IGNORECASE)
    # ПРАВКА #57: «Дата»/«Исх» там же — иначе шапка письма разваливается на два
    # блока и линия остаётся под одной строкой
    md_text = re.sub(
        r'^(\*\*(?!(?:Кому|От кого|Дата|Исх|(?i:Стадия|Фаза|Шаг|Этап|ВАЖНО)))[^*\n]{1,100}?:\*\*)  +\n(?!\n)',
        r'\1\n\n',
        md_text,
        flags=re.MULTILINE,
    )

    # ПРАВКА #45: строка, не начинающаяся с «|», завершает таблицу. Без пустой
    # строки перед ним абзац под таблицей всасывался лишней строкой таблицы.
    md_text = re.sub(r'^(\|.*)\n(?=[^|\n])', r'\1\n\n', md_text,
                     flags=re.MULTILINE)

    # --- Парсинг блоков Markdown ---
    blocks = md_text.split('\n\n')

    # ПРАВКА #57: шапка письма собирается из двух блоков, поэтому они выбираются
    # здесь, а не в цикле. Порядок блоков в markdown значения не имеет — шапка
    # всегда идёт первой, как и положено бланку.
    if style['header_table']:
        meta_block = _pop_block(blocks, is_letter_meta_block)
        komu_block = _pop_block(blocks, is_requisites_block)
        if komu_block:
            komu_block = _KOMU_LABEL_RE.sub('', komu_block)
        if meta_block or komu_block:
            add_letter_header(doc, meta_block, komu_block, content_width_cm, style)
            add_compact_spacer(doc)

    after_heading         = False
    # ПРАВКА #12: единый флаг — intro-блок только сразу после H1
    pending_intro_after_h1 = False
    last_list_paragraph   = None
    last_regular_paragraph = None  # ПРАВКА #10: для keep_with_next перед подписью
    current_numbered_num_id = None  # ПРАВКА #31: numId текущего нумерованного списка

    for block in blocks:
        block = block.strip()
        if not block:
            continue

        is_list_item = (block.startswith('- ')
                        or block.startswith('* ')
                        # ПРАВКА #31: пробел после точки обязателен, маркер — max 2
                        # цифры, иначе «2025. Год…» съедался как пункт списка
                        or bool(re.match(r'^\d{1,2}\. ', block)))
        if not is_list_item:
            # ПРАВКА #31: любой не-списочный блок завершает текущий
            # нумерованный список — следующий начнётся с 1
            current_numbered_num_id = None
        elif re.match(r'^1\. ', block):
            # ПРАВКА #44: два списка, разделённые только пустой строкой, делили
            # один numId и второй продолжал нумерацию первого («3, 4» вместо
            # «1, 2»). Блок, начинающийся с «1. », открывает новый список.
            current_numbered_num_id = None
        if not is_list_item and last_list_paragraph:
            last_list_paragraph.paragraph_format.space_after = Pt(10)
            last_list_paragraph = None

        # ── H1 ───────────────────────────────────────────────────────────────
        if block.startswith('# '):
            p = doc.add_paragraph()
            # ПРАВКА #56: «ПИСЬМО» в бланке стоит по центру
            p.paragraph_format.alignment    = (WD_ALIGN_PARAGRAPH.CENTER
                                               if style['h1_center']
                                               else WD_ALIGN_PARAGRAPH.LEFT)
            p.paragraph_format.space_before = Pt(24)
            p.paragraph_format.space_after  = Pt(12)
            parse_inline_markdown(p, block[2:], style['head_font'],
                                  style['h1_size'], style['h1_color'], style=style)
            for r in p.runs: r.bold = True
            # ПРАВКА #14: H1 не отрывается от контента ниже
            set_keep_with_next(p)
            # ПРАВКА #56: в строгом бланке декоративной линии под заголовком нет
            if style['h1_rule']:
                dec = doc.add_paragraph()
                dec.paragraph_format.space_before = Pt(0)
                dec.paragraph_format.space_after  = Pt(10)
                add_paragraph_border(dec, 'bottom', style['rule_color'], 12)
                # ПРАВКА #14: декоративная линия тоже держится с контентом ниже
                set_keep_with_next(dec)
            after_heading = True
            # ПРАВКА #12: intro-блок ожидается только сразу после H1
            pending_intro_after_h1 = True

        # ── H2 ───────────────────────────────────────────────────────────────
        elif block.startswith('## '):
            pending_intro_after_h1 = False   # ПРАВКА #12
            p = doc.add_paragraph()
            p.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(20)
            p.paragraph_format.space_after  = Pt(8)
            # ПРАВКА #3: H2 остаётся цветным 14pt
            parse_inline_markdown(p, block[3:], style['head_font'],
                                  style['h2_size'], style['h2_color'], style=style)
            for r in p.runs: r.bold = True
            # ПРАВКА #14: H2 не отрывается от контента ниже
            set_keep_with_next(p)
            after_heading = True

        # ── H3 ───────────────────────────────────────────────────────────────
        elif block.startswith('### '):
            pending_intro_after_h1 = False   # ПРАВКА #12
            p = doc.add_paragraph()
            p.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after  = Pt(6)
            # ПРАВКА #3: H3 — чёрный bold (отличается от H2 цветом и размером)
            # ПРАВКА #15: размер 11pt → 13pt, чтобы H3 был крупнее тела (12pt)
            parse_inline_markdown(p, block[4:], style['head_font'],
                                  style['h3_size'], style['text_color'], style=style)
            for r in p.runs: r.bold = True
            # ПРАВКА #14: H3 не отрывается от контента ниже
            set_keep_with_next(p)
            after_heading = True

        # ── H4–H6 ────────────────────────────────────────────────────────────
        elif re.match(r'^#{4,6} ', block):
            pending_intro_after_h1 = False   # ПРАВКА #12
            p = doc.add_paragraph()
            p.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(10)   # меньше, чем у H3 (14pt)
            p.paragraph_format.space_after  = Pt(4)
            # ПРАВКА #40: H5 и H6 оформляются как H4 — в деловых документах
            # глубже четвёртого уровня не ходят. Без декоративных линий.
            parse_inline_markdown(p, re.sub(r'^#{4,6} ', '', block),
                                  style['head_font'], style['h4_size'],
                                  style['text_color'], style=style)
            for r in p.runs: r.bold = True
            set_keep_with_next(p)
            after_heading = True

        # ── Цитаты > ─────────────────────────────────────────────────────────
        elif block.startswith('>'):
            pending_intro_after_h1 = False   # ПРАВКА #12
            clean = '\n'.join([l.lstrip('> ') for l in block.split('\n')])
            p = doc.add_paragraph()
            # ПРАВКА #5: увеличен отступ, лёгкая заливка
            p.paragraph_format.left_indent  = Cm(1.8)
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after  = Pt(10)
            add_paragraph_border(p, 'left', style['quote_color'], 18, space=4)
            add_paragraph_shading(p, style['quote_fill'])
            parse_inline_markdown(p, clean, style['body_font'], style['body_size'],
                                  style['quote_text'], is_italic_base=True, style=style)
            last_regular_paragraph = p
            after_heading = False

        # ── Callout-врезка !! текст !! ────────────────────────────────────────
        elif is_callout_block(block):
            pending_intro_after_h1 = False   # ПРАВКА #12
            # ПРАВКА #6: новый тип блока — оформляется как акцентная таблица
            add_callout_box(doc, block, content_width_cm, style)
            # ПРАВКА #30: спейсер, иначе callout склеивается со следующим w:tbl
            add_compact_spacer(doc)
            after_heading = False

        # ── Встроенные картинки (ПРАВКА #23) ─────────────────────────────────
        # ПРАВКА #32: ветка работает и при пустом images — раньше блок падал
        # в inline-парсер и превращался в «!» + мусорную гиперссылку
        elif _IMG_BLOCK_RE.match(block):
            pending_intro_after_h1 = False
            m = _IMG_BLOCK_RE.match(block)
            alt, img_src = m.group(1).strip(), m.group(2).strip()
            if img_src in _images_dict:
                _add_inline_image(doc, _images_dict[img_src], content_width_cm)
            elif is_photo_placeholder(alt):
                # alt — фото-плейсхолдер: рендерим как ветку 📷 ниже
                p = doc.add_paragraph()
                p.paragraph_format.left_indent  = Cm(1.0)
                p.paragraph_format.space_before = Pt(10)
                p.paragraph_format.space_after  = Pt(10)
                add_paragraph_border(p, 'left', style['photo_color'], 18)
                add_paragraph_shading(p, style['photo_fill'])
                parse_inline_markdown(p, alt, style['body_font'],
                                      style['photo_size'], style['photo_text'],
                                      is_italic_base=True, style=style)
            else:
                p = doc.add_paragraph()
                parse_inline_markdown(p, _missing_image_text(alt, img_src),
                                      style['body_font'], style['body_size'],
                                      style['text_color'], style=style)
            after_heading = False

        # ── Плейсхолдеры фото ────────────────────────────────────────────────
        elif is_photo_placeholder(block):
            pending_intro_after_h1 = False   # ПРАВКА #12
            p = doc.add_paragraph()
            p.paragraph_format.left_indent  = Cm(1.0)
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after  = Pt(10)
            add_paragraph_border(p, 'left', style['photo_color'], 18)
            add_paragraph_shading(p, style['photo_fill'])
            # ПРАВКА #37, известный предел: маркер цитаты снимается до парсера,
            # поэтому экранированный \> внутри цитаты не восстанавливается
            parse_inline_markdown(p, block.replace('>', '').strip(),
                                  style['body_font'], style['photo_size'],
                                  style['photo_text'], is_italic_base=True,
                                  style=style)
            after_heading = False

        # ── Стадии / ВАЖНО ───────────────────────────────────────────────────
        elif is_stage_paragraph(block):
            pending_intro_after_h1 = False   # ПРАВКА #12
            p = doc.add_paragraph()
            p.paragraph_format.left_indent  = Cm(0.75)
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after  = Pt(6)
            add_paragraph_border(p, 'left', style['accent_color'], 12)
            add_paragraph_shading(p, style['block_fill'])
            parse_inline_markdown(p, block, style['body_font'],
                                  style['body_size'], style['text_color'],
                                  style=style)
            after_heading = False

        # ── Списки ───────────────────────────────────────────────────────────
        elif is_list_item:
            pending_intro_after_h1 = False   # ПРАВКА #12
            for line in block.split('\n'):
                line = line.strip()
                if not line: continue
                is_num = bool(re.match(r'^\d{1,2}\. ', line))  # ПРАВКА #31: пробел обязателен, max 2 цифры
                # ПРАВКА #13: привязываем numbering напрямую через OXML
                # ПРАВКА #31: каждый новый нумерованный список — свой numId с рестартом
                if is_num and current_numbered_num_id is None:
                    current_numbered_num_id = new_numbered_num_id(doc)
                p = doc.add_paragraph()
                set_paragraph_numbering(p, current_numbered_num_id if is_num else bullet_num_id)
                p.paragraph_format.left_indent       = Cm(1.5)
                p.paragraph_format.first_line_indent = Cm(-0.75)
                p.paragraph_format.space_after       = Pt(4)
                parse_inline_markdown(p, re.sub(r'^(\- |\* |\d{1,2}\. )', '', line),  # ПРАВКА #31: синхронно с is_num
                                      style['body_font'], style['body_size'],
                                      style['text_color'],
                                      images=_images_dict, content_width_cm=content_width_cm,  # ПРАВКА #32
                                      style=style)
                last_list_paragraph = p
                last_regular_paragraph = p
            after_heading = False

        # ── Таблицы ──────────────────────────────────────────────────────────
        elif block.startswith('|') and '\n|' in block:
            pending_intro_after_h1 = False   # ПРАВКА #12
            lines = [l.strip() for l in block.split('\n')
                     if l.strip() and not re.match(r'^\|[-|: ]+\|$', l.strip())]
            if not lines: continue

            headers  = split_table_row(lines[0])   # ПРАВКА #33
            n_cols   = len(headers)
            table    = doc.add_table(rows=1, cols=n_cols)
            # ПРАВКА #16: autofit для распределения ширин по содержимому
            # ПРАВКА #46: table.autofit сам пишет w:tblLayout — ручной код
            # добавлял второй такой же элемент в каждую таблицу.
            table.autofit = True
            set_table_width_dxa(table, content_width_cm)

            # ПРАВКА #7: заголовочная строка — увеличена высота и шрифт
            set_row_height(table.rows[0], 560)   # ~1cm минимальная высота

            for i, h in enumerate(headers):
                if i < len(table.rows[0].cells):
                    cell = table.rows[0].cells[i]
                    set_cell_shading(cell, style['table_head_fill'])
                    set_cell_margins_and_borders(cell, BORDER_LIGHT, 4)
                    # ПРАВКА #19: минимальная ширина первой колонки для длинных подписей
                    if n_cols >= 3 and i == 0:
                        tcPr_w = cell._tc.get_or_add_tcPr()
                        existing_w = tcPr_w.find(qn('w:tcW'))
                        if existing_w is not None:
                            tcPr_w.remove(existing_w)
                        tcW = OxmlElement('w:tcW')
                        tcW.set(qn('w:w'), str(int(6.0 * 567)))
                        tcW.set(qn('w:type'), 'dxa')
                        insert_in_order(tcPr_w, tcW)
                    p = cell.paragraphs[0]
                    p.paragraph_format.alignment   = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_after = Pt(0)
                    # ПРАВКА #7: шрифт 10pt → 11pt в заголовке таблицы
                    parse_inline_markdown(p, h, style['head_font'],
                                          style['table_head_size'],
                                          style['table_head_text'], style=style)
                    for r in p.runs: r.bold = True

            for row_idx, line in enumerate(lines[1:]):
                cols = split_table_row(line)   # ПРАВКА #33
                # ПРАВКА #33: недостающие ячейки добиваем пустыми (стили ниже
                # применяются ко всем), лишние — приклеиваем к последней
                if len(cols) < n_cols:
                    cols += [''] * (n_cols - len(cols))
                elif len(cols) > n_cols:
                    cols = cols[:n_cols - 1] + [' '.join(cols[n_cols - 1:])]
                row_cells = table.add_row().cells
                is_even   = (row_idx % 2 == 1)

                for i, c in enumerate(cols):
                    if i < len(row_cells):
                        cell = row_cells[i]

                        # ПРАВКА #9: последняя колонка 3-колоночной таблицы
                        # получает фирменный голубой фон (выделяем «наш» столбец)
                        if n_cols == 3 and i == n_cols - 1:
                            bg = style['table_alt_fill']
                        else:
                            # BRAND_WHITE — «бумага», не фирменный цвет: в обоих
                            # профилях это белый, ключа не заводим
                            bg = style['table_row_fill'] if is_even else BRAND_WHITE

                        set_cell_shading(cell, bg)
                        set_cell_margins_and_borders(cell, BORDER_LIGHT, 4)
                        # ПРАВКА #19: минимальная ширина первой колонки для длинных подписей
                        if n_cols >= 3 and i == 0:
                            tcPr_w = cell._tc.get_or_add_tcPr()
                            existing_w = tcPr_w.find(qn('w:tcW'))
                            if existing_w is not None:
                                tcPr_w.remove(existing_w)
                            tcW = OxmlElement('w:tcW')
                            tcW.set(qn('w:w'), str(int(6.0 * 567)))
                            tcW.set(qn('w:type'), 'dxa')
                            insert_in_order(tcPr_w, tcW)
                        p = cell.paragraphs[0]
                        p.paragraph_format.space_after = Pt(0)
                        # ПРАВКА #8: автоматические ✓/✗ для Да/Нет/Отсутствует
                        add_table_cell_content(p, c, style['table_cell_size'], style)

            # ПРАВКА #35: шапка повторяется на каждой странице,
            # строка не разрывается пополам между страницами
            # ПРАВКА #49: порядок выровнен для единообразия с остальными
            # элементами (CT_TrPrBase — xsd:choice maxOccurs="unbounded",
            # схема порядок внутри trPr не регламентирует)
            for row in table.rows:
                set_row_flag(row, 'w:cantSplit')
            set_row_flag(table.rows[0], 'w:tblHeader')

            # Компактный отступ после таблицы
            sp = doc.add_paragraph()
            pPr = sp._p.get_or_add_pPr()
            s = OxmlElement('w:spacing')
            s.set(qn('w:before'), '0')
            s.set(qn('w:after'), '120')
            s.set(qn('w:line'), '120')
            s.set(qn('w:lineRule'), 'exact')
            insert_in_order(pPr, s)
            after_heading = False

        # ── Разделители --- ───────────────────────────────────────────────────
        elif block.startswith('---'):
            pending_intro_after_h1 = False   # ПРАВКА #12
            continue

        # ── Реквизиты «Кому / От кого» ───────────────────────────────────────
        elif is_requisites_block(block):
            pending_intro_after_h1 = False   # ПРАВКА #12
            p = doc.add_paragraph()
            # ПРАВКА #56: в письме адресат — обычный текст с выключкой вправо,
            # без голубой подложки
            p.paragraph_format.alignment    = (WD_ALIGN_PARAGRAPH.RIGHT
                                               if style['requisites_right']
                                               else WD_ALIGN_PARAGRAPH.LEFT)
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after  = Pt(14)
            p.paragraph_format.left_indent  = Cm(0.4)
            p.paragraph_format.right_indent = Cm(0.4)
            if style['requisites_fill']:
                add_paragraph_shading(p, style['block_fill'])
            for i, line in enumerate(block.split('\n')):
                line = line.strip()
                if not line: continue
                if i > 0: p.add_run().add_break()
                parse_inline_markdown(p, line, style['body_font'],
                                      style['body_size'], style['text_color'],
                                      style=style)
            last_regular_paragraph = p
            after_heading = False

        # ── Подпись «С уважением» ─────────────────────────────────────────────
        elif is_signature_block(block):
            pending_intro_after_h1 = False   # ПРАВКА #12
            # ПРАВКА #10: последний абзац перед подписью держим вместе с ней
            if last_regular_paragraph is not None:
                set_keep_with_next(last_regular_paragraph)

            p = doc.add_paragraph()
            # выключка остаётся LEFT и в письме: правый tab stop работает
            # только на абзаце, не выключенном вправо
            p.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(28)
            p.paragraph_format.space_after  = Pt(0)
            # ПРАВКА #56: красная линия над подписью — деталь ПЗ, не бланка
            if style['signature_rule']:
                add_paragraph_border(p, 'top', style['rule_color'], 4, space=8)
            # ПРАВКА #11: весь блок подписи не разрывается по страницам
            set_keep_together(p)

            sig_lines = [l.strip() for l in block.split('\n') if l.strip()]
            # ПРАВКА #56: в письме должность и фамилия стоят одной строкой —
            # должность слева, фамилия по правой табуляции. Склейка только при
            # трёх строках и больше («С уважением,» + должность + фамилия):
            # при двух строках склеивать нечего, остаётся поведение ПЗ.
            if style['signature_tab'] and len(sig_lines) >= 3:
                p.paragraph_format.tab_stops.add_tab_stop(
                    Cm(content_width_cm), WD_TAB_ALIGNMENT.RIGHT)
                sig_lines = sig_lines[:-2] + [sig_lines[-2] + '\t' + sig_lines[-1]]

            for i, line in enumerate(sig_lines):
                if i > 0: p.add_run().add_break()
                parse_inline_markdown(p, line, style['body_font'],
                                      style['body_size'], style['text_color'],
                                      style=style)
            after_heading = False

        # ── Обычные абзацы ────────────────────────────────────────────────────
        else:
            # ПРАВКА #4 + ПРАВКА #12: intro-блок только сразу после H1
            if pending_intro_after_h1:
                add_intro_paragraph(doc, block, content_width_cm, style)
                pending_intro_after_h1 = False
                after_heading = False
                # ПРАВКА #30: спейсер разделяет соседние w:tbl. ПРАВКА #56:
                # в письме врезки-таблицы нет, разделять нечего.
                if style['intro_band']:
                    sp = doc.add_paragraph()
                    pPr = sp._p.get_or_add_pPr()
                    s = OxmlElement('w:spacing')
                    s.set(qn('w:before'), '0')
                    s.set(qn('w:after'), '120')
                    s.set(qn('w:line'), '120')
                    s.set(qn('w:lineRule'), 'exact')
                    insert_in_order(pPr, s)
                continue

            p = doc.add_paragraph()
            if not after_heading:
                p.paragraph_format.first_line_indent = Cm(style['para_indent'])
            for i, line in enumerate(block.split('\n')):
                line = line.strip()
                if not line: continue
                if i > 0: p.add_run().add_break()
                parse_inline_markdown(p, line, style['body_font'],
                                      style['body_size'], style['text_color'],
                                      images=_images_dict,
                                      content_width_cm=content_width_cm,  # ПРАВКА #32
                                      style=style)
            last_regular_paragraph = p
            # ПРАВКА #20: лид-абзац (целиком жирный) держится со следующим блоком
            stripped_block = block.strip()
            if (stripped_block.startswith('**')
                    and stripped_block.endswith('**')
                    and stripped_block.count('**') == 2):
                set_keep_with_next(p)
                p.paragraph_format.first_line_indent = Cm(0)
                p.paragraph_format.space_after = Pt(2)
            after_heading = False

    # ПРАВКА #48: библиотечный код ничего не печатает. Эмодзи в stdout роняли
    # конвертацию на cp1251-консоли Windows — костыль с -X utf8 больше не нужен.
    doc.save(output_filename)


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
