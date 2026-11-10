# build_docx.py
"""
LatentHalt Universal Academic Report Builder
Full Architecture with 9 Advanced Academic Modules (V2.1 - Bugfixed):
- Fixed: Multi-line Callout/Admonitions Buffer
- Fixed: Table Caption Pre-Detection (GOST/Pandoc syntax)
- Fixed: Remote Image URL Downloader
- Fixed: Secure Mermaid CLI Invocation (shell=False)
- Documented: Endnotes Architecture for Footnotes
- MathML/OMML Equation Engine + Native Pygments Syntax Highlighting
"""

import os
import re
import sys
import shutil
import tempfile
import subprocess
import urllib.request
import urllib.parse
import yaml

# ------------------------------------------------------------------------------
# 1. ЗАВИСИМОСТИ И ФОЛЛБЭКИ
# ------------------------------------------------------------------------------
try:
    import docx
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
    from docx.enum.style import WD_STYLE_TYPE
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import qn, nsdecls
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
except ImportError:
    print("[!] Установите python-docx: pip install python-docx")
    sys.exit(1)

# Подсветка синтаксиса через Pygments
try:
    from pygments import lex
    from pygments.lexers import get_lexer_by_name, TextLexer
    from pygments.token import Token
    HAS_PYGMENTS = True
except ImportError:
    HAS_PYGMENTS = False

# Математический движок OMML
try:
    import latex2mathml.converter
    from mathml2omml import convert as mml2omml
    from lxml import etree
    HAS_LATEX = True
except ImportError:
    HAS_LATEX = False

# ------------------------------------------------------------------------------
# 2. КОНФИГУРАЦИЯ И ГЛОБАЛЬНЫЕ СЧЕТЧИКИ
# ------------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEXT_DIR = os.path.join(BASE_DIR, "text")
TITLE_PATH = os.path.join(TEXT_DIR, "title-page.md")
CODE_PATH = os.path.join(BASE_DIR, "comprehensive_experiment.py")
OUTPUT_DOCX = os.path.join(BASE_DIR, "Universal_Academic_Report.docx")

MAIN_FONT = "Arial"
CODE_FONT = "Consolas"

FIGURE_COUNTER = 0
TABLE_COUNTER = 0
FOOTNOTES_REGISTRY = {}
FOOTNOTE_COUNTER = 0

def read_file(path):
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        return f.read()

def detect_language(text):
    ru_chars = len(re.findall(r'[а-яА-ЯёЁ]', text))
    en_chars = len(re.findall(r'[a-zA-Z]', text))
    return "ru-RU" if ru_chars > en_chars else "en-US"

# ------------------------------------------------------------------------------
# МОДУЛЬ 1: МЕТАДАННЫЕ ДОКУМЕНТА
# ------------------------------------------------------------------------------
def extract_yaml_metadata(md_text):
    """Извлекает YAML Front Matter из начала Markdown-файла."""
    if md_text.startswith("---"):
        parts = md_text.split("---", 2)
        if len(parts) >= 3:
            try:
                metadata = yaml.safe_load(parts[1])
                return metadata or {}, parts[2].strip() # Возвращаем метаданные и оставшийся текст
            except Exception as e:
                print(f"[!] Ошибка парсинга YAML: {e}")
    return {}, md_text
    
def set_document_properties(doc, title=None, author=None, subject=None, keywords=None):
    props = doc.core_properties
    if title: props.title = title
    if author: props.author = author
    if subject: props.subject = subject
    if keywords: props.keywords = keywords
    props.category = "Scientific Preprint"
    props.content_status = "Final Draft"

# ------------------------------------------------------------------------------
# МОДУЛЬ 2: СТИЛИ ДОКУМЕНТА
# ------------------------------------------------------------------------------
def setup_styles(doc):
    for i in range(1, 4):
        style_name = f'Heading {i}'
        try:
            style = doc.styles[style_name]
        except KeyError:
            style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = doc.styles['Normal']
        style.font.name = MAIN_FONT
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.font.bold = True
        if i == 1:
            style.font.size = Pt(17); style.paragraph_format.space_before = Pt(20); style.paragraph_format.space_after = Pt(10)
        elif i == 2:
            style.font.size = Pt(13.5); style.paragraph_format.space_before = Pt(16); style.paragraph_format.space_after = Pt(8)
        elif i == 3:
            style.font.size = Pt(11.5); style.paragraph_format.space_before = Pt(12); style.paragraph_format.space_after = Pt(6)

    for i in range(1, 4):
        style_name = f'TOC {i}'
        try:
            style = doc.styles[style_name]
        except KeyError:
            style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = doc.styles['Normal']
        style.font.name = MAIN_FONT
        style.font.size = Pt(10.5)
        style.font.color.rgb = RGBColor(0, 0, 0)
        if i == 1:
            style.font.bold = True; style.paragraph_format.left_indent = Inches(0.0); style.paragraph_format.space_before = Pt(6)
        elif i == 2:
            style.paragraph_format.left_indent = Inches(0.25); style.paragraph_format.space_before = Pt(2)
        else:
            style.paragraph_format.left_indent = Inches(0.5); style.paragraph_format.space_before = Pt(1)

    try:
        caption_style = doc.styles['Caption']
    except KeyError:
        caption_style = doc.styles.add_style('Caption', WD_STYLE_TYPE.PARAGRAPH)
    caption_style.font.name = MAIN_FONT
    caption_style.font.size = Pt(9.5)
    caption_style.font.italic = True
    caption_style.font.color.rgb = RGBColor(71, 85, 105)

# ------------------------------------------------------------------------------
# МОДУЛЬ 3: ПОДСВЕТКА СИНТАКСИСА КОДА (PYGMENTS)
# ------------------------------------------------------------------------------
TOKEN_COLORS = {
    Token.Keyword: RGBColor(0, 51, 179),
    Token.Keyword.Constant: RGBColor(0, 51, 179),
    Token.String: RGBColor(6, 125, 23),
    Token.Comment: RGBColor(140, 140, 140),
    Token.Number: RGBColor(23, 80, 235),
    Token.Name.Function: RGBColor(0, 98, 122),
    Token.Name.Class: RGBColor(0, 98, 122),
    Token.Operator: RGBColor(30, 41, 59),
    Token.Punctuation: RGBColor(30, 41, 59),
}

def add_styled_code_block(doc, code_text, language='python'):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)

    set_cell_background(cell, "F8FAFC")
    set_cell_border(cell, top={'sz': 4, 'color': 'E2E8F0'}, bottom={'sz': 4, 'color': 'E2E8F0'},
                          left={'sz': 4, 'color': 'E2E8F0'}, right={'sz': 4, 'color': 'E2E8F0'})

    p = cell.paragraphs[0]
    p.paragraph_format.line_spacing = 1.05
    p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)

    rPr = p._p.get_or_add_pPr().get_or_add_rPr()
    rPr.append(OxmlElement('w:noProof'))

    if not HAS_PYGMENTS:
        r = p.add_run(code_text)
        r.font.name = CODE_FONT; r.font.size = Pt(8.5); r.font.color.rgb = RGBColor(30, 41, 59)
        return

    try:
        lexer = get_lexer_by_name(language, stripall=True)
    except:
        lexer = TextLexer()

    tokens = list(lex(code_text, lexer))
    for ttype, value in tokens:
        run = p.add_run(value)
        run.font.name = CODE_FONT
        run.font.size = Pt(8.5)

        matched_color = None
        for tok_key, col in TOKEN_COLORS.items():
            if ttype in tok_key:
                matched_color = col
                break

        run.font.color.rgb = matched_color if matched_color else RGBColor(30, 41, 59)
        if ttype in Token.Comment: run.font.italic = True
        elif ttype in Token.Keyword: run.font.bold = True

# ------------------------------------------------------------------------------
# МОДУЛЬ 4: ВЫЧИСЛИТЕЛЬНАЯ МАТЕМАТИКА (OMML + СТЕКОВЫЙ ПАРСЕР)
# ------------------------------------------------------------------------------
def parse_balanced_braces(text, start_pos):
    depth = 0
    start_idx = -1
    for i in range(start_pos, len(text)):
        if text[i] == '{':
            if depth == 0: start_idx = i
            depth += 1
        elif text[i] == '}':
            depth -= 1
            if depth == 0: return start_idx, i
    return -1, -1

def replace_latex_fractions(text):
    while True:
        idx = text.find(r'\frac')
        if idx == -1: break
        num_start, num_end = parse_balanced_braces(text, idx)
        if num_start == -1: break
        den_start, den_end = parse_balanced_braces(text, num_end + 1)
        if den_start == -1 or den_start > num_end + 3: break

        num = text[num_start + 1:num_end].strip()
        den = text[den_start + 1:den_end].strip()
        repl = '½' if num == '1' and den == '2' else f"[ {num} ] / [ {den} ]"
        text = text[:idx] + repl + text[den_end + 1:]
    return text

def clean_math_to_unicode(text):
    if not text: return ""
    s = text.replace('$$', '').replace('$', '').replace('\\[', '').replace('\\]', '')
    s = replace_latex_fractions(s)
    for cmd in ['text', 'mathbf', 'mathrm', 'mathbb']:
        s = re.sub(r'\\' + cmd + r'\{([^{}]+)\}', r'\1', s)

    s = s.replace(r'\Delta', 'Δ').replace(r'\cdot', '·').replace(r'\times', '×').replace(r'\approx', '≈')
    s = s.replace(r'\to', '→').replace(r'\in', '∈').replace(r'\parallel', '||').replace(r'\mid', '|')
    s = s.replace(r'\log_2', 'log₂').replace(r'\exp', 'exp').replace(r'\sum', '∑')
    s = s.replace('ℝ^d', 'ℝᵈ').replace('h_l', 'hₗ').replace('h_{l-1}', 'hₗ₋₁').replace('h_{18}', 'h₁₈')
    s = s.replace('h_{28}', 'h₂₈').replace('W_U', 'W_U').replace('P_T', 'Pₜ')
    s = s.replace('Δ_{geom}', 'Δ_geom').replace('Δ_{info}', 'Δ_info').replace('Δ_{friction}', 'Δ_friction')
    s = s.replace('\\', '').replace('{', '').replace('}', '')
    return re.sub(r' +', ' ', s).strip()

def latex_to_omml(latex_str):
    if not HAS_LATEX: return None
    try:
        clean = latex_str.replace('$$', '').replace('\\[', '').replace('\\]', '').strip()
        clean = clean.replace(r'\parallel', '||').replace(r'\|', '|').replace(r'\mid', '|')
        clean = re.sub(r'\\mathbf\{([^}]+)\}', r'\1', clean)
        mathml = latex2mathml.converter.convert(clean)
        omml_str = mml2omml(mathml)
        if not omml_str: return None
        omml_str = re.sub(r'<\?xml[^>]*\?>', '', omml_str).strip()
        wrapped = f'<root xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{omml_str}</root>'
        root = etree.fromstring(wrapped.encode('utf-8'), parser=etree.XMLParser(recover=True))
        return root[0] if len(root) > 0 else None
    except:
        return None

# ------------------------------------------------------------------------------
# МОДУЛЬ 5: БЛОКИ ЦИТАТ И CALLOUTS (ФИКС МНОГОСТРОЧНОСТИ)
# ------------------------------------------------------------------------------
def add_admonition(doc, kind, title, body_text):
    """Создает единую цветную плашку Callout для всего блока строк"""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)

    configs = {
        'NOTE': {'color': '2563EB', 'bg': 'EFF6FF', 'icon': 'ℹ️', 'label': 'Примечание'},
        'WARNING': {'color': 'D97706', 'bg': 'FFFBEB', 'icon': '⚠️', 'label': 'Внимание'},
        'TIP': {'color': '059669', 'bg': 'ECFDF5', 'icon': '💡', 'label': 'Совет'},
        'IMPORTANT': {'color': '7C3AED', 'bg': 'F5F3FF', 'icon': '📌', 'label': 'Важно'},
        'CAUTION': {'color': 'DC2626', 'bg': 'FEF2F2', 'icon': '🛑', 'label': 'Осторожно'},
    }
    cfg = configs.get(kind.upper(), configs['NOTE'])

    set_cell_background(cell, cfg['bg'])
    set_cell_border(cell, left={'sz': 24, 'color': cfg['color']},
                          top={'sz': 4, 'color': 'E2E8F0'},
                          bottom={'sz': 4, 'color': 'E2E8F0'},
                          right={'sz': 4, 'color': 'E2E8F0'})

    p_header = cell.paragraphs[0]
    p_header.paragraph_format.space_before = Pt(4); p_header.paragraph_format.space_after = Pt(2)
    h_run = p_header.add_run(f"{cfg['icon']} {title or cfg['label']}")
    h_run.font.name = MAIN_FONT; h_run.font.size = Pt(10); h_run.font.bold = True
    h_run.font.color.rgb = RGBColor.from_string(cfg['color'])

    if body_text:
        for b_line in body_text.split('\n'):
            b_line = b_line.strip()
            if not b_line: continue
            p_body = cell.add_paragraph()
            p_body.paragraph_format.space_before = Pt(1); p_body.paragraph_format.space_after = Pt(2)
            p_body.paragraph_format.line_spacing = 1.15
            add_inline_runs(p_body, clean_math_to_unicode(b_line), base_size=9.5, lang=detect_language(b_line))

def add_blockquote(doc, body_text):
    """Оформление многострочной цитаты единым блоком"""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_border(cell, left={'sz': 16, 'color': '94A3B8'})

    lines = [l.strip() for l in body_text.split('\n') if l.strip()]
    for i, l in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.1)
        p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
        run = p.add_run(clean_math_to_unicode(l))
        run.font.name = MAIN_FONT; run.font.size = Pt(10); run.font.italic = True
        run.font.color.rgb = RGBColor(71, 85, 105)

def flush_quote_buffer(doc, quote_lines):
    """Парсит буфер цитат и решает, это Callout или обычная цитата"""
    if not quote_lines: return
    first_line = quote_lines[0].strip()
    callout_m = re.match(r'^\[!([A-Z]+)\]\s*(.*)$', first_line, re.IGNORECASE)

    if callout_m:
        kind = callout_m.group(1).upper()
        title = callout_m.group(2).strip()
        body = "\n".join(quote_lines[1:])
        add_admonition(doc, kind, title, body)
    else:
        body = "\n".join(quote_lines)
        add_blockquote(doc, body)

# ------------------------------------------------------------------------------
# МОДУЛЬ 6: РИСУНКИ (ФИКС URL) И MERMAID (БЕЗОПАСНЫЙ CLI)
# ------------------------------------------------------------------------------
def add_figure(doc, img_source, caption):
    """
    Вставляет изображение (локальный путь или URL) с центровкой и подписью
    """
    global FIGURE_COUNTER
    temp_downloaded_file = None
    local_path = img_source

    # Загрузка изображения по URL
    if img_source.startswith(('http://', 'https://')):
        try:
            req = urllib.request.Request(img_source, headers={'User-Agent': 'Mozilla/5.0'})
            parsed_url = urllib.parse.urlparse(img_source)
            ext = os.path.splitext(parsed_url.path)[1] or '.png'
            with urllib.request.urlopen(req, timeout=15) as response:
                with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                    tmp.write(response.read())
                    temp_downloaded_file = tmp.name
                    local_path = temp_downloaded_file
        except Exception as e:
            print(f"[!] Ошибка загрузки изображения по URL '{img_source}': {e}")
            return

    if not os.path.exists(local_path):
        return

    try:
        FIGURE_COUNTER += 1
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(12); p_img.paragraph_format.space_after = Pt(4)
        p_img.add_run().add_picture(local_path, width=Inches(6.0))

        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(14)
        r = p_cap.add_run(f"Рис. {FIGURE_COUNTER}. {caption}")
        r.font.name = MAIN_FONT; r.font.size = Pt(9.5); r.font.italic = True
        r.font.color.rgb = RGBColor(71, 85, 105)
    finally:
        if temp_downloaded_file and os.path.exists(temp_downloaded_file):
            try: os.remove(temp_downloaded_file)
            except: pass

def render_mermaid_to_png(mmd_code):
    """
    Безопасный рендеринг Mermaid через shell=False
    """
    mmdc_bin = shutil.which("mmdc")
    if not mmdc_bin:
        return None

    in_path = None
    try:
        with tempfile.NamedTemporaryFile('w', delete=False, suffix='.mmd', encoding='utf-8') as f_in:
            f_in.write(mmd_code)
            in_path = f_in.name
        out_path = in_path.replace('.mmd', '.png')

        cmd = [mmdc_bin, '-i', in_path, '-o', out_path, '-b', 'transparent']
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=False, timeout=30)
        if res.returncode == 0 and os.path.exists(out_path):
            return out_path
    except Exception as e:
        print(f"[!] Mermaid render exception: {e}")
    finally:
        if in_path and os.path.exists(in_path):
            try: os.remove(in_path)
            except: pass
    return None

# ------------------------------------------------------------------------------
# МОДУЛЬ 7: ТАБЛИЦЫ И ПОДПИСИ (ФИКС CAPTIONS)
# ------------------------------------------------------------------------------
def add_table_caption(doc, text):
    global TABLE_COUNTER
    TABLE_COUNTER += 1
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(14); p.paragraph_format.space_after = Pt(4)
    r = p.add_run(f"Таблица {TABLE_COUNTER}. {text}")
    r.font.name = MAIN_FONT; r.font.size = Pt(9.5); r.font.bold = True
    r.font.color.rgb = RGBColor(30, 41, 59)

def auto_fit_table_widths(table, col_data):
    max_table_width = 6.5
    cols_count = len(table.columns)
    if cols_count == 0: return

    col_lens = []
    for c_idx in range(cols_count):
        lens = [len(str(row[c_idx])) for row in col_data if c_idx < len(row)]
        col_lens.append(max(lens) if lens else 1)

    total_len = sum(col_lens) or 1
    for c_idx, col in enumerate(table.columns):
        ratio = col_lens[c_idx] / total_len
        col_width = Inches(max(0.7, max_table_width * ratio))
        for cell in col.cells:
            cell.width = col_width

# ------------------------------------------------------------------------------
# МОДУЛЬ 8: ИНЛАЙН-ПАРСИНГ И КОНЦЕВЫЕ СНОСКИ (ENDNOTES)
# ------------------------------------------------------------------------------
# АРХИТЕКТУРНОЕ ПРИМЕЧАНИЕ:
# python-docx не поддерживает построение нативных постраничных сносок (w:footnote)
# без ручной модификации пакета OPC (document.xml.rels + footnotes.xml).
# Для научных публикаций и препринтов используется признанный стандарт Endnotes:
# маркер сноски [^1] внедряется верхним индексом, а расшифровка выносится
# в итоговый раздел "Сноски и примечания".

def parse_html_inline_tags(paragraph, text, base_size=11, lang="ru-RU"):
    tokens = re.split(r'(</?(?:sub|sup|b|i|br|code)>)', text)
    current_sub = False; current_sup = False; current_bold = False; current_italic = False; current_code = False

    for tok in tokens:
        if not tok: continue
        lower = tok.lower()
        if lower == '<sub>': current_sub = True; continue
        elif lower == '</sub>': current_sub = False; continue
        elif lower == '<sup>': current_sup = True; continue
        elif lower == '</sup>': current_sup = False; continue
        elif lower == '<b>': current_bold = True; continue
        elif lower == '</b>': current_bold = False; continue
        elif lower == '<i>': current_italic = True; continue
        elif lower == '</i>': current_italic = False; continue
        elif lower == '<code>': current_code = True; continue
        elif lower == '</code>': current_code = False; continue
        elif lower == '<br>': paragraph.add_run().add_break(); continue

        run = paragraph.add_run(tok)
        run.font.name = CODE_FONT if current_code else MAIN_FONT
        run.font.size = Pt(base_size - 1 if current_code else base_size)
        run.font.subscript = current_sub
        run.font.superscript = current_sup
        run.font.bold = current_bold
        run.font.italic = current_italic

        lang_el = OxmlElement('w:lang'); lang_el.set(qn('w:val'), lang)
        run._element.get_or_add_rPr().append(lang_el)

def add_footnote_reference(paragraph, ref_key):
    global FOOTNOTE_COUNTER
    if ref_key not in FOOTNOTES_REGISTRY:
        FOOTNOTE_COUNTER += 1
        FOOTNOTES_REGISTRY[ref_key] = {"num": FOOTNOTE_COUNTER, "text": ""}

    num = FOOTNOTES_REGISTRY[ref_key]["num"]
    r = paragraph.add_run(f"[{num}]")
    r.font.name = MAIN_FONT; r.font.size = Pt(8.5); r.font.superscript = True
    r.font.color.rgb = RGBColor(37, 99, 235)

def add_inline_runs(paragraph, text, base_size=11, lang="ru-RU"):
    pattern = r'(\[\^[^\]]+\]|\[[^\]]+\]\([^)]+\)|\*\*.*?\*\*|\*.*?\*|`.*?`|https?://[^\s\)]+)'
    parts = re.split(pattern, text)

    for part in parts:
        if not part: continue

        if part.startswith("[^") and part.endswith("]"):
            add_footnote_reference(paragraph, part[2:-1])
            continue

        md_link = re.match(r'\[([^\]]+)\]\(([^)]+)\)', part)
        if md_link:
            add_hyperlink(paragraph, md_link.group(2), md_link.group(1), base_size, lang)
            continue

        if part.startswith("http://") or part.startswith("https://"):
            add_hyperlink(paragraph, part, part, base_size, lang)
            continue

        if part.startswith("**") and part.endswith("**"):
            add_run_with_lang(paragraph, part[2:-2], base_size, bold=True, lang=lang); continue
        if part.startswith("*") and part.endswith("*"):
            add_run_with_lang(paragraph, part[1:-1], base_size, italic=True, lang=lang); continue

        if part.startswith("`") and part.endswith("`"):
            r = paragraph.add_run(part[1:-1])
            r.font.name = CODE_FONT; r.font.size = Pt(base_size - 1); r.font.color.rgb = RGBColor(30, 41, 59)
            continue

        if re.search(r'</?(?:sub|sup|b|i|br|code)>', part, re.IGNORECASE):
            parse_html_inline_tags(paragraph, part, base_size, lang)
        else:
            add_run_with_lang(paragraph, part, base_size, lang=lang)

# ------------------------------------------------------------------------------
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ XML
# ------------------------------------------------------------------------------
def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_border(cell, **kwargs):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = f'w:{edge}'
            element = OxmlElement(tag)
            element.set(qn('w:val'), edge_data.get('val', 'single'))
            element.set(qn('w:sz'), str(edge_data.get('sz', 4)))
            element.set(qn('w:space'), '0')
            element.set(qn('w:color'), edge_data.get('color', 'auto'))
            tcBorders.append(element)
    tcPr.append(tcBorders)

def add_hyperlink(paragraph, url, text, base_size, lang="ru-RU"):
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    c = OxmlElement('w:color'); c.set(qn('w:val'), '111827'); rPr.append(c)
    u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rPr.append(u)
    rFonts = OxmlElement('w:rFonts'); rFonts.set(qn('w:ascii'), MAIN_FONT); rFonts.set(qn('w:hAnsi'), MAIN_FONT); rPr.append(rFonts)
    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), str(int(base_size * 2))); rPr.append(sz)
    lang_el = OxmlElement('w:lang'); lang_el.set(qn('w:val'), lang); rPr.append(lang_el)
    new_run.append(rPr)
    t = OxmlElement('w:t'); t.text = text; new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def add_run_with_lang(paragraph, text, size, bold=False, italic=False, color=None, lang="ru-RU"):
    run = paragraph.add_run(text)
    run.font.name = MAIN_FONT; run.font.size = Pt(size); run.font.bold = bold; run.font.italic = italic
    if color: run.font.color.rgb = color
    lang_el = OxmlElement('w:lang'); lang_el.set(qn('w:val'), lang)
    run._element.get_or_add_rPr().append(lang_el)
    return run

def add_page_number_footer(section):
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin')
    txt = OxmlElement('w:instrText'); txt.set(qn('xml:space'), 'preserve'); txt.text = "PAGE"
    fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'end')
    run._r.append(fld1); run._r.append(txt); run._r.append(fld2)
    run.font.name = MAIN_FONT; run.font.size = Pt(10)

# ------------------------------------------------------------------------------
# МОДУЛЬ 9: ГЕНЕРАЦИЯ ОГЛАВЛЕНИЯ
# ------------------------------------------------------------------------------
def extract_headings(md_text):
    headings = []
    skip_toc = False
    for line in md_text.split('\n'):
        s = line.strip()
        if s.lower().startswith("## содержание") or s.lower().startswith("## table of contents"):
            skip_toc = True; continue
        if skip_toc and (s.startswith("## ") or s.startswith("# ")): skip_toc = False
        if skip_toc: continue
        if s.startswith("# "): headings.append((1, clean_math_to_unicode(s[2:])))
        elif s.startswith("## "): headings.append((2, clean_math_to_unicode(s[3:])))
        elif s.startswith("### "): headings.append((3, clean_math_to_unicode(s[4:])))
    return headings

def add_toc(doc, all_docs_headings):
    p_h = doc.add_paragraph(); p_h.paragraph_format.space_before = Pt(20)
    r_h = p_h.add_run("Содержание / Table of Contents")
    r_h.font.name = MAIN_FONT; r_h.font.size = Pt(16); r_h.font.bold = True

    p_toc_start = doc.add_paragraph()
    r_b = p_toc_start.add_run()
    fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin'); r_b._r.append(fld1)
    r_i = p_toc_start.add_run()
    txt = OxmlElement('w:instrText'); txt.set(qn('xml:space'), 'preserve'); txt.text = 'TOC \\o "1-3" \\h \\z \\u'; r_i._r.append(txt)
    r_s = p_toc_start.add_run()
    fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'separate'); r_s._r.append(fld2)

    for file_name, h_list in all_docs_headings:
        p_part = doc.add_paragraph(); p_part.paragraph_format.space_before = Pt(12)
        r_p = p_part.add_run(f"■ {file_name}")
        r_p.font.name = MAIN_FONT; r_p.font.size = Pt(11); r_p.font.bold = True; r_p.font.color.rgb = RGBColor(100, 100, 100)

        for level, title in h_list:
            p_entry = doc.add_paragraph()
            p_entry.paragraph_format.tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
            p_entry.paragraph_format.left_indent = Inches(0.3 * (level - 1))
            r_e = p_entry.add_run(title)
            r_e.font.name = MAIN_FONT; r_e.font.size = Pt(10.5); r_e.font.bold = (level == 1)
            p_entry.add_run('\t•')

    p_toc_end = doc.add_paragraph()
    r_e = p_toc_end.add_run()
    fld3 = OxmlElement('w:fldChar'); fld3.set(qn('w:fldCharType'), 'end'); r_e._r.append(fld3)
    doc.add_page_break()

# ------------------------------------------------------------------------------
# МОДУЛЬ 10: ПАРСИНГ MARKDOWN В DOCX (БУФЕРИЗИРОВАННЫЙ)
# ------------------------------------------------------------------------------
def parse_md_to_docx(doc, md_text, banner_text):
    p_banner = doc.add_paragraph(); p_banner.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run_with_lang(p_banner, f"■ {banner_text} ■", 11, bold=True, color=RGBColor(150, 150, 150))
    doc.add_paragraph()

    lines = md_text.split('\n')
    in_code = False; code_lang = 'python'; code_lines = []
    in_math = False; math_lines = []
    in_table = False; table_rows = []
    in_quote = False; quote_lines = []
    pending_table_caption = None

    for line in lines:
        raw_line = line.rstrip()
        stripped = raw_line.strip()

        # 1. Завершение буфера цитат/Callout, если текущая строка не '>'
        if in_quote and not stripped.startswith(">"):
            flush_quote_buffer(doc, quote_lines)
            in_quote = False
            quote_lines = []

        if not stripped:
            continue

        # 2. Сноски-определения: [^1]: Описание
        fn_def = re.match(r'^\[\^([^\]]+)\]:\s*(.+)$', stripped)
        if fn_def:
            key, val = fn_def.group(1), fn_def.group(2)
            if key in FOOTNOTES_REGISTRY:
                FOOTNOTES_REGISTRY[key]["text"] = val
            continue

        # 3. Блоки кода (включая Mermaid)
        if stripped.startswith("```"):
            if not in_code:
                in_code = True
                code_lang = stripped[3:].strip() or 'python'
                code_lines = []
            else:
                in_code = False
                joined_code = "\n".join(code_lines)
                if code_lang.lower() == 'mermaid':
                    png = render_mermaid_to_png(joined_code)
                    if png:
                        add_figure(doc, png, "Диаграмма процесса")
                        try: os.remove(png)
                        except: pass
                    else:
                        add_styled_code_block(doc, joined_code, language='text')
                else:
                    add_styled_code_block(doc, joined_code, language=code_lang)
            continue

        if in_code:
            code_lines.append(raw_line)
            continue

        # 4. Выключные математические формулы ($$ ... $$)
        if (stripped.startswith("$$") or stripped.startswith("\\[")) and not in_math:
            if (stripped.endswith("$$") or stripped.endswith("\\]")) and len(stripped) > 2:
                omml = latex_to_omml(stripped)
                p_eq = doc.add_paragraph(); p_eq.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_eq.paragraph_format.space_before = Pt(8); p_eq.paragraph_format.space_after = Pt(8)
                if omml is not None: p_eq._p.append(omml)
                else: add_run_with_lang(p_eq, clean_math_to_unicode(stripped), 11, italic=True)
                continue
            else:
                in_math = True; math_lines = [stripped]; continue

        if in_math:
            math_lines.append(stripped)
            if stripped.endswith("$$") or stripped.endswith("\\]"):
                in_math = False
                joined_math = " ".join(math_lines)
                omml = latex_to_omml(joined_math)
                p_eq = doc.add_paragraph(); p_eq.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_eq.paragraph_format.space_before = Pt(8); p_eq.paragraph_format.space_after = Pt(8)
                if omml is not None: p_eq._p.append(omml)
                else: add_run_with_lang(p_eq, clean_math_to_unicode(joined_math), 11, italic=True)
            continue

        # 5. Перехват подписей таблиц перед таблицей (GOST / Pandoc)
        table_cap_match = re.match(r'^(?:Table|Таблица)(?:\s*\d+)?[\.:]\s*(.+)$', stripped, re.IGNORECASE)
        if not table_cap_match:
            table_cap_match = re.match(r'^:\s*(.+)$', stripped)

        if table_cap_match and not in_table:
            pending_table_caption = table_cap_match.group(1).strip()
            continue

        # 6. Цитаты и Callout-блоки (накопление в буфер)
        if stripped.startswith(">"):
            in_quote = True
            cleaned_q = re.sub(r'^>\s?', '', raw_line)
            quote_lines.append(cleaned_q)
            continue

        # 7. Заголовки Markdown
        if stripped.startswith("# "):
            doc.add_heading(level=1); add_run_with_lang(doc.paragraphs[-1], clean_math_to_unicode(stripped[2:]), 17, bold=True); continue
        elif stripped.startswith("## "):
            doc.add_heading(level=2); add_run_with_lang(doc.paragraphs[-1], clean_math_to_unicode(stripped[3:]), 13.5, bold=True); continue
        elif stripped.startswith("### "):
            doc.add_heading(level=3); add_run_with_lang(doc.paragraphs[-1], clean_math_to_unicode(stripped[4:]), 11.5, bold=True); continue

        # 8. Картинки: ![Caption](path_or_url)
        img_m = re.match(r'^!\[(.*?)\]\((.*?)\)$', stripped)
        if img_m:
            cap, path_or_url = img_m.group(1), img_m.group(2)
            resolved_path = path_or_url if path_or_url.startswith(('http://', 'https://')) or os.path.isabs(path_or_url) else os.path.join(BASE_DIR, path_or_url)
            add_figure(doc, resolved_path, cap)
            continue

        # 9. Таблицы Markdown
        if stripped.startswith("|") and stripped.endswith("|"):
            cells = [clean_math_to_unicode(c.strip()) for c in stripped.split("|")[1:-1]]
            if all(set(c).issubset({'-', ':', ' '}) for c in cells):
                continue
            table_rows.append(cells)
            in_table = True
            continue
        else:
            if in_table and table_rows:
                # Вставка подписи к таблице
                add_table_caption(doc, pending_table_caption or "Сводные данные")
                pending_table_caption = None

                cols = max(len(r) for r in table_rows)
                tbl = doc.add_table(rows=len(table_rows), cols=cols)
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                tbl.autofit = False
                auto_fit_table_widths(tbl, table_rows)

                for r_idx, row in enumerate(table_rows):
                    for c_idx, val in enumerate(row):
                        if c_idx < cols:
                            cell = tbl.cell(r_idx, c_idx)
                            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                            p = cell.paragraphs[0]
                            p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(4)
                            p.paragraph_format.line_spacing = 1.0

                            # Числовое автовыравнивание
                            clean_val = re.sub(r'\*\*|\*|`', '', val).strip() # <-- Очистка
                            is_num = bool(re.match(r'^[\d\.\,\%\sx\+\-\/><=]+$', clean_val))
                            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_num and r_idx > 0 else WD_ALIGN_PARAGRAPH.LEFT

                            add_inline_runs(p, val, base_size=9, lang=detect_language(val))

                            if r_idx == 0:
                                for r in p.runs: r.font.bold = True
                                set_cell_background(cell, "F1F5F9")
                                set_cell_border(cell, top={'sz': 12, 'color': '000000'}, bottom={'sz': 6, 'color': '000000'})
                            else:
                                set_cell_border(cell, bottom={'sz': 4, 'color': 'E2E8F0'})

                table_rows = []; in_table = False
                doc.add_paragraph()

        # Если осталась неиспользованная подпись таблицы, сбрасываем ее как абзац
        if pending_table_caption:
            p_saved = doc.add_paragraph()
            add_inline_runs(p_saved, pending_table_caption, base_size=11, lang=detect_language(pending_table_caption))
            pending_table_caption = None

        # 10. Обычный текст и списки
        lang = detect_language(stripped)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.line_spacing = 1.35
        p.paragraph_format.space_after = Pt(6)

        if re.match(r'^\d+\.\s', stripped) or stripped.startswith("- ") or stripped.startswith("* "):
            leading_spaces = len(raw_line) - len(raw_line.lstrip())
            p.paragraph_format.left_indent = Inches(0.25 + leading_spaces * 0.08)
            bullet = "– " if leading_spaces > 0 else "• "
            text_part = stripped[2:] if stripped.startswith(("- ", "* ")) else stripped
            add_inline_runs(p, bullet + clean_math_to_unicode(text_part), base_size=11, lang=lang)
        else:
            add_inline_runs(p, clean_math_to_unicode(stripped), base_size=11, lang=lang)

    # Очистка буфера цитат, если файл заканчивается на цитату
        # Очистка буфера цитат, если файл заканчивается на цитату
    if in_quote:
        flush_quote_buffer(doc, quote_lines)
        
    # Очистка буфера таблицы, если файл заканчивается на таблицу
    if in_table and table_rows:
        add_table_caption(doc, pending_table_caption or "Сводные данные")
        cols = max(len(r) for r in table_rows)
        tbl = doc.add_table(rows=len(table_rows), cols=cols)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        auto_fit_table_widths(tbl, table_rows)
        for r_idx, row in enumerate(table_rows):
            for c_idx, val in enumerate(row):
                if c_idx < cols:
                    cell = tbl.cell(r_idx, c_idx)
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    p = cell.paragraphs[0]
                    p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(4)
                    p.paragraph_format.line_spacing = 1.0
                    clean_val = re.sub(r'\*\*|\*|`', '', val).strip()
                    is_num = bool(re.match(r'^[\d\.\,\%\sx\+\-\/><=]+$', clean_val))
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_num and r_idx > 0 else WD_ALIGN_PARAGRAPH.LEFT
                    add_inline_runs(p, val, base_size=9, lang=detect_language(val))
                    if r_idx == 0:
                        for r in p.runs: r.font.bold = True
                        set_cell_background(cell, "F1F5F9")
                        set_cell_border(cell, top={'sz': 12, 'color': '000000'}, bottom={'sz': 6, 'color': '000000'})
                    else:
                        set_cell_border(cell, bottom={'sz': 4, 'color': 'E2E8F0'})

    doc.add_page_break()

# ------------------------------------------------------------------------------
# МОДУЛЬ 11: СБОРКА И ПРИЛОЖЕНИЯ
# ------------------------------------------------------------------------------
def append_footnotes_section(doc):
    if not FOOTNOTES_REGISTRY: return

    p_h = doc.add_paragraph()
    p_h.paragraph_format.space_before = Pt(16)
    r_h = p_h.add_run("Сноски и примечания (Footnotes)")
    r_h.font.name = MAIN_FONT; r_h.font.size = Pt(12); r_h.font.bold = True

    for k, v in FOOTNOTES_REGISTRY.items():
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(2)
        r_num = p.add_run(f"[{v['num']}] ")
        r_num.font.name = MAIN_FONT; r_num.font.size = Pt(9); r_num.font.bold = True
        r_txt = p.add_run(v['text'] or k)
        r_txt.font.name = MAIN_FONT; r_txt.font.size = Pt(9)

def generate_document():
    global FIGURE_COUNTER, TABLE_COUNTER, FOOTNOTES_REGISTRY, FOOTNOTE_COUNTER
    FIGURE_COUNTER = 0
    TABLE_COUNTER = 0
    FOOTNOTES_REGISTRY = {}
    FOOTNOTE_COUNTER = 0
    print("[*] Поиск Markdown файлов в папке text/...")
    if not os.path.exists(TEXT_DIR):
        print(f"[!] Папка {TEXT_DIR} не найдена.")
        sys.exit(1)

    files = [f for f in os.listdir(TEXT_DIR) if f.endswith(".md") and f != "title-page.md"]
    files.sort()

    doc = Document()
    setup_styles(doc)

    for section in doc.sections:
        section.top_margin = Inches(0.9); section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(0.9); section.right_margin = Inches(0.9)
        hp = section.header.paragraphs[0]; hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("Universal Academic Report  |  2026")
        hrun.font.name = MAIN_FONT; hrun.font.size = Pt(8.5); hrun.font.color.rgb = RGBColor(140, 140, 140)
        section.different_first_page_header_footer = True
        add_page_number_footer(section)

    # Титульный лист
    title_text = read_file(TITLE_PATH)
    doc_title = "Universal Academic Report"
    doc_author = "author"
    doc_subject = "subject"
    doc_keywords = "keywords"

    if title_text:
        print("[*] Генерация титульного листа...")
        
        # ИЗВЛЕКАЕМ YAML
        meta, body_text = extract_yaml_metadata(title_text)
        
        if meta:
            doc_title = meta.get("title", doc_title)
            doc_author = meta.get("author", doc_author)
            doc_subject = meta.get("subject", doc_subject)
            doc_keywords = meta.get("keywords", doc_keywords)

        # РЕНДЕРИМ ВИЗУАЛЬНУЮ ЧАСТЬ
        for line in body_text.split('\n'):
            s = line.strip()

    if title_text:
        print("[*] Генерация титульного листа...")
        for line in title_text.split('\n'):
            s = line.strip()
            if not s: continue
            if s.startswith("# "):
                doc_title = s[2:]
                p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                add_run_with_lang(p, doc_title, 18, bold=True)
            elif s.startswith("## "):
                p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                add_run_with_lang(p, s[3:], 13, italic=True, color=RGBColor(100,100,100))
            elif s.startswith("### "):
                doc_author = s[4:]
                p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                add_run_with_lang(p, doc_author, 12, bold=True)
            elif s.startswith("#### ") or s.startswith("["):
                p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                add_inline_runs(p, s.lstrip('# '), 10.5)
            elif s.startswith(">"):
                p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.left_indent = Inches(0.4); p.paragraph_format.right_indent = Inches(0.4)
                add_inline_runs(p, s[1:].strip(), 9)
            elif s.startswith("---"):
                doc.add_paragraph()
            else:
                p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                add_run_with_lang(p, s, 9, color=RGBColor(120,120,120))
        doc.add_page_break()

    set_document_properties(doc, title=doc_title, author=doc_author, subject=doc_subject, keywords=doc_keywords)

    # Сбор заголовков для оглавления
    print("[*] Сбор оглавления...")
    all_headings = []
    docs_content = {}
    for f_name in files:
        content = read_file(os.path.join(TEXT_DIR, f_name))
        docs_content[f_name] = content
        all_headings.append((f_name, extract_headings(content)))

    add_toc(doc, all_headings)

    # Парсинг глав
    print("[*] Вставка глав...")
    for f_name in files:
        parse_md_to_docx(doc, docs_content[f_name], f_name)

    # Реестр сносок
    append_footnotes_section(doc)

    # Приложение: Исходный код
    if os.path.exists(CODE_PATH):
        print("[*] Добавление Приложения с кодом (Pygments)...")
        p_ban = doc.add_paragraph(); p_ban.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run_with_lang(p_ban, "ПРИЛОЖЕНИЕ A • ИСХОДНЫЙ КОД СТЕНДА", 13, bold=True)
        add_styled_code_block(doc, read_file(CODE_PATH), language='python')

    doc.save(OUTPUT_DOCX)
    print(f"\n[УСПЕХ] Академический файл успешно скомпилирован: {OUTPUT_DOCX}")

if __name__ == "__main__":
    generate_document()