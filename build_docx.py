import os
import re
import sys
import zipfile
import shutil
import logging
import argparse
import tempfile
import subprocess
from dataclasses import dataclass, field

import yaml
import requests
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from lxml import etree
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from markdown_it import MarkdownIt
from mdit_py_plugins.dollarmath import dollarmath_plugin
from mdit_py_plugins.footnote import footnote_plugin
import docx.opc.constants

from pygments import lex
from pygments.lexers import get_lexer_for_filename, TextLexer
from pygments.token import Token

logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s', datefmt='%H:%M:%S')
logger = logging.getLogger(__name__)

# ==========================================
# 1. КОНФИГУРАЦИЯ (config_loader)
# ==========================================

DEFAULT_CONFIG = {
    "paths": {"input_dir": "./text", "output_docx": "output.docx", "code_dir": None},
    "features": {
        "hyphenation": False,
        "numbering": {"figures": False, "tables": False}, 
        "footnotes": {"mode": "endnotes"}, 
        "mermaid": {"timeout_sec": 30}, 
        "downloads": {"max_size_mb": 50}
    },
    "toc": {"enabled": True, "depth": 3},
    "header_footer": {"header": {"enabled": False}, "footer": {"enabled": True}},
    # ДОБАВЛЯЕМ ДЕФОЛТНЫЕ ЗАГОЛОВКИ, ЧТОБЫ ТИТУЛЬНИК НЕ ПРОПАДАЛ
    "title_page": {"enabled": False},
    "appendix": {"enabled": True, "title": "Приложение A. Верификационный код"}
}

DEFAULT_THEME = {
    "page": {"size": "A4", "margins_inches": {"top": 1.0, "bottom": 1.0, "left": 1.25, "right": 1.25}},
    "text_body": {"font": "PT Serif", "size": 11, "color": "222222", "align": "left", "line_spacing": 1.4, "indent_inches": 0.0, "space_after": 8},
    "lists": {"bullet_marker": "▪", "bibliography_hanging_indent": 0.5},
    "styles": {
        "h1": {"font": "PT Serif", "size": 18, "bold": True, "color": "000000", "align": "left", "space_before": 28, "space_after": 12, "page_break_before": True},
        "h2": {"font": "PT Serif", "size": 16, "bold": True, "color": "000000", "align": "left", "space_before": 20, "space_after": 8},
        "h3": {"font": "PT Serif", "size": 13, "bold": True, "italic": True, "color": "111111", "align": "left", "space_before": 16, "space_after": 6},
        "quote": {"font": "PT Serif", "size": 10.5, "italic": True, "color": "444444", "left_indent": 0.4, "border_left_width": 12, "border_left_color": "CCCCCC"},
        "callout": {"font": "PT Serif", "size": 10.5, "color": "1A1A1A", "background_color": "F4F5F7", "left_indent": 0.2, "right_indent": 0.2},
        "code_inline": {"font": "PT Mono", "size": 9.5, "color": "333333", "background_color": "E8E8E8"},
        "code_block": {"font": "PT Mono", "size": 9, "color": "222222", "background_color": "F5F5F5", "left_indent": 0.2},
        "plaintext": {"font": "Consolas", "size": 9, "color": "333333", "border_top_width": 8, "border_bottom_width": 8, "border_color": "808080", "left_indent": 0.2},
        
        # ДОБАВЛЯЕМ ДЕФОЛТНЫЙ БЛОК ДЛЯ ПРИЛОЖЕНИЯ:
        "appendix_code": {
            "font": "Consolas", "size": 8.5, "color": "222222", 
            "background_color": "F9F9F9", "border_color": "CCCCCC", 
            "left_indent": 0.1, "line_spacing": 1.0, "space_before": 8, "space_after": 8,
            "pyg_keyword": "0000FF", "pyg_string": "A31515", "pyg_comment": "008000", "pyg_number": "098658"
        },
        
        "toc": {"font": "PT Serif", "size": 11, "color": "222222", "h1_bold": True, "indent_per_level": 0.25},
        "footnotes": {"font": "PT Serif", "size": 9, "color": "444444", "line_spacing": 1.1},
        "metadata": {"font": "PT Serif", "size": 10, "color": "333333", "line_spacing": 1.2},
        "caption": {"font": "PT Serif", "size": 9, "italic": True, "color": "555555"},
        "table_font_size": 9,
        "table": {"padding_inches": {"left": 0.04, "right": 0.04, "top": 0.03, "bottom": 0.03}, "border_color": "BFBFBF", "border_width_inner": 4, "border_width_outer": 8, "header_bg": "F2F2F2"}
    }
}

def deep_merge(source, destination):
    for key, value in source.items():
        if isinstance(value, dict):
            node = destination.setdefault(key, {})
            deep_merge(value, node)
        else:
            destination[key] = value
    return destination

def load_yaml(path, default_dict):
    if not os.path.exists(path):
        logger.warning(f"Файл {path} не найден. Используются дефолтные настройки.")
        return default_dict
    with open(path, 'r', encoding='utf-8') as f:
        user_config = yaml.safe_load(f) or {}
        return deep_merge(user_config, default_dict)

def get_runtime_settings():
    parser = argparse.ArgumentParser(description="Build DOCX from Markdown")
    # Меняем default на None, чтобы понять, передал ли пользователь путь явно
    parser.add_argument("-c", "--config", default=None, help="Путь к config.yaml")
    parser.add_argument("-t", "--theme", default=None, help="Путь к theme.yaml")
    parser.add_argument("-i", "--input", help="Папка с Markdown-файлами")
    parser.add_argument("-o", "--output", help="Имя итогового DOCX")
    parser.add_argument("--code", help="Путь к файлу или папке с кодом")
    args = parser.parse_args()

    # 1. Определяем папку с исходниками (откуда будем искать локальные конфиги)
    input_dir = args.input if args.input else DEFAULT_CONFIG["paths"]["input_dir"]
    
    # 2. Определяем путь к config.yaml
    if args.config:
        config_path = args.config  # Приоритет 1: явно указан в CLI
    elif os.path.exists(os.path.join(input_dir, "config.yaml")):
        config_path = os.path.join(input_dir, "config.yaml") # Приоритет 2: лежит в папке текста
    else:
        config_path = "config.yaml" # Приоритет 3: лежит рядом со скриптом

    # 3. Определяем путь к theme.yaml
    if args.theme:
        theme_path = args.theme
    elif os.path.exists(os.path.join(input_dir, "theme.yaml")):
        theme_path = os.path.join(input_dir, "theme.yaml")
    else:
        theme_path = "theme.yaml"

    # Загружаем конфиги
    config = load_yaml(config_path, DEFAULT_CONFIG.copy())
    theme = load_yaml(theme_path, DEFAULT_THEME.copy())

    # Переопределяем пути, если они были переданы в CLI
    if args.input: config["paths"]["input_dir"] = args.input
    if args.output: config["paths"]["output_docx"] = args.output
    if args.code: config["paths"]["code_dir"] = args.code
    
    return config, theme

# ==========================================
# 2. КОНТЕКСТ (context)
# ==========================================
@dataclass
class BuildContext:
    figure_counter: int = 0
    table_counter: int = 0
    footnotes_registry: dict = field(default_factory=dict)
    def next_figure(self) -> int:
        self.figure_counter += 1
        return self.figure_counter
    def next_table(self) -> int:
        self.table_counter += 1
        return self.table_counter

# ==========================================
# 3. СНОСКИ (footnotes_manager)
# ==========================================
class FootnotesManager:
    def __init__(self, mode="endnotes"):
        self.mode = mode
        self.notes = {}
        self.current_id = 1
        self.w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    def add_note(self, text):
        note_id = self.current_id
        self.notes[note_id] = text
        self.current_id += 1
        return note_id

    def inject_to_paragraph(self, paragraph, note_id):
        if self.mode == "footnotes":
            run = paragraph.add_run()
            rPr = run._element.get_or_add_rPr()
            rStyle = OxmlElement('w:rStyle')
            rStyle.set(qn('w:val'), 'FootnoteReference')
            rPr.append(rStyle)
            footnote_ref = OxmlElement('w:footnoteReference')
            footnote_ref.set(qn('w:id'), str(note_id))
            run._element.append(footnote_ref)
        else:
            run = paragraph.add_run(f"[{note_id}]")
            run.font.superscript = True

    def finalize_docx(self, docx_path):
        # Если нет сносок — ничего не делаем
        if not self.notes:
            return
            
        if self.mode == "endnotes":
            self._add_endnotes_section(docx_path)
        elif self.mode == "footnotes":
            self._inject_footnotes_xml(docx_path)

    def _add_endnotes_section(self, docx_path):
        doc = Document(docx_path)
        doc.add_page_break()
        doc.add_heading('Примечания', level=1)
        for note_id, text in self.notes.items():
            p = doc.add_paragraph()
            run = p.add_run(f"[{note_id}] ")
            run.font.superscript = True
            p.add_run(text)
        doc.save(docx_path)

    def _inject_footnotes_xml(self, docx_path):
        temp_dir = docx_path + "_temp"
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        footnotes_path = os.path.join(temp_dir, 'word', 'footnotes.xml')
        xml_template = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="{self.w_ns}">
    <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
    <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
'''
        for note_id, text in self.notes.items():
            xml_template += f'''
    <w:footnote w:id="{note_id}">
        <w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> {text}</w:t></w:r></w:p>
    </w:footnote>'''
        xml_template += "\n</w:footnotes>"
        with open(footnotes_path, 'w', encoding='utf-8') as f:
            f.write(xml_template)

        ct_path = os.path.join(temp_dir, '[Content_Types].xml')
        tree = etree.parse(ct_path)
        root = tree.getroot()
        ns = '{http://schemas.openxmlformats.org/package/2006/content-types}'
        if not any(elem.get('PartName') == '/word/footnotes.xml' for elem in root.findall(f'{ns}Override')):
            override = etree.SubElement(root, f'{ns}Override')
            override.set('PartName', '/word/footnotes.xml')
            override.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml')
            tree.write(ct_path, xml_declaration=True, encoding='UTF-8')

        rels_path = os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')
        tree = etree.parse(rels_path)
        root = tree.getroot()
        rel_ids = [rel.get('Id') for rel in root]
        max_id = max([int(rid[3:]) for rid in rel_ids if rid.startswith('rId')] + [0])
        new_rid = f"rId{max_id + 1}"
        rel = etree.SubElement(root, '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')
        rel.set('Id', new_rid); rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes'); rel.set('Target', 'footnotes.xml')
        tree.write(rels_path, xml_declaration=True, encoding='UTF-8')
        shutil.make_archive(temp_dir, 'zip', temp_dir)
        shutil.move(temp_dir + '.zip', docx_path)
        shutil.rmtree(temp_dir)

# ==========================================
# 4. УТИЛИТЫ (render_utils)
# ==========================================

def detect_language(text):
    return "ru-RU" if re.search(r'[а-яА-ЯёЁ]', text) else "en-US"

def set_run_language(run, lang_code):
    rPr = run._r.get_or_add_rPr()
    lang = OxmlElement('w:lang'); lang.set(qn('w:val'), lang_code); rPr.append(lang)

def apply_typography(text, lang="ru-RU"):
    if not text: return text
    NBSP = '\u00A0'
    
    # УДАЛЕНО: Замена дефиса на U+2011 (именно она съедала дефисы)
    
    # 2. Приклеиваем закрывающие скобки и знаки препинания
    text = re.sub(r'\s+([\)\]\}.,;:!?])', NBSP + r'\1', text)
    
    # 3. Приклеиваем открывающие скобки
    text = re.sub(r'([\(\[\{])\s+', r'\1' + NBSP, text)
    
    # 4. Приклеиваем короткие предлоги и союзы
    short_words = r'(?:в|во|на|по|с|со|к|ко|о|об|от|до|из|за|для|при|не|ни|и|а|но|да|же|ли|бы)' if lang == "ru-RU" else r'(?:a|an|the|of|in|to|at|on|by|for|with|and|or|but)'
    text = re.sub(rf'\b({short_words})\s+', r'\1' + NBSP, text, flags=re.IGNORECASE)
    
    # Заменяем дефис на неразрывный (U+2011) только между буквами
    text = re.sub(r'(?<=[a-zA-Zа-яА-ЯёЁ])-(?=[a-zA-Zа-яА-ЯёЁ])', '\u2011', text)
    
    return text

def add_hyperlink(paragraph, url, text, theme):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink'); hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r'); rPr = OxmlElement('w:rPr')
    
    # Берем цвет из theme.yaml (по умолчанию благородный 1F497D)
    link_color = theme.get("styles", {}).get("link_color", "1F497D").replace("#", "")
    
    c = OxmlElement('w:color'); c.set(qn('w:val'), link_color); rPr.append(c)
    u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rPr.append(u)
    new_run.append(rPr); t = OxmlElement('w:t'); t.text = text; new_run.append(t)
    hyperlink.append(new_run); paragraph._p.append(hyperlink)

def add_toc(doc, theme):
    toc_cfg = theme["styles"]["toc"]
    
    # Читаем выравнивание из конфига (по умолчанию left)
    align_str = toc_cfg.get("title_align", "left")
    align = WD_ALIGN_PARAGRAPH.LEFT if align_str == "left" else WD_ALIGN_PARAGRAPH.CENTER

    p_title = doc.add_paragraph()
    p_title.paragraph_format.page_break_before = True
    p_title.alignment = align  # <--- Применяем выравнивание
    p_title.paragraph_format.first_line_indent = Pt(0)
    p_title.paragraph_format.space_before = Pt(12)
    p_title.paragraph_format.space_after = Pt(12)
    
    run_title = p_title.add_run("Содержание / Table of Contents")
    run_title.bold = True; run_title.font.size = Pt(16); run_title.font.name = toc_cfg["font"]; run_title.font.color.rgb = RGBColor.from_string("1A1A1A")
    
    p = doc.add_paragraph(); p.paragraph_format.first_line_indent = Pt(0)
    run = p.add_run(); fldChar = OxmlElement('w:fldChar'); fldChar.set(qn('w:fldCharType'), 'begin'); run._r.append(fldChar)
    run2 = p.add_run(); instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = 'TOC \\o "1-3" \\h \\z \\u'; run2._r.append(instrText)
    run3 = p.add_run(); fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'separate'); run3._r.append(fldChar2)
    run4 = p.add_run("Обновите поле (F9), чтобы увидеть содержание...")
    run4.font.name = toc_cfg["font"]; run4.font.size = Pt(10); run4.font.italic = True; run4.font.color.rgb = RGBColor.from_string("666666")
    run5 = p.add_run(); fldChar3 = OxmlElement('w:fldChar'); fldChar3.set(qn('w:fldCharType'), 'end'); run5._r.append(fldChar3)

def latex_to_omml(latex_str):
    """Конвертирует LaTeX в OMML (нативный формат Word)."""
    try:
        import latex2mathml.converter
        from mathml2omml import convert as mml2omml
        from lxml import etree
        
        clean = latex_str.replace('$$', '').replace('\\[', '').replace('\\]', '').strip()
        
        # ИСПРАВЛЕНИЕ: \| и \parallel теперь заменяются на || (двойную черту)
        clean = clean.replace(r'\parallel', '||').replace(r'\|', '||').replace(r'\mid', '|')
        
        # Убираем \mathbf, если он ломает парсер
        clean = re.sub(r'\\mathbf\{([^}]+)\}', r'\1', clean)
        
        mathml = latex2mathml.converter.convert(clean)
        omml_str = mml2omml(mathml)
        if not omml_str:
            return None
            
        omml_str = re.sub(r'<\?xml[^>]*\?>', '', omml_str).strip()
        wrapped = f'<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{omml_str}</m:oMathPara>'
        root = etree.fromstring(wrapped.encode('utf-8'), parser=etree.XMLParser(recover=True))
        return root[0] if len(root) > 0 else None
    except Exception as e:
        logger.warning(f"[WARN] latex_to_omml failed: {e}")
        return None
        
def add_math(container, latex_str, is_block=True):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt
    
    # ЗАЩИТА ОТ CJK-СИМВОЛОВ (Китайский, Японский, Корейский)
    if re.search(r'[\u4e00-\u9fff]', latex_str):
        # Очищаем строку от LaTeX-мусора
        clean_text = latex_str.replace('$$', '').replace('\\[', '').replace('\\]', '')
        clean_text = clean_text.replace('\\text{', '').replace('}', '')
        clean_text = clean_text.replace('\\to', '→').strip()
        
        if is_block:
            p = container.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run(clean_text)
        else:
            run = container.add_run(clean_text)
        
        # Применяем моноширинный шрифт для аккуратности
        run.font.name = "Consolas"
        run.font.size = Pt(10)
        return
        
    # СТАНДАРТНАЯ ОБРАБОТКА ФОРМУЛ
    omml_element = latex_to_omml(latex_str)
    
    if omml_element is not None:
        if is_block:
            p = container.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.keep_together = True
        else:
            p = container
        p._p.append(omml_element)
    else:
        # Fallback, если конвертация не удалась
        logger.warning(f"[WARN] OMML failed for: {latex_str}")
        if is_block: 
            container.add_paragraph(f"[{latex_str}]")
        else: 
            container.add_run(f"[{latex_str}]")

def set_clean_grid_borders(table, theme):
    tbl_cfg = theme.get("styles", {}).get("table", {})
    color_inner = tbl_cfg.get("border_color", "BFBFBF").replace("#", "")
    color_outer = "404040"
    sz_inner = str(tbl_cfg.get("border_width_inner", 4)); sz_outer = str(tbl_cfg.get("border_width_outer", 8))
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn('w:tblBorders')): tblPr.remove(old)
    tblBorders = OxmlElement('w:tblBorders')
    top = OxmlElement('w:top'); top.set(qn('w:val'), 'single'); top.set(qn('w:sz'), sz_outer); top.set(qn('w:space'), '0'); top.set(qn('w:color'), color_outer); tblBorders.append(top)
    bottom = OxmlElement('w:bottom'); bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), sz_outer); bottom.set(qn('w:space'), '0'); bottom.set(qn('w:color'), color_outer); tblBorders.append(bottom)
    for edge in ('left', 'right', 'insideH', 'insideV'):
        el = OxmlElement(f'w:{edge}'); el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), sz_inner); el.set(qn('w:space'), '0'); el.set(qn('w:color'), color_inner); tblBorders.append(el)
    tblPr.append(tblBorders)
    if len(table.rows) > 0:
        for cell in table.rows[0].cells:
            tcPr = cell._tc.get_or_add_tcPr(); tcBorders = tcPr.find(qn('w:tcBorders'))
            if tcBorders is None: tcBorders = OxmlElement('w:tcBorders'); tcPr.append(tcBorders)
            hdr_bottom = OxmlElement('w:bottom'); hdr_bottom.set(qn('w:val'), 'single'); hdr_bottom.set(qn('w:sz'), sz_outer); hdr_bottom.set(qn('w:space'), '0'); hdr_bottom.set(qn('w:color'), color_outer); tcBorders.append(hdr_bottom)

def set_cell_shading(cell, color_hex="F2F2F2"):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:shd')): tcPr.remove(old)
    shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), color_hex.replace("#", "")); tcPr.append(shd)

def set_table_cell_margins(table, top=0.03, start=0.04, bottom=0.03, end=0.04):
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn('w:tblCellMar')): tblPr.remove(old)
    tblCellMar = OxmlElement('w:tblCellMar')
    for m, val in [('top', top), ('left', start), ('bottom', bottom), ('right', end)]:
        node = OxmlElement(f'w:{m}'); node.set(qn('w:w'), str(int(round(val * 1440)))); node.set(qn('w:type'), 'dxa'); tblCellMar.append(node)
    tblPr.append(tblCellMar)

def set_repeat_table_header(row):
    trPr = row._tr.get_or_add_trPr(); tblHeader = OxmlElement('w:tblHeader'); tblHeader.set(qn('w:val'), "true"); trPr.append(tblHeader)

def set_table_row_cant_split(row):
    trPr = row._tr.get_or_add_trPr(); cantSplit = OxmlElement('w:cantSplit'); trPr.append(cantSplit)

def disable_hyphenation(paragraph):
    pPr = paragraph._p.get_or_add_pPr(); suppress = OxmlElement('w:suppressAutoHyphens'); suppress.set(qn('w:val'), 'true'); pPr.append(suppress)

def apply_caption_style(p, theme):
    cap_cfg = theme.get("styles", {}).get("caption", {})
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(12)
    for run in p.runs:
        run.font.name = cap_cfg.get("font", "PT Serif"); run.font.size = Pt(cap_cfg.get("size", 9)); run.italic = True
        run.font.color.rgb = RGBColor.from_string(cap_cfg.get("color", "555555"))

def protect_table_text_integrity(text):
    if not text: return text
    NBSP = '\u00A0'
    text = re.sub(r'(\d+[\d.,]*)\s+([a-zA-Zа-яА-Я%]+)', r'\1' + NBSP + r'\2', text)
    text = re.sub(r'(\d+[\d.,]*)\s+(\([A-Za-zА-Яа-я]+\))', r'\1' + NBSP + r'\2', text)
    text = re.sub(r'([A-Za-z0-9]+)\s+(\([А-Яа-яA-Za-z]+\))', r'\1' + NBSP + r'\2', text)
    text = re.sub(r'(\d+)\s*([*+/–-])\s*(\d+)\s*(=)', r'\1' + NBSP + r'\2' + NBSP + r'\3' + NBSP + r'\4', text)
    text = re.sub(r'(\$[^\$]+\$)\s+(\((?:бит|bits)\))', r'\1' + NBSP + r'\2', text)
    return text

def is_numeric(text):
    if not text: return False
    t = text.strip().replace('−', '-').replace('\u00A0', ' ').strip('`"\'')
    t = re.sub(r'\s*\([A-Za-zА-Яа-я]+\)$', '', t)
    t = re.sub(r'[\s%x\^]+$', '', t)
    t = re.sub(r'^(?:L\d+|\d+)\s*(?:passes|слоев|слой|layers?|layer)?$', '', t, flags=re.IGNORECASE)
    if not t: return True
    return bool(re.match(r'^[+-]?\d*([.,]\d+)?$', t))

def add_mermaid(doc, code, config):
    timeout = config["features"]["mermaid"].get("timeout_sec", 30)
    with tempfile.NamedTemporaryFile(suffix=".mmd", delete=False, mode='w', encoding='utf-8') as f:
        f.write(code); in_path = f.name
    out_path = in_path.replace(".mmd", ".png")
    try:
        subprocess.run(["mmdc", "-i", in_path, "-o", out_path, "-b", "transparent"], check=True, timeout=timeout, capture_output=True)
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(out_path, width=Inches(5.5))
    except Exception as e:
        logger.warning(f"[WARN] Mermaid не отрендерена: {e}")
        p = doc.add_paragraph(); run = p.add_run(f"[Mermaid-диаграмма: {code}]"); run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
    finally:
        for path in (in_path, out_path):
            if os.path.exists(path):
                try: os.remove(path)
                except OSError: pass

def set_paragraph_borders(p, top_width=0, bottom_width=0, left_width=0, color="CCCCCC"):
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    for edge, sz in [('top', top_width), ('bottom', bottom_width), ('left', left_width)]:
        if sz > 0:
            el = OxmlElement(f'w:{edge}')
            el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz)); el.set(qn('w:space'), '6'); el.set(qn('w:color'), color.replace("#", ""))
            pBdr.append(el)
    if len(pBdr) > 0: pPr.append(pBdr)

def set_paragraph_background(p, color_hex):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), color_hex.replace("#", "")); pPr.append(shd)

def add_appendix(doc, config, theme):
    """Генерирует Приложение с исходным кодом, если указан --code или code_dir."""
    code_dir = config.get("paths", {}).get("code_dir")
    
    # Если путь не передан, выходим (никаких "Приложений" и пустых страниц)
    if not code_dir or not os.path.exists(code_dir):
        return

    app_cfg = config.get("appendix", {})
    # Разрешаем отключить приложение вручную в YAML, если нужно
    if not app_cfg.get("enabled", True):
        return

    logger.info(f"[INFO] Генерация Приложения с кодом из: {code_dir}")
    
    # 1. Заголовок H1 для приложения
    title = app_cfg.get("title", "Приложение A. Верификационный код")
    p_title = doc.add_paragraph(style='Heading 1')
    p_title.paragraph_format.page_break_before = True
    
    # Берем настройки из theme.yaml
    h1_cfg = theme.get("styles", {}).get("h1", {})
    run_title = p_title.add_run(title)
    run_title.font.name = h1_cfg.get("font", "PT Serif")
    run_title.font.size = Pt(h1_cfg.get("size", 18))
    run_title.bold = h1_cfg.get("bold", True)
    run_title.font.color.rgb = RGBColor.from_string(h1_cfg.get("color", "000000").replace("#", ""))

    # 2. Сбор файлов
    files_to_process = []
    if os.path.isfile(code_dir):
        files_to_process.append(code_dir)
    elif os.path.isdir(code_dir):
        # Рекурсивный обход папки
        for root, _, files in os.walk(code_dir):
            for f in files:
                # Игнорируем скрытые файлы и системный мусор
                if not f.startswith('.') and f.endswith(('.py', '.cpp', '.h', '.json', '.txt', '.md', '.js', '.css', '.html')):
                    files_to_process.append(os.path.join(root, f))

    files_to_process.sort()
    code_cfg = theme.get("styles", {}).get("appendix_code", {})

    # 3. Генерация листингов
    for filepath in files_to_process:
        filename = os.path.basename(filepath)
        
        # Заголовок H2 для имени файла
        p_h2 = doc.add_paragraph(style='Heading 2')
        p_h2.paragraph_format.space_before = Pt(16)
        
        # Берем настройки из theme.yaml
        h2_cfg = theme.get("styles", {}).get("h2", {})
        run_h2 = p_h2.add_run(f"Файл: {filename}")
        run_h2.font.name = h2_cfg.get("font", "PT Serif")
        run_h2.font.size = Pt(h2_cfg.get("size", 16))
        run_h2.bold = h2_cfg.get("bold", True)
        run_h2.font.color.rgb = RGBColor.from_string(h2_cfg.get("color", "000000").replace("#", ""))

        # Чтение кода
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                code_text = f.read()
        except Exception as e:
            logger.error(f"Не удалось прочитать {filepath}: {e}")
            continue

        # Создаем параграф для кода
        p_code = doc.add_paragraph()
        p_code.paragraph_format.first_line_indent = Pt(0)
        p_code.paragraph_format.left_indent = Inches(code_cfg.get("left_indent", 0.1))
        p_code.paragraph_format.space_before = Pt(code_cfg.get("space_before", 8))
        p_code.paragraph_format.space_after = Pt(code_cfg.get("space_after", 8))
        p_code.paragraph_format.line_spacing = code_cfg.get("line_spacing", 1.0)
        
        set_paragraph_background(p_code, code_cfg.get("background_color", "F9F9F9"))
        set_paragraph_borders(p_code, top_width=4, bottom_width=4, color=code_cfg.get("border_color", "CCCCCC"))

        # Pygments: токенизация
        try:
            lexer = get_lexer_for_filename(filename)
        except Exception:
            lexer = TextLexer()

        tokens = lex(code_text, lexer)
        
        for ttype, value in tokens:
            run = p_code.add_run(value)
            run.font.name = code_cfg.get("font", "Consolas")
            run.font.size = Pt(code_cfg.get("size", 8.5))
            
            # Отключаем проверку орфографии Word (красные линии)
            rPr = run._r.get_or_add_rPr()
            noProof = OxmlElement('w:noProof')
            rPr.append(noProof)

            # Применяем цвета в зависимости от токена
            color = code_cfg.get("color", "222222")
            if ttype in Token.Keyword: color = code_cfg.get("pyg_keyword", "0000FF")
            elif ttype in Token.String: color = code_cfg.get("pyg_string", "A31515")
            elif ttype in Token.Comment: color = code_cfg.get("pyg_comment", "008000")
            elif ttype in Token.Number: color = code_cfg.get("pyg_number", "098658")
            
            run.font.color.rgb = RGBColor.from_string(color)

    # Добавляем разрыв после приложения (если нужно)
    doc.add_page_break()

def apply_metadata(doc, config):
    """Применяет метаданные из config.yaml к свойствам документа."""
    meta = config.get("metadata", {})
    if not meta:
        return

    props = doc.core_properties
    if meta.get("title"): props.title = meta["title"]
    if meta.get("subject"): props.subject = meta["subject"]
    if meta.get("author"): props.author = meta["author"]
    if meta.get("keywords"): props.keywords = meta["keywords"]
    if meta.get("category"): props.category = meta["category"]
    
    # Записываем лицензию и доп. инфо в Comments (Описание)
    if meta.get("comments"): props.comments = meta["comments"]
    
    from datetime import datetime
    props.last_modified_by = meta.get("author", "build_docx")
    props.revision = 1
    
    logger.info("[INFO] Метаданные документа (включая лицензию) применены.")
    
# ==========================================
# 5. ТИТУЛЬНЫЙ ЛИСТ (title_page)
# ==========================================

def add_horizontal_line(paragraph, color_hex="404040"):
    pPr = paragraph._p.get_or_add_pPr(); pBdr = OxmlElement('w:pBdr'); bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), '8'); bottom.set(qn('w:space'), '1'); bottom.set(qn('w:color'), color_hex.replace("#", ""))
    pBdr.append(bottom); pPr.append(pBdr)

def add_title_page(doc, config, theme):
    tp = config.get("title_page", {})
    if not tp.get("enabled", False): return

    font_name = theme.get("text_body", {}).get("font", "PT Serif")
    font_size = theme.get("text_body", {}).get("size", 11)
    text_color = theme.get("text_body", {}).get("color", "222222").replace("#", "")

    def add_par(text, size, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, italic=False, space_after=0, color=None):
        c = color if color else text_color
        
        # Разбиваем текст по символу $ (формулы)
        parts = re.split(r'(?<!\\)\$', text)
        
        # Группируем элементы в строки (по символу \n)
        lines = []
        current_line = []
        for i, part in enumerate(parts):
            if i % 2 == 1: # Формула
                current_line.append(("math", part))
            else: # Текст
                sub_lines = part.split('\n')
                for j, sub_line in enumerate(sub_lines):
                    if sub_line:
                        current_line.append(("text", sub_line))
                    if j < len(sub_lines) - 1:
                        lines.append(current_line)
                        current_line = []
        if current_line:
            lines.append(current_line)
            
        last_p = None
        for line_elements in lines:
            # Создаем ОТДЕЛЬНЫЙ параграф для каждой строки
            p = doc.add_paragraph()
            p.alignment = align
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0) # Убираем отступ между строками!
            p.paragraph_format.line_spacing = 1.15
            disable_hyphenation(p)
            
            for elem_type, content in line_elements:
                if elem_type == "math":
                    add_math(p, content, is_block=False)
                else:
                    lang = detect_language(content)
                    run = p.add_run(apply_typography(content, lang))
                    run.font.name = font_name
                    run.font.size = Pt(size)
                    run.bold = bold
                    run.italic = italic
                    run.font.color.rgb = RGBColor.from_string(c)
            last_p = p
            
        # Отступ снизу применяем ТОЛЬКО к последней строке
        if last_p:
            last_p.paragraph_format.space_after = Pt(space_after)
            
        return last_p

    # --- ВОТ ЭТОГО КУСКА НЕ ХВАТАЛО ---
    add_par(tp.get("title_ru", ""), size=font_size + 5, bold=True, space_after=10)
    add_par(tp.get("title_en", ""), size=font_size + 2, italic=True, space_after=35)
    add_par(tp.get("author", ""), size=font_size + 2, bold=True, space_after=5)
    add_par(tp.get("role", ""), size=font_size, space_after=12)

    if tp.get("email"):
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(3)
        add_hyperlink(p, f"mailto:{tp.get('email').replace('Email: ', '')}", tp.get('email'), theme)
    if tp.get("repository"):
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(12)
        add_hyperlink(p, tp.get('repository').replace('Repository: ', ''), tp.get('repository'), theme)

    p_line = doc.add_paragraph(); p_line.paragraph_format.first_line_indent = Pt(0); p_line.paragraph_format.space_before = Pt(4); p_line.paragraph_format.space_after = Pt(12)
    add_horizontal_line(p_line, color_hex="404040")

    # --- Описание издания (Английский блок) ---
    if tp.get("edition_title_en"):
        add_par(tp.get("edition_title_en", ""), size=font_size - 1, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=2)
    if tp.get("edition_desc_en"):
        add_par(tp.get("edition_desc_en", ""), size=font_size - 1, italic=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=10)

    # --- Описание издания (Русский блок) ---
    if tp.get("edition_title_ru"):
        add_par(tp.get("edition_title_ru", ""), size=font_size - 1, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=2)
    if tp.get("edition_desc_ru"):
        add_par(tp.get("edition_desc_ru", ""), size=font_size - 1, italic=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=14)

    def add_abstract(title, text):
        if not text: return
        add_par(title, size=font_size - 1.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=2)
        p_abs = add_par(text, size=font_size - 1.5, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=12)
        p_abs.paragraph_format.line_spacing = 1.15

    add_abstract(tp.get("abstract_en_title"), tp.get("abstract_en_text"))
    add_abstract(tp.get("abstract_ru_title"), tp.get("abstract_ru_text"))
    
    logger.info("[INFO] Академический титульный лист сгенерирован.")
    # -----------------------------------

# ==========================================
# 6. ПАРСЕР (parser)
# ==========================================

def process_inline(paragraph, children, config, theme, ctx, fm):
    is_bold = False; is_italic = False; is_strike = False
    link_url = None; link_text = ""
    normal_font = theme.get("text_body", {}).get("font", "PT Serif")
    normal_color = theme.get("text_body", {}).get("color", "222222").replace("#", "")
    code_cfg = theme.get("styles", {}).get("code_inline", {})
    
    for child in children:
        if child.type == "link_open":
            link_url = child.attrs.get('href', ''); link_text = ""
        elif child.type == "link_close":
            if link_url: add_hyperlink(paragraph, link_url, link_text, theme)
            link_url = None
        elif child.type == "text":
            if link_url is not None:
                link_text += child.content
            else:
                lang = detect_language(child.content)
                # ИСПРАВЛЕНИЕ: Заменяем \n на пробел, чтобы не было переносов там, где их нет
                content = child.content.replace('\n', ' ')
                run = paragraph.add_run(apply_typography(content, lang))
                run.font.name = normal_font
                run.bold, run.italic, run.font.strike = is_bold, is_italic, is_strike
                run.font.color.rgb = RGBColor.from_string(normal_color)
                set_run_language(run, lang)

        elif child.type == "strong_open": is_bold = True
        elif child.type == "strong_close": is_bold = False
        elif child.type == "em_open": is_italic = True
        elif child.type == "em_close": is_italic = False
        elif child.type == "s_open": is_strike = True
        elif child.type == "s_close": is_strike = False
        elif child.type == "code_inline":
            run = paragraph.add_run(child.content)
            run.font.name = code_cfg.get("font", "PT Mono"); run.font.size = Pt(code_cfg.get("size", 9.5))
            run.font.color.rgb = RGBColor.from_string(code_cfg.get("color", "C7254E"))
            # Добавляем фон для инлайн-кода
            rPr = run._r.get_or_add_rPr()
            shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), code_cfg.get("background_color", "F9F2F4")); rPr.append(shd)
        elif child.type == "math_inline":
            # Если внутри формулы есть китайские иероглифы — выводим как текст, а не формулу
            if re.search(r'[\u4e00-\u9fff]', child.content):
                run = paragraph.add_run(child.content)
                run.font.name = normal_font
                run.font.size = Pt(tb_cfg.get("size", 11))
                run.font.color.rgb = RGBColor.from_string(normal_color)
            else:
                add_math(paragraph, child.content, is_block=False)
            
        elif child.type == "footnote_ref":
            fm.inject_to_paragraph(paragraph, int(child.meta['id']))
        elif child.type == "softbreak":
            paragraph.add_run(" ") # Мягкий перенос (одиночный Enter) просто превращаем в пробел
            
        elif child.type == "hardbreak":
            paragraph.add_run(" ").add_break() # Жесткий перенос (два пробела) оставляем с разрывом

def clean_for_measurement(text):
    t = text.replace('`', '').replace('"', '').replace("'", "").replace('$', '')
    t = re.sub(r'\\text\{([^}]+)\}', r'\1', t); t = re.sub(r'\\[a-zA-Z]+', 'Δ', t)
    t = t.replace('{', '').replace('}', ''); t = re.sub(r'[\r\n]+', ' ', t)
    return t.strip()

def get_indivisible_chunks(text):
    t = protect_table_text_integrity(text); t = clean_for_measurement(t)
    return [c for c in re.split(r'[ \t\r\n]+', t) if c]

def process_table(tokens, start_idx, doc, config, theme, ctx, fm):
    i = start_idx + 1
    raw_table_data = []; current_row = []
    while tokens[i].type != "table_close":
        if tokens[i].type == "tr_open": current_row = []
        elif tokens[i].type == "tr_close": raw_table_data.append(current_row)
        elif tokens[i].type in ["th_open", "td_open"]: current_row.append(tokens[i+1]); i += 1 
        i += 1
    table_data = [row for row in raw_table_data if any(inline.children for inline in row)]
    if not table_data: return i
    rows, cols = len(table_data), len(table_data[0])
    
    if config.get("features", {}).get("numbering", {}).get("tables", False):
        num = ctx.next_table(); p = doc.add_paragraph(); p.add_run(f"Таблица {num}."); apply_caption_style(p, theme)

    col_max_chunk = [0] * cols; col_avg_len = [0.0] * cols; is_num_col = [True] * cols
    for c_idx in range(cols):
        h_token = table_data[0][c_idx]
        h_text = "".join(c.content for c in h_token.children if hasattr(c, 'content')) if h_token.children else h_token.content
        col_max_chunk[c_idx] = max((len(ch) for ch in get_indivisible_chunks(h_text)), default=0)
    for c_idx in range(cols):
        col_total_len = 0; data_count = 0; has_non_empty = False
        for r_idx in range(1, rows):
            token = table_data[r_idx][c_idx]
            cell_text = ("".join(c.content for c in token.children if hasattr(c, 'content')) if token.children else token.content).strip()
            if cell_text:
                has_non_empty = True; data_count += 1; col_total_len += len(cell_text)
                if not is_numeric(cell_text): is_num_col[c_idx] = False
                for ch in get_indivisible_chunks(cell_text):
                    if len(ch) > col_max_chunk[c_idx]: col_max_chunk[c_idx] = len(ch)
        if not has_non_empty: is_num_col[c_idx] = False
        col_avg_len[c_idx] = (col_total_len / data_count) if data_count > 0 else 1.0

    table_font_size = 8.5 if cols >= 7 else float(theme.get("styles", {}).get("table_font_size", 9))
    pad_cfg = theme.get("styles", {}).get("table", {}).get("padding_inches", {"left": 0.04, "right": 0.04, "top": 0.03, "bottom": 0.03})
    pad_left = pad_cfg.get("left", 0.04); pad_right = pad_cfg.get("right", 0.04); pad_top = pad_cfg.get("top", 0.03); pad_bottom = pad_cfg.get("bottom", 0.03)
    
    section = doc.sections[0]; available_width_emu = int(section.page_width - section.left_margin - section.right_margin); available_width_in = available_width_emu / 914400.0
    char_coeff = 0.053 if table_font_size <= 8.5 else 0.060
    min_widths_in = [max((col_max_chunk[c] * char_coeff) + (pad_left + pad_right) + 0.02, 0.44 if is_num_col[c] else 0) for c in range(cols)]
    total_min_w = sum(min_widths_in)
    
    if total_min_w >= available_width_in:
        scale = available_width_in / total_min_w; final_widths_in = [w * scale for w in min_widths_in]
    else:
        remaining_in = available_width_in - total_min_w
        text_cols = [c for c in range(cols) if not is_num_col[c]] or list(range(cols))
        sum_text_weights = sum(max(col_avg_len[c], 2.0) for c in text_cols)
        final_widths_in = list(min_widths_in)
        for c in text_cols: final_widths_in[c] += remaining_in * (max(col_avg_len[c], 2.0) / sum_text_weights)

    table = doc.add_table(rows=rows, cols=cols); table.alignment = WD_TABLE_ALIGNMENT.LEFT; table.autofit = False
    tblPr = table._tbl.tblPr
    tblLayout = OxmlElement('w:tblLayout'); tblLayout.set(qn('w:type'), 'fixed'); tblPr.append(tblLayout)
    tblW = OxmlElement('w:tblW'); tblW.set(qn('w:w'), str(int(available_width_emu / 635))); tblW.set(qn('w:type'), 'dxa'); tblPr.append(tblW)
    
    set_clean_grid_borders(table, theme); set_table_cell_margins(table, top=pad_top, start=pad_left, bottom=pad_bottom, end=pad_right)
    set_repeat_table_header(table.rows[0])
    for row in table.rows: set_table_row_cant_split(row)

    header_bg = theme.get("styles", {}).get("table", {}).get("header_bg", "F2F2F2")
    normal_font = theme.get("text_body", {}).get("font", "PT Serif")

    for r_idx, row in enumerate(table_data):
        for c_idx, inline_token in enumerate(row):
            cell = table.cell(r_idx, c_idx); cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER; p = cell.paragraphs[0]
            p.paragraph_format.first_line_indent = Pt(0); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0); p.paragraph_format.line_spacing = 1.05
            disable_hyphenation(p)
            for child in inline_token.children:
                if child.type == "text": child.content = protect_table_text_integrity(child.content)
            process_inline(p, inline_token.children, config, theme, ctx, fm)
            if r_idx == 0:
                set_cell_shading(cell, header_bg); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs: run.font.name = normal_font; run.font.size = Pt(table_font_size); run.bold = True; run.font.color.rgb = RGBColor.from_string("1A1A1A")
            else:
                cell_text = cell.text.strip()
                if is_num_col[c_idx]: p.alignment = WD_ALIGN_PARAGRAPH.CENTER if re.match(r'^(?:L\d+|[A-Z]|\d+)$', cell_text) else WD_ALIGN_PARAGRAPH.RIGHT
                else: p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for run in p.runs:
                    run.font.name = normal_font; run.font.size = Pt(table_font_size)
                    if not run.font.color or not run.font.color.rgb: run.font.color.rgb = RGBColor.from_string("1A1A1A")
    
    for c_idx in range(cols):
        col_width_emu = int(final_widths_in[c_idx] * 914400); table.columns[c_idx].width = Emu(col_width_emu)
        for r_idx in range(rows): table.cell(r_idx, c_idx).width = Emu(col_width_emu)

    spacer = doc.add_paragraph(); spacer.paragraph_format.first_line_indent = Pt(0); spacer.paragraph_format.space_before = Pt(0); spacer.paragraph_format.space_after = Pt(8); spacer.paragraph_format.line_spacing = Pt(1); r_sp = spacer.add_run(); r_sp.font.size = Pt(1)
    return i

def parse_md_to_docx(doc, text, config, theme, ctx, fm):
    md = MarkdownIt("commonmark", {"html": True}).enable("table").enable("strikethrough")
    md.use(dollarmath_plugin); md.use(footnote_plugin); tokens = md.parse(text)
    
    i = 0; skip_next_list = False; list_stack = []
    list_counters = [] # <--- ВАЖНО: Массив для хранения текущих номеров
    in_bibliography = False
    
    normal_font = theme.get("text_body", {}).get("font", "PT Serif"); normal_color = theme.get("text_body", {}).get("color", "222222").replace("#", "")
    tb_cfg = theme.get("text_body", {})
    list_cfg = theme.get("lists", {})
    
    while i < len(tokens):
        token = tokens[i]
        
        # --- Отслеживаем заголовок списка литературы ---
        if token.type == "heading_open" and i + 1 < len(tokens):
            h_text = tokens[i+1].content.lower().strip()
            if any(k in h_text for k in ["литератур", "references", "библиограф"]):
                in_bibliography = True
            else:
                in_bibliography = False

            if h_text in ["содержание", "оглавление", "table of contents"]:
                skip_next_list = True; i += 3; continue
                
        if skip_next_list and token.type in ["bullet_list_open", "ordered_list_open"]:
            depth = 1; i += 1
            while i < len(tokens) and depth > 0:
                if tokens[i].type == token.type: depth += 1
                elif tokens[i].type == token.type.replace("_open", "_close"): depth -= 1
                i += 1
            skip_next_list = False; continue

        # --- ЛОГИКА СПИСКОВ ---
        if token.type in ["bullet_list_open", "ordered_list_open"]: 
            list_stack.append("bullet" if token.type == "bullet_list_open" else "ordered")
            list_counters.append(1) # Инициализируем счетчик для этого уровня
            i += 1; continue
            
        elif token.type in ["bullet_list_close", "ordered_list_close"]: 
            if list_stack: list_stack.pop()
            if list_counters: list_counters.pop()
            
            # Воздух после списка
            if len(doc.paragraphs) > 0:
                last_p = doc.paragraphs[-1]
                last_p.paragraph_format.space_after = Pt(tb_cfg.get("space_after", 8))
                
            i += 1; continue
            
        elif token.type == "list_item_close":
            # Увеличиваем счетчик только когда закрывается элемент нумерованного списка
            if list_stack and list_stack[-1] == "ordered" and list_counters:
                list_counters[-1] += 1
            i += 1; continue

        # --- ОБРАБОТКА ЗАГОЛОВКОВ ---
        if token.type == "heading_open":
            p = doc.add_heading(level=int(token.tag[1]))
            p.paragraph_format.first_line_indent = Pt(0); 
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT; 
            p.paragraph_format.keep_with_next = True
            disable_hyphenation(p) 
            
            h_key = f"h{token.tag[1]}"
            if h_key in theme.get("styles", {}):
                h_cfg = theme["styles"][h_key]
                p.paragraph_format.space_before = Pt(h_cfg.get("space_before", 12)); p.paragraph_format.space_after = Pt(h_cfg.get("space_after", 6))
            process_inline(p, tokens[i+1].children, config, theme, ctx, fm)
            for run in p.runs: run.bold = True
            i += 3
            
        # --- ОБРАБОТКА ПАРАГРАФОВ ---
        elif token.type == "paragraph_open":
            p = doc.add_paragraph()
            if list_stack:
                p.paragraph_format.first_line_indent = Pt(0)
                p.paragraph_format.left_indent = Inches(list_cfg.get("indent_inches", 0.25) * len(list_stack))
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_before = Pt(list_cfg.get("space_before", 2))
                p.paragraph_format.space_after = Pt(list_cfg.get("space_after", 4))
                p.paragraph_format.line_spacing = 1.15
                
                if in_bibliography:
                    p.paragraph_format.left_indent = Inches(list_cfg.get("bibliography_hanging_indent", 0.5))
                    p.paragraph_format.first_line_indent = Inches(-list_cfg.get("bibliography_hanging_indent", 0.5))
                    p.paragraph_format.space_after = Pt(6) 
                else:
                    if list_stack[-1] == "bullet":
                        run_bullet = p.add_run(f"{list_cfg.get('bullet_marker', '▪')}\u00A0")
                        run_bullet.font.name = normal_font
                        run_bullet.font.size = Pt(tb_cfg.get("size", 11))
                        run_bullet.font.color.rgb = RGBColor.from_string(normal_color)
                    elif list_stack[-1] == "ordered":
                        num = list_counters[-1] if list_counters else 1
                        run_num = p.add_run(f"{num}.\u00A0")
                        run_num.font.name = normal_font
                        run_num.font.size = Pt(tb_cfg.get("size", 11))
                        run_num.font.color.rgb = RGBColor.from_string(normal_color)
            else:
                p.paragraph_format.first_line_indent = Inches(tb_cfg.get("indent_inches", 0.0))
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.line_spacing = tb_cfg.get("line_spacing", 1.4)
                p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(tb_cfg.get("space_after", 8))
            process_inline(p, tokens[i+1].children, config, theme, ctx, fm); 
            
            if p.text.strip().endswith((':', '：')):
                p.paragraph_format.keep_with_next = True
            i += 3
            
        # ... (Остальной код для fence, math_block, table_open, blockquote_open, footnote_block оставьте как был)
        elif token.type == "fence":
            lang = token.info.strip().lower(); p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Pt(0)
            if lang == "mermaid":
                add_mermaid(doc, token.content, config); i += 1; continue
            elif lang in ("text", "txt", "plaintext", ""):
                cfg = theme["styles"]["plaintext"]
                p.paragraph_format.left_indent = Inches(cfg.get("left_indent", 0.2))
                p.paragraph_format.space_before = Pt(cfg.get("space_before", 12)); p.paragraph_format.space_after = Pt(cfg.get("space_after", 12))
                p.paragraph_format.line_spacing = cfg.get("line_spacing", 1.0)
                
                clean_content = token.content.rstrip('\n') 
                # ВНИМАНИЕ: используем clean_content, а не token.content!
                run = p.add_run(clean_content) 
                run.font.name = cfg["font"]; run.font.size = Pt(cfg["size"]); run.font.color.rgb = RGBColor.from_string(cfg["color"])
                set_paragraph_borders(p, top_width=cfg.get("border_top_width", 0), bottom_width=cfg.get("border_bottom_width", 0), color=cfg.get("border_color", "808080"))
            else:
                cfg = theme["styles"]["code_block"]
                p.paragraph_format.left_indent = Inches(cfg.get("left_indent", 0.2))
                p.paragraph_format.space_before = Pt(cfg.get("space_before", 12)); p.paragraph_format.space_after = Pt(cfg.get("space_after", 12))
                p.paragraph_format.line_spacing = cfg.get("line_spacing", 1.0)
                
                clean_content = token.content.rstrip('\n') 
                # ВНИМАНИЕ: используем clean_content
                run = p.add_run(clean_content) 
                run.font.name = cfg["font"]; run.font.size = Pt(cfg["size"]); run.font.color.rgb = RGBColor.from_string(cfg["color"])
                if cfg.get("background_color"): set_paragraph_background(p, cfg["background_color"])
            i += 1
        elif token.type == "math_block":
            add_math(doc, token.content, is_block=True); i += 1
        elif token.type == "table_open":
            i = process_table(tokens, i, doc, config, theme, ctx, fm)
        elif token.type == "blockquote_open":
            # ... (без изменений)
            inline_token = tokens[i + 2]; first_child = inline_token.children[0] if inline_token.children else None
            if first_child and first_child.type == "text" and first_child.content.startswith("[!"):
                match = re.match(r"\[!(\w+)\]\s*(.*)", first_child.content, re.DOTALL)
                if match:
                    first_child.content = match.group(2); p = doc.add_paragraph()
                    cfg = theme["styles"]["callout"]
                    p.paragraph_format.first_line_indent = Pt(0); p.paragraph_format.left_indent = Inches(cfg.get("left_indent", 0.2)); p.paragraph_format.right_indent = Inches(cfg.get("right_indent", 0.2))
                    p.paragraph_format.space_before = Pt(cfg.get("space_before", 12)); p.paragraph_format.space_after = Pt(cfg.get("space_after", 12))
                    p.paragraph_format.line_spacing = 1.2
                    set_paragraph_background(p, cfg.get("background_color", "F4F5F7"))
                    run = p.add_run(f"[{match.group(1).upper()}] "); run.bold = True; run.font.name = cfg["font"]; run.font.size = Pt(cfg["size"]); run.font.color.rgb = RGBColor.from_string(cfg["color"])
                    process_inline(p, inline_token.children, config, theme, ctx, fm); i += 5; continue
            p = doc.add_paragraph(); cfg = theme["styles"]["quote"]
            p.paragraph_format.first_line_indent = Pt(0); p.paragraph_format.left_indent = Inches(cfg.get("left_indent", 0.4))
            p.paragraph_format.space_before = Pt(12); p.paragraph_format.space_after = Pt(12); p.paragraph_format.line_spacing = 1.2
            process_inline(p, inline_token.children, config, theme, ctx, fm)
            for run in p.runs: run.italic = True; run.font.name = cfg["font"]; run.font.size = Pt(cfg["size"]); run.font.color.rgb = RGBColor.from_string(cfg["color"])
            set_paragraph_borders(p, left_width=cfg.get("border_left_width", 12), color=cfg.get("border_left_color", "CCCCCC"))
            i += 5
        elif token.type == "footnote_block":
            # ... (без изменений)
            for fn_token in token.children:
                if fn_token.type == "footnote_open":
                    note_id = int(fn_token.meta['id']); inline = fn_token.children[1]
                    text_fn = "".join(c.content for c in inline.children if c.type == "text"); fm.add_note(text_fn)
            i += 1
        else: i += 1

# ==========================================
# 7. СБОРЩИК ДОКУМЕНТА (build_docx)
# ==========================================

def setup_headers_footers(doc, config, theme):
    hf_cfg = config.get("header_footer", {})
    section = doc.sections[0]
    section.different_first_page_header_footer = True 
    normal_font = theme.get("text_body", {}).get("font", "PT Serif")
    
    def get_align(align_str):
        if align_str == "left": return WD_ALIGN_PARAGRAPH.LEFT
        if align_str == "right": return WD_ALIGN_PARAGRAPH.RIGHT
        return WD_ALIGN_PARAGRAPH.CENTER

    # --- ВЕРХНИЙ КОЛОНТИТУЛ ---
    header_cfg = hf_cfg.get("header", {})
    if header_cfg.get("enabled", False):
        header = section.header
        header.is_linked_to_previous = False # Отвязываем от предыдущих секций
        h_p = header.paragraphs[0]
        h_p.alignment = get_align(header_cfg.get("alignment", "right"))
        h_p.paragraph_format.first_line_indent = Pt(0)
        
        if header_cfg.get("text"):
            run_h = h_p.add_run(header_cfg["text"])
            run_h.font.name = normal_font
            # Читаем размер и цвет из конфига (по умолчанию 9pt и серый)
            run_h.font.size = Pt(header_cfg.get("font_size", 9))
            color_hex = header_cfg.get("color", "888888")
            run_h.font.color.rgb = RGBColor.from_string(color_hex)

    # --- НИЖНИЙ КОЛОНТИТУЛ (Титульная страница) ---
    footer_cfg = hf_cfg.get("footer", {})
    if footer_cfg.get("enabled", True):
        first_footer = section.first_page_footer
        first_footer.is_linked_to_previous = False
        fp = first_footer.paragraphs[0]
        fp.alignment = get_align(footer_cfg.get("alignment", "center"))
        fp.paragraph_format.first_line_indent = Pt(0)
        
        first_text = footer_cfg.get("first_page_text", "")
        if first_text:
            run_fp = fp.add_run(first_text)
            run_fp.font.name = normal_font
            run_fp.font.size = Pt(9) # Для титула оставляем 9pt
            run_fp.font.color.rgb = RGBColor(128, 128, 128)
    
    # --- НИЖНИЙ КОЛОНТИТУЛ (Остальные страницы) ---
    if footer_cfg.get("enabled", True) and footer_cfg.get("show_page_numbers", True):
        footer = section.footer
        footer.is_linked_to_previous = False
        p = footer.paragraphs[0]
        p.alignment = get_align(footer_cfg.get("alignment", "center"))
        p.paragraph_format.first_line_indent = Pt(0)
        
        run = p.add_run()
        run.font.name = normal_font
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(128, 128, 128)
        
        # Вставка поля PAGE
        fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = 'PAGE'
        fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'separate')
        t = OxmlElement('w:t'); t.text = "1"
        fldChar3 = OxmlElement('w:fldChar'); fldChar3.set(qn('w:fldCharType'), 'end')
        run._r.append(fldChar1); run._r.append(instrText); run._r.append(fldChar2); run._r.append(t); run._r.append(fldChar3)

def set_document_settings(doc, config):
    settings = doc.settings.element
    updateFields = OxmlElement('w:updateFields'); updateFields.set(qn('w:val'), "true"); settings.append(updateFields)
    
    # Читаем настройку из конфига (по умолчанию False)
    enable_hyphen = config.get("features", {}).get("hyphenation", False)
    if enable_hyphen:
        autoHyphen = OxmlElement('w:autoHyphenation'); autoHyphen.set(qn('w:val'), "true"); settings.append(autoHyphen)

def setup_document_styles(doc, theme):
    section = doc.sections[0]
    page = theme["page"]
    if page["size"].upper() == "A4": section.page_width, section.page_height = Inches(8.27), Inches(11.69)
    for side, margin in page["margins_inches"].items(): setattr(section, f"{side}_margin", Inches(margin))

    styles = doc.styles
    tb_cfg = theme["text_body"]
    text_color = tb_cfg.get("color", "222222").replace("#", "")

    normal_style = styles['Normal']
    normal_style.font.name = tb_cfg["font"]; normal_style.font.size = Pt(tb_cfg["size"]); normal_style.font.color.rgb = RGBColor.from_string(text_color)
    normal_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    normal_style.paragraph_format.line_spacing = tb_cfg.get("line_spacing", 1.4)
    normal_style.paragraph_format.first_line_indent = Inches(tb_cfg.get("indent_inches", 0.0))
    normal_style.paragraph_format.space_before = Pt(0); normal_style.paragraph_format.space_after = Pt(tb_cfg.get("space_after", 8))

    for level in (1, 2, 3):
        h_name = f'Heading {level}'; h_cfg = theme["styles"][f'h{level}']
        try: h_style = styles[h_name]
        except KeyError: h_style = styles.add_style(h_name, WD_STYLE_TYPE.PARAGRAPH)
        h_style.font.name = h_cfg["font"]; h_style.font.size = Pt(h_cfg["size"]); h_style.font.bold = h_cfg.get("bold", True); h_style.font.italic = h_cfg.get("italic", False)
        h_style.font.color.rgb = RGBColor.from_string(h_cfg.get("color", text_color).replace("#", ""))
        h_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        h_style.paragraph_format.first_line_indent = Pt(0)
        h_style.paragraph_format.keep_with_next = True
        
        h_style.paragraph_format.line_spacing = 1.0
        # НОВОЕ: Принудительный перенос страницы перед H1 (если включено в YAML)
        h_style.paragraph_format.page_break_before = h_cfg.get("page_break_before", False)
        
        h_style.paragraph_format.space_before = Pt(h_cfg.get("space_before", 12)); h_style.paragraph_format.space_after = Pt(h_cfg.get("space_after", 6))

    try: fn_style = styles['FootnoteText']
    except KeyError: fn_style = styles.add_style('FootnoteText', WD_STYLE_TYPE.PARAGRAPH)
    fn_cfg = theme["styles"]["footnotes"]
    fn_style.font.name = fn_cfg["font"]; fn_style.font.size = Pt(fn_cfg["size"]); fn_style.font.color.rgb = RGBColor.from_string(fn_cfg["color"]); fn_style.paragraph_format.first_line_indent = Pt(0); fn_style.paragraph_format.line_spacing = fn_cfg.get("line_spacing", 1.1)

    try: fn_ref = styles['FootnoteReference']
    except KeyError: fn_ref = styles.add_style('FootnoteReference', WD_STYLE_TYPE.CHARACTER)
    fn_ref.font.name = tb_cfg["font"]; fn_ref.font.size = Pt(tb_cfg["size"])

    for level in range(1, 4):
        style_name = f'TOC {level}'
        try: toc_style = styles[style_name]
        except KeyError: toc_style = styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
        toc_cfg = theme["styles"]["toc"]
        toc_style.font.name = toc_cfg["font"]; toc_style.font.size = Pt(toc_cfg["size"]); toc_style.font.color.rgb = RGBColor.from_string(toc_cfg["color"])
        toc_style.font.bold = (level == 1) and toc_cfg.get("h1_bold", True)
        toc_style.paragraph_format.first_line_indent = Pt(0)
        if level > 1: toc_style.paragraph_format.left_indent = Inches(toc_cfg.get("indent_per_level", 0.25) * (level - 1))
        toc_style.paragraph_format.space_after = Pt(toc_cfg.get("space_after", 4)); toc_style.paragraph_format.space_before = Pt(toc_cfg.get("space_before_h1", 12) if level == 1 else 0)

def build_document(config_path=None, theme_path=None):
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    if config_path and theme_path:
        config = load_yaml(config_path, DEFAULT_CONFIG.copy()); theme = load_yaml(theme_path, DEFAULT_THEME.copy())
    else:
        config, theme = get_runtime_settings() 
    
    input_dir = config["paths"]["input_dir"]; output_docx = config["paths"]["output_docx"]
    ctx = BuildContext(); fm = FootnotesManager(mode=config.get("features", {}).get("footnotes", {}).get("mode", "endnotes"))

    doc = Document()
    setup_document_styles(doc, theme)
    setup_headers_footers(doc, config, theme)
    set_document_settings(doc, config) 

    add_title_page(doc, config, theme)
    add_toc(doc, theme)
    
    # Обработка Markdown
    if os.path.exists(input_dir):
        files = sorted([f for f in os.listdir(input_dir) if f.endswith(".md")])
        for f in files:
            logging.info(f"Обработка: {f}")
            with open(os.path.join(input_dir, f), "r", encoding="utf-8") as file:
                parse_md_to_docx(doc, file.read(), config, theme, ctx, fm)
    else:
        logging.error(f"Папка {input_dir} не найдена!")

    # НОВОЕ: Добавление Приложения
    add_appendix(doc, config, theme)

    apply_metadata(doc, config)

    doc.save(output_docx)
    fm.finalize_docx(output_docx)
    logging.info(f"[SUCCESS] Документ сохранен: {output_docx}")

if __name__ == "__main__":
    build_document()