import os
import sys
import re
import argparse
import asyncio
from bs4 import BeautifulSoup
import openpyxl
from playwright.async_api import async_playwright

try:
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Mm, Pt, RGBColor
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

QUESTIONNAIRE_STYLES = """
*, *::before, *::after { box-sizing: border-box; }

body {
    margin: 0;
    padding: 20px;
    font-family: "Arial Unicode MS", "Arial Unicode", "Segoe UI", Arial, sans-serif;
    font-size: 14px;
    line-height: 1.5;
    color: #111827;
    background: #f8fafc;
}

.questionnaire-container {
    width: 100%;
    max-width: 820px;
    margin: 0 auto;
    background: #fff;
    padding: 32px 40px;
    border-radius: 8px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
}

.qn-title {
    text-align: center;
    font-size: 22px;
    font-weight: 700;
    margin: 0 0 6px 0;
    color: #0f172a;
}

.qn-subtitle {
    text-align: center;
    font-size: 13.5px;
    color: #475569;
    margin: 0 0 24px 0;
    padding-bottom: 12px;
    border-bottom: 2px solid #334155;
}

.section-header {
    font-size: 15px;
    font-weight: 700;
    color: #0f172a;
    background: #f0f4f8;
    padding: 10px 14px;
    margin: 28px 0 14px 0;
    border-top: 2px solid #334155;
    border-bottom: 2px solid #334155;
    page-break-inside: avoid;
    page-break-after: avoid;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
}

.note-item {
    background: #f8fafc;
    border-left: 4px solid #3b82f6;
    padding: 12px 16px;
    margin: 16px 0;
    font-size: 13.5px;
    color: #334155;
    line-height: 1.55;
    page-break-inside: avoid;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
}

.note-item p {
    margin: 0 0 8px 0;
}

.note-item p:last-child {
    margin-bottom: 0;
}

.question-item {
    margin-bottom: 16px;
    padding-bottom: 12px;
    border-bottom: 1px solid #f1f5f9;
    page-break-inside: avoid;
}

.question-prompt {
    font-size: 14px;
    font-weight: 600;
    color: #0f172a;
    margin-bottom: 4px;
    line-height: 1.45;
}

.question-prompt p {
    margin: 0 0 4px 0;
}

.question-prompt p:last-child {
    margin-bottom: 0;
}

.hint {
    font-size: 12.5px;
    font-style: italic;
    color: #1e40af;
    margin: 2px 0 6px 12px;
    line-height: 1.4;
}

.choice-list {
    margin: 6px 0 6px 14px;
    display: flex;
    flex-direction: column;
    gap: 4px;
}

.choice-item {
    display: flex;
    align-items: baseline;
    font-size: 13.5px;
    color: #1e293b;
    line-height: 1.45;
}

.choice-marker {
    display: inline-block;
    width: 24px;
    flex-shrink: 0;
    color: #64748b;
    font-family: monospace;
    font-size: 13px;
}

.choice-label {
    flex: 1;
}

.answer-blank {
    margin: 6px 0 4px 14px;
    color: #64748b;
    font-size: 13px;
    line-height: 1.4;
}

@media screen and (max-width: 768px) {
    body {
        padding: 10px;
    }
    .questionnaire-container {
        padding: 16px 18px;
    }
}

@media print {
    @page {
        size: A4 portrait;
        margin: 15mm 12mm;
    }

    body {
        background: #fff;
        padding: 0;
        font-size: 10pt;
        line-height: 1.4;
    }

    .questionnaire-container {
        max-width: 100%;
        padding: 0;
        box-shadow: none;
        border-radius: 0;
    }

    .section-header {
        background: #f0f4f8 !important;
        color: #000 !important;
        border-top: 1.5pt solid #333 !important;
        border-bottom: 1.5pt solid #333 !important;
        page-break-after: avoid;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
    }

    .note-item {
        background: #f8fafc !important;
        border-left: 3pt solid #3b82f6 !important;
        page-break-inside: avoid;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
    }

    .question-item {
        page-break-inside: avoid;
        border-bottom: 0.5pt solid #e2e8f0;
        margin-bottom: 10pt;
        padding-bottom: 8pt;
    }
}
"""

SYSTEM_METADATA_TYPES = {
    'start', 'end', 'today', 'deviceid', 'phonenumber', 'simserial',
    'subscriberid', 'username', 'email', 'audit', 'text-audit',
    'start-geopoint', 'background-audio', 'caseid', 'caseread', 'casesave'
}

def split_type(raw_type):
    """Split XLSForm type string into base_type and list_name."""
    if not raw_type:
        return '', ''
    parts = raw_type.strip().split(None, 1)
    base = parts[0].strip()
    arg = parts[1].strip() if len(parts) > 1 else ''
    return base, arg

def is_section_start(raw_type):
    """Check if row begins a group or repeat."""
    t = raw_type.lower().strip()
    return t in ['begin group', 'begin_group', 'begin repeat', 'begin_repeat'] or t.startswith(('begin group', 'begin_group', 'begin repeat', 'begin_repeat'))

def is_section_end(raw_type):
    """Check if row closes a group or repeat."""
    t = raw_type.lower().strip()
    return t in ['end group', 'end_group', 'end repeat', 'end_repeat'] or t.startswith(('end group', 'end_group', 'end repeat', 'end_repeat'))

def extract_language_name(header_str):
    """Extract language name from label::Language header."""
    match = re.match(r'^(?:label|hint|constraint_message)::(.+)$', header_str.strip())
    if match:
        return match.group(1).strip()
    return None

def clean_language_suffix(lang_name):
    """Clean language name for filename suffix (e.g. 'English (en)' -> 'English')."""
    clean = re.sub(r'\(.*?\)', '', lang_name).strip()
    clean = re.sub(r'[^a-zA-Z0-9_-]', '_', clean)
    clean = re.sub(r'_+', '_', clean).strip('_')
    return clean

def get_header_index(header_map, *keys):
    """Get index from header_map matching one of the keys case-insensitively."""
    for key in keys:
        target = key.strip().lower()
        for h, idx in header_map.items():
            if h.strip().lower() == target:
                return idx
    return None

def get_lang_value(dict_of_langs, target_lang):
    """Look up language value with exact, substring, or default fallback."""
    if not dict_of_langs:
        return ''
    if target_lang in dict_of_langs and dict_of_langs[target_lang]:
        return dict_of_langs[target_lang]
    for l_key, val in dict_of_langs.items():
        if target_lang.lower() in l_key.lower() and val:
            return val
    if 'Default' in dict_of_langs and dict_of_langs['Default']:
        return dict_of_langs['Default']
    for val in dict_of_langs.values():
        if val:
            return val
    return ''

class XLSFormParser:
    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.wb = openpyxl.load_workbook(excel_path, data_only=True)
        self.sheets = {name.lower().strip(): name for name in self.wb.sheetnames}
        
        self.settings = self._parse_settings()
        self.languages = self._detect_languages()
        self.choices = self._parse_choices()
        self.survey_rows = self._parse_survey()

    def _get_sheet(self, target_name):
        actual_name = self.sheets.get(target_name.lower())
        if actual_name:
            return self.wb[actual_name]
        return None

    def _parse_settings(self):
        settings = {}
        sheet = self._get_sheet('settings')
        if not sheet or sheet.max_row < 2:
            return settings
        
        headers = [str(cell.value or '').strip() for cell in sheet[1]]
        values = [str(cell.value or '').strip() if cell.value is not None else '' for cell in sheet[2]]
        
        for h, v in zip(headers, values):
            if h:
                settings[h] = v
        return settings

    def _detect_languages(self):
        languages = []
        for sheet_name in ['survey', 'choices']:
            sheet = self._get_sheet(sheet_name)
            if not sheet or sheet.max_row < 1:
                continue
            headers = [str(cell.value or '').strip() for cell in sheet[1]]
            for h in headers:
                if h.startswith('label::'):
                    lang = extract_language_name(h)
                    if lang and lang not in languages:
                        languages.append(lang)
        if not languages:
            languages.append('Default')
        return languages

    def _parse_choices(self):
        choices = {}
        sheet = self._get_sheet('choices')
        if not sheet or sheet.max_row < 2:
            return choices

        headers = [str(cell.value or '').strip() for cell in sheet[1]]
        header_map = {h: i for i, h in enumerate(headers) if h}

        list_name_col = get_header_index(header_map, 'list_name', 'list name')
        name_col = get_header_index(header_map, 'name', 'value')

        if list_name_col is None or name_col is None:
            return choices

        for row_idx in range(2, sheet.max_row + 1):
            row_cells = sheet[row_idx]
            list_name_val = str(row_cells[list_name_col].value or '').strip()
            name_val = str(row_cells[name_col].value or '').strip()

            if not list_name_val or not name_val:
                continue

            if list_name_val not in choices:
                choices[list_name_val] = []

            labels = {}
            for lang in self.languages:
                col_idx = get_header_index(header_map, f'label::{lang}', 'label')
                if col_idx is None and lang != 'Default':
                    for h, idx in header_map.items():
                        if h.startswith('label::') and lang.lower() in h.lower():
                            col_idx = idx
                            break
                if col_idx is None:
                    col_idx = get_header_index(header_map, 'label')

                if col_idx is not None and col_idx < len(row_cells):
                    val = row_cells[col_idx].value
                    labels[lang] = str(val).strip() if val is not None else name_val
                else:
                    labels[lang] = name_val

            choices[list_name_val].append({
                'name': name_val,
                'labels': labels
            })

        return choices

    def _parse_survey(self):
        survey_rows = []
        sheet = self._get_sheet('survey')
        if not sheet or sheet.max_row < 2:
            return survey_rows

        headers = [str(cell.value or '').strip() for cell in sheet[1]]
        header_map = {h: i for i, h in enumerate(headers) if h}

        type_col = get_header_index(header_map, 'type')
        name_col = get_header_index(header_map, 'name')
        if type_col is None:
            return survey_rows

        for row_idx in range(2, sheet.max_row + 1):
            row_cells = sheet[row_idx]
            type_val = str(row_cells[type_col].value or '').strip()
            if not type_val:
                continue

            name_val = ''
            if name_col is not None and name_col < len(row_cells):
                name_val = str(row_cells[name_col].value or '').strip()

            labels = {}
            hints = {}

            for lang in self.languages:
                # Label
                col_idx = get_header_index(header_map, f'label::{lang}')
                if col_idx is None and lang != 'Default':
                    for h, idx in header_map.items():
                        if h.startswith('label::') and lang.lower() in h.lower():
                            col_idx = idx
                            break
                if col_idx is None:
                    col_idx = get_header_index(header_map, 'label')
                labels[lang] = str(row_cells[col_idx].value or '').strip() if (col_idx is not None and col_idx < len(row_cells) and row_cells[col_idx].value is not None) else ''

                # Hint
                col_idx = get_header_index(header_map, f'hint::{lang}')
                if col_idx is None and lang != 'Default':
                    for h, idx in header_map.items():
                        if h.startswith('hint::') and lang.lower() in h.lower():
                            col_idx = idx
                            break
                if col_idx is None:
                    col_idx = get_header_index(header_map, 'hint')
                hints[lang] = str(row_cells[col_idx].value or '').strip() if (col_idx is not None and col_idx < len(row_cells) and row_cells[col_idx].value is not None) else ''

            survey_rows.append({
                'type': type_val,
                'name': name_val,
                'labels': labels,
                'hints': hints,
            })

        return survey_rows

def get_choice_list_by_name(choices_dict, list_name):
    """Flexible lookup for choice list matching name case-insensitively."""
    if not list_name or not choices_dict:
        return []
    if list_name in choices_dict:
        return choices_dict[list_name]
    for key, items in choices_dict.items():
        if key.strip().lower() == list_name.strip().lower():
            return items
    return []

def get_html_response_placeholder(base_type, item, q_label):
    """Generate HTML response blank based on open-ended question type."""
    name_lower = (item.get('name') or '').lower()
    label_lower = (q_label or '').lower()
    
    if base_type == 'text':
        is_long = any(k in name_lower for k in ['note', 'detail', 'comment', 'specify', 'other', 'explain', 'desc', 'reason']) or any(k in label_lower for k in ['explain', 'detail', 'specify', 'notes', 'comments', 'describe', 'reasons'])
        if is_long:
            return '<div class="answer-blank">____________________________________________________________________<br>____________________________________________________________________</div>'
        return '<div class="answer-blank">____________________________________________________________________</div>'
    elif base_type in ('integer', 'decimal', 'range'):
        return '<div class="answer-blank">[ ____________________ ]</div>'
    elif base_type == 'date':
        return '<div class="answer-blank">[ ____ / ____ / ________ ] (DD/MM/YYYY)</div>'
    elif base_type == 'time':
        return '<div class="answer-blank">[ ____ : ____ ] (HH:MM)</div>'
    elif base_type == 'datetime':
        return '<div class="answer-blank">[ ____/____/________  ____:____ ]</div>'
    elif base_type in ('geopoint', 'geotrace', 'geoshape'):
        return '<div class="answer-blank"><em>[ GPS Location Coordinates: ________________________________ ]</em></div>'
    elif base_type in ('image', 'audio', 'video', 'file'):
        return f'<div class="answer-blank"><em>[ Attach / Record {base_type.capitalize()}: ________________________________ ]</em></div>'
    elif base_type == 'barcode':
        return '<div class="answer-blank"><em>[ Barcode / QR: ________________________________ ]</em></div>'
    elif base_type == 'acknowledge':
        return '<div class="answer-blank">[  ] Confirmed / Acknowledged</div>'
    else:
        return '<div class="answer-blank">____________________________________________________________________</div>'

def build_questionnaire_html(parser, language='Default', landscape=False):
    """Build responsive table-free HTML questionnaire for interviewers/respondents."""
    soup = BeautifulSoup('<!DOCTYPE html><html><head></head><body></body></html>', 'html.parser')
    
    head = soup.head
    head.append(soup.new_tag('meta', attrs={'charset': 'utf-8'}))
    head.append(soup.new_tag('meta', attrs={'name': 'viewport', 'content': 'width=device-width, initial-scale=1.0'}))
    
    title_text = parser.settings.get('form_title') or parser.settings.get('title') or os.path.splitext(os.path.basename(parser.excel_path))[0]
    title_tag = soup.new_tag('title')
    title_tag.string = f"{title_text} - Questionnaire ({language})"
    head.append(title_tag)
    
    style_content = QUESTIONNAIRE_STYLES
    if landscape:
        style_content = style_content.replace('size: A4 portrait;', 'size: A4 landscape;')
        style_content = style_content.replace('max-width: 820px;', 'max-width: 1100px;')

    style_tag = soup.new_tag('style')
    style_tag.string = style_content
    head.append(style_tag)

    body = soup.body
    container = soup.new_tag('div', attrs={'class': 'questionnaire-container'})
    body.append(container)

    # Document Header Title
    h1 = soup.new_tag('div', attrs={'class': 'qn-title'})
    h1.string = title_text
    container.append(h1)

    sub_title = soup.new_tag('div', attrs={'class': 'qn-subtitle'})
    sub_title.string = f"Interview Questionnaire — Language: {language}" if language != 'Default' else "Interview Questionnaire"
    container.append(sub_title)

    for item in parser.survey_rows:
        raw_type = item['type'].strip()
        base_type, list_name = split_type(raw_type)

        # Skip background system metadata types
        if base_type in SYSTEM_METADATA_TYPES:
            continue

        # Skip calculation rows and hidden logic
        if base_type in ('calculate', 'hidden') or raw_type == 'calculate':
            continue

        # Skip closing section tags
        if is_section_end(raw_type):
            continue

        # Section header check for group and repeat beginnings
        if is_section_start(raw_type):
            label = get_lang_value(item['labels'], language) or item['name']
            if not label and 'group' in raw_type.lower():
                label = item['name']
            if not label:
                continue

            sec_div = soup.new_tag('div', attrs={'class': 'section-header'})
            span = soup.new_tag('span')
            span.string = label
            sec_div.append(span)
            container.append(sec_div)
            continue

        # Visible label check
        q_label = get_lang_value(item['labels'], language)
        if not q_label:
            continue

        # Note Rows: render as full-width script / instruction block
        if base_type == 'note':
            note_item = soup.new_tag('div', attrs={'class': 'note-item'})
            note_div = soup.new_tag('div', attrs={'class': 'note-text'})
            note_div.append(BeautifulSoup(q_label, 'html.parser'))
            note_item.append(note_div)
            container.append(note_item)
            continue

        # Standard Interview Question (without table)
        q_item = soup.new_tag('div', attrs={'class': 'question-item'})

        # Question Prompt (no * required)
        div_prompt = soup.new_tag('div', attrs={'class': 'question-prompt'})
        div_prompt.append(BeautifulSoup(q_label, 'html.parser'))
        q_item.append(div_prompt)

        # Hint
        q_hint = get_lang_value(item['hints'], language)
        if q_hint:
            div_hint = soup.new_tag('div', attrs={'class': 'hint'})
            div_hint.string = f"Note: {q_hint}"
            q_item.append(div_hint)

        # Choices or open-ended response blank (NO option codes)
        if base_type.startswith(('select_one', 'select_multiple', 'select_or_other', 'rank')) and list_name:
            choice_list = get_choice_list_by_name(parser.choices, list_name)
            if choice_list:
                is_multi = base_type.startswith('select_multiple')
                marker_str = '[ ]' if is_multi else '( )'

                c_list_div = soup.new_tag('div', attrs={'class': 'choice-list'})
                for choice in choice_list:
                    c_item_div = soup.new_tag('div', attrs={'class': 'choice-item'})
                    c_marker = soup.new_tag('span', attrs={'class': 'choice-marker'})
                    c_marker.string = marker_str

                    c_lbl = soup.new_tag('span', attrs={'class': 'choice-label'})
                    choice_label = get_lang_value(choice['labels'], language) or choice['name']
                    c_lbl.append(BeautifulSoup(choice_label, 'html.parser'))

                    c_item_div.append(c_marker)
                    c_item_div.append(c_lbl)
                    c_list_div.append(c_item_div)

                q_item.append(c_list_div)
            else:
                blank_html = '<div class="answer-blank"><em>[ Options from external file ]</em></div>'
                q_item.append(BeautifulSoup(blank_html, 'html.parser'))
        else:
            blank_html = get_html_response_placeholder(base_type, item, q_label)
            q_item.append(BeautifulSoup(blank_html, 'html.parser'))

        container.append(q_item)

    return str(soup)

# Alias for backwards compatibility
build_codebook_html = build_questionnaire_html

DOCX_FONT = 'Arial Unicode MS'
DOCX_MONO_FONT = 'Consolas'

DOCX_TITLE_PT = 16.0
DOCX_HEADER_PT = 11.0
DOCX_SECTION_PT = 12.0
DOCX_QUESTION_PT = 10.5
DOCX_BODY_PT = 9.5
DOCX_SMALL_PT = 8.5
DOCX_TINY_PT = 7.5

DOCX_COLOR_TEXT = '0F172A'
DOCX_COLOR_HINT = '1E40AF'
DOCX_COLOR_NOTE_BORDER = '3B82F6'

DOCX_FILL_SECTION = 'F0F4F8'
DOCX_FILL_NOTE = 'F8FAFC'

DOCX_RULE_SECTION = '334155'
DOCX_RULE_LIGHT = 'E2E8F0'

# Dimensions for Portrait A4
DOCX_PORTRAIT_WIDTH_MM = 210.0
DOCX_PORTRAIT_HEIGHT_MM = 297.0
DOCX_MARGIN_LEFT_RIGHT_MM = 16.0
DOCX_MARGIN_TOP_BOTTOM_MM = 15.0

# Dimensions for Landscape A4
DOCX_LANDSCAPE_WIDTH_MM = 297.0
DOCX_LANDSCAPE_HEIGHT_MM = 210.0

_DOCX_PPR_ORDER = (
    'w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr', 'w:widowControl',
    'w:numPr', 'w:suppressLineNumbers', 'w:pBdr', 'w:shd', 'w:tabs', 'w:suppressAutoHyphens',
    'w:kinsoku', 'w:wordWrap', 'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE',
    'w:autoSpaceDN', 'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
    'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc', 'w:textDirection',
    'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr',
    'w:sectPr', 'w:pPrChange',
)

_DOCX_RPR_ORDER = (
    'w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps', 'w:strike',
    'w:dstrike', 'w:outline', 'w:shadow', 'w:emboss', 'w:imprint', 'w:noProof', 'w:snapToGrid',
    'w:vanish', 'w:webHidden', 'w:color', 'w:spacing', 'w:w', 'w:kern', 'w:position', 'w:sz',
    'w:szCs', 'w:highlight', 'w:u', 'w:effect', 'w:bdr', 'w:shd', 'w:fitText', 'w:vertAlign',
    'w:rtl', 'w:cs', 'w:em', 'w:lang', 'w:eastAsianLayout', 'w:specVanish', 'w:oMath',
)

def _docx_set_child(parent, order, tag, element=None, replace=True):
    """Insert (or replace) a property child at the position its schema demands."""
    existing = parent.find(qn(tag))
    if existing is not None:
        if not replace:
            return existing
        parent.remove(existing)
    if element is None:
        element = OxmlElement(tag)
    successors = order[order.index(tag) + 1:]
    parent.insert_element_before(element, *successors)
    return element

def _docx_toggle(parent, order, tag, value='true'):
    el = _docx_set_child(parent, order, tag)
    el.set(qn('w:val'), value)
    return el

def _docx_border(tag, sz='4', color='auto', val='single', space='0'):
    """Build a single w:<edge> border element."""
    el = OxmlElement(f'w:{tag}')
    el.set(qn('w:val'), val)
    el.set(qn('w:sz'), sz)
    el.set(qn('w:space'), space)
    el.set(qn('w:color'), color)
    return el

def _docx_shading(fill):
    """Build a w:shd element for solid fill."""
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill)
    return shd

def _docx_style_run(run, size=DOCX_BODY_PT, color=None, bold=False, italic=False, mono=False):
    """Apply font, size, weight, and colour to a run, covering complex-script attributes."""
    name = DOCX_MONO_FONT if mono else DOCX_FONT
    font = run.font
    font.name = name
    font.size = Pt(size)
    font.bold = bold
    font.italic = italic
    if color:
        font.color.rgb = RGBColor.from_string(color)

    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rfonts.set(qn(attr), name)

    _docx_toggle(rpr, _DOCX_RPR_ORDER, 'w:szCs', str(int(round(size * 2))))
    if bold:
        _docx_toggle(rpr, _DOCX_RPR_ORDER, 'w:bCs')
    if italic:
        _docx_toggle(rpr, _DOCX_RPR_ORDER, 'w:iCs')
    return run

def _docx_tight_paragraph(paragraph, space_before=0.0, space_after=1.0):
    """Collapse paragraph spacing so stacked blocks stay compact."""
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(space_before)
    fmt.space_after = Pt(space_after)
    fmt.line_spacing = 1.1

def _docx_setup_document(title_text, subtitle_text=None, landscape=False):
    """Create document with portrait or landscape A4 page setup and clean title banner."""
    doc = Document()

    normal = doc.styles['Normal']
    normal.font.name = DOCX_FONT
    normal.font.size = Pt(DOCX_BODY_PT)
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rfonts.set(qn(attr), DOCX_FONT)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(2)
    normal.paragraph_format.line_spacing = 1.1

    section = doc.sections[0]
    if landscape:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Mm(DOCX_LANDSCAPE_WIDTH_MM)
        section.page_height = Mm(DOCX_LANDSCAPE_HEIGHT_MM)
    else:
        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width = Mm(DOCX_PORTRAIT_WIDTH_MM)
        section.page_height = Mm(DOCX_PORTRAIT_HEIGHT_MM)

    section.top_margin = Mm(DOCX_MARGIN_TOP_BOTTOM_MM)
    section.bottom_margin = Mm(DOCX_MARGIN_TOP_BOTTOM_MM)
    section.left_margin = Mm(DOCX_MARGIN_LEFT_RIGHT_MM)
    section.right_margin = Mm(DOCX_MARGIN_LEFT_RIGHT_MM)

    # Title
    heading = doc.add_paragraph()
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _docx_tight_paragraph(heading, space_before=0.0, space_after=2.0)
    _docx_style_run(heading.add_run(title_text), size=DOCX_TITLE_PT, bold=True, color='0F172A')

    if subtitle_text:
        sub_p = doc.add_paragraph()
        sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _docx_tight_paragraph(sub_p, space_before=1.0, space_after=14.0)
        ppr = sub_p._element.get_or_add_pPr()
        pbdr = _docx_set_child(ppr, _DOCX_PPR_ORDER, 'w:pBdr')
        pbdr.append(_docx_border('bottom', sz='12', color='334155', space='6'))
        _docx_style_run(sub_p.add_run(subtitle_text), size=DOCX_HEADER_PT, color='475569')
    else:
        ppr = heading._element.get_or_add_pPr()
        pbdr = _docx_set_child(ppr, _DOCX_PPR_ORDER, 'w:pBdr')
        pbdr.append(_docx_border('bottom', sz='12', color='334155', space='6'))

    return doc

def _docx_add_formatted_paragraphs(
    doc, text_or_html,
    base_size=DOCX_BODY_PT,
    base_color=None,
    base_bold=False,
    base_italic=False,
    space_before=0.0,
    space_after=2.0,
    left_indent_pt=0.0,
    right_indent_pt=0.0,
    keep_with_next=False,
    shading=None,
    left_border=None,
    bottom_border=None,
    top_border=None
):
    """Parse text_or_html into one or more docx Paragraphs with styled runs and paragraph properties."""
    if not text_or_html:
        p = doc.add_paragraph()
        _docx_tight_paragraph(p, space_before=space_before, space_after=space_after)
        return [p]

    soup = BeautifulSoup(str(text_or_html), 'html.parser')
    p_tags = soup.find_all(['p', 'div'])
    if not p_tags:
        p_tags = [soup]

    paragraphs = []
    num_tags = len(p_tags)

    for i, p_tag in enumerate(p_tags):
        p = doc.add_paragraph()
        
        sb = space_before if i == 0 else 2.0
        sa = space_after if i == num_tags - 1 else 3.0
        _docx_tight_paragraph(p, space_before=sb, space_after=sa)

        ppr = p._element.get_or_add_pPr()

        if keep_with_next:
            _docx_toggle(ppr, _DOCX_PPR_ORDER, 'w:keepNext')

        if left_indent_pt > 0 or right_indent_pt > 0:
            ind = _docx_set_child(ppr, _DOCX_PPR_ORDER, 'w:ind')
            if left_indent_pt > 0:
                ind.set(qn('w:left'), str(int(round(left_indent_pt * 20))))
            if right_indent_pt > 0:
                ind.set(qn('w:right'), str(int(round(right_indent_pt * 20))))

        if shading:
            _docx_set_child(ppr, _DOCX_PPR_ORDER, 'w:shd', _docx_shading(shading))

        if left_border or bottom_border or top_border:
            pbdr = _docx_set_child(ppr, _DOCX_PPR_ORDER, 'w:pBdr')
            if left_border:
                sz, color, space = left_border
                pbdr.append(_docx_border('left', sz=str(sz), color=color, space=str(space)))
            if top_border:
                sz, color, space = top_border
                pbdr.append(_docx_border('top', sz=str(sz), color=color, space=str(space)))
            if bottom_border:
                sz, color, space = bottom_border
                pbdr.append(_docx_border('bottom', sz=str(sz), color=color, space=str(space)))

        def walk(node, is_bold, is_italic):
            for child in getattr(node, 'children', []):
                if isinstance(child, str):
                    text_str = str(child)
                    if text_str:
                        run = p.add_run(text_str)
                        _docx_style_run(run, size=base_size, color=base_color,
                                        bold=is_bold or base_bold,
                                        italic=is_italic or base_italic)
                elif child.name in ['b', 'strong']:
                    walk(child, True, is_italic)
                elif child.name in ['i', 'em']:
                    walk(child, is_bold, True)
                elif child.name == 'br':
                    run = p.add_run('\n')
                    _docx_style_run(run, size=base_size)
                else:
                    walk(child, is_bold, is_italic)

        walk(p_tag, False, False)
        if not p.runs:
            _docx_style_run(p.add_run(''), size=base_size)

        paragraphs.append(p)

    return paragraphs

def _docx_add_section_header(doc, label):
    """Add a styled section header without a table."""
    p = doc.add_paragraph()
    _docx_tight_paragraph(p, space_before=18.0, space_after=6.0)
    ppr = p._element.get_or_add_pPr()
    _docx_toggle(ppr, _DOCX_PPR_ORDER, 'w:keepNext')

    pbdr = _docx_set_child(ppr, _DOCX_PPR_ORDER, 'w:pBdr')
    pbdr.append(_docx_border('top', sz='12', color=DOCX_RULE_SECTION, space='4'))
    pbdr.append(_docx_border('bottom', sz='12', color=DOCX_RULE_SECTION, space='4'))

    _docx_set_child(ppr, _DOCX_PPR_ORDER, 'w:shd', _docx_shading(DOCX_FILL_SECTION))

    ind = _docx_set_child(ppr, _DOCX_PPR_ORDER, 'w:ind')
    ind.set(qn('w:left'), '120')
    ind.set(qn('w:right'), '120')

    run = p.add_run(label)
    _docx_style_run(run, size=DOCX_SECTION_PT, bold=True, color='0F172A')
    return p

def _docx_add_response_placeholder_flow(doc, base_type, item, q_label):
    """Add response blank line or brackets below an open-ended question without tables."""
    p = doc.add_paragraph()
    _docx_tight_paragraph(p, space_before=2.0, space_after=6.0)
    ppr = p._element.get_or_add_pPr()
    ind = _docx_set_child(ppr, _DOCX_PPR_ORDER, 'w:ind')
    ind.set(qn('w:left'), '240')

    name_lower = (item.get('name') or '').lower()
    label_lower = (q_label or '').lower()

    if base_type == 'text':
        is_long = any(k in name_lower for k in ['note', 'detail', 'comment', 'specify', 'other', 'explain', 'desc', 'reason']) or any(k in label_lower for k in ['explain', 'detail', 'specify', 'notes', 'comments', 'describe', 'reasons'])
        if is_long:
            _docx_style_run(p.add_run('____________________________________________________________________\n____________________________________________________________________'), size=DOCX_SMALL_PT, color='94A3B8')
        else:
            _docx_style_run(p.add_run('____________________________________________________________________'), size=DOCX_SMALL_PT, color='94A3B8')
    elif base_type in ('integer', 'decimal', 'range'):
        _docx_style_run(p.add_run('[ ____________________ ]'), size=DOCX_SMALL_PT, color='94A3B8')
    elif base_type == 'date':
        _docx_style_run(p.add_run('[ ____ / ____ / ________ ]  (DD/MM/YYYY)'), size=DOCX_SMALL_PT, color='94A3B8')
    elif base_type == 'time':
        _docx_style_run(p.add_run('[ ____ : ____ ]  (HH:MM)'), size=DOCX_SMALL_PT, color='94A3B8')
    elif base_type == 'datetime':
        _docx_style_run(p.add_run('[ ____/____/________   ____:____ ]'), size=DOCX_SMALL_PT, color='94A3B8')
    elif base_type in ('geopoint', 'geotrace', 'geoshape'):
        _docx_style_run(p.add_run('[ GPS Location Coordinates: ________________________________ ]'), size=DOCX_SMALL_PT, color='94A3B8', italic=True)
    elif base_type in ('image', 'audio', 'video', 'file'):
        _docx_style_run(p.add_run(f'[ Attach / Record {base_type.capitalize()}: ________________________________ ]'), size=DOCX_SMALL_PT, color='94A3B8', italic=True)
    elif base_type == 'barcode':
        _docx_style_run(p.add_run('[ Barcode / QR: ________________________________ ]'), size=DOCX_SMALL_PT, color='94A3B8', italic=True)
    elif base_type == 'acknowledge':
        _docx_style_run(p.add_run('[  ] Confirmed / Acknowledged'), size=DOCX_SMALL_PT, color='94A3B8')
    else:
        _docx_style_run(p.add_run('____________________________________________________________________'), size=DOCX_SMALL_PT, color='94A3B8')

def build_questionnaire_docx(parser, language='Default', landscape=False):
    """Build an editable, clean table-free DOCX questionnaire for interviewers/respondents."""
    if not DOCX_AVAILABLE:
        raise RuntimeError("python-docx is not installed. Run 'pip install python-docx' or pass --no-docx.")

    title_text = parser.settings.get('form_title') or parser.settings.get('title') or os.path.splitext(os.path.basename(parser.excel_path))[0]
    subtitle_text = f"Interview Questionnaire — Language: {language}" if language != 'Default' else "Interview Questionnaire"
    doc = _docx_setup_document(title_text, subtitle_text=subtitle_text, landscape=landscape)

    for item in parser.survey_rows:
        raw_type = item['type'].strip()
        base_type, list_name = split_type(raw_type)

        # Skip background system metadata types
        if base_type in SYSTEM_METADATA_TYPES:
            continue

        # Skip calculation rows and hidden logic
        if base_type in ('calculate', 'hidden') or raw_type == 'calculate':
            continue

        # Skip closing section tags
        if is_section_end(raw_type):
            continue

        # Section header check for group and repeat beginnings
        if is_section_start(raw_type):
            label = get_lang_value(item['labels'], language) or item['name']
            if not label and 'group' in raw_type.lower():
                label = item['name']
            if not label:
                continue
            _docx_add_section_header(doc, label)
            continue

        # Visible label check
        q_label = get_lang_value(item['labels'], language)
        if not q_label:
            continue

        # Note Rows: full-width styled callout block (no tables)
        if base_type == 'note':
            _docx_add_formatted_paragraphs(
                doc, q_label,
                base_size=DOCX_BODY_PT,
                base_color='334155',
                base_italic=False,
                space_before=4.0,
                space_after=4.0,
                left_indent_pt=14.0,
                right_indent_pt=8.0,
                shading=DOCX_FILL_NOTE,
                left_border=('24', DOCX_COLOR_NOTE_BORDER, '8')
            )
            continue

        # Question Prompt (no * required marker)
        q_hint = get_lang_value(item['hints'], language)
        has_choices = base_type.startswith(('select_one', 'select_multiple', 'select_or_other', 'rank')) and list_name

        _docx_add_formatted_paragraphs(
            doc, q_label,
            base_size=DOCX_QUESTION_PT,
            base_color='0F172A',
            base_bold=True,
            space_before=10.0,
            space_after=2.0 if q_hint else (2.0 if has_choices else 2.0),
            keep_with_next=True
        )

        # Question Hint
        if q_hint:
            hint_p = doc.add_paragraph()
            _docx_tight_paragraph(hint_p, space_before=0.0, space_after=3.0)
            hint_ppr = hint_p._element.get_or_add_pPr()
            _docx_toggle(hint_ppr, _DOCX_PPR_ORDER, 'w:keepNext')
            ind = _docx_set_child(hint_ppr, _DOCX_PPR_ORDER, 'w:ind')
            ind.set(qn('w:left'), '200')
            run = hint_p.add_run(f"Note: {q_hint}")
            _docx_style_run(run, size=DOCX_SMALL_PT, color=DOCX_COLOR_HINT, italic=True)

        # Choices (without option codes) or Open-Ended Blank
        if has_choices:
            choice_list = get_choice_list_by_name(parser.choices, list_name)
            if choice_list:
                is_multi = base_type.startswith('select_multiple')
                marker_str = '[ ]  ' if is_multi else '( )  '

                for idx, choice in enumerate(choice_list):
                    cp = doc.add_paragraph()
                    _docx_tight_paragraph(cp, space_before=1.0, space_after=1.5)
                    cppr = cp._element.get_or_add_pPr()
                    if idx < 2:  # Keep first choices close to question
                        _docx_toggle(cppr, _DOCX_PPR_ORDER, 'w:keepNext')
                    ind = _docx_set_child(cppr, _DOCX_PPR_ORDER, 'w:ind')
                    ind.set(qn('w:left'), '260')

                    m_run = cp.add_run(marker_str)
                    _docx_style_run(m_run, size=DOCX_BODY_PT, color='64748B')

                    choice_label = get_lang_value(choice['labels'], language) or choice['name']
                    l_run = cp.add_run(choice_label)
                    _docx_style_run(l_run, size=DOCX_BODY_PT, color='1E293B')
            else:
                bp = doc.add_paragraph()
                _docx_tight_paragraph(bp, space_before=2.0, space_after=4.0)
                ind = _docx_set_child(bp._element.get_or_add_pPr(), _DOCX_PPR_ORDER, 'w:ind')
                ind.set(qn('w:left'), '240')
                _docx_style_run(bp.add_run('[ Options from external file ]'), size=DOCX_SMALL_PT, color='94A3B8', italic=True)
        else:
            _docx_add_response_placeholder_flow(doc, base_type, item, q_label)

    return doc

# Alias for backwards compatibility
build_codebook_docx = build_questionnaire_docx

def generate_docx(parser, docx_path, language='Default', landscape=False):
    """Write the editable DOCX questionnaire for one language to disk."""
    doc = build_questionnaire_docx(parser, language=language, landscape=landscape)
    doc.save(docx_path)
    return docx_path

async def generate_pdf(html_path, pdf_path, landscape=False):
    """Render HTML questionnaire to A4 PDF via Playwright Chromium."""
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()
        await page.goto(f'file:///{os.path.abspath(html_path)}', wait_until='networkidle')
        await page.wait_for_timeout(1000)
        await page.pdf(
            path=pdf_path,
            format='A4',
            landscape=landscape,
            margin={'top': '15mm', 'right': '12mm', 'bottom': '15mm', 'left': '12mm'},
            print_background=True
        )
        await browser.close()

def process_xlsform(excel_path, target_lang=None, generate_pdf_flag=True, generate_docx_flag=True, landscape=False):
    """Process a single XLSForm spreadsheet and generate HTML, PDF, and DOCX questionnaires."""
    print(f"Processing XLSForm Questionnaire: {os.path.basename(excel_path)}...")
    parser = XLSFormParser(excel_path)

    if generate_docx_flag and not DOCX_AVAILABLE:
        print("  python-docx not installed, skipping DOCX output. Install it with: pip install python-docx")
        generate_docx_flag = False

    langs_to_process = parser.languages
    if target_lang:
        matching = [l for l in parser.languages if target_lang.lower() in l.lower()]
        if matching:
            langs_to_process = matching
        else:
            langs_to_process = [target_lang]

    base_name = os.path.splitext(excel_path)[0]

    for lang in langs_to_process:
        suffix = clean_language_suffix(lang)
        out_html_name = f"{base_name}_{suffix}.html" if suffix else f"{base_name}.html"
        out_pdf_name = f"{base_name}_{suffix}.pdf" if suffix else f"{base_name}.pdf"
        out_docx_name = f"{base_name}_{suffix}.docx" if suffix else f"{base_name}.docx"

        print(f"  Building HTML questionnaire for language: {lang} -> {os.path.basename(out_html_name)}")
        html_content = build_questionnaire_html(parser, language=lang, landscape=landscape)

        with open(out_html_name, 'w', encoding='utf-8') as f:
            f.write(html_content)
        print(f"    Saved HTML: {out_html_name}")

        if generate_pdf_flag:
            print(f"    Generating PDF: {out_pdf_name}...")
            asyncio.run(generate_pdf(out_html_name, out_pdf_name, landscape=landscape))
            print(f"    Saved PDF: {out_pdf_name}")

        if generate_docx_flag:
            print(f"    Generating DOCX: {out_docx_name}...")
            generate_docx(parser, out_docx_name, language=lang, landscape=landscape)
            print(f"    Saved DOCX: {out_docx_name}")

def main():
    parser = argparse.ArgumentParser(description="Direct XLSForm to Printable Interview Questionnaire (Word, HTML, PDF)")
    parser.add_argument('--input', '-i', help="Path to input XLSForm (.xlsx) file")
    parser.add_argument('--dir', '-d', help="Directory containing XLSForm files")
    parser.add_argument('--lang', '-l', help="Target language (e.g. English, Hindi). If omitted, processes all languages.")
    parser.add_argument('--landscape', action='store_true', help="Render in A4 landscape orientation instead of default portrait.")
    parser.add_argument('--no-pdf', action='store_true', help="Skip PDF generation.")
    parser.add_argument('--no-docx', action='store_true', help="Skip editable DOCX generation.")

    args = parser.parse_args()

    options = {
        'target_lang': args.lang,
        'generate_pdf_flag': not args.no_pdf,
        'generate_docx_flag': not args.no_docx,
        'landscape': args.landscape,
    }

    if not args.input and not args.dir:
        cwd = os.getcwd()
        xlsx_files = [os.path.join(cwd, f) for f in os.listdir(cwd) if f.endswith('.xlsx') and not f.startswith('~$')]
        if not xlsx_files:
            print("No .xlsx files found in current directory.")
            sys.exit(1)
        for f in xlsx_files:
            process_xlsform(f, **options)
    elif args.input:
        process_xlsform(args.input, **options)
    elif args.dir:
        xlsx_files = [os.path.join(args.dir, f) for f in os.listdir(args.dir) if f.endswith('.xlsx') and not f.startswith('~$')]
        for f in xlsx_files:
            process_xlsform(f, **options)

if __name__ == '__main__':
    main()
