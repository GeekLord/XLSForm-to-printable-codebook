"""
xlsform_to_questionnaire.py - Direct XLSForm to Printable Interview Questionnaire

Converts ODK / KoboToolbox / SurveyCTO XLSForm spreadsheets directly into clean,
field-ready, table-free interview questionnaires in Microsoft Word (.docx),
HTML, and PDF formats.

Key Architectural Characteristics:
----------------------------------
1. Human-Centric Questionnaire Flow:
   Unlike technical codebooks which present 3-column data dictionaries (Field Name,
   Question/Logic, Answer Choices/Codes), this script formats surveys as natural,
   conversational interview documents. Technical metadata, variable names,
   relevance constraints, calculation expressions, and raw numeric choice codes
   are omitted so field enumerators and respondents see only clean prompts, hints,
   and checkboxes.

2. Adaptive Response Blanks:
   Open-ended questions automatically receive formatted paper blanks matching their
   expected input type:
   - Short text: Single underline rule
   - Extended text / comments / reasons: Double underline rule
   - Numeric / integers: Bracketed box [ ________ ]
   - Dates / Times: Pre-formatted date/time guides [ __/__/____ ]
   - GPS / Media / Barcodes: Field protocol instruction callouts

3. Multi-Language Intelligence:
   Automatically inspects `survey` and `choices` sheets for localized headers
   (e.g., `label::English`, `label::Hindi`, `hint::English`), generating separate,
   fully localized HTML, PDF, and DOCX documents per language.

4. OpenXML Schema Precision & Complex Script Typography:
   For Microsoft Word (.docx) generation, low-level OXML manipulation ensures strict
   adherence to ECMA-376 schema child ordering (_DOCX_PPR_ORDER, _DOCX_RPR_ORDER)
   to prevent corrupted document warnings. High-coverage unicode typography
   (Arial Unicode MS) is bound across standard and complex-script XML attributes
   (w:ascii, w:hAnsi, w:cs, w:eastAsia, w:szCs, w:bCs, w:iCs) to guarantee crisp
   rendering of Indic (Devanagari, Odia, Bengali, Tamil), Arabic, and Latin scripts.

5. Page Break & Orphan Prevention:
   Section headers, question prompts, and initial choice items enforce `w:keepNext`
   (and CSS `page-break-after: avoid` / `page-break-inside: avoid`) to prevent
   orphaned headers and awkward page splits during printing.
"""

import os
import sys
import re
import argparse
import asyncio
from bs4 import BeautifulSoup
import openpyxl
from playwright.async_api import async_playwright

# Optional python-docx import: Enables direct generation of editable Word (.docx) documents.
# If python-docx is not installed in the environment, the script gracefully falls back
# to HTML and PDF generation without failing.
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


# ==============================================================================
# CSS & PRINT STYLES
# ==============================================================================
# Clean, modern typography and responsive container styling for questionnaire view.
# Includes both desktop screen reading styles and strict @media print rules
# tuned for A4 portrait/landscape paper pagination.
QUESTIONNAIRE_STYLES = """
*, *::before, *::after { box-sizing: border-box; }

/* Base typography: Arial Unicode MS ensures multi-script coverage (Indic, Latin, etc.) */
body {
    margin: 0;
    padding: 20px;
    font-family: "Arial Unicode MS", "Arial Unicode", "Segoe UI", Arial, sans-serif;
    font-size: 14px;
    line-height: 1.5;
    color: #111827;
    background: #f8fafc;
}

/* Centered reading container for screen view */
.questionnaire-container {
    width: 100%;
    max-width: 820px;
    margin: 0 auto;
    background: #fff;
    padding: 32px 40px;
    border-radius: 8px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
}

/* Questionnaire Document Header */
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

/* Group & Repeat Section Header Band */
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

/* Note / Instruction Callout Box */
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

/* Question Prompt & Layout */
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

/* Enumerator Instruction / Guidance Hint */
.hint {
    font-size: 12.5px;
    font-style: italic;
    color: #1e40af;
    margin: 2px 0 6px 12px;
    line-height: 1.4;
}

/* Multiple-choice option list */
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

/* Handwritten answer blank lines */
.answer-blank {
    margin: 6px 0 4px 14px;
    color: #64748b;
    font-size: 13px;
    line-height: 1.4;
}

/* Mobile responsive adjustments */
@media screen and (max-width: 768px) {
    body {
        padding: 10px;
    }
    .questionnaire-container {
        padding: 16px 18px;
    }
}

/* Print Stylesheet: Optimized for crisp physical paper pagination */
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


# ==============================================================================
# CONSTANTS & METADATA CONFIGURATION
# ==============================================================================
# Automated background metadata types injected by mobile data collection engines
# (ODK Collect, KoboToolbox, SurveyCTO). These are invisible background sensors
# (start/end timestamps, GPS location, device serial numbers, audit logs)
# and must be excluded from physical interview questionnaires.
SYSTEM_METADATA_TYPES = {
    'start', 'end', 'today', 'deviceid', 'phonenumber', 'simserial',
    'subscriberid', 'username', 'email', 'audit', 'text-audit',
    'start-geopoint', 'background-audio', 'caseid', 'caseread', 'casesave'
}


# ==============================================================================
# XLSFORM PARSING HELPERS
# ==============================================================================

def split_type(raw_type):
    """
    Split an XLSForm `type` cell into its base type and trailing list name or argument.

    Example:
        'select_one gender_list' -> ('select_one', 'gender_list')
        'text'                   -> ('text', '')
        'begin group'            -> ('begin group', '')

    Args:
        raw_type (str): Raw string value from the 'type' column of the survey sheet.

    Returns:
        tuple[str, str]: (base_type, list_name_or_arg)
    """
    if not raw_type:
        return '', ''
    parts = raw_type.strip().split(None, 1)
    base = parts[0].strip()
    arg = parts[1].strip() if len(parts) > 1 else ''
    return base, arg


def is_section_start(raw_type):
    """
    Determine if a row represents the opening of a question group or repeat block.

    Supports various authoring conventions ('begin group', 'begin_group',
    'begin repeat', 'begin_repeat').

    Args:
        raw_type (str): The raw XLSForm question type.

    Returns:
        bool: True if row opens a group/repeat container.
    """
    t = raw_type.lower().strip()
    return t in ['begin group', 'begin_group', 'begin repeat', 'begin_repeat'] or t.startswith(('begin group', 'begin_group', 'begin repeat', 'begin_repeat'))


def is_section_end(raw_type):
    """
    Determine if a row represents the closing tag of a question group or repeat block.

    Supports 'end group', 'end_group', 'end repeat', 'end_repeat'.

    Args:
        raw_type (str): The raw XLSForm question type.

    Returns:
        bool: True if row closes a group/repeat container.
    """
    t = raw_type.lower().strip()
    return t in ['end group', 'end_group', 'end repeat', 'end_repeat'] or t.startswith(('end group', 'end_group', 'end repeat', 'end_repeat'))


def extract_language_name(header_str):
    """
    Extract language name from an XLSForm multi-language header string.

    Examples:
        'label::English (en)' -> 'English (en)'
        'hint::Hindi (hi)'    -> 'Hindi (hi)'
        'label'               -> None

    Args:
        header_str (str): Raw Excel column header text.

    Returns:
        str | None: Extracted language name, or None if not a localized header.
    """
    match = re.match(r'^(?:label|hint|constraint_message)::(.+)$', header_str.strip())
    if match:
        return match.group(1).strip()
    return None


def clean_language_suffix(lang_name):
    """
    Sanitize language name for use in output file suffixes.

    Strips parenthetical locale codes and unsafe filesystem characters.
    Example: 'English (en)' -> 'English', 'Français (fr)' -> 'Français'

    Args:
        lang_name (str): Raw language identifier.

    Returns:
        str: Cleaned filesystem-safe language string.
    """
    clean = re.sub(r'\(.*?\)', '', lang_name).strip()
    clean = re.sub(r'[^a-zA-Z0-9_-]', '_', clean)
    clean = re.sub(r'_+', '_', clean).strip('_')
    return clean


def get_header_index(header_map, *keys):
    """
    Look up 0-based column index in header_map matching one of the target keys case-insensitively.

    Using this helper prevents common Python 0-as-falsy bugs (e.g. `if col_idx:` failing on index 0).

    Args:
        header_map (dict[str, int]): Mapping of column header name to its 0-based index.
        *keys (str): Candidate column names to test (e.g. 'list_name', 'list name').

    Returns:
        int | None: Matching column index, or None if none of the keys were found.
    """
    for key in keys:
        target = key.strip().lower()
        for h, idx in header_map.items():
            if h.strip().lower() == target:
                return idx
    return None


def get_lang_value(dict_of_langs, target_lang):
    """
    Retrieve localized text string from language dictionary with intelligent fallbacks.

    Fallback Resolution Sequence:
    1. Exact key match (e.g. 'English (en)')
    2. Case-insensitive substring match (e.g. 'english' in 'English (en)')
    3. 'Default' language entry
    4. First non-empty value in the dictionary
    5. Empty string '' if no text found

    Args:
        dict_of_langs (dict[str, str]): Dictionary mapping language names to translated strings.
        target_lang (str): Desired target language.

    Returns:
        str: Resolved text string.
    """
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

# ==============================================================================
# EXCEL WORKBOOK PARSER (XLSFormParser)
# ==============================================================================

class XLSFormParser:
    """
    Parses an ODK/KoboToolbox/SurveyCTO XLSForm workbook into structured Python dictionaries.

    Responsible for:
    - Normalizing sheet names case-insensitively ('survey', 'choices', 'settings').
    - Parsing form-level metadata from the 'settings' sheet (e.g., form_title, id_string).
    - Discovering all language translations defined across 'label::Language' headers.
    - Extracting multiple-choice options and their localized labels from 'choices'.
    - Sequential extraction of survey rows from 'survey' while capturing localized labels & hints.

    Attributes:
        excel_path (str): Filesystem path to the .xlsx workbook.
        wb (openpyxl.Workbook): Openpyxl workbook loaded in data_only mode (evaluates cached values).
        sheets (dict[str, str]): Lowercase-to-actual sheet name lookup mapping.
        settings (dict[str, str]): Key-value settings parsed from 'settings' sheet.
        languages (list[str]): List of detected language names (e.g. ['English (en)', 'Hindi (hi)']).
        choices (dict[str, list[dict]]): Parsed choice options grouped by list_name.
        survey_rows (list[dict]): Ordered survey rows parsed from 'survey' sheet.
    """

    def __init__(self, excel_path):
        """
        Initialize workbook and execute full parse sequence.

        Args:
            excel_path (str): Path to input XLSForm .xlsx file.
        """
        self.excel_path = excel_path
        # data_only=True ensures openpyxl retrieves cached calculation values instead of formula strings
        self.wb = openpyxl.load_workbook(excel_path, data_only=True)
        # Build lowercase lookup dictionary to tolerate casing differences across authoring tools
        self.sheets = {name.lower().strip(): name for name in self.wb.sheetnames}
        
        self.settings = self._parse_settings()
        self.languages = self._detect_languages()
        self.choices = self._parse_choices()
        self.survey_rows = self._parse_survey()

    def _get_sheet(self, target_name):
        """
        Retrieve openpyxl worksheet by name case-insensitively.

        Args:
            target_name (str): Expected sheet name (e.g. 'survey', 'choices', 'settings').

        Returns:
            openpyxl.worksheet.worksheet.Worksheet | None: Worksheet instance or None if not found.
        """
        actual_name = self.sheets.get(target_name.lower())
        if actual_name:
            return self.wb[actual_name]
        return None

    def _parse_settings(self):
        """
        Extract form-level configuration from the 'settings' sheet.

        Captures properties like 'form_title', 'form_id', 'version', and 'default_language'.

        Returns:
            dict[str, str]: Key-value configuration dictionary.
        """
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
        """
        Scan column headers in 'survey' and 'choices' sheets to identify all defined languages.

        Detects patterns such as 'label::English (en)' and 'label::Hindi'. If no localized
        headers exist, defaults to ['Default'].

        Returns:
            list[str]: Unique list of detected language names.
        """
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
        """
        Parse the 'choices' sheet to construct a dictionary of option lists.

        Maps each `list_name` to an array of choice objects containing:
        - `name`: The raw stored value/code.
        - `labels`: Dictionary mapping each language to its human-readable choice text.

        Returns:
            dict[str, list[dict]]: Structured choice lists keyed by list_name.
        """
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
        """
        Parse rows sequentially from the 'survey' sheet.

        Extracts core question attributes:
        - `type`: Question type string (e.g. 'select_one q1_list', 'text', 'begin group').
        - `name`: Variable field name.
        - `labels`: Localized question prompts across all detected languages.
        - `hints`: Localized enumerator guidance notes across all detected languages.

        Returns:
            list[dict]: Ordered collection of survey question dictionaries.
        """
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
                # Resolve localized question label
                col_idx = get_header_index(header_map, f'label::{lang}')
                if col_idx is None and lang != 'Default':
                    for h, idx in header_map.items():
                        if h.startswith('label::') and lang.lower() in h.lower():
                            col_idx = idx
                            break
                if col_idx is None:
                    col_idx = get_header_index(header_map, 'label')
                labels[lang] = str(row_cells[col_idx].value or '').strip() if (col_idx is not None and col_idx < len(row_cells) and row_cells[col_idx].value is not None) else ''

                # Resolve localized enumerator hint
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
    """
    Look up a choice list from the choices dictionary with case-insensitive tolerance.

    Args:
        choices_dict (dict[str, list[dict]]): Parsed choice options keyed by list_name.
        list_name (str): Target choice list name from the survey type cell.

    Returns:
        list[dict]: Array of choice items, or empty list if not found.
    """
    if not list_name or not choices_dict:
        return []
    if list_name in choices_dict:
        return choices_dict[list_name]
    for key, items in choices_dict.items():
        if key.strip().lower() == list_name.strip().lower():
            return items
    return []


# ==============================================================================
# RESPONSE BLANKS & HTML QUESTIONNAIRE BUILDER
# ==============================================================================

def get_html_response_placeholder(base_type, item, q_label):
    """
    Generate an HTML response blank formatted specifically for paper-based handwriting.

    Adaptive Heuristics:
    - Text questions containing keywords like 'explain', 'detail', 'comment', 'reason'
      automatically receive a double-ruled underline blank to accommodate longer written responses.
    - Standard text receives a single underline rule.
    - Numbers/integers receive boxed brackets `[ ____________________ ]`.
    - Dates and times receive segmented guides `[ ____ / ____ / ________ ] (DD/MM/YYYY)`.
    - Hardware/sensor fields (GPS, audio, photos, barcodes) receive formatted protocol callouts.

    Args:
        base_type (str): Base question type (e.g. 'text', 'integer', 'date').
        item (dict): Question row dictionary containing 'name' and other attributes.
        q_label (str): Human-readable question prompt text.

    Returns:
        str: Raw HTML snippet representing the printable answer blank.
    """
    name_lower = (item.get('name') or '').lower()
    label_lower = (q_label or '').lower()
    
    if base_type == 'text':
        # Detect whether question warrants extended writing space
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
    """
    Construct a responsive, table-free HTML questionnaire document for field interviewers.

    Structural Pipeline:
    1. Generates complete HTML5 boilerplate with meta viewport and UTF-8 charset.
    2. Injects responsive screen and @media print CSS rules.
    3. Traverses survey rows sequentially:
       - Skips automated sensor metadata (start, deviceid, audit, etc.).
       - Skips internal calculation / hidden variables (calculate, hidden).
       - Skips closing group/repeat tags.
       - Translates 'begin group' / 'begin repeat' into styled section banner bands.
       - Formats 'note' items as prominent instructional callouts.
       - Formats questions with bold prompts, italicized guidance hints, and clean options:
         * Select questions: '( )' for single-choice, '[ ]' for multiple-choice (codes hidden).
         * Open questions: Formatted response blanks via `get_html_response_placeholder`.

    Args:
        parser (XLSFormParser): Parsed XLSForm workbook instance.
        language (str): Target language to extract and render.
        landscape (bool): If True, adjusts container width and print page rule to landscape A4.

    Returns:
        str: Serialized HTML document string.
    """
    soup = BeautifulSoup('<!DOCTYPE html><html><head></head><body></body></html>', 'html.parser')
    
    head = soup.head
    head.append(soup.new_tag('meta', attrs={'charset': 'utf-8'}))
    head.append(soup.new_tag('meta', attrs={'name': 'viewport', 'content': 'width=device-width, initial-scale=1.0'}))
    
    # Form Title resolution from settings or filename
    title_text = parser.settings.get('form_title') or parser.settings.get('title') or os.path.splitext(os.path.basename(parser.excel_path))[0]
    title_tag = soup.new_tag('title')
    title_tag.string = f"{title_text} - Questionnaire ({language})"
    head.append(title_tag)
    
    # Invert dimensions if landscape mode is requested
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

        # 1. Skip background system metadata types
        if base_type in SYSTEM_METADATA_TYPES:
            continue

        # 2. Skip calculation rows and hidden logic
        if base_type in ('calculate', 'hidden') or raw_type == 'calculate':
            continue

        # 3. Skip closing section tags
        if is_section_end(raw_type):
            continue

        # 4. Section header check for group and repeat beginnings
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

        # 5. Visible label check: suppress questions without a label in this language
        q_label = get_lang_value(item['labels'], language)
        if not q_label:
            continue

        # 6. Note Rows: render as full-width script / instruction block
        if base_type == 'note':
            note_item = soup.new_tag('div', attrs={'class': 'note-item'})
            note_div = soup.new_tag('div', attrs={'class': 'note-text'})
            note_div.append(BeautifulSoup(q_label, 'html.parser'))
            note_item.append(note_div)
            container.append(note_item)
            continue

        # 7. Standard Interview Question (clean flowing layout, no tables)
        q_item = soup.new_tag('div', attrs={'class': 'question-item'})

        # Question Prompt (clean bold text)
        div_prompt = soup.new_tag('div', attrs={'class': 'question-prompt'})
        div_prompt.append(BeautifulSoup(q_label, 'html.parser'))
        q_item.append(div_prompt)

        # Guidance Hint (italicized notes)
        q_hint = get_lang_value(item['hints'], language)
        if q_hint:
            div_hint = soup.new_tag('div', attrs={'class': 'hint'})
            div_hint.string = f"Note: {q_hint}"
            q_item.append(div_hint)

        # Multiple-Choice Options OR Open-Ended Response Blank (Codes stripped)
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

# Alias for backwards compatibility with earlier API references
build_codebook_html = build_questionnaire_html

# ==============================================================================
# OPENXML / WORD (.DOCX) CONSTANTS & SCHEMA ENGINE
# ==============================================================================

# Typography: Arial Unicode MS provides comprehensive glyph coverage across
# Devanagari, Bengali, Odia, Tamil, Telugu, Arabic, and Latin alphabets,
# preventing missing-glyph 'tofu' boxes in multilingual survey instruments.
DOCX_FONT = 'Arial Unicode MS'
DOCX_MONO_FONT = 'Consolas'

# Font sizing scale (in points)
DOCX_TITLE_PT = 16.0
DOCX_HEADER_PT = 11.0
DOCX_SECTION_PT = 12.0
DOCX_QUESTION_PT = 10.5
DOCX_BODY_PT = 9.5
DOCX_SMALL_PT = 8.5
DOCX_TINY_PT = 7.5

# Hex Color Palette (without '#' prefix for OpenXML hex strings)
DOCX_COLOR_TEXT = '0F172A'         # Deep Slate for high-contrast legible text
DOCX_COLOR_HINT = '1E40AF'         # Rich Blue for enumerator guidance notes
DOCX_COLOR_NOTE_BORDER = '3B82F6'  # Accent Blue for note callout left border

DOCX_FILL_SECTION = 'F0F4F8'       # Soft grey-blue tint for section header band
DOCX_FILL_NOTE = 'F8FAFC'          # Very light tint for instruction callouts

DOCX_RULE_SECTION = '334155'       # Dark slate border rule for sections
DOCX_RULE_LIGHT = 'E2E8F0'         # Subtle divider rule between questions

# Standard Page Setup: Dimensions for Portrait A4 (ISO 216) in millimeters
DOCX_PORTRAIT_WIDTH_MM = 210.0
DOCX_PORTRAIT_HEIGHT_MM = 297.0
DOCX_MARGIN_LEFT_RIGHT_MM = 16.0
DOCX_MARGIN_TOP_BOTTOM_MM = 15.0

# Dimensions for Landscape A4
DOCX_LANDSCAPE_WIDTH_MM = 297.0
DOCX_LANDSCAPE_HEIGHT_MM = 210.0

# ------------------------------------------------------------------------------
# OpenXML ECMA-376 Strict Schema Ordering Definitions
# ------------------------------------------------------------------------------
# In Microsoft Word's OpenXML schema (ISO/IEC 29500-1:2016), child elements of
# paragraph properties (<w:pPr>) and run properties (<w:rPr>) must appear in a
# strictly defined canonical order. If elements are inserted out of sequence
# (e.g. inserting <w:shd> before <w:pBdr>), Microsoft Word will flag the document
# as damaged and display an "unreadable content - repair" warning.
# The tuples below define the precise schema sequence required for validation.

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


# ==============================================================================
# OPENXML LOW-LEVEL PROPERTY HELPERS
# ==============================================================================

def _docx_set_child(parent, order, tag, element=None, replace=True):
    """
    Insert or replace a child OXML element in its strictly schema-mandated position.

    Finds the first existing successor element defined in `order` and inserts the new
    element immediately before it, preventing schema validation violations in Word.

    Args:
        parent (docx.oxml.xmlchemy.BaseOxmlElement): Parent XML element (e.g. pPr or rPr).
        order (tuple[str]): Canonical schema order tuple (_DOCX_PPR_ORDER or _DOCX_RPR_ORDER).
        tag (str): Prefixed tag name (e.g. 'w:keepNext', 'w:shd').
        element (OxmlElement | None): Existing element to insert, or None to create fresh.
        replace (bool): If True, replaces any existing element with the same tag.

    Returns:
        OxmlElement: The inserted child XML element.
    """
    existing = parent.find(qn(tag))
    if existing is not None:
        if not replace:
            return existing
        parent.remove(existing)
    if element is None:
        element = OxmlElement(tag)
    # Find all elements that are legally allowed to follow this tag
    successors = order[order.index(tag) + 1:]
    parent.insert_element_before(element, *successors)
    return element


def _docx_toggle(parent, order, tag, value='true'):
    """
    Insert a boolean toggle element (e.g. <w:keepNext w:val="true"/> or <w:bCs/>).

    Args:
        parent: Parent XML container element.
        order (tuple[str]): Canonical order tuple.
        tag (str): Prefixed XML tag name.
        value (str): Value attribute string (default: 'true').

    Returns:
        OxmlElement: Configured toggle element.
    """
    el = _docx_set_child(parent, order, tag)
    el.set(qn('w:val'), value)
    return el


def _docx_border(tag, sz='4', color='auto', val='single', space='0'):
    """
    Construct a single OpenXML border descriptor element (<w:left>, <w:top>, etc.).

    Args:
        tag (str): Border edge name ('top', 'left', 'bottom', 'right').
        sz (str): Border width in 1/8th of a point (e.g. '4' = 0.5pt, '12' = 1.5pt).
        color (str): Hex RGB color string without '#' prefix, or 'auto'.
        val (str): Border style ('single', 'dotted', 'dashed', 'double').
        space (str): Spacing distance from text in points.

    Returns:
        OxmlElement: Configured border XML node.
    """
    el = OxmlElement(f'w:{tag}')
    el.set(qn('w:val'), val)
    el.set(qn('w:sz'), sz)
    el.set(qn('w:space'), space)
    el.set(qn('w:color'), color)
    return el


def _docx_shading(fill):
    """
    Construct an OpenXML shading element (<w:shd>) for background color fill.

    Args:
        fill (str): Hex RGB color code without '#' prefix.

    Returns:
        OxmlElement: Shading XML node.
    """
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill)
    return shd


def _docx_style_run(run, size=DOCX_BODY_PT, color=None, bold=False, italic=False, mono=False):
    """
    Apply comprehensive font typography to a docx text run, ensuring Indic script fidelity.

    Complex-Script Note:
    Microsoft Word maintains separate font attributes for Latin ('w:ascii', 'w:hAnsi')
    and Complex Scripts ('w:cs', 'w:szCs', 'w:bCs', 'w:iCs'). For Indic languages
    (Devanagari, Odia, Bengali, Tamil), setting standard `font.bold = True` does not bold
    the characters unless `w:bCs` is explicitly toggled on the run properties.
    Similarly, font size in `w:szCs` must be supplied in half-points (points * 2).

    Args:
        run (docx.text.run.Run): The run instance to format.
        size (float): Font size in typographical points.
        color (str | None): Hex RGB color code (e.g. '0F172A').
        bold (bool): Whether text should be bold.
        italic (bool): Whether text should be italicized.
        mono (bool): Whether to use fixed-width monospace font (Consolas).

    Returns:
        docx.text.run.Run: The styled run instance.
    """
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
    # Map font name across all text categories (Western, High-ANSI, Complex Script, East Asian)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rfonts.set(qn(attr), name)

    # Complex script font size: specified in half-points (size * 2)
    _docx_toggle(rpr, _DOCX_RPR_ORDER, 'w:szCs', str(int(round(size * 2))))
    if bold:
        _docx_toggle(rpr, _DOCX_RPR_ORDER, 'w:bCs')
    if italic:
        _docx_toggle(rpr, _DOCX_RPR_ORDER, 'w:iCs')
    return run


def _docx_tight_paragraph(paragraph, space_before=0.0, space_after=1.0):
    """
    Set compact paragraph spacing to ensure survey items stay visually unified.

    Args:
        paragraph (docx.text.paragraph.Paragraph): Target paragraph.
        space_before (float): Space before paragraph in points.
        space_after (float): Space after paragraph in points.
    """
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(space_before)
    fmt.space_after = Pt(space_after)
    fmt.line_spacing = 1.1


def _docx_setup_document(title_text, subtitle_text=None, landscape=False):
    """
    Initialize a blank Word document with A4 page setup, tight margins, and title header.

    Args:
        title_text (str): Primary survey title string.
        subtitle_text (str | None): Subtitle string detailing language / purpose.
        landscape (bool): If True, configures A4 landscape orientation; otherwise portrait.

    Returns:
        docx.Document: Configured Word document instance.
    """
    doc = Document()

    # Set base document 'Normal' style defaults
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

    # Apply A4 geometry and clean print margins
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

    # Document Header Title
    heading = doc.add_paragraph()
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _docx_tight_paragraph(heading, space_before=0.0, space_after=2.0)
    _docx_style_run(heading.add_run(title_text), size=DOCX_TITLE_PT, bold=True, color='0F172A')

    # Subtitle with underline divider
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
    """
    Parse an HTML or plain text string into one or more styled Word paragraphs.

    Walks the HTML DOM recursively, handling inline formatting tags:
    - Bold: `<b>`, `<strong>`
    - Italic: `<i>`, `<em>`
    - Line Breaks: `<br>`
    - Paragraph blocks: `<p>`, `<div>`

    Applies paragraph-level properties strictly adhering to OpenXML schema sequence:
    - Indentations (`w:ind` in twips, 1 pt = 20 twips)
    - Shading (`w:shd`)
    - Custom borders (`w:pBdr`)
    - Pagination keep-with-next (`w:keepNext`)

    Args:
        doc (docx.Document): Active Word document.
        text_or_html (str): Text or HTML fragment to render.
        base_size (float): Base font size in points.
        base_color (str | None): Hex text color code.
        base_bold (bool): Base bold styling.
        base_italic (bool): Base italic styling.
        space_before (float): Space before paragraph in points.
        space_after (float): Space after paragraph in points.
        left_indent_pt (float): Left indent in points.
        right_indent_pt (float): Right indent in points.
        keep_with_next (bool): If True, attaches w:keepNext to prevent orphan across page.
        shading (str | None): Background fill hex color string.
        left_border (tuple | None): (sz, color, space) tuple for left border rule.
        bottom_border (tuple | None): (sz, color, space) tuple for bottom border rule.
        top_border (tuple | None): (sz, color, space) tuple for top border rule.

    Returns:
        list[docx.text.paragraph.Paragraph]: List of created Paragraph instances.
    """
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

        # Indentations in OpenXML are measured in twips (1/20th of a point)
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

        # Recursive DOM walker for inline formatting
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
    """
    Add a styled group/repeat section header band to the Word document.

    Applies:
    - Light grey-blue background tint (#F0F4F8).
    - Top and bottom border rules (#334155).
    - `w:keepNext` to prevent this section header from being orphaned at the bottom of a page.

    Args:
        doc (docx.Document): Active Word document.
        label (str): Section header title text.

    Returns:
        docx.text.paragraph.Paragraph: Created section header paragraph.
    """
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
    """
    Insert a table-free, styled response blank into the Word document for open-ended questions.

    Adaptive Heuristics:
    - Text questions containing keywords ('detail', 'explain', 'comment', 'reason', etc.)
      receive a double-ruled underline blank for extended writing space.
    - Standard text questions receive a single underline rule.
    - Numeric fields receive bracketed placeholders `[ ____________________ ]`.
    - Date/Time questions receive structured entry slots.
    - Sensor fields receive italicized guidance labels.

    Args:
        doc (docx.Document): Active Word document.
        base_type (str): Base question type (e.g. 'text', 'integer', 'date').
        item (dict): Question row dictionary.
        q_label (str): Human-readable question prompt text.
    """
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


# ==============================================================================
# WORD DOCUMENT BUILDER & QUESTION FLOW
# ==============================================================================

def build_questionnaire_docx(parser, language='Default', landscape=False):
    """
    Build an editable, clean, table-free Microsoft Word (.docx) questionnaire document.

    Orchestration Flow:
    1. Validates `python-docx` availability.
    2. Initializes document with A4 page layout and title banner.
    3. Traverses survey rows sequentially:
       - Suppresses automated sensor metadata (`start`, `deviceid`, `audit`, etc.).
       - Suppresses internal calculation variables (`calculate`, `hidden`).
       - Translates group/repeat starts into styled section header bands.
       - Renders 'note' items as full-width callout boxes with blue left accent border.
       - Renders questions with bold prompts, guidance hints, and choice options.
       - Applies `w:keepNext` on question prompts, hints, and top choices to ensure
         questions do not break awkwardly across page splits.

    Args:
        parser (XLSFormParser): Parsed XLSForm workbook instance.
        language (str): Target language to extract and render.
        landscape (bool): If True, formats page geometry as landscape A4.

    Returns:
        docx.Document: Fully populated Word document instance.

    Raises:
        RuntimeError: If `python-docx` is not installed in the Python environment.
    """
    if not DOCX_AVAILABLE:
        raise RuntimeError("python-docx is not installed. Run 'pip install python-docx' or pass --no-docx.")

    title_text = parser.settings.get('form_title') or parser.settings.get('title') or os.path.splitext(os.path.basename(parser.excel_path))[0]
    subtitle_text = f"Interview Questionnaire — Language: {language}" if language != 'Default' else "Interview Questionnaire"
    doc = _docx_setup_document(title_text, subtitle_text=subtitle_text, landscape=landscape)

    for item in parser.survey_rows:
        raw_type = item['type'].strip()
        base_type, list_name = split_type(raw_type)

        # 1. Skip background system metadata types
        if base_type in SYSTEM_METADATA_TYPES:
            continue

        # 2. Skip calculation rows and hidden logic
        if base_type in ('calculate', 'hidden') or raw_type == 'calculate':
            continue

        # 3. Skip closing section tags
        if is_section_end(raw_type):
            continue

        # 4. Section header check for group and repeat beginnings
        if is_section_start(raw_type):
            label = get_lang_value(item['labels'], language) or item['name']
            if not label and 'group' in raw_type.lower():
                label = item['name']
            if not label:
                continue
            _docx_add_section_header(doc, label)
            continue

        # 5. Visible label check
        q_label = get_lang_value(item['labels'], language)
        if not q_label:
            continue

        # 6. Note Rows: full-width styled callout block (no tables)
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

        # 7. Question Prompt (bold question text with keep-with-next pagination protection)
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

        # 8. Guidance Hint (italicized enumerator note)
        if q_hint:
            hint_p = doc.add_paragraph()
            _docx_tight_paragraph(hint_p, space_before=0.0, space_after=3.0)
            hint_ppr = hint_p._element.get_or_add_pPr()
            _docx_toggle(hint_ppr, _DOCX_PPR_ORDER, 'w:keepNext')
            ind = _docx_set_child(hint_ppr, _DOCX_PPR_ORDER, 'w:ind')
            ind.set(qn('w:left'), '200')
            run = hint_p.add_run(f"Note: {q_hint}")
            _docx_style_run(run, size=DOCX_SMALL_PT, color=DOCX_COLOR_HINT, italic=True)

        # 9. Multiple-Choice Options OR Open-Ended Blank (Option codes stripped)
        if has_choices:
            choice_list = get_choice_list_by_name(parser.choices, list_name)
            if choice_list:
                is_multi = base_type.startswith('select_multiple')
                marker_str = '[ ]  ' if is_multi else '( )  '

                for idx, choice in enumerate(choice_list):
                    cp = doc.add_paragraph()
                    _docx_tight_paragraph(cp, space_before=1.0, space_after=1.5)
                    cppr = cp._element.get_or_add_pPr()
                    if idx < 2:  # Keep first two choices close to the question prompt
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

# Alias for backwards compatibility with earlier API references
build_codebook_docx = build_questionnaire_docx


def generate_docx(parser, docx_path, language='Default', landscape=False):
    """
    Generate and serialize the editable DOCX questionnaire for a given language to disk.

    Args:
        parser (XLSFormParser): Parsed workbook instance.
        docx_path (str): Destination file path.
        language (str): Target language string.
        landscape (bool): Landscape orientation toggle.

    Returns:
        str: Absolute or relative output file path.
    """
    doc = build_questionnaire_docx(parser, language=language, landscape=landscape)
    doc.save(docx_path)
    return docx_path


# ==============================================================================
# PDF RENDERING ENGINE (PLAYWRIGHT CHROMIUM)
# ==============================================================================

async def generate_pdf(html_path, pdf_path, landscape=False):
    """
    Render an HTML questionnaire to pixel-perfect A4 PDF using Playwright Chromium.

    Enables `print_background=True` so background shading tints and border rules
    on section banners and note callouts are rendered faithfully.

    Args:
        html_path (str): Path to source HTML file.
        pdf_path (str): Destination PDF file path.
        landscape (bool): If True, renders A4 landscape; otherwise portrait.
    """
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()
        await page.goto(f'file:///{os.path.abspath(html_path)}', wait_until='networkidle')
        # Allow web fonts and typography engines to complete layout calculation
        await page.wait_for_timeout(1000)
        await page.pdf(
            path=pdf_path,
            format='A4',
            landscape=landscape,
            margin={'top': '15mm', 'right': '12mm', 'bottom': '15mm', 'left': '12mm'},
            print_background=True
        )
        await browser.close()


# ==============================================================================
# PIPELINE ORCHESTRATION & CLI ENTRY POINT
# ==============================================================================

def process_xlsform(excel_path, target_lang=None, generate_pdf_flag=True, generate_docx_flag=True, landscape=False):
    """
    Process a single XLSForm workbook and generate localized questionnaires across formats.

    For each target language:
    1. Builds and writes localized HTML file.
    2. Renders PDF via Playwright Chromium (if enabled).
    3. Builds and writes editable Word (.docx) file (if enabled).

    Args:
        excel_path (str): Path to input XLSForm .xlsx file.
        target_lang (str | None): Specific language to filter (e.g. 'English'). If None, processes all.
        generate_pdf_flag (bool): Flag to toggle PDF output.
        generate_docx_flag (bool): Flag to toggle Word (.docx) output.
        landscape (bool): Flag to toggle landscape A4 orientation.
    """
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
    """
    Parse command-line arguments and execute questionnaire generation pipeline.
    """
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
