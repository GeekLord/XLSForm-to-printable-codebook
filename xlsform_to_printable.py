"""
xlsform_to_printable.py - XLSForm to Printable 3-Column Codebook Generator

Converts ODK / KoboToolbox / SurveyCTO XLSForm spreadsheets into professional,
printable 3-column data dictionaries and technical codebooks in HTML, PDF,
and editable Microsoft Word (.docx) formats.

Architectural Paradigm:
-----------------------
1. 3-Column Data Dictionary Layout:
   - Column 1: Field Name (22% width): Stored variable name, highlighted in bold,
     with explicit '(required)' tags for mandatory responses.
   - Column 2: Question & Logic (50% width): Human-readable question prompt,
     enumerator guidance hints, skip logic / relevance conditions, automated
     platform metadata descriptors, and categorized calculation formulas.
   - Column 3: Answer Options (28% width): Two-column nested choice table displaying
     the raw stored choice value alongside its localized label text.

2. Automated Calculation Taxonomy:
   Inspects calculation expressions and categorizes them with contextual badges
   and human-readable descriptions (e.g., CSV lookups via `pulldata()`, choice
   label lookups via `jr:choice-name()`, timestamps, string processing, and conditional logic).
   Long expressions are trimmed for clean visual presentation while retaining full
   code tooltips.

3. Platform Metadata Annotations:
   Recognizes background metadata fields (start, end, today, deviceid, audit,
   simserial, subscriberid, start-geopoint) and automatically attaches explanatory
   callouts explaining their data collection role.

4. Multi-Language Codebook Generation:
   Automatically detects localized `label::[Lang]` and `hint::[Lang]` columns,
   generating separate, synchronized HTML, PDF, and DOCX codebooks per language.

5. OpenXML Schema Precision & Complex Script Typography:
   Generates native Microsoft Word (.docx) documents in landscape A4 orientation.
   Utilizes strict ECMA-376 schema ordering to prevent file repair warnings.
   Binds high-coverage unicode typography (Arial Unicode MS) across standard and 
   complex-script attributes ensuring flawless rendering for Indic and Arabic scripts.
"""

import os
import sys
import re
import argparse
import asyncio
import base64
import mimetypes
from bs4 import BeautifulSoup
import openpyxl
from playwright.async_api import async_playwright

# Optional PIL / Pillow import for proportional image sizing
try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# Optional python-docx import: Enables direct generation of editable Word (.docx) documents.
# If python-docx is not installed, gracefully falls back to HTML and PDF generation.
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
# CSS & PRINT STYLES (3-COLUMN FIXED TABLE LAYOUT)
# ==============================================================================
# Responsive and print-optimized styles engineered for technical survey codebooks.
# Features a rigid 3-column fixed table layout (22% / 50% / 28%), dark section bands,
# high-contrast metadata badges, and @media print pagination controls.
RESPONSIVE_STYLES = """
.spacer {
    display: none !important;
    width: 0 !important;
    padding: 0 !important;
    margin: 0 !important;
}

*, *::before, *::after { box-sizing: border-box; }

body {
    margin: 0;
    padding: 10px;
    font-family: "Arial Unicode MS", "Arial Unicode", Arial, sans-serif;
    font-size: 14px;
    line-height: 20px;
    color: #000;
    background: #fff;
}

.codebook-container {
    width: 100%;
    max-width: 1200px;
    margin: auto;
    padding: 0 10px;
}

h4 {
    text-align: center;
    font-size: 19px;
    margin: 12px 0;
    padding-bottom: 8px;
    border-bottom: 2px solid #333;
}

table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 20px;
    table-layout: fixed;
}

th, td {
    padding: 7px 9px;
    text-align: left;
    border: 1px solid #ddd;
    vertical-align: top;
    overflow-wrap: break-word;
    word-break: break-word;
}

td[colspan="3"],
.section-header-dark {
    background: #2b2d42 !important;
    color: #fff !important;
    font-weight: bold;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
}

.headerCell {
    background: #4a4e69;
    color: #fff;
}

.headerCell h6 {
    margin: 0;
    font-size: 13px;
    font-weight: bold;
    color: #fff;
}

.fieldCell {
    font-weight: bold;
    color: #222;
    width: 22%;
    font-size: 13.5px;
    white-space: normal;
    overflow-wrap: break-word;
    word-break: break-word;
}

.fieldCell .required {
    color: #c00;
    font-weight: normal;
}

.questionCell {
    width: 50%;
    font-size: 13.5px;
    overflow-wrap: break-word;
    word-break: break-word;
}

.answerCell {
    width: 28%;
    font-size: 13px;
    overflow-wrap: break-word;
    word-break: break-word;
}

tr.gray td {
    background: #fff;
}

td .table {
    margin: 0;
    font-size: 12.5px;
    width: 100%;
    table-layout: fixed;
}

td .table td {
    padding: 2px 4px;
    border: none;
    border-bottom: 1px dotted #ccc;
    vertical-align: top;
    overflow-wrap: break-word;
    word-break: break-word;
}

.response-note-cell {
    display: none;
}

@media screen and (max-width: 768px) {
    body {
        font-size: 13px;
        line-height: 18px;
        padding: 5px;
    }
    
    h4 {
        font-size: 16px;
    }
    
    .table-responsive {
        overflow-x: auto;
        -webkit-overflow-scrolling: touch;
    }
    
    table {
        min-width: 600px;
    }
    
    th, td {
        padding: 5px;
    }
    
    .headerCell h6 {
        font-size: 12px;
    }
}

@media screen and (max-width: 480px) {
    body {
        font-size: 12px;
        line-height: 16px;
        padding: 5px;
    }
    
    h4 {
        font-size: 14px;
    }
    
    th, td {
        padding: 4px;
        font-size: 11.5px;
    }
}

@media print {
    @page { 
        size: A4 landscape; 
        margin: 8mm 10mm 8mm 10mm; 
    }
    @page :first { 
        margin-top: 8mm; 
    }
    
    body {
        font-size: 9.5pt;
        line-height: 13.5pt;
        padding: 0;
        background: #fff;
    }
    
    h4 {
        font-size: 14pt;
        margin: 6px 0;
        border-bottom: 1.5pt solid #333;
    }
    
    table {
        page-break-inside: auto;
        table-layout: fixed;
        width: 100%;
    }
    
    thead {
        display: table-header-group;
    }
    
    tr {
        page-break-inside: avoid;
    }
    
    th, td {
        padding: 4pt 6pt;
        border: 0.5pt solid #999;
        overflow-wrap: break-word;
        word-break: break-word;
    }
    
    .headerCell {
        background: #e9ecef !important;
        color: #000 !important;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
    }
    
    .headerCell h6 {
        color: #000 !important;
        font-size: 10pt;
    }
    
    td[colspan="3"],
    .section-header-dark {
        background: #f0f4f8 !important;
        color: #000 !important;
        font-weight: bold;
        font-size: 10.5pt;
        border-top: 1.5pt solid #333 !important;
        border-bottom: 1.5pt solid #333 !important;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
    }
    
    tr.gray td {
        background: #fff !important;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
    }

    img.questionPrompt {
        max-width: 100%;
        max-height: 200px;
        height: auto;
        object-fit: contain;
        page-break-inside: avoid;
        display: block;
        margin: 4px 0;
    }

    img.choicePrompt {
        max-width: 100px;
        max-height: 80px;
        height: auto;
        object-fit: contain;
        page-break-inside: avoid;
    }
}

.media-image-container {
    margin: 4px 0;
}

img.questionPrompt { 
    max-width: 100%; 
    max-height: 260px;
    height: auto;
    object-fit: contain;
    padding: 2px; 
    border: 1px solid #ddd; 
    border-radius: 3px;
    display: block; 
    margin: 4px 0; 
}

img.choicePrompt { 
    max-width: 120px; 
    max-height: 90px;
    height: auto;
    object-fit: contain;
    padding: 2px; 
    border: 1px solid #ddd; 
    border-radius: 3px;
    vertical-align: middle;
    margin: 2px 4px 2px 0;
    display: inline-block;
}

.media-image-missing {
    font-size: 11px;
    font-style: italic;
    color: #868e96;
    background: #f8f9fa;
    border: 1px dashed #ced4da;
    border-radius: 3px;
    padding: 3px 6px;
    margin: 3px 0;
    display: inline-block;
}

.required { 
    color: red; 
    font-style: italic; 
}

.relevance {
    font-size: 11.5px;
    font-style: italic;
    color: #2b7013;
}

.hint {
    color: #1864ab;
    font-size: 11.5px;
}

.metadata-tag {
    font-size: 11.5px;
    color: #495057;
    background: #f1f3f5;
    border-left: 3px solid #6c757d;
    padding: 2px 6px;
    margin-top: 3px;
    border-radius: 2px;
}

.calculation {
    font-size: 11.5px;
    color: #004085;
    background: #e7f5ff;
    border-left: 3px solid #1864ab;
    padding: 3px 6px;
    margin-top: 3px;
    word-break: break-all;
    border-radius: 2px;
}

.calc-desc {
    font-size: 10.5px;
    font-style: italic;
    color: #495057;
    margin-top: 1px;
}
"""


# ==============================================================================
# PLATFORM METADATA & CALCULATION TAXONOMY
# ==============================================================================

# Dictionary of recognized mobile data collection metadata fields.
# Maps variable type strings to a tuple of (Platform Information, Purpose Description).
METADATA_DESCRIPTIONS = {
    'start': ('ODK / SurveyCTO / KoboToolbox Metadata', 'Automated timestamp recorded when the survey session starts.'),
    'end': ('ODK / SurveyCTO / KoboToolbox Metadata', 'Automated timestamp recorded when the survey session ends.'),
    'today': ('ODK / SurveyCTO / KoboToolbox Metadata', 'Current date recorded when the survey is conducted.'),
    'deviceid': ('ODK / SurveyCTO / KoboToolbox Metadata', 'Unique hardware identifier (IMEI/UUID) of the collection device.'),
    'phonenumber': ('ODK / SurveyCTO / KoboToolbox Metadata', 'Phone number associated with the SIM card in the device.'),
    'simserial': ('SurveyCTO / ODK Metadata', 'Serial number of the SIM card installed in the device.'),
    'subscriberid': ('SurveyCTO / ODK Metadata', 'IMSI subscriber identifier of the SIM card.'),
    'username': ('ODK / KoboToolbox / SurveyCTO Metadata', 'Username of the logged-in field enumerator.'),
    'email': ('ODK / KoboToolbox Metadata', 'Email address of the logged-in enumerator.'),
    'audit': ('ODK / SurveyCTO / KoboToolbox Metadata', 'User activity audit log file tracking timestamped interaction events.'),
    'text-audit': ('SurveyCTO Metadata', 'Keystroke and text editing audit log file.'),
    'start-geopoint': ('ODK / SurveyCTO Metadata', 'Initial GPS location captured upon opening the form.'),
    'background-audio': ('ODK Metadata', 'Automated background audio recording captured during the survey.'),
    'caseid': ('SurveyCTO Metadata', 'Case Management unique entity identifier.'),
    'caseread': ('SurveyCTO Metadata', 'Case Management pre-loaded read field.'),
    'casesave': ('SurveyCTO Metadata', 'Case Management saved status field.'),
}


def classify_calculation(formula_str):
    """
    Classify an XLSForm calculation formula into an architectural category with explanatory details.

    Detects specialized functions:
    - `pulldata()`: Pre-loaded CSV dataset lookups
    - `jr:choice-name()`: Choice label resolutions
    - `now()`, `once(now())`: Timestamp captures
    - `format-date-time()`: Date/time formatting
    - `decimal-date-time()`: Duration mathematics
    - `selected-at()`, `count-selected()`: Multi-select evaluations
    - `concat()`, `substr()`, `string-length()`: String manipulations
    - `if()`: Conditional logic branching

    Args:
        formula_str (str): Raw calculation formula expression from survey sheet.

    Returns:
        tuple[str, str]: (Category Badge Title, Descriptive Purpose Explanation)
    """
    if not formula_str:
        return 'Calculation', ''
    
    formula_lower = formula_str.lower()
    
    if 'pulldata(' in formula_lower:
        return 'Calculation (SurveyCTO / ODK External Dataset Query)', 'Queries pre-loaded CSV dataset or server case data using unique lookup key.'
    elif 'jr:choice-name(' in formula_lower:
        return 'Calculation (ODK / SurveyCTO Choice Label Lookup)', 'Resolves raw choice option code into human-readable label text.'
    elif 'once(now())' in formula_lower or 'now()' in formula_lower:
        return 'Calculation (Timestamp Capture)', 'Captures current date/time upon evaluation.'
    elif 'format-date-time(' in formula_lower:
        return 'Calculation (Date/Time Formatter)', 'Formats raw timestamp into custom formatted date/time string.'
    elif 'decimal-date-time(' in formula_lower:
        return 'Calculation (Duration / Datetime Math)', 'Converts datetime to decimal days for duration math.'
    elif 'selected-at(' in formula_lower or 'count-selected(' in formula_lower:
        return 'Calculation (ODK Multi-Select / Repeat Function)', 'Extracts or counts selected response options in multiple-choice field.'
    elif 'substr(' in formula_lower or 'concat(' in formula_lower or 'string-length(' in formula_lower:
        return 'Calculation (String Processing)', 'Performs text manipulation, substring extraction, or string concatenation.'
    elif 'if(' in formula_lower:
        return 'Calculation (Conditional Logic)', 'Evaluates conditional logic expression.'
    else:
        return 'Calculation', ''


# Maximum character length for displayed calculation formulas before truncation
CALC_TRIM_LENGTH = 65

# Group/Repeat section tag tuples
SECTION_START_PREFIXES = ('begin group', 'begin_group', 'begin repeat', 'begin_repeat')
SECTION_END_TYPES = ('end_group', 'end group', 'end_repeat', 'end repeat')


# ==============================================================================
# XLSFORM PARSING HELPERS
# ==============================================================================

def split_type(raw_type):
    """
    Split an XLSForm `type` cell into its lowercase base type and trailing list name.

    Example:
        'select_one village_list' -> ('select_one', 'village_list')
        'text'                    -> ('text', '')

    Args:
        raw_type (str): Raw string value from 'type' column.

    Returns:
        tuple[str, str]: (base_type, list_name)
    """
    parts = str(raw_type or '').strip().split(None, 1)
    base_type = parts[0].lower() if parts else ''
    list_name = parts[1].strip() if len(parts) > 1 else ''
    return base_type, list_name


def is_section_start(raw_type):
    """
    Determine if a row begins a group or repeat container block.

    Args:
        raw_type (str): Raw question type string.

    Returns:
        bool: True if row starts a group or repeat.
    """
    return str(raw_type or '').strip().lower().startswith(SECTION_START_PREFIXES)


def is_section_end(raw_type):
    """
    Determine if a row closes a group or repeat container block.

    Args:
        raw_type (str): Raw question type string.

    Returns:
        bool: True if row closes a group or repeat.
    """
    return str(raw_type or '').strip().lower() in SECTION_END_TYPES


def trim_calculation(formula, max_len=CALC_TRIM_LENGTH):
    """
    Shorten a formula for clean tabular display while retaining ellipsis cue.

    Args:
        formula (str): Raw calculation expression.
        max_len (int): Maximum allowable character count.

    Returns:
        str: Truncated formula string with '...' if length exceeded.
    """
    formula = formula or ''
    if len(formula) > max_len:
        return formula[:max_len].rstrip() + '...'
    return formula


def extract_language_name(header_str):
    """
    Extract language identifier from an XLSForm localized header string.

    Example:
        'label::English (en)' -> 'English (en)'
        'label'               -> 'Default'

    Args:
        header_str (str): Raw column header text.

    Returns:
        str: Extracted language name.
    """
    if '::' in header_str:
        return header_str.split('::', 1)[1].strip()
    return 'Default'


def clean_language_suffix(lang_name):
    """
    Sanitize language name for clean output file naming.

    Strips parenthetical locale codes (e.g., 'English (en)' -> 'English').

    Args:
        lang_name (str): Raw language name.

    Returns:
        str: Sanitized language string suitable for filenames.
    """
    if not lang_name or lang_name == 'Default':
        return ''
    match = re.match(r'^([^\(\)]+)', lang_name)
    if match:
        return match.group(1).strip()
    return lang_name.strip()


def get_header_index(header_map, *keys):
    """
    Safely look up column index for any candidate key in a header map.

    Avoids Python's 0-evaluating-as-falsy pitfall.

    Args:
        header_map (dict[str, int]): Header string to column index mapping.
        *keys (str): Candidate column names to search.

    Returns:
        int | None: Matching column index or None.
    """
    for key in keys:
        if key in header_map:
            return header_map[key]
    return None


def get_lang_value(dict_of_langs, target_lang):
    """
    Retrieve localized string with intelligent multi-stage fallback resolution.

    Args:
        dict_of_langs (dict[str, str]): Mapping of language names to text.
        target_lang (str): Desired language code/name.

    Returns:
        str: Resolved localized string or empty string.
    """
    if not dict_of_langs:
        return ''
    if target_lang in dict_of_langs and dict_of_langs[target_lang]:
        return dict_of_langs[target_lang]
    clean_target = clean_language_suffix(target_lang).lower()
    for key, val in dict_of_langs.items():
        if val and clean_target in key.lower():
            return val
    if 'Default' in dict_of_langs and dict_of_langs['Default']:
        return dict_of_langs['Default']
    for val in dict_of_langs.values():
        if val:
            return val
    return ''


def resolve_image_path(image_filename, excel_path):
    """
    Locate an image file referenced in an XLSForm.

    Searches:
    1. The directory containing the XLSForm spreadsheet.
    2. A 'media' sub-folder within the form's directory.
    3. An ODK standard '<form_base>-media' sub-folder.
    4. Relative sub-directories if image_filename contains path separators.
    5. The current working directory and 'media/' in CWD.
    6. Case-insensitive filesystem matching across all candidate locations.

    Args:
        image_filename (str): Name or relative path of the media image file.
        excel_path (str): Filesystem path to the source XLSForm .xlsx file.

    Returns:
        str | None: Absolute path to the resolved image file, or None if not found.
    """
    if not image_filename or not str(image_filename).strip():
        return None

    image_filename = str(image_filename).strip().replace('\\', '/')
    base_image_name = os.path.basename(image_filename)

    form_dir = os.path.dirname(os.path.abspath(excel_path)) if excel_path else os.getcwd()
    form_base = os.path.splitext(os.path.basename(excel_path))[0] if excel_path else ''

    candidate_dirs = [
        form_dir,
        os.path.join(form_dir, 'media'),
        os.path.join(form_dir, f"{form_base}-media"),
    ]
    if os.path.abspath(os.getcwd()) != os.path.abspath(form_dir):
        candidate_dirs.extend([
            os.getcwd(),
            os.path.join(os.getcwd(), 'media')
        ])

    if '/' in image_filename:
        for cdir in list(candidate_dirs):
            candidate_dirs.append(os.path.join(cdir, os.path.dirname(image_filename)))

    # Direct match check
    for cdir in candidate_dirs:
        p1 = os.path.normpath(os.path.join(cdir, image_filename))
        if os.path.isfile(p1):
            return p1
        p2 = os.path.normpath(os.path.join(cdir, base_image_name))
        if os.path.isfile(p2):
            return p2

    # Case-insensitive match check across candidate directories
    target_lower = base_image_name.lower()
    for cdir in candidate_dirs:
        if os.path.isdir(cdir):
            try:
                for entry in os.listdir(cdir):
                    if entry.lower() == target_lower:
                        full_p = os.path.normpath(os.path.join(cdir, entry))
                        if os.path.isfile(full_p):
                            return full_p
            except OSError:
                pass

    return None


def image_to_base64_data_uri(image_path):
    """
    Encode an image file on disk into a self-contained Base64 data URI string.

    Args:
        image_path (str): Filesystem path to the image file.

    Returns:
        str | None: 'data:image/<type>;base64,<payload>' or None on failure.
    """
    if not image_path or not os.path.isfile(image_path):
        return None
    try:
        mime, _ = mimetypes.guess_type(image_path)
        if not mime or not mime.startswith('image/'):
            ext = os.path.splitext(image_path)[1].lower()
            ext_map = {
                '.jpg': 'image/jpeg',
                '.jpeg': 'image/jpeg',
                '.png': 'image/png',
                '.gif': 'image/gif',
                '.webp': 'image/webp',
                '.svg': 'image/svg+xml',
                '.bmp': 'image/bmp'
            }
            mime = ext_map.get(ext, 'image/jpeg')

        with open(image_path, 'rb') as f:
            b64_str = base64.b64encode(f.read()).decode('ascii')
        return f"data:{mime};base64,{b64_str}"
    except Exception:
        return None


def get_scaled_image_dimensions(image_path, max_width_mm, max_height_mm):
    """
    Calculate proportional image dimensions (in mm) constrained by max width and height.

    Args:
        image_path (str): Filesystem path to the image.
        max_width_mm (float): Maximum allowable width in millimeters.
        max_height_mm (float): Maximum allowable height in millimeters.

    Returns:
        tuple[float, float]: (scaled_width_mm, scaled_height_mm)
    """
    if not PIL_AVAILABLE or not image_path or not os.path.isfile(image_path):
        return max_width_mm, max_height_mm
    try:
        with Image.open(image_path) as img:
            w_px, h_px = img.size
            if w_px <= 0 or h_px <= 0:
                return max_width_mm, max_height_mm
            aspect = w_px / h_px
            w_from_h = max_height_mm * aspect
            if w_from_h <= max_width_mm:
                return round(w_from_h, 2), round(max_height_mm, 2)
            else:
                return round(max_width_mm, 2), round(max_width_mm / aspect, 2)
    except Exception:
        return max_width_mm, max_height_mm

# ==============================================================================
# EXCEL WORKBOOK PARSER (XLSFormParser)
# ==============================================================================

class XLSFormParser:
    """
    Parses an ODK / KoboToolbox / SurveyCTO XLSForm workbook into structured Python dictionaries.

    Engineered specifically for technical codebook generation:
    - Normalizes worksheet names case-insensitively ('survey', 'choices', 'settings').
    - Evaluates cached formula values via openpyxl's `data_only=True` mode.
    - Captures form-wide settings (title, form_id, version).
    - Detects all localized language column sets (label::[Language], hint::[Language], constraint_message::[Language]).
    - Parses option choices with both raw value codes and localized labels.
    - Extracts complete technical metadata per survey question (name, type, required status,
      hints, constraint expressions, constraint messages, relevance expressions, and calculations).

    Attributes:
        excel_path (str): Filesystem path to source .xlsx file.
        wb (openpyxl.Workbook): Openpyxl workbook handle in data_only mode.
        sheets (dict[str, str]): Lowercase-to-actual worksheet name lookup mapping.
        settings (dict[str, str]): Settings key-value pairs.
        languages (list[str]): List of detected language names.
        choices (dict[str, list[dict]]): Parsed option choice lists keyed by list_name.
        survey_rows (list[dict]): Parsed survey rows with full technical logic.
    """

    def __init__(self, excel_path):
        """
        Load workbook and execute parsing sequence across all sheets.

        Args:
            excel_path (str): Path to input XLSForm spreadsheet (.xlsx).
        """
        self.excel_path = excel_path
        # data_only=True loads evaluated formula results instead of raw formula strings
        self.wb = openpyxl.load_workbook(excel_path, data_only=True)
        self.sheets = {name.lower().strip(): name for name in self.wb.sheetnames}
        
        self.settings = self._parse_settings()
        self.languages = self._detect_languages()
        self.choices = self._parse_choices()
        self.survey_rows = self._parse_survey()

    def _get_sheet(self, target_name):
        """
        Retrieve worksheet by case-insensitive name.

        Args:
            target_name (str): Target sheet name (e.g. 'survey', 'choices', 'settings').

        Returns:
            openpyxl.worksheet.worksheet.Worksheet | None: Openpyxl sheet or None if absent.
        """
        actual_name = self.sheets.get(target_name.lower())
        if actual_name:
            return self.wb[actual_name]
        return None

    def _parse_settings(self):
        """
        Parse form-level metadata from the 'settings' sheet.

        Returns:
            dict[str, str]: Configuration settings dictionary.
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
        Scan headers in 'survey' and 'choices' sheets to identify all defined language translations.

        Returns:
            list[str]: Unique list of detected language names (defaults to ['Default']).
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
        Parse option choice rows from the 'choices' sheet.

        Constructs a dictionary where each `list_name` maps to an array of option dicts
        containing the raw stored code ('name') and localized text labels ('labels').

        Returns:
            dict[str, list[dict]]: Choice list dictionary keyed by list_name.
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
            images = {}
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

                # Option Media Image per language
                img_col_idx = get_header_index(header_map, f'media::image::{lang}', f'image::{lang}')
                if img_col_idx is None and lang != 'Default':
                    for h, idx in header_map.items():
                        h_lower = h.lower()
                        if (h_lower.startswith('media::image::') or h_lower.startswith('image::')) and lang.lower() in h_lower:
                            img_col_idx = idx
                            break
                if img_col_idx is None:
                    img_col_idx = get_header_index(header_map, 'media::image', 'image')

                if img_col_idx is not None and img_col_idx < len(row_cells):
                    img_val = row_cells[img_col_idx].value
                    images[lang] = str(img_val).strip() if img_val is not None else ''
                else:
                    images[lang] = ''

            # Fallback for unqualified choice image column
            def_img_col = get_header_index(header_map, 'media::image', 'image')
            if def_img_col is not None and def_img_col < len(row_cells):
                def_img_val = str(row_cells[def_img_col].value or '').strip()
                if def_img_val:
                    images['Default'] = def_img_val
                    for l in self.languages:
                        if not images.get(l):
                            images[l] = def_img_val

            choices[list_name_val].append({
                'name': name_val,
                'labels': labels,
                'images': images
            })

        return choices

    def _parse_survey(self):
        """
        Parse all technical survey question rows from the 'survey' sheet.

        Extracts comprehensive question metadata:
        - `type`: Question type string (e.g. 'select_one q1', 'calculate', 'begin group').
        - `name`: Variable field name.
        - `required`: Boolean flag indicating if response is mandatory.
        - `labels`: Dictionary of localized prompt labels.
        - `hints`: Dictionary of localized enumerator notes.
        - `constraint_msgs`: Dictionary of localized error messages.
        - `relevance`: Skip logic relevance expression.
        - `constraint`: Validation check constraint expression.
        - `calculation`: Calculation formula expression.

        Returns:
            list[dict]: Ordered collection of parsed question definitions.
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

            req_col = get_header_index(header_map, 'required')
            is_required = False
            if req_col is not None and req_col < len(row_cells):
                req_val = str(row_cells[req_col].value or '').strip().lower()
                is_required = req_val in ['yes', '1', 'true', 'ok', 'required']

            labels = {}
            hints = {}
            images = {}
            constraint_msgs = {}

            for lang in self.languages:
                # Localized question label
                col_idx = get_header_index(header_map, f'label::{lang}')
                if col_idx is None and lang != 'Default':
                    for h, idx in header_map.items():
                        if h.startswith('label::') and lang.lower() in h.lower():
                            col_idx = idx
                            break
                if col_idx is None:
                    col_idx = get_header_index(header_map, 'label')
                labels[lang] = str(row_cells[col_idx].value or '').strip() if (col_idx is not None and col_idx < len(row_cells) and row_cells[col_idx].value is not None) else ''

                # Localized enumerator hint
                col_idx = get_header_index(header_map, f'hint::{lang}')
                if col_idx is None and lang != 'Default':
                    for h, idx in header_map.items():
                        if h.startswith('hint::') and lang.lower() in h.lower():
                            col_idx = idx
                            break
                if col_idx is None:
                    col_idx = get_header_index(header_map, 'hint')
                hints[lang] = str(row_cells[col_idx].value or '').strip() if (col_idx is not None and col_idx < len(row_cells) and row_cells[col_idx].value is not None) else ''

                # Localized question media image
                col_idx = get_header_index(header_map, f'media::image::{lang}', f'image::{lang}')
                if col_idx is None and lang != 'Default':
                    for h, idx in header_map.items():
                        h_lower = h.lower()
                        if (h_lower.startswith('media::image::') or h_lower.startswith('image::')) and lang.lower() in h_lower:
                            col_idx = idx
                            break
                if col_idx is None:
                    col_idx = get_header_index(header_map, 'media::image', 'image')
                images[lang] = str(row_cells[col_idx].value or '').strip() if (col_idx is not None and col_idx < len(row_cells) and row_cells[col_idx].value is not None) else ''

                # Localized constraint validation message
                col_idx = get_header_index(header_map, f'constraint_message::{lang}')
                if col_idx is None and lang != 'Default':
                    for h, idx in header_map.items():
                        if h.startswith('constraint_message::') and lang.lower() in h.lower():
                            col_idx = idx
                            break
                if col_idx is None:
                    col_idx = get_header_index(header_map, 'constraint_message')
                constraint_msgs[lang] = str(row_cells[col_idx].value or '').strip() if (col_idx is not None and col_idx < len(row_cells) and row_cells[col_idx].value is not None) else ''

            # Fallback for unqualified question image column
            def_img_col = get_header_index(header_map, 'media::image', 'image')
            if def_img_col is not None and def_img_col < len(row_cells):
                def_img_val = str(row_cells[def_img_col].value or '').strip()
                if def_img_val:
                    images['Default'] = def_img_val
                    for l in self.languages:
                        if not images.get(l):
                            images[l] = def_img_val

            rel_col = get_header_index(header_map, 'relevant', 'relevance')
            relevance = str(row_cells[rel_col].value or '').strip() if (rel_col is not None and rel_col < len(row_cells) and row_cells[rel_col].value is not None) else ''

            con_col = get_header_index(header_map, 'constraint')
            constraint = str(row_cells[con_col].value or '').strip() if (con_col is not None and con_col < len(row_cells) and row_cells[con_col].value is not None) else ''

            calc_col = get_header_index(header_map, 'calculation', 'calculate')
            calculation = str(row_cells[calc_col].value or '').strip() if (calc_col is not None and calc_col < len(row_cells) and row_cells[calc_col].value is not None) else ''

            survey_rows.append({
                'type': type_val,
                'name': name_val,
                'required': is_required,
                'labels': labels,
                'hints': hints,
                'images': images,
                'constraint_msgs': constraint_msgs,
                'relevance': relevance,
                'constraint': constraint,
                'calculation': calculation,
            })

        return survey_rows


def get_choice_list_by_name(choices_dict, list_name):
    """
    Retrieve option list from choices dictionary with case-insensitive tolerance.

    Args:
        choices_dict (dict[str, list[dict]]): Parsed choice options grouped by list_name.
        list_name (str): Expected choice list name.

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
# HTML CODEBOOK BUILDER (3-COLUMN TABLE)
# ==============================================================================

def build_codebook_html(parser, language='Default'):
    """
    Construct a responsive, 3-column HTML codebook for a designated survey language.

    Layout Architecture:
    - Title Banner: Centered form title with underline rule.
    - Outer Table: Fixed-width responsive table with `table-layout: fixed`.
    - Column 1 (22% width): Variable field name in bold, required marker in red italic.
    - Column 2 (50% width): Question text, platform metadata annotations,
      categorized calculation badges (with hover tooltip of raw formula),
      enumerator hints, and skip logic / relevance conditions.
    - Column 3 (28% width): Two-column nested choices table (choice value | label),
      or blank area for freeform responses.
    - Section Rows: Full-width spanning cells (`colspan="3"`) with dark background
      styling (`#2B2D42`) for groups and repeats.

    Args:
        parser (XLSFormParser): Parsed XLSForm workbook instance.
        language (str): Target language to extract and display.

    Returns:
        str: Serialized HTML document string.
    """
    soup = BeautifulSoup('<!DOCTYPE html><html><head></head><body></body></html>', 'html.parser')
    
    head = soup.head
    head.append(soup.new_tag('meta', attrs={'charset': 'utf-8'}))
    head.append(soup.new_tag('meta', attrs={'name': 'viewport', 'content': 'width=device-width, initial-scale=1.0'}))
    
    title_text = parser.settings.get('form_title') or parser.settings.get('title') or os.path.splitext(os.path.basename(parser.excel_path))[0]
    title_tag = soup.new_tag('title')
    title_tag.string = title_text
    head.append(title_tag)
    
    style_tag = soup.new_tag('style')
    style_tag.string = RESPONSIVE_STYLES
    head.append(style_tag)

    body = soup.body
    container = soup.new_tag('div', attrs={'class': 'codebook-container'})
    body.append(container)

    h4 = soup.new_tag('h4')
    h4.string = title_text
    container.append(h4)

    table_wrapper = soup.new_tag('div', attrs={'class': 'table-responsive'})
    container.append(table_wrapper)

    table = soup.new_tag('table', attrs={'class': 'table table-bordered table-condensed'})
    table_wrapper.append(table)

    # Table Header with 3-column fixed widths
    thead = soup.new_tag('thead')
    table.append(thead)
    tr_head = soup.new_tag('tr')
    thead.append(tr_head)

    cols = [('22%', 'Field'), ('50%', 'Question'), ('28%', 'Answer')]
    for width, label in cols:
        th = soup.new_tag('th', attrs={'class': 'headerCell', 'style': f'width: {width};'})
        h6 = soup.new_tag('h6')
        h6.string = label
        th.append(h6)
        tr_head.append(th)

    tbody = soup.new_tag('tbody')
    table.append(tbody)

    for item in parser.survey_rows:
        raw_type = item['type'].strip()
        base_type, list_name = split_type(raw_type)

        # 1. Section Header: Check for group and repeat container beginnings
        if is_section_start(raw_type):
            label = get_lang_value(item['labels'], language) or item['name']
            if not label and 'group' in raw_type.lower():
                label = item['name']
            
            tr = soup.new_tag('tr', attrs={'class': 'entryRow'})
            td = soup.new_tag('td', attrs={'colspan': '3', 'class': 'section-header-dark', 'style': 'padding-left: 5px;'})
            span = soup.new_tag('span')
            span.string = label
            td.append(span)
            tr.append(td)
            tbody.append(tr)
            continue

        # 2. Suppress closing group / repeat tags
        if is_section_end(raw_type):
            continue

        # 3. Standard Field Row
        tr = soup.new_tag('tr', attrs={'class': 'gray entryRow' if item['name'] else 'entryRow'})
        
        # Column 1: Field Name (22%)
        td_field = soup.new_tag('td', attrs={'class': 'fieldCell', 'style': 'padding-left: 5px;'})
        span_field = soup.new_tag('span')
        span_field.string = item['name'] if item['name'] else ''
        if item['required']:
            span_req = soup.new_tag('span', attrs={'class': 'required'})
            span_req.string = ' (required)'
            span_field.append(span_req)
        td_field.append(span_field)
        tr.append(td_field)

        # Column 2: Question Prompt & Details (50%)
        td_q = soup.new_tag('td', attrs={'class': 'questionCell'})
        q_label = get_lang_value(item['labels'], language)
        if q_label:
            td_q.append(q_label)

        # Media Image Rendering for Question Prompt
        q_image = get_lang_value(item.get('images', {}), language)
        if q_image:
            img_path = resolve_image_path(q_image, parser.excel_path)
            if img_path:
                data_uri = image_to_base64_data_uri(img_path)
                if data_uri:
                    img_div = soup.new_tag('div', attrs={'class': 'media-image-container'})
                    img_tag = soup.new_tag('img', attrs={
                        'class': 'questionPrompt',
                        'src': data_uri,
                        'alt': q_label or q_image
                    })
                    img_div.append(img_tag)
                    td_q.append(img_div)
            else:
                missing_div = soup.new_tag('div', attrs={'class': 'media-image-missing'})
                missing_div.string = f"[Image: {q_image} (file not found)]"
                td_q.append(missing_div)

        # Platform Metadata Tagging: If variable type is a known background metadata field
        if base_type in METADATA_DESCRIPTIONS:
            platform_info, desc_text = METADATA_DESCRIPTIONS[base_type]
            div_meta = soup.new_tag('div', attrs={'class': 'metadata-tag'})
            div_meta.append(BeautifulSoup(f"<b>[{platform_info}]</b>: {desc_text}", 'html.parser'))
            td_q.append(div_meta)

        # Platform Calculation Display: Categorize calculation formula and display category badge + description
        if item['calculation'] or base_type == 'calculate':
            cat_title, cat_desc = classify_calculation(item['calculation'])
            div_calc = soup.new_tag('div', attrs={'class': 'calculation'})
            
            raw_calc = item['calculation']
            trimmed_calc = trim_calculation(raw_calc)

            calc_html = f"<b>{cat_title}</b>"
            if trimmed_calc:
                safe_raw = raw_calc.replace('&', '&amp;').replace('"', '&quot;')
                calc_html += f": <code title=\"{safe_raw}\">{trimmed_calc}</code>"
                
            div_calc.append(BeautifulSoup(calc_html, 'html.parser'))
            if cat_desc:
                div_desc = soup.new_tag('div', attrs={'class': 'calc-desc'})
                div_desc.string = cat_desc
                div_calc.append(div_desc)
            td_q.append(div_calc)

        # Guidance Hint
        q_hint = get_lang_value(item['hints'], language)
        if q_hint:
            div_hint = soup.new_tag('div', attrs={'class': 'hint'})
            div_hint.string = q_hint
            td_q.append(div_hint)

        # Skip Logic / Relevance Condition
        if item['relevance']:
            div_rel = soup.new_tag('div', attrs={'class': 'relevance'})
            div_rel.string = f"Relevance: {item['relevance']}"
            td_q.append(div_rel)

        tr.append(td_q)

        # Column 3: Answer / Choice Table (28%)
        td_a = soup.new_tag('td', attrs={'class': 'answerCell'})

        if base_type.startswith(('select_one', 'select_multiple', 'select_or_other')) and list_name:
            choice_list = get_choice_list_by_name(parser.choices, list_name)
            if choice_list:
                c_table = soup.new_tag('table', attrs={'class': 'table borderless', 'style': 'width: 100%; table-layout: fixed; margin: 0;'})
                c_tbody = soup.new_tag('tbody')
                c_table.append(c_tbody)

                for choice in choice_list:
                    c_tr = soup.new_tag('tr')
                    c_td_val = soup.new_tag('td', attrs={'style': 'width: 32px; min-width: 32px; font-weight: bold; text-align: center; vertical-align: top; padding: 2px 4px 2px 0;'})
                    c_td_val.string = choice['name']
                    
                    c_td_lbl = soup.new_tag('td', attrs={'style': 'vertical-align: top; padding: 2px 0 2px 5px; border-left: 1px solid #ddd; overflow-wrap: break-word; word-break: break-word;'})

                    # Choice Media Image
                    c_image = get_lang_value(choice.get('images', {}), language)
                    if c_image:
                        c_img_path = resolve_image_path(c_image, parser.excel_path)
                        if c_img_path:
                            c_data_uri = image_to_base64_data_uri(c_img_path)
                            if c_data_uri:
                                c_img_tag = soup.new_tag('img', attrs={
                                    'class': 'choicePrompt',
                                    'src': c_data_uri,
                                    'alt': c_image
                                })
                                c_td_lbl.append(c_img_tag)
                        else:
                            c_missing = soup.new_tag('div', attrs={'class': 'media-image-missing'})
                            c_missing.string = f"[Image: {c_image} (not found)]"
                            c_td_lbl.append(c_missing)

                    c_span = soup.new_tag('span')
                    c_span.string = get_lang_value(choice['labels'], language) or choice['name']
                    c_td_lbl.append(c_span)

                    c_tr.append(c_td_val)
                    c_tr.append(c_td_lbl)
                    c_tbody.append(c_tr)

                td_a.append(c_table)

        tr.append(td_a)
        tbody.append(tr)

    return str(soup)


# ==============================================================================
# OPENXML / WORD (.DOCX) CONSTANTS & SCHEMA ENGINE
# ==============================================================================

# Typography: Arial Unicode MS provides comprehensive glyph coverage across
# Devanagari, Bengali, Odia, Tamil, Telugu, Arabic, and Latin alphabets.
DOCX_FONT = 'Arial Unicode MS'
DOCX_MONO_FONT = 'Consolas'

# Font sizing scale (in typographical points)
DOCX_TITLE_PT = 14.0
DOCX_HEADER_PT = 10.0
DOCX_SECTION_PT = 10.5
DOCX_BODY_PT = 9.5
DOCX_SMALL_PT = 8.0
DOCX_TINY_PT = 7.5

# Hex Color Palette (without '#' prefix for WordprocessingML)
DOCX_COLOR_TEXT = '000000'
DOCX_COLOR_FIELD = '222222'
DOCX_COLOR_REQUIRED = 'CC0000'
DOCX_COLOR_HINT = '1864AB'
DOCX_COLOR_RELEVANCE = '2B7013'
DOCX_COLOR_META_TEXT = '495057'
DOCX_COLOR_META_RULE = '6C757D'
DOCX_COLOR_CALC_TEXT = '004085'
DOCX_COLOR_CALC_RULE = '1864AB'

DOCX_FILL_HEADER = 'E9ECEF'
DOCX_FILL_SECTION = 'F0F4F8'
DOCX_FILL_META = 'F1F3F5'
DOCX_FILL_CALC = 'E7F5FF'

DOCX_RULE_SECTION = '333333'
DOCX_RULE_CELL = '999999'
DOCX_RULE_CHOICE = 'CCCCCC'

# Landscape A4 Page Geometry (ISO 216)
DOCX_PAGE_WIDTH_MM = 297.0
DOCX_PAGE_HEIGHT_MM = 210.0
DOCX_MARGIN_MM = 10.0
DOCX_CONTENT_WIDTH_MM = DOCX_PAGE_WIDTH_MM - (2 * DOCX_MARGIN_MM)

# 3-Column Relative Width Distribution (22% / 50% / 28%)
DOCX_COL_FRACTIONS = (0.22, 0.50, 0.28)


# ------------------------------------------------------------------------------
# OpenXML ECMA-376 Strict Schema Ordering Definitions
# ------------------------------------------------------------------------------
# WordprocessingML validates property child elements by schema position rather than
# mere presence. Appending child elements out of order results in damaged document warnings.
# The tuples below enforce the exact schema sequence for paragraphs, runs, tables, cells, and rows.

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

_DOCX_TBLPR_ORDER = (
    'w:tblStyle', 'w:tblpPr', 'w:tblOverlap', 'w:bidiVisual', 'w:tblStyleRowBandSize',
    'w:tblStyleColBandSize', 'w:tblW', 'w:jc', 'w:tblCellSpacing', 'w:tblInd', 'w:tblBorders',
    'w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook', 'w:tblCaption', 'w:tblDescription',
    'w:tblPrChange',
)

_DOCX_TCPR_ORDER = (
    'w:cnfStyle', 'w:tcW', 'w:gridSpan', 'w:hMerge', 'w:vMerge', 'w:tcBorders', 'w:shd', 'w:noWrap',
    'w:tcMar', 'w:textDirection', 'w:tcFitText', 'w:vAlign', 'w:hideMark', 'w:headers', 'w:cellIns',
    'w:cellDel', 'w:cellMerge', 'w:tcPrChange',
)

_DOCX_TRPR_ORDER = (
    'w:cnfStyle', 'w:divId', 'w:gridBefore', 'w:gridAfter', 'w:wBefore', 'w:wAfter', 'w:cantSplit',
    'w:trHeight', 'w:tblHeader', 'w:tblPrEx', 'w:jc', 'w:hidden', 'w:ins', 'w:del', 'w:trPrChange',
)


# ==============================================================================
# OPENXML LOW-LEVEL PROPERTY HELPERS
# ==============================================================================

def _docx_set_child(parent, order, tag, element=None, replace=True):
    """
    Insert (or replace) a property child at the position its schema demands.

    Args:
        parent: Parent XML element.
        order (tuple[str]): Canonical schema order tuple.
        tag (str): Prefixed tag name.
        element: Element instance or None to construct new.
        replace (bool): Whether to replace matching existing child.

    Returns:
        OxmlElement: Inserted child element.
    """
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
    """Insert a boolean toggle child element (e.g. w:cantSplit, w:tblHeader)."""
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
    """Build a w:shd element for solid cell or paragraph fill."""
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill)
    return shd


def _docx_style_run(run, size=DOCX_BODY_PT, color=None, bold=False, italic=False, mono=False):
    """
    Apply font, size, weight, and colour to a run, covering complex-script attributes.

    Word only picks up Devanagari, Odia, Tamil, and Bengali glyphs when w:cs and
    w:eastAsia name the font too, so every run sets all four rFonts attributes.
    Complex-script twins (szCs, bCs, iCs) drive non-Latin font rendering; without
    them Word falls back to 10pt regular for those runs.
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
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rfonts.set(qn(attr), name)

    _docx_toggle(rpr, _DOCX_RPR_ORDER, 'w:szCs', str(int(round(size * 2))))
    if bold:
        _docx_toggle(rpr, _DOCX_RPR_ORDER, 'w:bCs')
    if italic:
        _docx_toggle(rpr, _DOCX_RPR_ORDER, 'w:iCs')
    return run


def _docx_tight_paragraph(paragraph, space_before=0.0, space_after=1.0):
    """Collapse paragraph spacing so stacked cell blocks stay compact."""
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(space_before)
    fmt.space_after = Pt(space_after)
    fmt.line_spacing = 1.0
    return paragraph


def _docx_decorate_paragraph(paragraph, fill=None, accent=None, indent_mm=0.0):
    """Give a paragraph the accent bar and tint used by metadata and calculation blocks."""
    ppr = paragraph._element.get_or_add_pPr()
    if accent:
        pbdr = _docx_set_child(ppr, _DOCX_PPR_ORDER, 'w:pBdr')
        pbdr.append(_docx_border('left', sz='12', color=accent, space='2'))
    if fill:
        _docx_set_child(ppr, _DOCX_PPR_ORDER, 'w:shd', _docx_shading(fill))
    if indent_mm:
        paragraph.paragraph_format.left_indent = Mm(indent_mm)
    return paragraph


def _docx_reset_cell(cell):
    """Strip the default empty paragraph so cell content starts flush at the top."""
    for paragraph in list(cell.paragraphs):
        paragraph._element.getparent().remove(paragraph._element)
    return cell


def _docx_add_cell_paragraph(cell, space_before=0.0, space_after=1.0):
    """Add a compact paragraph to a table cell."""
    paragraph = cell.add_paragraph()
    return _docx_tight_paragraph(paragraph, space_before=space_before, space_after=space_after)


def _docx_shade_cell(cell, fill):
    """Apply background shading color to a table cell."""
    tcpr = cell._element.get_or_add_tcPr()
    _docx_set_child(tcpr, _DOCX_TCPR_ORDER, 'w:shd', _docx_shading(fill))
    return cell


def _docx_cell_borders(cell, **edges):
    """Set explicit borders on one cell, e.g. top=('12', '333333')."""
    tcpr = cell._element.get_or_add_tcPr()
    borders = _docx_set_child(tcpr, _DOCX_TCPR_ORDER, 'w:tcBorders', replace=False)
    for edge, spec in edges.items():
        if spec is None:
            borders.append(_docx_border(edge, val='none', sz='0'))
        else:
            sz, color = spec
            borders.append(_docx_border(edge, sz=sz, color=color))
    return cell


def _docx_table_borders(table, sz='4', color=DOCX_RULE_CELL, inside_h=None, inside_v=None):
    """Replace a table's border set, defaulting inside rules to the outer rule."""
    tblpr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    outer = {'val': 'single', 'sz': sz, 'color': color} if sz != '0' else {'val': 'none', 'sz': '0', 'color': 'auto'}
    for edge in ('top', 'left', 'bottom', 'right'):
        borders.append(_docx_border(edge, val=outer['val'], sz=outer['sz'], color=outer['color']))
    for edge, spec in (('insideH', inside_h), ('insideV', inside_v)):
        if spec is None:
            borders.append(_docx_border(edge, val=outer['val'], sz=outer['sz'], color=outer['color']))
        elif spec == 'none':
            borders.append(_docx_border(edge, val='none', sz='0'))
        else:
            val, edge_sz, edge_color = spec
            borders.append(_docx_border(edge, val=val, sz=edge_sz, color=edge_color))
    _docx_set_child(tblpr, _DOCX_TBLPR_ORDER, 'w:tblBorders', borders)
    return table


def _docx_cell_margins(table, top=40, bottom=40, left=80, right=80):
    """Set uniform cell padding in twips (20 twips to the point)."""
    cell_mar = OxmlElement('w:tblCellMar')
    for edge, value in (('top', top), ('left', left), ('bottom', bottom), ('right', right)):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:w'), str(value))
        el.set(qn('w:type'), 'dxa')
        cell_mar.append(el)
    _docx_set_child(table._tbl.tblPr, _DOCX_TBLPR_ORDER, 'w:tblCellMar', cell_mar)
    return table


def _docx_fixed_layout(table, widths_mm):
    """Pin the table to a fixed layout with explicit column widths."""
    table.autofit = False
    tblpr = table._tbl.tblPr

    layout = OxmlElement('w:tblLayout')
    layout.set(qn('w:type'), 'fixed')
    _docx_set_child(tblpr, _DOCX_TBLPR_ORDER, 'w:tblLayout', layout)

    total_twips = int(sum(Mm(width).twips for width in widths_mm))
    tbl_w = OxmlElement('w:tblW')
    tbl_w.set(qn('w:w'), str(total_twips))
    tbl_w.set(qn('w:type'), 'dxa')
    _docx_set_child(tblpr, _DOCX_TBLPR_ORDER, 'w:tblW', tbl_w)

    grid = table._tbl.find(qn('w:tblGrid'))
    if grid is not None:
        for idx, col in enumerate(grid.findall(qn('w:gridCol'))):
            if idx < len(widths_mm):
                col.set(qn('w:w'), str(int(Mm(widths_mm[idx]).twips)))

    for idx, column in enumerate(table.columns):
        if idx >= len(widths_mm):
            break
        width = Mm(widths_mm[idx])
        column.width = width
        for cell in column.cells:
            cell.width = width
    return table


def _docx_apply_row_widths(row, widths_mm):
    """
    Stamp column widths onto a freshly added row.

    Table.add_row() copies an even split of the original table width onto each new cell,
    which overrides the fixed grid and leaves ragged columns down the page.
    """
    for cell, width in zip(row.cells, widths_mm):
        cell.width = Mm(width)
    return row


def _docx_row_properties(row, cant_split=True, header=False):
    """Keep a row off a page break and optionally repeat it as a header row."""
    trpr = row._tr.get_or_add_trPr()
    if cant_split:
        _docx_toggle(trpr, _DOCX_TRPR_ORDER, 'w:cantSplit')
    if header:
        _docx_toggle(trpr, _DOCX_TRPR_ORDER, 'w:tblHeader')
    return row


def _docx_setup_document(title_text):
    """Create a landscape A4 document with the shared Unicode-safe base style."""
    doc = Document()

    normal = doc.styles['Normal']
    normal.font.name = DOCX_FONT
    normal.font.size = Pt(DOCX_BODY_PT)
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rfonts.set(qn(attr), DOCX_FONT)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(1)
    normal.paragraph_format.line_spacing = 1.0

    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Mm(DOCX_PAGE_WIDTH_MM)
    section.page_height = Mm(DOCX_PAGE_HEIGHT_MM)
    section.top_margin = Mm(DOCX_MARGIN_MM)
    section.bottom_margin = Mm(DOCX_MARGIN_MM)
    section.left_margin = Mm(DOCX_MARGIN_MM)
    section.right_margin = Mm(DOCX_MARGIN_MM)

    heading = doc.add_paragraph()
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _docx_tight_paragraph(heading, space_after=6.0)
    ppr = heading._element.get_or_add_pPr()
    pbdr = _docx_set_child(ppr, _DOCX_PPR_ORDER, 'w:pBdr')
    pbdr.append(_docx_border('bottom', sz='12', color=DOCX_RULE_SECTION, space='4'))
    _docx_style_run(heading.add_run(title_text), size=DOCX_TITLE_PT, bold=True)

    return doc


def _docx_add_choice_table(cell, choice_list, language, width_mm, excel_path=None):
    """Render the two-column code/label choice list inside the Answer cell."""
    code_width = min(9.0, width_mm * 0.16)
    label_width = max(width_mm - code_width, code_width)

    table = cell.add_table(rows=0, cols=2)
    _docx_fixed_layout(table, [code_width, label_width])
    _docx_table_borders(table, sz='0', inside_h=('dotted', '4', DOCX_RULE_CHOICE), inside_v='none')
    _docx_cell_margins(table, top=10, bottom=10, left=0, right=20)

    for choice in choice_list:
        row = table.add_row()
        _docx_row_properties(row, cant_split=True)
        _docx_apply_row_widths(row, [code_width, label_width])

        code_cell = _docx_reset_cell(row.cells[0])
        code_para = _docx_add_cell_paragraph(code_cell, space_after=0.0)
        code_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _docx_style_run(code_para.add_run(choice['name']), size=DOCX_SMALL_PT, bold=True)

        label_cell = _docx_reset_cell(row.cells[1])
        _docx_cell_borders(label_cell, left=('4', 'DDDDDD'))

        # Choice image in DOCX
        c_image = get_lang_value(choice.get('images', {}), language)
        if c_image and excel_path:
            img_path = resolve_image_path(c_image, excel_path)
            if img_path and os.path.isfile(img_path):
                scaled_w, scaled_h = get_scaled_image_dimensions(img_path, max_width_mm=min(28.0, label_width - 2.0), max_height_mm=20.0)
                img_para = _docx_add_cell_paragraph(label_cell, space_after=1.0)
                r_img = img_para.add_run()
                r_img.add_picture(img_path, width=Mm(scaled_w), height=Mm(scaled_h))
            else:
                miss_para = _docx_add_cell_paragraph(label_cell, space_after=1.0)
                _docx_style_run(miss_para.add_run(f"[Image: {c_image} (not found)]"), size=DOCX_TINY_PT, color=DOCX_COLOR_META_TEXT, italic=True)

        label_para = _docx_add_cell_paragraph(label_cell, space_after=0.0)
        label_text = get_lang_value(choice['labels'], language) or choice['name']
        _docx_style_run(label_para.add_run(label_text), size=DOCX_SMALL_PT)

    # _Cell.add_table appends a trailing paragraph; shrink it so it costs no space.
    for paragraph in cell.paragraphs:
        _docx_tight_paragraph(paragraph, space_after=0.0)
        for run in paragraph.runs:
            _docx_style_run(run, size=1.0)
        if not paragraph.runs:
            _docx_style_run(paragraph.add_run(''), size=1.0)
    return table


def _docx_add_section_row(table, label, widths_mm):
    """Add the full-width group/repeat band, merged across all three columns."""
    row = table.add_row()
    _docx_row_properties(row, cant_split=True)
    _docx_apply_row_widths(row, widths_mm)
    cell = row.cells[0].merge(row.cells[2])
    cell.width = Mm(sum(widths_mm))
    _docx_reset_cell(cell)
    _docx_shade_cell(cell, DOCX_FILL_SECTION)
    _docx_cell_borders(cell, top=('12', DOCX_RULE_SECTION), bottom=('12', DOCX_RULE_SECTION))
    paragraph = _docx_add_cell_paragraph(cell, space_after=0.0)
    _docx_style_run(paragraph.add_run(label), size=DOCX_SECTION_PT, bold=True)
    return row


# ==============================================================================
# WORD DOCUMENT BUILDER & CODEBOOK TABLE GENERATION
# ==============================================================================

def build_codebook_docx(parser, language='Default'):
    """
    Build an editable DOCX codebook from the parsed XLSForm, mirroring the HTML layout.

    Reads the same parsed survey rows and choice lists as build_codebook_html, so the
    Word file, the HTML file, and the PDF stay in step without a conversion round trip.
    """
    if not DOCX_AVAILABLE:
        raise RuntimeError("python-docx is not installed. Run 'pip install python-docx' or pass --no-docx.")

    title_text = parser.settings.get('form_title') or parser.settings.get('title') or os.path.splitext(os.path.basename(parser.excel_path))[0]
    doc = _docx_setup_document(title_text)

    col_widths = [DOCX_CONTENT_WIDTH_MM * fraction for fraction in DOCX_COL_FRACTIONS]

    table = doc.add_table(rows=1, cols=3)
    _docx_fixed_layout(table, col_widths)
    _docx_table_borders(table, sz='4', color=DOCX_RULE_CELL)
    _docx_cell_margins(table)

    header_row = table.rows[0]
    _docx_row_properties(header_row, cant_split=True, header=True)
    for cell, heading in zip(header_row.cells, ('Field', 'Question', 'Answer')):
        _docx_reset_cell(cell)
        _docx_shade_cell(cell, DOCX_FILL_HEADER)
        paragraph = _docx_add_cell_paragraph(cell, space_after=0.0)
        _docx_style_run(paragraph.add_run(heading), size=DOCX_HEADER_PT, bold=True)

    for item in parser.survey_rows:
        raw_type = item['type'].strip()
        base_type, list_name = split_type(raw_type)

        if is_section_start(raw_type):
            label = get_lang_value(item['labels'], language) or item['name']
            _docx_add_section_row(table, label, col_widths)
            continue

        if is_section_end(raw_type):
            continue

        row = table.add_row()
        _docx_row_properties(row, cant_split=True)
        _docx_apply_row_widths(row, col_widths)
        field_cell, question_cell, answer_cell = row.cells[0], row.cells[1], row.cells[2]

        # Column 1: field name, with the required marker in red
        _docx_reset_cell(field_cell)
        field_para = _docx_add_cell_paragraph(field_cell, space_after=0.0)
        _docx_style_run(field_para.add_run(item['name'] or ''), color=DOCX_COLOR_FIELD, bold=True)
        if item['required']:
            _docx_style_run(field_para.add_run(' (required)'), size=DOCX_SMALL_PT, color=DOCX_COLOR_REQUIRED, italic=True)

        # Column 2: label, metadata note, calculation badge, hint, relevance
        _docx_reset_cell(question_cell)
        q_label = get_lang_value(item['labels'], language)
        label_para = _docx_add_cell_paragraph(question_cell, space_after=0.0)
        _docx_style_run(label_para.add_run(q_label or ''))

        # Question Media Image in DOCX
        q_image = get_lang_value(item.get('images', {}), language)
        if q_image:
            img_path = resolve_image_path(q_image, parser.excel_path)
            if img_path and os.path.isfile(img_path):
                scaled_w, scaled_h = get_scaled_image_dimensions(
                    img_path,
                    max_width_mm=min(120.0, col_widths[1] - 8.0),
                    max_height_mm=75.0
                )
                img_para = _docx_add_cell_paragraph(question_cell, space_before=2.0, space_after=3.0)
                r_img = img_para.add_run()
                r_img.add_picture(img_path, width=Mm(scaled_w), height=Mm(scaled_h))
            else:
                miss_para = _docx_add_cell_paragraph(question_cell, space_before=2.0, space_after=2.0)
                _docx_style_run(miss_para.add_run(f"[Image: {q_image} (file not found)]"), size=DOCX_TINY_PT, color=DOCX_COLOR_META_TEXT, italic=True)

        if base_type in METADATA_DESCRIPTIONS:
            platform_info, desc_text = METADATA_DESCRIPTIONS[base_type]
            meta_para = _docx_add_cell_paragraph(question_cell, space_after=0.0)
            _docx_tight_paragraph(meta_para, space_before=2.0, space_after=0.0)
            _docx_decorate_paragraph(meta_para, fill=DOCX_FILL_META, accent=DOCX_COLOR_META_RULE, indent_mm=1.2)
            _docx_style_run(meta_para.add_run(f"[{platform_info}]"), size=DOCX_TINY_PT, color=DOCX_COLOR_META_TEXT, bold=True)
            _docx_style_run(meta_para.add_run(f": {desc_text}"), size=DOCX_TINY_PT, color=DOCX_COLOR_META_TEXT)

        if item['calculation'] or base_type == 'calculate':
            cat_title, cat_desc = classify_calculation(item['calculation'])
            calc_para = _docx_add_cell_paragraph(question_cell, space_after=0.0)
            _docx_tight_paragraph(calc_para, space_before=2.0, space_after=0.0)
            _docx_decorate_paragraph(calc_para, fill=DOCX_FILL_CALC, accent=DOCX_COLOR_CALC_RULE, indent_mm=1.2)
            _docx_style_run(calc_para.add_run(cat_title), size=DOCX_TINY_PT, color=DOCX_COLOR_CALC_TEXT, bold=True)
            trimmed_calc = trim_calculation(item['calculation'])
            if trimmed_calc:
                _docx_style_run(calc_para.add_run(': '), size=DOCX_TINY_PT, color=DOCX_COLOR_CALC_TEXT)
                _docx_style_run(calc_para.add_run(trimmed_calc), size=DOCX_TINY_PT, color=DOCX_COLOR_CALC_TEXT, mono=True)
            if cat_desc:
                desc_para = _docx_add_cell_paragraph(question_cell, space_after=0.0)
                _docx_decorate_paragraph(desc_para, fill=DOCX_FILL_CALC, indent_mm=1.2)
                _docx_style_run(desc_para.add_run(cat_desc), size=DOCX_TINY_PT, color=DOCX_COLOR_META_TEXT, italic=True)

        q_hint = get_lang_value(item['hints'], language)
        if q_hint:
            hint_para = _docx_add_cell_paragraph(question_cell, space_after=0.0)
            _docx_tight_paragraph(hint_para, space_before=2.0, space_after=0.0)
            _docx_style_run(hint_para.add_run(q_hint), size=DOCX_TINY_PT, color=DOCX_COLOR_HINT)

        if item['relevance']:
            rel_para = _docx_add_cell_paragraph(question_cell, space_after=0.0)
            _docx_tight_paragraph(rel_para, space_before=2.0, space_after=0.0)
            _docx_style_run(rel_para.add_run(f"Relevance: {item['relevance']}"), size=DOCX_TINY_PT, color=DOCX_COLOR_RELEVANCE, italic=True)

        # Column 3: choice list for select types, otherwise left blank for writing on
        _docx_reset_cell(answer_cell)
        choice_list = []
        if base_type.startswith(('select_one', 'select_multiple', 'select_or_other')) and list_name:
            choice_list = get_choice_list_by_name(parser.choices, list_name)

        if choice_list:
            _docx_add_choice_table(answer_cell, choice_list, language, col_widths[2] - 4.0, excel_path=parser.excel_path)
        else:
            blank_para = _docx_add_cell_paragraph(answer_cell, space_after=0.0)
            _docx_style_run(blank_para.add_run(''))

    return doc


def generate_docx(parser, docx_path, language='Default'):
    """Write the editable DOCX codebook for one language to disk."""
    doc = build_codebook_docx(parser, language=language)
    doc.save(docx_path)
    return docx_path


# ==============================================================================
# PDF RENDERING ENGINE (PLAYWRIGHT CHROMIUM)
# ==============================================================================

async def generate_pdf(html_path, pdf_path):
    """Render HTML codebook to landscape A4 PDF via Playwright Chromium."""
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()
        await page.goto(f'file:///{os.path.abspath(html_path)}', wait_until='networkidle')
        await page.wait_for_timeout(1000)
        await page.pdf(
            path=pdf_path,
            format='A4',
            landscape=True,
            margin={'top': '10mm', 'right': '10mm', 'bottom': '10mm', 'left': '10mm'},
            print_background=True
        )
        await browser.close()


# ==============================================================================
# PIPELINE ORCHESTRATION & CLI ENTRY POINT
# ==============================================================================

def process_xlsform(excel_path, target_lang=None, generate_pdf_flag=True, generate_docx_flag=True):
    """Process a single XLSForm spreadsheet and generate HTML, PDF, and DOCX codebooks."""
    print(f"Processing XLSForm: {os.path.basename(excel_path)}...")
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

        print(f"  Building HTML codebook for language: {lang} -> {os.path.basename(out_html_name)}")
        html_content = build_codebook_html(parser, language=lang)

        with open(out_html_name, 'w', encoding='utf-8') as f:
            f.write(html_content)
        print(f"    Saved HTML: {out_html_name}")

        if generate_pdf_flag:
            print(f"    Generating PDF: {out_pdf_name}...")
            asyncio.run(generate_pdf(out_html_name, out_pdf_name))
            print(f"    Saved PDF: {out_pdf_name}")

        if generate_docx_flag:
            print(f"    Generating DOCX: {out_docx_name}...")
            generate_docx(parser, out_docx_name, language=lang)
            print(f"    Saved DOCX: {out_docx_name}")


def main():
    """Parse CLI arguments and run codebook generator pipeline."""
    parser = argparse.ArgumentParser(description="Direct XLSForm to Printable Codebook, PDF & DOCX Generator")
    parser.add_argument('--input', '-i', help="Path to input XLSForm (.xlsx) file")
    parser.add_argument('--dir', '-d', help="Directory containing XLSForm files")
    parser.add_argument('--lang', '-l', help="Target language (e.g. English, Hindi). If omitted, processes all languages.")
    parser.add_argument('--no-pdf', action='store_true', help="Skip PDF generation.")
    parser.add_argument('--no-docx', action='store_true', help="Skip editable DOCX generation.")

    args = parser.parse_args()

    options = {
        'target_lang': args.lang,
        'generate_pdf_flag': not args.no_pdf,
        'generate_docx_flag': not args.no_docx,
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
