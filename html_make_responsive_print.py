"""
html_make_responsive_print.py
=============================
SurveyCTO HTML Table to Print-Optimized Questionnaire Formatter.

Purpose:
  SurveyCTO HTML codebook exports structure surveys as wide, multi-column HTML tables.
  When printed or exported to paper/PDF, wide tables often split awkwardly across pages,
  clip text, or leave huge blank gaps.

  This utility transforms tabular SurveyCTO HTML exports into a clean, vertical
  questionnaire layout (card/block format):
    1. DOM Restructuring: Recursively converts <table> elements into semantic card
       containers (<div class="question-item">).
    2. Column Decomposition:
       - Column 1 (Variable/Field Name) -> Header badge (<div class="field-label">).
       - Column 2 (Question Prompt/Relevance) -> Prominent text block (<div class="question-text">).
       - Column 3 (Choices/Answers) -> Choice list with checkboxes or lined fill-in answer areas.
    3. Section Headers: Detects group rows and converts them into full-width dividing bars.
    4. Print-Optimized Styling: Injects an A4 portrait CSS stylesheet with strict
       page-break avoidance rules (page-break-inside: avoid) on questions and choices,
       notebook-style ruled answer lines, and dual-mode styling (screen shadow preview vs.
       clean print graphics).

Usage:
  Run directly in a folder containing SurveyCTO HTML exports:
      python html_make_responsive_print.py
"""

import os
import re
from bs4 import BeautifulSoup
from bs4.element import Tag

# ==============================================================================
# PRINT-OPTIMIZED STYLESHEET
# ==============================================================================
# This stylesheet configures portrait A4 pagination, card-based question layouts,
# ruled writing lines for manual field entry, and page-break management.
PRINT_STYLES = """

/* ========== PRINT OPTIMIZED CSS FOR SURVEYCTO ========== */
*, *::before, *::after {
    box-sizing: border-box;
}

body {
    margin: 0;
    padding: 0;
    font-family: "Arial Unicode MS", "Arial Unicode", Arial, sans-serif;
    font-size: 10pt;
    line-height: 1.4;
    color: #000;
    background-color: #ffffff;
}

/* ========== PAGE SETUP FOR PRINTING ========== */
@page {
    size: A4;
    margin: 15mm 12mm 15mm 12mm;
}

@page :first {
    margin-top: 10mm;
}

/* ========== MAIN CONTAINER ========== */
.questionnaire-container {
    width: 100%;
    max-width: 210mm; /* A4 width */
    margin: 0 auto;
    padding: 0;
}

/* ========== QUESTIONNAIRE HEADER ========== */
.qn-header {
    text-align: center;
    padding: 10px 0 15px 0;
    border-bottom: 2px solid #333;
    margin-bottom: 20px;
}

.qn-header h1 {
    font-size: 16pt;
    font-weight: bold;
    margin: 0 0 5px 0;
    color: #000;
}

.qn-header h2 {
    font-size: 12pt;
    font-weight: normal;
    margin: 0;
    color: #333;
}

/* ========== QUESTION ITEM ========== */
.question-item {
    border: 1px solid #999;
    border-left: 4px solid #333;
    margin-bottom: 12px;
    page-break-inside: avoid;
    background: #fff;
}

.question-item.section-header {
    border: none;
    border-left: 4px solid #666;
    background: #f0f0f0;
    font-weight: bold;
    font-size: 11pt;
    padding: 8px 10px;
    margin-top: 15px;
}

/* ========== FIELD LABEL ========== */
.field-label {
    background: #f5f5f5;
    padding: 4px 8px;
    font-size: 8pt;
    font-weight: bold;
    color: #666;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    border-bottom: 1px solid #ddd;
}

.field-label .required {
    color: #c00;
    font-weight: normal;
}

/* ========== QUESTION TEXT ========== */
.question-text {
    padding: 10px 12px;
    font-size: 10pt;
    line-height: 1.5;
    border-bottom: 1px solid #eee;
    background: #fafafa;
}

.question-text strong {
    font-weight: bold;
}

.question-text .relevance {
    font-size: 8pt;
    color: #666;
    font-style: italic;
    margin-top: 5px;
    padding: 3px 6px;
    background: #f0f0f0;
    border-radius: 3px;
}

/* ========== ANSWER AREA ========== */
.answer-area {
    padding: 10px 12px;
    min-height: 25px;
}

/* Empty answer placeholder */
.answer-placeholder {
    border: 1px dashed #999;
    min-height: 30px;
    background: repeating-linear-gradient(
        0deg,
        #fff,
        #fff 24px,
        #eee 24px,
        #eee 25px
    );
    margin: 5px 0;
}

/* ========== CHOICE OPTIONS ========== */
.choice-list {
    list-style: none;
    padding: 0;
    margin: 0;
}

.choice-item {
    display: flex;
    align-items: flex-start;
    padding: 6px 0;
    border-bottom: 1px dotted #ddd;
    page-break-inside: avoid;
}

.choice-item:last-child {
    border-bottom: none;
}

.choice-number {
    width: 30px;
    flex-shrink: 0;
    font-weight: bold;
    color: #333;
    text-align: center;
    padding: 2px 5px;
    background: #f5f5f5;
    border: 1px solid #ccc;
    margin-right: 10px;
    font-size: 9pt;
}

.choice-text {
    flex: 1;
    font-size: 10pt;
    padding: 2px 0;
}

/* Checkbox/radio style */
.choice-checkbox {
    width: 16px;
    height: 16px;
    border: 2px solid #333;
    display: inline-block;
    margin-right: 8px;
    vertical-align: middle;
    flex-shrink: 0;
}

/* ========== NESTED QUESTIONS ========== */
.nested-question {
    margin-left: 15px;
    border-left: 2px solid #ccc;
    margin-top: 8px;
}

/* ========== TABLE FOR COMPLEX DATA ========== */
.data-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 9pt;
    margin: 8px 0;
}

.data-table th,
.data-table td {
    border: 1px solid #999;
    padding: 6px 8px;
    text-align: left;
}

.data-table th {
    background: #f0f0f0;
    font-weight: bold;
    font-size: 8pt;
    text-transform: uppercase;
}

.data-table td {
    vertical-align: top;
}

/* ========== IMAGE HANDLING ========== */
img {
    max-width: 100%;
    height: auto;
    max-height: 150px;
    display: block;
    margin: 8px 0;
}

img.questionPrompt,
img.choicePrompt {
    max-width: 200px;
    border: 1px solid #ddd;
    padding: 3px;
    background: #f9f9f9;
}

/* ========== NOTES AND HIGHLIGHTS ========== */
.note-cell {
    background: #fffacd;
    padding: 8px 12px;
    border-left: 4px solid #ffd700;
    font-style: italic;
}

.section-divider {
    border-top: 2px solid #333;
    margin: 20px 0 15px 0;
    padding-top: 10px;
}

/* ========== PRINT SPECIFIC ========== */
@media print {
    body {
        font-size: 10pt;
    }
    
    .question-item {
        page-break-inside: avoid;
    }
    
    .choice-item {
        page-break-inside: avoid;
    }
    
    .section-header {
        page-break-after: avoid;
    }
    
    /* Ensure images don't break awkwardly */
    img {
        page-break-inside: avoid;
        page-break-before: auto;
        page-break-after: auto;
    }
    
    /* Show URL after links */
    a[href]:after {
        content: " (" attr(href) ")";
        font-size: 8pt;
        color: #666;
    }
}

/* ========== SCREEN PREVIEW ========== */
@media screen {
    body {
        background: #e0e0e0;
        padding: 20px;
    }
    
    .questionnaire-container {
        background: #fff;
        padding: 20mm;
        box-shadow: 0 0 10px rgba(0,0,0,0.2);
    }
    
    .question-item {
        margin-bottom: 15px;
    }
}
"""

# ==============================================================================
# BASE UTILITY STYLES
# ==============================================================================
BASE_STYLES = """
.text-left { text-align: left; }
.text-right { text-align: right; }
.text-center { text-align: center; }
.clearfix { overflow: auto; }
"""


# ==============================================================================
# DOM TEXT & ELEMENT EXTRACTION HELPERS
# ==============================================================================

def extract_text_content(element):
    """
    Safely extract clean text from a BeautifulSoup element or string.

    Handles None, raw strings, and Tag instances with nested children, stripping
    leading and trailing whitespace.

    Args:
        element (bs4.element.Tag | str | None): The DOM element or text node.

    Returns:
        str: Cleaned text content, or an empty string if element is None.
    """
    if element is None:
        return ""
    if isinstance(element, str):
        return element
    return element.get_text(strip=True)


def parse_cell_content(cell, soup):
    """
    Parse an HTML table cell and extract structured content while preserving formatting.

    Traverses direct children of a table cell, converting block and inline elements:
      - '<p>' tags become '<div class="question-paragraph">' to avoid conflicting margin collapsing.
      - '<span>' text is stripped and extracted.
      - '<br>' tags are retained for intentional linebreaks in questionnaire prompts.
      - '<img>' tags (e.g. SurveyCTO survey logos or item icons) are preserved.
      - '<strong>' / '<b>' and '<em>' / '<i>' tags are normalized.

    Args:
        cell (bs4.element.Tag): The <td> or <th> element to parse.
        soup (bs4.BeautifulSoup): The root BeautifulSoup instance used for creating new tags.

    Returns:
        list: List of parsed text strings and BeautifulSoup Tag elements ready to be appended.
    """
    content = []
    
    # Process all children preserving semantic meaning
    for child in cell.children:
        if isinstance(child, str):
            text = child.strip()
            if text:
                content.append(text)
        elif child.name:
            # Handle specific formatting elements
            if child.name == 'p':
                p = soup.new_tag('div')
                p['class'] = 'question-paragraph'
                p.string = child.get_text(strip=True)
                content.append(p)
            elif child.name == 'span':
                content.append(child.get_text(strip=True))
            elif child.name == 'br':
                content.append(soup.new_tag('br'))
            elif child.name == 'img':
                content.append(child)
            elif child.name in ['strong', 'b']:
                strong = soup.new_tag('strong')
                strong.string = child.get_text(strip=True)
                content.append(strong)
            elif child.name in ['em', 'i']:
                em = soup.new_tag('em')
                em.string = child.get_text(strip=True)
                content.append(em)
            else:
                # Fallback for arbitrary inline elements
                text = child.get_text(strip=True)
                if text:
                    content.append(text)
    
    return content



# ==============================================================================
# TABLE TO QUESTIONNAIRE CARD TRANSFORMER
# ==============================================================================

def table_to_questionnaire(soup, table):
    """
    Convert a SurveyCTO HTML table structure into a print-optimized questionnaire card flow.

    SurveyCTO renders survey forms as rigid HTML tables where each row is a question or section.
    This function decomposes the table rows (<tr>) into modular, vertical card containers:
      - Section Headers: Single-cell rows with colspan="3" or dark background colors become
        full-width section dividers (<div class="question-item section-header">).
      - 3-Column Question Rows:
          * Cell 0 (Variable/Field Name) -> Top badge (<div class="field-label">)
          * Cell 1 (Question Prompt & Relevance) -> Formatted prompt (<div class="question-text">)
          * Cell 2 (Choices / Answers) -> Nested table parsed into choice items (<div class="choice-item">)
            or an empty notebook-style fill-in line block (<div class="answer-placeholder">).
      - 2-Column Rows: Rendered as compact Field / Value pairs.
      - 1-Column Rows: Rendered as highlighted notes (<div class="note-cell">).

    Args:
        soup (bs4.BeautifulSoup): Root BeautifulSoup instance used for creating replacement tags.
        table (bs4.element.Tag): The HTML <table> element to transform.

    Returns:
        bs4.element.Tag: A <div> container enclosing the transformed questionnaire cards.
    """
    container = soup.new_tag('div')
    
    # Extract all rows from the table
    rows = table.find_all('tr')
    if not rows:
        return container
    
    # Check for <thead> to extract column headers if present
    thead = table.find('thead')
    headers = []
    if thead:
        header_row = thead.find('tr')
        if header_row:
            headers = [th.get_text(strip=True) for th in header_row.find_all(['th', 'td'])]
    
    # Determine the number of columns from the first data row
    first_cells = rows[0].find_all(['td', 'th'])
    num_cols = len(first_cells)
    
    # If thead was present, skip the header row from row iteration
    data_rows = rows[1:] if thead else rows
    
    for row in data_rows:
        cells = row.find_all(['td', 'th'])
        if not cells:
            continue
        
        # ----------------------------------------------------------------------
        # Case 1: Section Header Row
        # ----------------------------------------------------------------------
        # SurveyCTO marks section breaks with a single <td> spanning all columns (colspan=3)
        # or having a hardcoded background-color.
        is_section_header = False
        if len(cells) == 1:
            cell = cells[0]
            if cell.get('colspan') == '3' or 'background-color' in str(cell.get('style', '')):
                is_section_header = True
        
        if is_section_header:
            section = soup.new_tag('div')
            section['class'] = 'question-item section-header'
            section.string = cells[0].get_text(strip=True)
            container.append(section)
            continue
        
        # ----------------------------------------------------------------------
        # Case 2: Standard Question Card
        # ----------------------------------------------------------------------
        question_item = soup.new_tag('div')
        question_item['class'] = 'question-item'
        
        # Subcase A: Standard 3-Column Layout [Field Name | Prompt | Choices/Answer]
        if num_cols >= 3:
            field_cell = cells[0] if len(cells) > 0 else None
            question_cell = cells[1] if len(cells) > 1 else None
            answer_cell = cells[2] if len(cells) > 2 else None
            
            # 1. Field Label Badge (Variable Name)
            if field_cell:
                field_text = field_cell.get_text(strip=True)
                if field_text:
                    field_label = soup.new_tag('div')
                    field_label['class'] = 'field-label'
                    field_label.string = field_text
                    question_item.append(field_label)
            
            # 2. Question Text Block
            if question_cell:
                question_text = soup.new_tag('div')
                question_text['class'] = 'question-text'
                
                # Parse rich formatting (paragraphs, bold, linebreaks, images)
                content = parse_cell_content(question_cell, soup)
                for item in content:
                    if isinstance(item, str):
                        question_text.append(item)
                    else:
                        question_text.append(item)
                
                question_item.append(question_text)
            
            # 3. Answer Area (Choices or Hand-Filling Placeholder)
            answer_area = soup.new_tag('div')
            answer_area['class'] = 'answer-area'
            
            if answer_cell:
                # Check if this question has a nested choices table (select_one, select_multiple)
                nested_table = answer_cell.find('table')
                if nested_table:
                    # Deconstruct nested choice rows into a vertical choice list
                    choice_list = soup.new_tag('div')
                    choice_list['class'] = 'choice-list'
                    
                    choice_rows = nested_table.find_all('tr')
                    for choice_row in choice_rows:
                        choice_cells = choice_row.find_all(['td', 'th'])
                        if len(choice_cells) >= 2:
                            choice_item = soup.new_tag('div')
                            choice_item['class'] = 'choice-item'
                            
                            # Choice Number/Code Badge
                            choice_num = soup.new_tag('span')
                            choice_num['class'] = 'choice-number'
                            choice_num_text = choice_cells[1].get_text(strip=True) if len(choice_cells) > 1 else ''
                            choice_num.string = choice_num_text
                            choice_item.append(choice_num)
                            
                            # Choice Label Text
                            choice_text = soup.new_tag('span')
                            choice_text['class'] = 'choice-text'
                            choice_text_content = choice_cells[2].get_text(strip=True) if len(choice_cells) > 2 else ''
                            choice_text.string = choice_text_content
                            choice_item.append(choice_text)
                            
                            choice_list.append(choice_item)
                    
                    answer_area.append(choice_list)
                else:
                    # Non-choice questions (text, integer, decimal): render ruled writing lines
                    placeholder = soup.new_tag('div')
                    placeholder['class'] = 'answer-placeholder'
                    answer_area.append(placeholder)
            else:
                placeholder = soup.new_tag('div')
                placeholder['class'] = 'answer-placeholder'
                answer_area.append(placeholder)
            
            question_item.append(answer_area)
            
        # Subcase B: 2-Column Layout [Field Name | Value]
        elif num_cols == 2:
            field_cell = cells[0]
            value_cell = cells[1] if len(cells) > 1 else None
            
            field_label = soup.new_tag('div')
            field_label['class'] = 'field-label'
            field_label.string = field_cell.get_text(strip=True)
            question_item.append(field_label)
            
            if value_cell:
                value_div = soup.new_tag('div')
                value_div['class'] = 'answer-area'
                value_div.string = value_cell.get_text(strip=True)
                question_item.append(value_div)
        
        # Subcase C: 1-Column Layout [Note / Instruction]
        elif num_cols == 1:
            content_text = cells[0].get_text(strip=True)
            if content_text:
                note_div = soup.new_tag('div')
                note_div['class'] = 'note-cell'
                note_div.string = content_text
                question_item.append(note_div)
        
        container.append(question_item)
    
    return container


# ==============================================================================
# NESTED TABLE RESOLUTION
# ==============================================================================

def convert_nested_tables(soup, element):
    """
    Recursively find and transform nested HTML tables into questionnaire cards.

    Performs a bottom-up depth-first conversion so that innermost choice tables
    are transformed before parent wrapper tables are parsed.

    Args:
        soup (bs4.BeautifulSoup): Root BeautifulSoup instance.
        element (bs4.element.Tag): The parent DOM container to search.
    """
    tables = element.find_all('table', recursive=False)
    for table in tables:
        convert_nested_tables(soup, table)
        questionnaire_div = table_to_questionnaire(soup, table)
        table.replace_with(questionnaire_div)


# ==============================================================================
# DOCUMENT-LEVEL HTML RESPONSIVE & PRINT CONVERSION
# ==============================================================================

def make_html_responsive(html_content):
    """
    Transform a complete SurveyCTO HTML document into a print-ready questionnaire.

    Orchestrates the full pipeline:
      1. Head Configuration: Removes legacy stylesheets and injects PRINT_STYLES and BASE_STYLES.
      2. Container Reset: Converts fixed-width container divs (<div style="width: ...">)
         into flexible <div class="questionnaire-container"> containers.
      3. Table Transformation: Converts the primary survey table (and any subsidiary tables)
         into sequential card items.
      4. Header Prominence: Promotes generic <h4> survey titles into prominent
         <div class="qn-header"><h1>...</h1></div> headers.

    Args:
        html_content (str): The raw SurveyCTO HTML document content.

    Returns:
        str: Transformed HTML document string.
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # --------------------------------------------------------------------------
    # Step 1: Head Configuration & Style Injection
    # --------------------------------------------------------------------------
    head = soup.find('head')
    if head:
        # Clear existing SurveyCTO inline styles to avoid specificity conflicts
        style_tags = head.find_all('style')
        for style in style_tags:
            style.extract()
        
        # Inject print and screen stylesheets
        new_style = soup.new_tag('style')
        new_style.string = BASE_STYLES + PRINT_STYLES
        head.append(new_style)
    
    # --------------------------------------------------------------------------
    # Step 2: Unconstrain Container Widths
    # --------------------------------------------------------------------------
    main_div = soup.find('div', style=lambda x: x and ('width' in str(x).lower()))
    if main_div:
        main_div['class'] = 'questionnaire-container'
        if main_div.get('style'):
            del main_div['style']
    
    # --------------------------------------------------------------------------
    # Step 3: Transform Main and Subsidiary Tables
    # --------------------------------------------------------------------------
    main_table = soup.find('table', class_=lambda x: x and 'table' in str(x))
    if main_table:
        questionnaire_div = table_to_questionnaire(soup, main_table)
        main_table.replace_with(questionnaire_div)
    
    # Clean up any remaining loose tables
    all_tables = soup.find_all('table')
    for table in all_tables:
        questionnaire_div = table_to_questionnaire(soup, table)
        table.replace_with(questionnaire_div)
    
    # --------------------------------------------------------------------------
    # Step 4: Elevate Survey Title
    # --------------------------------------------------------------------------
    title = soup.find('h4')
    if title:
        title['class'] = 'qn-header'
        title.name = 'div'
        h1 = soup.new_tag('h1')
        h1.string = title.get_text(strip=True)
        title.clear()
        title.append(h1)
    
    return str(soup)


# ==============================================================================
# DIRECTORY BATCH PROCESSOR & CLI ENTRYPOINT
# ==============================================================================

def process_directory():
    """
    Batch-process all HTML codebook files in the current working directory.

    Finds all files matching '*.html' or '*.htm', applies the print questionnaire
    transformation, and saves the updated content in-place.
    """
    current_dir = os.getcwd()
    html_files = [f for f in os.listdir(current_dir) if f.endswith('.html') or f.endswith('.htm')]
    
    if not html_files:
        print("No HTML files found in the current directory.")
        return
    
    print(f"Found {len(html_files)} HTML file(s) to process.")
    
    for filename in html_files:
        file_path = os.path.join(current_dir, filename)
        print(f"\nProcessing {filename}...")
        
        try:
            # Read raw HTML file
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
            
            # Transform to print-optimized questionnaire layout
            responsive_content = make_html_responsive(content)
            
            # Write modified content back in-place
            with open(file_path, 'w', encoding='utf-8') as file:
                file.write(responsive_content)
            
            print(f"  Successfully converted to print-optimized questionnaire!")
            
        except Exception as e:
            print(f"  Error processing {filename}: {str(e)}")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    process_directory()

