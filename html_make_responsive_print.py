import os
import re
from bs4 import BeautifulSoup
from bs4.element import Tag

# Print-optimized CSS for SurveyCTO questionnaire printing
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

# Minimal base styles
BASE_STYLES = """
.text-left { text-align: left; }
.text-right { text-align: right; }
.text-center { text-align: center; }
.clearfix { overflow: auto; }
"""


def extract_text_content(element):
    """Extract text content from element, handling nested elements."""
    if element is None:
        return ""
    if isinstance(element, str):
        return element
    return element.get_text(strip=True)


def parse_cell_content(cell, soup):
    """Parse cell content and return list of elements to append."""
    content = []
    
    # Process all children
    for child in cell.children:
        if isinstance(child, str):
            text = child.strip()
            if text:
                content.append(text)
        elif child.name:
            # Handle specific elements
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
                # For other elements, just get text
                text = child.get_text(strip=True)
                if text:
                    content.append(text)
    
    return content


def table_to_questionnaire(soup, table):
    """
    Convert SurveyCTO table structure to print-optimized questionnaire format.
    """
    container = soup.new_tag('div')
    
    # Get all rows
    rows = table.find_all('tr')
    if not rows:
        return container
    
    # Check for thead to get headers
    thead = table.find('thead')
    headers = []
    if thead:
        header_row = thead.find('tr')
        if header_row:
            headers = [th.get_text(strip=True) for th in header_row.find_all(['th', 'td'])]
    
    # Determine number of columns from first row
    first_cells = rows[0].find_all(['td', 'th'])
    num_cols = len(first_cells)
    
    # Skip header row in data rows if thead exists
    data_rows = rows[1:] if thead else rows
    
    for row in data_rows:
        cells = row.find_all(['td', 'th'])
        if not cells:
            continue
        
        # Check if this is a section header row (single cell with colspan or gray background)
        is_section_header = False
        if len(cells) == 1:
            cell = cells[0]
            if cell.get('colspan') == '3' or 'background-color' in str(cell.get('style', '')):
                is_section_header = True
        
        if is_section_header:
            # Create section header
            section = soup.new_tag('div')
            section['class'] = 'question-item section-header'
            section.string = cells[0].get_text(strip=True)
            container.append(section)
            continue
        
        # Regular question row
        question_item = soup.new_tag('div')
        question_item['class'] = 'question-item'
        
        # Extract field, question, answer based on column structure
        if num_cols >= 3:
            # Standard Field/Question/Answer layout
            field_cell = cells[0] if len(cells) > 0 else None
            question_cell = cells[1] if len(cells) > 1 else None
            answer_cell = cells[2] if len(cells) > 2 else None
            
            # Field label
            if field_cell:
                field_text = field_cell.get_text(strip=True)
                if field_text:
                    field_label = soup.new_tag('div')
                    field_label['class'] = 'field-label'
                    field_label.string = field_text
                    question_item.append(field_label)
            
            # Question text
            if question_cell:
                question_text = soup.new_tag('div')
                question_text['class'] = 'question-text'
                
                # Parse question content
                content = parse_cell_content(question_cell, soup)
                for item in content:
                    if isinstance(item, str):
                        question_text.append(item)
                    else:
                        question_text.append(item)
                
                question_item.append(question_text)
            
            # Answer area
            answer_area = soup.new_tag('div')
            answer_area['class'] = 'answer-area'
            
            if answer_cell:
                # Check for nested table (choices/options)
                nested_table = answer_cell.find('table')
                if nested_table:
                    # Convert nested table to choice list
                    choice_list = soup.new_tag('div')
                    choice_list['class'] = 'choice-list'
                    
                    choice_rows = nested_table.find_all('tr')
                    for choice_row in choice_rows:
                        choice_cells = choice_row.find_all(['td', 'th'])
                        if len(choice_cells) >= 2:
                            choice_item = soup.new_tag('div')
                            choice_item['class'] = 'choice-item'
                            
                            # Choice number
                            choice_num = soup.new_tag('span')
                            choice_num['class'] = 'choice-number'
                            choice_num_text = choice_cells[1].get_text(strip=True) if len(choice_cells) > 1 else ''
                            choice_num.string = choice_num_text
                            choice_item.append(choice_num)
                            
                            # Choice text
                            choice_text = soup.new_tag('span')
                            choice_text['class'] = 'choice-text'
                            choice_text_content = choice_cells[2].get_text(strip=True) if len(choice_cells) > 2 else ''
                            choice_text.string = choice_text_content
                            choice_item.append(choice_text)
                            
                            choice_list.append(choice_item)
                    
                    answer_area.append(choice_list)
                else:
                    # Empty answer area with placeholder
                    placeholder = soup.new_tag('div')
                    placeholder['class'] = 'answer-placeholder'
                    answer_area.append(placeholder)
            else:
                # Empty answer area
                placeholder = soup.new_tag('div')
                placeholder['class'] = 'answer-placeholder'
                answer_area.append(placeholder)
            
            question_item.append(answer_area)
            
        elif num_cols == 2:
            # Simple label/value format
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
        
        elif num_cols == 1:
            # Single cell - treat as note or simple content
            content_text = cells[0].get_text(strip=True)
            if content_text:
                note_div = soup.new_tag('div')
                note_div['class'] = 'note-cell'
                note_div.string = content_text
                question_item.append(note_div)
        
        container.append(question_item)
    
    return container


def convert_nested_tables(soup, element):
    """Recursively convert nested tables."""
    tables = element.find_all('table', recursive=False)
    for table in tables:
        convert_nested_tables(soup, table)
        questionnaire_div = table_to_questionnaire(soup, table)
        table.replace_with(questionnaire_div)


def make_html_responsive(html_content):
    """Transform HTML to print-optimized questionnaire format."""
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Add viewport meta for screen preview
    head = soup.find('head')
    if head:
        # Clear existing styles and add print-optimized CSS
        style_tags = head.find_all('style')
        for style in style_tags:
            style.extract()
        
        new_style = soup.new_tag('style')
        new_style.string = BASE_STYLES + PRINT_STYLES
        head.append(new_style)
    
    # Find main container div
    main_div = soup.find('div', style=lambda x: x and ('width' in str(x).lower()))
    if main_div:
        # Convert to questionnaire container
        main_div['class'] = 'questionnaire-container'
        if main_div.get('style'):
            del main_div['style']
    
    # Convert main table to questionnaire format
    main_table = soup.find('table', class_=lambda x: x and 'table' in str(x))
    if main_table:
        questionnaire_div = table_to_questionnaire(soup, main_table)
        main_table.replace_with(questionnaire_div)
    
    # Handle any remaining nested tables
    all_tables = soup.find_all('table')
    for table in all_tables:
        questionnaire_div = table_to_questionnaire(soup, table)
        table.replace_with(questionnaire_div)
    
    # Make title prominent
    title = soup.find('h4')
    if title:
        title['class'] = 'qn-header'
        title.name = 'div'
        h1 = soup.new_tag('h1')
        h1.string = title.get_text(strip=True)
        title.clear()
        title.append(h1)
    
    return str(soup)


def process_directory():
    """Process all HTML files in the current directory."""
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
            # Read the file
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
            
            # Transform to print-optimized format
            responsive_content = make_html_responsive(content)
            
            # Write the modified content back
            with open(file_path, 'w', encoding='utf-8') as file:
                file.write(responsive_content)
            
            print(f"  Successfully converted to print-optimized questionnaire!")
            
        except Exception as e:
            print(f"  Error processing {filename}: {str(e)}")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    process_directory()
