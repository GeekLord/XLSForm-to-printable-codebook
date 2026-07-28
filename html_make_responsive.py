import os
import asyncio
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright

# Responsive + Print-optimized styles for SurveyCTO codebook
RESPONSIVE_STYLES = """
.spacer {
    display: none !important;
    width: 0 !important;
    padding: 0 !important;
    margin: 0 !important;
}

/* Viewport and base */
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

/* Responsive container */
.codebook-container {
    width: 100%;
    max-width: 1200px;
    margin: auto;
    padding: 0 15px;
}

h4 {
    text-align: center;
    font-size: 18px;
    margin: 15px 0;
    padding-bottom: 10px;
    border-bottom: 2px solid #333;
}

/* Table base styles */
table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 20px;
}

th, td {
    padding: 8px;
    text-align: left;
    border: 1px solid #ddd;
    vertical-align: top;
}

/* Section headers - comprehensive selectors for all variations */
td[colspan="3"],
td[style*="background-color: #707070"],
td[style*="background-color: #8C8C8C"],
td[style*="background-color:#707070"],
td[style*="background-color:#8C8C8C"],
td[style*="background-color: rgb(112, 112, 112)"],
td[style*="background-color: rgb(140, 140, 140)"],
.section-header-dark,
tr.gray td.section-header-dark {
    background: #1a1a1a !important;
    color: #fff !important;
    font-weight: bold;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
}

/* Also target any td with color white that looks like a header */
td[style*="color: #FFFFFF"],
td[style*="color:#FFFFFF"],
td[style*="color: rgb(255, 255, 255)"] {
    background: #1a1a1a !important;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
}

/* Header styles */
.headerCell {
    background: #707070;
    color: #fff;
}

.headerCell h6 {
    margin: 0;
    font-size: 12px;
    font-weight: bold;
}

/* Field column */
.fieldCell {
    font-weight: bold;
    color: #333;
    width: 20%;
    white-space: normal;
    word-break: break-word;
}

.fieldCell .required {
    color: #c00;
    font-weight: normal;
}

/* Question column */
.questionCell {
    width: 60%;
}

/* Answer column */
.answerCell {
    width: 20%;
}

/* Section headers (gray rows) */
tr.gray td {
    background: #fff;
}

/* Nested choice tables */
td .table {
    margin: 0;
    font-size: 13px;
}

td .table td {
    padding: 4px;
    border: none;
    border-bottom: 1px dotted #ccc;
}




.response-note-cell {
    width: 10px;
}

/* Responsive breakpoints */
@media screen and (max-width: 768px) {
    body {
        font-size: 14px;
        line-height: 22px;
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
        padding: 6px;
    }
    
    .headerCell h6 {
        font-size: 11px;
    }
}

@media screen and (max-width: 480px) {
    body {
        font-size: 13px;
        line-height: 20px;
        padding: 5px;
    }
    
    h4 {
        font-size: 14px;
    }
    
    th, td {
        padding: 4px;
        font-size: 12px;
    }
    
    .fieldCell {
        width: 25%;
    }
    
    .questionCell {
        width: 50%;
    }
    
    .answerCell {
        width: 25%;
    }
}

/* Print styles */
@media print {
    @page { size: A4 landscape; margin: 10mm; }
    @page :first { margin-top: 8mm; }
    
    body {
        font-size: 9pt;
        line-height: 14pt;
        padding: 0;
        background: #fff;
    }
    
    h4 {
        font-size: 14pt;
        margin: 10px 0;
        border-bottom: 1pt solid #333;
    }
    
    table {
        page-break-inside: auto;
    }
    
    tr {
        page-break-inside: avoid;
    }
    
    th, td {
        padding: 4pt 6pt;
        border: 0.5pt solid #999;
    }
    
    .headerCell {
        background: #1a1a1a !important;
        color: #fff !important;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
    }
    
    td[colspan="3"],
    td[style*="background-color: #707070"],
    td[style*="background-color: #8C8C8C"],
    td[style*="background-color:#707070"],
    td[style*="background-color:#8C8C8C"],
    td[style*="background-color: rgb(112, 112, 112)"],
    td[style*="background-color: rgb(140, 140, 140)"],
    td[style*="color: #FFFFFF"],
    td[style*="color:#FFFFFF"],
    .section-header-dark,
    tr.gray td.section-header-dark {
        background: #1a1a1a !important;
        color: #fff !important;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
    }
    
    tr.gray td {
        background: #fff !important;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
    }
    .section-header-dark {
        background-color: #1a1a1a !important;
        color: #fff !important;
    }
}

/* User added styles */
img.questionPrompt { 
    max-width: 580px; 
    padding: 3px; 
    border: 1px solid #ddd; 
    display: block; 
    margin: 5px 0; 
}

img.choicePrompt { 
    max-width: 130px; 
    padding: 3px; 
    border: 1px solid #ddd; 
}

.required { 
    color: red; 
    font-style: italic; 
}

.relevance {
    font-size: 12px;
    font-style: italic;
    color: green;
}

.hint {
    color: blue;
    font-size: 13px;
}
"""


def make_html_responsive(html_content):
    """Add responsive CSS and structure to codebook HTML."""
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Remove problematic inline styles from section headers BEFORE adding CSS
    for td in soup.find_all('td', style=True):
        style = td.get('style', '')
        style_lower = style.lower()
        # Check if it has light gray background that's hard to read
        if ('#707070' in style or '#8C8C8C' in style or 
            'background-color: rgb(112' in style or 'background-color: rgb(140' in style or
            'background-color:#707070' in style or 'background-color:#8C8C8C' in style):
            # Parse and remove background-color and color properties
            new_styles = []
            for part in style.split(';'):
                part_stripped = part.strip()
                if not part_stripped:
                    continue
                prop_name = part_stripped.split(':')[0].strip().lower()
                # Skip background-color and color properties
                if prop_name == 'background-color' or prop_name == 'color':
                    continue
                new_styles.append(part_stripped)
            # Rebuild style without the background/color
            if new_styles:
                td['style'] = '; '.join(new_styles)
            else:
                del td['style']
            # Add a class so CSS can target it
            td['class'] = td.get('class', []) + ['section-header-dark']
    
    head = soup.find('head')
    if not head:
        head = soup.new_tag('head')
        soup.html.insert(0, head)
    
    # Add viewport meta
    viewport = head.find('meta', attrs={'name': 'viewport'})
    if not viewport:
        viewport = soup.new_tag('meta', attrs={
            'name': 'viewport',
            'content': 'width=device-width, initial-scale=1.0'
        })
        head.insert(0, viewport)
    
    # Replace or add styles
    style_tag = head.find('style')
    if style_tag:
        style_tag.string = RESPONSIVE_STYLES
    else:
        style_tag = soup.new_tag('style')
        style_tag.string = RESPONSIVE_STYLES
        head.append(style_tag)
    
    # Make container responsive
    container = soup.find('div', style=lambda x: x and 'width: 1000px' in str(x))
    if container:
        container['class'] = 'codebook-container'
        del container['style']
    
    # Add table-responsive wrapper to main table
    main_table = soup.find('table', class_='table')
    if main_table:
        parent = main_table.parent
        if parent.name != 'div' or 'table-responsive' not in str(parent.get('class', [])):
            wrapper = soup.new_tag('div', class_='table-responsive')
            main_table.wrap(wrapper)
            
    # Add border to answer column via inline style as requested
    for td in soup.find_all('td', style=True):
        style = td.get('style', '')
        # Match <td style="width: 100%; padding-left: 3px; padding-right: 3px;">
        if 'width: 100%' in style and 'padding-left: 3px' in style and 'padding-right: 3px' in style:
            # Append border-left style
            if 'border-left' not in style:
                td['style'] = style.rstrip(';') + '; border-left: 1px solid #999;'
    
    return str(soup)

async def generate_pdf(html_path, pdf_path):
    """Generate PDF from HTML using Playwright."""
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()
        
        # Load HTML file
        await page.goto(f'file:///{html_path}', wait_until='networkidle')
        
        # Wait for fonts to load
        await page.wait_for_timeout(1000)
        
        # Generate PDF in landscape A4
        await page.pdf(
            path=pdf_path,
            format='A4',
            landscape=True,
            margin={'top': '10mm', 'right': '10mm', 'bottom': '10mm', 'left': '10mm'},
            print_background=True
        )
        
        await browser.close()
        return pdf_path

def process_directory():
    """Process all HTML files and generate PDFs."""
    current_dir = os.getcwd()
    html_files = [f for f in os.listdir(current_dir) if f.endswith('.html')]
    
    if not html_files:
        print("No HTML files found.")
        return
    
    print(f"Found {len(html_files)} HTML file(s).")
    
    for filename in html_files:
        file_path = os.path.join(current_dir, filename)
        print(f"\nProcessing {filename}...")
        
        try:
            # Read and make responsive
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
            
            responsive_content = make_html_responsive(content)
            
            # Save responsive HTML
            with open(file_path, 'w', encoding='utf-8') as file:
                file.write(responsive_content)
            
            print(f"  Made responsive!")
            
            # Generate PDF
            pdf_filename = filename.replace('.html', '.pdf')
            pdf_path = os.path.join(current_dir, pdf_filename)
            
            print(f"  Generating PDF: {pdf_filename}...")
            asyncio.run(generate_pdf(file_path, pdf_path))
            print(f"  PDF created!")
            
        except Exception as e:
            print(f"  Error: {str(e)}")

if __name__ == "__main__":
    process_directory()