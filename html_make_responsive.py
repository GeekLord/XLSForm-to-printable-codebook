"""
html_make_responsive.py - SurveyCTO Codebook HTML Post-Processor & PDF Generator

A legacy post-processing utility designed to transform default, rigid SurveyCTO
HTML codebook exports into mobile-responsive, print-ready documents and landscape PDFs.

Core Functions:
- Strips hardcoded inline width styles (e.g. `width: 1000px`) and injects responsive container classes.
- Normalizes section header inline colors (`#707070`, `#8c8c8c`) into theme-managed dark header bands (`#2B2D42`).
- Wraps tables inside `.table-responsive` viewports to prevent horizontal overflow on smaller screens.
- Resolves nested choice table width clipping (`width: 100%` on table cells).
- Injects print stylesheets tuned for landscape A4 output.
- Renders pixel-accurate landscape PDFs using Playwright Chromium.
"""

import os
import asyncio
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright

# ==============================================================================
# CSS & PRINT STYLES
# ==============================================================================
# Responsive + Print-optimized styles for SurveyCTO codebook
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
    font-size: 13px;
    line-height: 18px;
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
    font-size: 18px;
    margin: 10px 0;
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
    padding: 6px 8px;
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
    font-size: 12px;
    font-weight: bold;
    color: #fff;
}

.fieldCell {
    font-weight: bold;
    color: #222;
    width: 22%;
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
    overflow-wrap: break-word;
    word-break: break-word;
}

.answerCell {
    width: 28%;
    overflow-wrap: break-word;
    word-break: break-word;
}

tr.gray td {
    background: #fff;
}

td .table {
    margin: 0;
    font-size: 12px;
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
        font-size: 15px;
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
        font-size: 11px;
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
        font-size: 11px;
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
        font-size: 8.5pt;
        line-height: 12pt;
        padding: 0;
        background: #fff;
    }
    
    h4 {
        font-size: 13pt;
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
        padding: 3pt 5pt;
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
        font-size: 9pt;
    }
    
    td[colspan="3"],
    .section-header-dark {
        background: #f0f4f8 !important;
        color: #000 !important;
        font-weight: bold;
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
}

img.questionPrompt { 
    max-width: 100%; 
    padding: 2px; 
    border: 1px solid #ddd; 
    display: block; 
    margin: 4px 0; 
}

img.choicePrompt { 
    max-width: 120px; 
    padding: 2px; 
    border: 1px solid #ddd; 
}

.required { 
    color: red; 
    font-style: italic; 
}

.relevance {
    font-size: 11px;
    font-style: italic;
    color: #2b7013;
}

.hint {
    color: #1864ab;
    font-size: 11px;
}

.metadata-tag {
    font-size: 11px;
    color: #495057;
    background: #f1f3f5;
    border-left: 3px solid #6c757d;
    padding: 2px 5px;
    margin-top: 3px;
    border-radius: 2px;
}

.calculation {
    font-size: 11px;
    color: #004085;
    background: #e7f5ff;
    border-left: 3px solid #1864ab;
    padding: 3px 6px;
    margin-top: 3px;
    word-break: break-all;
    border-radius: 2px;
}

.calc-desc {
    font-size: 10px;
    font-style: italic;
    color: #495057;
    margin-top: 1px;
}
"""


# ==============================================================================
# DOM MANIPULATION & RESPONSIVE CONVERSION
# ==============================================================================

def make_html_responsive(html_content):
    """
    Transform raw SurveyCTO HTML exports into responsive, print-ready documents.

    SurveyCTO's native codebook exports come with rigid styling limitations:
      1. Fixed-width wrappers (e.g., <div style="width: 1000px">) causing horizontal
         scrolling on tablets and mobile screens.
      2. Dark gray section header cells with inline styles (#707070, #8c8c8c, rgb(112, ...))
         that create poor contrast and clash with print styles.
      3. Missing viewport meta tags required for modern mobile scaling.
      4. Naked tables lacking responsive scroll wrappers.
      5. Nested choice tables with 'width: 100%' inline styles that overflow their cells.

    This function parses the HTML via BeautifulSoup, sanitizes inline styles, injects
    modern responsive CSS, wraps tables in scroll containers, and normalizes choice cells.

    Args:
        html_content (str): The raw HTML string exported from SurveyCTO.

    Returns:
        str: Transformed HTML string with injected responsive styles and markup.
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # --------------------------------------------------------------------------
    # Step 1: Strip Problematic Inline Header Colors
    # --------------------------------------------------------------------------
    # SurveyCTO embeds hardcoded inline styles on group/section header <td> tags
    # (#707070, #8c8c8c, rgb(112, 112, 112), etc.). We strip 'background-color' and
    # 'color' from these tags and attach the semantic CSS class 'section-header-dark'
    # so the new stylesheet can control contrast cleanly across screen and print media.
    for td in soup.find_all('td', style=True):
        style = td.get('style', '')
        style_lower = style.lower()
        if ('#707070' in style or '#8c8c8c' in style or 
            'background-color: rgb(112' in style or 'background-color: rgb(140' in style or
            'background-color:#707070' in style or 'background-color:#8c8c8c' in style):
            new_styles = []
            for part in style.split(';'):
                part_stripped = part.strip()
                if not part_stripped:
                    continue
                prop_name = part_stripped.split(':')[0].strip().lower()
                # Remove hardcoded color properties
                if prop_name == 'background-color' or prop_name == 'color':
                    continue
                new_styles.append(part_stripped)
            if new_styles:
                td['style'] = '; '.join(new_styles)
            else:
                del td['style']
            td['class'] = td.get('class', []) + ['section-header-dark']
    
    # --------------------------------------------------------------------------
    # Step 2: Ensure <head> and Mobile Viewport Meta Tag Exist
    # --------------------------------------------------------------------------
    head = soup.find('head')
    if not head:
        head = soup.new_tag('head')
        soup.html.insert(0, head)
    
    viewport = head.find('meta', attrs={'name': 'viewport'})
    if not viewport:
        viewport = soup.new_tag('meta', attrs={
            'name': 'viewport',
            'content': 'width=device-width, initial-scale=1.0'
        })
        head.insert(0, viewport)
    
    # --------------------------------------------------------------------------
    # Step 3: Inject Modern Responsive Stylesheet
    # --------------------------------------------------------------------------
    style_tag = head.find('style')
    if style_tag:
        style_tag.string = RESPONSIVE_STYLES
    else:
        style_tag = soup.new_tag('style')
        style_tag.string = RESPONSIVE_STYLES
        head.append(style_tag)
    
    # --------------------------------------------------------------------------
    # Step 4: Unconstrain Fixed-Width Containers
    # --------------------------------------------------------------------------
    # SurveyCTO wraps the entire codebook table in <div style="width: 1000px">.
    # Convert this to our flexible .codebook-container class.
    container = soup.find('div', style=lambda x: x and 'width: 1000px' in str(x))
    if container:
        container['class'] = 'codebook-container'
        del container['style']
    
    # --------------------------------------------------------------------------
    # Step 5: Wrap Main Table in Responsive Overflow Wrapper
    # --------------------------------------------------------------------------
    # On mobile viewports, large multi-column tables cause document horizontal blowout.
    # Wrapping the table in a .table-responsive div enables smooth horizontal scrolling.
    main_table = soup.find('table', class_='table')
    if main_table:
        parent = main_table.parent
        if parent.name != 'div' or 'table-responsive' not in str(parent.get('class', [])):
            wrapper = soup.new_tag('div', class_='table-responsive')
            main_table.wrap(wrapper)
            
    # --------------------------------------------------------------------------
    # Step 6: Fix Choice Cell Width Overflow
    # --------------------------------------------------------------------------
    # Nested choice tables often specify 'width: 100%' alongside fixed padding,
    # causing the third column to exceed the table boundary in Chromium/WebKit.
    for td in soup.find_all('td', style=True):
        style = td.get('style', '')
        if 'width: 100%' in style and 'padding-left: 3px' in style and 'padding-right: 3px' in style:
            # Strip width: 100% and add a subtle left border to delineate choice options
            new_style = style.replace('width: 100%;', '').replace('width:100%;', '').replace('width: 100%', '').strip()
            if 'border-left' not in new_style:
                new_style = new_style.rstrip(';') + '; border-left: 1px solid #ddd;'
            td['style'] = new_style
    
    return str(soup)


# ==============================================================================
# PLAYWRIGHT CHROMIUM PDF GENERATOR
# ==============================================================================

async def generate_pdf(html_path, pdf_path):
    """
    Render an HTML codebook to a landscape A4 PDF using headless Chromium.

    Uses Playwright to spin up a headless browser instance, loads the local file,
    waits for network and web fonts to settle, and renders the PDF with full
    background graphics enabled.

    Args:
        html_path (str): Absolute or relative filesystem path to the HTML source.
        pdf_path (str): Destination filesystem path for the rendered PDF file.

    Returns:
        str: The path of the generated PDF file.
    """
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()
        
        # Load local HTML file and wait for all DOM assets to finish loading
        await page.goto(f'file:///{html_path}', wait_until='networkidle')
        
        # Buffer delay to ensure complex scripts (e.g. Indic/Devanagari fonts)
        # and CSS layout calculations settle before print rasterization
        await page.wait_for_timeout(1000)
        
        # Generate PDF in Landscape A4 format:
        # - print_background=True preserves table zebra striping and header colors.
        # - 10mm uniform margins provide balanced printable area.
        await page.pdf(
            path=pdf_path,
            format='A4',
            landscape=True,
            margin={'top': '10mm', 'right': '10mm', 'bottom': '10mm', 'left': '10mm'},
            print_background=True
        )
        
        await browser.close()
        return pdf_path


# ==============================================================================
# DIRECTORY BATCH PROCESSOR & CLI ENTRYPOINT
# ==============================================================================

def process_directory():
    """
    Batch-process all HTML codebook files in the current working directory.

    For each '*.html' file discovered:
      1. Reads raw HTML content.
      2. Applies make_html_responsive() transformations.
      3. Overwrites the HTML file in-place with the responsive markup.
      4. Triggers Playwright via asyncio to compile a matching '*.pdf' document.
    """
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
            
            # Save responsive HTML in-place
            with open(file_path, 'w', encoding='utf-8') as file:
                file.write(responsive_content)
            
            print(f"  Made responsive!")
            
            # Generate companion PDF file
            pdf_filename = filename.replace('.html', '.pdf')
            pdf_path = os.path.join(current_dir, pdf_filename)
            
            print(f"  Generating PDF: {pdf_filename}...")
            asyncio.run(generate_pdf(file_path, pdf_path))
            print(f"  PDF created!")
            
        except Exception as e:
            print(f"  Error: {str(e)}")


if __name__ == "__main__":
    process_directory()