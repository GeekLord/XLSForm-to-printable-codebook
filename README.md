# XLSForm to Printable Codebook & PDF Generator

This repository provides tools and scripts to convert XLSForm spreadsheets (`.xlsx`) from ODK, SurveyCTO, and KoboToolbox platforms into responsive, print-ready HTML questionnaires and high-quality A4 PDF codebooks.

## Project Overview

Survey forms created for ODK, SurveyCTO, or KoboToolbox are authored as XLSForm spreadsheets containing `survey`, `choices`, and `settings` worksheets. Previously, generating a printable codebook required uploading the XLSForm to a survey server, exporting the printable HTML file, and running post-processing scripts (`html_make_responsive.py` or `html_make_responsive_print.py`) to inject responsive CSS and render PDFs.

The core objective of this project is to implement a direct Python pipeline (`xlsform_to_printable.py`) that parses `.xlsx` XLSForm files directly, handles platform variations across ODK, SurveyCTO, and KoboToolbox, auto-detects all form languages (including un-designated labels), and renders print-ready HTML and PDF codebooks for all detected languages without requiring any server upload or export step.

## Multi-Platform & Multi-Language Support

- **Cross-Platform Compatibility**: Automatically normalizes syntax differences between ODK, SurveyCTO, and KoboToolbox (e.g. `begin_group` vs `begin group`, `media::image` vs `image`, `select_one_from_file`, `or_other`, and platform-specific metadata).
- **Automated Multi-Language Processing**: Automatically scans column headers for language tags (`label::English (en)`, `label::Hindi (hi)`, `label::Spanish`, etc.) as well as un-tagged `label` columns. If multiple languages are present, codebooks are generated for all detected languages automatically.

## Repository Structure

- `xlsform_to_printable.py` — Standalone script to directly parse XLSForm `.xlsx` files and output responsive HTML and A4 PDF codebooks across languages and platforms.
- `html_make_responsive.py` — Existing script that processes SurveyCTO-exported HTML codebooks, applies responsive and print CSS, and generates A4 landscape PDFs using Playwright Chromium.
- `html_make_responsive_print.py` — Alternative script that converts SurveyCTO HTML tables into a card-style `.question-item` layout for portrait A4 printing.
- `backup_html_make_responsive.py` — Frozen reference backup of the HTML responsive transformation script.
- `xlsform_to_printable_plan.md` — Detailed technical architecture, XLSForm element mapping, and implementation roadmap for the direct XLSForm parser.
- `PA_KAP_Endline_CR_Programme_UP_Bihar.xlsx` — Sample XLSForm spreadsheet for the KAP Endline survey.
- `PA_Panel_Diary_CR_Programme_UP_Bihar.xlsx` — Sample XLSForm spreadsheet for the Panel Diary survey.
- `*_English.html` / `*_Hindi.html` — Sample exported codebooks for English and Hindi form variations.

## Prerequisites and Setup

1. Python 3.8+ installed on your system.
2. Install required Python packages:
   ```bash
   pip install openpyxl beautifulsoup4 playwright
   ```
3. Install Playwright browser binaries (required for PDF generation):
   ```bash
   playwright install chromium
   ```

## Usage Instructions

### Direct XLSForm Processing (New Workflow)
Run `xlsform_to_printable.py` directly against XLSForm Excel files to generate codebooks for all detected languages:
```bash
python xlsform_to_printable.py --input PA_KAP_Endline_CR_Programme_UP_Bihar.xlsx --pdf
```
To target a specific language explicitly:
```bash
python xlsform_to_printable.py --input PA_KAP_Endline_CR_Programme_UP_Bihar.xlsx --lang English --pdf
```

### Legacy HTML Post-Processing Workflow
To format existing SurveyCTO HTML exports and render landscape A4 PDFs:
```bash
python html_make_responsive.py
```
To transform existing HTML exports into portrait card layouts:
```bash
python html_make_responsive_print.py
```

## Key Technical Specifications & Styling Rules

- **Encoding & Typography**: All files are read and written using `utf-8` encoding. The primary font stack begins with `"Arial Unicode MS"` followed by Arial to ensure seamless rendering of Devanagari (Hindi) script and international characters.
- **Section Headers**: Form groups (`begin group` / `begin_group`) are styled as full-width section headers with dark background `#1a1a1a` and white text, configured with `-webkit-print-color-adjust: exact` for background retention during printing.
- **Choice Lists**: Single-choice (`select_one`) and multiple-choice (`select_multiple`) questions display option values and labels in clean nested choice tables inside the Answer column.
- **Form Metadata & Logic**: Field names, required indicators (`(required)`), hints (styled in blue), relevance expressions (styled in green italic), and constraints are parsed and rendered preserving field visibility rules.
- **Print Optimization**: Configured with `@media print` rules, `page-break-inside: avoid` on question rows, and exact color preservation for A4 landscape and portrait printing.

## Markdown Guidelines for Contributor Documentation

All Markdown documentation in this project strictly follows a single-line paragraph rule. Paragraphs and bullet points must be written as a single unwrapped line without manual column-limit line breaks.
