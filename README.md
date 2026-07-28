# XLSForm to Printable Codebook & PDF Generator

This repository provides tools and scripts to convert XLSForm spreadsheets (`.xlsx`) from ODK, SurveyCTO, and KoboToolbox platforms directly into responsive, print-ready HTML questionnaires and high-quality A4 PDF codebooks.

## Project Overview

Survey forms created for ODK, SurveyCTO, or KoboToolbox are authored as XLSForm spreadsheets containing `survey`, `choices`, and `settings` worksheets. Previously, generating a printable codebook required uploading the XLSForm to a survey server, exporting the printable HTML file, and running post-processing scripts (`html_make_responsive.py` or `html_make_responsive_print.py`) to inject responsive CSS and render PDFs.

The core objective of this project is to provide a direct Python pipeline (`xlsform_to_printable.py`) that parses `.xlsx` XLSForm files directly, handles platform variations across ODK, SurveyCTO, and KoboToolbox, auto-detects all form languages (including un-designated labels), displays platform-specific metadata explanations and categorized calculation formulas in the Question column, renders nested choice tables in the Answer column, and outputs print-ready HTML and PDF codebooks for all detected languages without requiring any server upload or export step.

## Key Features & Capabilities

- **Direct XLSForm Parsing**: Parses `.xlsx` files directly without requiring server uploads, processing `survey`, `choices`, and `settings` worksheets seamlessly.
- **Cross-Platform Compatibility**: Automatically normalizes syntax differences between ODK, SurveyCTO, and KoboToolbox (e.g. `begin_group` vs `begin group`, `end_group` vs `end group`, `begin_repeat` vs `begin repeat`, `select_one_from_file`, `select_multiple_from_file`, `or_other`, and platform-specific metadata).
- **Balanced Table Layout**: Enforces `table-layout: fixed; width: 100%` with balanced column proportions: **Field (22%)**, **Question (50%)**, and **Answer (28%)**. Uses `overflow-wrap: break-word` so variable names like `start-geopoint` or `phone_number` wrap naturally without single-character splits.
- **Eco-Friendly Print Styling**: Replaced solid black background bars with light-tinted headers (`#f0f4f8` with 1.5pt solid top/bottom borders for section headers and `#e9ecef` for table headers) in `@media print`. This saves up to **80% printer ink/toner** while preserving clear visual hierarchy.
- **Calculation Trimming & Height Reduction**: Complex formulas (such as `indexed-repeat()` or long `if()` conditions) are trimmed to 65 characters with an ellipsis (`...`) to prevent rows from expanding 5-6 lines vertically. Full formulas are preserved in HTML `title` tooltip attributes alongside categorized title badges and human-readable descriptions.
- **Choice Table Sizing & Overflow Protection**: Renders nested choice tables in the Answer column using a 2-column fixed layout (32px option code + auto-flex label with word wrapping). Removes layout clipping so long choice labels (e.g. respondent names or multi-line options) never extend past the right page margin.
- **Enhanced Typography & Readability**: Boosted print typography to `9.5pt` base font (`13.5pt` line height) and screen font to `14px` (`13.5px` cell font) for optimal legibility across desktop screens, tablets, and printed paper.
- **Automated Multi-Language Processing**: Automatically scans column headers for language tags (`label::English (en)`, `label::Odia (or)`, `label::Hindi (hi)`, etc.) as well as un-tagged `label` columns. Generates distinct codebook files for all detected languages automatically (e.g. `Survey_English.html`, `Survey_Odia.html`, `Survey_English.pdf`, `Survey_Odia.pdf`).
- **Platform Metadata Descriptions**: Automatically tags and explains platform-specific metadata variables (`start`, `end`, `today`, `deviceid`, `phonenumber`, `simserial`, `subscriberid`, `username`, `email`, `audit`, `text-audit`, `start-geopoint`, `caseid`, `caseread`, `casesave`) in the Question column with human-readable explanations.

## Supported Platform Metadata Reference

- **`start`**: Automated timestamp recorded when the survey session starts.
- **`end`**: Automated timestamp recorded when the survey session ends.
- **`today`**: Current date recorded when the survey is conducted.
- **`deviceid`**: Unique hardware identifier (IMEI/UUID) of the data collection device.
- **`phonenumber`**: Phone number associated with the SIM card in the device.
- **`simserial`**: Serial number of the SIM card installed in the device.
- **`subscriberid`**: IMSI subscriber identifier of the SIM card.
- **`username`**: Username of the logged-in field enumerator.
- **`email`**: Email address of the logged-in field enumerator.
- **`audit`**: User activity audit log file tracking timestamped interaction events.
- **`text-audit`**: Keystroke and text editing audit log file.
- **`start-geopoint`**: Initial GPS location captured upon opening the form.
- **`background-audio`**: Automated background audio recording captured during the survey.
- **`caseid`**: Case Management unique entity identifier (SurveyCTO).
- **`caseread`**: Case Management pre-loaded read field (SurveyCTO).
- **`casesave`**: Case Management saved status field (SurveyCTO).

## Supported Calculation Function Categories

- **SurveyCTO / ODK External Dataset Query (`pulldata`)**: Queries pre-loaded CSV datasets or server case data using unique lookup keys.
- **ODK / SurveyCTO Choice Label Lookup (`jr:choice-name`)**: Resolves raw choice option codes into human-readable label strings.
- **Timestamp Capture (`once(now())`, `now()`)**: Captures current date and time upon evaluation.
- **Date/Time Formatter (`format-date-time`)**: Formats raw timestamps into custom formatted date/time strings.
- **Duration / Datetime Math (`decimal-date-time`)**: Converts datetime to decimal days for duration math calculations.
- **Multi-Select & Repeat Functions (`selected-at`, `count-selected`)**: Extracts or counts selected response options in multiple-choice fields.
- **String Processing (`substr`, `concat`, `string-length`)**: Performs text manipulation, substring extraction, or string concatenation.
- **Conditional Logic (`if`)**: Evaluates conditional expressions.

## Repository Structure

- `xlsform_to_printable.py` — Primary standalone script to directly parse XLSForm `.xlsx` files and output responsive HTML and A4 PDF codebooks across languages and platforms.
- `html_make_responsive.py` — Post-processing script that formats existing HTML codebooks, applies responsive/print styles, and renders landscape A4 PDFs using Playwright Chromium.
- `html_make_responsive_print.py` — Alternative script that converts SurveyCTO HTML tables into a card-style layout for portrait A4 printing.
- `backup_html_make_responsive.py` — Frozen reference backup of the HTML responsive transformation script.
- `xlsform_to_printable_plan.md` — Detailed technical architecture, XLSForm element mapping, and implementation roadmap.
- `Brick Kiln Literacy RCT - Household Survey.xlsx` — Sample XLSForm spreadsheet containing multi-language (English and Odia) survey definitions.
- `*_English.html` / `*_Odia.html` — Sample generated codebooks for English and Odia form variations.
- `*_English.pdf` / `*_Odia.pdf` — Sample generated eco-friendly A4 landscape codebook PDFs.

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

### Direct XLSForm Processing (Recommended Workflow)
Run `xlsform_to_printable.py` directly against XLSForm Excel files to generate HTML and PDF codebooks for all detected languages:
```bash
python xlsform_to_printable.py --input "Brick Kiln Literacy RCT - Household Survey.xlsx"
```
To target a specific language explicitly:
```bash
python xlsform_to_printable.py --input "Brick Kiln Literacy RCT - Household Survey.xlsx" --lang English
```
To process an entire directory of XLSForm files:
```bash
python xlsform_to_printable.py --dir /path/to/forms/
```
To generate HTML codebooks only without rendering PDFs:
```bash
python xlsform_to_printable.py --input "Brick Kiln Literacy RCT - Household Survey.xlsx" --no-pdf
```

### Post-Processing Existing HTML Files
To process an existing HTML codebook export and generate landscape A4 PDFs:
```bash
python html_make_responsive.py --input codebook.html
```

## Technical Specifications & Print Optimization

- **Page & Margin Setup**: `@page { size: A4 landscape; margin: 8mm 10mm; }` for compact, high-density codebook layout.
- **Repeating Headers**: `thead { display: table-header-group; }` automatically repeats table column headers (`Field`, `Question`, `Answer`) at the top of every page break.
- **Typography & Encoding**: All files use UTF-8 encoding. Uses `"Arial Unicode MS", Arial, sans-serif` to ensure complete rendering of Odia, Devanagari (Hindi), Tamil, Bengali, and other international scripts.
- **Ink & Toner Savings**: Solid black header fills (`#1a1a1a`) are automatically replaced during printing with light gray tint headers (`#f0f4f8` / `#e9ecef`) and crisp dark text, preventing heavy ink saturation while maintaining clear visual distinction.
- **Word Wrapping**: Enforces `overflow-wrap: break-word; word-break: break-word;` across all table cells to eliminate horizontal clipping and text overflow across page boundaries.

## Markdown Guidelines for Contributor Documentation

All Markdown documentation in this project strictly follows a single-line paragraph rule. Paragraphs and bullet points must be written as a single unwrapped line without manual column-limit line breaks.
