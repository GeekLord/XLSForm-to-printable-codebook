# XLSForm to Printable Codebook & PDF Generator

This repository provides tools and scripts to convert XLSForm spreadsheets (`.xlsx`) from ODK, SurveyCTO, and KoboToolbox platforms directly into responsive, print-ready HTML questionnaires and high-quality A4 PDF codebooks.

## Project Overview

Survey forms created for ODK, SurveyCTO, or KoboToolbox are authored as XLSForm spreadsheets containing `survey`, `choices`, and `settings` worksheets. Previously, generating a printable codebook required uploading the XLSForm to a survey server, exporting the printable HTML file, and running post-processing scripts (`html_make_responsive.py` or `html_make_responsive_print.py`) to inject responsive CSS and render PDFs.

The core objective of this project is to provide a direct Python pipeline (`xlsform_to_printable.py`) that parses `.xlsx` XLSForm files directly, handles platform variations across ODK, SurveyCTO, and KoboToolbox, auto-detects all form languages (including un-designated labels), displays platform-specific metadata explanations and categorized calculation formulas in the Question column, renders nested choice tables in the Answer column, and outputs print-ready HTML and PDF codebooks for all detected languages without requiring any server upload or export step.

## Key Features & Capabilities

- **Cross-Platform Compatibility**: Automatically normalizes syntax differences between ODK, SurveyCTO, and KoboToolbox (e.g. `begin_group` vs `begin group`, `end_group` vs `end group`, `begin_repeat` vs `begin repeat`, `select_one_from_file`, `select_multiple_from_file`, `or_other`, and platform-specific metadata).
- **Platform Metadata Descriptions**: Automatically tags and explains platform-specific metadata variables (such as `start`, `end`, `today`, `deviceid`, `phonenumber`, `simserial`, `subscriberid`, `username`, `email`, `audit`, `text-audit`, `start-geopoint`, `caseid`, `caseread`, `casesave`) in the Question column with human-readable descriptions (e.g. `[ODK / SurveyCTO / KoboToolbox Metadata]: Automated timestamp recorded when the survey session starts`).
- **Categorized Calculation & Formula Presentation**: Classifies and explains platform-specific calculation functions in the Question column, categorizing formulas such as `pulldata()` (External Dataset Query), `jr:choice-name()` (Choice Label Lookup), `once(now())` (Timestamp Capture), `format-date-time()` (Date/Time Formatter), `decimal-date-time()` (Duration Math), `selected-at()` / `count-selected()` (Multi-Select Processing), `substr()` / `concat()` (String Processing), and `if()` (Conditional Logic).
- **Automated Multi-Language Processing**: Automatically scans column headers for language tags (`label::English (en)`, `label::Hindi (hi)`, `label::Spanish`, etc.) as well as un-tagged `label` columns. If multiple languages are present, codebooks are generated for all detected languages automatically (e.g. `Survey_English.html`, `Survey_Hindi.html`, `Survey_English.pdf`, `Survey_Hindi.pdf`).
- **Populated Choice Tables**: Parses the `choices` sheet safely using zero-index-aware header resolution (`get_header_index()`) and renders clean nested tables in the Answer column mapping choice values (`name`) to choice labels for each target language.
- **Responsive & Print-Optimized CSS**: Features dark section headers (`#1a1a1a`) for form groups, exact print color retention (`-webkit-print-color-adjust: exact`), Devanagari font fallback (`Arial Unicode MS`), page-break protection (`page-break-inside: avoid`), and landscape A4 Playwright PDF rendering.

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
- `html_make_responsive.py` — Existing script that processes SurveyCTO-exported HTML codebooks, applies responsive and print CSS, and generates A4 landscape PDFs using Playwright Chromium.
- `html_make_responsive_print.py` — Alternative script that converts SurveyCTO HTML tables into a card-style `.question-item` layout for portrait A4 printing.
- `backup_html_make_responsive.py` — Frozen reference backup of the HTML responsive transformation script.
- `xlsform_to_printable_plan.md` — Detailed technical architecture, XLSForm element mapping, and implementation roadmap for the direct XLSForm parser.
- `PA_KAP_Endline_CR_Programme_UP_Bihar.xlsx` — Sample XLSForm spreadsheet for the KAP Endline survey.
- `PA_Panel_Diary_CR_Programme_UP_Bihar.xlsx` — Sample XLSForm spreadsheet for the Panel Diary survey.
- `*_English.html` / `*_Hindi.html` — Sample generated codebooks for English and Hindi form variations.
- `*_English.pdf` / `*_Hindi.pdf` — Sample generated A4 landscape codebook PDFs.

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
Run `xlsform_to_printable.py` directly against XLSForm Excel files to generate HTML and PDF codebooks for all detected languages:
```bash
python xlsform_to_printable.py --input PA_KAP_Endline_CR_Programme_UP_Bihar.xlsx
```
To target a specific language explicitly:
```bash
python xlsform_to_printable.py --input PA_KAP_Endline_CR_Programme_UP_Bihar.xlsx --lang English
```
To process an entire directory of XLSForm files:
```bash
python xlsform_to_printable.py --dir /path/to/forms/
```
To generate HTML codebooks only without rendering PDFs:
```bash
python xlsform_to_printable.py --input PA_KAP_Endline_CR_Programme_UP_Bihar.xlsx --no-pdf
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
- **Platform Metadata**: Automatically tagged in Question column with gray left-bordered containers (`.metadata-tag`).
- **Calculation Display**: Calculation fields display their categorized formula and explanation in blue left-bordered containers (`.calculation`) inside the Question column.
- **Form Metadata & Logic**: Field names, required indicators (`(required)`), hints (styled in blue), relevance expressions (styled in green italic), and constraints are parsed and rendered preserving field visibility rules.
- **Print Optimization**: Configured with `@media print` rules, `page-break-inside: avoid` on question rows, and exact color preservation for A4 landscape and portrait printing.

## Markdown Guidelines for Contributor Documentation

All Markdown documentation in this project strictly follows a single-line paragraph rule. Paragraphs and bullet points must be written as a single unwrapped line without manual column-limit line breaks.
