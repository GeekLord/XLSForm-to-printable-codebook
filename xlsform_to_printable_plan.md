# Direct XLSForm to Printable Codebook & PDF Generator Implementation Plan

This document outlines the technical design, architecture, and step-by-step roadmap for developing `xlsform_to_printable.py`, a standalone Python utility that directly parses XLSForm files (`.xlsx`) and generates responsive HTML codebooks and print-ready A4 PDFs across ODK, SurveyCTO, and KoboToolbox platform variants.

## Executive Summary & Objectives

Creating codebook PDFs previously required uploading XLSForm spreadsheets (`.xlsx`) to a survey server (such as SurveyCTO), downloading the generated HTML preview, and post-processing it. The objective of `xlsform_to_printable.py` is to parse XLSForm worksheets (`survey`, `choices`, `settings`) directly, account for platform-specific variations across ODK, SurveyCTO, and KoboToolbox, automatically detect and process all present languages (or un-designated default labels), construct structured HTML tables, embed responsive and print CSS, and render high-quality A4 PDFs using Playwright.

## Multi-Platform Compatibility (ODK, SurveyCTO, KoboToolbox)

XLSForm standards are shared across ODK, SurveyCTO, and KoboToolbox, but platforms introduce specific syntax variations that the parser must normalize:
- **Keyword Normalization**: Supports underscore and space variations in row types e.g. `begin_group` / `begin group`, `end_group` / `end group`, `begin_repeat` / `begin repeat`, `end_repeat` / `end repeat`.
- **Select Types**: Normalizes `select_one <list_name>`, `select_multiple <list_name>`, `select_one_from_file <file>`, `select_multiple_from_file <file>`, and `or_other` constructs.
- **Media & Attachment Columns**: Handles `image`, `audio`, `video`, `media::image`, `media::image::<lang>`, `image::English`, and `questionPrompt` / `choicePrompt` image rendering.
- **Form Metadata**: Gracefully reads `settings` columns common across platforms (`form_title` / `title`, `id_string` / `form_id`, `version`, `default_language`, `style`).

## Comprehensive Multi-Language Auto-Detection & Batch Rendering

Forms may contain one language, multiple languages, or plain columns without explicit language qualifiers:
- **Header Parsing & Normalization**: Regex parser scans header columns for language tags like `label::English (en)`, `label::Hindi (hi)`, `label::Spanish`, `label::es`, `hint::English`, `constraint_message::Hindi`, and un-tagged `label` / `hint`.
- **Language Inventory**: Constructs a list of all detected languages. If no language tags are present, defaults to `Default` (unspecified).
- **Automated Multi-Language Codebook Generation**: When executing without a restrictive `--lang` flag, the script automatically loops over all identified languages in the XLSForm and generates separate codebooks for each language (e.g. `Survey_English.html`, `Survey_Hindi.html`, `Survey_English.pdf`, `Survey_Hindi.pdf`).

## System Architecture & Data Flow

1. **XLSForm Spreadsheet Reader**: Utilizes `openpyxl` to load the workbook, inspect sheet headers, and extract data from `survey`, `choices`, and `settings` sheets.
2. **Language & Platform Normalizer**: Scans column headers, identifies all language variations, normalizes platform keyword differences, and resolves `choices` lists per language.
3. **Survey Object Model**: Constructs an in-memory hierarchy representing form items, section groups (`begin_group`), repeat blocks (`begin_repeat`), question definitions, choices, logic constraints (`relevant`, `required`, `constraint`, `calculation`, `hint`), and media.
4. **HTML Codebook Renderer**: Constructs standard 3-column codebook table markup (`Field`, `Question`, `Answer`), renders section header rows for groups, embeds nested choice tables for select items, and applies badge styles for hints, relevance expressions, and media references.
5. **CSS & Style Injection**: Injects `RESPONSIVE_STYLES` into the HTML `<head>` ensuring exact print color adjustment, responsive viewport meta tags, and Devanagari/international font support (`Arial Unicode MS`).
6. **Async Playwright PDF Engine**: Launches headless Chromium, loads synthesized HTML files with `file:///` protocol, waits for fonts to render, and exports A4 landscape or portrait PDFs.

## XLSForm Element to HTML Component Mapping

- **Form Title**: Extracted from `settings` (`form_title` / `title`) or file name, rendered as top-level `<h4>` header.
- **Section Headers (`begin_group` / `begin_repeat`)**: Rendered as full-width `<tr class="entryRow"><td colspan="3" class="section-header-dark">` containing group label or group name.
- **Question Field Name**: Placed in column 1 (`td.fieldCell`), appending `<span class="required">(required)</span>` if `required` is set to `yes`, `1`, `true`, or `OK`.
- **Question Text & Metadata**: Placed in column 2 (`td.questionCell`), displaying question label, hint (in `<div class="hint">`), relevance expression (in `<div class="relevance">Relevance: ...</div>`), and constraint details.
- **Choice Lists (`select_one` / `select_multiple`)**: Placed in column 3 (`td.answerCell`), containing a nested `<table class="table borderless">` mapping choice values (`name`) to choice labels for the targeted language.
- **Open-Ended Inputs (`text`, `integer`, `decimal`, `date`, `time`, `geopoint`)**: Placed in column 3 (`td.answerCell`), showing variable data type indicator or blank input line.
- **Notes (`note`)**: Rendered in column 2 as informational prompt text with an empty answer cell.
- **Calculations & Hidden Fields (`calculate`, `start`, `end`, `deviceid`, `subscriberid`)**: Optionally rendered or flagged as internal metadata fields based on CLI flags.

## Detailed Phased Implementation Roadmap

### Phase 1: Multi-Platform Parser & Language Extractor
- Implement `XLSFormParser` class in `xlsform_to_printable.py` to open `.xlsx` files using `openpyxl`.
- Build dynamic header parser to identify all language variants (`label::*`, `hint::*`, `constraint_message::*`, `label`) and map choices per language.
- Implement platform syntax normalizers for ODK, SurveyCTO, and KoboToolbox row types and column names.

### Phase 2: Multi-Language HTML Synthesizer
- Build `HTMLBuilder` generating clean, valid HTML5 with `utf-8` encoding.
- Structure standard 3-column table (`Field`, `Question`, `Answer`) with group section headers and nested choice tables.
- Iterate over each detected language to build language-specific HTML codebooks automatically.
- Embed `RESPONSIVE_STYLES` CSS into the `<head>` of each output document.

### Phase 3: Playwright PDF Generator & CLI Interface
- Integrate `async_playwright` to render A4 landscape PDFs for all generated language HTML files.
- Add CLI arguments using `argparse` supporting single file conversion (`--input`), directory batch conversion (`--dir`), specific language override (`--lang`), and PDF export toggle (`--pdf`).
- Add robust error handling so batch processing continues cleanly if a file fails.

### Phase 4: Validation across ODK, SurveyCTO, and KoboToolbox Forms
- Test parser against sample forms (`PA_KAP_Endline_CR_Programme_UP_Bihar.xlsx` and `PA_Panel_Diary_CR_Programme_UP_Bihar.xlsx`).
- Verify multi-language output (English and Hindi), single-language plain label forms, and platform syntax variations.

## Verification & Quality Assurance Strategy

- **Multi-Language Output Verification**: Confirm that forms with multiple languages generate separate output codebooks for each language, and forms without explicit language tags generate a Default codebook.
- **Cross-Platform Compatibility**: Validate parsing across ODK, SurveyCTO, and KoboToolbox column conventions.
- **Visual & Print Audit**: Verify table layout, column widths, font scaling, and page-break rules in rendered PDFs.
