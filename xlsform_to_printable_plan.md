# Direct XLSForm to Printable Codebook, PDF & DOCX Generator Implementation Plan

This document outlines the technical design, architecture, and step-by-step roadmap for developing `xlsform_to_printable.py`, a standalone Python utility that directly parses XLSForm files (`.xlsx`) and generates responsive HTML codebooks, print-ready A4 PDFs, and editable Word DOCX codebooks across ODK, SurveyCTO, and KoboToolbox platform variants.

## Executive Summary & Objectives

Creating codebook PDFs previously required uploading XLSForm spreadsheets (`.xlsx`) to a survey server (such as SurveyCTO), downloading the generated HTML preview, and post-processing it. The objective of `xlsform_to_printable.py` is to parse XLSForm worksheets (`survey`, `choices`, `settings`) directly, account for platform-specific variations across ODK, SurveyCTO, and KoboToolbox, automatically detect and process all present languages (or un-designated default labels), explain platform-specific metadata variables, categorize calculation formulas in the Question column, populate nested choice tables in the Answer column, embed responsive and print CSS, and render high-quality A4 PDFs using Playwright.

## Multi-Platform Compatibility & Metadata Handling (ODK, SurveyCTO, KoboToolbox)

XLSForm standards are shared across ODK, SurveyCTO, and KoboToolbox, but platforms introduce specific metadata fields and syntax variations that the parser normalizes:
- **Keyword & Row Type Normalization**: Supports underscore and space variations in row types e.g. `begin_group` / `begin group`, `end_group` / `end group`, `begin_repeat` / `begin repeat`, `end_repeat` / `end repeat`.
- **Platform Metadata Variables**: Automatically tags and provides explanatory descriptions for metadata types (`start`, `end`, `today`, `deviceid`, `phonenumber`, `simserial`, `subscriberid`, `username`, `email`, `audit`, `text-audit`, `start-geopoint`, `caseid`, `caseread`, `casesave`) in the Question column (e.g. `[ODK / SurveyCTO / KoboToolbox Metadata]: Automated timestamp recorded when the survey session starts`).
- **Select Types**: Normalizes `select_one <list_name>`, `select_multiple <list_name>`, `select_one_from_file <file>`, `select_multiple_from_file <file>`, `select_one_or_other <list_name>`, and `select_multiple_or_other <list_name>` constructs.
- **Media & Attachment Columns**: Handles `image`, `audio`, `video`, `media::image`, `media::image::<lang>`, `image::English`, and `questionPrompt` / `choicePrompt` image rendering.
- **Form Metadata**: Gracefully reads `settings` columns common across platforms (`form_title` / `title`, `id_string` / `form_id`, `version`, `default_language`, `style`).

## Categorized Platform Calculation & Formula Presentation

Calculations in XLSForms serve distinct operational purposes across platforms (dataset lookup, label resolution, time tracking, string parsing, conditional logic):
- **SurveyCTO / ODK External Dataset Query (`pulldata`)**: Categorized and explained as querying pre-loaded CSV datasets or server case data using unique lookup keys.
- **ODK / SurveyCTO Choice Label Lookup (`jr:choice-name`)**: Categorized and explained as resolving raw choice option codes into human-readable label strings.
- **Timestamp Capture (`once(now())`, `now()`)**: Categorized and explained as capturing current date and time upon evaluation.
- **Date/Time Formatter (`format-date-time`)**: Categorized and explained as formatting raw timestamps into custom formatted date/time strings.
- **Duration / Datetime Math (`decimal-date-time`)**: Categorized and explained as converting datetime to decimal days for duration math calculations.
- **Multi-Select & Repeat Functions (`selected-at`, `count-selected`)**: Categorized and explained as extracting or counting selected response options in multiple-choice fields.
- **String Processing (`substr`, `concat`, `string-length`)**: Categorized and explained as performing text manipulation, substring extraction, or string concatenation.
- **Conditional Logic (`if`)**: Categorized and explained as evaluating conditional logic expressions.

## Comprehensive Multi-Language Auto-Detection & Batch Rendering

Forms may contain one language, multiple languages, or plain columns without explicit language qualifiers:
- **Header Parsing & Normalization**: Parser scans header columns for language tags like `label::English (en)`, `label::Hindi (hi)`, `label::Spanish`, `label::es`, `hint::English`, `constraint_message::Hindi`, and un-tagged `label` / `hint`.
- **Language Inventory & Fallback**: Constructs a list of all detected languages. If no language tags are present, defaults to `Default` (unspecified). `get_lang_value()` provides clean fallback to base language names or default strings.
- **Automated Multi-Language Codebook Generation**: When executing without a restrictive `--lang` flag, the script automatically loops over all identified languages in the XLSForm and generates separate codebooks for each language (e.g. `Survey_English.html`, `Survey_Hindi.html`, `Survey_English.pdf`, `Survey_Hindi.pdf`).

## Safe Choice Parsing & Zero-Index Resolution

- **Header Index Resolution**: Employs `get_header_index(header_map, *keys)` to ensure that column index `0` (the 1st column of a worksheet, e.g. `list_name` in `choices`) is recognized as a valid integer and not evaluated as falsy.
- **Flexible List Lookup**: Employs `get_choice_list_by_name(choices_dict, list_name)` for trimmed and case-insensitive choice list matching.
- **Choice Table Rendering**: Renders nested `<table>` elements inside the Answer column (`td.answerCell`) mapping choice values (`name`) to language-specific choice labels.

## System Architecture & Data Flow

1. **XLSForm Spreadsheet Reader**: Utilizes `openpyxl` to load the workbook, inspect sheet headers, and extract data from `survey`, `choices`, and `settings` sheets.
2. **Language & Platform Normalizer**: Scans column headers, identifies all language variations, normalizes platform keyword differences, maps platform metadata, and resolves `choices` lists per language.
3. **Survey Object Model**: Constructs an in-memory hierarchy representing form items, section groups (`begin_group`), repeat blocks (`begin_repeat`), question definitions, choices, logic constraints (`relevant`, `required`, `constraint`, `calculation`, `hint`), and media.
4. **HTML Codebook Renderer**: Constructs standard 3-column codebook table markup (`Field`, `Question`, `Answer`), renders section header rows for groups, embeds nested choice tables for select items, renders categorized calculation formulas and metadata explanations, and applies badge styles for hints and relevance expressions.
5. **CSS & Style Injection**: Injects `RESPONSIVE_STYLES` into the HTML `<head>` ensuring exact print color adjustment, responsive viewport meta tags, and Devanagari/international font support (`Arial Unicode MS`).
6. **Async Playwright PDF Engine**: Launches headless Chromium, loads synthesized HTML files with `file:///` protocol, waits for fonts to render, and exports A4 landscape or portrait PDFs.
7. **Editable DOCX Renderer**: Consumes the same Survey Object Model as the HTML renderer (never the rendered HTML or PDF) and emits a landscape A4 Word document through `python-docx`, reproducing the 3-column table, section bands, metadata and calculation blocks, and nested choice tables in WordprocessingML.

## Shared Renderer Contract Between HTML, PDF & DOCX

The HTML renderer and the DOCX renderer are siblings over one parsed model rather than a pipeline, so the Word file is not a conversion artefact and cannot drift from the printable output:
- **Single Source of Truth**: `build_codebook_html()` and `build_codebook_docx()` both read `XLSFormParser.survey_rows`, `XLSFormParser.choices`, and `XLSFormParser.settings`.
- **Shared Row Semantics**: `split_type()` separates base type from list name, `is_section_start()` and `is_section_end()` classify group and repeat rows, and `trim_calculation()` applies the shared 65-character formula display limit (`CALC_TRIM_LENGTH`).
- **Shared Explanatory Layer**: `METADATA_DESCRIPTIONS`, `classify_calculation()`, `get_lang_value()`, and `get_choice_list_by_name()` are called identically by both renderers.
- **Optional Dependency Isolation**: `python-docx` is imported behind a `DOCX_AVAILABLE` guard so a missing install degrades to HTML and PDF output with a printed hint, never an exception.
- **Independent Output Toggles**: `process_xlsform()` accepts `generate_pdf_flag` and `generate_docx_flag`, exposed as `--no-pdf` and `--no-docx`, and writes `<base>_<Language>.html`, `.pdf`, and `.docx` side by side.

## XLSForm Element to HTML Component Mapping

- **Form Title**: Extracted from `settings` (`form_title` / `title`) or file name, rendered as top-level `<h4>` header.
- **Section Headers (`begin_group` / `begin_repeat`)**: Rendered as full-width `<tr class="entryRow"><td colspan="3" class="section-header-dark">` containing group label or group name.
- **Question Field Name**: Placed in column 1 (`td.fieldCell`), appending `<span class="required">(required)</span>` if `required` is set to `yes`, `1`, `true`, or `ok`.
- **Question Text & Metadata**: Placed in column 2 (`td.questionCell`), displaying question label, platform metadata tag (in `<div class="metadata-tag">`), calculation formula category & explanation (in `<div class="calculation">`), hint (in `<div class="hint">`), relevance expression (in `<div class="relevance">Relevance: ...</div>`), and constraint details.
- **Choice Lists (`select_one` / `select_multiple`)**: Placed in column 3 (`td.answerCell`), containing a nested `<table class="table borderless">` mapping choice values (`name`) to choice labels for the targeted language.
- **Open-Ended Inputs (`text`, `integer`, `decimal`, `date`, `time`, `geopoint`)**: Placed in column 3 (`td.answerCell`), showing variable data type indicator or blank input line.
- **Notes (`note`)**: Rendered in column 2 as informational prompt text with an empty answer cell.
- **Calculations (`calculate`)**: Rendered with field name in column 1 and calculation formula category & description in column 2.

## XLSForm Element to DOCX Component Mapping

- **Form Title**: Rendered as a centred bold 14pt paragraph carrying a 1.5pt bottom border, the Word equivalent of the HTML `<h4>` rule.
- **Page Geometry**: One landscape A4 section (297mm x 210mm) with 10mm margins, giving a 277mm text block divided 22% / 50% / 28% across a fixed-layout table (`w:tblLayout` type `fixed`, explicit `w:gridCol` and `w:tcW` widths).
- **Column Header Row**: `Field`, `Question`, `Answer` shaded `#E9ECEF` and flagged `w:tblHeader` so Word repeats it on every page, matching `thead { display: table-header-group }`.
- **Section Headers (`begin_group` / `begin_repeat`)**: A row merged across all three columns, shaded `#F0F4F8` with 1.5pt `w:tcBorders` above and below, matching the print stylesheet rather than the darker screen treatment.
- **Question Field Name**: Column 1 in bold, with an italic red ` (required)` run appended when `required` resolves true.
- **Question Text & Metadata**: Column 1 label paragraph followed by stacked blocks in the HTML order, each drawn with paragraph shading (`w:shd`) and a left accent border (`w:pBdr`): metadata note (`#F1F3F5` on `#6C757D`), calculation category badge with the trimmed formula in a monospace run (`#E7F5FF` on `#1864AB`), category description, hint, then relevance.
- **Choice Lists (`select_one` / `select_multiple` / `select_or_other`)**: A borderless nested table inside the Answer cell, a narrow centred bold code column beside a wrapping label column, separated by dotted `w:insideH` rules.
- **Open-Ended Inputs & Notes**: Answer cell left empty for handwritten responses, matching the HTML.
- **Page-Break Protection**: `w:cantSplit` on every outer and nested row, the Word equivalent of `page-break-inside: avoid`.
- **International Script Support**: Every run sets `Arial Unicode MS` on `w:ascii`, `w:hAnsi`, `w:cs`, and `w:eastAsia`, plus the complex-script twins `w:szCs`, `w:bCs`, and `w:iCs`, without which Word renders Devanagari and Odia runs at a default 10pt regular.

## WordprocessingML Correctness Constraints

Word rejects a document whose property children appear out of schema order, surfacing as an unreadable-content repair prompt rather than a parse error, so the DOCX layer encodes ordering explicitly:
- **Ordered Insertion**: `_docx_set_child(parent, order, tag)` inserts or replaces a child at its schema slot using the `_DOCX_PPR_ORDER`, `_DOCX_RPR_ORDER`, `_DOCX_TBLPR_ORDER`, `_DOCX_TCPR_ORDER`, and `_DOCX_TRPR_ORDER` sequences instead of appending.
- **Row Width Restamping**: `_docx_apply_row_widths()` runs on every row added after `_docx_fixed_layout()`, because `Table.add_row()` copies an even split of the original table width onto new cells and silently overrides the fixed grid.
- **Cell Content Reset**: `_docx_reset_cell()` strips the placeholder paragraph so content starts flush at the top of a cell, and the trailing paragraph Word requires after a nested table is shrunk to 1pt.

## Detailed Phased Implementation Roadmap

### Phase 1: Multi-Platform Parser & Metadata Extractor
- Implement `XLSFormParser` class in `xlsform_to_printable.py` to open `.xlsx` files using `openpyxl`.
- Build dynamic header parser to identify all language variants (`label::*`, `hint::*`, `constraint_message::*`, `label`) and map choices per language using `get_header_index()`.
- Implement platform metadata dictionary (`METADATA_DESCRIPTIONS`) and calculation classifier (`classify_calculation()`).

### Phase 2: Multi-Language HTML Synthesizer
- Build `HTMLBuilder` generating clean, valid HTML5 with `utf-8` encoding.
- Structure standard 3-column table (`Field`, `Question`, `Answer`) with group section headers, metadata tags, categorized calculation formula badges, and nested choice tables.
- Iterate over each detected language to build language-specific HTML codebooks automatically.
- Embed `RESPONSIVE_STYLES` CSS into the `<head>` of each output document.

### Phase 3: Playwright PDF Generator & CLI Interface
- Integrate `async_playwright` to render A4 landscape PDFs for all generated language HTML files.
- Add CLI arguments using `argparse` supporting single file conversion (`--input`), directory batch conversion (`--dir`), specific language override (`--lang`), and PDF export toggle (`--no-pdf`).
- Add robust error handling so batch processing continues cleanly if a file fails.

### Phase 4: Validation across ODK, SurveyCTO, and KoboToolbox Forms
- Test parser against sample forms (`PA_KAP_Endline_CR_Programme_UP_Bihar.xlsx` and `PA_Panel_Diary_CR_Programme_UP_Bihar.xlsx`).
- Verify multi-language output (English and Hindi), single-language plain label forms, choice table population (54 choice tables in KAP Endline, 16 in Panel Diary), metadata explanations (`start`, `end`, `today`, `deviceid`, `username`, `start-geopoint`), calculation formula displays (24 in KAP Endline, 30 in Panel Diary), and platform syntax variations.

### Phase 5: Editable DOCX Renderer
- Extract the row semantics shared by both renderers (`split_type()`, `is_section_start()`, `is_section_end()`, `trim_calculation()`) so the HTML output stays byte-identical while the DOCX renderer reuses them.
- Implement `build_codebook_docx()` and `generate_docx()` over `python-docx`, behind the `DOCX_AVAILABLE` import guard.
- Implement the WordprocessingML helper layer: schema-ordered child insertion, cell and paragraph shading, explicit borders, fixed table layout, repeating header rows, `w:cantSplit`, and complex-script run styling.
- Extend `process_xlsform()` with `generate_docx_flag` and the CLI with `--no-docx`, writing `<base>_<Language>.docx` beside the existing HTML and PDF.

## Verification & Quality Assurance Strategy

- **Multi-Language Output Verification**: Confirm that forms with multiple languages generate separate output codebooks for each language, and forms without explicit language tags generate a Default codebook.
- **Metadata & Calculation Audit**: Ensure all platform metadata fields display explanatory descriptions in column 2, and all `calculate` fields display formula categories and explanations in column 2.
- **Choice Table Audit**: Ensure all `select_one` and `select_multiple` questions display populated choice option tables in column 3.
- **Cross-Platform Compatibility**: Validate parsing across ODK, SurveyCTO, and KoboToolbox column conventions.
- **Visual & Print Audit**: Verify table layout, column widths, font scaling, and page-break rules in rendered PDFs.
- **HTML Regression Guard**: Confirm that refactors extracting shared helpers leave the generated HTML byte-identical to the previous revision for every language of a reference form.
- **HTML to DOCX Parity Audit**: Confirm equal body row counts, section band counts, and choice table counts, and identical field name and section label sequences between the two renderers.
- **DOCX Structural Validation**: Confirm zero schema-ordering violations across every `w:pPr`, `w:rPr`, `w:tblPr`, `w:tcPr`, and `w:trPr` element, and one consistent column width pattern across all rows.
- **Word Open Audit**: Open each generated DOCX in Word and confirm it loads without a repair prompt, reporting landscape orientation, A4 page dimensions, a repeating header row, and the expected 22% / 50% / 28% cell widths.
