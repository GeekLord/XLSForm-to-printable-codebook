# XLSForm to printable codebook and questionnaire

Print your survey before you field it. Point `xlsform_to_printable.py` at an XLSForm `.xlsx` file to generate a responsive HTML codebook, an A4 landscape PDF, and an editable Word `.docx` for every language in the form, without uploading files to an external server. For an enumerator-facing or respondent-facing survey ready for paper administration, `Questionnaire/xlsform_to_questionnaire.py` produces a clean, table-free printable interview questionnaire.

Direct conversion eliminates the server round trip. Older workflows required uploading the spreadsheet to ODK, SurveyCTO, or KoboToolbox, exporting the platform's default printable HTML, and running post-processing scripts to make tables readable on paper. Two post-processing scripts (`html_make_responsive.py` and `html_make_responsive_print.py`) remain available for exported HTML. The primary converters skip the server step and parse the `survey`, `choices`, and `settings` sheets directly.

## Codebook versus questionnaire

| Feature | Codebook (`xlsform_to_printable.py`) | Interview questionnaire (`Questionnaire/xlsform_to_questionnaire.py`) |
| --- | --- | --- |
| Primary audience | Data managers, research PIs, programmers, and M&E officers | Field enumerators, respondents, review boards, and training sessions |
| Document layout | 3-column fixed table (Field 22%, Question 50%, Answer 28%) | Flowing document without tables, using question prompts, notes, choices, and blanks |
| Default page setup | A4 landscape (10mm margins) | A4 portrait (15mm top and bottom, 12mm and 16mm margins; `--landscape` optional) |
| Variable names and logic | Variable names, required flags, calculation badges, formulas, and relevance rules | Omitted; presents only human-readable question text and instructions |
| Choice lists | Choice values paired with labels in a nested table | Radio `( )` or checkbox `[ ]` markers with labels only; raw codes hidden |
| Open-ended inputs | Empty answer cell (data type shown in Field column) | Tailored handwriting blanks (ruled lines, date/time formats, numeric brackets) |
| Notes (`note` type) | Standard table row with note text in Question column | Full-width callout box with a blue accent bar for scripts and consent forms |
| Calculations and metadata | Documented with plain-language explanations and badges | Excluded completely |
| Output formats | Responsive HTML, landscape PDF, editable Word `.docx` | Responsive HTML, portrait/landscape PDF, editable Word `.docx` |

## Setup

Python 3.8 or newer:

```bash
pip install -r requirements.txt
playwright install chromium
```

Chromium is required only for PDF generation. Pass `--no-pdf` to run without it. `python-docx` is required only for Word output; if the package is missing, both scripts print a notice and generate HTML and PDF. Pass `--no-docx` to disable Word file generation explicitly.

## Usage

### Generating codebooks (xlsform_to_printable.py)

Process every language in the form into HTML, PDF, and DOCX:

```bash
python xlsform_to_printable.py --input "Pradhan Qualitative Survey 2026.xlsx"
```

Generate a single language using a case-insensitive substring match (for example, `Hindi` matches the `label::Hindi (hi)` column):

```bash
python xlsform_to_printable.py --input "Pradhan Qualitative Survey 2026.xlsx" --lang Hindi
```

Process every `.xlsx` file in a folder, skipping temporary Excel lock files:

```bash
python xlsform_to_printable.py --dir /path/to/forms/
```

Generate HTML and DOCX without rendering PDFs:

```bash
python xlsform_to_printable.py --input "Pradhan Qualitative Survey 2026.xlsx" --no-pdf
```

Generate HTML and PDF without writing Word files:

```bash
python xlsform_to_printable.py --input "Pradhan Qualitative Survey 2026.xlsx" --no-docx
```

When run without `--input` or `--dir`, the script processes every `.xlsx` file in the current working directory.

### Generating interview questionnaires (Questionnaire/xlsform_to_questionnaire.py)

Process every language in the form into HTML, portrait PDF, and portrait DOCX:

```bash
python Questionnaire/xlsform_to_questionnaire.py --input "Pradhan Qualitative Survey 2026.xlsx"
```

Generate a single language:

```bash
python Questionnaire/xlsform_to_questionnaire.py --input "Pradhan Qualitative Survey 2026.xlsx" --lang English
```

Render in A4 landscape instead of the default portrait layout:

```bash
python Questionnaire/xlsform_to_questionnaire.py --input "Pradhan Qualitative Survey 2026.xlsx" --landscape
```

Process every `.xlsx` file in a directory:

```bash
python Questionnaire/xlsform_to_questionnaire.py --dir /path/to/forms/
```

Skip PDF rendering:

```bash
python Questionnaire/xlsform_to_questionnaire.py --input "Pradhan Qualitative Survey 2026.xlsx" --no-pdf
```

Skip DOCX generation:

```bash
python Questionnaire/xlsform_to_questionnaire.py --input "Pradhan Qualitative Survey 2026.xlsx" --no-docx
```

Outputs are saved in the same directory as the source spreadsheet, using the language name as a suffix: `<form>_<Language>.html`, `<form>_<Language>.pdf`, and `<form>_<Language>.docx`.

## What the codebook looks like

Three fixed-width columns prevent rows from expanding off the page: Field at 22%, Question at 50%, and Answer at 28%.

The Field column displays the variable name, tagged with a red `(required)` indicator when the `required` column is set to yes, 1, true, ok, or required.

The Question column presents row information in top-to-bottom order:
- The localized question label.
- A metadata note when the row represents an automatic platform variable like `start` or `deviceid`.
- A calculation badge classifying the formula, followed by the formula trimmed to 65 characters. The full expression is preserved in the HTML `title` attribute of the `<code>` element so it remains visible on hover without expanding row height in print.
- Hint text.
- Relevance conditions prefixed with `Relevance:`.

The Answer column contains a nested two-column table for `select_one`, `select_multiple`, and `select_or_other` fields: a 32px code cell on the left and the choice label wrapping on the right. Choice lookups against the `choices` sheet are case-insensitive, and choices without labels display their raw value. External file choices leave this cell empty because option rows do not live in the workbook.

`begin group`, `begin_group`, `begin repeat`, and `begin_repeat` rows render as full-width section bands showing the group label or name. Matching closing tags are omitted.

`constraint` and `constraint_message` columns are parsed from the sheet but not rendered in the codebook table.

## What the interview questionnaire looks like

The questionnaire generator builds a paper-first interview document designed for field administration and ethics review without table borders:

The document header displays the survey title from the `settings` sheet (or the workbook filename), the interview language subtitle, and a divider line.

Section headers from `begin group` and `begin repeat` rows appear as full-width tinted bands (`#f0f4f8` with 2px borders) to mark survey modules. Closing tags (`end group`, `end repeat`) are omitted.

Enumerator instructions, consent statements, and read-aloud prompts from `note` rows become full-width callout boxes with a blue left accent bar (`#3b82f6`) and a light background tint (`#f8fafc`).

Question prompts appear in bold text with standard paragraph spacing. Hints appear in italic blue text indented below the question prompt (`Note: <hint>`). Technical variable names, required flags, relevance conditions, and constraint formulas are omitted.

Multiple-choice questions display radio markers `( )` for `select_one` and checkbox markers `[ ]` for `select_multiple`. Choice values such as `1`, `2`, or `-99` are hidden so only localized choice labels appear.

Open-ended questions provide dedicated handwriting blanks matching the field's data type instead of blank lines.

Platform metadata rows (`start`, `end`, `deviceid`, `today`, `username`, `email`, `audit`) and calculation rows (`calculate`, `hidden`) are omitted from the questionnaire.

### Response blanks by question type

When rendering open-ended questions in the questionnaire, `xlsform_to_questionnaire.py` provides formatted handwriting blanks corresponding to each XLSForm data type:

| Question type | Rendered response blank | Notes |
| --- | --- | --- |
| `text` (standard) | `____________________________________________________________________` | Single ruled line for short free-text answers |
| `text` (detailed/notes) | Double ruled line `____________________________________________________________________` | Triggered when the variable name or label mentions notes, details, comments, specify, other, explain, or reasons |
| `integer`, `decimal`, `range` | `[ ____________________ ]` | Bracketed blank for numeric values |
| `date` | `[ ____ / ____ / ________ ] (DD/MM/YYYY)` | Segmented date blank |
| `time` | `[ ____ : ____ ] (HH:MM)` | Segmented 24-hour time blank |
| `datetime` | `[ ____/____/________  ____:____ ]` | Combined timestamp blank |
| `geopoint`, `geotrace`, `geoshape` | `[ GPS Location Coordinates: ________________________________ ]` | GPS capture indicator |
| `image`, `audio`, `video`, `file` | `[ Attach / Record <Type>: ________________________________ ]` | Media attachment placeholder |
| `barcode` | `[ Barcode / QR: ________________________________ ]` | Barcode or QR scanner placeholder |
| `acknowledge` | `[  ] Confirmed / Acknowledged` | Single acknowledgment checkbox |

## Languages

The parser inspects the header rows in `survey` and `choices` and collects every `label::` column, so headers like `label::Hindi (hi)` and `label::English (en)` become target languages without manual configuration. Untagged `label`, `hint`, and `constraint_message` columns serve as fallbacks for languages lacking dedicated translations. Forms without tagged label columns default to a single `Default` language.

Output filenames drop the parenthetical language code, so `label::Hindi (hi)` produces `_Hindi.html`, not `_Hindi (hi).html`.

## Media images (media::image)

When an XLSForm includes media image references in `survey` or `choices` sheets:
- **Supported Columns**: `media::image::<language>`, `media::image`, `image::<language>`, and `image`.
- **Automatic Resolution**: Image files are discovered in the same folder as the XLSForm spreadsheet, a `media/` sub-folder, or an ODK standard `<form_base>-media/` directory, with case-insensitive matching across platforms.
- **Embedded in HTML & PDF**: Images are encoded as self-contained Base64 data URIs (`data:image/<type>;base64,...`), ensuring generated HTML and PDF files remain completely portable without broken links.
- **Embedded in DOCX**: Images are scaled proportionally to fit page geometry and inserted into WordprocessingML drawings in codebook table cells and questionnaire flows.
- **Missing Image Safety**: If an image referenced in the form is not found, a subtle indicator (`[Image: filename (file not found)]`) is rendered rather than failing or disrupting document layout.

## Platform metadata reference

When a row type matches an automatic survey platform variable, the codebook Question column provides a plain-language explanation instead of an empty cell. (These rows are omitted from the interview questionnaire.)

| Type | What the codebook says |
| --- | --- |
| `start` | Timestamp recorded when the survey session starts |
| `end` | Timestamp recorded when the survey session ends |
| `today` | Date the survey was conducted |
| `deviceid` | Hardware identifier (IMEI or UUID) of the collection device |
| `phonenumber` | Phone number of the SIM card in the device |
| `simserial` | Serial number of the SIM card |
| `subscriberid` | IMSI subscriber identifier of the SIM card |
| `username` | Username of the logged-in enumerator |
| `email` | Email address of the logged-in enumerator |
| `audit` | Activity log of timestamped interaction events |
| `text-audit` | Keystroke and text editing log (SurveyCTO) |
| `start-geopoint` | GPS location captured when the form opens |
| `background-audio` | Background audio recorded during the survey (ODK) |
| `caseid` | Case Management entity identifier (SurveyCTO) |
| `caseread` | Case Management pre-loaded read field (SurveyCTO) |
| `casesave` | Case Management saved status field (SurveyCTO) |

## Calculation categories

Formulas are matched against these patterns in sequence, and the first matching pattern assigns the category. Unmatched formulas are labelled `Calculation`. (Calculations are omitted from the interview questionnaire.)

| Pattern | Category |
| --- | --- |
| `pulldata(` | External dataset query, reading a pre-loaded CSV or server case data |
| `jr:choice-name(` | Choice label lookup, turning a stored code into readable text |
| `once(now())` or `now()` | Timestamp capture |
| `format-date-time(` | Date and time formatting |
| `decimal-date-time(` | Duration math, converting a datetime to decimal days |
| `selected-at(` or `count-selected(` | Multi-select extraction or counting |
| `substr(`, `concat(`, `string-length(` | String processing |
| `if(` | Conditional logic |

Evaluation order determines classification. For example, a formula that nests `now()` inside `format-date-time()` is categorized as a timestamp capture because the `now()` pattern is evaluated first.

## Print behavior

Screen and print layouts use the same stylesheet in both tools, with print formatting specified in `@media print` blocks. DOCX generators replicate these styling rules in WordprocessingML.

### Codebook print layout (xlsform_to_printable.py)

- `@page` specifies A4 landscape with 8mm vertical and 10mm horizontal margins. Playwright sets 10mm margins on all four sides for PDF output.
- `thead` uses `table-header-group` so Field, Question, and Answer headers repeat at the top of each page.
- `page-break-inside: avoid` on table rows keeps questions and choice lists on the same page.
- Section bands and column headers switch from dark navy fills on screen (`#2b2d42` and `#4a4e69` with white text) to light tints in print (`#f0f4f8` with 1.5pt rules above and below, and `#e9ecef` for column headers with dark text). Both views include `print-color-adjust: exact` so browser print engines do not drop background shading.
- Typography uses 9.5pt on 13.5pt leading in print, and 14px body with 13px to 13.5px table cells on screen.
- Cells include `overflow-wrap: break-word` and `word-break: break-word` so long variable names wrap cleanly without horizontal scrollbars.
- Responsive breakpoints at 768px and 480px adjust font sizes and enable horizontal scrolling on small screens rather than compressing columns.
- UTF-8 encoding is used throughout, and the font stack prioritizes Arial Unicode MS so Hindi, Odia, Bengali, and Tamil scripts render accurately.

### Questionnaire print layout (xlsform_to_questionnaire.py)

- `@page` defaults to A4 portrait (`size: A4 portrait; margin: 15mm 12mm;`), switching to landscape when `--landscape` is supplied.
- `page-break-inside: avoid` on `.question-item` and `.note-item` prevents questions and notes from breaking across page margins.
- `page-break-after: avoid` on `.section-header` keeps section headings from separating from following questions at the bottom of a page.
- Section headers use a light tint (`#f0f4f8`) with 1.5pt top and bottom rules (`#333333`) and `print-color-adjust: exact`.
- Note blocks maintain a light background tint (`#f8fafc`) and a 3pt blue left border (`#3b82f6`) on paper.
- Container width is capped at 820px in portrait (1100px in landscape) on screen, centered with a subtle drop shadow, and expands to full width during print.

## Editable DOCX output

Word documents are generated directly from parsed survey data using `python-docx` rather than converted from HTML or PDF, preventing rendering drift between output formats.

### Codebook Word document (xlsform_to_printable.py)

- Uses a three-column fixed table matching the HTML width distribution: 22%, 50%, and 28% of the 277mm printable width (172.8pt, 392.6pt, and 219.9pt).
- Configured for A4 landscape with 10mm margins, matching the PDF geometry.
- Header rows repeat across pages using Word's `tblHeader` property.
- Table rows include `cantSplit` to prevent questions and choice lists from splitting across pages.
- Group and repeat rows span all three columns with `#F0F4F8` shading and 1.5pt rules above and below.
- Metadata notes and calculation badges preserve shading and left accent borders via paragraph properties.
- Choice lists use a borderless nested table inside the Answer cell, pairing choice codes with wrapping labels.
- Formulas longer than 65 characters are trimmed to fit table cells, matching the PDF display.

### Questionnaire Word document (xlsform_to_questionnaire.py)

- Formatted as a table-free flowing document in A4 portrait (210mm x 297mm) with 16mm side margins and 15mm top and bottom margins, or A4 landscape (297mm x 210mm) when `--landscape` is passed.
- Question prompts use 10.5pt bold text with paragraph spacing and `keepWithNext` (`w:keepNext`) to keep prompts attached to choices and blanks across page breaks.
- Hints indent 10pt using italic blue text (`#1E40AF`) with `keepWithNext` enabled.
- Choice options indent 13pt with `( )` or `[ ]` markers in muted slate (`#64748B`) followed by the choice label, setting `keepWithNext` on initial choices to prevent split lists.
- Open-ended response blanks indent 12pt using ruled lines or brackets.
- Section headers use 12pt bold text with `#F0F4F8` paragraph shading, top and bottom rules (`#334155`), and `keepWithNext` enabled.
- Note callout blocks use 9.5pt text with light gray shading (`#F8FAFC`), 7pt left indentation, and a 3pt blue left border (`#3B82F6`).

### Typography and schema compliance

- Every run in both Word generators specifies Arial Unicode MS across `w:ascii`, `w:hAnsi`, `w:cs`, and `w:eastAsia`, and sets the complex-script attributes `w:szCs`, `w:bCs`, and `w:iCs`. This ensures Hindi, Bengali, Odia, and Tamil text retains the specified font, size, and weight instead of reverting to Word's default 10pt styling.
- OpenXML ECMA-376 requires child property elements to appear in a strict order. Both generators use schema-ordered helpers (`_docx_set_child`, `_DOCX_PPR_ORDER`, `_DOCX_RPR_ORDER`) to place XML tags in their valid positions, preventing unreadable-content repair warnings when opening files in Word.

## Repository files

- `xlsform_to_printable.py`: Direct codebook converter. Parses `.xlsx` workbooks, constructs 3-column HTML tables, and generates landscape A4 PDFs and DOCX codebooks.
- `Questionnaire/xlsform_to_questionnaire.py`: Direct interview questionnaire converter. Parses `.xlsx` workbooks and produces table-free flowing HTML, portrait or landscape A4 PDFs, and editable DOCX questionnaires for field administration.
- `Pradhan Qualitative Survey 2026.xlsx`: Sample XLSForm workbook containing survey questions, choices, and settings in English and Hindi.
- `Questionnaire/Pradhan_Qualitative_Survey_2026.xlsx`: Companion sample workbook located in the Questionnaire directory.
- `html_make_responsive.py`: Legacy post-processor for survey-platform exported HTML. Adds responsive and print styles to existing HTML tables and renders landscape A4 PDFs.
- `html_make_responsive_print.py`: Alternative legacy post-processor that restructures exported HTML tables into card-style question blocks for portrait A4 printing.
- `backup_html_make_responsive.py`: Reference archive of the earlier HTML post-processor.
- `xlsform_to_printable_plan.md`: Technical documentation containing the XLSForm element mapping matrix, calculation rules, and parser architecture.
- `requirements.txt`: Pinned Python package dependencies (`openpyxl`, `beautifulsoup4`, `playwright`, `python-docx`).

Both post-processing scripts process every `.html` file in the current working directory, overwrite each file with the responsive version, and write a PDF beside it:

```bash
python html_make_responsive.py
```

## Writing docs in this repo

Markdown paragraphs and list items are single unwrapped lines. Do not hard-wrap prose at a column limit or break a line inside a paragraph or bullet.
