# XLSForm to printable codebook and questionnaire

Print your survey before you field it. Point `xlsform_to_printable.py` at an XLSForm `.xlsx` file and you get a responsive HTML codebook, an A4 landscape PDF, and an editable Word `.docx` for every language in the form, without uploading anything to a server. For an enumerator-facing or respondent-facing survey ready for paper administration, `Questionnaire/xlsform_to_questionnaire.py` produces a clean, table-free printable interview questionnaire.

That server round trip is the whole reason this exists. The old way to get a paper copy of an ODK, SurveyCTO, or KoboToolbox form was to upload the spreadsheet, export the platform's printable HTML, then run a post-processing script over it to make the tables usable on paper. Two of those scripts still live here (`html_make_responsive.py` and `html_make_responsive_print.py`) and still work on exported HTML. The direct converters skip the upload and read the `survey`, `choices`, and `settings` sheets directly.

## Codebook vs. Questionnaire

| Feature | Codebook (`xlsform_to_printable.py`) | Interview Questionnaire (`Questionnaire/xlsform_to_questionnaire.py`) |
| --- | --- | --- |
| Primary audience | Data managers, research PIs, programmers, M&E officers | Field enumerators, survey respondents, IRB review boards, training sessions |
| Document layout | 3-column fixed table (Field 22%, Question 50%, Answer 28%) | Table-free flowing document with question prompts, notes, choices, and blanks |
| Default page setup | A4 landscape (10mm margins) | A4 portrait (15mm top/bottom, 12mm/16mm margins; `--landscape` optional) |
| Variable names & logic | Variable names, required flags, calculation badges, formulas, and relevance rules | Omitted; presents only human-readable question text and instructions |
| Choice lists | Choice values (codes) paired with labels in a nested table | Radio `( )` or checkbox `[ ]` markers with labels only; raw codes hidden |
| Open-ended inputs | Empty answer cell (data type shown in Field column) | Tailored handwriting blanks (ruled lines, date/time formats, numeric brackets) |
| Notes (`note` type) | Standard table row with note text in Question column | Full-width callout box with a blue accent bar for scripts and consent forms |
| Calculations & metadata | Fully documented with plain-language explanations and badges | Excluded completely |
| Output formats | Responsive HTML, landscape PDF, editable Word `.docx` | Responsive HTML, portrait/landscape PDF, editable Word `.docx` |

## Setup

Python 3.8 or newer, then:

```bash
pip install -r requirements.txt
playwright install chromium
```

Chromium is only needed for PDF rendering. If you pass `--no-pdf`, you can skip it. `python-docx` is only needed for the Word file, and both scripts say so and carry on with HTML and PDF if the import fails, so `--no-docx` runs fine without it.

## Usage

### Generating codebooks (xlsform_to_printable.py)

Every language in the form, HTML, PDF, and DOCX:

```bash
python xlsform_to_printable.py --input "Brick Kiln Literacy RCT - Household Survey.xlsx"
```

One language only. The match is a case-insensitive substring, so `Odia` matches the `Odia (or)` column:

```bash
python xlsform_to_printable.py --input "Brick Kiln Literacy RCT - Household Survey.xlsx" --lang Odia
```

Every `.xlsx` in a folder, skipping Excel lock files:

```bash
python xlsform_to_printable.py --dir /path/to/forms/
```

HTML and DOCX only, no browser needed:

```bash
python xlsform_to_printable.py --input form.xlsx --no-pdf
```

Skip the Word file and keep the print-ready PDF:

```bash
python xlsform_to_printable.py --input form.xlsx --no-docx
```

With no `--input` and no `--dir`, the script processes every `.xlsx` in the current working directory.

### Generating interview questionnaires (Questionnaire/xlsform_to_questionnaire.py)

Every language in the form, HTML, portrait PDF, and portrait DOCX:

```bash
python Questionnaire/xlsform_to_questionnaire.py --input "Brick Kiln Literacy RCT - Household Survey.xlsx"
```

One language only (case-insensitive substring match):

```bash
python Questionnaire/xlsform_to_questionnaire.py --input "Brick Kiln Literacy RCT - Household Survey.xlsx" --lang English
```

Render in A4 landscape instead of default portrait:

```bash
python Questionnaire/xlsform_to_questionnaire.py --input form.xlsx --landscape
```

Every `.xlsx` in a folder, skipping Excel lock files:

```bash
python Questionnaire/xlsform_to_questionnaire.py --dir /path/to/forms/
```

HTML and DOCX only (skip Playwright/PDF rendering):

```bash
python Questionnaire/xlsform_to_questionnaire.py --input form.xlsx --no-pdf
```

Skip DOCX generation and keep HTML and PDF:

```bash
python Questionnaire/xlsform_to_questionnaire.py --input form.xlsx --no-docx
```

With no `--input` and no `--dir`, the script processes every `.xlsx` in the current working directory. Output files sit next to the source spreadsheet and take the language name as a suffix: `<form>_<Language>.html`, `<form>_<Language>.pdf`, and `<form>_<Language>.docx`.

## What the codebook looks like

Three columns, fixed width, so a row never blows out the page: Field at 22%, Question at 50%, Answer at 28%.

The Field column holds the variable name, with a red `(required)` marker when the `required` column says yes, 1, true, ok, or required.

The Question column stacks whatever applies to the row, in this order:

- The question label in the target language.
- A metadata note, if the row's type is a platform variable like `start` or `deviceid`. See the table below for what each one says.
- A calculation badge naming the formula's category, followed by the formula itself trimmed to 65 characters. The full formula lives in the `title` attribute of the `<code>` element, so it survives in HTML on hover and stays out of the way in print. Left whole, a long `indexed-repeat()` or a nested `if()` stretches its row five or six lines deep.
- The hint text.
- The relevance expression, prefixed with `Relevance:`.

The Answer column holds a nested two-column table for `select_one`, `select_multiple`, and `select_or_other` fields: a 32px code cell on the left, the choice label wrapping freely on the right. The choice list lookup is case-insensitive against the `choices` sheet, and a choice with no label falls back to its own value. Types that read options from an external file leave the Answer column empty, since those options are not in the sheet.

`begin group`, `begin_group`, `begin repeat`, and `begin_repeat` rows become a full-width section band carrying the group label, or the group name when there is no label. The matching end rows are dropped.

`constraint` and `constraint_message` are parsed but not yet rendered.

## What the interview questionnaire looks like

The questionnaire script produces a paper-first, table-free interview document designed for field enumerators, respondents, and ethics reviews:

- Document header banner: Displays the form title from the `settings` sheet (or filename) with an interview language subtitle and divider rule.
- Section headers: `begin group` and `begin repeat` rows become full-width tinted header bands (`#f0f4f8` with 2px borders) marking survey modules. Closing tags (`end group`, `end repeat`) are dropped.
- Notes and instructions: `note` rows become full-width callout blocks with a 4px left accent bar (`#3b82f6`) and light tint (`#f8fafc`). These house enumerator instructions, consent statements, and read-aloud introductions without breaking question flow.
- Questions: Question prompts display in bold text with clean paragraph spacing. Hints are rendered in italic blue text indented below the question (`Note: <hint>`). Technical variable names, required flags, relevance conditions, and constraints are omitted.
- Multiple-choice questions: `select_one` rows display with radio markers `( )`; `select_multiple` rows display with checkbox markers `[ ]`. Technical choice values (e.g. numeric codes like `1`, `2`, `-99`) are stripped so only the localized choice labels appear.
- Open-ended response blanks: Non-choice questions display tailored handwriting blanks matching the field's data type, avoiding plain blank lines.
- Excluded elements: System metadata (`start`, `end`, `deviceid`, `today`, `username`, `email`, `audit`, etc.) and internal calculation rows (`calculate`, `hidden`) are completely excluded from the questionnaire.

### Response blanks by question type

When rendering open-ended questions in the questionnaire, `xlsform_to_questionnaire.py` provides formatted handwriting blanks corresponding to each XLSForm data type:

| Question Type | Rendered Response Blank | Notes |
| --- | --- | --- |
| `text` (standard) | `____________________________________________________________________` | Single ruled line for short free-text answers |
| `text` (detailed/notes) | Double ruled line `____________________________________________________________________` | Triggered when variable name or label mentions notes, details, comments, specify, other, explain, or reasons |
| `integer`, `decimal`, `range` | `[ ____________________ ]` | Bracketed blank for numeric values |
| `date` | `[ ____ / ____ / ________ ] (DD/MM/YYYY)` | Segmented date blank |
| `time` | `[ ____ : ____ ] (HH:MM)` | Segmented 24-hour time blank |
| `datetime` | `[ ____/____/________  ____:____ ]` | Combined timestamp blank |
| `geopoint`, `geotrace`, `geoshape` | `[ GPS Location Coordinates: ________________________________ ]` | GPS capture indicator |
| `image`, `audio`, `video`, `file` | `[ Attach / Record <Type>: ________________________________ ]` | Media attachment placeholder |
| `barcode` | `[ Barcode / QR: ________________________________ ]` | Barcode or QR scanner placeholder |
| `acknowledge` | `[  ] Confirmed / Acknowledged` | Single acknowledgment checkbox |

## Languages

The parser reads the header rows of the `survey` and `choices` sheets and collects every `label::` column it finds, so `label::English (en)` and `label::Odia (or)` both become target languages with no configuration. Untagged `label`, `hint`, and `constraint_message` columns are used as the fallback for any language that lacks its own column, and a form with no tagged labels at all is treated as a single `Default` language.

Filename suffixes drop the language code in parentheses, so `label::English (en)` produces `_English.html`, not `_English (en).html`.

## Platform metadata reference

When a row's type matches one of these, the codebook Question column explains it in plain language instead of leaving the cell blank. (These rows are skipped entirely in the interview questionnaire.)

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

Formulas are matched against these patterns in order, and the first hit wins. Anything unmatched is labelled plain `Calculation`. (Calculations are skipped entirely in the interview questionnaire.)

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

Order matters here. A formula that wraps `now()` inside `format-date-time()` is tagged as a timestamp capture, because the `now()` test runs first.

## Print behaviour

The screen layout and the print layout come from the same stylesheet in both scripts, with the print rules in a `@media print` block. The DOCX renderers mirror these decisions in Word's own vocabulary; see the section below.

### Codebook print layout (`xlsform_to_printable.py`)

- `@page` is A4 landscape with 8mm vertical and 10mm horizontal margins. The Playwright call passes 10mm on all four sides, and that is what the finished PDF uses.
- `thead` is a `table-header-group`, so the Field, Question, and Answer headers reappear at the top of each page.
- `page-break-inside: avoid` on rows keeps a question and its choice list together.
- Section bands and column headers switch from dark navy on screen (`#2b2d42` and `#4a4e69`, white text) to light tints in print (`#f0f4f8` with 1.5pt rules above and below, `#e9ecef` for column headers, black text). Paper gets the same visual hierarchy without a solid bar of toner on every section. Both states set `print-color-adjust: exact` so browsers do not helpfully drop the fills.
- Type is 9.5pt on 13.5pt in print, and 14px body with 13 to 13.5px cells on screen.
- Every cell sets `overflow-wrap: break-word` and `word-break: break-word`, so long variable names like `start-geopoint` wrap without a horizontal scrollbar and without splitting one character per line.
- Screen breakpoints at 768px and 480px shrink the type and let `.table-responsive` scroll sideways rather than crush the columns.
- The scripts read and write UTF-8 throughout, and the font stack leads with `"Arial Unicode MS"` so Odia, Devanagari, Tamil, and Bengali labels render instead of showing boxes.

### Questionnaire print layout (`xlsform_to_questionnaire.py`)

- `@page` is A4 portrait (`size: A4 portrait; margin: 15mm 12mm;`) by default, switching to landscape if `--landscape` is passed.
- `page-break-inside: avoid` on `.question-item` and `.note-item` prevents questions and note blocks from breaking across page margins.
- `page-break-after: avoid` on `.section-header` ensures section headers never get stranded as orphans at the bottom of a page.
- Section headers use light tint (`#f0f4f8`) with 1.5pt top and bottom rules (`#333333`) and `print-color-adjust: exact`.
- Note blocks preserve their light background tint (`#f8fafc`) and 3pt blue left accent border (`#3b82f6`) on paper.
- Container width is capped at 820px for portrait (1100px for landscape) on screen, centered with subtle drop shadow, and expands to full printable width in `@media print`.

## Editable DOCX output

The Word files are built directly from parsed survey data using `python-docx`, not by converting HTML or PDF. Nothing is scraped back out of a rendered file, so the outputs cannot drift apart.

### Codebook Word document (`xlsform_to_printable.py`)

- One three-column table, fixed layout, matching the HTML proportions: 22%, 50%, 28% of the 277mm text block, so a Word column measures 172.8pt, 392.6pt, and 219.9pt.
- A4 landscape with 10mm margins, the same page setup the PDF uses.
- The Field, Question, and Answer row is marked as a repeating header row, so it reappears on every page the way `thead` does in print.
- Rows carry `cantSplit`, which is the Word equivalent of `page-break-inside: avoid`, keeping a question with its choice list.
- Group and repeat rows become a band merged across all three columns, tinted `#F0F4F8` with 1.5pt rules above and below, matching the print stylesheet rather than the darker screen colour.
- Metadata notes and calculation badges keep their tint and left accent bar, drawn as paragraph shading and a paragraph border.
- Choice lists are a borderless nested table inside the Answer cell, a narrow centred code column beside the wrapping label, with dotted row rules.
- Long formulas are trimmed to 65 characters, as in the PDF.

### Questionnaire Word document (`xlsform_to_questionnaire.py`)

- Table-free flowing document structure in A4 portrait (210mm x 297mm) with 16mm left/right and 15mm top/bottom margins (or A4 landscape with 297mm x 210mm when `--landscape` is passed).
- Question prompts use 10.5pt bold text with paragraph spacing and `keepWithNext` (`w:keepNext`) enabled so questions stay attached to their choices and response blanks across page breaks.
- Hints are indented 10pt with italic blue text (`#1E40AF`) and `keepWithNext` enabled.
- Choice lists indent each option by 13pt with `( )` or `[ ]` markers in muted slate (`#64748B`) followed by the choice label, with `keepWithNext` on the leading choices to avoid page splits.
- Open-ended response blanks are indented 12pt using muted ruled lines or brackets.
- Section headers use 12pt bold text inside a full-width shaded paragraph (`#F0F4F8`) with top and bottom rules (`#334155`) and `keepWithNext` enabled.
- Note callout blocks use 9.5pt text with light gray shading (`#F8FAFC`), 7pt left margin indentation, and a 3pt left paragraph border (`#3B82F6`).

### Typography and schema compliance

- Every run in both DOCX generators names `Arial Unicode MS` on all four of `w:ascii`, `w:hAnsi`, `w:cs`, and `w:eastAsia`, and sets the complex-script twins `w:szCs`, `w:bCs`, and `w:iCs`. Without those, Word renders Devanagari, Odia, Tamil, and Bengali runs at a default 10pt regular instead of the size and weight requested.
- Word validates property elements by position, not merely by presence. Both scripts use dedicated helpers (`_docx_set_child`) with ordered property schemas (`_DOCX_PPR_ORDER`, `_DOCX_RPR_ORDER`) to insert elements at their exact schema slots, eliminating Word's unreadable-content repair prompt on open.

## Repository files

- `xlsform_to_printable.py`: the direct codebook converter. Parses `.xlsx` files, builds 3-column HTML tables, renders landscape A4 PDFs, and writes landscape DOCX codebooks.
- `Questionnaire/xlsform_to_questionnaire.py`: the direct interview questionnaire converter. Parses `.xlsx` files and produces clean, table-free flowing HTML, portrait A4 PDFs, and portrait DOCX questionnaires for field administration.
- `Questionnaire/Pradhan_Qualitative_Survey_2026.xlsx`: sample questionnaire spreadsheet for qualitative field interviews.
- `html_make_responsive.py`: post-processor for HTML exported from a survey platform. Keeps the original table layout, injects responsive and print CSS, writes a landscape A4 PDF.
- `html_make_responsive_print.py`: alternative post-processor that rebuilds exported tables as card-style question blocks for portrait A4.
- `backup_html_make_responsive.py`: frozen reference copy of the post-processor. Not an entry point.
- `xlsform_to_printable_plan.md`: architecture notes, the XLSForm element mapping matrix, and the implementation roadmap.
- `requirements.txt`: the four third-party packages, pinned. Chromium is not in here; it comes from `playwright install chromium`.
- `Brick Kiln Literacy RCT - Household Survey.xlsx`: sample form with English and Odia labels.

Both post-processing scripts take no arguments. They process every `.html` file in the current working directory, overwrite each one with the responsive version, and write a PDF beside it:

```bash
python html_make_responsive.py
```

## Writing docs in this repo

Markdown paragraphs and list items are single unwrapped lines. Do not hard-wrap prose at a column limit or break a line inside a paragraph or bullet.
