# XLSForm to printable codebook

Print your survey before you field it. Point `xlsform_to_printable.py` at an XLSForm `.xlsx` file and you get a responsive HTML codebook, an A4 landscape PDF, and an editable Word `.docx` for every language in the form, without uploading anything to a server.

That server round trip is the whole reason this exists. The old way to get a paper copy of an ODK, SurveyCTO, or KoboToolbox form was to upload the spreadsheet, export the platform's printable HTML, then run a post-processing script over it to make the tables usable on paper. Two of those scripts still live here (`html_make_responsive.py` and `html_make_responsive_print.py`) and still work on exported HTML. The direct converter skips the upload and reads the `survey`, `choices`, and `settings` sheets itself.

## Setup

Python 3.8 or newer, then:

```bash
pip install -r requirements.txt
playwright install chromium
```

Chromium is only needed for PDF rendering. If you pass `--no-pdf`, you can skip it. `python-docx` is only needed for the Word file, and the script says so and carries on with HTML and PDF if the import fails, so `--no-docx` runs fine without it.

## Usage

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

Output files sit next to the source spreadsheet and take the language name as a suffix: `Brick Kiln Literacy RCT - Household Survey_English.html`, `..._Odia.html`, and the matching `.pdf` and `.docx` files. A form with no `label::` columns produces one unsuffixed set.

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

## Languages

The parser reads the header rows of the `survey` and `choices` sheets and collects every `label::` column it finds, so `label::English (en)` and `label::Odia (or)` both become target languages with no configuration. Untagged `label`, `hint`, and `constraint_message` columns are used as the fallback for any language that lacks its own column, and a form with no tagged labels at all is treated as a single `Default` language.

Filename suffixes drop the language code in parentheses, so `label::English (en)` produces `_English.html`, not `_English (en).html`.

## Platform metadata reference

When a row's type matches one of these, the Question column explains it in plain language instead of leaving the cell blank.

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

Formulas are matched against these patterns in order, and the first hit wins. Anything unmatched is labelled plain `Calculation`.

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

The screen layout and the print layout come from the same stylesheet, `RESPONSIVE_STYLES`, with the print rules in a `@media print` block. The DOCX is a separate renderer and mirrors these decisions in Word's own vocabulary; see the section below.

- `@page` is A4 landscape with 8mm vertical and 10mm horizontal margins. The Playwright call passes 10mm on all four sides, and that is what the finished PDF uses.
- `thead` is a `table-header-group`, so the Field, Question, and Answer headers reappear at the top of each page.
- `page-break-inside: avoid` on rows keeps a question and its choice list together.
- Section bands and column headers switch from dark navy on screen (`#2b2d42` and `#4a4e69`, white text) to light tints in print (`#f0f4f8` with 1.5pt rules above and below, `#e9ecef` for column headers, black text). Paper gets the same visual hierarchy without a solid bar of toner on every section. Both states set `print-color-adjust: exact` so browsers do not helpfully drop the fills.
- Type is 9.5pt on 13.5pt in print, and 14px body with 13 to 13.5px cells on screen.
- Every cell sets `overflow-wrap: break-word` and `word-break: break-word`, so long variable names like `start-geopoint` wrap without a horizontal scrollbar and without splitting one character per line.
- Screen breakpoints at 768px and 480px shrink the type and let `.table-responsive` scroll sideways rather than crush the columns.
- The scripts read and write UTF-8 throughout, and the font stack leads with `"Arial Unicode MS"` so Odia, Devanagari, Tamil, and Bengali labels render instead of showing boxes.

## Editable DOCX output

The Word file is built by `build_codebook_docx()` from the same parsed survey rows and choice lists that feed the HTML, not by converting the HTML or the PDF. Nothing is scraped back out of a rendered file, so the three outputs cannot drift apart, and every element the HTML shows in a row appears in the same order in the Word table.

It is a real Word table, so a research manager can retitle a section, reword a prompt, add an enumerator instruction, or drop a question, and the layout reflows on its own.

- One three-column table, fixed layout, matching the HTML proportions: 22%, 50%, 28% of the 277mm text block, so a Word column measures 172.8pt, 392.6pt, and 219.9pt.
- A4 landscape with 10mm margins, the same page setup the PDF uses.
- The Field, Question, and Answer row is marked as a repeating header row, so it reappears on every page the way `thead` does in print.
- Rows carry `cantSplit`, which is the Word equivalent of `page-break-inside: avoid`, keeping a question with its choice list.
- Group and repeat rows become a band merged across all three columns, tinted `#F0F4F8` with 1.5pt rules above and below, matching the print stylesheet rather than the darker screen colour.
- Metadata notes and calculation badges keep their tint and left accent bar, drawn as paragraph shading and a paragraph border.
- Choice lists are a borderless nested table inside the Answer cell, a narrow centred code column beside the wrapping label, with dotted row rules.
- Every run names `Arial Unicode MS` on all four of `w:ascii`, `w:hAnsi`, `w:cs`, and `w:eastAsia`, and sets the complex-script twins `w:szCs`, `w:bCs`, and `w:iCs`. Without those, Word renders Devanagari, Odia, Tamil, and Bengali runs at a default 10pt regular instead of the size and weight the row asked for.

Two things the Word file cannot carry over from the HTML. Long formulas are trimmed to 65 characters, as in the PDF, and there is no hover text to hold the rest, since `title` attributes have no Word equivalent. The `--no-pdf` and `--no-docx` flags are independent, so the source spreadsheet is still the only complete record of a long `pulldata()` or nested `if()`.

Word validates property elements by position, not merely by presence, so the helpers that write raw WordprocessingML insert each child at its schema slot instead of appending. Getting this wrong is what produces the unreadable-content repair prompt on open.

## Repository files

- `xlsform_to_printable.py`: the direct converter. Parses the `.xlsx`, builds the HTML, renders the PDF, writes the DOCX.
- `html_make_responsive.py`: post-processor for HTML exported from a survey platform. Keeps the original table layout, injects responsive and print CSS, writes a landscape A4 PDF.
- `html_make_responsive_print.py`: alternative post-processor that rebuilds the exported tables as card-style question blocks for portrait A4.
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
