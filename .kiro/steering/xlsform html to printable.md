---
inclusion: always
---

# XLSForm HTML to Printable Codebook

## Purpose

This workspace converts SurveyCTO/XLSForm-exported HTML codebooks into responsive,
print-ready questionnaires (and optionally PDFs). Source `.xlsx` XLSForms sit alongside
their exported `.html` siblings for reference.

## Layout

- `html_make_responsive.py` — primary script. Keeps the original table layout, injects
  responsive + print CSS, then renders a landscape A4 PDF via Playwright/Chromium.
- `html_make_responsive_print.py` — alternative script. Restructures tables into
  card-style `.question-item` blocks for portrait A4 printing. No PDF step.
- `backup_html_make_responsive.py` — frozen reference copy. Do not treat as an entry point
  and do not edit unless asked.
- `*_English.html` / `*_Hindi.html` — exported codebooks, one pair per form.
- `*.xlsx` — the source XLSForms the HTML was generated from.

Each script is self-contained: CSS lives in a module-level string constant
(`RESPONSIVE_STYLES`, `PRINT_STYLES`, `BASE_STYLES`), transformation logic in
`make_html_responsive(html_content) -> str`, and batch I/O in `process_directory()`.
Preserve this shape when extending; do not split CSS into external stylesheets, since the
deliverable must be a single portable HTML file.

## Critical Behavior

Both scripts **overwrite the HTML files in place**. There is no backup step. Before
running or modifying a script, confirm with the user, and prefer copying a sample file to
a scratch name for testing over running `process_directory()` across the whole folder.

Transformations are also **not idempotent** — re-running on already-processed output can
double-wrap containers or strip structure the parser expected. Assume input is a fresh
export unless proven otherwise.

## Source HTML Structure

Understand the export format before writing selectors:

- Outer wrapper: `<div style="margin: auto; width: 1000px">` — becomes
  `.codebook-container` / `.questionnaire-container`.
- Title: a single `<h4>`.
- Main table: `table.table.table-bordered.table-condensed` with a three-column
  `thead` (Field / Question / Answer) and cells classed `.fieldCell`, `.questionCell`,
  `.answerCell`.
- Section headers: a row whose single `<td>` carries `colspan="3"` plus inline
  `background-color: #707070` (or `#8C8C8C`) and `color: #FFFFFF`.
- Choice lists: a nested `table` inside the answer cell; columns are
  `.response-note-cell`, the choice value, then the choice label.
- Indentation for grouped fields uses nested `table.borderless` with `.spacer` cells.
- Other meaningful classes: `.required`, `.hint`, `.relevance`, `.constraint`,
  `.noteHolder`, `img.questionPrompt`, `img.choicePrompt`, `tr.gray`, `tr.entryRow`.

## Conventions

- Parse with BeautifulSoup (`'html.parser'`); build nodes with `soup.new_tag()` and
  mutate via `wrap()` / `replace_with()`. Never regex over HTML markup.
- Always read and write with `encoding='utf-8'` — Hindi (Devanagari) content depends on it.
  Keep the `"Arial Unicode MS"` font stack ahead of Arial for the same reason.
- Match inline styles defensively. SurveyCTO emits colors as `#707070`, `#8C8C8C`,
  `rgb(112, 112, 112)`, and `rgb(140, 140, 140)`, with and without a space after the colon.
  Add every variant when introducing a new selector.
- Preferred pattern for restyling exported cells: strip the offending inline
  `background-color` / `color` declarations, attach a semantic class such as
  `.section-header-dark`, then style via the injected CSS. Avoid fighting inline styles
  with `!important` alone.
- Any background or color that must survive printing needs both
  `-webkit-print-color-adjust: exact` and `print-color-adjust: exact`.
- Guard against page splits with `page-break-inside: avoid` on rows, question blocks,
  and choice items.
- Wrap every file operation in try/except and print a short per-file progress line
  (`Processing <name>...`, then the outcome). Batch runs must not abort on one bad file.
- Use `asyncio.run()` per PDF; keep Playwright calls inside `async with async_playwright()`.
- Standard library plus `beautifulsoup4` and `playwright` only. Ask before adding
  dependencies.

## Print Targets

- `html_make_responsive.py`: A4 **landscape**, 10mm margins, `print_background=True`,
  body ~9pt.
- `html_make_responsive_print.py`: A4 **portrait**, 15mm/12mm margins, body 10pt.

Screen breakpoints in the responsive script are 768px and 480px; the mobile path relies on
a `.table-responsive` wrapper with horizontal scroll rather than reflowing columns.

## When Making Changes

- Styling-only requests: edit the relevant CSS constant, nothing else.
- Structural requests: change `make_html_responsive()` and its helpers, and keep the
  three-column fallback chain (`num_cols >= 3`, `== 2`, `== 1`) intact.
- Playwright needs its browser binary; if PDF generation fails, check
  `playwright install chromium` before debugging the script.
- Verify output by opening the generated HTML/PDF, not just by a clean exit code.
