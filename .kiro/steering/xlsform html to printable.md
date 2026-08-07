---
inclusion: always
---

# XLSForm HTML to Printable Codebook

## Purpose

This workspace converts XLSForm spreadsheets (`.xlsx`) directly into responsive, print-ready HTML questionnaires, PDFs, and editable Word DOCX codebooks across ODK, SurveyCTO, and KoboToolbox platforms, as well as post-processing legacy SurveyCTO exported HTML codebooks.

## Layout

- `xlsform_to_printable.py` — primary direct converter script. Parses XLSForm `.xlsx` files directly, normalizes multi-platform syntax, auto-detects all languages, explains platform metadata (`start`, `end`, `deviceid`, `username`, `audit`), categorizes calculation formulas (`pulldata`, `jr:choice-name`, `once(now())`, `format-date-time`, `decimal-date-time`, `substr`), renders choice tables in the Answer column, embeds responsive/print CSS, and outputs HTML, PDF, and DOCX codebooks for all languages.
- `html_make_responsive.py` — legacy script. Keeps original SurveyCTO table layout, injects responsive + print CSS, then renders a landscape A4 PDF via Playwright/Chromium.
- `html_make_responsive_print.py` — alternative legacy script. Restructures SurveyCTO tables into card-style `.question-item` blocks for portrait A4 printing.
- `backup_html_make_responsive.py` — frozen reference copy. Do not treat as an entry point and do not edit unless asked.
- `*_English.html` / `*_Hindi.html` — generated HTML codebooks, one pair per form.
- `*_English.pdf` / `*_Hindi.pdf` — generated landscape A4 PDF codebooks.
- `*_English.docx` / `*_Hindi.docx` — generated editable landscape A4 Word codebooks.
- `*.xlsx` — source XLSForm spreadsheets.
- `requirements.txt` — pinned third-party dependencies (`openpyxl`, `beautifulsoup4`, `playwright`, `python-docx`). Update this alongside any dependency change.
- `README.md` — project overview, setup, multi-platform & multi-language details, platform metadata descriptions, calculation classification reference, choice table parsing, and CLI usage.
- `xlsform_to_printable_plan.md` — technical architecture, XLSForm element mapping matrix, platform metadata reference, calculation categories, zero-index safe choice parsing, and implementation roadmap.

## Key Technical Specifications & Conventions

- Each script is self-contained: CSS lives in a module-level string constant (`RESPONSIVE_STYLES`, `PRINT_STYLES`, `BASE_STYLES`), transformation logic in `build_codebook_html()` and `build_codebook_docx()`, and batch I/O in `process_xlsform()`.
- DOCX is built from the parsed XLSForm model, never by converting the rendered HTML or PDF. Both renderers share `split_type()`, `is_section_start()`, `is_section_end()`, `trim_calculation()`, `get_lang_value()`, `classify_calculation()`, and `METADATA_DESCRIPTIONS`, so the outputs cannot drift.
- `python-docx` is an optional import guarded by `DOCX_AVAILABLE`. A missing install prints a hint and skips DOCX; it must never break HTML or PDF output.
- Raw WordprocessingML written by hand goes through `_docx_set_child()` with the matching `_DOCX_*_ORDER` tuple. Word validates property children by position, and appending out of order triggers the repair prompt on open.
- Every DOCX run goes through `_docx_style_run()`, which names the font on `w:ascii`, `w:hAnsi`, `w:cs`, and `w:eastAsia` and sets `w:szCs`, `w:bCs`, and `w:iCs`. Devanagari and Odia runs lose their size and weight without the complex-script twins.
- Call `_docx_apply_row_widths()` on every row added after `_docx_fixed_layout()`. `Table.add_row()` stamps an even split of the original table width onto new cells and overrides the fixed grid.
- Always read and write with `encoding='utf-8'` — Hindi (Devanagari) content depends on it. Keep the `"Arial Unicode MS"` font stack ahead of Arial for the same reason.
- Section headers: styled as `.section-header-dark` with `#1a1a1a` background and white text, configured with `-webkit-print-color-adjust: exact` and `print-color-adjust: exact`.
- Choice tables: nested `<table class="table borderless">` inside the Answer column mapping choice values (`name`) to choice labels for each language.
- Metadata tags: rendered inside `<div class="metadata-tag">` containers explaining platform metadata variables.
- Calculation formulas: rendered inside `<div class="calculation">` containers with category badges and explanations for all calculation variables.
- Guard against page splits with `page-break-inside: avoid` on rows, question blocks, and choice items, and with `w:cantSplit` on DOCX rows.
- Standard library plus `beautifulsoup4`, `openpyxl`, `playwright`, and `python-docx` only.

## Print Targets

- `xlsform_to_printable.py` & `html_make_responsive.py`: A4 **landscape**, 10mm margins, `print_background=True`.
- `xlsform_to_printable.py` DOCX: A4 **landscape**, 10mm margins, 277mm text block split 22% / 50% / 28%, repeating header row.
- `html_make_responsive_print.py`: A4 **portrait**, 15mm/12mm margins.

## Markdown in This Project

Write Markdown paragraphs and list items as a single unwrapped line each. Do not hard-wrap prose at a column limit, and do not insert line breaks inside a paragraph or bullet.
