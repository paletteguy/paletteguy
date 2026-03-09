# CV Profile Generator

## Project Overview
CV/resume generator for **Karsten Sperling Opdal** — produces `.docx`, `.pdf`, and `README.md` from a single Python script + JSON data file.

## Files
- `cv_data.json` — Single source of truth for all CV content
- `generate_cv.py` — Generator script (python-docx). Run with `python3 generate_cv.py`
- `README.md` — Auto-generated markdown CV (do not edit directly)
- `me.jpeg` — Profile photo used in .docx output
- `Karsten Opdal - CV.docx` — Generated output (do not edit directly)
- `Karsten Opdal - CV.pdf` — Generated via LibreOffice headless (do not edit directly)

## Dependencies
- `python-docx` — .docx generation
- `lxml` (via python-docx) — XML manipulation for advanced formatting
- LibreOffice — PDF conversion (headless mode)

## Architecture & Patterns

### Data Flow
`cv_data.json` → `generate_cv.py` → `.docx` + `README.md` + `.pdf`

All CV content lives in `cv_data.json`. The script loads it and feeds it to `generate_docx()` and `generate_md()`. Never duplicate content between files.

### Experience Periods
Experience entries use `"start"` and `"end"` fields (e.g. `"October 2010"`, `"Present"`). The `format_period()` function calculates durations automatically. Education entries use a plain `"period"` string (no calculation). All month names must be in English.

### Docx Generation
- **Name**: Standalone `doc.add_heading(level=0)` (Title style) above the table — NOT inside a table cell
- **Title underline removal**: Explicit `w:pBdr` with "none" borders via XML
- **Header table**: 2-column table (contact left, photo right) with fixed layout
  - Contact column: PAGE_WIDTH - PHOTO_COL_WIDTH
  - Photo column: 1.6 inches, image 1.5 inches, centered
  - Zero cell margins via `w:tblCellMar` to prevent photo clipping
  - Column widths set via `w:tblGrid` gridCol elements + cell-level `w:tcW`
- **Table borders**: Removed at 3 levels: table style removed, `w:tblBorders` set to "none", cell-level `w:tcBorders` set to "none"
- **Named constants**: All sizes, fonts, and colors defined at top of `generate_docx()`
- **Headings**: Custom `add_heading_styled()` with color `#2E4A62`
- **Experience entries**: `add_experience()` helper — company (bold 12pt), title (italic 10.5pt), period (gray 9pt), bullet list
- **Page breaks**: Controlled via `"page_break_after": true` in experience data + page break before Experience section
- **Font**: Calibri 10.5pt default

### Markdown Generation
- Plain markdown, no HTML tables or photo references
- Trailing `  ` (double space) for line breaks in contact details and languages
- Sections joined with `\n\n---\n\n` separator
- Education and experience use `###` subheadings with italic descriptions

### Known Issues
- **LibreOffice table borders**: Despite all XML border removal, LibreOffice may still show faint table boundaries. These are display artifacts (View > Table Boundaries), not actual printed borders.
- **EMU arithmetic**: `Inches()` returns EMU objects but arithmetic on them returns plain `int`, losing `.inches` property. Use raw float math for twips: `str(int(inches_value * 1440))`

## Content Notes
- Github repos section comes before Education
- `"page_break_after"` in experience entries controls docx pagination (ignored by markdown)
