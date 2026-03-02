# CV Profile Generator

## Project Overview
CV/resume generator for **Karsten Sperling Opdal** — produces a `.docx` file and a matching `README.md` from a single Python script.

## Files
- `generate_cv.py` — Main generator script (python-docx + Pillow). Run with `python3 generate_cv.py`
- `README.md` — Markdown version of the CV (synced manually with generate_cv.py)
- `me.png` — Profile photo used in both outputs
- `Karsten Opdal - CV.docx` — Generated output (do not edit directly)
- `Karsten Opdal - CV edited.docx` — User's manually edited reference (used for syncing changes)

## Dependencies
- `python-docx` — .docx generation
- `Pillow` (PIL) — Image resizing before insertion
- `lxml` (via python-docx) — XML manipulation for advanced formatting

## Architecture & Patterns

### Docx Generation
- **Name**: Standalone `doc.add_heading(level=0)` (Title style) above the table — NOT inside a table cell (font size doesn't render correctly in cells)
- **Title underline removal**: Explicit `w:pBdr` with "none" borders via XML to remove the blue bottom border that Title style adds
- **Header table**: 2-column table (contact left, photo right) with fixed layout (`w:tblLayout type="fixed"`)
  - Left column: ~5.1 inches (contact details)
  - Right column: 1.4 inches (photo)
  - Column widths set via `w:tblGrid` gridCol elements + cell-level `w:tcW`
- **Table borders**: Removed at 3 levels: table style removed, `w:tblBorders` set to "none", cell-level `w:tcBorders` set to "none"
- **Headings**: Custom `add_heading_styled()` with color `#2E4A62`
- **Experience entries**: `add_experience()` helper — company (bold 12pt), title (italic 10.5pt), period (gray 9pt), bullet list
- **Page breaks**: Manually placed between experience groups
- **Font**: Calibri 10.5pt default

### Known Issues
- **LibreOffice table borders**: Despite all XML border removal, LibreOffice may still show faint table boundaries. These are display artifacts (View > Table Boundaries), not actual printed borders.
- **EMU arithmetic**: `Inches()` returns EMU objects but arithmetic on them (`Inches(6.5) - Inches(1.4)`) returns plain `int`, losing `.inches` property. Use raw float math for twips: `str(int(inches_value * 1440))`

### Keeping Files in Sync
When updating CV content, update **both** `generate_cv.py` and `README.md`. The .md uses:
- HTML `<table>` for header layout (photo left, contact right)
- `<br>` tags for line breaks inside HTML table cells
- Standard markdown for everything else

## Content Notes
- Experience order: Laerdal → Oscilloscope → Nokia Specialist → Nokia Mobile → Infineon → Thorsø → Tang-Data → Nokia SE → EC-Soft Danmark Senior → Telenor → EC-Soft Norge → Nokia Contract → EC-Soft Danmark Consultant → Vizion → Greve → Opdal Enterprise → Home development
- Core Skills: `C++ (C++23) · Embedded Linux · Yocto · Linux kernel · Qt/QML (Qt 3–6) · Python · WiFi/Bluetooth · CI/CD · Boost · AI-assisted engineering`
- Github repos section comes before Education
