#!/usr/bin/env python3

import json
import os
import subprocess
from datetime import date
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

MONTHS = {
    'january': 1, 'february': 2, 'march': 3, 'april': 4,
    'may': 5, 'june': 6, 'july': 7, 'august': 8,
    'september': 9, 'october': 10, 'november': 11, 'december': 12,
}


def parse_month_year(s):
    """Parse 'Month Year' into (year, month)."""
    parts = s.strip().split()
    return int(parts[1]), MONTHS[parts[0].lower()]


def format_period(start_str, end_str):
    """Format 'Month Year – Month Year (duration)' from start/end strings."""
    start_y, start_m = parse_month_year(start_str)
    if end_str.lower() == 'present':
        today = date.today()
        end_y, end_m = today.year, today.month
        end_label = 'Present'
    else:
        end_y, end_m = parse_month_year(end_str)
        end_label = end_str

    total_months = (end_y - start_y) * 12 + (end_m - start_m) + 1
    years, months = divmod(total_months, 12)

    parts = []
    if years:
        parts.append(f'{years} year{"s" if years != 1 else ""}')
    if months:
        parts.append(f'{months} month{"s" if months != 1 else ""}')
    duration = ' '.join(parts) if parts else '1 month'

    return f'{start_str} – {end_label} ({duration})'


# ============================================================
# DOCX GENERATION
# ============================================================

def generate_docx(data, photo_path):
    doc = Document()

    # -- Size constants (inches) --
    PAGE_WIDTH = 6.5           # usable width (letter 8.5 minus 1-inch margins)
    PHOTO_COL_WIDTH = 1.6      # right column for photo
    PHOTO_IMG_WIDTH = 1.5      # photo slightly narrower than column to avoid clipping
    CONTACT_COL_WIDTH = PAGE_WIDTH - PHOTO_COL_WIDTH

    # -- Font sizes --
    NAME_SIZE = Pt(28)
    TAGLINE_SIZE = Pt(11)
    DEFAULT_SIZE = Pt(10.5)
    COMPANY_SIZE = Pt(12)
    PERIOD_SIZE = Pt(9)
    CONTACT_LABEL_SIZE = Pt(9)
    CONTACT_VALUE_SIZE = Pt(10)
    CONTACT_LINE_SPACING = Pt(11)

    # -- Colors --
    HEADING_COLOR = RGBColor(0x2E, 0x4A, 0x62)
    GRAY = RGBColor(0x66, 0x66, 0x66)

    # -- Twips for XML (1 inch = 1440 twips) --
    def to_twips(inches):
        return str(int(inches * 1440))

    page_width_twips = to_twips(PAGE_WIDTH)
    contact_col_twips = to_twips(CONTACT_COL_WIDTH)
    photo_col_twips = to_twips(PHOTO_COL_WIDTH)

    # -- Style defaults --
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = DEFAULT_SIZE

    # -- Helper functions --
    def add_heading_styled(text, level=1):
        h = doc.add_heading(text, level=level)
        for run in h.runs:
            run.font.color.rgb = HEADING_COLOR
        return h

    def add_experience(company, title, period, description_lines):
        p = doc.add_paragraph()
        run = p.add_run(company)
        run.bold = True
        run.font.size = COMPANY_SIZE
        p.space_after = Pt(0)

        p2 = doc.add_paragraph()
        run2 = p2.add_run(title)
        run2.italic = True
        run2.font.size = DEFAULT_SIZE
        p2.space_after = Pt(0)

        p3 = doc.add_paragraph()
        run3 = p3.add_run(period)
        run3.font.size = PERIOD_SIZE
        run3.font.color.rgb = GRAY
        p3.space_after = Pt(4)

        for line in description_lines:
            p4 = doc.add_paragraph(line, style='List Bullet')
            p4.paragraph_format.space_after = Pt(2)

        # Small spacer
        doc.add_paragraph().paragraph_format.space_before = Pt(4)

    # -- Header: name --
    h = doc.add_heading(data['name'], level=0)
    h.paragraph_format.space_after = Pt(0)
    h.paragraph_format.space_before = Pt(0)
    for run in h.runs:
        run.font.color.rgb = HEADING_COLOR
        run.font.size = NAME_SIZE
    # Remove the bottom border that Title style adds
    pPr = h._p.get_or_add_pPr()
    pBdr = pPr.find(qn('w:pBdr'))
    if pBdr is not None:
        pPr.remove(pBdr)
    pBdr = pPr.makeelement(qn('w:pBdr'), {})
    for edge in ('top', 'left', 'bottom', 'right'):
        el = pBdr.makeelement(qn(f'w:{edge}'), {
            qn('w:val'): 'none', qn('w:sz'): '0',
            qn('w:space'): '0', qn('w:color'): 'auto',
        })
        pBdr.append(el)
    pPr.append(pBdr)

    # -- Tagline --
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(0)
    run = p.add_run(data['tagline'])
    run.bold = True
    run.italic = True
    run.font.size = TAGLINE_SIZE

    # -- Contact table (left: contact, right: photo) --

    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.LEFT

    tbl = header_table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else tbl._add_tblPr()

    # Remove any table style
    tblStyle = tblPr.find(qn('w:tblStyle'))
    if tblStyle is not None:
        tblPr.remove(tblStyle)

    # Fixed table layout
    tblLayout = tblPr.makeelement(qn('w:tblLayout'), {qn('w:type'): 'fixed'})
    tblPr.append(tblLayout)

    # Explicit table width
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = tblPr.makeelement(qn('w:tblW'), {})
        tblPr.append(tblW)
    tblW.set(qn('w:w'), page_width_twips)
    tblW.set(qn('w:type'), 'dxa')

    # Remove default cell margins so photo fills its column
    tblCellMar = tblPr.makeelement(qn('w:tblCellMar'), {})
    for edge in ('top', 'left', 'bottom', 'right'):
        el = tblCellMar.makeelement(qn(f'w:{edge}'), {
            qn('w:w'): '0', qn('w:type'): 'dxa',
        })
        tblCellMar.append(el)
    old_mar = tblPr.find(qn('w:tblCellMar'))
    if old_mar is not None:
        tblPr.remove(old_mar)
    tblPr.append(tblCellMar)

    # Table borders: none
    borders = tblPr.find(qn('w:tblBorders'))
    if borders is not None:
        tblPr.remove(borders)
    borders = tblPr.makeelement(qn('w:tblBorders'), {})
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = borders.makeelement(qn(f'w:{edge}'), {
            qn('w:val'): 'none', qn('w:sz'): '0',
            qn('w:space'): '0', qn('w:color'): 'auto',
        })
        borders.append(el)
    tblPr.append(borders)

    # Grid columns
    tblGrid = tbl.find(qn('w:tblGrid'))
    if tblGrid is not None:
        tbl.remove(tblGrid)
    tblGrid = tbl.makeelement(qn('w:tblGrid'), {})
    gridCol1 = tblGrid.makeelement(qn('w:gridCol'), {qn('w:w'): contact_col_twips})
    gridCol2 = tblGrid.makeelement(qn('w:gridCol'), {qn('w:w'): photo_col_twips})
    tblGrid.append(gridCol1)
    tblGrid.append(gridCol2)
    tblPr.addnext(tblGrid)

    # Cell widths and borders
    for i, cell in enumerate(header_table.rows[0].cells):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcW = tcPr.find(qn('w:tcW'))
        if tcW is None:
            tcW = tcPr.makeelement(qn('w:tcW'), {})
            tcPr.append(tcW)
        tcW.set(qn('w:w'), contact_col_twips if i == 0 else photo_col_twips)
        tcW.set(qn('w:type'), 'dxa')
        tcBorders = tcPr.makeelement(qn('w:tcBorders'), {})
        for edge_name in ('top', 'left', 'bottom', 'right'):
            el = tcBorders.makeelement(qn(f'w:{edge_name}'), {
                qn('w:val'): 'none', qn('w:sz'): '0',
                qn('w:space'): '0', qn('w:color'): 'auto',
            })
            tcBorders.append(el)
        old = tcPr.find(qn('w:tcBorders'))
        if old is not None:
            tcPr.remove(old)
        tcPr.append(tcBorders)

    # Left cell: contact details
    left_cell = header_table.cell(0, 0)
    left_cell.paragraphs[0].clear()
    p = left_cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = CONTACT_LINE_SPACING
    contact_lines = [('Contact details', ''), ('', '')] + data['contact']
    for i, (label, value) in enumerate(contact_lines):
        if i > 0:
            p.add_run('\n')
        run_label = p.add_run(label + ' ')
        run_label.bold = True
        run_label.font.size = CONTACT_LABEL_SIZE
        run_label.font.color.rgb = GRAY
        run_value = p.add_run(value)
        run_value.font.size = CONTACT_VALUE_SIZE

    # Right cell: photo
    right_cell = header_table.cell(0, 1)
    right_cell.paragraphs[0].clear()
    right_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    right_cell.paragraphs[0].paragraph_format.space_before = Pt(0)
    right_cell.paragraphs[0].paragraph_format.space_after = Pt(0)
    right_cell.paragraphs[0].add_run().add_picture(photo_path, width=Inches(PHOTO_IMG_WIDTH))

    # -- Summary --
    add_heading_styled('Summary', level=1)
    doc.add_paragraph(data['summary'])

    # -- Core Skills --
    add_heading_styled('Core Skills', level=1)
    doc.add_paragraph(data['skills'])

    # -- Languages --
    add_heading_styled('Languages', level=1)
    for lang, desc in data['languages']:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(lang + ': ')
        run.bold = True
        p.add_run(desc)

    doc.add_page_break()

    # -- Experience --
    add_heading_styled('Experience', level=1)
    for exp in data['experience']:
        period = format_period(exp['start'], exp['end'])
        add_experience(exp['company'], exp['title'], period, exp['bullets'])
        if exp.get('page_break_after'):
            doc.add_page_break()

    doc.add_page_break()

    # -- Github Repositories --
    add_heading_styled('Github repositories', level=1)
    for repo in data['repos']:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.add_run(repo)

    # -- Education --
    add_heading_styled('Education', level=1)

    p = doc.add_paragraph()
    run = p.add_run(data['education_heading'])
    run.bold = True

    for edu in data['education']:
        p = doc.add_paragraph()
        run = p.add_run(edu['institution'])
        run.bold = True
        p.add_run('\n' + edu['degree'])
        p.add_run('\n' + edu['period'])

    return doc


# ============================================================
# MARKDOWN GENERATION
# ============================================================

def generate_md(data):
    contact = '\n'.join(f'**{label}** {value}  ' for label, value in data['contact'])
    header = f'# {data["name"]}\n\n**{data["tagline"]}**\n\n**Contact details**\n\n{contact}'

    languages = '\n'.join(f'**{lang}:** {desc}  ' for lang, desc in data['languages'])

    experience_entries = []
    for exp in data['experience']:
        bullets = '\n'.join(f'- {b}' for b in exp['bullets'])
        experience_entries.append(
            f'### {exp["company"]}\n*{exp["title"]}*  \n{format_period(exp["start"], exp["end"])}\n\n{bullets}'
        )
    experience = '## Experience\n\n' + '\n\n'.join(experience_entries)

    repos = '\n'.join(f'- {repo}' for repo in data['repos'])

    edu_entries = '\n\n'.join(
        f'### {e["institution"]}\n*{e["degree"]}*  \n{e["period"]}' for e in data['education']
    )
    education = f'**{data["education_heading"]}**\n\n{edu_entries}'

    sections = [
        header,
        f'## Summary\n\n{data["summary"]}',
        f'## Core Skills\n\n{data["skills"]}',
        f'## Languages\n\n{languages}',
        experience,
        f'## Github repositories\n\n{repos}',
        f'## Education\n\n{education}',
    ]

    return '\n\n---\n\n'.join(sections) + '\n'


# ============================================================
# MAIN
# ============================================================

if __name__ == '__main__':
    script_dir = os.path.dirname(os.path.abspath(__file__))
    with open(os.path.join(script_dir, 'cv_data.json'), encoding='utf-8') as f:
        cv_data = json.load(f)
    photo_path = os.path.join(script_dir, cv_data['photo'])

    # Generate .docx
    docx_name = 'Karsten Opdal - CV.docx'
    doc = generate_docx(cv_data, photo_path)
    doc.save(docx_name)
    print(f'Saved to: {docx_name}')

    # Generate README.md
    md_name = os.path.join(script_dir, 'README.md')
    md_content = generate_md(cv_data)
    with open(md_name, 'w', encoding='utf-8') as f:
        f.write(md_content)
    print(f'Saved to: {md_name}')

    # Convert to PDF via LibreOffice
    output_dir = os.path.dirname(os.path.abspath(docx_name)) or '.'
    subprocess.run([
        'libreoffice', '--headless', '--convert-to', 'pdf',
        '--outdir', output_dir, docx_name
    ], check=True)
    pdf_name = os.path.splitext(docx_name)[0] + '.pdf'
    print(f'Saved to: {pdf_name}')
