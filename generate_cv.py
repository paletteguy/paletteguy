#!/usr/bin/env python3

import os
import io
from PIL import Image
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

doc = Document()

# -- Style defaults --
style = doc.styles['Normal']
font = style.font
font.name = 'Calibri'
font.size = Pt(10.5)

# -- Helper functions --
def add_heading_styled(text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x2E, 0x4A, 0x62)
    return h

def add_experience(company, title, period, description_lines):
    p = doc.add_paragraph()
    run = p.add_run(company)
    run.bold = True
    run.font.size = Pt(12)
    p.space_after = Pt(0)

    p2 = doc.add_paragraph()
    run2 = p2.add_run(title)
    run2.italic = True
    run2.font.size = Pt(10.5)
    p2.space_after = Pt(0)

    p3 = doc.add_paragraph()
    run3 = p3.add_run(period)
    run3.font.size = Pt(9)
    run3.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    p3.space_after = Pt(4)

    for line in description_lines:
        p4 = doc.add_paragraph(line, style='List Bullet')
        p4.paragraph_format.space_after = Pt(2)

    # Small spacer
    doc.add_paragraph().paragraph_format.space_before = Pt(4)


# ============================================================
# HEADER — photo + contact info side by side
# ============================================================
script_dir = os.path.dirname(os.path.abspath(__file__))
photo_path = os.path.join(script_dir, 'me.png')

# Name — use Title heading so it renders large
h = doc.add_heading('Karsten Sperling Opdal', level=0)
h.paragraph_format.space_after = Pt(0)
h.paragraph_format.space_before = Pt(0)
for run in h.runs:
    run.font.color.rgb = RGBColor(0x2E, 0x4A, 0x62)
    run.font.size = Pt(28)
# Remove the bottom border that Title style adds
pPr = h._p.get_or_add_pPr()
pBdr = pPr.find(qn('w:pBdr'))
if pBdr is not None:
    pPr.remove(pBdr)
# Explicitly set no paragraph borders
pBdr = pPr.makeelement(qn('w:pBdr'), {})
for edge in ('top', 'left', 'bottom', 'right'):
    el = pBdr.makeelement(qn(f'w:{edge}'), {
        qn('w:val'): 'none', qn('w:sz'): '0',
        qn('w:space'): '0', qn('w:color'): 'auto',
    })
    pBdr.append(el)
pPr.append(pBdr)

# Tagline — standalone paragraph
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(2)
p.paragraph_format.space_before = Pt(0)
run = p.add_run('Developing Embedded Systems into the future of AI')
run.bold = True
run.italic = True
run.font.size = Pt(11)

# Table for contact info (left) + photo (right)
photo_width = Inches(1.4)
photo_width_twips = str(int(1.4 * 1440))
left_width_twips = str(int((6.5 - 1.4) * 1440))  # remaining page width
page_width_twips = str(int(6.5 * 1440))

header_table = doc.add_table(rows=1, cols=2)
header_table.alignment = WD_TABLE_ALIGNMENT.LEFT

# Configure table: fixed layout, full width, no borders, no style
tbl = header_table._tbl
tblPr = tbl.tblPr if tbl.tblPr is not None else tbl._add_tblPr()

# Remove any table style (prevents LibreOffice from inheriting TableGrid borders)
tblStyle = tblPr.find(qn('w:tblStyle'))
if tblStyle is not None:
    tblPr.remove(tblStyle)

# Set fixed table layout so column widths are respected
tblLayout = tblPr.makeelement(qn('w:tblLayout'), {qn('w:type'): 'fixed'})
tblPr.append(tblLayout)

# Set explicit table width
tblW = tblPr.find(qn('w:tblW'))
if tblW is None:
    tblW = tblPr.makeelement(qn('w:tblW'), {})
    tblPr.append(tblW)
tblW.set(qn('w:w'), page_width_twips)
tblW.set(qn('w:type'), 'dxa')

# Set table-level borders to none
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

# Set grid columns (w:tblGrid) for fixed layout
tblGrid = tbl.find(qn('w:tblGrid'))
if tblGrid is not None:
    tbl.remove(tblGrid)
tblGrid = tbl.makeelement(qn('w:tblGrid'), {})
gridCol1 = tblGrid.makeelement(qn('w:gridCol'), {qn('w:w'): left_width_twips})
gridCol2 = tblGrid.makeelement(qn('w:gridCol'), {qn('w:w'): photo_width_twips})
tblGrid.append(gridCol1)
tblGrid.append(gridCol2)
# Insert tblGrid after tblPr
tblPr.addnext(tblGrid)

# Set cell-level widths and borders to none
for i, cell in enumerate(header_table.rows[0].cells):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    # Set cell width
    tcW = tcPr.find(qn('w:tcW'))
    if tcW is None:
        tcW = tcPr.makeelement(qn('w:tcW'), {})
        tcPr.append(tcW)
    tcW.set(qn('w:w'), left_width_twips if i == 0 else photo_width_twips)
    tcW.set(qn('w:type'), 'dxa')
    # Set cell borders to none
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
p.paragraph_format.line_spacing = Pt(11)
contact_lines = [
    ('Contact details', ''),
    ('',''),
    ('Company:', 'Opdal Enterprise'),
    ('Address:', 'Nordvangen 13, DK-4600 Køge'),
    ('Phone:', '+45 42 40 42 82'),
    ('Email:', 'karsten.s.opdal@gmail.com'),
    ('LinkedIn:', 'https://linkedin.com/in/karstenopdal-7a11512'),
    ('Web:', 'https://opdal.dk'),
    ('GitHub profile:', 'https://github.com/paletteguy'),
]
for i, (label, value) in enumerate(contact_lines):
    if i > 0:
        p.add_run('\n')
    run_label = p.add_run(label + ' ')
    run_label.bold = True
    run_label.font.size = Pt(9)
    run_label.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    run_value = p.add_run(value)
    run_value.font.size = Pt(10)

# Right cell: photo (resize to thumbnail first)
right_cell = header_table.cell(0, 1)
right_cell.paragraphs[0].clear()
right_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
right_cell.paragraphs[0].paragraph_format.space_before = Pt(0)
right_cell.paragraphs[0].paragraph_format.space_after = Pt(0)
# Insert photo
img = Image.open(photo_path)
img.thumbnail((350, 350))
img_stream = io.BytesIO()
img.save(img_stream, format='PNG')
img_stream.seek(0)
right_cell.paragraphs[0].add_run().add_picture(img_stream, width=photo_width)

# ============================================================
# SUMMARY
# ============================================================
add_heading_styled('Summary', level=1)

summary = (
    "I make embedded Linux do what it's told — on hardware where failure isn't an option.\n\n"
    "For 15 years at Laerdal Medical, I was the architect behind SimPad's embedded platform — owning "
    "everything from the Yocto build system delivering production Linux across ARM and x86, to the real-time "
    "simulation engine that brings medical manikins to life via CAN bus, IOC controllers, and physiological "
    "modelling. I evolved the build system through six major Yocto releases (Rocko to Scarthgap), migrated "
    "the platform from Ubuntu to a custom distribution with RAUC A/B OTA updates and HSM-signed security, "
    "and architected CI/CD pipelines with SBOM generation and CVE tracking for medical device compliance. "
    "Along the way I built Qt/QML touch-screen UIs, ALSA audio pipelines, WiFi/Bluetooth integration "
    "(with patches contributed to BlueZ and Qt Bluetooth), and a suite of internal tools — from an "
    "AI-powered diagnostic platform to an SSH certificate portal and a cross-platform device imager.\n\n"
    "Before Laerdal, I spent a decade in the Nokia ecosystem — developing the user interfaces for the "
    "iconic 3210 and 3310, implementing concatenated SMS and T9 input, and building simulation tools that "
    "let engineers develop phone software in Visual Studio without touching hardware. I debugged firmware "
    "issues nobody else could crack, including a battery drain bug that improved battery lifetime by 30%. "
    "At Infineon, I shipped the Nokia 2110 low-cost platform where every kilobyte counted.\n\n"
    "I own the full stack from silicon to screen: board bring-up on i.MX6, i.MX8, and x86, Linux kernel "
    "debugging and Device Tree customization, C++ services in real-time, and secure OTA deployment. "
    "I've integrated AI deeply into my engineering workflow — from AI-assisted code generation and review "
    "to building custom tooling around LLMs for debugging and documentation.\n\n"
    "30+ years from Nokia handsets to medical devices. Now looking for the next platform worth building.\n\n"
    "Copenhagen, Denmark | Open to remote across Europe"
)
doc.add_paragraph(summary)

# ============================================================
# CORE SKILLS
# ============================================================
add_heading_styled('Core Skills', level=1)

doc.add_paragraph(
    'C++ (C++23) · Embedded Linux · Yocto · Linux kernel · Qt/QML (Qt 3–6) · Python · '
    'WiFi/Bluetooth · CI/CD · Boost · AI-assisted engineering'
)

# ============================================================
# EXPERIENCE
# ============================================================
add_heading_styled('Experience', level=1)

add_experience(
    'Laerdal Medical', 'Senior Software Engineer Consulting',
    'October 2010 – February 2026 (15 years 5 months)',
    [
        'Senior Embedded Software Engineer / Build System Architect for SimPad devices / simulators and SimMan3G product lines.',
        'Lead maintainer of Yocto/OpenEmbedded build platform delivering production Linux across ARM32, ARM64, x86-32, and x86-64. Evolved through six major Yocto releases (Rocko to Scarthgap). Migrated from Ubuntu to custom distribution with RAUC A/B OTA updates and Azure Key Vault HSM signing.',
        'Architected multi-platform CI/CD pipelines with GitHub Actions, SBOM generation, and automated CVE tracking for medical device compliance.',
        'Linux kernel configuration, debugging, patching, and Device Tree customization across i.MX6, i.MX8M Plus, and Intel x86 — board bring-up, peripheral enablement, pin muxing, clock trees, and driver troubleshooting.',
        'Developed real-time simulation engine controlling medical manikins (SimMom, SimBaby, MammaAnne, SimMan ALS) via CAN bus or IOC controller and physiological modelling. C++ (up to C++23) with Boost, Qt 3\u20136 with deep QML expertise for embedded UIs. Proficient in C#.',
        'ALSA audio pipelines for clinical simulation. Bluetooth expert with patches to BlueZ and Qt Bluetooth stacks.',
        'Created SimServer Imager (Qt6/C++/QML) for managed WIC/VSI/SPU deployment and CDN distribution. Built VEX Kernel Checker \u2014 AI-assisted CVE analysis integrated with Dependency-Track. Built cross-platform SimServer companion app.',
        'Built SSH Certificate Portal (FastAPI/Python) \u2014 self-service time-limited SSH certificates with Azure AD OIDC, HSM-protected CA signing (RSA/ED25519/ECDSA), deployed on Azure App Service.',
        'Designed Azure Key Vault HSM signing workflow for RAUC bundles and secure boot \u2014 non-exportable keys, automated renewal via Azure Automation, multi-environment CI integration.',
        'Built SimServer Companion \u2014 AI-powered diagnostic platform using Claude/Gemini/OpenAI for crash analysis across SimPad, LinkBox, and CAN firmware. Jira automation, source context integration. Delivered as CLI, VS Code extension, and Tauri desktop app.',
    ]
)

doc.add_page_break()

add_experience(
    'Oscilloscope', 'Senior Software Engineer Consulting',
    'October 2010 – June 2019 (8 years 9 months)',
    [
        'Freelance',
    ]
)

add_experience(
    'Nokia', 'Software Engineering Specialist',
    'April 2010 – October 2010 (7 months)',
    [
        'I excelled in identifying and resolving complex firmware issues at Nokia, significantly enhancing product performance.',
        'Specialized in debugging intricate hardware/software interaction bugs on mobile platforms.',
        'Successfully fixed a battery drain issue that had persisted for years, improving battery lifetime by 30%.',
        'Developed critical skills in embedded systems analysis and problem-solving within a leading technology company.',
    ]
)

add_experience(
    'Nokia Mobile Phones', 'Senior Software Engineering Consultant',
    'September 2008 – February 2010 (1 year 6 months)',
    [
        'Provided expert consultation on the S40 platform software, focusing on bug fixing and error correction.',
        'Debugged firmware issues across the S40 software stack to enhance performance and reliability.',
        'Collaborated with cross-functional teams to ensure timely resolution of software defects, improving user experience.',
    ]
)

doc.add_page_break()

add_experience(
    'Infineon', 'Software Expert',
    'October 2007 – September 2008 (1 year)',
    [
        'Developed the Nokia 2110 low-cost phone platform during an expat assignment in Copenhagen.',
        'Optimized embedded software for resource-constrained hardware, ensuring efficient use of memory.',
        'Delivered high-quality software under tight deadlines, focusing on minimizing resource usage.',
    ]
)

add_experience(
    'Thorsø Data', 'Software Expert',
    'October 2007 – September 2008 (1 year)',
    [
        'In house Senior Contract Software Engineer',
    ]
)

add_experience(
    'Tang-Data A/S', 'Senior Software Engineer',
    'March 2007 – October 2007 (8 months)',
    [
        'Developed a comprehensive veterinary CRM system utilizing Qt 3 and Qt 4, enhancing client management.',
        'Provided on-site support for veterinary computer setups, ensuring seamless CRM hardware installations.',
        'Collaborated with cross-functional teams to address customer needs and improve system functionality.',
    ]
)

doc.add_page_break()

add_experience(
    'Nokia', 'Senior Software Engineer',
    'October 2000 – February 2007 (6 years 5 months)',
    [
        'Enhanced the software development lifecycle at Nokia through innovative UI development.',
        'Developed internal tools that streamlined handset software testing processes.',
        'Enabled mobile phone development in a simulated environment using Microsoft Visual Studio.',
    ]
)

add_experience(
    'EC-Soft Danmark A/S', 'Senior System Software Engineer',
    'October 2000 – December 2001 (1 year 3 months)',
    [
        'Consultant work for Nokia Denmark A/S developing S30 phones',
    ]
)

add_experience(
    'Telenor', 'Contract Software Engineer',
    'April 2000 – June 2000 (3 months)',
    [
        'Developed a web service for read outlook mails and calendar on phones',
    ]
)

add_experience(
    'EC-Soft Norge AS', 'Senior Software Engineer',
    'May 2000 – October 2000 (6 months)',
    [
        'In house Software Consultant',
    ]
)

doc.add_page_break()

add_experience(
    'Nokia', 'Contract Software Engineer',
    'May 1999 – March 2000 (11 months)',
    [
        'Developed user interfaces for Nokia\'s iconic handset models, including the 3210 and 3310.',
        'Collaborated with a fellow developer to implement concatenated SMS messaging, enhancing user communication.',
        'Integrated T9 input support, improving text input efficiency for users.',
    ]
)

add_experience(
    'EC-Soft Danmark A/S', 'Software Consultant',
    'May 1999 – March 2000 (11 months)',
    [
        'In house software consultant',
    ]
)

add_experience(
    'Vizion Factory ApS', 'Software Engineer',
    'October 1995 – April 1999 (3 years 7 months)',
    [
        'Played a key role in software development and early web solutions, contributing to the burgeoning internet landscape.',
        'Built applications that addressed emerging needs and developed a training engine paired with a web interface to streamline user interaction.',
    ]
)

add_experience(
    'Greve Kommune', 'Network System Specialist',
    'January 1991 – October 1995 (4 years 10 months)',
    [
        'Set up new networks to enhance connectivity and efficiency within the organization.',
        'Provided training to employees on computer use, fostering a tech-savvy workplace.',
        'Maintained hardware by adding and replacing components, ensuring optimal performance.',
    ]
)
doc.add_page_break()

add_experience(
    'Opdal Enterprise', 'Owner', 'September 2008 – Present (17 years 7 months)',
    ['Køge Municipality']
)

add_experience(
    'Home development', 'Developer', 'April 1983 – Present (43 years)',
    [
        'Developed Game for C-64, demos for C-64 and Amiga and a task / thread / resource manager for Amiga OS.',
        'Developed music editor software C-64, Amiga and PC Dos.',
        'Currently developing Android application for learning purpose and application with Rust and Svelte',
    ]
)

doc.add_page_break()

# ============================================================
# GITHUB REPOSITORIES
# ============================================================
add_heading_styled('Github repositories', level=1)

github_repos = [
    'https://github.com/paletteguy/profile',
    'https://github.com/Laerdal/vex-kernel-checker',
    'https://github.com/Laerdal/linux-fslc',
    'https://github.com/Laerdal/meta-dependencytrack',
    'https://github.com/Laerdal-Medical/simserver-imager',
]
for repo in github_repos:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(repo)

# ============================================================
# EDUCATION
# ============================================================
add_heading_styled('Education', level=1)

p = doc.add_paragraph()
run = p.add_run('Software development')
run.bold = True

p = doc.add_paragraph()
run = p.add_run('Niels Brock')
run.bold = True
p.add_run("\nBachelor's Degree, Computer Science")
p.add_run('\nSeptember 1990 – June 1992')

p = doc.add_paragraph()
run = p.add_run('Niels Brock')
run.bold = True
p.add_run("\nBachelor's Degree, Business/Commerce, General")
p.add_run('\nAugust 1988 – June 1990')

p = doc.add_paragraph()
run = p.add_run('EFG Handel og Kontor')
run.bold = True
p.add_run('\nBasic business school')
p.add_run('\nAugust 1987 – June 1988')

p = doc.add_paragraph()
run = p.add_run('Krogaardskolen')
run.bold = True
p.add_run('\n1976 – 1987')

# ============================================================
# SAVE
# ============================================================
output = 'Karsten Opdal - CV.docx'
doc.save(output)
print(f'Saved to: {output}')
