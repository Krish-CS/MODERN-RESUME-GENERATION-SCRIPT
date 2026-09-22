"""
================================================================================
MODERN ATS RESUME GENERATION SCRIPT (DOCX & PDF)
================================================================================
Author: Krish-CS
Repository: https://github.com/Krish-CS/MODERN-RESUME-GENERATION-SCRIPT
Format: Strict 1-Page Layout (US Letter 8.5" x 11", 0.25" Margins)
Typography: Times New Roman (ATS Standard)

================================================================================
INSTRUCTIONS FOR USERS & AI CODING AGENTS:
================================================================================
This script is designed to be fully automated and easy to customize for any user
or AI agent.

WHERE TO EDIT:
- All resume content is centralized in the `RESUME_DATA` dictionary below.
- Do NOT alter the layout math (margins, tab stops, font sizes, line spacing)
  unless you are adding or removing significant amounts of text, as the current
  configuration guarantees a strict, beautiful 1-page fit with ~70-90pt of
  bottom margin clearance.

WHAT TO REPLACE IN `RESUME_DATA`:
1. "header":
   - "name": Replace with your Full Name (all caps or title case).
   - "location": City, State / Country.
   - "phone": Contact number with country code.
   - "email": Professional email address (will generate a clickable mailto: link).
   - "linkedin_url" & "linkedin_text": Your full LinkedIn URL and display handle.
   - "github_url" & "github_text": Your full GitHub profile URL and display handle.

2. "summary_runs":
   - List of tuples: (text_string, is_bold_boolean, is_italic_boolean).
   - Replace "[Target Role / Professional Title]" with your target job role.
   - Replace "[Skill 1], [Skill 2]..." with your 4-5 core technical competencies.
   - Keep the summary between 3 to 4 lines to preserve the 1-page layout.

3. "education":
   - List of educational qualifications (College, Higher Secondary, Secondary).
   - "institution": Name of the university, college, or school.
   - "score": CGPA (e.g. "[CGPA]" or "CGPA: 8.5 / 10.0") or percentage.
   - "location": City, Country.
   - "degree": Degree / Program / Major name (e.g., B.Tech / B.E. in CSE).
   - "year": Graduation or attendance timeline (e.g., "20XX – 20YY" or "20XX").

4. "technical_skills":
   - List of tuples: (Category Name, Comma-separated skills).
   - Group your skills logically (Programming Languages, Frontend, Backend, etc.).

5. "soft_skills":
   - List of tuples: (Skill Category, Description sentence).
   - Keep descriptions concise (1-2 lines per skill) focused on engineering impact.

6. "experience":
   - List of work experience or internship entries.
   - "role": Job title and company name.
   - "year": Timeframe (e.g., "20XX – 20YY").
   - "department": Team or division name.
   - "location": City, Country or "Remote".
   - "bullets": 2 action-verb oriented achievements with quantifiable metrics.

7. "projects":
   - List of 2 to 3 featured technical projects.
   - "title": Project name and domain.
   - "tech_stack": Comma-separated list of frameworks, languages, and tools.
   - "year": Year of completion (e.g., "20XX – 20YY" or "20XX").
   - "github_url" & "github_text": Optional project repository link.
   - "bullets": 2 concise bullet points per project detailing architecture,
     implementation, and results.

8. "certifications":
   - List of tuples: (Issuing Organization, Certification title & year).

9. "languages":
   - List of tuples: (Language Name, Proficiency level).
================================================================================
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# ==============================================================================
# 1. RESUME DATA CONFIGURATION
# (USERS & AI AGENTS: EDIT CONTENT INSIDE THIS DICTIONARY)
# ==============================================================================
RESUME_DATA = {
    # --------------------------------------------------------------------------
    # Header Information
    # --------------------------------------------------------------------------
    "header": {
        "name": "[YOUR FULL NAME]",
        "location": "City, State / Country",
        "phone": "+1 (555) 012-3456",
        "email": "your.email@example.com",
        "linkedin_url": "https://linkedin.com/in/yourusername",
        "linkedin_text": "linkedin.com/in/yourusername",
        "github_url": "https://github.com/yourusername",
        "github_text": "github.com/yourusername",
    },

    # --------------------------------------------------------------------------
    # Professional Summary (Structured as formatted runs for bold keywords)
    # Format: (text, is_bold, is_italic)
    # --------------------------------------------------------------------------
    "summary_runs": [
        ("[Target Role / Professional Title]", True, False),
        (" with a strong foundation in ", False, False),
        ("[Skill 1], [Skill 2], [Skill 3], [Skill 4], and [Skill 5]", True, False),
        (". Academic and project experience building modular software applications, database-driven architectures, and responsive user interfaces. Skilled in frontend development, RESTful API integration, database management, and collaborative Git workflows. Committed to writing clean, maintainable code, solving complex technical challenges, and continuous learning in modern technologies.", False, False),
    ],

    # --------------------------------------------------------------------------
    # Education
    # Format: Institution, Score ([CGPA] / [Percentage]), Location, Degree, Year
    # --------------------------------------------------------------------------
    "education": [
        {
            "institution": "[University / College Name]",
            "score": "[CGPA]",
            "location": "[City, Country]",
            "degree": "[Bachelor of Engineering / Technology – Computer Science]",
            "year": "20XX – 20YY",
        },
        {
            "institution": "[Higher Secondary School Name]",
            "score": "[Percentage]",
            "location": "[City, Country]",
            "degree": "[Higher Secondary Certificate (HSC / 12th Grade)]",
            "year": "20XX – 20YY",
        },
        {
            "institution": "[Secondary School Name]",
            "score": "[Percentage]",
            "location": "[City, Country]",
            "degree": "[Secondary School Leaving Certificate (SSLC / 10th Grade)]",
            "year": "20XX",
        },
    ],

    # --------------------------------------------------------------------------
    # Technical Skills
    # Format: (Category Name, Items)
    # --------------------------------------------------------------------------
    "technical_skills": [
        ("Programming Languages", "[Language 1], [Language 2], [Language 3], [Language 4]"),
        ("Web & Frontend Technologies", "[HTML5], [CSS3], [JavaScript], [Frontend Framework (e.g. React.js)], [Responsive UI/UX]"),
        ("Backend & API Development", "[Backend Framework (e.g. Node.js / Spring Boot)], [RESTful APIs], [Microservices], [CRUD Operations]"),
        ("Database Management", "[Relational DB (e.g. MySQL / PostgreSQL)], [NoSQL DB (e.g. MongoDB)], [SQL Queries & Schema Design]"),
        ("Developer Tools & Platforms", "[Git], [GitHub], [VS Code], [Postman], [Docker / Cloud Platforms (e.g. AWS / Azure)]"),
        ("Core CS Fundamentals", "[Data Structures & Algorithms], [OOPs], [DBMS], [Operating Systems], [Computer Networks]"),
    ],

    # --------------------------------------------------------------------------
    # Soft Skills
    # Format: (Skill Name, Description)
    # --------------------------------------------------------------------------
    "soft_skills": [
        ("Problem Solving & Analytical Thinking: ", "Deconstructs complex engineering challenges into modular solutions, applying systematic logic and iterative debugging to produce robust solutions."),
        ("Team Leadership & Collaboration: ", "Coordinates module delivery, conducts peer code reviews, and communicates effectively across cross-functional teams in Agile/Scrum environments."),
        ("Time Management & Prioritization: ", "Balances project deliverables, technical research, and academic commitments effectively without compromising output quality."),
        ("Adaptability & Continuous Learning: ", "Quickly masters emerging technologies, modern development tools, and industry best practices to adapt to dynamic project requirements."),
    ],

    # --------------------------------------------------------------------------
    # Professional Experience / Internship
    # Format: Role, Year, Department, Location, Bullet Points
    # --------------------------------------------------------------------------
    "experience": [
        {
            "role": "[Job Title / Intern Role – Company / Organization Name]",
            "year": "20XX – 20YY",
            "department": "[Department / Division / Team Name]",
            "location": "[City, Country / Remote]",
            "bullets": [
                "Engineered and optimized [application/module/feature] using [Technologies/Tools], enhancing [performance/scalability/efficiency] by [X%].",
                "Collaborated with cross-functional development teams on sprint planning, code reviews, and CI/CD pipelines to deliver high-quality releases on schedule.",
            ],
        },
    ],

    # --------------------------------------------------------------------------
    # Projects (Supports 3 full projects with clickable repository links)
    # Format: Title, Tech Stack, Year, GitHub URL, GitHub Display Text, Bullets
    # --------------------------------------------------------------------------
    "projects": [
        {
            "title": "[Project Title 1 – Full-Stack Web Application]",
            "tech_stack": "[Tech 1, Tech 2, Tech 3, Tech 4]",
            "year": "20XX – 20YY",
            "github_url": "https://github.com/yourusername/project-one",
            "github_text": "github.com/yourusername/project-one",
            "bullets": [
                "Designed and implemented a full-stack [application type] providing [core feature 1], [core feature 2], and [core feature 3] for [target user/use case].",
                "Implemented secure CRUD operations, parameterized SQL/NoSQL queries, and RESTful APIs with comprehensive authentication and validation mechanisms.",
            ],
        },
        {
            "title": "[Project Title 2 – Automated Monitoring Tool / Platform]",
            "tech_stack": "[Tech 1, Tech 2, Tech 3, Tech 4]",
            "year": "20XX",
            "github_url": "https://github.com/yourusername/project-two",
            "github_text": "github.com/yourusername/project-two",
            "bullets": [
                "Developed an automated [system/monitoring service] to track [key metrics/events] in real time, significantly reducing incident response time.",
                "Integrated third-party APIs and notification gateways to trigger automated alerts and streamline background telemetry logging.",
            ],
        },
        {
            "title": "[Project Title 3 – Data Pipeline / Distributed Service]",
            "tech_stack": "[Tech 1, Tech 2, Tech 3, Tech 4]",
            "year": "20XX",
            "github_url": "https://github.com/yourusername/project-three",
            "github_text": "github.com/yourusername/project-three",
            "bullets": [
                "Architected and deployed a scalable [system/model/service], implementing [core algorithms/features] to streamline [data processing/operational workflow].",
                "Configured automated CI/CD pipelines and unit testing suites, achieving [X%] test coverage and ensuring regression-free deployments.",
            ],
        },
    ],

    # --------------------------------------------------------------------------
    # Certifications
    # Format: (Issuing Authority, Credential Title & Year)
    # --------------------------------------------------------------------------
    "certifications": [
        ("[Cloud Provider / Authority]: ", "[Certification Title, e.g., Cloud / Solutions Architecture] (20XX)."),
        ("[Technical Institute / Academy]: ", "[Certification Title, e.g., Full Stack Development / Java Certification] (20XX)."),
        ("[Online Learning Platform]: ", "[Specialization Title, e.g., Database Design & Algorithms] (20XX)."),
        ("[Industry Partner / Organization]: ", "[Certification Title, e.g., Artificial Intelligence & Automation] (20XX)."),
    ],

    # --------------------------------------------------------------------------
    # Languages
    # Format: (Language Name, Proficiency Level)
    # --------------------------------------------------------------------------
    "languages": [
        ("[Language 1]: ", "Professional Working Proficiency / Native."),
        ("[Language 2]: ", "Professional Working Proficiency / Fluent."),
    ],
}


# ==============================================================================
# 2. LOW-LEVEL WORD XML & FORMATTING HELPER FUNCTIONS
# ==============================================================================

def set_run_font(run, font_name="Times New Roman", size_pt=9.25, bold=False, italic=False, color_rgb=(0, 0, 0)):
    """
    Applies strict typography styling to an individual Word text run.
    
    Why this function is necessary:
    In python-docx, simply setting `run.font.name` is often insufficient for Word
    to embed the font properly when converting to PDF or opening on different OS.
    This function explicitly injects the `<w:rFonts>` XML element for ascii, hAnsi,
    and cs (complex scripts), ensuring 100% font consistency across platforms.

    Parameters:
    - run: The docx.text.run.Run object to style.
    - font_name: The font family name (default: "Times New Roman").
    - size_pt: Font size in points (default: 9.25 pt).
    - bold: Boolean flag for bold weight.
    - italic: Boolean flag for italic styling.
    - color_rgb: Tuple of (R, G, B) integers (default: (0, 0, 0) for pure black).
    """
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(*color_rgb)
    
    # Inject w:rFonts XML element to guarantee font binding
    rPr = run._r.get_or_add_rPr()
    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" w:cs="{font_name}"/>')
    rPr.append(rFonts)


def add_hyperlink(paragraph, url, text, font_name="Times New Roman", font_size_pt=8.8, color="000000", underline=True):
    """
    Inserts a clickable, ATS-compliant hyperlink into a paragraph using Word OPC.

    Why this function is necessary:
    Standard python-docx does not have a high-level API for clickable hyperlinks.
    This function creates a relationship ID in the document part, constructs
    the `<w:hyperlink>` and `<w:r>` XML nodes, and binds them to the URL.

    Parameters:
    - paragraph: The docx.text.paragraph.Paragraph object to append the link to.
    - url: The target URL (e.g. 'https://github.com/user' or 'mailto:user@mail.com').
    - text: The visible anchor text displayed in the document.
    - font_name: Font family (default: 'Times New Roman').
    - font_size_pt: Font size in points (default: 8.8 pt).
    - color: Hex color string without '#' (default: '000000' for clean black).
    - underline: Boolean flag to enable standard hyperlink underline (default: True).
    """
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w", "r")} r:id="{r_id}" />')
    new_run = parse_xml(f'<w:r {nsdecls("w")} />')
    rPr = parse_xml(f'<w:rPr {nsdecls("w")} />')
    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" w:cs="{font_name}"/>')
    rPr.append(rFonts)
    if color:
        c = parse_xml(f'<w:color {nsdecls("w")} w:val="{color}"/>')
        rPr.append(c)
    if underline:
        u = parse_xml(f'<w:u {nsdecls("w")} w:val="single"/>')
        rPr.append(u)
    val_sz = str(int(font_size_pt * 2))
    sz = parse_xml(f'<w:sz {nsdecls("w")} w:val="{val_sz}"/>')
    rPr.append(sz)
    new_run.append(rPr)
    new_run_text = parse_xml(f'<w:t {nsdecls("w")}>{text}</w:t>')
    new_run.append(new_run_text)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_icon(paragraph, icon_name, size_pt=8.0, offset_val="-2"):
    """
    Inserts a small contact icon with vertical baseline alignment.

    Why this function is necessary:
    Icons in contact lines often appear misaligned (floating above text).
    This function applies a `<w:position>` vertical offset in half-points
    (e.g., -2 half points) to shift the icon down so it aligns seamlessly
    with the middle of the adjacent text.
    If the icon file is missing, it skips gracefully without throwing an error.

    Parameters:
    - paragraph: Target paragraph.
    - icon_name: Base name of the PNG file in icons/ folder (e.g. 'phone', 'mail').
    - size_pt: Width and height of the icon in points (default: 8.0 pt).
    - offset_val: Vertical position offset in half-points (default: '-2').
    """
    icon_path = os.path.join("icons", f"{icon_name}.png")
    if os.path.exists(icon_path):
        run = paragraph.add_run()
        if offset_val and offset_val != "0":
            rPr = run._r.get_or_add_rPr()
            rPr.append(parse_xml(f'<w:position {nsdecls("w")} w:val="{offset_val}"/>'))
        run.add_picture(icon_path, width=Pt(size_pt), height=Pt(size_pt))


# ==============================================================================
# 3. CORE RESUME BUILDER ENGINE
# ==============================================================================

def create_resume(data, output_path):
    """
    Assembles the complete 1-page resume DOCX from a data dictionary.

    Architecture & Layout Philosophy:
    - Standard US Letter dimensions (8.5" x 11.0").
    - Strict margins: Top: 0.25", Bottom: 0.25", Left: 0.40", Right: 0.40".
    - Total content width = 8.5" - (2 * 0.40") = 7.70".
    - Section headings use a 0.75pt bottom border line `<w:pBdr>` for clean separation.
    - Uses exact tab stops for 2-column and 3-column rows instead of tables.
      (Tables often break ATS parsing or add unwanted internal cell padding).
    - Font sizes: Header Name: 18.5 pt, Section Headings: 10.5 pt, Body Text: 9.25 pt,
      Contact Line: 8.8 pt.

    Parameters:
    - data: A dictionary matching the structure of `RESUME_DATA`.
    - output_path: Filepath where the generated .docx file should be saved.
    """
    doc = docx.Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    section.top_margin = Inches(0.25)
    section.bottom_margin = Inches(0.25)
    section.left_margin = Inches(0.40)
    section.right_margin = Inches(0.40)
    content_width = Inches(7.70)

    # Configure default Normal style
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(9.25)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.04
    normal_style.paragraph_format.space_before = Pt(0)
    normal_style.paragraph_format.space_after = Pt(0)

    FS = 9.25  # Standard base font size throughout the document

    # --------------------------------------------------------------------------
    # Internal Helper Closures for Structured Components
    # --------------------------------------------------------------------------

    def add_section_heading(title, space_before=3.6, space_after=1.2):
        """Adds an ATS-friendly section heading with an underlying border line."""
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        pPr = p._p.get_or_add_pPr()
        pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/></w:pBdr>')
        pPr.append(pBdr)
        run = p.add_run(title)
        set_run_font(run, size_pt=10.5, bold=True)
        return p

    def add_two_column_row(left_text, right_text, left_bold=True, left_italic=False,
                           right_bold=False, right_italic=False, space_before=1.1, space_after=0):
        """Creates a row with left-aligned and right-aligned text using a right tab stop."""
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.tab_stops.add_tab_stop(content_width, WD_TAB_ALIGNMENT.RIGHT)
        r_left = p.add_run(left_text)
        set_run_font(r_left, size_pt=FS, bold=left_bold, italic=left_italic)
        r_right = p.add_run('\t' + right_text)
        set_run_font(r_right, size_pt=FS, bold=right_bold, italic=right_italic)
        return p

    def add_three_column_row(left_text, mid_text, right_text, left_bold=True, left_italic=False,
                             mid_bold=False, mid_italic=False, right_bold=False, right_italic=False,
                             space_before=1.1, space_after=0):
        """Creates a 3-column row (Institution \t Score \t Location) with precision tab stops."""
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        # Tab 1: Mid text at 5.8 inches (left-aligned)
        p.paragraph_format.tab_stops.add_tab_stop(Inches(5.8), WD_TAB_ALIGNMENT.LEFT)
        # Tab 2: Right text at content_width 7.70 inches (right-aligned)
        p.paragraph_format.tab_stops.add_tab_stop(content_width, WD_TAB_ALIGNMENT.RIGHT)
        r_left = p.add_run(left_text)
        set_run_font(r_left, size_pt=FS, bold=left_bold, italic=left_italic)
        r_mid = p.add_run('\t' + mid_text)
        set_run_font(r_mid, size_pt=FS, bold=mid_bold, italic=mid_italic)
        r_right = p.add_run('\t' + right_text)
        set_run_font(r_right, size_pt=FS, bold=right_bold, italic=right_italic)
        return p

    def add_project_header(title, tech_stack, date_text, github_url=None, github_text=None, space_before=1.2, space_after=0):
        """Creates a project header line with title, optional clickable repo link, tech stack, and year."""
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.tab_stops.add_tab_stop(content_width, WD_TAB_ALIGNMENT.RIGHT)
        r1 = p.add_run(title)
        set_run_font(r1, size_pt=FS, bold=True)
        if github_url and github_text:
            r_sp = p.add_run(' (')
            set_run_font(r_sp, size_pt=8.2)
            add_hyperlink(p, github_url, github_text, font_size_pt=8.2)
            r_cl = p.add_run(')')
            set_run_font(r_cl, size_pt=8.2)
        r2 = p.add_run(' | ')
        set_run_font(r2, size_pt=FS)
        r3 = p.add_run(tech_stack)
        set_run_font(r3, size_pt=FS, italic=True)
        r4 = p.add_run('\t' + date_text)
        set_run_font(r4, size_pt=FS)
        return p

    def add_bullet_point(runs_data, space_before=0.8, space_after=0.8):
        """Creates an ATS-compliant bullet point with a custom hanging indent."""
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.left_indent = Inches(0.18)
        p.paragraph_format.first_line_indent = Inches(-0.12)
        p.paragraph_format.line_spacing = 1.06
        bullet_run = p.add_run('• ')
        set_run_font(bullet_run, size_pt=FS)
        for text, bold, italic in runs_data:
            r = p.add_run(text)
            set_run_font(r, size_pt=FS, bold=bold, italic=italic)
        return p

    def add_skill_row(category, items, space_before=1.0, space_after=1.0):
        """Creates a technical skill row with a bold category label followed by skill items."""
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.06
        r1 = p.add_run(f"{category}: ")
        set_run_font(r1, size_pt=FS, bold=True)
        r2 = p.add_run(items)
        set_run_font(r2, size_pt=FS)
        return p

    # --------------------------------------------------------------------------
    # 1. HEADER SECTION (Name + Contact Row with Icons & Hyperlinks)
    # --------------------------------------------------------------------------
    hdr = data.get("header", {})
    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(1.0)
    r_name = p_name.add_run(hdr.get("name", "[YOUR FULL NAME]"))
    set_run_font(r_name, size_pt=18.5, bold=True)

    # Contact line
    p_contact = doc.add_paragraph()
    p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_contact.paragraph_format.space_before = Pt(0)
    p_contact.paragraph_format.space_after = Pt(2.0)

    # Location
    add_icon(p_contact, 'pin', size_pt=8.0, offset_val="0")
    r_loc = p_contact.add_run(f" {hdr.get('location', 'City, Country')}")
    set_run_font(r_loc, size_pt=8.8)
    r_s1 = p_contact.add_run(' | ')
    set_run_font(r_s1, size_pt=8.8)

    # Phone
    add_icon(p_contact, 'phone', size_pt=8.0, offset_val="-1")
    r_ph = p_contact.add_run(f" {hdr.get('phone', '+1 (555) 012-3456')}")
    set_run_font(r_ph, size_pt=8.8)
    r_s2 = p_contact.add_run(' | ')
    set_run_font(r_s2, size_pt=8.8)

    # Email
    add_icon(p_contact, 'mail', size_pt=8.0, offset_val="-3")
    r_sp1 = p_contact.add_run(' ')
    set_run_font(r_sp1, size_pt=8.8)
    email = hdr.get("email", "your.email@example.com")
    add_hyperlink(p_contact, f"mailto:{email}", email, font_size_pt=8.8)
    r_s3 = p_contact.add_run(' | ')
    set_run_font(r_s3, size_pt=8.8)

    # LinkedIn
    add_icon(p_contact, 'linkedin', size_pt=8.0, offset_val="-2")
    r_sp2 = p_contact.add_run(' ')
    set_run_font(r_sp2, size_pt=8.8)
    add_hyperlink(p_contact, hdr.get("linkedin_url", "https://linkedin.com"), hdr.get("linkedin_text", "linkedin.com/in/username"), font_size_pt=8.8)
    r_s4 = p_contact.add_run(' | ')
    set_run_font(r_s4, size_pt=8.8)

    # GitHub
    add_icon(p_contact, 'github', size_pt=8.0, offset_val="-2")
    r_sp3 = p_contact.add_run(' ')
    set_run_font(r_sp3, size_pt=8.8)
    add_hyperlink(p_contact, hdr.get("github_url", "https://github.com"), hdr.get("github_text", "github.com/username"), font_size_pt=8.8)

    # --------------------------------------------------------------------------
    # 2. PROFESSIONAL SUMMARY
    # --------------------------------------------------------------------------
    add_section_heading('PROFESSIONAL SUMMARY', space_before=2.8, space_after=1.0)
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(0.8)
    p_sum.paragraph_format.space_after = Pt(1.0)
    p_sum.paragraph_format.line_spacing = 1.05
    p_sum.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    for text, bold, italic in data.get("summary_runs", []):
        r = p_sum.add_run(text)
        set_run_font(r, size_pt=FS, bold=bold, italic=italic)

    # --------------------------------------------------------------------------
    # 3. EDUCATION
    # --------------------------------------------------------------------------
    add_section_heading('EDUCATION', space_before=3.2, space_after=1.0)
    for edu in data.get("education", []):
        add_three_column_row(
            edu.get("institution", ""),
            edu.get("score", ""),
            edu.get("location", ""),
            left_bold=True, space_before=1.0, space_after=0
        )
        add_two_column_row(
            edu.get("degree", ""),
            edu.get("year", ""),
            left_bold=False, left_italic=True, right_bold=False, right_italic=True,
            space_before=0, space_after=0.8
        )

    # --------------------------------------------------------------------------
    # 4. TECHNICAL SKILLS
    # --------------------------------------------------------------------------
    add_section_heading('TECHNICAL SKILLS', space_before=3.2, space_after=1.0)
    for category, items in data.get("technical_skills", []):
        add_skill_row(category, items, space_before=0.8, space_after=0.8)

    # --------------------------------------------------------------------------
    # 5. SOFT SKILLS
    # --------------------------------------------------------------------------
    add_section_heading('SOFT SKILLS', space_before=3.2, space_after=1.0)
    for category, desc in data.get("soft_skills", []):
        add_bullet_point([(category, True, False), (desc, False, False)], space_before=0.8, space_after=0.8)

    # --------------------------------------------------------------------------
    # 6. PROFESSIONAL EXPERIENCE / INTERNSHIP
    # --------------------------------------------------------------------------
    add_section_heading('PROFESSIONAL EXPERIENCE', space_before=3.2, space_after=1.0)
    for exp in data.get("experience", []):
        add_two_column_row(
            exp.get("role", ""),
            exp.get("year", ""),
            left_bold=True, right_bold=False, space_before=1.0, space_after=0
        )
        add_two_column_row(
            exp.get("department", ""),
            exp.get("location", ""),
            left_bold=False, left_italic=True, right_bold=False, right_italic=True,
            space_before=0, space_after=0.2
        )
        for bullet in exp.get("bullets", []):
            add_bullet_point([(bullet, False, False)], space_before=0.7, space_after=0.7)

    # --------------------------------------------------------------------------
    # 7. PROJECTS
    # --------------------------------------------------------------------------
    add_section_heading('PROJECTS', space_before=3.2, space_after=1.0)
    for proj in data.get("projects", []):
        add_project_header(
            proj.get("title", ""),
            proj.get("tech_stack", ""),
            proj.get("year", ""),
            github_url=proj.get("github_url"),
            github_text=proj.get("github_text"),
            space_before=1.0, space_after=0
        )
        for bullet in proj.get("bullets", []):
            add_bullet_point([(bullet, False, False)], space_before=0.7, space_after=0.7)

    # --------------------------------------------------------------------------
    # 8. CERTIFICATIONS
    # --------------------------------------------------------------------------
    add_section_heading('CERTIFICATIONS', space_before=3.2, space_after=1.0)
    for provider, cert in data.get("certifications", []):
        add_bullet_point([(provider, True, False), (cert, False, False)], space_before=0.7, space_after=0.7)

    # --------------------------------------------------------------------------
    # 9. LANGUAGES
    # --------------------------------------------------------------------------
    add_section_heading('LANGUAGES', space_before=3.2, space_after=1.0)
    for lang, level in data.get("languages", []):
        add_bullet_point([(lang, True, False), (level, False, False)], space_before=0.6, space_after=0.6)

    # Save to disk
    doc.save(output_path)
    print(f"Successfully generated DOCX: {output_path}")


# ==============================================================================
# 4. DOCX TO PDF CONVERSION UTILITY
# ==============================================================================

def convert_docx_to_pdf(docx_path, pdf_path):
    """
    Converts a DOCX file into an exact PDF representation.

    Conversion Strategy:
    1. Primary method: Windows COM automation (`win32com.client.Dispatch('Word.Application')`).
       This uses Microsoft Word's own print engine, ensuring pixel-perfect fidelity,
       exact tab stop matching, and clean vector font embedding.
    2. Fallback method: `docx2pdf.convert`, which handles systems where Word COM
       dispatch might need a wrapped interface.

    Parameters:
    - docx_path: Path to the input .docx file.
    - pdf_path: Destination path for the output .pdf file.
    """
    try:
        import win32com.client
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        in_file = os.path.abspath(docx_path)
        out_file = os.path.abspath(pdf_path)
        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=17)  # 17 represents wdFormatPDF
        doc.Close()
        word.Quit()
        print(f"Successfully generated PDF: {pdf_path}")
    except Exception as e:
        print(f"win32com notice ({e}), attempting docx2pdf fallback...")
        try:
            from docx2pdf import convert
            convert(docx_path, pdf_path)
            print(f"Successfully generated PDF via docx2pdf: {pdf_path}")
        except Exception as e2:
            print(f"Error during PDF conversion: {e2}")


# ==============================================================================
# 5. PYMUPDF PAGE COUNT & MARGIN VALIDATOR
# ==============================================================================

def check_pdf_pages(pdf_path):
    """
    Analyzes the generated PDF using PyMuPDF (fitz) to guarantee a 1-page fit.

    What this function does:
    1. Opens the PDF and checks the total page count.
    2. Extracts all non-empty text blocks on page 1.
    3. Finds the maximum vertical coordinate (`max_y`) where content ends.
    4. Calculates the remaining bottom margin (`page_height - max_y`).
    5. Alerts the user if content spilled onto page 2.

    Parameters:
    - pdf_path: Filepath to the PDF to validate.
    """
    try:
        import fitz
        doc = fitz.open(pdf_path)
        page_count = len(doc)
        print(f"Validation: Total pages in {pdf_path} = {page_count}")
        if page_count > 0:
            page = doc[0]
            rect = page.rect
            blocks = page.get_text('blocks')
            if blocks:
                max_y = max(b[3] for b in blocks if b[4].strip())
                margin = rect.height - max_y
                print(f"Page Height: {rect.height:.1f} pt | Content Max Y: {max_y:.1f} pt | Bottom Margin: {margin:.1f} pt")
                if page_count == 1:
                    print("Status: PERFECT! Fits on exactly 1 page.")
                else:
                    print(f"Status: WARNING! Resume spilled onto {page_count} pages. Tighten spacing or content.")
        doc.close()
    except Exception as e:
        print(f"Page validation notice: {e}")


# ==============================================================================
# 6. MAIN EXECUTION PIPELINE
# ==============================================================================

if __name__ == '__main__':
    DOCX_OUTPUT = 'General_Resume_Template.docx'
    PDF_OUTPUT = 'General_Resume_Template.pdf'

    print("Generating general resume template...")
    # Step 1: Build the DOCX document from data
    create_resume(RESUME_DATA, DOCX_OUTPUT)
    
    # Step 2: Convert DOCX to PDF using native Word engine
    convert_docx_to_pdf(DOCX_OUTPUT, PDF_OUTPUT)
    
    # Step 3: Verify strict 1-page compliance
    check_pdf_pages(PDF_OUTPUT)
