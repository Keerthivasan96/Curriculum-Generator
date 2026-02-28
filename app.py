"""
Woxsen University — Course Outline Generator
School of Arts and Design | B.Des (Hons) Communication Design
"""

import os
import io
import requests
import streamlit as st

from dotenv import load_dotenv

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, Frame, PageTemplate, BaseDocTemplate
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ---------------------------------------------------------------------------
# Environment & Configuration
# ---------------------------------------------------------------------------

load_dotenv()

API_ENDPOINT = "https://api.groq.com/openai/v1/chat/completions"
API_MODEL    = "llama-3.1-8b-instant"
API_KEY      = os.getenv("GROQ_API_KEY") or st.secrets.get("GROQ_API_KEY", "")

LOGO_PATH    = os.path.join(os.path.dirname(__file__), "woxsen_logo.jpg")

# ---------------------------------------------------------------------------
# Brand Colors
# ---------------------------------------------------------------------------

WOXSEN_RED  = RGBColor(0xC0, 0x00, 0x00)
WOXSEN_DARK = RGBColor(0x1F, 0x1F, 0x1F)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)

RL_RED      = colors.HexColor("#C00000")
RL_DARK     = colors.HexColor("#1F1F1F")
RL_GREY_HDR = colors.HexColor("#595659")
RL_LIGHT    = colors.HexColor("#F2F2F2")
RL_FOOTER   = colors.HexColor("#3C3C3C")

# ---------------------------------------------------------------------------
# Content Generation Template
# ---------------------------------------------------------------------------

CONTENT_TEMPLATE = """You are an academic curriculum architect generating Woxsen University course outlines for B.Des (Hons) Communication Design, School of Arts and Design, Hyderabad.

Generate output ONLY using these exact section headers (copy verbatim):

SECTION 1: COURSE OVERVIEW
SECTION 2: BRIEF DESCRIPTION
SECTION 3: PROGRAM LEARNING OUTCOMES
SECTION 4: COURSE LEARNING OUTCOMES
SECTION 5: REFERENCES
SECTION 6: SESSION-WISE TOPICS
SECTION 7: EVALUATION COMPONENTS
SECTION 8: ASSESSMENT POLICY
SECTION 9: EVALUATION CLO ALIGNMENT
SECTION 10: RUBRICS
SECTION 11: APPENDICES
SECTION 12: ATTENDANCE AND ETHICS

Formatting Requirements:

SECTION 1 — one field per line:
  Course Title: [title]
  Course Code: [BDXX34XXX format]
  Credits: [number]
  Semester: [semester]
  Programme: B.Des (Hons) Communication Design
  School: School of Arts and Design

SECTION 2 — 2-3 formal academic paragraphs. No bullet points.

SECTION 3 — exactly 5 PLOs:
  PLO1: Visual Communication | [definition 15-20 words] | Yes
  PLO2: Brand Identity Development | [definition 15-20 words] | Yes
  PLO3: User Experience (UX) Design | [definition 15-20 words] | Yes
  PLO4: Research and Analysis | [definition 15-20 words] | Yes
  PLO5: Storytelling and Narrative | [definition 15-20 words] | Yes

SECTION 4 — 4-5 CLOs:
  CLO1: [outcome statement] | PLO1, PLO2

SECTION 5 — 8 references:
  R1 | [Book Title] | [Author, Edition Year] | CLO1, CLO2
  (continue R2-R8)

SECTION 6 — MINIMUM 16 sessions. Do not produce fewer than 16:
  Session 1 | [Topic] | [Session Intended Learning Outcome] | [Pedagogy] | CLO1 | R1
  (Session 16 or 17 = Final Jury, Pedagogy = Jury)
  Pedagogy types: Lecture, Case Study, Activity, Workshop, Practical, Discussion, Research, Jury, Review

SECTION 7 — evaluation breakdown (Internal 60 + End Term 40 = 100 marks total):
  Component 1 | Sessions 1-4 | 15 marks | Appendix A | CLO1, CLO2
  Component 2 | Sessions 5-8 | 15 marks | Appendix B | CLO2, CLO3
  Component 3 | Sessions 9-14 | 30 marks | Appendix C | CLO2, CLO3, CLO4
  End Term Jury | Sessions 15-16 | 40 marks | Appendix D | CLO4

SECTION 8 — one row per component:
  Sessions 1-4 | 15 | Appendix A | No AI | Scale 1 | CLO1, CLO2 | [outcome measured]
  Sessions 5-8 | 15 | Appendix B | No AI | Scale 1 | CLO2, CLO3 | [outcome measured]
  Sessions 9-14 | 30 | Appendix C | No AI | Scale 1 | CLO2-CLO4 | [outcome measured]
  End Term | 40 | Appendix D | No AI | Scale 1 | CLO4 | [outcome measured]

SECTION 9 — narrative (2-3 paragraphs) then alignment table:
  [Component Name] | [CLO Alignment] | [Description 20-30 words]

SECTION 10 — one rubric row per CLO:
  CLO1 | [Exceeds >80 description] | [Meets 55-79 description] | [Below <55 description]

SECTION 11 — 4 appendices (A, B, C, D). For each:
  APPENDIX A: [Name]
  Objective: [one sentence]
  Brief Description: [2-3 sentences]
  Submission Requirements:
  - [item]
  Evaluation Criteria:
  [Criteria] | [Marks] | [Description]
  Appendix D = Final Jury (40 marks): Product (15), Process & Reasoning (15), Presentation (5), Personality & Confidence (5)

SECTION 12 — use this exact text verbatim:

Attendance & Punctuality
Learning is an interactive process. Students are expected to be present in all the classes. Absence is only appropriate in exceptional circumstances. Voluntary activities are never valid reasons for missing any class.
Students may refer to the student handbook for regulations covering attendance.
Students who do not meet attendance requirements will not be permitted to write the end term examination and will be required to repeat the course with the next batch of students.
Late arrival is disruptive to the learning environment; students are expected to be in class before the scheduled commencement time. Students arriving for class after the scheduled commencement time will be turned away unless they have a valid reason to be permitted to attend.

Copyright
The content provided by the faculty in the class is copy-righted. Students are instructed not to distribute or share content used during courses with external entities.

Student Code of Ethics
Each student enrolled in this course accepts personal responsibility to uphold and defend academic integrity and to promote an atmosphere in which all individuals may flourish. The Students' Code of Ethics strives to set a standard of honest behaviour that reflects well on students and the school. All students enrolled in these courses are expected to follow the Students' Code of Ethics contained in the student handbook. Unethical and unfair practices adopted by students may lead to penalties such as having to repeat the course or having the student's enrollment cancelled.

Output Rules:
- No markdown (no **, ###, ---)
- No commentary outside the 12 sections
- All pipe-separated rows must have consistent columns
- Section 6 must contain a minimum of 16 sessions
- Section 7 total marks must equal exactly 100"""


# ---------------------------------------------------------------------------
# Section Keys
# ---------------------------------------------------------------------------

SECTION_KEYS = [
    "SECTION 1: COURSE OVERVIEW",
    "SECTION 2: BRIEF DESCRIPTION",
    "SECTION 3: PROGRAM LEARNING OUTCOMES",
    "SECTION 4: COURSE LEARNING OUTCOMES",
    "SECTION 5: REFERENCES",
    "SECTION 6: SESSION-WISE TOPICS",
    "SECTION 7: EVALUATION COMPONENTS",
    "SECTION 8: ASSESSMENT POLICY",
    "SECTION 9: EVALUATION CLO ALIGNMENT",
    "SECTION 10: RUBRICS",
    "SECTION 11: APPENDICES",
    "SECTION 12: ATTENDANCE AND ETHICS",
]


# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------

def generate_outline(topic: str, semester: str, credits: int) -> str:
    if not API_KEY:
        raise ValueError("API key not configured. Add GROQ_API_KEY to your environment or Streamlit secrets.")

    user_message = (
        f"Generate a complete Woxsen University course outline for:\n"
        f"Course Topic: {topic}\n"
        f"Semester: {semester}\n"
        f"Credits: {credits}\n"
        f"Programme: B.Des (Hons) Communication Design\n"
        f"School: School of Arts and Design, Woxsen University, Hyderabad\n"
        f"Section 6 must have at least 16 sessions. Section 7 must total exactly 100 marks."
    )

    response = requests.post(
        API_ENDPOINT,
        headers={"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"},
        json={
            "model": API_MODEL,
            "messages": [
                {"role": "system", "content": CONTENT_TEMPLATE},
                {"role": "user",   "content": user_message},
            ],
            "temperature": 0.3,
            "max_tokens": 4000,
        },
        timeout=120,
    )

    if response.status_code != 200:
        raise Exception(f"Request failed ({response.status_code}): {response.text}")

    return response.json()["choices"][0]["message"]["content"]


# ---------------------------------------------------------------------------
# Parsing Utilities
# ---------------------------------------------------------------------------

def parse_sections(raw: str) -> dict:
    sections = {}
    for i, key in enumerate(SECTION_KEYS):
        next_key = SECTION_KEYS[i + 1] if i + 1 < len(SECTION_KEYS) else None
        start = raw.find(key)
        if start == -1:
            sections[key] = ""
            continue
        start += len(key)
        end = raw.find(next_key, start) if next_key else len(raw)
        sections[key] = raw[start:end].strip()
    return sections


def extract_meta(text: str, semester: str, credits: int) -> dict:
    meta = {
        "title":     "Course Title",
        "code":      "BDXX34XXX",
        "credits":   str(credits),
        "semester":  semester,
        "programme": "B.Des (Hons) Communication Design",
        "school":    "School of Arts and Design",
    }
    for line in text.split("\n"):
        l   = line.strip()
        low = l.lower()
        if low.startswith("course title"):   meta["title"]     = l.split(":", 1)[-1].strip()
        elif low.startswith("course code"):  meta["code"]      = l.split(":", 1)[-1].strip()
        elif low.startswith("credits"):      meta["credits"]   = l.split(":", 1)[-1].strip()
        elif low.startswith("semester"):     meta["semester"]  = l.split(":", 1)[-1].strip()
        elif low.startswith("programme"):    meta["programme"] = l.split(":", 1)[-1].strip()
        elif low.startswith("school"):       meta["school"]    = l.split(":", 1)[-1].strip()
    return meta


def pipe_rows(text: str, min_cols: int = 2) -> list:
    rows = []
    for line in text.split("\n"):
        line = line.strip()
        if "|" in line:
            parts = [p.strip() for p in line.split("|") if p.strip()]
            if len(parts) >= min_cols:
                rows.append(parts)
    return rows


def pad_rows(rows: list, n_cols: int) -> list:
    result = []
    for r in rows:
        r = list(r)
        while len(r) < n_cols:
            r.append("")
        result.append(r[:n_cols])
    return result


# ---------------------------------------------------------------------------
# DOCX Helpers
# ---------------------------------------------------------------------------

def set_cell_bg(cell, hex_color: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)


def add_red_border_bottom(para, size: int = 12):
    pPr  = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot  = OxmlElement("w:bottom")
    bot.set(qn("w:val"),   "single")
    bot.set(qn("w:sz"),    str(size))
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), "C00000")
    pBdr.append(bot)
    pPr.append(pBdr)


def add_section_header(doc: Document, title: str):
    p   = doc.add_paragraph()
    run = p.add_run(title)
    run.font.bold      = True
    run.font.size      = Pt(11)
    run.font.color.rgb = WOXSEN_RED
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(4)
    add_red_border_bottom(p, 6)


def add_body_text(doc: Document, text: str, bold: bool = False):
    p   = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(10)
    run.font.bold = bold
    p.paragraph_format.space_after = Pt(3)


def add_styled_table(doc: Document, headers: list, rows: list, col_widths_cm: list = None):
    n   = len(headers)
    tbl = doc.add_table(rows=1 + len(rows), cols=n)
    tbl.style = "Table Grid"

    for i, h in enumerate(headers):
        cell = tbl.rows[0].cells[i]
        cell.text = h
        set_cell_bg(cell, "595659")
        if col_widths_cm:
            cell.width = Cm(col_widths_cm[i])
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.bold      = True
                r.font.color.rgb = WHITE
                r.font.size      = Pt(9)

    for ri, row_data in enumerate(rows):
        row = tbl.rows[ri + 1]
        bg  = "F2F2F2" if ri % 2 == 0 else "FFFFFF"
        for ci, txt in enumerate(row_data):
            cell      = row.cells[ci]
            cell.text = str(txt)
            set_cell_bg(cell, bg)
            if col_widths_cm and ci < len(col_widths_cm):
                cell.width = Cm(col_widths_cm[ci])
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(9)

    doc.add_paragraph()


def configure_page_header(doc: Document):
    section = doc.sections[0]
    header  = section.header
    header.is_linked_to_previous = False

    for p in header.paragraphs:
        for run in p.runs:
            run.text = ""

    hdr_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    hdr_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    if os.path.exists(LOGO_PATH):
        run = hdr_para.add_run()
        run.add_picture(LOGO_PATH, width=Cm(4.2))
    else:
        run = hdr_para.add_run("WOXSEN UNIVERSITY")
        run.font.bold = True


def configure_page_footer(doc: Document):
    section = doc.sections[0]
    footer  = section.footer
    footer.is_linked_to_previous = False

    for p in footer.paragraphs:
        p.clear()

    fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    pPr = fp._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  "3C3C3C")
    pPr.append(shd)

    run = fp.add_run()
    run.font.color.rgb = WHITE
    run.font.size      = Pt(9)

    for tag, val in [("begin", None), (None, " PAGE "), ("end", None)]:
        if tag:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tag)
            run._r.append(el)
        else:
            instr = OxmlElement("w:instrText")
            instr.text = val
            run._r.append(instr)


# ---------------------------------------------------------------------------
# DOCX Builder
# ---------------------------------------------------------------------------

def build_docx(raw: str, semester: str, credits: int,
               faculty_name: str = "", faculty_email: str = "",
               course_year: str = "2025-26") -> bytes:

    secs = parse_sections(raw)
    meta = extract_meta(secs.get("SECTION 1: COURSE OVERVIEW", ""), semester, credits)

    doc = Document()
    for sec in doc.sections:
        sec.top_margin      = Cm(2.5)
        sec.bottom_margin   = Cm(2.0)
        sec.left_margin     = Cm(2.5)
        sec.right_margin    = Cm(2.5)
        sec.header_distance = Cm(1.0)
        sec.footer_distance = Cm(0.8)

    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(10)

    configure_page_header(doc)
    configure_page_footer(doc)

    # Cover — red divider line
    red_line = doc.add_paragraph()
    add_red_border_bottom(red_line, 18)
    red_line.paragraph_format.space_before = Pt(0)
    red_line.paragraph_format.space_after  = Pt(2)

    # "Course Outline" label
    co = doc.add_paragraph()
    co.add_run("Course Outline").font.size = Pt(10)
    co.alignment = WD_ALIGN_PARAGRAPH.CENTER
    co.paragraph_format.space_before = Pt(2)
    co.paragraph_format.space_after  = Pt(14)

    # Course title
    title_p = doc.add_paragraph()
    title_r = title_p.add_run(meta["title"].upper())
    title_r.font.bold      = True
    title_r.font.size      = Pt(13)
    title_r.font.color.rgb = WOXSEN_DARK
    title_p.paragraph_format.space_after = Pt(2)

    for line in [
        f"Course Code: {meta['code']}",
        f"Semester {meta['semester']}",
        meta["programme"],
        f"Communication Design: {course_year}",
        f"{meta['credits']} Credits",
    ]:
        p   = doc.add_paragraph()
        run = p.add_run(line)
        run.font.size      = Pt(10)
        run.font.color.rgb = WOXSEN_DARK
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)

    doc.add_paragraph()
    doc.add_paragraph()

    # Prepared by table
    prep = doc.add_table(rows=2, cols=2)
    prep.style = "Table Grid"
    for i, (lbl, val) in enumerate([("Prepared by", faculty_name), ("Email ID", faculty_email)]):
        prep.rows[i].cells[0].text  = lbl
        prep.rows[i].cells[1].text  = val
        prep.rows[i].cells[0].width = Cm(4)
        prep.rows[i].cells[1].width = Cm(12)
        for cell in prep.rows[i].cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(10)

    doc.add_page_break()

    # Brief Description
    add_section_header(doc, "Brief Description and Relevance of the Course")
    for line in secs.get("SECTION 2: BRIEF DESCRIPTION", "").split("\n"):
        if line.strip():
            add_body_text(doc, line.strip())

    # Programme Learning Outcomes
    add_section_header(doc, "Programme Learning Outcomes")
    plo_rows = []
    for line in secs.get("SECTION 3: PROGRAM LEARNING OUTCOMES", "").split("\n"):
        l = line.strip()
        if l.lower().startswith("plo") and "|" in l:
            parts = [p.strip() for p in l.split("|")]
            while len(parts) < 3:
                parts.append("")
            plo_rows.append(parts[:3])
    if plo_rows:
        add_styled_table(doc, ["PLO", "Definition", "Addressed"], plo_rows, [3.5, 10, 2.5])

    # Course Learning Outcomes
    add_section_header(doc, "Course Learning Outcomes")
    clo_rows = []
    for line in secs.get("SECTION 4: COURSE LEARNING OUTCOMES", "").split("\n"):
        l = line.strip()
        if l.lower().startswith("clo") and "|" in l:
            parts = [p.strip() for p in l.split("|")]
            while len(parts) < 2:
                parts.append("")
            clo_rows.append(parts[:2])
    if clo_rows:
        add_styled_table(doc, ["Course Learning Outcome", "Mapping to Programme LO"], clo_rows, [12, 4])

    # Prerequisites
    add_section_header(doc, "Prerequisites")
    add_body_text(doc, "NIL")

    # References
    add_section_header(doc, "Reference / Reading Material Recommended")
    ref_rows = []
    for line in secs.get("SECTION 5: REFERENCES", "").split("\n"):
        l = line.strip()
        if "|" in l and l and (l[0] in "R0123456789"):
            parts = [p.strip() for p in l.split("|")]
            while len(parts) < 4:
                parts.append("")
            ref_rows.append(parts[:4])
    if ref_rows:
        add_styled_table(doc,
            ["Code", "Textbook / Article / Report Name", "Edition / Year / Link", "CLO Mapped"],
            ref_rows, [1.5, 8, 4, 2.5])

    # Session Table
    add_section_header(doc, "Session-Wise Topics and Reading / References")
    sess_rows = pipe_rows(secs.get("SECTION 6: SESSION-WISE TOPICS", ""), min_cols=3)
    if sess_rows:
        norm = pad_rows(sess_rows, 6)
        add_styled_table(doc,
            ["Sn.", "Topic", "Session Intended Learning Outcome", "Pedagogy", "CLO", "Reading Material"],
            norm, [1.0, 4.5, 5.5, 2.5, 1.0, 2.0])

    # Evaluation Components
    add_section_header(doc, "Performance Evaluation Components for the Course")
    eval_rows = pipe_rows(secs.get("SECTION 7: EVALUATION COMPONENTS", ""), min_cols=3)
    if eval_rows:
        add_styled_table(doc, ["Session No.", "Marks", "Evaluation Form", "CLO"],
                         pad_rows(eval_rows, 4), [3, 2, 8, 3])

    # Assessment Policy (formerly "AI Assessment")
    add_section_header(doc, "Assessment Policy")
    ap_rows = pipe_rows(secs.get("SECTION 8: ASSESSMENT POLICY", ""), min_cols=5)
    if ap_rows:
        add_styled_table(doc,
            ["Session No.", "Marks", "Evaluation Form", "Assessment Level", "Scale", "CLO", "Outcome Measured"],
            pad_rows(ap_rows, 7), [2.0, 1.5, 3.5, 2.5, 1.5, 2.0, 3.5])

    # Evaluation & CLO Alignment
    add_section_header(doc, "Evaluation Components & CLO Alignment")
    sec9 = secs.get("SECTION 9: EVALUATION CLO ALIGNMENT", "")
    table9, narrative9 = [], []
    for line in sec9.split("\n"):
        l = line.strip()
        if not l:
            continue
        if "|" in l:
            parts = [p.strip() for p in l.split("|")]
            if len(parts) >= 2:
                table9.append(parts)
        else:
            narrative9.append(l)
    for nl in narrative9:
        add_body_text(doc, nl)
    if table9:
        add_styled_table(doc, ["Evaluation Component", "CLO Alignment", "Description"],
                         pad_rows(table9, 3), [5, 3, 8.5])

    # Rubrics
    add_section_header(doc, "Evaluation Rubrics")
    rubric_rows = pipe_rows(secs.get("SECTION 10: RUBRICS", ""), min_cols=4)
    if rubric_rows:
        add_styled_table(doc,
            ["CLO", "Exceeds Expectations (>80)", "Meets Expectations (55–79)", "Does Not Meet Expectations (<55)"],
            pad_rows(rubric_rows, 4), [2, 5, 5, 4])

    # Appendices
    doc.add_page_break()
    sec11        = secs.get("SECTION 11: APPENDICES", "")
    first_app    = True
    in_criteria  = False
    criteria_buf = []

    def flush_criteria():
        if criteria_buf:
            add_styled_table(doc, ["Criteria", "Marks", "Description"],
                             pad_rows(criteria_buf, 3), [5, 2, 9.5])
            criteria_buf.clear()

    for line in sec11.split("\n"):
        line = line.strip()
        if not line:
            continue

        if line.upper().startswith("APPENDIX"):
            flush_criteria()
            in_criteria = False
            if not first_app:
                doc.add_page_break()
            first_app = False
            p   = doc.add_paragraph()
            run = p.add_run(line)
            run.font.bold      = True
            run.font.size      = Pt(13)
            run.font.color.rgb = WOXSEN_DARK
            add_red_border_bottom(p, 8)
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after  = Pt(8)

        elif "evaluation criteria" in line.lower():
            flush_criteria()
            in_criteria = True
            add_body_text(doc, line, bold=True)

        elif in_criteria and "|" in line:
            parts = [p.strip() for p in line.split("|")]
            if len(parts) >= 2:
                criteria_buf.append(parts)

        elif line.startswith("-") or line.startswith("•"):
            p   = doc.add_paragraph()
            run = p.add_run(line.lstrip("-•").strip())
            run.font.size = Pt(9)
            p.paragraph_format.left_indent = Cm(0.5)
            p.paragraph_format.space_after  = Pt(2)

        elif any(line.lower().startswith(k) for k in ["objective", "brief description",
                                                        "submission requirements"]):
            add_body_text(doc, line, bold=True)
        else:
            add_body_text(doc, line)

    flush_criteria()

    # Attendance, Copyright, Ethics
    doc.add_page_break()
    for line in secs.get("SECTION 12: ATTENDANCE AND ETHICS", "").split("\n"):
        line = line.strip()
        if not line:
            continue
        if any(line.lower().startswith(k) for k in ["attendance", "copyright", "student code"]):
            p   = doc.add_paragraph()
            run = p.add_run(line)
            run.font.bold = True
            run.font.size = Pt(10)
            p.paragraph_format.space_before = Pt(8)
        else:
            add_body_text(doc, line)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# PDF — Page Template with logo + footer on every page
# ---------------------------------------------------------------------------

class WoxsenTemplate(BaseDocTemplate):
    def __init__(self, filename, logo_path, **kwargs):
        self.logo_path = logo_path
        BaseDocTemplate.__init__(self, filename, **kwargs)
        frame    = Frame(self.leftMargin, self.bottomMargin, self.width, self.height, id="main")
        template = PageTemplate(id="All", frames=frame, onPage=self._draw_chrome)
        self.addPageTemplates([template])

    def _draw_chrome(self, canvas, doc):
        canvas.saveState()

        # Dark footer bar
        canvas.setFillColor(RL_FOOTER)
        canvas.rect(0, 0, doc.pagesize[0], 18, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica", 9)
        canvas.drawRightString(doc.pagesize[0] - 15, 5, str(canvas.getPageNumber()))

        # Logo top-right
        if os.path.exists(self.logo_path):
            logo_w = 4.2 * cm
            logo_h = 2.2 * cm
            x = doc.pagesize[0] - doc.rightMargin - logo_w
            y = doc.pagesize[1] - doc.topMargin - logo_h + 0.4 * cm
            canvas.drawImage(self.logo_path, x, y,
                             width=logo_w, height=logo_h,
                             preserveAspectRatio=True, mask="auto")

        canvas.restoreState()


def build_pdf(raw: str, semester: str, credits: int,
              faculty_name: str = "", faculty_email: str = "",
              course_year: str = "2025-26") -> bytes:

    secs = parse_sections(raw)
    meta = extract_meta(secs.get("SECTION 1: COURSE OVERVIEW", ""), semester, credits)

    buf = io.BytesIO()
    doc = WoxsenTemplate(
        buf,
        logo_path=LOGO_PATH,
        pagesize=A4,
        leftMargin=2.5 * cm,
        rightMargin=2.5 * cm,
        topMargin=3.2 * cm,
        bottomMargin=1.5 * cm,
    )

    def style(name, **kw):
        defaults = dict(fontName="Helvetica", fontSize=9.5, leading=14, spaceAfter=3)
        defaults.update(kw)
        return ParagraphStyle(name, **defaults)

    s_normal = style("normal")
    s_detail = style("detail", textColor=RL_DARK, spaceAfter=1)
    s_title  = style("title",  fontName="Helvetica-Bold", fontSize=13, leading=16, textColor=RL_DARK)
    s_center = style("center", alignment=TA_CENTER, spaceAfter=10)
    s_sec    = style("sec",    fontName="Helvetica-Bold", fontSize=10.5, textColor=RL_RED, spaceBefore=12)
    s_sub    = style("sub",    fontName="Helvetica-Bold", spaceBefore=6)
    s_app    = style("app",    fontName="Helvetica-Bold", fontSize=12, leading=16, spaceBefore=10, spaceAfter=6)
    s_bullet = style("bullet", fontSize=9, leftIndent=12, spaceAfter=2)

    story = []

    def section_header(title):
        story.append(Paragraph(title, s_sec))
        story.append(HRFlowable(width="100%", thickness=1.5, color=RL_RED, spaceAfter=4))

    def styled_table(headers, rows, col_widths):
        data = [[
            Paragraph(f"<b>{h}</b>",
                      ParagraphStyle("th", fontName="Helvetica-Bold", fontSize=8,
                                     leading=11, textColor=colors.white))
            for h in headers
        ]]
        for row in rows:
            data.append([
                Paragraph(str(v), ParagraphStyle("td", fontName="Helvetica", fontSize=8, leading=11))
                for v in row
            ])
        tbl = Table(data, colWidths=col_widths)
        tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0),  RL_GREY_HDR),
            ("TEXTCOLOR",     (0, 0), (-1, 0),  colors.white),
            ("BOX",           (0, 0), (-1, -1), 0.5, colors.black),
            ("INNERGRID",     (0, 0), (-1, -1), 0.3, colors.grey),
            ("ROWBACKGROUNDS",(0, 1), (-1, -1), [RL_LIGHT, colors.white]),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING",   (0, 0), (-1, -1), 5),
            ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 4 * mm))

    W = doc.width

    # Cover
    story.append(Spacer(1, 4 * mm))
    story.append(HRFlowable(width="100%", thickness=3, color=RL_RED, spaceAfter=2))
    story.append(Paragraph("Course Outline", s_center))
    story.append(Paragraph(meta["title"].upper(), s_title))
    story.append(Paragraph(f"<b>Course Code:</b> {meta['code']}", s_detail))
    story.append(Paragraph(f"Semester {meta['semester']}", s_detail))
    story.append(Paragraph(meta["programme"], s_detail))
    story.append(Paragraph(f"Communication Design: {course_year}", s_detail))
    story.append(Paragraph(f"{meta['credits']} Credits", s_detail))
    story.append(Spacer(1, 10 * mm))

    prep_data = [
        [Paragraph("<b>Prepared by</b>", s_normal), Paragraph(faculty_name, s_normal)],
        [Paragraph("<b>Email ID</b>",    s_normal), Paragraph(faculty_email, s_normal)],
    ]
    prep_tbl = Table(prep_data, colWidths=[4 * cm, 12 * cm])
    prep_tbl.setStyle(TableStyle([
        ("BOX",          (0, 0), (-1, -1), 0.5, colors.black),
        ("INNERGRID",    (0, 0), (-1, -1), 0.5, colors.black),
        ("TOPPADDING",   (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 5),
        ("LEFTPADDING",  (0, 0), (-1, -1), 6),
    ]))
    story.append(prep_tbl)
    story.append(PageBreak())

    # Brief Description
    section_header("Brief Description and Relevance of the Course")
    for line in secs.get("SECTION 2: BRIEF DESCRIPTION", "").split("\n"):
        if line.strip():
            story.append(Paragraph(line.strip(), s_normal))

    # PLO
    section_header("Programme Learning Outcomes")
    plo_rows = []
    for line in secs.get("SECTION 3: PROGRAM LEARNING OUTCOMES", "").split("\n"):
        l = line.strip()
        if l.lower().startswith("plo") and "|" in l:
            parts = [p.strip() for p in l.split("|")]
            while len(parts) < 3:
                parts.append("")
            plo_rows.append(parts[:3])
    if plo_rows:
        styled_table(["PLO", "Definition", "Addressed"], plo_rows, [3.5*cm, 10*cm, 2.5*cm])

    # CLO
    section_header("Course Learning Outcomes")
    clo_rows = []
    for line in secs.get("SECTION 4: COURSE LEARNING OUTCOMES", "").split("\n"):
        l = line.strip()
        if l.lower().startswith("clo") and "|" in l:
            parts = [p.strip() for p in l.split("|")]
            while len(parts) < 2:
                parts.append("")
            clo_rows.append(parts[:2])
    if clo_rows:
        styled_table(["Course Learning Outcome", "Mapping to Programme LO"], clo_rows, [12*cm, 4*cm])

    # Prerequisites
    section_header("Prerequisites")
    story.append(Paragraph("NIL", s_normal))

    # References
    section_header("Reference / Reading Material Recommended")
    ref_rows = []
    for line in secs.get("SECTION 5: REFERENCES", "").split("\n"):
        l = line.strip()
        if "|" in l and l and l[0] in "R0123456789":
            parts = [p.strip() for p in l.split("|")]
            while len(parts) < 4:
                parts.append("")
            ref_rows.append(parts[:4])
    if ref_rows:
        styled_table(["Code", "Textbook / Article / Report Name", "Edition / Year / Link", "CLO Mapped"],
                     ref_rows, [1.5*cm, 7.5*cm, 4*cm, 3*cm])

    # Sessions
    section_header("Session-Wise Topics and Reading / References")
    sess_rows = pipe_rows(secs.get("SECTION 6: SESSION-WISE TOPICS", ""), min_cols=3)
    if sess_rows:
        styled_table(
            ["Sn.", "Topic", "Session Intended Learning Outcome", "Pedagogy", "CLO", "Reading Material"],
            pad_rows(sess_rows, 6),
            [1*cm, 3.8*cm, 5*cm, 2.5*cm, 1.2*cm, 2.5*cm])

    # Evaluation Components
    section_header("Performance Evaluation Components for the Course")
    eval_rows = pipe_rows(secs.get("SECTION 7: EVALUATION COMPONENTS", ""), min_cols=3)
    if eval_rows:
        styled_table(["Session No.", "Marks", "Evaluation Form", "CLO"],
                     pad_rows(eval_rows, 4), [3*cm, 2*cm, 8*cm, 3*cm])

    # Assessment Policy
    section_header("Assessment Policy")
    ap_rows = pipe_rows(secs.get("SECTION 8: ASSESSMENT POLICY", ""), min_cols=5)
    if ap_rows:
        styled_table(
            ["Session No.", "Marks", "Evaluation Form", "Assessment Level", "Scale", "CLO", "Outcome Measured"],
            pad_rows(ap_rows, 7),
            [1.8*cm, 1.3*cm, 3*cm, 2.5*cm, 1.5*cm, 2*cm, 3.9*cm])

    # CLO Alignment
    section_header("Evaluation Components & CLO Alignment")
    sec9 = secs.get("SECTION 9: EVALUATION CLO ALIGNMENT", "")
    table9, narrative9 = [], []
    for line in sec9.split("\n"):
        l = line.strip()
        if not l:
            continue
        if "|" in l:
            parts = [p.strip() for p in l.split("|")]
            if len(parts) >= 2:
                table9.append(parts)
        else:
            narrative9.append(l)
    for nl in narrative9:
        story.append(Paragraph(nl, s_normal))
    if table9:
        styled_table(["Evaluation Component", "CLO Alignment", "Description"],
                     pad_rows(table9, 3), [5*cm, 3*cm, 8*cm])

    # Rubrics
    section_header("Evaluation Rubrics")
    rubric_rows = pipe_rows(secs.get("SECTION 10: RUBRICS", ""), min_cols=4)
    if rubric_rows:
        styled_table(
            ["CLO", "Exceeds Expectations (>80)", "Meets Expectations (55–79)", "Does Not Meet Expectations (<55)"],
            pad_rows(rubric_rows, 4),
            [1.5*cm, 5*cm, 5*cm, 4.5*cm])

    # Appendices
    story.append(PageBreak())
    sec11       = secs.get("SECTION 11: APPENDICES", "")
    first_app   = True
    in_criteria = False
    crit_buf    = []

    def flush_pdf_criteria():
        if crit_buf:
            styled_table(["Criteria", "Marks", "Description"], pad_rows(crit_buf, 3), [5*cm, 2*cm, 9*cm])
            crit_buf.clear()

    for line in sec11.split("\n"):
        line = line.strip()
        if not line:
            continue
        if line.upper().startswith("APPENDIX"):
            flush_pdf_criteria()
            in_criteria = False
            if not first_app:
                story.append(PageBreak())
            first_app = False
            story.append(Paragraph(line, s_app))
            story.append(HRFlowable(width="100%", thickness=1.5, color=RL_RED, spaceAfter=4))
        elif "evaluation criteria" in line.lower():
            flush_pdf_criteria()
            in_criteria = True
            story.append(Paragraph(f"<b>{line}</b>", s_normal))
        elif in_criteria and "|" in line:
            parts = [p.strip() for p in line.split("|")]
            if len(parts) >= 2:
                crit_buf.append(parts)
        elif line.startswith("-") or line.startswith("•"):
            story.append(Paragraph("• " + line.lstrip("-•").strip(), s_bullet))
        elif any(line.lower().startswith(k) for k in ["objective", "brief description", "submission"]):
            story.append(Paragraph(f"<b>{line}</b>", s_normal))
        else:
            story.append(Paragraph(line, s_normal))

    flush_pdf_criteria()

    # Attendance, Copyright, Ethics
    story.append(PageBreak())
    for line in secs.get("SECTION 12: ATTENDANCE AND ETHICS", "").split("\n"):
        line = line.strip()
        if not line:
            continue
        if any(line.lower().startswith(k) for k in ["attendance", "copyright", "student code"]):
            story.append(Paragraph(f"<b>{line}</b>", s_sub))
        else:
            story.append(Paragraph(line, s_normal))

    story.append(Spacer(1, 8 * mm))
    story.append(HRFlowable(width="100%", thickness=2, color=RL_RED))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Streamlit Interface
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Woxsen Course Outline Generator",
    page_icon="🎓",
    layout="wide",
)

st.markdown("""
<style>
.main { background-color: #f9f9f9; }
section[data-testid="stSidebar"] { background-color: #1a1a2e; }
section[data-testid="stSidebar"] * { color: #ffffff !important; }
.stButton > button {
    background-color: #C00000;
    color: white;
    border: none;
    border-radius: 8px;
    height: 3em;
    font-weight: 600;
    letter-spacing: 0.03em;
}
.stButton > button:hover { background-color: #a00000; }
.stTextInput > div > div > input,
.stSelectbox > div > div { border-radius: 6px; }
.block-container { padding-top: 2rem; }
</style>
""", unsafe_allow_html=True)

# Header
col_logo, col_title = st.columns([1, 5])
with col_logo:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=90)
with col_title:
    st.markdown("## Woxsen University — Course Outline Generator")
    st.markdown("<span style='color:#888;font-size:0.9em'>School of Arts & Design &nbsp;|&nbsp; B.Des (Hons) Communication Design</span>", unsafe_allow_html=True)

st.divider()

# Sidebar
with st.sidebar:
    st.markdown("### ⚙️ Course Parameters")
    semester    = st.selectbox("Semester",      ["I", "II", "III", "IV", "V", "VI"], index=5)
    credits     = st.selectbox("Credits",       [2, 3, 4, 5], index=1)
    course_year = st.text_input("Academic Year", value="2025-26")

    st.markdown("---")
    st.markdown("### 👤 Faculty Details")
    faculty_name  = st.text_input("Faculty Name",  placeholder="Prof. Jane Doe, Assistant Professor")
    faculty_email = st.text_input("Faculty Email", placeholder="jane.doe@woxsen.edu.in")

    st.markdown("---")
    if os.path.exists(LOGO_PATH):
        st.success("✅ Logo file found")
    else:
        st.warning("⚠️ Place woxsen_logo.jpg in the same folder as app.py")

    if not API_KEY:
        st.error("❌ GROQ_API_KEY not set")
    else:
        st.success("✅ API key configured")

# Main Input Area
with st.container():
    st.markdown("### 📚 Course Topic")
    col_input, col_btn = st.columns([4, 1])
    with col_input:
        course_topic = st.text_input(
            "Enter the course topic",
            placeholder="e.g. Introduction to Metaverse Design, Typography for Digital Media...",
            label_visibility="collapsed",
        )
    with col_btn:
        generate_btn = st.button("✨ Generate", use_container_width=True, type="primary")

# Generation
if generate_btn:
    if not course_topic.strip():
        st.error("Please enter a course topic before generating.")
        st.stop()

    progress = st.progress(0, text="Initialising content engine...")
    try:
        progress.progress(20, text="Sending request...")
        raw = generate_outline(course_topic.strip(), semester, credits)
        progress.progress(80, text="Processing output...")
        st.session_state.update({
            "raw":          raw,
            "topic":        course_topic.strip(),
            "semester":     semester,
            "credits":      credits,
            "faculty_name": faculty_name,
            "faculty_email":faculty_email,
            "course_year":  course_year,
        })
        progress.progress(100, text="Done.")
        st.success("✅ Course outline generated. Switch to the Download tab.")
    except Exception as e:
        progress.empty()
        st.error(f"❌ {e}")
        st.stop()

# Output Tabs
if "raw" in st.session_state:
    st.divider()
    tab_preview, tab_download = st.tabs(["📄 Preview", "⬇️ Download"])

    with tab_preview:
        st.text_area("Generated Content", value=st.session_state["raw"],
                     height=520, label_visibility="collapsed")

    with tab_download:
        st.markdown("### Download Course Outline")
        fname = st.session_state["topic"].replace(" ", "_")[:40]

        col_a, col_b, col_c = st.columns(3)

        with col_a:
            st.markdown("**Word Document**")
            if st.button("📝 Build DOCX", use_container_width=True):
                with st.spinner("Building document..."):
                    try:
                        data = build_docx(
                            st.session_state["raw"],
                            st.session_state["semester"],
                            st.session_state["credits"],
                            st.session_state.get("faculty_name", ""),
                            st.session_state.get("faculty_email", ""),
                            st.session_state.get("course_year", "2025-26"),
                        )
                        st.download_button(
                            "⬇️ Download .docx",
                            data=data,
                            file_name=f"Woxsen_{fname}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.error(f"Build failed: {e}")

        with col_b:
            st.markdown("**PDF Document**")
            if st.button("📄 Build PDF", use_container_width=True):
                with st.spinner("Building document..."):
                    try:
                        data = build_pdf(
                            st.session_state["raw"],
                            st.session_state["semester"],
                            st.session_state["credits"],
                            st.session_state.get("faculty_name", ""),
                            st.session_state.get("faculty_email", ""),
                            st.session_state.get("course_year", "2025-26"),
                        )
                        st.download_button(
                            "⬇️ Download .pdf",
                            data=data,
                            file_name=f"Woxsen_{fname}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.error(f"Build failed: {e}")

        with col_c:
            st.markdown("**Regenerate**")
            if st.button("🔄 Try Again", use_container_width=True):
                with st.spinner("Regenerating..."):
                    try:
                        raw = generate_outline(
                            st.session_state["topic"],
                            st.session_state["semester"],
                            st.session_state["credits"],
                        )
                        st.session_state["raw"] = raw
                        st.rerun()
                    except Exception as e:
                        st.error(f"Failed: {e}")