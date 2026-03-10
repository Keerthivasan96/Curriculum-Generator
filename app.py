"""
Woxsen University — AI Curriculum Assistant
School of Arts and Design | B.Des (Hons) Communication Design

FIXES (v2):
  - Extract & Apply Changes (Tab 3) now reliably renders after button click
  - st.rerun() also fires in except-block so errors surface correctly
  - Guard added: button blocked when chat history is empty
  - Empty-result guard: warning shown instead of silent blank UI
  - API calls wrapped with explicit error messages surfaced to UI

NEW FEATURES (v2):
  - 📊 Curriculum Health Score — instant at-a-glance scorecard after upload
  - 💾 Save / Load Chat Session — export chat as JSON, reload later
  - 📝 Quick Edit Mode — direct in-page text editing of generated curriculum
  - 🔁 Change History — track applied change sets with timestamps
  - 🎯 Suggested Changes Toolbar — one-click common improvements
"""

import os
import io
import re
import json
import datetime
import requests
import streamlit as st

from dotenv import load_dotenv

import PyPDF2
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
# API — Curriculum Generator
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
# Curriculum Advisor Helpers
# ---------------------------------------------------------------------------

def extract_text_from_upload(uploaded_file) -> str:
    """Extract plain text from PDF, DOCX, or TXT upload."""
    name = uploaded_file.name.lower()
    if name.endswith(".txt"):
        return uploaded_file.read().decode("utf-8", errors="ignore")
    elif name.endswith(".pdf"):
        reader = PyPDF2.PdfReader(uploaded_file)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    elif name.endswith(".docx"):
        doc = Document(uploaded_file)
        return "\n".join(p.text for p in doc.paragraphs)
    else:
        return uploaded_file.read().decode("utf-8", errors="ignore")


def analyze_curriculum(curriculum_text: str) -> str:
    """Ask the LLM to analyze the uploaded curriculum."""
    system = (
        "You are an expert academic curriculum consultant for Woxsen University, "
        "School of Arts and Design, B.Des Communication Design programme. "
        "You review curriculum documents and provide structured, constructive feedback."
    )
    prompt = (
        "Analyze the following curriculum document and provide a structured review with these headings:\n\n"
        "STRENGTHS\n"
        "List 3-4 things that are well-designed in this curriculum.\n\n"
        "AREAS FOR IMPROVEMENT\n"
        "List 3-4 specific improvements with brief explanations.\n\n"
        "MISSING OR OUTDATED TOPICS\n"
        "Identify any topics that are missing or outdated given current industry trends.\n\n"
        "SESSION AND EVALUATION HEALTH\n"
        "Comment on session count, pedagogy variety, and evaluation balance.\n\n"
        "QUICK RECOMMENDATIONS\n"
        "3 actionable next steps for the faculty.\n\n"
        "Keep each point concise (1-2 sentences). No markdown symbols.\n\n"
        f"CURRICULUM:\n{curriculum_text[:6000]}"
    )
    response = requests.post(
        API_ENDPOINT,
        headers={"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"},
        json={
            "model": API_MODEL,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user",   "content": prompt},
            ],
            "temperature": 0.4,
            "max_tokens": 1200,
        },
        timeout=120,
    )
    if response.status_code != 200:
        raise Exception(f"Analysis failed ({response.status_code}): {response.text}")
    return response.json()["choices"][0]["message"]["content"]


# NEW FEATURE: Curriculum Health Score
def score_curriculum(curriculum_text: str) -> dict:
    """Return a quick numeric health score across 5 dimensions."""
    system = (
        "You are an academic curriculum quality assessor. Score the curriculum on five dimensions. "
        "Respond ONLY with a valid JSON object — no preamble, no commentary, no markdown. "
        "Example format: "
        '{"session_count": 72, "pedagogy_variety": 65, "evaluation_balance": 80, '
        '"clo_coverage": 70, "industry_relevance": 60, "overall": 69, '
        '"session_count_note": "Only 12 sessions found, minimum is 16.", '
        '"pedagogy_variety_note": "Good mix of lectures and workshops.", '
        '"evaluation_balance_note": "Internal/End-term split looks correct.", '
        '"clo_coverage_note": "All CLOs are mapped to sessions.", '
        '"industry_relevance_note": "Topics feel current but AI tools are missing."}'
    )
    prompt = (
        "Score this curriculum (0-100 each):\n"
        "- session_count: Are there at least 16 sessions?\n"
        "- pedagogy_variety: Is there a mix of pedagogy types?\n"
        "- evaluation_balance: Are internal (60) and end-term (40) marks correct?\n"
        "- clo_coverage: Are CLOs mapped across sessions and evaluation?\n"
        "- industry_relevance: Does the content reflect current industry practices?\n"
        "Also compute 'overall' as the average.\n"
        "Include a one-sentence '_note' for each dimension.\n\n"
        f"CURRICULUM:\n{curriculum_text[:5000]}"
    )
    response = requests.post(
        API_ENDPOINT,
        headers={"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"},
        json={
            "model": API_MODEL,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user",   "content": prompt},
            ],
            "temperature": 0.1,
            "max_tokens": 500,
        },
        timeout=60,
    )
    if response.status_code != 200:
        raise Exception(f"Scoring failed: {response.text}")
    raw = response.json()["choices"][0]["message"]["content"]
    # Strip any accidental markdown fences
    raw = re.sub(r"```[a-z]*", "", raw).replace("```", "").strip()
    return json.loads(raw)


def chat_with_advisor(curriculum_text: str, chat_history: list, user_question: str) -> str:
    """Conversational advisor with full memory. Chat history is the full conversation so far."""
    system = (
        "You are an expert AI Curriculum Advisor for Woxsen University, School of Arts and Design, "
        "B.Des Communication Design programme. You are reviewing a specific curriculum document with the professor.\n\n"
        "Rules:\n"
        "- You have full memory of this conversation. Always refer back to what was discussed earlier.\n"
        "- When the professor says things like 'top 3 of those', 'change those', 'update that' — "
        "you understand they refer to what was just discussed.\n"
        "- When you suggest a change to the curriculum, be specific (e.g. which session, which module, which reference).\n"
        "- Be concise and practical. No markdown symbols like ** or ###.\n\n"
        f"CURRICULUM DOCUMENT:\n{curriculum_text[:5000]}"
    )
    messages = [{"role": "system", "content": system}]
    for h in chat_history:
        messages.append({"role": h["role"], "content": h["content"]})
    messages.append({"role": "user", "content": user_question})

    response = requests.post(
        API_ENDPOINT,
        headers={"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"},
        json={
            "model": API_MODEL,
            "messages": messages,
            "temperature": 0.4,
            "max_tokens": 1000,
        },
        timeout=120,
    )
    if response.status_code != 200:
        raise Exception(f"Chat failed ({response.status_code}): {response.text}")
    return response.json()["choices"][0]["message"]["content"]


def extract_proposed_changes(curriculum_text: str, chat_history: list) -> str:
    """
    Scan the full conversation and produce a numbered list of proposed changes.
    FIX: Validates chat_history is non-empty before calling. Returns empty string
    (not an exception) when no changes were discussed.
    """
    if not chat_history:
        return ""

    conversation_text = "\n".join(
        f"{'Professor' if h['role'] == 'user' else 'Advisor'}: {h['content']}"
        for h in chat_history
    )
    system = (
        "You are an academic curriculum editor. Read the conversation between a professor and an AI advisor, "
        "then extract every specific curriculum change that was discussed or suggested. "
        "Output ONLY a numbered list. Each item must be one concrete change, written as an action. "
        "Example: '1. Replace Session 4 topic with Generative AI Tools overview.'\n"
        "If no specific changes were discussed, output exactly: NO_CHANGES_FOUND\n"
        "No preamble, no commentary, no markdown. Just the numbered list."
    )
    prompt = (
        f"CURRICULUM:\n{curriculum_text[:3000]}\n\n"
        f"CONVERSATION:\n{conversation_text[:4000]}\n\n"
        "List all proposed changes:"
    )
    response = requests.post(
        API_ENDPOINT,
        headers={"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"},
        json={
            "model": API_MODEL,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user",   "content": prompt},
            ],
            "temperature": 0.2,
            "max_tokens": 800,
        },
        timeout=120,
    )
    if response.status_code != 200:
        raise Exception(f"Change extraction failed ({response.status_code}): {response.text}")
    result = response.json()["choices"][0]["message"]["content"].strip()
    if result == "NO_CHANGES_FOUND":
        return ""
    return result


def apply_changes_to_curriculum(curriculum_text: str, confirmed_changes: list) -> str:
    """
    Apply confirmed changes to curriculum and return output in the exact Woxsen
    12-section pipe-delimited format so it can be fed directly into build_docx()
    and build_pdf() — producing the full styled Woxsen document, not plain text.
    """
    changes_str = "\n".join(f"{i+1}. {c}" for i, c in enumerate(confirmed_changes))

    system = (
        "You are an academic curriculum editor for Woxsen University, School of Arts and Design, "
        "B.Des Communication Design programme.\n\n"
        "You will receive an existing curriculum and a list of confirmed changes to apply.\n"
        "Apply every change precisely. Output the COMPLETE updated curriculum using EXACTLY these "
        "12 section headers (copy verbatim):\n\n"
        "SECTION 1: COURSE OVERVIEW\n"
        "SECTION 2: BRIEF DESCRIPTION\n"
        "SECTION 3: PROGRAM LEARNING OUTCOMES\n"
        "SECTION 4: COURSE LEARNING OUTCOMES\n"
        "SECTION 5: REFERENCES\n"
        "SECTION 6: SESSION-WISE TOPICS\n"
        "SECTION 7: EVALUATION COMPONENTS\n"
        "SECTION 8: ASSESSMENT POLICY\n"
        "SECTION 9: EVALUATION CLO ALIGNMENT\n"
        "SECTION 10: RUBRICS\n"
        "SECTION 11: APPENDICES\n"
        "SECTION 12: ATTENDANCE AND ETHICS\n\n"
        "STRICT FORMATTING RULES — follow exactly:\n"
        "SECTION 1: one field per line, e.g. 'Course Title: X'\n"
        "SECTION 2: 2-3 plain paragraphs, no bullets\n"
        "SECTION 3: pipe-separated rows: PLO1: Name | Definition | Yes\n"
        "SECTION 4: pipe-separated rows: CLO1: Statement | PLO1, PLO2\n"
        "SECTION 5: pipe-separated rows: R1 | Book Title | Author, Year | CLO1\n"
        "SECTION 6: pipe-separated rows: Session 1 | Topic | Learning Outcome | Pedagogy | CLO | R1\n"
        "           MINIMUM 16 sessions. Session 16 or 17 = Final Jury (Pedagogy: Jury)\n"
        "SECTION 7: pipe-separated rows: Component Name | Sessions X-Y | Z marks | Appendix X | CLO1\n"
        "           Internal total = 60 marks, End Term = 40 marks, Grand total = 100 marks\n"
        "SECTION 8: pipe-separated rows: Sessions | Marks | Form | No AI | Scale 1 | CLO | Outcome\n"
        "SECTION 9: 2-3 narrative paragraphs then pipe rows: Component | CLO | Description\n"
        "SECTION 10: pipe rows: CLO1 | Exceeds >80 | Meets 55-79 | Below <55\n"
        "SECTION 11: 4 appendices (A, B, C, D) with Objective, Brief Description, "
        "Submission Requirements (bullet list), then pipe rows: Criteria | Marks | Description\n"
        "SECTION 12: use the exact standard Woxsen attendance/ethics text verbatim\n\n"
        "OUTPUT RULES:\n"
        "- No markdown (no **, ###, ---)\n"
        "- No commentary or preamble outside the 12 sections\n"
        "- All pipe rows must have consistent column counts\n"
        "- Preserve all unchanged content exactly; only modify what the confirmed changes specify"
    )

    prompt = (
        f"ORIGINAL CURRICULUM:\n{curriculum_text[:6000]}\n\n"
        f"CONFIRMED CHANGES TO APPLY:\n{changes_str}\n\n"
        "Output the complete updated curriculum in the exact 12-section Woxsen format:"
    )

    response = requests.post(
        API_ENDPOINT,
        headers={"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"},
        json={
            "model": API_MODEL,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user",   "content": prompt},
            ],
            "temperature": 0.2,
            "max_tokens": 4000,
        },
        timeout=180,
    )
    if response.status_code != 200:
        raise Exception(f"Apply changes failed ({response.status_code}): {response.text}")
    return response.json()["choices"][0]["message"]["content"]


def build_updated_docx_full(updated_text: str, semester: str = "IV", credits: int = 3,
                             faculty_name: str = "", faculty_email: str = "",
                             course_year: str = "2025-26") -> bytes:
    """
    Build a fully styled Woxsen DOCX from updated curriculum text.
    Routes through the same build_docx() pipeline used by the generator —
    so the output is identical in format to a freshly generated curriculum.
    """
    return build_docx(updated_text, semester, credits, faculty_name, faculty_email, course_year)


def build_updated_pdf_full(updated_text: str, semester: str = "IV", credits: int = 3,
                            faculty_name: str = "", faculty_email: str = "",
                            course_year: str = "2025-26") -> bytes:
    """
    Build a fully styled Woxsen PDF from updated curriculum text.
    Routes through the same build_pdf() pipeline used by the generator.
    """
    return build_pdf(updated_text, semester, credits, faculty_name, faculty_email, course_year)


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

    co = doc.add_paragraph()
    co.add_run("Course Outline").font.size = Pt(10)
    co.alignment = WD_ALIGN_PARAGRAPH.CENTER
    co.paragraph_format.space_before = Pt(2)
    co.paragraph_format.space_after  = Pt(14)

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

    add_section_header(doc, "Brief Description and Relevance of the Course")
    for line in secs.get("SECTION 2: BRIEF DESCRIPTION", "").split("\n"):
        if line.strip():
            add_body_text(doc, line.strip())

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

    add_section_header(doc, "Prerequisites")
    add_body_text(doc, "NIL")

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

    add_section_header(doc, "Session-Wise Topics and Reading / References")
    sess_rows = pipe_rows(secs.get("SECTION 6: SESSION-WISE TOPICS", ""), min_cols=3)
    if sess_rows:
        norm = pad_rows(sess_rows, 6)
        add_styled_table(doc,
            ["Sn.", "Topic", "Session Intended Learning Outcome", "Pedagogy", "CLO", "Reading Material"],
            norm, [1.0, 4.5, 5.5, 2.5, 1.0, 2.0])

    add_section_header(doc, "Performance Evaluation Components for the Course")
    eval_rows = pipe_rows(secs.get("SECTION 7: EVALUATION COMPONENTS", ""), min_cols=3)
    if eval_rows:
        add_styled_table(doc, ["Session No.", "Marks", "Evaluation Form", "CLO"],
                         pad_rows(eval_rows, 4), [3, 2, 8, 3])

    add_section_header(doc, "Assessment Policy")
    ap_rows = pipe_rows(secs.get("SECTION 8: ASSESSMENT POLICY", ""), min_cols=5)
    if ap_rows:
        add_styled_table(doc,
            ["Session No.", "Marks", "Evaluation Form", "Assessment Level", "Scale", "CLO", "Outcome Measured"],
            pad_rows(ap_rows, 7), [2.0, 1.5, 3.5, 2.5, 1.5, 2.0, 3.5])

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

    add_section_header(doc, "Evaluation Rubrics")
    rubric_rows = pipe_rows(secs.get("SECTION 10: RUBRICS", ""), min_cols=4)
    if rubric_rows:
        add_styled_table(doc,
            ["CLO", "Exceeds Expectations (>80)", "Meets Expectations (55–79)", "Does Not Meet Expectations (<55)"],
            pad_rows(rubric_rows, 4), [2, 5, 5, 4])

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
        canvas.setFillColor(RL_FOOTER)
        canvas.rect(0, 0, doc.pagesize[0], 18, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica", 9)
        canvas.drawRightString(doc.pagesize[0] - 15, 5, str(canvas.getPageNumber()))
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

    section_header("Brief Description and Relevance of the Course")
    for line in secs.get("SECTION 2: BRIEF DESCRIPTION", "").split("\n"):
        if line.strip():
            story.append(Paragraph(line.strip(), s_normal))

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

    section_header("Prerequisites")
    story.append(Paragraph("NIL", s_normal))

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

    section_header("Session-Wise Topics and Reading / References")
    sess_rows = pipe_rows(secs.get("SECTION 6: SESSION-WISE TOPICS", ""), min_cols=3)
    if sess_rows:
        styled_table(
            ["Sn.", "Topic", "Session Intended Learning Outcome", "Pedagogy", "CLO", "Reading Material"],
            pad_rows(sess_rows, 6),
            [1*cm, 3.8*cm, 5*cm, 2.5*cm, 1.2*cm, 2.5*cm])

    section_header("Performance Evaluation Components for the Course")
    eval_rows = pipe_rows(secs.get("SECTION 7: EVALUATION COMPONENTS", ""), min_cols=3)
    if eval_rows:
        styled_table(["Session No.", "Marks", "Evaluation Form", "CLO"],
                     pad_rows(eval_rows, 4), [3*cm, 2*cm, 8*cm, 3*cm])

    section_header("Assessment Policy")
    ap_rows = pipe_rows(secs.get("SECTION 8: ASSESSMENT POLICY", ""), min_cols=5)
    if ap_rows:
        styled_table(
            ["Session No.", "Marks", "Evaluation Form", "Assessment Level", "Scale", "CLO", "Outcome Measured"],
            pad_rows(ap_rows, 7),
            [1.8*cm, 1.3*cm, 3*cm, 2.5*cm, 1.5*cm, 2*cm, 3.9*cm])

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

    section_header("Evaluation Rubrics")
    rubric_rows = pipe_rows(secs.get("SECTION 10: RUBRICS", ""), min_cols=4)
    if rubric_rows:
        styled_table(
            ["CLO", "Exceeds Expectations (>80)", "Meets Expectations (55–79)", "Does Not Meet Expectations (<55)"],
            pad_rows(rubric_rows, 4),
            [1.5*cm, 5*cm, 5*cm, 4.5*cm])

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
# Streamlit UI
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Woxsen AI Curriculum Assistant",
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
.score-card {
    background: white;
    border-radius: 10px;
    padding: 12px 16px;
    margin: 4px 0;
    border-left: 4px solid #C00000;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
}
</style>
""", unsafe_allow_html=True)

# Header
col_logo, col_title = st.columns([1, 5])
with col_logo:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=90)
with col_title:
    st.markdown("## Woxsen University — AI Curriculum Assistant")
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

    # NEW: Change History in sidebar
    if st.session_state.get("change_history"):
        st.markdown("---")
        st.markdown("### 📋 Change History")
        for entry in reversed(st.session_state["change_history"][-5:]):
            st.markdown(f"**{entry['timestamp']}**")
            st.caption(f"{entry['count']} change(s) applied")

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
    tab_preview, tab_download, tab_quickedit = st.tabs(["📄 Preview", "⬇️ Download", "✏️ Quick Edit"])

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

    # NEW FEATURE: Quick Edit tab
    with tab_quickedit:
        st.markdown("### ✏️ Quick Edit Generated Curriculum")
        st.caption("Edit the raw curriculum text directly and save your changes.")
        edited_raw = st.text_area(
            "Edit curriculum",
            value=st.session_state.get("raw", ""),
            height=500,
            label_visibility="collapsed",
            key="quick_edit_area",
        )
        if st.button("💾 Save Edits", type="primary"):
            st.session_state["raw"] = edited_raw
            st.success("✅ Changes saved. Go to the Download tab to export.")


# ===========================================================================
# AI CURRICULUM ADVISOR
# ===========================================================================

st.divider()
st.markdown("## 🧠 AI Curriculum Advisor")
st.markdown(
    "<span style='color:#888;font-size:0.9em'>"
    "Upload an existing curriculum · Get a health score · Chat with full memory · Review & apply changes"
    "</span>",
    unsafe_allow_html=True,
)

advisor_tab1, advisor_tab2, advisor_tab3 = st.tabs(
    ["📂 Upload & Analyze", "💬 Chat Advisor", "✏️ Review & Apply Changes"]
)

# ---------------------------------------------------------------------------
# TAB 1 — Upload & Analyze
# ---------------------------------------------------------------------------
with advisor_tab1:
    st.markdown("### Upload Existing Curriculum")
    uploaded_file = st.file_uploader(
        "Upload your curriculum document",
        type=["pdf", "docx", "txt"],
        key="curriculum_upload",
        help="Supports PDF, Word (.docx), and plain text files",
    )

    if uploaded_file:
        # KEY FIX: Only re-read and reset state when a NEW file is uploaded.
        # Without this guard, every st.rerun() (e.g. after clicking "Review Proposed Changes")
        # re-enters this block because the file_uploader widget still holds the file,
        # wiping proposed_changes_raw and making Tab 3 blank every time.
        current_file_id = f"{uploaded_file.name}_{uploaded_file.size}"
        if st.session_state.get("_loaded_file_id") != current_file_id:
            with st.spinner("Reading document..."):
                try:
                    curriculum_text = extract_text_from_upload(uploaded_file)
                    st.session_state["uploaded_curriculum"] = curriculum_text
                    st.session_state["_loaded_file_id"]     = current_file_id
                    st.session_state["advisor_chat"]        = []
                    st.session_state.pop("proposed_changes_raw",  None)
                    st.session_state.pop("confirmed_changes",     None)
                    st.session_state.pop("updated_curriculum",    None)
                    st.session_state.pop("curriculum_score",      None)
                    st.success(f"✅ Document loaded — {len(curriculum_text.split())} words extracted.")
                except Exception as e:
                    st.error(f"Could not read file: {e}")

    if "uploaded_curriculum" in st.session_state:
        with st.expander("📄 View extracted text", expanded=False):
            st.text_area("Curriculum Text", value=st.session_state["uploaded_curriculum"],
                         height=250, label_visibility="collapsed")

        # NEW FEATURE: Health Score
        st.markdown("### 📊 Curriculum Health Score")
        score_col1, score_col2 = st.columns([2, 1])
        with score_col1:
            if st.button("⚡ Generate Health Score", type="primary"):
                with st.spinner("Scoring curriculum… ~10 seconds"):
                    try:
                        score_data = score_curriculum(st.session_state["uploaded_curriculum"])
                        st.session_state["curriculum_score"] = score_data
                    except Exception as e:
                        st.error(f"Scoring failed: {e}")

        if "curriculum_score" in st.session_state:
            sc = st.session_state["curriculum_score"]
            dims = [
                ("session_count",       "📅 Session Count"),
                ("pedagogy_variety",    "🎓 Pedagogy Variety"),
                ("evaluation_balance",  "⚖️ Evaluation Balance"),
                ("clo_coverage",        "🎯 CLO Coverage"),
                ("industry_relevance",  "🔥 Industry Relevance"),
            ]
            overall = sc.get("overall", 0)
            color = "#2ecc71" if overall >= 75 else "#f39c12" if overall >= 50 else "#C00000"
            st.markdown(
                f"<div style='font-size:1.6em;font-weight:700;color:{color};margin-bottom:8px'>"
                f"Overall Score: {overall}/100</div>",
                unsafe_allow_html=True,
            )
            for key, label in dims:
                val  = sc.get(key, 0)
                note = sc.get(f"{key}_note", "")
                bar_color = "#2ecc71" if val >= 75 else "#f39c12" if val >= 50 else "#C00000"
                st.markdown(
                    f"<div class='score-card'><b>{label}: {val}/100</b><br>"
                    f"<span style='color:#555;font-size:0.88em'>{note}</span>"
                    f"<div style='background:#eee;border-radius:4px;margin-top:6px;height:8px'>"
                    f"<div style='background:{bar_color};width:{val}%;height:8px;border-radius:4px'></div>"
                    f"</div></div>",
                    unsafe_allow_html=True,
                )

        st.markdown("---")
        st.markdown("### AI Analysis")
        if st.button("🔍 Analyze Curriculum", use_container_width=False, type="primary"):
            with st.spinner("Analyzing… this takes ~15 seconds"):
                try:
                    analysis = analyze_curriculum(st.session_state["uploaded_curriculum"])
                    st.session_state["curriculum_analysis"] = analysis
                except Exception as e:
                    st.error(f"Analysis failed: {e}")

        if "curriculum_analysis" in st.session_state:
            st.markdown("---")
            analysis_text   = st.session_state["curriculum_analysis"]
            section_headers = ["STRENGTHS", "AREAS FOR IMPROVEMENT",
                                "MISSING OR OUTDATED TOPICS",
                                "SESSION AND EVALUATION HEALTH", "QUICK RECOMMENDATIONS"]
            icons = {
                "STRENGTHS":                    "✅",
                "AREAS FOR IMPROVEMENT":        "⚠️",
                "MISSING OR OUTDATED TOPICS":   "🔍",
                "SESSION AND EVALUATION HEALTH":"📊",
                "QUICK RECOMMENDATIONS":        "🚀",
            }
            lines = analysis_text.split("\n")
            current_section  = None
            section_content  = {}
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                matched = next((h for h in section_headers if line.upper().startswith(h)), None)
                if matched:
                    current_section = matched
                    section_content[current_section] = []
                elif current_section:
                    section_content[current_section].append(line)

            if section_content:
                for section in section_headers:
                    if section in section_content:
                        st.markdown(f"**{icons.get(section,'•')} {section.title()}**")
                        for item in section_content[section]:
                            if item:
                                st.markdown(f"- {item}")
                        st.markdown("")
            else:
                st.text_area("Analysis Result", value=analysis_text, height=400,
                             label_visibility="collapsed")

# ---------------------------------------------------------------------------
# TAB 2 — Chat Advisor (full memory)
# ---------------------------------------------------------------------------
with advisor_tab2:
    st.markdown("### Chat with Your Curriculum Advisor")
    st.caption("The advisor remembers everything discussed in this session. Ask follow-ups freely.")

    if "uploaded_curriculum" not in st.session_state:
        st.info("👆 Please upload a curriculum document in the **Upload & Analyze** tab first.")
    else:
        if "advisor_chat" not in st.session_state:
            st.session_state["advisor_chat"] = []

        # NEW FEATURE: Save / Load chat session
        with st.expander("💾 Save / Load Chat Session", expanded=False):
            save_col, load_col = st.columns(2)
            with save_col:
                if st.session_state.get("advisor_chat"):
                    chat_json = json.dumps(st.session_state["advisor_chat"], indent=2)
                    st.download_button(
                        "⬇️ Save chat as JSON",
                        data=chat_json.encode("utf-8"),
                        file_name=f"advisor_chat_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.json",
                        mime="application/json",
                        use_container_width=True,
                    )
                else:
                    st.info("No chat to save yet.")
            with load_col:
                uploaded_chat = st.file_uploader("Load previous chat JSON", type=["json"],
                                                  key="load_chat_upload")
                if uploaded_chat:
                    try:
                        loaded = json.loads(uploaded_chat.read().decode("utf-8"))
                        if isinstance(loaded, list):
                            st.session_state["advisor_chat"] = loaded
                            st.success(f"✅ Loaded {len(loaded)} messages.")
                            st.rerun()
                        else:
                            st.error("Invalid chat file format.")
                    except Exception as e:
                        st.error(f"Could not load chat: {e}")

        # Suggestion buttons
        st.markdown("**Quick questions:**")
        suggestions = [
            "Which modules are outdated?",
            "Give me top 3 improvements",
            "How can I add AI tools to this course?",
            "Suggest more practical activities",
        ]
        s_cols = st.columns(4)
        for i, s in enumerate(suggestions):
            with s_cols[i]:
                if st.button(s, key=f"sug_{i}", use_container_width=True):
                    st.session_state["advisor_prefill"] = s

        # NEW FEATURE: Suggested change shortcuts
        st.markdown("**Suggested changes (click to propose):**")
        change_suggestions = [
            "Add a session on Generative AI tools",
            "Update references to 2023-2024 editions",
            "Add more hands-on workshop sessions",
            "Include industry guest lecture session",
            "Add UX research methodology session",
        ]
        cs_cols = st.columns(5)
        for i, cs in enumerate(change_suggestions):
            with cs_cols[i]:
                if st.button(cs, key=f"cs_{i}", use_container_width=True,
                             help="Click to propose this change in chat"):
                    st.session_state["advisor_prefill"] = f"Please propose the following change to the curriculum: {cs}"

        st.markdown("---")

        # Render full chat history
        for msg in st.session_state["advisor_chat"]:
            with st.chat_message(msg["role"]):
                st.write(msg["content"])

        # Handle prefill from suggestion buttons
        prefill    = st.session_state.pop("advisor_prefill", "") if "advisor_prefill" in st.session_state else ""
        user_input = st.chat_input("Ask anything about your curriculum…", key="advisor_chat_input")
        if prefill and not user_input:
            user_input = prefill

        if user_input:
            st.session_state["advisor_chat"].append({"role": "user", "content": user_input})
            with st.chat_message("user"):
                st.write(user_input)

            with st.chat_message("assistant"):
                with st.spinner("Thinking…"):
                    try:
                        reply = chat_with_advisor(
                            st.session_state["uploaded_curriculum"],
                            st.session_state["advisor_chat"][:-1],
                            user_input,
                        )
                        st.write(reply)
                        st.session_state["advisor_chat"].append({"role": "assistant", "content": reply})
                    except Exception as e:
                        err = f"Error: {e}"
                        st.error(err)
                        st.session_state["advisor_chat"].append({"role": "assistant", "content": err})

        # Bottom controls
        if st.session_state.get("advisor_chat"):
            bot_col1, bot_col2 = st.columns([1, 1])
            with bot_col1:
                if st.button("🗑️ Clear Chat", key="clear_chat_btn"):
                    st.session_state["advisor_chat"] = []
                    st.session_state.pop("proposed_changes_raw", None)
                    st.session_state.pop("confirmed_changes",    None)
                    st.session_state.pop("updated_curriculum",   None)
                    st.rerun()

            with bot_col2:
                # FIX: Check that chat contains at least one assistant message before allowing extract
                chat_has_content = any(
                    m["role"] == "assistant"
                    for m in st.session_state.get("advisor_chat", [])
                )
                if not chat_has_content:
                    st.info("💬 Chat with the advisor first to generate proposals.")
                else:
                    if st.button("📋 Review Proposed Changes →", key="goto_changes_btn", type="primary"):
                        with st.spinner("Scanning conversation for proposed changes…"):
                            try:
                                raw_changes = extract_proposed_changes(
                                    st.session_state["uploaded_curriculum"],
                                    st.session_state["advisor_chat"],
                                )
                                # Parse numbered list into Python list
                                items = []
                                if raw_changes:
                                    for line in raw_changes.split("\n"):
                                        line = line.strip()
                                        if not line:
                                            continue
                                        clean = re.sub(r"^\d+[\.\)]\s*", "", line).strip()
                                        if clean:
                                            items.append(clean)

                                # Write to session state. Do NOT call st.rerun() here —
                                # the extra rerun re-triggers the file_uploader block which
                                # wipes these keys. Streamlit's natural button rerun is enough.
                                st.session_state["proposed_changes_raw"] = items
                                st.session_state["confirmed_changes"]    = items.copy()
                                st.session_state.pop("updated_curriculum", None)

                                if items:
                                    st.success(f"✅ {len(items)} change(s) extracted! Now click the **✏️ Review & Apply Changes** tab above.")
                                else:
                                    st.warning("No specific changes found in the conversation. Try discussing more concrete changes, then click this button again.")

                            except Exception as e:
                                st.session_state["proposed_changes_raw"] = []
                                st.session_state["confirmed_changes"]    = []
                                st.error(f"Extraction failed: {e}")

# ---------------------------------------------------------------------------
# TAB 3 — Review & Apply Changes  (THE FIXED TAB)
# ---------------------------------------------------------------------------
with advisor_tab3:
    st.markdown("### Review Proposed Changes")

    # FIX: Surface any error that was stored during the last extract attempt
    # Using pop() so it only shows once
    if "_extract_error" in st.session_state:
        st.error(f"Could not extract changes: {st.session_state.pop('_extract_error')}")

    if "uploaded_curriculum" not in st.session_state:
        st.info("👆 Upload a curriculum and chat with the advisor first.")

    elif "proposed_changes_raw" not in st.session_state:
        # FIX: Clear instruction so user knows exactly what to do
        st.info("💬 Go to the **Chat Advisor** tab, discuss changes with the advisor, then click **Review Proposed Changes →**.")

    else:
        proposed = st.session_state["proposed_changes_raw"]

        if not proposed:
            st.warning(
                "No specific changes were found in the conversation. "
                "Try discussing concrete changes — for example: "
                "'Replace Session 4 with a Generative AI tools overview' or "
                "'Add a session on UX research methods after Session 8.'"
            )
            # Still let them add manual changes
            st.markdown("---")
            st.markdown("**You can still add custom changes manually:**")

        else:
            st.markdown(
                f"**{len(proposed)} change(s) identified** from your conversation. "
                "Edit any item, uncheck ones you don't want, then click **Apply Changes**."
            )
            st.markdown("---")

        # Editable checklist — sync from proposed_changes_raw
        confirmed_changes = []
        for i, change in enumerate(st.session_state.get("confirmed_changes", proposed)):
            col_chk, col_txt = st.columns([0.05, 0.95])
            with col_chk:
                keep = st.checkbox("", value=True, key=f"chk_{i}", label_visibility="collapsed")
            with col_txt:
                edited = st.text_input(f"change_{i}", value=change,
                                       key=f"edit_{i}", label_visibility="collapsed")
            if keep:
                confirmed_changes.append(edited)

        # Add custom change
        st.markdown("---")
        st.markdown("**Add a custom change:**")
        add_col1, add_col2 = st.columns([4, 1])
        with add_col1:
            new_change = st.text_input(
                "Custom change",
                placeholder="e.g. Add a session on Prompt Engineering after Session 10",
                key="new_change_input",
                label_visibility="collapsed",
            )
        with add_col2:
            if st.button("➕ Add", key="add_change_btn", use_container_width=True):
                if new_change.strip():
                    updated_list = st.session_state["confirmed_changes"] + [new_change.strip()]
                    st.session_state["confirmed_changes"] = updated_list
                    st.rerun()

        st.markdown("---")

        # Summary before applying
        if confirmed_changes:
            st.markdown(f"**{len(confirmed_changes)} change(s) will be applied:**")
            for i, c in enumerate(confirmed_changes, 1):
                st.markdown(f"{i}. {c}")

            st.markdown("")
            if st.button("✅ Apply Changes to Curriculum", type="primary", use_container_width=False):
                with st.spinner("Applying changes and rebuilding in Woxsen format… 20-40 seconds"):
                    try:
                        updated = apply_changes_to_curriculum(
                            st.session_state["uploaded_curriculum"],
                            confirmed_changes,
                        )
                        st.session_state["updated_curriculum"] = updated

                        # Log to change history
                        if "change_history" not in st.session_state:
                            st.session_state["change_history"] = []
                        st.session_state["change_history"].append({
                            "timestamp": datetime.datetime.now().strftime("%d %b %Y, %H:%M"),
                            "count":     len(confirmed_changes),
                            "changes":   confirmed_changes.copy(),
                        })

                        st.success("✅ Curriculum updated! Download below in full Woxsen format.")
                    except Exception as e:
                        st.error(f"Failed to apply changes: {e}")
        else:
            st.warning("No changes selected. Check at least one item above.")

        # Show updated curriculum + download in full Woxsen format
        if "updated_curriculum" in st.session_state:
            st.divider()
            st.markdown("### Updated Curriculum — Download")
            st.caption("Both downloads use the full Woxsen format: cover page, styled tables, red headers, logo, appendices — identical to the generator output.")

            # Pull faculty/course metadata from sidebar inputs or session state
            _semester    = st.session_state.get("semester",     semester)
            _credits     = st.session_state.get("credits",      credits)
            _fname       = st.session_state.get("faculty_name", faculty_name)
            _femail      = st.session_state.get("faculty_email",faculty_email)
            _year        = st.session_state.get("course_year",  course_year)

            with st.expander("📄 Preview raw updated text", expanded=False):
                st.text_area("Updated", value=st.session_state["updated_curriculum"],
                             height=350, label_visibility="collapsed")

            dl_col1, dl_col2, dl_col3 = st.columns(3)

            with dl_col1:
                if st.button("📝 Build Woxsen DOCX", use_container_width=True, type="primary"):
                    with st.spinner("Building full Woxsen Word document…"):
                        try:
                            docx_data = build_updated_docx_full(
                                st.session_state["updated_curriculum"],
                                _semester, _credits, _fname, _femail, _year,
                            )
                            st.download_button(
                                "⬇️ Download Updated .docx",
                                data=docx_data,
                                file_name="Woxsen_Updated_Curriculum.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True,
                            )
                        except Exception as e:
                            st.error(f"DOCX build failed: {e}")

            with dl_col2:
                if st.button("📄 Build Woxsen PDF", use_container_width=True, type="primary"):
                    with st.spinner("Building full Woxsen PDF…"):
                        try:
                            pdf_data = build_updated_pdf_full(
                                st.session_state["updated_curriculum"],
                                _semester, _credits, _fname, _femail, _year,
                            )
                            st.download_button(
                                "⬇️ Download Updated .pdf",
                                data=pdf_data,
                                file_name="Woxsen_Updated_Curriculum.pdf",
                                mime="application/pdf",
                                use_container_width=True,
                            )
                        except Exception as e:
                            st.error(f"PDF build failed: {e}")

            with dl_col3:
                st.download_button(
                    "⬇️ Download raw .txt",
                    data=st.session_state["updated_curriculum"].encode("utf-8"),
                    file_name="Woxsen_Updated_Curriculum.txt",
                    mime="text/plain",
                    use_container_width=True,
                    help="Plain text — useful if you want to paste into the generator for a full fresh rebuild",
                )

            # NEW FEATURE: Show full change history for this session
            if st.session_state.get("change_history"):
                st.markdown("---")
                with st.expander("📋 Full Change History (this session)", expanded=False):
                    for idx, entry in enumerate(reversed(st.session_state["change_history"]), 1):
                        st.markdown(f"**Batch {idx} — {entry['timestamp']}** ({entry['count']} changes)")
                        for c in entry["changes"]:
                            st.markdown(f"  - {c}")