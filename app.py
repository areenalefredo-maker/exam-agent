"""
Exam Worksheet Generator
A Streamlit app that takes Question Paper (QP), Mark Scheme (MS), and a classification
spreadsheet, and produces a formatted DOCX worksheet using Claude AI.
"""

import streamlit as st
import anthropic
import base64
import json
import io
import re
import time
from pathlib import Path

import pandas as pd
import fitz  # PyMuPDF
from PIL import Image

from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------------------------- Page Setup ----------------------------
st.set_page_config(
    page_title="Exam Worksheet Generator",
    page_icon="📝",
    layout="centered"
)

st.title("📝 Exam Worksheet Generator")
st.caption("Upload your exam files → AI processes them → Download a formatted worksheet")

# ---------------------------- Helpers ----------------------------
MODEL = "claude-sonnet-4-5"

def call_claude_with_pdf(client, pdf_bytes, prompt, max_tokens=8000):
    """Send a PDF to Claude with a prompt and return the text response."""
    response = client.messages.create(
        model=MODEL,
        max_tokens=max_tokens,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": base64.b64encode(pdf_bytes).decode()
                    }
                },
                {"type": "text", "text": prompt}
            ]
        }]
    )
    return response.content[0].text

def call_claude_with_image(client, img_bytes, prompt, max_tokens=1000):
    """Send an image to Claude with a prompt."""
    response = client.messages.create(
        model=MODEL,
        max_tokens=max_tokens,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/png",
                        "data": base64.b64encode(img_bytes).decode()
                    }
                },
                {"type": "text", "text": prompt}
            ]
        }]
    )
    return response.content[0].text

def parse_json_from_text(text):
    """Extract JSON from Claude's response, handling code fences if present."""
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return json.loads(text)

# ---------------------------- Extraction Functions ----------------------------

EXTRACTION_PROMPT = """You are extracting questions from an exam paper to recreate them in a worksheet.

Return a JSON object with this exact structure (and NOTHING else):
{
  "questions": [
    {
      "number": 1,
      "page": 2,
      "body": "the question text including any introductory paragraph",
      "options": ["A.  ...", "B.  ...", "C.  ...", "D.  ..."],
      "has_figure": true,
      "figure_position": "after_options",
      "figure_description": "brief description of the figure for image extraction",
      "shared_with": []
    }
  ]
}

Rules:
- Include EVERY question. Number them as they appear in the paper.
- "page" is the PDF page number (1-indexed) the question appears on.
- "body" should be the main question text. If there's an introductory paragraph that belongs to this question, include it.
- "options" should each start with "A. ", "B. ", "C. ", "D. " (with two spaces after the period for alignment).
- Preserve scientific notation properly: use Unicode subscripts (CO₂, H₂O), Greek letters (α, β), arrows (→), and degree signs (°C).
- "has_figure": true if this question has an associated diagram, graph, chart, or chemical structure.
- "figure_position": one of "before_body", "between_body_and_options", "after_options". This is where the figure appears relative to the question content.
- "figure_description": short description of what's in the figure (e.g., "graph of enzyme rate vs time", "Davson-Danielli membrane diagram").
- "shared_with": if multiple questions share ONE figure (like Q2 and Q3 both reference one image), list the OTHER question numbers here. Only set has_figure=true on the FIRST question that introduces the shared figure.
- For tables that are intrinsic to the question (e.g., a table of options), include them as part of "body" using a simple text format, OR set has_figure=true if the table is image-based.

Return ONLY the JSON object, no preamble, no explanation, no code fences.
"""

ANSWERS_PROMPT = """Extract the correct answer letter for each question from this mark scheme.

Return JSON only with this structure:
{"answers": {"1": "B", "2": "D", "3": "B", ...}}

Use only the letter (A, B, C, or D). If a question has no clear answer, omit it.
Return ONLY the JSON, no preamble.
"""

def extract_questions(client, qp_pdf_bytes, status_callback=None):
    """Extract structured question data from QP."""
    if status_callback:
        status_callback("Asking Claude to read the question paper...")
    text = call_claude_with_pdf(client, qp_pdf_bytes, EXTRACTION_PROMPT, max_tokens=16000)
    data = parse_json_from_text(text)
    return data.get("questions", [])

def extract_answers(client, ms_pdf_bytes, status_callback=None):
    """Extract answer key from Mark Scheme."""
    if status_callback:
        status_callback("Asking Claude to read the mark scheme...")
    text = call_claude_with_pdf(client, ms_pdf_bytes, ANSWERS_PROMPT, max_tokens=2000)
    data = parse_json_from_text(text)
    return data.get("answers", {})

def read_classifications(xlsx_bytes):
    """Read chapter / difficulty / reference from the Excel sheet."""
    df = pd.read_excel(io.BytesIO(xlsx_bytes))
    # Normalize column names (case-insensitive matching)
    cols = {c.lower().strip(): c for c in df.columns}

    def find_col(*candidates):
        for c in candidates:
            if c.lower() in cols:
                return cols[c.lower()]
        return None

    qnum_col = find_col("Question No.", "Question No", "Question", "Q", "Q#", "QNum")
    chapter_col = find_col("Topic", "Chapter", "Subject", "Unit")
    diff_col = find_col("Difficulty", "Level", "Level of question")
    ref_col = find_col("Reference", "Date", "Year")

    classifications = {}
    if qnum_col is None:
        return classifications

    for _, row in df.iterrows():
        try:
            qnum = int(float(str(row[qnum_col]).strip()))
        except (ValueError, TypeError):
            continue
        classifications[qnum] = {
            "chapter": str(row[chapter_col]) if chapter_col else "",
            "difficulty": str(row[diff_col]) if diff_col else "Medium",
            "reference": str(row[ref_col]).strip() if ref_col else "",
        }
    return classifications

# ---------------------------- Image Extraction ----------------------------

def render_page_image(pdf_bytes, page_index, dpi=200):
    """Render a single PDF page as a PIL Image."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[page_index]
    pix = page.get_pixmap(dpi=dpi)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    doc.close()
    return img

FIGURE_BOUNDS_PROMPT = """This is page {page} of an exam paper. Find the bounding box of the figure/diagram/graph/chart for Question {qnum}.

Description of the figure: {description}

Return JSON only:
{{"top_pct": 0.0, "left_pct": 0.0, "bottom_pct": 1.0, "right_pct": 1.0}}

All values are percentages (0.0 to 1.0) of the page dimensions.
- top_pct: y-coordinate of the top of the figure (0 = page top)
- bottom_pct: y-coordinate of the bottom of the figure (1 = page bottom)
- left_pct, right_pct: similar for horizontal extent

Return tight bounds around ONLY the figure, excluding question text, options, and headers.
Return ONLY the JSON, no preamble.
"""

def crop_figure_for_question(client, pdf_bytes, page_index, qnum, description):
    """Use Claude vision to identify and crop the figure for a question."""
    page_img = render_page_image(pdf_bytes, page_index, dpi=200)

    # Send the page image to Claude to get bounds
    buf = io.BytesIO()
    page_img.save(buf, format="PNG")
    bounds_text = call_claude_with_image(
        client,
        buf.getvalue(),
        FIGURE_BOUNDS_PROMPT.format(page=page_index + 1, qnum=qnum, description=description),
        max_tokens=300
    )
    try:
        bounds = parse_json_from_text(bounds_text)
    except Exception:
        # Fallback: return the whole page
        return page_img

    w, h = page_img.size
    left = max(0, int(bounds.get("left_pct", 0.05) * w))
    top = max(0, int(bounds.get("top_pct", 0) * h))
    right = min(w, int(bounds.get("right_pct", 0.95) * w))
    bottom = min(h, int(bounds.get("bottom_pct", 1) * h))

    # Sanity check
    if right <= left or bottom <= top:
        return page_img
    if (right - left) < 100 or (bottom - top) < 50:
        return page_img

    return page_img.crop((left, top, right, bottom))

# ---------------------------- DOCX Building ----------------------------

def add_horizontal_rule(paragraph):
    """Add a bottom border to a paragraph (used as a horizontal rule)."""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "999999")
    pBdr.append(bottom)
    pPr.append(pBdr)

def add_run(paragraph, text, bold=False, size=11, font="Arial"):
    run = paragraph.add_run(text)
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    return run

def add_blank(doc, size=11):
    p = doc.add_paragraph()
    add_run(p, " ", size=size)
    return p

def add_lined_blank(doc):
    """Add a paragraph with a bottom border (for student solution lines)."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "4")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "BFBFBF")
    pBdr.append(bottom)
    pPr.append(pBdr)
    add_run(p, "", size=11)
    return p

def add_question_header(doc, q_num, level, marks, reference, chapter, page_break=False):
    """Add the standard question header block."""
    p = doc.add_paragraph()
    if page_break:
        p.paragraph_format.page_break_before = True
    add_run(p, f"Question: {q_num}", bold=True, size=13)

    p2 = doc.add_paragraph()
    add_run(p2, "Level of question", bold=True, size=11)
    add_run(p2, f": {level}  |  ", size=11)
    add_run(p2, "Number of Marks: ", bold=True, size=11)
    add_run(p2, f"{marks}  |  ", size=11)
    add_run(p2, "Reference:", bold=True, size=11)
    add_run(p2, f" {reference}  |", size=11)

    p3 = doc.add_paragraph()
    add_run(p3, "Chapter", bold=True, size=11)
    add_run(p3, f" :  {chapter}", size=11)

    p4 = doc.add_paragraph()
    add_horizontal_rule(p4)

def add_body_text(doc, text):
    p = doc.add_paragraph()
    add_run(p, text, size=11)

def add_option(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.25)
    add_run(p, text, size=11)

def add_image(doc, pil_image, max_width_inches=5.0):
    """Add a centered image to the document."""
    buf = io.BytesIO()
    pil_image.save(buf, format="PNG")
    buf.seek(0)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    # Calculate height to preserve aspect ratio
    w, h = pil_image.size
    width = Inches(max_width_inches)
    run.add_picture(buf, width=width)

def add_student_solution_section(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    add_run(p, "Student's Solution:", bold=True, size=11)
    for _ in range(5):
        add_lined_blank(doc)

def add_answer_section(doc, answer, motivation):
    p_hr = doc.add_paragraph()
    add_horizontal_rule(p_hr)

    p1 = doc.add_paragraph()
    add_run(p1, "Answer from Mark Scheme:", bold=True, size=11)

    p2 = doc.add_paragraph()
    add_run(p2, answer, bold=True, size=12)

    p_hr2 = doc.add_paragraph()
    add_horizontal_rule(p_hr2)

    p3 = doc.add_paragraph()
    add_run(p3, "Keep it up", bold=True, size=11)
    add_run(p3, f" : {motivation}", size=11)

    p_hr3 = doc.add_paragraph()
    add_horizontal_rule(p_hr3)

MOTIVATIONS = [
    "Keep up the momentum!", "Success is within reach.", "Knowledge is power.",
    "Stay focused and positive.", "You are a fast learner.", "Every challenge is an opportunity.",
    "Wonderful work!", "Stay motivated.", "Brilliant work!", "Keep moving forward.",
    "You are unlimited.", "Great effort leads to success!", "You are a star!",
    "You are doing a fantastic job.", "Keep challenging yourself.", "Keep exploring new ideas.",
    "Superb effort!", "Excellent progress!", "You are making a difference.",
    "You are capable of amazing things.", "You're getting better every day.",
    "Great job staying focused.", "Stay curious and keep learning.", "Aim for the stars.",
    "You've got this!", "Keep shining!", "Keep aiming high.", "Your attitude determines your direction."
]

def build_document(questions, answers, classifications, qp_pdf_bytes, client, status_callback=None, default_reference="", figure_cache=None):
    """Build the final DOCX from extracted data."""
    doc = Document()

    # Page setup: US Letter, 1" margins
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

    # Default style
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)

    # Build a map of which questions share figures
    shared_figures = {}  # q_num -> primary_q_num (the one that has the actual figure)
    for q in questions:
        if q.get("has_figure") and q.get("shared_with"):
            for shared in q["shared_with"]:
                shared_figures[shared] = q["number"]

    # Cache extracted figures so we don't re-extract for shared questions
    if figure_cache is None:
        figure_cache = {}

    for idx, q in enumerate(questions):
        q_num = q["number"]
        cls = classifications.get(q_num, {})
        difficulty = cls.get("difficulty", "Medium").strip() or "Medium"
        chapter = cls.get("chapter", "").strip() or "—"
        reference = cls.get("reference", "").strip() or default_reference
        answer = answers.get(str(q_num), "—")
        motivation = MOTIVATIONS[idx % len(MOTIVATIONS)]

        if status_callback:
            status_callback(f"Building question {q_num}/{len(questions)}...")

        add_question_header(
            doc, q_num, difficulty, 1, reference, chapter,
            page_break=(idx > 0)
        )

        # Body text
        body = q.get("body", "").strip()
        if body:
            add_body_text(doc, body)

        figure_position = q.get("figure_position", "after_options")
        has_figure = q.get("has_figure", False)
        is_shared_consumer = q_num in shared_figures

        # Determine which figure to use
        figure_image = None
        if has_figure:
            page_index = q.get("page", 1) - 1
            try:
                if q_num not in figure_cache:
                    figure_cache[q_num] = crop_figure_for_question(
                        client, qp_pdf_bytes, page_index, q_num,
                        q.get("figure_description", "the figure")
                    )
                figure_image = figure_cache[q_num]
            except Exception as e:
                if status_callback:
                    status_callback(f"⚠️ Couldn't extract figure for Q{q_num}: {e}")

        # Insert figure based on position (before options is the default for a single figure)
        if figure_image and figure_position in ("before_body", "between_body_and_options"):
            add_image(doc, figure_image)

        # Options
        for opt in q.get("options", []):
            add_option(doc, opt)

        # Figure after options
        if figure_image and figure_position == "after_options":
            add_image(doc, figure_image)

        # Student solution section
        add_student_solution_section(doc)

        # Answer section
        add_answer_section(doc, answer, motivation)

    return doc

# ---------------------------- Streamlit UI ----------------------------

with st.expander("ℹ️ How to use this app", expanded=False):
    st.markdown("""
    1. **Upload your Question Paper** (PDF) — the original exam paper.
    2. **Upload the Mark Scheme** (PDF) — contains the correct answers.
    3. **Upload your Classification Sheet** (Excel/CSV) — has chapter, difficulty, etc. per question.
    4. Click **Generate** and wait a couple of minutes.
    5. **Download** the formatted DOCX.

    **Notes:**
    - The first generation per session may take 2-5 minutes depending on paper length.
    - Image extraction is approximate — you may want to check images after download.
    - Your API key is never stored; it stays in your browser session.
    """)

# API key input
st.subheader("🔑 API Key")
api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
if not api_key:
    api_key = st.text_input(
        "Anthropic API Key",
        type="password",
        help="Get one at console.anthropic.com — needed for Claude access."
    )
else:
    st.success("API key loaded from server config ✓")

# File uploads
st.subheader("📁 Upload Files")
col1, col2 = st.columns(2)
with col1:
    qp_file = st.file_uploader("Question Paper (PDF)", type=["pdf"], key="qp")
    xlsx_file = st.file_uploader("Classification Sheet (Excel)", type=["xlsx", "xls"], key="xlsx")
with col2:
    ms_file = st.file_uploader("Mark Scheme (PDF)", type=["pdf"], key="ms")
    default_ref = st.text_input(
        "Default Reference (optional)",
        placeholder="e.g. (May 2021)",
        help="Used if the Excel doesn't include a reference column."
    )

# Generate button
ready = bool(api_key and qp_file and ms_file and xlsx_file)
if st.button("🚀 Generate Worksheet", disabled=not ready, type="primary", use_container_width=True):
    try:
        client = anthropic.Anthropic(api_key=api_key)
        qp_bytes = qp_file.read()
        ms_bytes = ms_file.read()
        xlsx_bytes = xlsx_file.read()

        progress_box = st.empty()
        log_box = st.empty()
        log = []

        def update_status(msg):
            log.append(msg)
            log_box.info(msg)

        with st.spinner("Working..."):
            update_status("📖 Extracting questions from question paper...")
            questions = extract_questions(client, qp_bytes)

            update_status(f"✓ Found {len(questions)} questions.")
            update_status("📖 Extracting answers from mark scheme...")
            answers = extract_answers(client, ms_bytes)

            update_status(f"✓ Got {len(answers)} answers.")
            update_status("📖 Reading classification sheet...")
            classifications = read_classifications(xlsx_bytes)
            update_status(f"✓ Got classifications for {len(classifications)} questions.")

            update_status("🖼️ Building document (this is the slow part — extracting figures)...")
            doc = build_document(
                questions, answers, classifications,
                qp_bytes, client,
                status_callback=update_status,
                default_reference=default_ref
            )

            update_status("💾 Saving document...")
            output = io.BytesIO()
            doc.save(output)
            output.seek(0)

        st.success(f"✅ Done! Generated {len(questions)} questions.")

        st.download_button(
            label="⬇️ Download Worksheet",
            data=output,
            file_name="worksheet.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

    except anthropic.AuthenticationError:
        st.error("❌ Invalid API key. Please check your Anthropic API key.")
    except json.JSONDecodeError as e:
        st.error(f"❌ Couldn't parse Claude's response as JSON. Try again. Error: {e}")
    except Exception as e:
        st.error(f"❌ Error: {type(e).__name__}: {e}")
        st.exception(e)

elif not ready:
    missing = []
    if not api_key: missing.append("API key")
    if not qp_file: missing.append("Question Paper")
    if not ms_file: missing.append("Mark Scheme")
    if not xlsx_file: missing.append("Classification Sheet")
    st.info(f"⏳ Waiting for: {', '.join(missing)}")

st.markdown("---")
st.caption("Built with Streamlit + Claude API. Your files are processed in-memory and never stored.")
