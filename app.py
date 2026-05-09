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

def clean_latex_text(text):
    """Convert simple LaTeX math notation in spreadsheet cells to readable text.

    Spreadsheets prepared with LaTeX-style math (e.g. ``$O_2$``, ``$1.2 \\times 10^{24}$``,
    ``$Br^-$``) need to be converted to readable Unicode for the worksheet output,
    since python-docx doesn't render LaTeX. This is a best-effort conversion that
    handles the most common chemistry/math patterns. Anything we can't convert is
    left in place, with the surrounding ``$`` markers stripped.

    Conversions performed:
      - Subscripts: ``X_2``, ``X_{12}`` → ``X₂``, ``X₁₂`` (digit subscripts only)
      - Superscripts: ``X^2``, ``X^{2+}``, ``X^-`` → ``X²``, ``X²⁺``, ``X⁻``
      - Common operators: ``\\times`` → ``×``, ``\\rightarrow`` → ``→``, ``\\circ`` → ``°``
      - Fractions: ``\\frac{a}{b}`` → ``a/b``
      - Greek letters: ``\\alpha`` → ``α``, etc.
      - Strip remaining LaTeX control sequences and ``$`` markers
    """
    if not text or not isinstance(text, str):
        return text

    s = text

    # First: replace common LaTeX operators and symbols. Order matters — do
    # multi-char replacements before single-char ones, and do these BEFORE
    # super/subscript expansion so that things like "40^{\circ}C" → "40°C"
    # work (otherwise the °  ends up trapped inside ^{...} which is then
    # converted to a superscript).
    replacements = {
        r"\\times": "×",
        r"\\cdot": "·",
        r"\\div": "÷",
        r"\\pm": "±",
        r"\\mp": "∓",
        r"\\rightarrow": "→",
        r"\\leftarrow": "←",
        r"\\Rightarrow": "⇒",
        r"\\leftrightarrow": "↔",
        r"\\to": "→",
        r"\\circ": "°",
        r"\\degree": "°",
        r"\\infty": "∞",
        r"\\leq": "≤",
        r"\\geq": "≥",
        r"\\neq": "≠",
        r"\\approx": "≈",
        r"\\sim": "~",
        r"\\equiv": "≡",
        r"\\sqrt": "√",
        r"\\Delta": "Δ", r"\\delta": "δ",
        r"\\alpha": "α", r"\\beta": "β", r"\\gamma": "γ",
        r"\\theta": "θ", r"\\lambda": "λ", r"\\mu": "μ",
        r"\\pi": "π", r"\\sigma": "σ", r"\\omega": "ω",
        r"\\Omega": "Ω", r"\\Sigma": "Σ", r"\\Phi": "Φ",
        r"\\ominus": "⊖", r"\\oplus": "⊕",
    }
    for pattern, replacement in replacements.items():
        s = re.sub(pattern, replacement, s)

    # Strip the `^{X}` wrapper around already-substituted symbols. After the
    # operator substitutions above, "40^{°}C" is "40^{°}C"; we want "40°C".
    # We strip ^{...} when its content is already a single Unicode symbol.
    s = re.sub(r"\^\{(°|×|÷|±|∓|→|←|⇒|↔|⊖|⊕|∞|·|√)\}", r"\1", s)

    # Now subscripts and superscripts (digits/sign/parens only).
    sub_map = str.maketrans("0123456789+-=()", "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎")
    s = re.sub(r"_\{([0-9+\-=()]+)\}", lambda m: m.group(1).translate(sub_map), s)
    s = re.sub(r"_([0-9])", lambda m: m.group(1).translate(sub_map), s)

    sup_map = str.maketrans("0123456789+-=()n", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿ")
    s = re.sub(r"\^\{([0-9+\-=()n]+)\}", lambda m: m.group(1).translate(sup_map), s)
    s = re.sub(r"\^([0-9+\-n])", lambda m: m.group(1).translate(sup_map), s)

    # Fractions: \frac{a}{b} → (a)/(b) (best-effort, single-level)
    s = re.sub(r"\\frac\{([^{}]*)\}\{([^{}]*)\}", r"(\1)/(\2)", s)

    # \text{X} / \mathrm{X} → X (preserve content)
    s = re.sub(r"\\text\{([^{}]*)\}", r"\1", s)
    s = re.sub(r"\\mathrm\{([^{}]*)\}", r"\1", s)

    # Strip $ math delimiters
    s = s.replace("$", "")

    # Strip any remaining \xxx commands we didn't recognize
    s = re.sub(r"\\[a-zA-Z]+", "", s)

    # Markdown bold around question numbers: **5.** → 5.
    s = re.sub(r"\*\*(\d+\.)\*\*", r"\1", s)

    # Collapse whitespace
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_question_from_excel_row(row, position):
    """Build a question dict from a spreadsheet row.

    Returns the same shape as extract_questions returns (number, body, options,
    page, has_figure, etc.) so build_document can consume either source.

    Detects when a question has missing/empty options (e.g. "A.\\nB.\\nC.\\nD.")
    and marks it as needing a figure placeholder, so the user knows to insert
    the figure manually after downloading.
    """
    raw_text = str(row.get("Question Text", "") or "")

    # Split into body lines and option lines first (before LaTeX cleaning,
    # which collapses whitespace including newlines).
    # Lines that begin with "A. ", "B. ", "C. ", "D. " (or "A.", "B.", etc.
    # for figure-only options) become options.
    lines = [l.strip() for l in raw_text.split("\n")]
    body_lines = []
    options = []
    for line in lines:
        m = re.match(r"^([A-D])\.\s*(.*)$", line)
        if m:
            letter, content = m.group(1), m.group(2).strip()
            content_clean = clean_latex_text(content) if content else ""
            options.append(f"{letter}.  {content_clean}" if content_clean else f"{letter}.")
        elif line:
            body_lines.append(line)

    # Strip leading "**N.**" markdown numbering from the body
    if body_lines:
        body_lines[0] = re.sub(r"^\*\*\d+\.\*\*\s*", "", body_lines[0])

    body = " ".join(body_lines).strip()
    body = clean_latex_text(body)

    # Detect when options are empty (figure-based question with image options)
    empty_options = sum(1 for o in options if re.match(r"^[A-D]\.$", o.strip())) >= 3
    # Detect when the body contains a markdown table (also a figure indicator)
    has_table_in_body = "|" in body and re.search(r"\|.{1,30}\|", body) is not None

    needs_figure_placeholder = empty_options or has_table_in_body

    # If the body has a markdown-style table, strip it and flag for placeholder
    if has_table_in_body:
        body = re.sub(r"\|[^|\n]*\|", "", body).strip()
        body = re.sub(r"\s+", " ", body)

    qnum_raw = row.get("Question No.", position)
    try:
        qnum = int(qnum_raw) if not pd.isna(qnum_raw) else position
    except (ValueError, TypeError):
        qnum = position

    page_raw = row.get("Page Number", 1)
    try:
        page = int(page_raw) if not pd.isna(page_raw) else 1
    except (ValueError, TypeError):
        page = 1

    return {
        "number": qnum,
        "position": position,
        "body": body,
        "options": options,
        "page": page,
        "has_figure": needs_figure_placeholder,
        "needs_manual_figure": needs_figure_placeholder,
        "figure_position": "between_body_and_options",
    }


def extract_answer_from_mark_scheme_cell(cell):
    """Extract a single answer letter (A-D) from a Mark Scheme cell.

    Cell values look like: "1) D", "MCQ) A", "A-D) B", "C", "A. ...".
    We find the first standalone uppercase letter A-D after stripping
    common prefixes.
    """
    if cell is None or pd.isna(cell):
        return None
    s = str(cell).strip()
    # Strip common prefixes: "1)", "1.", "MCQ)", "A-D)"
    s = re.sub(r"^(MCQ|A-D|\d+)[\)\.]\s*", "", s, flags=re.IGNORECASE)
    # Find first standalone A-D
    m = re.match(r"\s*([A-D])\b", s)
    return m.group(1) if m else None


def extract_questions_from_excel(xlsx_bytes, qp_file_name=None, status_callback=None):
    """Read questions, options, and answers directly from the spreadsheet.

    This is an alternative to extract_questions() / extract_answers() that
    bypasses the Claude API entirely. It's faster, free, and avoids PDF token
    limit errors — but requires the spreadsheet to contain the full question
    text and mark scheme letters in the appropriate columns.

    Returns a tuple (questions, answers) where:
      - questions is a list of dicts in the same shape as extract_questions
      - answers is a dict {str(qnum): "A"|"B"|"C"|"D"}
    """
    df = pd.read_excel(io.BytesIO(xlsx_bytes))
    cols = {c.lower().strip(): c for c in df.columns}

    def find_col(*candidates):
        for c in candidates:
            if c.lower() in cols:
                return cols[c.lower()]
        return None

    text_col = find_col("Question Text", "Question", "Text")
    ms_col = find_col("Mark Scheme", "Answer", "MS")
    file_col = find_col("File Name", "FileName", "File", "Source")

    if text_col is None:
        raise ValueError(
            "The spreadsheet doesn't have a 'Question Text' column. "
            "Cannot extract questions from Excel."
        )

    target_name = None
    if qp_file_name and file_col:
        target_name = _normalize_filename(qp_file_name)

    questions = []
    answers = {}
    position = 0

    for _, row in df.iterrows():
        if target_name is not None:
            row_file = _normalize_filename(str(row[file_col]))
            if row_file != target_name and target_name not in row_file and row_file not in target_name:
                continue

        position += 1
        q = parse_question_from_excel_row(row, position)
        questions.append(q)

        if ms_col is not None:
            ans = extract_answer_from_mark_scheme_cell(row[ms_col])
            if ans:
                # Key answers by both the printed Q# and the position so either
                # lookup works in build_document.
                answers[str(q["number"])] = ans
                answers[f"pos:{position}"] = ans

    if status_callback:
        figure_count = sum(1 for q in questions if q.get("needs_manual_figure"))
        status_callback(
            f"📑 Extracted {len(questions)} questions from spreadsheet. "
            f"{figure_count} need a figure inserted manually."
        )

    return questions, answers


# ---------------------------- PDF Chunking ----------------------------
# Claude's context window is 200K tokens. PDFs sent as documents are token-heavy
# because Claude processes each page as both text AND an image. Empirically,
# a typical printed exam page uses ~1.5–3K tokens, while a scanned/high-res page
# can use 6–10K tokens or more.
#
# To stay safely under the 200K limit (and leave headroom for the prompt and
# response), we split large PDFs into small chunks and process each chunk
# independently. We also estimate token usage upfront so we can warn the user
# before making any API calls.

# Default pages per chunk. Conservative — keeps each request well under
# 50K tokens for typical text PDFs, and under 100K even for scanned PDFs.
DEFAULT_PAGES_PER_CHUNK = 3

# Token budget per chunk. We aim to stay below this to leave room for the
# prompt, response, and safety margin. Claude's hard limit is 200K input tokens.
TARGET_TOKENS_PER_CHUNK = 80_000

# Hard ceiling on a single chunk's estimated tokens. If even one page exceeds
# this, we'll still try to send it on its own (and let Claude reject it if it
# truly can't fit).
MAX_TOKENS_PER_CHUNK = 180_000

# Estimated tokens per byte of PDF content. Very rough heuristic that works
# for most exam PDFs. Pure-text PDFs are ~0.4 tokens/byte, scanned image-heavy
# PDFs are ~0.05 tokens/byte (because the image gets counted differently).
# We use a middle estimate that's slightly conservative.
EST_TOKENS_PER_BYTE = 0.15

# Hard cap on total pages we'll accept, to avoid runaway cost / time.
MAX_TOTAL_PAGES = 200


class InputTooLargeError(Exception):
    """Raised when input exceeds what Claude can process, even after chunking."""
    pass


def get_pdf_page_count(pdf_bytes):
    """Return the number of pages in a PDF."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    n = doc.page_count
    doc.close()
    return n


def estimate_pdf_tokens(pdf_bytes):
    """Estimate how many tokens a PDF will consume when sent to Claude.

    This is an approximation. The actual count depends on the page contents
    (text density, embedded images, scan resolution). We err on the side of
    overestimating so we don't accidentally exceed the limit.
    """
    return int(len(pdf_bytes) * EST_TOKENS_PER_BYTE)


def get_page_pdf_bytes(pdf_bytes, page_index):
    """Extract a single page from a PDF as its own PDF byte string."""
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    new_doc = fitz.open()
    new_doc.insert_pdf(src, from_page=page_index, to_page=page_index)
    buf = io.BytesIO()
    new_doc.save(buf)
    new_doc.close()
    src.close()
    return buf.getvalue()


def build_size_aware_chunks(pdf_bytes, target_tokens=TARGET_TOKENS_PER_CHUNK):
    """Split a PDF into chunks based on estimated token cost, not just page count.

    Walks pages sequentially, accumulating them into a chunk until adding another
    page would exceed `target_tokens`. Single pages that already exceed the target
    become their own chunk.

    Returns a list of (chunk_pdf_bytes, start_page_1indexed, end_page_1indexed,
    estimated_tokens) tuples. Page numbers refer to the ORIGINAL PDF.
    """
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    total = src.page_count

    # Pre-compute per-page estimated tokens by extracting each page individually
    # and measuring its serialized size.
    page_costs = []
    for i in range(total):
        single = fitz.open()
        single.insert_pdf(src, from_page=i, to_page=i)
        buf = io.BytesIO()
        single.save(buf)
        single.close()
        page_costs.append(estimate_pdf_tokens(buf.getvalue()))

    # Greedily pack pages into chunks
    chunks = []
    chunk_start = 0  # 0-indexed
    chunk_tokens = 0
    chunk_end = 0  # exclusive

    for i in range(total):
        page_cost = page_costs[i]
        # If adding this page would exceed the target AND the current chunk has
        # at least one page, close the current chunk and start a new one.
        if chunk_end > chunk_start and chunk_tokens + page_cost > target_tokens:
            new_doc = fitz.open()
            new_doc.insert_pdf(src, from_page=chunk_start, to_page=chunk_end - 1)
            buf = io.BytesIO()
            new_doc.save(buf)
            new_doc.close()
            chunks.append((buf.getvalue(), chunk_start + 1, chunk_end, chunk_tokens))
            chunk_start = i
            chunk_tokens = 0

        chunk_end = i + 1
        chunk_tokens += page_cost

    # Flush the final chunk
    if chunk_end > chunk_start:
        new_doc = fitz.open()
        new_doc.insert_pdf(src, from_page=chunk_start, to_page=chunk_end - 1)
        buf = io.BytesIO()
        new_doc.save(buf)
        new_doc.close()
        chunks.append((buf.getvalue(), chunk_start + 1, chunk_end, chunk_tokens))

    src.close()
    return chunks


def split_pdf_into_chunks(pdf_bytes, pages_per_chunk=DEFAULT_PAGES_PER_CHUNK):
    """Split a PDF into chunks of N pages each (legacy fixed-size splitter).

    Kept for the recursive fallback splitter. For the main extraction path,
    prefer `build_size_aware_chunks`.

    Returns a list of (chunk_pdf_bytes, start_page_1indexed, end_page_1indexed)
    tuples. Page numbers refer to the ORIGINAL PDF.
    """
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    total = src.page_count
    chunks = []
    for start in range(0, total, pages_per_chunk):
        end = min(start + pages_per_chunk, total)
        new_doc = fitz.open()
        new_doc.insert_pdf(src, from_page=start, to_page=end - 1)
        buf = io.BytesIO()
        new_doc.save(buf)
        new_doc.close()
        chunks.append((buf.getvalue(), start + 1, end))
    src.close()
    return chunks


def is_too_large_error(exc):
    """Detect whether an Anthropic error is specifically a context-length error.

    Newer SDK versions use different message formats, so we check several
    signals: the error message text, the error body type, and the status code.
    """
    msg = str(exc).lower()
    # Most common phrasings used by the API
    if "prompt is too long" in msg:
        return True
    if "exceeds" in msg and "tokens" in msg:
        return True
    if "context window" in msg:
        return True
    if "context_length_exceeded" in msg:
        return True
    # Try the structured body if available
    try:
        body = getattr(exc, "body", None) or {}
        err = body.get("error", {}) if isinstance(body, dict) else {}
        if "too long" in str(err.get("message", "")).lower():
            return True
        if err.get("type") == "invalid_request_error" and "tokens" in str(err.get("message", "")).lower():
            return True
    except Exception:
        pass
    return False


def call_claude_with_pdf(client, pdf_bytes, prompt, max_tokens=8000):
    """Send a PDF to Claude with a prompt and return the text response.

    Raises InputTooLargeError if Anthropic rejects the request as too large.
    Callers that want to handle large inputs should split the PDF first.
    """
    try:
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
    except anthropic.BadRequestError as e:
        if is_too_large_error(e):
            raise InputTooLargeError(str(e)) from e
        raise

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

def process_chunk_with_adaptive_split(
    client, chunk_bytes, prompt, max_tokens, start_page,
    status_callback=None
):
    """Try to send a chunk to Claude. If it's rejected as too large, recursively
    split it in half and try the smaller pieces. If even a single page is too
    large, fall back to sending it as a downscaled image instead of a PDF.

    Returns a list of (response_text, sub_start_page) tuples.
    """
    # Try the chunk as-is first
    try:
        text = call_claude_with_pdf(client, chunk_bytes, prompt, max_tokens=max_tokens)
        return [(text, start_page)]
    except InputTooLargeError:
        pass  # Fall through to splitting

    # The chunk is too large. How many pages does it have?
    chunk_pages = get_pdf_page_count(chunk_bytes)

    if chunk_pages == 1:
        # Can't split further by pages. Try sending as a downscaled image instead.
        if status_callback:
            status_callback(
                f"⚠️  Page {start_page} is too large as PDF. "
                f"Retrying as a downscaled image..."
            )
        try:
            page_img = render_page_image(chunk_bytes, 0, dpi=120)
            # Cap dimensions to control image token cost
            page_img.thumbnail((1568, 1568))
            buf = io.BytesIO()
            page_img.save(buf, format="JPEG", quality=80)
            text = call_claude_with_image(
                client, buf.getvalue(), prompt, max_tokens=max_tokens
            )
            return [(text, start_page)]
        except Exception as fallback_err:
            raise InputTooLargeError(
                f"Page {start_page} is too large to process even as a downscaled image. "
                f"This usually means the page contains very high-resolution scanned content. "
                f"Try re-saving the PDF with reduced image quality and upload it again."
            ) from fallback_err

    # Multi-page chunk: split it in half and recurse.
    new_size = max(1, chunk_pages // 2)
    if status_callback:
        status_callback(
            f"⚠️  Chunk at pages {start_page}–{start_page + chunk_pages - 1} "
            f"was too large. Splitting into pieces of {new_size} page(s) and retrying..."
        )

    sub_chunks = split_pdf_into_chunks(chunk_bytes, new_size)
    results = []
    for sub_bytes, sub_start, sub_end in sub_chunks:
        # sub_start is 1-indexed within the original chunk; map back to absolute.
        absolute_start = start_page + (sub_start - 1)
        results.extend(
            process_chunk_with_adaptive_split(
                client, sub_bytes, prompt, max_tokens, absolute_start,
                status_callback=status_callback
            )
        )
    return results


def extract_questions(client, qp_pdf_bytes, status_callback=None):
    """Extract structured question data from QP, chunking the PDF if needed.

    The PDF is split into chunks sized by estimated token cost (not just page
    count) so that even image-heavy or scanned PDFs are processed in pieces
    that fit comfortably under Claude's 200K context window. Each chunk is
    sent to Claude separately. If any chunk is rejected as too large, it is
    automatically split further (and ultimately falls back to image-mode for
    single pages). Results are merged into a single list, with page numbers
    offset to match the original PDF and duplicate question numbers across
    chunk boundaries de-duplicated.
    """
    total_pages = get_pdf_page_count(qp_pdf_bytes)
    if total_pages > MAX_TOTAL_PAGES:
        raise InputTooLargeError(
            f"This PDF has {total_pages} pages, which exceeds the maximum of "
            f"{MAX_TOTAL_PAGES} we can process in one run. Please split it into "
            f"smaller files (for example, one paper per upload) and try again."
        )

    # Build chunks based on estimated token cost
    chunks = build_size_aware_chunks(qp_pdf_bytes, target_tokens=TARGET_TOKENS_PER_CHUNK)
    all_questions = []
    seen_numbers = set()

    if status_callback:
        status_callback(
            f"📑 Question paper split into {len(chunks)} chunk(s) for processing."
        )

    for i, (chunk_bytes, start_page, end_page, est_tokens) in enumerate(chunks):
        if status_callback:
            status_callback(
                f"📖 Reading question paper — chunk {i + 1}/{len(chunks)} "
                f"(pages {start_page}–{end_page}, ~{est_tokens:,} tokens)..."
            )
        results = process_chunk_with_adaptive_split(
            client, chunk_bytes, EXTRACTION_PROMPT,
            max_tokens=8000, start_page=start_page,
            status_callback=status_callback
        )

        for text, sub_start_page in results:
            try:
                data = parse_json_from_text(text)
            except json.JSONDecodeError:
                if status_callback:
                    status_callback(
                        f"⚠️  Couldn't parse JSON from chunk at page {sub_start_page}, skipping."
                    )
                continue

            for q in data.get("questions", []):
                # Page numbers from each chunk are 1-indexed within that chunk;
                # convert them to absolute page numbers in the original PDF.
                local_page = q.get("page", 1)
                q["page"] = sub_start_page + (local_page - 1)

                # De-duplicate by question number (in case a question straddles
                # the boundary between two chunks).
                qnum = q.get("number")
                if qnum is None or qnum in seen_numbers:
                    continue
                seen_numbers.add(qnum)
                all_questions.append(q)

    # Sort by question number for safety
    all_questions.sort(key=lambda q: q.get("number", 0))

    # Tag each question with its 1-indexed position in the QP. The user's
    # spreadsheet is read in row order, and we look up classifications by
    # this position rather than by the printed Q# (which can be ambiguous
    # if the spreadsheet has typos in the Q# column).
    for i, q in enumerate(all_questions):
        q["position"] = i + 1

    return all_questions

def extract_answers(client, ms_pdf_bytes, status_callback=None):
    """Extract answer key from Mark Scheme, chunking the PDF if needed."""
    total_pages = get_pdf_page_count(ms_pdf_bytes)
    if total_pages > MAX_TOTAL_PAGES:
        raise InputTooLargeError(
            f"The mark scheme has {total_pages} pages, which exceeds the maximum "
            f"of {MAX_TOTAL_PAGES} we can process in one run."
        )

    chunks = build_size_aware_chunks(ms_pdf_bytes, target_tokens=TARGET_TOKENS_PER_CHUNK)
    answers = {}

    if status_callback:
        status_callback(
            f"📑 Mark scheme split into {len(chunks)} chunk(s) for processing."
        )

    for i, (chunk_bytes, start_page, end_page, est_tokens) in enumerate(chunks):
        if status_callback:
            status_callback(
                f"📖 Reading mark scheme — chunk {i + 1}/{len(chunks)} "
                f"(pages {start_page}–{end_page}, ~{est_tokens:,} tokens)..."
            )
        results = process_chunk_with_adaptive_split(
            client, chunk_bytes, ANSWERS_PROMPT,
            max_tokens=2000, start_page=start_page,
            status_callback=status_callback
        )

        for text, sub_start_page in results:
            try:
                data = parse_json_from_text(text)
            except json.JSONDecodeError:
                if status_callback:
                    status_callback(
                        f"⚠️  Couldn't parse mark scheme chunk at page {sub_start_page}, skipping."
                    )
                continue

            for qnum, ans in data.get("answers", {}).items():
                # Don't overwrite if we already have this answer from an earlier chunk.
                if qnum not in answers:
                    answers[qnum] = ans

    return answers

def read_classifications(xlsx_bytes, qp_file_name=None, status_callback=None):
    """Read chapter / difficulty / reference from the Excel sheet.

    The Excel sheet is read **in row order** — the first row corresponds to
    the first question in the QP, the second row to the second question, and
    so on. The 'Question No.' column is IGNORED for keying because user-prepared
    spreadsheets often contain typos in question numbers (e.g. multiple rows
    incorrectly labeled Q#=1).

    Args:
        xlsx_bytes: Raw bytes of the uploaded Excel file.
        qp_file_name: If given, only rows whose 'File Name' column matches this
            (case-insensitive, file-extension-stripped) are returned BEFORE
            sequential numbering is assigned. This handles spreadsheets that
            contain rows from multiple exam papers.
        status_callback: Optional function called with diagnostic messages
            about which columns were detected and how many rows were kept.
            Useful for surfacing column-name mismatches in the UI.

    Returns:
        A dict mapping sequential 1-indexed position (int) to a dict with keys
        'chapter', 'difficulty', and 'reference'. The position corresponds to
        the question's position in the QP (1 = first question, 2 = second, etc.),
        NOT to the value in the 'Question No.' column.
    """
    df = pd.read_excel(io.BytesIO(xlsx_bytes))
    cols = {c.lower().strip(): c for c in df.columns}

    def find_col(*candidates):
        for c in candidates:
            if c.lower() in cols:
                return cols[c.lower()]
        return None

    chapter_col = find_col("Topic", "Chapter", "Subject", "Unit")
    diff_col = find_col("Difficulty", "Level", "Level of question")
    ref_col = find_col("Reference", "Date", "Year")
    file_col = find_col("File Name", "FileName", "File", "Source")

    # Diagnostic: report which columns were detected so the user can see if
    # any expected column is missing from their spreadsheet.
    if status_callback:
        status_callback(
            f"📋 Excel columns detected — "
            f"Chapter: {chapter_col!r}, "
            f"Difficulty: {diff_col!r}, "
            f"Reference: {ref_col!r}, "
            f"File Name: {file_col!r}"
        )
        all_cols = list(df.columns)
        status_callback(f"📋 All columns in your spreadsheet: {all_cols}")

    # Filter by file name first (so the position-counter only counts kept rows).
    target_name = None
    if qp_file_name and file_col:
        target_name = _normalize_filename(qp_file_name)

    classifications = {}
    position = 0  # 1-indexed position assigned in row order

    for _, row in df.iterrows():
        if target_name is not None:
            row_file = _normalize_filename(str(row[file_col]))
            if row_file != target_name and target_name not in row_file and row_file not in target_name:
                continue

        position += 1
        classifications[position] = {
            "chapter": str(row[chapter_col]).strip() if chapter_col else "",
            "difficulty": str(row[diff_col]).strip() if diff_col else "Medium",
            "reference": str(row[ref_col]).strip() if ref_col else "",
        }
    return classifications


def _normalize_filename(name):
    """Normalize a filename for comparison.

    Handles common variations between how a file is named on disk vs how it's
    referenced in the spreadsheet:
      - case differences ("Biology" vs "biology")
      - file extensions (.pdf, .docx)
      - underscores vs spaces ("Higher_level" vs "Higher level")
      - dashes ("-" vs "_")
      - "Copy" / "Copy 2" / etc. suffixes added by file managers
      - trailing/leading whitespace
    """
    name = str(name).strip().lower()
    # Strip common file extensions
    for ext in (".pdf", ".docx", ".doc"):
        if name.endswith(ext):
            name = name[:-len(ext)]
    # Replace underscores and dashes with spaces
    name = name.replace("_", " ").replace("-", " ")
    # Remove "copy" suffixes that some file managers append
    name = re.sub(r"\s+copy(\s+\d+)?$", "", name)
    # Collapse all whitespace to single spaces
    name = re.sub(r"\s+", " ", name).strip()
    return name

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
    """Add a single writing line (paragraph with a bottom border).

    Note: kept for backward-compatibility. New code should prefer
    add_writing_lines_table() which produces more reliable output across
    Word and LibreOffice renderers.
    """
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(8)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "4")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "BFBFBF")
    pBdr.append(bottom)
    pPr.append(pBdr)
    add_run(p, "\u00A0", size=11)
    return p


def add_writing_lines_table(doc, num_lines):
    """Add a single-column table with `num_lines` rows, where each row's
    bottom border draws one writing line. This renders reliably in both
    Word and LibreOffice (unlike empty-paragraph + border, which collapses).
    """
    from docx.shared import Cm
    table = doc.add_table(rows=num_lines, cols=1)
    table.autofit = False
    # Make the table span the full text width
    for row in table.rows:
        row.height = Cm(0.9)  # ~25 points per line — enough room to write
        cell = row.cells[0]
        # Set cell width (approximately 6.5" = page width minus 1" margins)
        cell.width = Inches(6.5)
        # Configure cell borders: only show the bottom border
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement("w:tcBorders")
        for side in ("top", "left", "right"):
            b = OxmlElement(f"w:{side}")
            b.set(qn("w:val"), "nil")
            tcBorders.append(b)
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "4")
        bottom.set(qn("w:color"), "BFBFBF")
        tcBorders.append(bottom)
        tcPr.append(tcBorders)
        # Empty paragraph in the cell
        cell.paragraphs[0].add_run("")
    return table

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


def add_figure_placeholder(doc, q_num):
    """Add a visible placeholder box where a figure or table should be inserted.

    Used when the question has an associated figure but we couldn't extract it
    automatically (e.g. when running in Excel-only mode with no PDF access).
    The user can then paste a screenshot or insert an image into this placeholder
    after downloading the document.
    """
    table = doc.add_table(rows=1, cols=1)
    table.autofit = False
    cell = table.rows[0].cells[0]
    cell.width = Inches(6.5)

    # Configure cell borders: a dashed border on all sides to make the
    # placeholder visually distinct from the writing-line tables.
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "bottom", "left", "right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "dashed")
        b.set(qn("w:sz"), "8")
        b.set(qn("w:color"), "808080")
        tcBorders.append(b)
    tcPr.append(tcBorders)

    # Vertical centering and a bit of padding inside the cell
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), "center")
    tcPr.append(vAlign)
    tcMar = OxmlElement("w:tcMar")
    for side in ("top", "bottom"):
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"), "240")  # 240 twips = ~12pt
        m.set(qn("w:type"), "dxa")
        tcMar.append(m)
    tcPr.append(tcMar)

    # Friendly label centered in the box
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"📷  [Insert figure / table for Q{q_num} here]")
    run.font.name = "Arial"
    run.font.size = Pt(11)
    run.font.italic = True
    run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    # Add a small spacing paragraph after the placeholder
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(4)


# Number of blank lines printed under "Student's Solution:" for the student to write in.
STUDENT_SOLUTION_LINES = 4

def add_student_solution_section(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    add_run(p, "Student's Solution:", bold=True, size=11)
    add_writing_lines_table(doc, STUDENT_SOLUTION_LINES)

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

# ---------------------------- Sorting Helpers ----------------------------

# Difficulty ordering: Easy questions first within each chapter, then Medium, then Hard.
DIFFICULTY_ORDER = {
    "easy": 0,
    "medium": 1,
    "hard": 2,
    "difficult": 2,
}

def difficulty_rank(diff_str):
    """Map a difficulty string to a sort key. Unknown values sort last."""
    return DIFFICULTY_ORDER.get((diff_str or "").strip().lower(), 99)


def chapter_sort_key(chapter_str):
    """Extract a sortable key from a chapter string.

    Tries to find a leading number (e.g. "ch1. Cell Biology", "1. Atomic Structure",
    "Chapter 3 - Genetics") so chapters sort numerically. Falls back to the raw
    string for chapters without a number.

    Returns (numeric_key, original_text) so chapters with numbers sort by number
    first, then alphabetically, and chapters without numbers go to the end.
    """
    if not chapter_str:
        return (9999, "")
    # Strip common prefixes and find the first integer in the string
    text = chapter_str.strip()
    match = re.search(r"\d+", text)
    if match:
        return (int(match.group()), text.lower())
    return (9999, text.lower())


def year_sort_key(reference_str):
    """Extract a year from a reference string for sorting.

    References look like "(May 2021)", "8825-6220 (31 October 2025)", "May 2023",
    "2024", etc. We extract the first 4-digit number that looks like a year
    (1900-2099). Returns 0 for missing/unparseable references so they sort first
    (treated as "no year info").
    """
    if not reference_str:
        return 0
    match = re.search(r"\b(19|20)\d{2}\b", reference_str)
    return int(match.group()) if match else 0


def sort_questions_for_worksheet(questions, classifications, default_reference=""):
    """Sort questions in worksheet display order.

    Order: Chapter (numeric) → Difficulty (Easy → Medium → Hard) → Year (oldest first)
    → Position in extraction order (stable tiebreak).

    Returns a new list. Does not modify the input.
    """
    def key(q):
        qnum = q.get("number", 0)
        position = q.get("position", qnum)
        cls = classifications.get(position, {})
        chapter = (cls.get("chapter") or "—").strip()
        difficulty = (cls.get("difficulty") or "Medium").strip()
        reference = (cls.get("reference") or default_reference or "").strip()
        return (
            chapter_sort_key(chapter),
            difficulty_rank(difficulty),
            year_sort_key(reference),
            position,  # stable tiebreak by extraction order
        )
    return sorted(questions, key=key)


def add_chapter_heading(doc, chapter_text, page_break=True):
    """Add a large, centered chapter heading paragraph."""
    p = doc.add_paragraph()
    if page_break:
        p.paragraph_format.page_break_before = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(20)
    p.paragraph_format.space_after = Pt(20)
    run = p.add_run(chapter_text)
    run.font.name = "Arial"
    run.font.size = Pt(20)
    run.font.bold = True


def add_difficulty_subheading(doc, difficulty_text):
    """Add a smaller subheading for the difficulty group within a chapter."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(f"— {difficulty_text} questions —")
    run.font.name = "Arial"
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.italic = True


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
    """Build the final DOCX from extracted data.

    Questions are organized into the worksheet in this order:
      1. By chapter (numeric, 1, 2, 3, ...)
      2. Within each chapter, by difficulty (Easy → Medium → Hard)
      3. Within each difficulty group, by year (oldest first)
      4. Original question number as tiebreak

    Each chapter starts on a new page with a chapter heading. Difficulty groups
    inside a chapter are introduced with a smaller subheading.

    The `classifications` dict is keyed by 1-indexed POSITION (Excel row order),
    NOT by the question number printed in the QP. This is because the user's
    spreadsheet may contain typos in the Q# column. We assume the spreadsheet
    rows are in the same order as the questions appear in the QP.
    """
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

    # Tag each question with a 1-indexed position in extraction order. This is
    # how we'll look up its classification (chapter/difficulty/reference), not
    # by the printed Q# which may be ambiguous if the spreadsheet has typos.
    for i, q in enumerate(questions):
        q.setdefault("position", i + 1)

    # Build a map of which questions share figures
    shared_figures = {}  # q_num -> primary_q_num (the one that has the actual figure)
    for q in questions:
        if q.get("has_figure") and q.get("shared_with"):
            for shared in q["shared_with"]:
                shared_figures[shared] = q["number"]

    # Cache extracted figures so we don't re-extract for shared questions
    if figure_cache is None:
        figure_cache = {}

    # Sort the questions for worksheet display order
    sorted_questions = sort_questions_for_worksheet(
        questions, classifications, default_reference=default_reference
    )

    # Walk the sorted list and emit chapter / difficulty headings as we go.
    # We track the "current" chapter and difficulty so we only emit a heading
    # when one of them changes.
    current_chapter = None
    current_difficulty = None
    is_first_block = True  # so the very first chapter doesn't get a leading page break

    for idx, q in enumerate(sorted_questions):
        q_num = q["number"]
        position = q.get("position", q_num)  # fall back to q_num if missing
        cls = classifications.get(position, {})
        difficulty = (cls.get("difficulty") or "Medium").strip() or "Medium"
        chapter = (cls.get("chapter") or "—").strip() or "—"
        reference = (cls.get("reference") or "").strip() or default_reference
        # Look up the answer first by question number, then by position as a
        # fallback. The Excel-extraction path keys answers by both, so either
        # lookup works there. The PDF-extraction path only uses qnum.
        answer = (
            answers.get(str(q_num))
            or answers.get(f"pos:{position}")
            or "—"
        )
        motivation = MOTIVATIONS[idx % len(MOTIVATIONS)]

        if status_callback:
            status_callback(
                f"Building question {idx + 1}/{len(sorted_questions)} "
                f"(Q{q_num}, {chapter}, {difficulty})..."
            )

        # Emit a chapter heading when the chapter changes
        if chapter != current_chapter:
            add_chapter_heading(doc, chapter, page_break=not is_first_block)
            current_chapter = chapter
            current_difficulty = None  # reset so we re-emit the difficulty heading
            is_first_block = False
            # Also force a page break before the first question in this chapter
            question_needs_page_break = False  # the chapter heading already broke
        else:
            question_needs_page_break = True

        # Emit a difficulty subheading when the difficulty changes within a chapter
        if difficulty != current_difficulty:
            add_difficulty_subheading(doc, difficulty)
            current_difficulty = difficulty
            question_needs_page_break = False  # subheading flows into the question

        add_question_header(
            doc, q_num, difficulty, 1, reference, chapter,
            page_break=question_needs_page_break
        )

        # Body text
        body = q.get("body", "").strip()
        if body:
            add_body_text(doc, body)

        figure_position = q.get("figure_position", "after_options")
        has_figure = q.get("has_figure", False)
        needs_manual = q.get("needs_manual_figure", False)
        is_shared_consumer = q_num in shared_figures

        # Determine which figure to use. Three cases:
        #   1. We already have an extracted figure cached → use it
        #   2. We have a Claude client and a PDF to extract from → call Claude
        #   3. Excel-based extraction with no figure available → emit a
        #      placeholder paragraph so the user knows where to paste a figure
        figure_image = None
        if has_figure and not needs_manual:
            page_index = q.get("page", 1) - 1
            try:
                if q_num not in figure_cache:
                    if client is None or not qp_pdf_bytes:
                        # No way to extract — fall through to placeholder logic
                        raise RuntimeError("no client or PDF available")
                    figure_cache[q_num] = crop_figure_for_question(
                        client, qp_pdf_bytes, page_index, q_num,
                        q.get("figure_description", "the figure")
                    )
                figure_image = figure_cache[q_num]
            except Exception as e:
                if status_callback:
                    status_callback(f"⚠️ Couldn't extract figure for Q{q_num}: {e}")
                needs_manual = True

        # Insert figure (or placeholder) based on position
        def emit_figure_or_placeholder():
            if figure_image:
                add_image(doc, figure_image)
            elif needs_manual:
                add_figure_placeholder(doc, q_num)

        if (figure_image or needs_manual) and figure_position in ("before_body", "between_body_and_options"):
            emit_figure_or_placeholder()

        # Options
        for opt in q.get("options", []):
            add_option(doc, opt)

        # Figure after options
        if (figure_image or needs_manual) and figure_position == "after_options":
            emit_figure_or_placeholder()

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

# Mode toggle — PDF mode (default) extracts text and figures from QP via Claude.
# Excel-only mode bypasses the API entirely for faster/cheaper runs.
excel_only_mode = st.toggle(
    "📊 Use Excel sheet only (skip PDF + API, faster but figures need manual insertion)",
    value=False,
    help=(
        "When OFF (default), the app uses Claude to read your QP and MS PDFs to "
        "extract question text, options, and figures automatically. Figures are "
        "embedded in the document.\n\n"
        "When ON, the app reads questions, options, and answers directly from your "
        "spreadsheet — no Claude API calls, no token-limit errors, no PDF parsing. "
        "Faster and free, but figures get a placeholder box and you need to paste "
        "them in manually after downloading."
    ),
)

col1, col2 = st.columns(2)
with col1:
    qp_file = st.file_uploader(
        "Question Paper (PDF)" + (" — optional in Excel-only mode" if excel_only_mode else ""),
        type=["pdf"], key="qp",
    )
    xlsx_file = st.file_uploader("Classification Sheet (Excel)", type=["xlsx", "xls"], key="xlsx")
with col2:
    ms_file = st.file_uploader(
        "Mark Scheme (PDF)" + (" — optional in Excel-only mode" if excel_only_mode else ""),
        type=["pdf"], key="ms",
    )
    default_ref = st.text_input(
        "Default Reference (optional)",
        placeholder="e.g. (May 2021)",
        help="Used if the Excel doesn't include a reference column."
    )

# Generate button — different requirements per mode
if excel_only_mode:
    ready = bool(xlsx_file)
    if not ready:
        st.info("⏳ Excel-only mode: just upload the spreadsheet and click Generate.")
else:
    ready = bool(api_key and qp_file and ms_file and xlsx_file)

if st.button("🚀 Generate Worksheet", disabled=not ready, type="primary", use_container_width=True):
    try:
        client = anthropic.Anthropic(api_key=api_key) if api_key else None
        qp_bytes = qp_file.read() if qp_file else b""
        ms_bytes = ms_file.read() if ms_file else b""
        xlsx_bytes = xlsx_file.read()

        progress_box = st.empty()
        log_box = st.empty()
        log = []

        def update_status(msg):
            log.append(msg)
            log_box.info(msg)

        with st.spinner("Working..."):
            if excel_only_mode:
                # Fast path: extract everything directly from the spreadsheet.
                # No API calls, no PDF parsing, no token limits.
                update_status("📊 Extracting questions and answers from spreadsheet...")
                qp_name_for_filter = qp_file.name if qp_file else None
                questions, answers = extract_questions_from_excel(
                    xlsx_bytes,
                    qp_file_name=qp_name_for_filter,
                    status_callback=update_status,
                )

                if not questions and qp_name_for_filter:
                    # Try again without filename filter
                    update_status(
                        "⚠️ No rows matched the QP filename. Retrying without filename filter..."
                    )
                    questions, answers = extract_questions_from_excel(
                        xlsx_bytes,
                        qp_file_name=None,
                        status_callback=update_status,
                    )

                if not questions:
                    st.error(
                        "❌ Couldn't extract any questions from the spreadsheet. "
                        "Make sure it has a 'Question Text' column with the question content."
                    )
                    st.stop()

                update_status(f"✓ Got {len(questions)} questions and "
                              f"{sum(1 for v in answers.values() if not v.startswith('pos:'))//1} answers.")
            else:
                # PDF + API extraction path (original behavior).
                # Quick upfront check so the user sees what's about to happen
                try:
                    qp_pages = get_pdf_page_count(qp_bytes)
                    ms_pages = get_pdf_page_count(ms_bytes)
                    qp_total_tokens = estimate_pdf_tokens(qp_bytes)
                    ms_total_tokens = estimate_pdf_tokens(ms_bytes)

                    # Warn upfront if the input looks unusually large.
                    if qp_total_tokens > 800_000 or ms_total_tokens > 800_000:
                        st.warning(
                            f"⚠️ Your files are large: estimated ~{qp_total_tokens:,} tokens "
                            f"for the QP and ~{ms_total_tokens:,} tokens for the MS. "
                            f"Each is well over Claude's 200K-token limit, so the app will "
                            f"split them into many small chunks. **This will take longer "
                            f"and cost more API credits than usual.** If you'd rather not, "
                            f"cancel and re-upload smaller / lower-resolution PDFs."
                        )

                    qp_chunks_preview = build_size_aware_chunks(
                        qp_bytes, target_tokens=TARGET_TOKENS_PER_CHUNK
                    )
                    ms_chunks_preview = build_size_aware_chunks(
                        ms_bytes, target_tokens=TARGET_TOKENS_PER_CHUNK
                    )
                    update_status(
                        f"📊 Question Paper: {qp_pages} pages, ~{qp_total_tokens:,} tokens "
                        f"→ {len(qp_chunks_preview)} chunk(s). "
                        f"Mark Scheme: {ms_pages} pages, ~{ms_total_tokens:,} tokens "
                        f"→ {len(ms_chunks_preview)} chunk(s)."
                    )
                except InputTooLargeError:
                    raise  # Surface this with the standard handler below
                except Exception:
                    pass  # Non-fatal; continue regardless

                update_status("📖 Extracting questions from question paper...")
                questions = extract_questions(client, qp_bytes, status_callback=update_status)

                update_status(f"✓ Found {len(questions)} questions.")
                update_status("📖 Extracting answers from mark scheme...")
                answers = extract_answers(client, ms_bytes, status_callback=update_status)

                update_status(f"✓ Got {len(answers)} answers.")

            # Read classifications regardless of mode (Excel column lookup is the same).
            update_status("📖 Reading classification sheet...")
            classifications = read_classifications(
                xlsx_bytes,
                qp_file_name=qp_file.name if qp_file else None,
                status_callback=update_status,
            )

            # If filtering by File Name returned nothing, the spreadsheet is
            # either for a different QP, OR the user's spreadsheet doesn't have
            # a 'File Name' column / has it filled with different values. Try a
            # fallback that ignores file-name filtering and warn the user.
            if not classifications:
                fallback = read_classifications(
                    xlsx_bytes,
                    qp_file_name=None,
                    status_callback=update_status,
                )
                if fallback:
                    st.warning(
                        f"⚠️ **No spreadsheet rows matched the uploaded QP file name** "
                        f"`{qp_file.name}`.\n\n"
                        f"The spreadsheet has {len(fallback)} rows. The 'File Name' "
                        f"column values in the spreadsheet didn't match the uploaded "
                        f"QP's filename, so the chapter / difficulty / reference data "
                        f"can't be matched correctly.\n\n"
                        f"**The worksheet will still be generated**, but **without "
                        f"chapter, difficulty, or reference info** for each question.\n\n"
                        f"**To fix this:** make sure the spreadsheet's 'File Name' "
                        f"column contains a value that matches your QP's filename "
                        f"(spaces, underscores, and `.pdf` extension are handled "
                        f"automatically). For example, if your QP is "
                        f"`Chemistry Paper 1.pdf`, the spreadsheet should have "
                        f"`Chemistry Paper 1` (or similar) in the File Name column."
                    )
                    # Use the fallback so the user at least gets *something*.
                    classifications = fallback
                else:
                    st.error(
                        "❌ The spreadsheet appears to be empty or doesn't have "
                        "the expected columns (Topic, Difficulty, Reference)."
                    )

            update_status(
                f"✓ Got classifications for {len(classifications)} questions "
                f"(matched to '{qp_file.name}')."
            )

            # Show a quick breakdown of how questions will be organized
            from collections import Counter
            chapter_counts = Counter()
            for i, q in enumerate(questions):
                position = q.get("position", i + 1)
                cls = classifications.get(position, {})
                ch = (cls.get("chapter") or "—").strip() or "Uncategorized"
                chapter_counts[ch] += 1
            if chapter_counts:
                breakdown = " · ".join(
                    f"{ch} ({n})" for ch, n in sorted(
                        chapter_counts.items(),
                        key=lambda x: chapter_sort_key(x[0])
                    )
                )
                update_status(f"📚 Chapter breakdown: {breakdown}")
                update_status(
                    "📑 Worksheet will be sorted by Chapter → Difficulty (Easy→Medium→Hard) → Year."
                )

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

    except InputTooLargeError as e:
        st.error("❌ **The input is too large for Claude to process.**")
        st.warning(str(e))
        st.info(
            "💡 **What this means and how to fix it:**\n\n"
            "Claude has a 200,000-token limit per request. The app already splits "
            "your PDFs into small chunks before sending them, so this error usually "
            "means a **single page** of your PDF is too dense (very high-resolution "
            "scans or many embedded images on one page).\n\n"
            "**Try one of these:**\n\n"
            "1. **Reduce PDF image quality.** Open your PDF in any tool that has "
            "a 'Reduce file size' or 'Optimize PDF' option (Adobe Acrobat, "
            "Smallpdf, ILovePDF, PDF24). Re-export at lower resolution and re-upload.\n"
            "2. **Split the PDF into smaller files.** Upload one exam at a time "
            "instead of a combined file containing multiple papers.\n"
            "3. **Convert to text-based PDF.** If your PDF is a scanned image, "
            "run it through OCR first (most PDF tools have a 'Make searchable / OCR' "
            "option). Text-based PDFs use ~10x fewer tokens than scanned ones."
        )
    except anthropic.AuthenticationError:
        st.error("❌ Invalid API key. Please check your Anthropic API key.")
    except anthropic.BadRequestError as e:
        # Catch any size errors that slipped past our adaptive splitter
        if is_too_large_error(e):
            st.error("❌ **The input is too large for Claude to process.**")
            st.warning(
                "Even after splitting your PDF into single pages and falling back "
                "to image-based extraction, Claude rejected the request. This is rare.\n\n"
                "**Most likely cause:** your PDF has very high-resolution scanned content. "
                "Try re-saving the PDF with reduced image quality (any PDF tool that has "
                "a 'Reduce file size' or 'Optimize PDF' option will work) and upload it again."
            )
        else:
            st.error(f"❌ API error: {e}")
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
