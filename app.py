"""
IB Exam Worksheet Generator
- QP PDF  → questions, options, tables, diagrams, graphs (verbatim)
- MS PDF  → answers only
- Excel   → question numbers, chapter, difficulty, marks, quotes
"""
import streamlit as st
import anthropic
import openpyxl
import json
import re
import io
import base64
import datetime

import fitz                           # PyMuPDF — PDF inspection + rendering
from PIL import Image
import numpy as np

from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.oxml.ns import qn as docx_qn
from docx.oxml import OxmlElement


# ═══════════════════════════════════════════════════════════════════════════════
#  STREAMLIT UI
# ═══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Exam Worksheet Generator",
    page_icon="📄",
    layout="centered",
)

st.title("📄 Exam Worksheet Generator")
st.caption("QP PDF → questions · MS PDF → answers · Excel → metadata")

with st.expander("ℹ️ Source rules", expanded=False):
    st.markdown(
        "| Source | Used for |\n"
        "|--------|----------|\n"
        "| **QP PDF** | Question text, options, tables, diagrams, graphs — verbatim |\n"
        "| **MS PDF** | Answers only |\n"
        "| **Excel** | Question numbers, chapter, difficulty, marks, quotes |\n\n"
        "- Visual questions (graphs/diagrams/tables) → cropped image from QP, no duplicate text\n"
        "- Excel `Topic = Error` or empty → `Unclassified`\n"
        "- Excel `Marks = 0` or missing → `1` (standard MCQ)"
    )

c1, c2 = st.columns(2)
with c1:
    qp_file = st.file_uploader("📋 Question Paper (QP)", type=["pdf"], key="qp")
    ms_file = st.file_uploader("✅ Mark Scheme (MS)",     type=["pdf"], key="ms")
with c2:
    xl_file = st.file_uploader("📊 Excel Sheet", type=["xlsx", "xls"], key="xl")

st.divider()

# ── Mode selector ──────────────────────────────────────────────────────────
mode = st.radio(
    "Processing mode",
    options=["MCQ Mode", "Structured Mark Scheme Mode"],
    horizontal=True,
    key="mode",
    help=(
        "**MCQ Mode** — for Chemistry / Biology Paper 1 style. Mark Scheme is "
        "a grid of A/B/C/D answers.\n\n"
        "**Structured Mark Scheme Mode** — for Mathematics or any subject with "
        "full worked solutions (M1/A1/R1 marks, steps, diagrams). The Mark "
        "Scheme is extracted as a cropped image from the MS PDF for each "
        "question."
    ),
)

st.divider()


# ═══════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
def read_bytes(uploaded_file) -> bytes:
    if uploaded_file is None:
        return b""
    if isinstance(uploaded_file, bytes):
        return uploaded_file
    if hasattr(uploaded_file, "getvalue"):
        return uploaded_file.getvalue()
    if hasattr(uploaded_file, "seek"):
        try:
            uploaded_file.seek(0)
        except Exception:
            pass
    if hasattr(uploaded_file, "read"):
        return uploaded_file.read()
    raise TypeError(f"Unsupported uploaded file type: {type(uploaded_file)}")


def to_b64(data: bytes) -> str:
    return base64.standard_b64encode(data).decode()


def safe_json(text: str):
    text = re.sub(r"```json\s*", "", text)
    text = re.sub(r"```\s*", "", text)
    text = text.strip()
    # Try direct parse, else find the first JSON-looking substring
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        for opener, closer in [("[", "]"), ("{", "}")]:
            i = text.find(opener)
            j = text.rfind(closer)
            if i != -1 and j != -1 and j > i:
                try:
                    return json.loads(text[i:j + 1])
                except json.JSONDecodeError:
                    continue
        return None


# ───────────────────────────────────────────────────────────────────────────────
#  EXCEL  →  metadata only
# ───────────────────────────────────────────────────────────────────────────────
def parse_excel(uploaded_file) -> list[dict]:
    """Read question numbers + metadata from Excel.
    Never used for question text or answers.
    CHANGE 3: always seeks to 0 — no cached reads.
    CHANGE 4: Paper 2 rows are silently skipped.
    CHANGE 5: 'October' is normalised to 'November' in Reference strings.
    CHANGE 6: Date-encoded multi-page refs (e.g. 2026-06-07 → '6-7') kept.
    CHANGE 7: marks=0 is kept as-is (will be corrected from QP later).
    """
    # CHANGE 3: ensure clean read, no stale buffer
    uploaded_file.seek(0)
    wb = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)
    ws = wb.active
    raw_headers = [c.value for c in next(ws.iter_rows(max_row=1))]
    headers = [str(h).strip() if h else "" for h in raw_headers]
    # Lower-cased + normalised version for fuzzy matching
    norm_headers = [re.sub(r"[\s_\-\.]+", " ", h.lower()).strip()
                    for h in headers]

    def col_idx(*candidates):
        """Find a column index by trying each candidate name.
        Match is case-insensitive and ignores spaces/underscores/dots/dashes.
        """
        for name in candidates:
            target = re.sub(r"[\s_\-\.]+", " ", str(name).lower()).strip()
            for i, h in enumerate(norm_headers):
                if h == target:
                    return i
        # Substring fallback (handles 'PDF Page Number' matching 'Page Number')
        for name in candidates:
            target = re.sub(r"[\s_\-\.]+", " ", str(name).lower()).strip()
            for i, h in enumerate(norm_headers):
                if target in h or h in target:
                    return i
        return None

    qn_idx   = col_idx("Question No.", "Question No", "Q No", "Q#",
                       "Question Number", "Q", "QuestionNumber")
    page_idx = col_idx("Page Number", "Page", "page_number",
                       "PDF Page", "PDF Page Number", "pdf_page",
                       "QP Page", "Page No.", "Page No", "PageNum")
    top_idx  = col_idx("Topic", "Chapter", "Section", "Sub-topic", "Subtopic")
    dif_idx  = col_idx("Difficulty", "Level", "Hardness")
    mrk_idx  = col_idx("Marks", "Mark", "Points", "Score")
    qut_idx  = col_idx("Quote", "Note", "Motivational quote")
    ref_idx  = col_idx("Reference", "Ref", "Session", "Year", "Paper",
                       "Exam Session", "Paper Reference")

    # Diagnostics — stash on the function so the UI can show them
    parse_excel.last_headers = headers
    parse_excel.last_column_map = {
        "Question No.": (qn_idx,  headers[qn_idx]   if qn_idx   is not None else None),
        "Page Number":  (page_idx, headers[page_idx] if page_idx is not None else None),
        "Topic":        (top_idx,  headers[top_idx]  if top_idx  is not None else None),
        "Difficulty":   (dif_idx,  headers[dif_idx]  if dif_idx  is not None else None),
        "Marks":        (mrk_idx,  headers[mrk_idx]  if mrk_idx  is not None else None),
        "Reference":    (ref_idx,  headers[ref_idx]  if ref_idx  is not None else None),
        "Quote":        (qut_idx,  headers[qut_idx]  if qut_idx  is not None else None),
    }

    # Fall back to fixed-position 2 only if NOTHING matched (very old format)
    if qn_idx is None:
        qn_idx = 2

    def cell_val(row, idx, fallback=""):
        if idx is None:
            return fallback
        try:
            v = list(row)[idx].value
            return str(v).strip() if v is not None else fallback
        except IndexError:
            return fallback

    rows = []
    seen_keys = set()
    duplicates = []

    for row in ws.iter_rows(min_row=2):
        try:
            q_num = int(float(str(list(row)[qn_idx].value)))
        except (TypeError, ValueError):
            continue
        if q_num <= 0:
            continue

        topic = cell_val(row, top_idx)
        # CHANGE 7: only mark as Unclassified when truly missing/error —
        # do NOT override a valid chapter name like "Chapter 1 : Number & Algebra"
        if not topic or topic.strip().lower() in ("error", "none", "n/a", "nan"):
            topic = "Unclassified"

        try:
            marks = int(float(cell_val(row, mrk_idx, "0")))
        except ValueError:
            marks = 0
        # CHANGE 7: do NOT override marks=0 → keep it so the merge step
        # can detect "[Maximum mark: N]" from the QP and correct it.
        if marks < 0:
            marks = 0

        # Page number — handle:
        #   - integers:   91
        #   - strings:    "91", "p.91", "Page 91"
        #   - ranges:     "82-83", "p.82-83"  → take first number
        #   - datetimes:  Excel auto-converts "2-3" → 2026-02-03 → BAD!
        #                 We try to recover from .number_format or skip.
        raw_page = ""
        if page_idx is not None:
            try:
                cell = list(row)[page_idx]
                v = cell.value
                if v is None:
                    raw_page = ""
                elif isinstance(v, datetime.datetime):
                    # Excel converted "X-Y" → date. Try to recover X and Y
                    # from the month and day.
                    raw_page = f"{v.month}-{v.day}"
                else:
                    raw_page = str(v).strip()
            except (IndexError, AttributeError):
                raw_page = ""

        page_num = 0
        if raw_page:
            # Find first integer in the string (handles "82-83" → 82,
            # "p.91" → 91, "Page 91 (left)" → 91)
            m = re.search(r"\d+", raw_page)
            if m:
                try:
                    candidate = int(m.group())
                    # Reject obviously-wrong values (Excel YYYY conversions)
                    if 1 <= candidate <= 9999:
                        page_num = candidate
                except ValueError:
                    page_num = 0

        difficulty = cell_val(row, dif_idx, "Unspecified") or "Unspecified"
        ref        = cell_val(row, ref_idx, "")
        quote      = cell_val(row, qut_idx, "")

        # CHANGE 5: normalise October → November (IB has no October session;
        # some Excel files label the November session as "October")
        ref = re.sub(r'\boctober\b', 'November', ref, flags=re.IGNORECASE)
        ref = re.sub(r'\boct\b',     'Nov',      ref, flags=re.IGNORECASE)

        # CHANGE 4: skip Paper 2 rows entirely.
        # P2 paper codes contain: 7205, 7210, 7215, 7310, 7315, 7320
        # Also skip explicit "Paper 2" / "HP2" / "SP2" references.
        _P2_CODES = ('7205', '7210', '7215', '7310', '7315', '7320')
        if any(code in ref for code in _P2_CODES):
            continue
        if re.search(r'\bpaper[\s\-]?2\b|[/\-_](?:h|s)p2[/\-_]', ref,
                     re.IGNORECASE):
            continue

        # ── Deduplication: only when EVERY identifying field is identical ───
        # Q1 from different references / pages / topics / difficulties are
        # NOT duplicates — they're legitimately distinct questions.
        # We only skip a row when (Q#, Reference, Page, Topic, Difficulty,
        # Marks, Quote) are all the exact same as a previously-seen row.
        dedup_key = (
            q_num,
            ref.strip().lower(),
            page_num,
            topic.strip().lower(),
            difficulty.strip().lower(),
            marks,
            quote.strip().lower(),
        )
        if dedup_key in seen_keys:
            duplicates.append((q_num, ref))
            continue
        seen_keys.add(dedup_key)

        rows.append({
            "qn":         q_num,
            "page_num":   page_num,
            "topic":      topic,
            "difficulty": difficulty,
            "marks":      marks,
            "quote":      quote,
            "ref":        ref,
        })

    parse_excel.last_duplicates = duplicates
    return rows


# ───────────────────────────────────────────────────────────────────────────────
#  PYMUPDF  →  find EVERY occurrence of every "N." marker in the PDF
# ───────────────────────────────────────────────────────────────────────────────
def find_question_locations(pdf_bytes: bytes) -> dict:
    """Scan all pages of the QP PDF and locate EVERY question marker.

    Critical: when the QP PDF contains multiple papers concatenated, the
    same question number (e.g. Q1) appears once per paper. We MUST keep
    every occurrence — not just the first — so that callers can pick the
    right Q1 by page number.

    Returns a dict keyed by (page_idx_1based, q_num):
        {(page_idx, q_num): {"page_idx": int,    # 0-based
                             "top_y": float,
                             "bottom_y": float,
                             "page_height": float,
                             "page_width": float,
                             "marker_right_x": float}}
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    raw = []   # (q_num, page_idx_0based, top_y, page_h, page_w, marker_right_x)

    pat = re.compile(r"^(\d+)\.$")

    for page_idx in range(len(doc)):
        page = doc[page_idx]
        ph, pw = page.rect.height, page.rect.width
        text_dict = page.get_text("dict")

        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue

                first_text = spans[0].get("text", "").strip()
                bbox = line["bbox"]

                m = pat.match(first_text)
                if not m:
                    continue
                if bbox[0] > 110:
                    continue

                q_num = int(m.group(1))
                if q_num < 1 or q_num > 99:
                    continue

                marker_right_x = spans[0]["bbox"][2]
                raw.append((q_num, page_idx, bbox[1], ph, pw, marker_right_x))

    doc.close()

    # Sort by document order so we can compute each marker's bottom_y
    raw.sort(key=lambda e: (e[1], e[2]))

    # Pre-compute per-page "content bottom" — the y of the last real content
    # line before any footer (page number / copyright / "Turn over").
    # This lets the LAST question on each page extend down to actual content
    # rather than being clipped at an arbitrary fraction of page height.
    doc2 = fitz.open(stream=pdf_bytes, filetype="pdf")
    content_bottom_per_page: dict = {}

    # Footer detection patterns — IB exam footers vary but include:
    #   - "International Baccalaureate Organization 2022"
    #   - "Turn over"
    #   - "References:" section header (appears in May 2022 P2)
    #   - Bare paper-number codes: "8821 – 6101", "2221 – 6113"
    # NOTE: bare page numbers like "137" or "204" need extra y-check
    # (they only count as footer when in the bottom strip of the page).
    footer_pat = re.compile(
        r"(international\s+baccalaureate|turn\s+over|references?\s*:|"
        r"^\s*\d{4}\s*[\-\u2013]\s*\d{3,4}\s*$|"
        r"^\s*[\-\u2013]\s*\d{1,3}\s*[\-\u2013]\s*$)",
        re.IGNORECASE
    )
    # Bare page-number pattern (only counts as footer in bottom strip)
    bare_num_pat = re.compile(r"^\s*\d{1,3}\s*$")

    # IB QP "writing lines" — dotted answer-box lines. The PDF contains
    # text like ". . . . . . . . . . . . . . . . . . . ." for each blank line.
    # We must NOT include these in the question crop (the worksheet has its
    # own Student's Solution box). A line counts as a writing line when ≥ 80%
    # of its non-space chars are dots.
    def is_writing_line(txt: str) -> bool:
        s = (txt or "").replace(" ", "").replace("\t", "")
        if len(s) < 8:
            return False
        dots = s.count(".")
        return dots / max(len(s), 1) >= 0.8

    for pi in range(len(doc2)):
        page = doc2[pi]
        ph = page.rect.height
        td = page.get_text("dict")
        # Collect every line's (y0, y1, text)
        lines_info = []
        for block in td.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue
                full = " ".join(s.get("text", "") for s in spans).strip()
                if not full:
                    continue
                bb = line["bbox"]
                lines_info.append((bb[1], bb[3], full))   # (y0, y1, text)

        def is_footer(y0, txt):
            """Context-aware footer detection."""
            if footer_pat.search(txt):
                return True
            # Bare numbers count as footer ONLY in bottom 10% of page
            if bare_num_pat.match(txt) and y0 > ph * 0.90:
                return True
            return False

        # Find the FIRST footer line in the lower half of the page.
        first_footer_y0 = ph * 0.92
        for y0, y1, txt in sorted(lines_info, key=lambda t: t[0]):
            if y0 < ph * 0.5:
                continue
            if is_footer(y0, txt):
                first_footer_y0 = min(first_footer_y0, y0)
                break

        # Find the FIRST writing line — the answer-box dotted lines.
        # Everything below this is the answer box and must be excluded
        # from the question crop. Writing lines can appear above 0.5 too
        # (short questions on a half-page), so we don't restrict by y here.
        first_writing_y0 = None
        for y0, y1, txt in sorted(lines_info, key=lambda t: t[0]):
            if is_writing_line(txt):
                first_writing_y0 = y0
                break

        # Effective cutoff: the earlier of (footer top, writing-line top)
        upper_bound = first_footer_y0
        if first_writing_y0 is not None and first_writing_y0 < upper_bound:
            upper_bound = first_writing_y0

        # Now find the last content line above that boundary.
        # Writing lines and footers are NOT content.
        content_bottom = 0
        for y0, y1, txt in lines_info:
            if is_footer(y0, txt) or is_writing_line(txt):
                continue
            if y1 < upper_bound - 2:
                if y1 > content_bottom:
                    content_bottom = y1

        # If we never found any content, fall back to a safe ratio
        if content_bottom == 0:
            content_bottom = ph * 0.88

        content_bottom_per_page[pi] = min(
            content_bottom + 8,
            upper_bound - 6,
        )

    doc2.close()

    result = {}
    for i, (qn, page_idx, top_y, ph, pw, mrx) in enumerate(raw):
        # Default bottom_y = the page's actual content bottom (not a fixed %)
        bottom_y = content_bottom_per_page.get(page_idx, ph * 0.88)
        # Hard ceiling at 0.92 to keep IB footer out even if detector missed
        bottom_y = min(bottom_y, ph * 0.92)

        # If there's a NEXT marker on the same page, bottom = that marker
        for j in range(i + 1, len(raw)):
            nq, np_idx, ntop, _, _, _ = raw[j]
            if np_idx == page_idx:
                bottom_y = ntop - 4
                break
            if np_idx > page_idx:
                break

        # Use 1-based page number for the key (matches Excel "Page Number")
        page_1based = page_idx + 1
        key = (page_1based, qn)

        # If the same (page, qn) appears twice (extremely rare), keep first
        if key in result:
            continue

        result[key] = {
            "page_idx":       page_idx,
            "top_y":          top_y,
            "bottom_y":       bottom_y,
            "page_height":    ph,
            "page_width":     pw,
            "marker_right_x": mrx,
        }

    return result


def find_question_for_excel_row(
    locations: dict,           # {(page, qn): {...}}
    q_num: int,
    excel_page: int,
    tolerance: int = 5,        # tolerate ±N pages of mismatch
) -> dict | None:
    """Pick the right `locations` entry for an Excel row.

    Strategy:
      1) Exact match on (excel_page, q_num)
      2) Within ±tolerance pages of excel_page — pick the nearest
      3) Otherwise → None (caller decides: error / fallback)
    """
    if excel_page and excel_page > 0:
        # Try exact match first
        if (excel_page, q_num) in locations:
            return locations[(excel_page, q_num)]

        # Try nearby pages (text-layer offsets, page-number mismatches)
        candidates = [
            (abs(p - excel_page), p)
            for (p, qn) in locations
            if qn == q_num and abs(p - excel_page) <= tolerance
        ]
        if candidates:
            candidates.sort()  # nearest first
            _, p = candidates[0]
            return locations[(p, q_num)]

        # No match within tolerance → caller handles
        return None

    # No page hint in Excel — best-effort: use first occurrence of this qn
    for (p, qn), loc in sorted(locations.items()):
        if qn == q_num:
            return loc
    return None


def find_question_with_segment_offset(
    locations: dict,
    q_num: int,
    excel_page: int,
    segment_offset: int = 0,
) -> dict | None:
    """Like find_question_for_excel_row, but applies a learned per-segment
    page offset before searching. If excel_page=16 and offset=-4, we look
    for the question near PDF page 12.

    Falls back to plain matching if offset-based lookup fails.
    """
    if excel_page and excel_page > 0:
        # Try shifted page first
        shifted = excel_page + segment_offset
        if shifted > 0:
            if (shifted, q_num) in locations:
                return locations[(shifted, q_num)]
            # Wider tolerance after offset
            candidates = [
                (abs(p - shifted), p)
                for (p, qn) in locations
                if qn == q_num and abs(p - shifted) <= 5
            ]
            if candidates:
                candidates.sort()
                _, p = candidates[0]
                return locations[(p, q_num)]

    # Fall back to plain matching (no offset)
    return find_question_for_excel_row(locations, q_num, excel_page)


def infer_segment_page_offsets(
    locations: dict,
    xl_rows: list,
    segments: list,
) -> list:
    """For each Excel paper-segment, infer the offset between Excel page
    numbers and QP PDF page numbers, by looking at Q1 positions.

    Returns a list of offsets, one per segment.
    """
    offsets = []
    for seg_rows in segments:
        seg_offset = 0   # default
        # Find Q1 in this segment
        q1_rows = [xl_rows[ri] for ri in seg_rows if xl_rows[ri]["qn"] == 1]
        if q1_rows:
            excel_q1_page = q1_rows[0]["page_num"]
            if excel_q1_page > 0:
                # Find nearest QP Q1 location
                q1_qp_pages = sorted(p for (p, qn) in locations if qn == 1)
                if q1_qp_pages:
                    # Pick the QP Q1 closest to excel_q1_page
                    nearest = min(q1_qp_pages,
                                  key=lambda p: abs(p - excel_q1_page))
                    # Only use offset if within reason (< 20 pages off)
                    if abs(nearest - excel_q1_page) < 20:
                        seg_offset = nearest - excel_q1_page
        offsets.append(seg_offset)
    return offsets


def crop_question_png(pdf_bytes: bytes, loc: dict, q_num: int = None,
                      dpi: int = 200) -> bytes | None:
    """Crop the question area from the QP PDF and return PNG bytes.
       `loc` is now a single location dict (from find_question_for_excel_row),
       NOT the full locations dict.
    """
    if not loc:
        return None

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[loc["page_idx"]]
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    iw, ih = img.size
    doc.close()

    scale = dpi / 72
    top_px    = max(0,  int(loc["top_y"]    * scale) - 10)
    bottom_px = min(ih, int(loc["bottom_y"] * scale) + 5)

    # Skip the question-number column ("5.") so the marker doesn't
    # appear inside the embedded image — it's already in the heading.
    mrx = loc.get("marker_right_x", 0)
    left_px = max(0, int((mrx + 6) * scale)) if mrx else 0
    # Trim the right margin — IB papers have:
    #   - a vertical page-border line around x ≈ 0.93 × page_width
    #   - sometimes a writing/notes column inside that border
    # Cropping at 0.92 × page_width removes the border but keeps "[2]" marks.
    pw = loc.get("page_width", 0) or (iw / scale)
    right_px = iw - int(pw * 0.05 * scale)

    if bottom_px <= top_px or right_px <= left_px:
        return None

    cropped = img.crop((left_px, top_px, right_px, bottom_px))
    buf = io.BytesIO()
    cropped.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


# ───────────────────────────────────────────────────────────────────────────────
#  CLAUDE  →  classify (visual?) + extract text for text-only questions
# ───────────────────────────────────────────────────────────────────────────────
def classify_and_extract(client: anthropic.Anthropic, qp_b64: str,
                         q_nums_or_rows,
                         qp_bytes: bytes = None,
                         locations: dict = None) -> list[dict]:
    """Deterministic extraction using PyMuPDF — no Claude calls, no placeholders.

    `q_nums_or_rows` may be:
      - a list of ints — legacy single-paper mode, will lookup the first
        occurrence of each q_num in locations
      - a list of (q_num, excel_page) tuples — preferred for multi-paper:
        each Excel row is matched to its specific page in the PDF
    """
    if qp_bytes is None or locations is None:
        return []

    doc = fitz.open(stream=qp_bytes, filetype="pdf")
    results = []

    for entry in q_nums_or_rows:
        if isinstance(entry, tuple):
            q_num, excel_page = entry
        else:
            q_num, excel_page = entry, 0

        loc = find_question_for_excel_row(locations, q_num, excel_page)
        if not loc:
            results.append({
                "qn": q_num, "found": False, "needs_image": False,
                "text": "", "A": "", "B": "", "C": "", "D": "",
                "_loc": None, "_excel_page": excel_page,
            })
            continue

        page = doc[loc["page_idx"]]
        top_y, bottom_y = loc["top_y"], loc["bottom_y"]
        clip = fitz.Rect(0, top_y, page.rect.width, bottom_y)

        # ── Collect spans inside the question's bbox ─────────────────────────
        # We work at the SPAN level (not the line level) so we can detect
        # subscripts / superscripts that are encoded as smaller-font spans
        # whose baseline is above or below the main text baseline.
        text_dict = page.get_text("dict", clip=clip)

        # First pass: gather all spans + figure out the dominant body font size
        all_spans: list[dict] = []
        size_counts: dict = {}
        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = span.get("text", "")
                    if not text:
                        continue
                    sz = round(span.get("size", 0), 1)
                    bb = span.get("bbox", (0, 0, 0, 0))
                    all_spans.append({
                        "text": text,
                        "size": sz,
                        "x0":   bb[0],
                        "y0":   bb[1],
                        "y1":   bb[3],
                        "flags": span.get("flags", 0),
                    })
                    if text.strip():
                        size_counts[sz] = size_counts.get(sz, 0) + len(text.strip())

        if not all_spans:
            results.append({
                "qn": q_num, "found": True, "needs_image": True,
                "text": "", "A": "", "B": "", "C": "", "D": "",
            })
            continue

        body_size = max(size_counts.items(), key=lambda kv: kv[1])[0] \
                    if size_counts else 11.0

        # Unicode tables for safe sub/superscript conversion
        SUBS  = str.maketrans("0123456789+-=()n",
                              "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₙ")
        SUPS  = str.maketrans("0123456789+-=()n",
                              "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿ")

        # Track whether any sub/superscript span had characters we could
        # NOT translate safely (e.g. letters other than 'n' in subscripts,
        # mixed content). If yes, force image fallback to avoid distortion.
        unsafe_sub_super = False

        # ── Group spans by text line (y bucket) ──────────────────────────────
        Y_BUCKET = 3
        line_buckets: dict = {}
        for sp in all_spans:
            key = round(sp["y0"] / Y_BUCKET) * Y_BUCKET
            line_buckets.setdefault(key, []).append(sp)

        # Build full-line text WITH subscript/superscript conversion
        rendered_lines: list[tuple] = []   # (y, x0, rendered_text)
        for key in sorted(line_buckets):
            spans_on_line = sorted(line_buckets[key], key=lambda s: s["x0"])
            # Identify the dominant baseline on this line
            full_size_spans = [s for s in spans_on_line
                               if s["size"] >= body_size - 0.5]
            if full_size_spans:
                main_y0 = full_size_spans[0]["y0"]
                main_y1 = full_size_spans[0]["y1"]
            else:
                main_y0 = spans_on_line[0]["y0"]
                main_y1 = spans_on_line[0]["y1"]

            parts = []
            for sp in spans_on_line:
                text = sp["text"]
                if not text.strip():
                    parts.append(text)
                    continue
                is_smaller = sp["size"] < body_size - 1.0
                # Subscript: smaller font, baseline below main
                if is_smaller and sp["y1"] > main_y1 + 0.5:
                    if all(c in "0123456789+-=()n" for c in text.strip()):
                        parts.append(text.translate(SUBS))
                    else:
                        unsafe_sub_super = True
                        parts.append(text)
                # Superscript: smaller font, baseline above main
                elif is_smaller and sp["y0"] < main_y0 - 0.5:
                    if all(c in "0123456789+-=()n" for c in text.strip()):
                        parts.append(text.translate(SUPS))
                    else:
                        unsafe_sub_super = True
                        parts.append(text)
                else:
                    parts.append(text)

            rendered = "".join(parts)
            # Normalise weird whitespace
            rendered = re.sub(r"[\u2009\u00a0]", " ", rendered).strip()
            if rendered:
                rendered_lines.append((key, spans_on_line[0]["x0"], rendered))

        # ── Detect option markers (A. / B. / C. / D.) ────────────────────────
        opt_re = re.compile(r"^\s*([ABCD])\.\s*(.*)$")
        opt_markers: dict = {}
        for y, x0, text in rendered_lines:
            m = opt_re.match(text)
            if m and m.group(1) not in opt_markers:
                opt_markers[m.group(1)] = (y, m.group(2).strip())

        # ── Build stem ───────────────────────────────────────────────────────
        stem_lines: list[str] = []
        first_opt_y = (min(y for y, _ in opt_markers.values())
                       if opt_markers else float("inf"))
        qn_strip_re = re.compile(rf"^\s*{q_num}\.\s+(.*)$")
        qn_alone_re = re.compile(rf"^\s*{q_num}\.\s*$")
        for y, x0, text in rendered_lines:
            if y >= first_opt_y:
                break
            if not stem_lines:
                m_strip = qn_strip_re.match(text)
                if m_strip:
                    rest = m_strip.group(1).strip()
                    if rest:
                        stem_lines.append(rest)
                    continue
                if qn_alone_re.match(text):
                    continue
            stem_lines.append(text)
        stem = "\n".join(stem_lines).strip()

        # ── Build option texts ───────────────────────────────────────────────
        options = {"A": "", "B": "", "C": "", "D": ""}
        sorted_opts = sorted(opt_markers.items(), key=lambda kv: kv[1][0])
        for i, (letter, (start_y, init)) in enumerate(sorted_opts):
            end_y = (sorted_opts[i + 1][1][0]
                     if i + 1 < len(sorted_opts) else float("inf"))
            parts = [init] if init else []
            for y, x0, text in rendered_lines:
                if y > start_y and y < end_y and not opt_re.match(text):
                    parts.append(text)
            options[letter] = " ".join(parts).strip()

        # ── Detect visual content (drawings / images inside bbox) ────────────
        needs_image = False
        n_drawings_in_bbox = 0
        big_draw_area = 0.0

        try:
            drawings = page.get_drawings()
            for d in drawings:
                r = d.get("rect")
                if r is None:
                    continue
                if r.y1 < top_y or r.y0 > bottom_y:
                    continue
                n_drawings_in_bbox += 1
                w, h = r.width, r.height
                if w >= 4 and h >= 4:
                    big_draw_area += w * h
        except Exception:
            pass

        if big_draw_area > 1500:
            needs_image = True
        if not needs_image and n_drawings_in_bbox >= 15:
            needs_image = True

        if not needs_image:
            try:
                for img in page.get_image_info(xrefs=True):
                    bb = img.get("bbox")
                    if not bb:
                        continue
                    if bb[3] < top_y or bb[1] > bottom_y:
                        continue
                    if (bb[2] - bb[0]) * (bb[3] - bb[1]) > 200:
                        needs_image = True
                        break
            except Exception:
                pass

        # ── Force image when text accuracy is at risk ────────────────────────
        # Any sub/superscript content that we couldn't translate safely:
        if unsafe_sub_super:
            needs_image = True

        def _has_chem_complexity(s: str) -> bool:
            """True if the text contains any chemistry/math notation that
            risks distortion in plain text. Strict: when in doubt → image.
            """
            if not s:
                return False

            # ANY Unicode subscript or superscript present
            if re.search(r"[₀-₉₊₋₌₍₎ₙ⁰-⁹⁺⁻⁼⁽⁾ⁿ]", s):
                return True

            # Any element-and-digit chemistry formula in the option text:
            # "C₂₀H₃₀O" extracted as "CHO 2030" or "C 20H30O" or "C20H30O"
            # Heuristic: 2+ uppercase letters in a row (=multiple elements)
            # followed by anything — that's a formula and may be corrupted.
            if re.search(r"[A-Z]{2,}", s):
                # If it's an all-caps acronym in plain English ("WHO", "DNA"),
                # length is bounded and there's no digit nearby. Block when
                # there's a digit anywhere in the string.
                if re.search(r"\d", s):
                    return True

            # Letter directly followed by a digit:
            # C2, NH4, BF3, 10⁻⁹
            if re.search(r"[A-Z][a-z]?\d", s):
                return True

            # Letter + space + digit (corrupted formula ordering):
            # "C 20", "NH 4", "BF 3"
            if re.search(r"\b[A-Z][a-z]?\s+\d", s):
                return True

            # Multi-element formula stretches without separators:
            # "CHCHCHCHCHCHOHCHCOCH"
            if re.search(r"[A-Z][A-Z][A-Z]{4,}", s):
                return True

            # Arithmetic operators in mathematical context
            if re.search(r"[=×]", s):
                return True

            # Trailing minus signs glued at end (corrupted superscripts)
            if re.search(r"\s[-+]+\s*$", s):
                return True

            # Bare double-element word missing subscripts ("BF", "PO" alone)
            if re.fullmatch(r"\s*[A-Z][a-z]?[A-Z][a-z]?\s*", s):
                return True

            return False

        if not needs_image:
            if _has_chem_complexity(stem):
                needs_image = True
            else:
                for L in "ABCD":
                    if _has_chem_complexity(options[L]):
                        needs_image = True
                        break

        # Final fallback: incomplete extraction
        non_empty_opts = [v for v in options.values()
                          if v and len(v.strip()) >= 1]
        if len(non_empty_opts) < 4 or not stem:
            needs_image = True

        if needs_image:
            stem = ""
            options = {"A": "", "B": "", "C": "", "D": ""}

        results.append({
            "qn":          q_num,
            "found":       True,
            "needs_image": needs_image,
            "text":        stem,
            "A": options["A"], "B": options["B"],
            "C": options["C"], "D": options["D"],
            "_loc":        loc,                # picked location dict
            "_excel_page": excel_page,
        })

    doc.close()
    return results


def _norm_for_match(s: str) -> str:
    """Lower-case, strip whitespace/punctuation for fuzzy filename↔reference matching."""
    if not s:
        return ""
    s = s.lower()
    # CHANGE 5: October = November in IB calendar
    s = re.sub(r'\boctober\b', 'november', s)
    s = re.sub(r'\boct\b',     'nov',      s)
    s = re.sub(r"[\(\)\[\]\{\}_\-,\.]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def match_papers_to_references(
    pdf_files: list,                      # list of Streamlit UploadedFile
    refs: list[str],                      # unique Reference strings from Excel
) -> dict[str, dict]:
    """Match each Excel reference to one of the uploaded PDF files.

    Returns: {reference_string: {"name": filename, "bytes": pdf_bytes,
                                  "matched": bool}}

    Matching strategy (in order):
      1. Exact / substring match between reference and filename
      2. Token-overlap scoring (best filename per reference)
      3. If only one PDF was uploaded → all references map to it
    """
    if not pdf_files:
        return {}

    # Cache file bytes by name
    files_data = {}
    for f in pdf_files:
        try:
            f.seek(0)
            files_data[f.name] = f.read()
        except Exception:
            continue

    if not files_data:
        return {}

    # Single-file shortcut: all references map to the only PDF
    if len(files_data) == 1:
        only_name, only_bytes = next(iter(files_data.items()))
        return {ref: {"name": only_name, "bytes": only_bytes, "matched": True}
                for ref in refs}

    # Multi-file: token-overlap scoring
    norm_files = {n: set(_norm_for_match(n).split()) for n in files_data}

    result = {}
    for ref in refs:
        ref_tokens = set(_norm_for_match(ref).split())
        if not ref_tokens:
            # Fallback to first file
            n = next(iter(files_data))
            result[ref] = {"name": n, "bytes": files_data[n], "matched": False}
            continue

        # Score each file by intersection size (Jaccard-ish)
        best_name, best_score = None, 0
        for fname, ftokens in norm_files.items():
            score = len(ref_tokens & ftokens)
            if score > best_score:
                best_score = score
                best_name  = fname

        if best_name and best_score >= 1:
            result[ref] = {"name": best_name, "bytes": files_data[best_name],
                           "matched": True}
        else:
            # Couldn't match — fall back to first file but mark as unmatched
            n = next(iter(files_data))
            result[ref] = {"name": n, "bytes": files_data[n], "matched": False}

    return result


def _extract_qn_ans_pairs(text: str) -> list[tuple]:
    """Extract every (qn, answer_letter) pair from a chunk of MS text.

    Handles all common IB mark-scheme formats:
        '1.  D'    '1   D'    'Q1  D'    '1: D'
        '1)  D'    '1 D'      '1.D'      '  1.D'

    Returns pairs in reading order (left-to-right, top-to-bottom).
    Does NOT match digits inside option text (e.g. 'D 1.5' won't be parsed
    as Q1=… because the digit must come before the letter).
    """
    pattern = re.compile(
        r"(?:^|[\s\|])"            # start of line or after whitespace/pipe
        r"(?:Q\s*)?"               # optional 'Q' prefix
        r"(\d{1,2})"               # question number
        r"[\.\):]?"                # optional . ) :
        r"\s*"                     # zero or more spaces (allow '1.D')
        r"([ABCD])"                # answer letter
        r"(?=$|[\s\.\,\;\|])"      # boundary after letter
    )
    pairs = []
    for m in pattern.finditer(text):
        qn = int(m.group(1))
        if 1 <= qn <= 99:
            pairs.append((qn, m.group(2)))
    return pairs


def extract_answers_pymupdf_per_section(ms_bytes: bytes) -> list[dict]:
    """Extract answers from MS PDF, grouped by section.

    A 'section' = a contiguous range of MS pages that hold answers for ONE
    paper. We detect section boundaries by looking for the answer-grid restart
    (i.e. Q1 appearing again, OR any Q# appearing twice).

    Returns:
        [{"start_page": int_1based, "end_page": int,
          "answers": {qn_str: 'A'|'B'|'C'|'D'}}, …]
    """
    sections: list[dict] = []
    current: dict = {"start_page": 1, "answers": {}, "max_qn": 0}

    try:
        doc = fitz.open(stream=ms_bytes, filetype="pdf")
    except Exception:
        return []

    for page_idx, page in enumerate(doc):
        # Try the page text as-is first (PyMuPDF preserves visual order
        # reasonably well for grid layouts).
        text = page.get_text()
        page_pairs = _extract_qn_ans_pairs(text)

        # Skip pages that look like prose (mostly text, few qn matches).
        # Real answer-grid pages have many short lines with N. X patterns.
        if len(page_pairs) < 3:
            continue

        for qn, ans in page_pairs:
            qn_str = str(qn)
            # New paper boundary detection:
            #  - the same qn appears again, OR
            #  - qn resets to 1 after we already saw qn >= 5
            is_boundary = (
                qn_str in current["answers"]
                or (qn == 1 and current["max_qn"] >= 5)
            )
            if is_boundary and current["answers"]:
                sections.append({
                    "start_page": current["start_page"],
                    "end_page":   page_idx + 1,
                    "answers":    current["answers"],
                })
                current = {"start_page": page_idx + 1, "answers": {}, "max_qn": 0}

            current["answers"][qn_str] = ans
            if qn > current["max_qn"]:
                current["max_qn"] = qn

    if current["answers"]:
        sections.append({
            "start_page": current["start_page"],
            "end_page":   len(doc),
            "answers":    current["answers"],
        })

    doc.close()
    return sections


def extract_answers_pymupdf(ms_bytes: bytes) -> dict:
    """Flat extraction — returns the union of all answers in all sections,
    using first-occurrence resolution for duplicates. Kept for backward
    compatibility with single-paper callers.
    """
    sections = extract_answers_pymupdf_per_section(ms_bytes)
    answers: dict = {}
    for sec in sections:
        for q, a in sec["answers"].items():
            answers.setdefault(q, a)
    return answers


# ───────────────────────────────────────────────────────────────────────────────
#  STRUCTURED MARK SCHEME (Math / Bio with worked solutions)
# ───────────────────────────────────────────────────────────────────────────────
def _detect_paper_code(text: str) -> str | None:
    """Pull a normalised paper code from a chunk of PDF text.

    IB papers carry one of two code formats:
      - "2221 – 7209"  →  numeric (year+month + paper#-tz)
      - "M21/5/MATHX/HP1/ENG/TZ1/XX"  →  alphanumeric
    We normalise both to a common form like "M21/HP1/TZ1" so a QP and MS
    for the same paper match even when they carry different code styles.
    Returns the normalised code or None.
    """
    if not text:
        return None

    # IB alphanumeric: "M21/5/MATHX/HP1/ENG/TZ1"
    m = re.search(r'([MN])(\d{2})/\d/MATH\w+/([HS])P(\d)/\w+/(TZ\d|TZ0)', text)
    if m:
        session, year, level, paper, tz = m.groups()
        return f"{session}{year}/{level}P{paper}/{tz}"

    # Numeric: "2221-7209".
    # IB convention (for Math Analysis & Approaches and AI):
    #   first 4 digits = YYSS where YY = year, SS = session marker
    #     - 22 → May 2021 paper-set (the year is YY+1 — see note)
    #     Actually the rule is: "2221" means 2021 May (subject-paper-year),
    #     where "22" is the math-subject code and "21" the year.
    # The last 4 digits are paper-level-TZ:
    #   71XX = HL Paper 1, 72XX = SL/AI Paper 1, etc.
    # Since the mapping isn't trivially recoverable, we keep both code
    # styles and rely on the user uploading a matching pair (so QP code
    # set ⊆ MS code set for at least some sessions, or vice versa).
    m = re.search(r'\b(\d{4})\s*[\-–]\s*(\d{4})\b', text)
    if m:
        return f"{m.group(1)}-{m.group(2)}"

    return None


def _build_page_to_paper_map(pdf_bytes: bytes) -> dict:
    """For every page in the PDF, identify its paper code. Returns
    {page_idx: paper_code_string}. Pages without a detectable code keep
    the code from the previous page (so cover/instruction pages of a
    paper still get tagged)."""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception:
        return {}
    page_to_code = {}
    last_code = None
    for pi in range(len(doc)):
        text = doc[pi].get_text()
        code = _detect_paper_code(text)
        if code:
            last_code = code
        if last_code:
            page_to_code[pi] = last_code
    doc.close()
    return page_to_code


# Continuation text patterns — skip in MS crops
MS_CONT_PAT = re.compile(r'question\s*\d*\s*continued\.?|this question continues', re.I)

def find_ms_question_locations(ms_bytes: bytes) -> list[dict]:
    """For structured MS (e.g. Math): find every question's bounding box
    across all pages of the MS PDF.

    Math MS layout differs from QP — markers can be:
      "1.", "2.", "3." (start of line, in body or table)
      "1)", "2)" sometimes
      "1." in big bold at left margin
    A question's content extends down to the next question marker.

    Returns a list of section dicts, one per detected MS paper:
        [{"start_page": int, "end_page": int,
          "questions": {qn: {"page_idx": int, "top_y": float,
                             "bottom_y": float, "end_page_idx": int,
                             "end_y": float}}}, …]

    `end_page_idx` and `end_y` together mark where the question stops
    (it may span across pages).
    """
    try:
        doc = fitz.open(stream=ms_bytes, filetype="pdf")
    except Exception:
        return []

    # Step 1: scan every page for "N." question markers.
    # Math MS markers in IB format are typically "N." on a line by themselves
    # at the LEFT margin (often x0 < 50). The number is followed by either:
    #   - end of line (just "N.")
    #   - or a space then question content (rare in MS — the (a)/(b) parts
    #     live on the next line)
    # We REJECT lines like "1.9 and..." or "1.5 20" because the digit after
    # the dot means it's a decimal, not a marker.
    raw_markers = []   # (qn, page_idx, top_y, page_height, marker_right_x)
    # Markers can be:
    #   "1."  (most common)
    #   "1. (a)"  (when question content starts on same line)
    #   "1"   alone in bold (no dot — some Math MS use this style)
    #   "Q1." / "Q1. (a)" — IB May 2025 TZ3 format (Q-prefix)
    pat_strict  = re.compile(r"^(\d{1,2})\.\s*$")
    pat_qa      = re.compile(r"^(\d{1,2})\.\s+\(?\w")
    pat_nodot   = re.compile(r"^(\d{1,2})\s*$")     # bare "1", "2", "10"
    pat_qprefix = re.compile(r"^Q(\d{1,2})\.\s*(?:\(?\w|$)")  # "Q1.", "Q1. (a)"

    for pi in range(len(doc)):
        page = doc[pi]
        ph = page.rect.height
        td = page.get_text("dict")
        for block in td.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue
                full_line = "".join(s.get("text", "") for s in spans).strip()
                bb = line["bbox"]
                if bb[0] > 80:   # MS question markers sit at the far left
                    continue

                qn = None
                m = pat_strict.match(full_line) or pat_qa.match(full_line)
                if m:
                    qn = int(m.group(1))
                else:
                    # Q-prefix format: "Q1.", "Q2. EITHER", "Q10."
                    mq = pat_qprefix.match(full_line)
                    if mq:
                        qn = int(mq.group(1))
                    else:
                        # Try bare "1", "2", … but ONLY if it's bold AND sits at
                        # the very left margin (x0 < 50pt). Instructions pages
                        # also have bold "1", "2" but indented further right.
                        m2 = pat_nodot.match(full_line)
                        if m2 and bb[0] < 50:
                            font = spans[0].get("font", "")
                            size = spans[0].get("size", 0)
                            # Bold AND reasonable body-text size (10-13pt)
                            if ("Bold" in font or "bold" in font) and 9 <= size <= 14:
                                qn = int(m2.group(1))

                if qn is None:
                    continue
                if qn < 1 or qn > 30:
                    continue
                # Reject decimal numbers ("1.5", "1.9")
                if re.match(rf"^{qn}\.\d", full_line):
                    continue
                # Marker right edge for cropping
                marker_right_x = spans[0]["bbox"][2] if spans else (bb[0] + 20)
                raw_markers.append((qn, pi, bb[1], ph, marker_right_x))

    doc.close()

    if not raw_markers:
        return []

    # Step 2: split into sections by Q# restart pattern
    sections: list[dict] = []
    current: dict = {"questions": {}, "_order": [],
                     "start_page": raw_markers[0][1] + 1, "max_qn": 0}

    for qn, pi, top_y, ph, mrx in raw_markers:
        # New section if qn repeats OR resets to 1 after Q ≥ 5
        if (qn in current["questions"] or
                (qn == 1 and current["max_qn"] >= 5)):
            if current["questions"]:
                sections.append(current)
                current = {"questions": {}, "_order": [],
                           "start_page": pi + 1, "max_qn": 0}

        current["questions"][qn] = {
            "page_idx":       pi,
            "top_y":          top_y,
            "bottom_y":       ph * 0.92,   # tentative — refined below
            "end_page_idx":   pi,
            "end_y":          ph * 0.92,
            "page_height":    ph,
            "marker_right_x": mrx,
        }
        current["_order"].append((qn, pi, top_y))
        if qn > current["max_qn"]:
            current["max_qn"] = qn

    if current["questions"]:
        sections.append(current)

    # Step 3: for each section, compute each question's end (where next
    # question begins, possibly on a later page)
    for sec in sections:
        order = sec["_order"]
        for idx, (qn, pi, top_y) in enumerate(order):
            if idx + 1 < len(order):
                next_qn, next_pi, next_top_y = order[idx + 1]
                if next_pi == pi:
                    # Next question on same page → end_y is just above it
                    sec["questions"][qn]["end_page_idx"] = pi
                    sec["questions"][qn]["end_y"] = next_top_y - 4
                else:
                    # Next question on later page → this question ends
                    # at bottom of its starting page (and may span pages)
                    sec["questions"][qn]["end_page_idx"] = next_pi
                    sec["questions"][qn]["end_y"] = next_top_y - 4
            # else: last question in section → keeps tentative bottom

        # Set section's end_page
        last_qn, last_pi, _ = order[-1]
        sec["end_page"] = sec["questions"][last_qn]["end_page_idx"] + 1
        # Remove internal _order key
        sec.pop("_order", None)

    # Filter out tiny "sections" (< 5 questions). These are typically
    # table-of-contents / index pages, not real worked-solution pages.
    sections = [s for s in sections if len(s["questions"]) >= 5]

    # Also filter out sections that sit inside instructions/cover pages.
    # IB MS PDFs have ~5 pages of "Instructions to Examiners", "Abbreviations",
    # "Marks awarded for...", etc., which contain numbered lists that look
    # like question markers. We detect this by checking the page text.
    instr_pat = re.compile(
        r"(instructions to examiners|abbreviations|marks? awarded for|"
        r"using the markscheme|method of marking|implied marks|"
        r"misread|brackets in working)",
        re.IGNORECASE
    )
    try:
        doc2 = fitz.open(stream=ms_bytes, filetype="pdf")
        filtered = []
        for s in sections:
            start = s["start_page"] - 1
            # Look at the section's first 1-2 pages
            instr_score = 0
            for pi in range(start, min(start + 2, len(doc2))):
                txt = doc2[pi].get_text()
                if instr_pat.search(txt):
                    instr_score += 1
            if instr_score == 0:
                filtered.append(s)
        doc2.close()
        sections = filtered
    except Exception:
        pass

    # Tag each section with the paper code on its first page (used to match
    # the right MS section for each Excel paper segment).
    page_codes = _build_page_to_paper_map(ms_bytes)
    for s in sections:
        first_page_idx = s["start_page"] - 1
        s["paper_code"] = page_codes.get(first_page_idx)

    return sections


def _topic_keywords(text: str) -> set:
    """Extract distinctive content words from text, ignoring boilerplate.
    Used to compare QP question topic vs MS answer topic — if the keyword
    overlap is too small, the MS is for a different question."""
    if not text:
        return set()
    # Normalise: lowercase, strip math markers, keep letters/digits
    t = text.lower()
    # Remove IB boilerplate that appears in every question
    boilerplate = re.compile(
        r"(maximum\s+mark|marks?\s*\]|"
        r"answers?\s+must\s+be|working\s+and/?or|"
        r"international\s+baccalaureate|"
        r"©\s*\d{4}|turn\s+over|"
        r"award\s+(a\d|m\d|r\d)|note\s*:|"
        r"\b(m\d|a\d|r\d|ft|ag)\b|"
        r"calculator|gdc|gradient|substitute|"
        r"\bsolve\b|\bfind\b|\bcalculate\b|\bdetermine\b|\bgiven\b|\bshow\b|"
        r"\bvalue\b|\bvalues\b|\bequation\b|\bexpression\b|\banswer\b|"
        r"\bcorrect\b|\bworking\b|\busing\b|\bfollowing\b|"
        r"\bthe\b|\bof\b|\ba\b|\bto\b|\bin\b|\bis\b|\bfor\b|\bbe\b|"
        r"\bby\b|\bwith\b|\bare\b|\band\b|\bor\b|\bif\b|\bnot\b|\bcan\b|"
        r"\bone\b|\btwo\b|\bthree\b|\bfour\b|\bfive\b|\bsix\b|\bseven\b|"
        r"\beight\b|\bnine\b|\bten\b|"
        r"\bmark\b|\bmarks\b|\bpart\b|\bparts\b|\bpoint\b|\bpoints\b)",
        re.IGNORECASE
    )
    t = boilerplate.sub(" ", t)
    # Extract alphabetic tokens of length ≥ 4 (skip short noise)
    tokens = re.findall(r"\b[a-z][a-z]{3,}\b", t)
    return set(tokens)


def _topics_match(qp_text: str, ms_text: str, threshold: float = 0.10) -> bool:
    """Return True if QP question and MS answer share enough topic-keywords
    to plausibly be about the same question. Used as a safety check before
    embedding an MS image, to catch wrong-paper pairings."""
    qp_kw = _topic_keywords(qp_text)
    ms_kw = _topic_keywords(ms_text)
    if not qp_kw or not ms_kw:
        return True   # can't tell — give benefit of doubt
    overlap = qp_kw & ms_kw
    # Match ratio over the smaller set
    smaller = min(len(qp_kw), len(ms_kw))
    if smaller == 0:
        return True
    ratio = len(overlap) / smaller
    return ratio >= threshold


def crop_ms_question_image(ms_bytes: bytes, q_info: dict,
                           dpi: int = 150) -> bytes | None:
    """Crop the worked solution for one MS question.

    The question may span multiple pages. We render each spanned page,
    stack them vertically into one tall image, and return PNG bytes.

    Crop rules (per page slice):
      - Top:      a few px above the question marker
      - Bottom:   just above the next question / footer
      - Left:     trims the page's left margin (~5% in)
      - Right:    trims the page's right margin (~5% in)
    Headers (top 5%) and footers (below 92%) are always excluded.
    """
    if not q_info:
        return None

    start_pi = q_info["page_idx"]
    end_pi   = q_info.get("end_page_idx", start_pi)
    start_y  = q_info["top_y"]
    end_y    = q_info["end_y"]
    mrx      = q_info.get("marker_right_x", 0)   # right edge of "N." marker

    try:
        doc = fitz.open(stream=ms_bytes, filetype="pdf")
    except Exception:
        return None

    scale = dpi / 72
    pieces = []

    # Footer detectors. Strong patterns (always footer): IB exam codes,
    # copyright. Weak pattern (bare numbers) is only treated as footer
    # when the line is in the very bottom strip (y > 0.93 × ph), so that
    # equation subscripts like "u₁ = -6, d = 2" aren't mistaken for page
    # numbers when they appear mid-page.
    ms_strong_footer = re.compile(
        r"(M\d{2}/\d/|N\d{2}/\d/|"
        r"international\s+baccalaureate|"
        r"©\s*\d{4})",
        re.IGNORECASE
    )
    ms_bare_num = re.compile(r"^\s*\d{1,3}\s*$")

    def is_ms_footer(text: str, y0: float, ph_: float) -> bool:
        if ms_strong_footer.search(text):
            return True
        if ms_bare_num.match(text) and y0 > ph_ * 0.93:
            return True
        return False

    def find_ms_footer_top(page) -> float:
        """Return the y0 of the first footer line in lower half of page."""
        ph_ = page.rect.height
        td = page.get_text("dict")
        first_y = ph_ * 0.95
        for block in td.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue
                text = "".join(s.get("text", "") for s in spans).strip()
                if not text:
                    continue
                bb = line["bbox"]
                if bb[1] < ph_ * 0.85:
                    continue
                if is_ms_footer(text, bb[1], ph_):
                    if bb[1] < first_y:
                        first_y = bb[1]
        return first_y

    def find_header_bottom(page) -> float:
        """Return the y1 of the page-header block (page number / exam code
        at the top of the page). Also skips 'Question N continued' lines."""
        ph_ = page.rect.height
        td = page.get_text("dict")
        header_y1 = 0
        header_pat = re.compile(
            r"(^\s*[–\-]\s*\d{1,3}\s*[–\-]\s*$|"   # "– 9 –"
            r"M\d{2}/\d/|N\d{2}/\d/)",
            re.IGNORECASE
        )
        for block in td.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue
                text = "".join(s.get("text", "") for s in spans).strip()
                if not text:
                    continue
                bb = line["bbox"]
                if bb[1] > ph_ * 0.20:
                    continue
                if header_pat.search(text) or MS_CONT_PAT.search(text):
                    if bb[3] > header_y1:
                        header_y1 = bb[3]
        return header_y1

    def find_last_content_y(page, top_y: float, hard_bottom: float) -> float:
        """Return the y1 of the LAST real content line between top_y and
        hard_bottom on this page. Skips continuation headers."""
        ph_ = page.rect.height
        td = page.get_text("dict")
        last_y1 = top_y
        for block in td.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue
                text = "".join(s.get("text", "") for s in spans).strip()
                if not text:
                    continue
                bb = line["bbox"]
                if bb[1] < top_y - 2:
                    continue
                if bb[1] >= hard_bottom - 2:
                    continue
                if is_ms_footer(text, bb[1], ph_):
                    continue
                if MS_CONT_PAT.search(text):
                    continue
                if bb[3] > last_y1:
                    last_y1 = bb[3]
        return last_y1

    for pi in range(start_pi, end_pi + 1):
        if pi >= len(doc):
            break
        page = doc[pi]
        ph = page.rect.height
        pw = page.rect.width
        mat = fitz.Matrix(scale, scale)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        full = Image.open(io.BytesIO(pix.tobytes("png")))
        iw, ih = full.size

        # Detect this page's footer top (page number / exam code)
        footer_top = find_ms_footer_top(page)
        page_bottom_y = max(ph * 0.5, footer_top - 6)

        # Horizontal trim
        right_px = iw - int(pw * 0.06 * scale)

        if pi == start_pi and mrx:
            left_px = max(0, int((mrx + 6) * scale))
        else:
            left_px = int(pw * 0.04 * scale)

        # Detect this page's header bottom (page number + exam code at top)
        header_bottom = find_header_bottom(page)
        page_top_y = max(ph * 0.06, header_bottom + 6) if header_bottom else ph * 0.06

        # Determine vertical bounds per page, tight to actual content
        if pi == start_pi and pi == end_pi:
            # Single-page question: known end_y from next-marker detection
            top_y_pi = start_y
            bot_y_pi = min(end_y, page_bottom_y)
        elif pi == start_pi:
            # First page of multi-page: trim to the last real content line on
            # this page (no big empty area before page break)
            top_y_pi = start_y
            bot_y_pi = find_last_content_y(page, start_y, page_bottom_y)
            bot_y_pi = min(bot_y_pi + 6, page_bottom_y)
        elif pi == end_pi:
            # Last page: start below the page header
            top_y_pi = page_top_y
            actual_end = find_last_content_y(page, top_y_pi, min(end_y, page_bottom_y))
            bot_y_pi = min(actual_end + 6, end_y, page_bottom_y)
        else:
            # Middle page: from below header to last real content
            top_y_pi = page_top_y
            bot_y_pi = find_last_content_y(page, top_y_pi, page_bottom_y)
            bot_y_pi = min(bot_y_pi + 6, page_bottom_y)

        top_px = max(0,  int(top_y_pi * scale) - 6)
        bot_px = min(ih, int(bot_y_pi * scale) + 4)

        if bot_px <= top_px or right_px <= left_px:
            continue
        pieces.append(full.crop((left_px, top_px, right_px, bot_px)))

    doc.close()

    if not pieces:
        return None

    # ── Split long answers into page-sized chunks ─────────────────────────────
    # Each page in the Word document is ~A4 height at 150dpi ≈ 1754px.
    # Content area (minus margins) ≈ 1400px. We allow ~1300px per chunk so
    # the image is readable without shrinking to illegibility.
    # If the total height fits in one chunk, we return one image; otherwise
    # we split into multiple PNG chunks that the Word generator inserts
    # sequentially (Part 1, Part 2, …).
    MAX_CHUNK_H = 1300   # px at 150dpi — safe readable height per Word page

    # First stack so we can measure total height
    total_w = max(p.size[0] for p in pieces)
    total_h = sum(p.size[1] for p in pieces)

    if total_h <= MAX_CHUNK_H:
        # Short enough — return single image as before (bytes)
        stacked = Image.new("RGB", (total_w, total_h), "white")
        y_off = 0
        for p in pieces:
            stacked.paste(p, (0, y_off))
            y_off += p.size[1]
        buf = io.BytesIO()
        stacked.save(buf, format="PNG", optimize=True)
        return buf.getvalue()

    # Tall answer — build full stacked image then slice into chunks
    stacked = Image.new("RGB", (total_w, total_h), "white")
    y_off = 0
    for p in pieces:
        stacked.paste(p, (0, y_off))
        y_off += p.size[1]

    chunks = []
    y_cursor = 0
    while y_cursor < total_h:
        chunk = stacked.crop((0, y_cursor, total_w, min(y_cursor + MAX_CHUNK_H, total_h)))
        buf = io.BytesIO()
        chunk.save(buf, format="PNG", optimize=True)
        chunks.append(buf.getvalue())
        y_cursor += MAX_CHUNK_H

    return chunks   # list of bytes — caller must handle multi-part insertion


def extract_answers(client: anthropic.Anthropic, ms_b64: str,
                    q_nums: list[int],
                    ms_bytes: bytes = None) -> dict:
    """Primary: deterministic PyMuPDF extraction (regex on text grid).
       Fallback: Claude — only for question numbers that PyMuPDF didn't find.
       The fallback is rare and inexpensive.
    """
    answers: dict = {}

    # 1) PyMuPDF deterministic pass
    if ms_bytes is not None:
        answers = extract_answers_pymupdf(ms_bytes)

    # 2) Claude fallback for whatever's still missing
    missing = [n for n in q_nums if str(n) not in answers]
    if missing and client is not None:
        nums = ", ".join(str(n) for n in missing)
        try:
            response = client.messages.create(
                model="claude-opus-4-5",
                max_tokens=1500,
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "document",
                         "source": {"type": "base64",
                                    "media_type": "application/pdf",
                                    "data": ms_b64}},
                        {"type": "text", "text": (
                            f"Extract the correct answers for question numbers "
                            f"{nums} from this IB mark scheme.\n\n"
                            "The mark scheme is a grid of letters A–D.\n\n"
                            "Return ONLY this JSON object — no markdown, no "
                            "explanation:\n"
                            '{"1":"D","2":"A",...}\n\n'
                            'Use "NOT_FOUND" for any question that is genuinely '
                            "absent."
                        )},
                    ],
                }]
            )
            raw = "".join(b.text for b in response.content if hasattr(b, "text"))
            data = safe_json(raw)
            if isinstance(data, dict):
                for k, v in data.items():
                    if k not in answers and v in ("A", "B", "C", "D"):
                        answers[k] = v
        except Exception:
            pass   # silent fallback — UI shows missing answers

    return answers


# ───────────────────────────────────────────────────────────────────────────────
#  LaTeX → Unicode (for inline equations in text-only questions)
# ───────────────────────────────────────────────────────────────────────────────
_SUPS = {"0":"⁰","1":"¹","2":"²","3":"³","4":"⁴","5":"⁵","6":"⁶","7":"⁷",
         "8":"⁸","9":"⁹","+":"⁺","-":"⁻","n":"ⁿ","m":"ᵐ","x":"ˣ","a":"ᵃ","b":"ᵇ"}
_SUBS = {"0":"₀","1":"₁","2":"₂","3":"₃","4":"₄","5":"₅","6":"₆","7":"₇",
         "8":"₈","9":"₉","+":"₊","-":"₋","n":"ₙ","e":"ₑ","r":"ᵣ","x":"ₓ","a":"ₐ"}

def _to_sup(s): return "".join(_SUPS.get(c, c) for c in s)
def _to_sub(s): return "".join(_SUBS.get(c, c) for c in s)

def _math(m: str) -> str:
    t = re.sub(r"\\frac\{([^{}]+)\}\{([^{}]+)\}", r"(\1)/(\2)", m)
    for k, v in {
        r"\times":"×", r"\rightarrow":"→", r"\rightleftharpoons":"⇌",
        r"\to":"→", r"\Delta":"Δ", r"\delta":"δ", r"\ominus":"⊖",
        r"\alpha":"α", r"\beta":"β", r"\gamma":"γ", r"\lambda":"λ",
        r"\mu":"μ", r"\pi":"π", r"\sigma":"σ", r"\cdot":"·", r"\pm":"±",
        r"\geq":"≥", r"\leq":"≤", r"\neq":"≠", r"\approx":"≈",
        r"\circ":"°",
    }.items():
        t = t.replace(k, v)
    t = re.sub(r"\^\{([^{}]{1,12})\}", lambda m: _to_sup(m.group(1)), t)
    t = re.sub(r"\^([a-zA-Z0-9+\-])",  lambda m: _to_sup(m.group(1)), t)
    t = re.sub(r"_\{([^{}]{1,12})\}",  lambda m: _to_sub(m.group(1)), t)
    t = re.sub(r"_([a-zA-Z0-9])",      lambda m: _to_sub(m.group(1)), t)
    t = re.sub(r"\\(?:text|mathrm|mbox)\{([^{}]+)\}", r"\1", t)
    return re.sub(r"[{}]", "", t).strip()

def latex_to_text(text: str) -> str:
    if not text:
        return ""
    t = str(text)
    t = re.sub(r"\$\$([^$]+)\$\$", lambda m: _math(m.group(1)), t)
    t = re.sub(r"\$([^$\n]+)\$",   lambda m: _math(m.group(1)), t)
    t = re.sub(r"\\\[([^\]]+)\\\]", lambda m: _math(m.group(1)), t)
    t = re.sub(r"\*\*([^*]+)\*\*", r"\1", t)
    t = re.sub(r"\*([^*]+)\*",     r"\1", t)
    return t


# ═══════════════════════════════════════════════════════════════════════════════
#  WORD DOCUMENT BUILDER
# ═══════════════════════════════════════════════════════════════════════════════
DIFF_ORDER = ["Easy", "Medium", "Hard"]

# A4 page with 2 cm margins → content width ≈ 17 cm
CONTENT_WIDTH_CM = 17.0


def _set_cell_borders(cell, top="single", left="single", right="single",
                      bottom="single", color="999999",
                      bot_color=None, top_color=None,
                      sz_top=8, sz_bot=8):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side, style, sz, c in [
        ("top",    top,    sz_top, top_color or color),
        ("left",   left,   8,      color),
        ("bottom", bottom, sz_bot, bot_color or color),
        ("right",  right,  8,      color),
    ]:
        b = OxmlElement(f"w:{side}")
        b.set(docx_qn("w:val"),   style)
        b.set(docx_qn("w:sz"),    str(sz))
        b.set(docx_qn("w:color"), c)
        tcBorders.append(b)
    tcPr.append(tcBorders)


def _run(para, text, *, bold=False, italic=False, color=None, size_pt=11):
    r = para.add_run(str(text))
    r.bold        = bold
    r.italic      = italic
    r.font.size   = Pt(size_pt)
    r.font.name   = "Arial"
    if color:
        r.font.color.rgb = RGBColor(*bytes.fromhex(color))
    return r


def _insert_pBdr(p, pBdr):
    """OOXML schema requires pBdr to come BEFORE spacing/ind/jc/rPr in pPr."""
    pPr = p._p.get_or_add_pPr()
    insert_idx = len(pPr)
    for i, child in enumerate(pPr):
        tag = child.tag.split('}')[-1]
        if tag in ('spacing', 'ind', 'jc', 'contextualSpacing',
                   'mirrorIndents', 'rPr', 'sectPr'):
            insert_idx = i
            break
    pPr.insert(insert_idx, pBdr)


def _hr(doc, color="CCCCCC"):
    p    = doc.add_paragraph()
    pBdr = OxmlElement("w:pBdr")
    bot  = OxmlElement("w:bottom")
    bot.set(docx_qn("w:val"),   "single")
    bot.set(docx_qn("w:sz"),    "6")
    bot.set(docx_qn("w:color"), color)
    pBdr.append(bot)
    _insert_pBdr(p, pBdr)
    return p


def _add_solution_box(doc, n_lines: int = 4, row_height_twips: int = 560):
    """Table with bottom-border writing lines.
       row_height_twips: 560 = 28pt (default), 510 = ~25pt (compact)
    """
    n_lines = max(4, min(n_lines, 10))
    table = doc.add_table(rows=n_lines, cols=1)
    table.autofit = False
    for cell in table.columns[0].cells:
        cell.width = Cm(CONTENT_WIDTH_CM)

    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        trH  = OxmlElement("w:trHeight")
        trH.set(docx_qn("w:val"), str(row_height_twips))
        trPr.append(trH)

        cell = row.cells[0]
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = OxmlElement("w:tcBorders")
        for side, val, sz, color in [
            ("top",    "nil",    "0", "auto"),
            ("left",   "nil",    "0", "auto"),
            ("right",  "nil",    "0", "auto"),
            ("bottom", "single", "6", "BFBFBF"),
        ]:
            b = OxmlElement(f"w:{side}")
            b.set(docx_qn("w:val"), val)
            if val != "nil":
                b.set(docx_qn("w:sz"),    sz)
                b.set(docx_qn("w:color"), color)
            tcBorders.append(b)
        tcPr.append(tcBorders)

        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)


def build_word_document(
    questions:  list[dict],
    qp_bytes:   bytes | None = None,
    locations:  dict | None = None,
    progress_cb=None,
) -> bytes:
    """Build the Word worksheet.
       - Visual questions → cropped image only (no text duplication)
       - Text-only questions → text + options
       The optional `qp_bytes` / `locations` are LEGACY single-paper fallback.
       For multi-paper worksheets, each question carries its own
       `_qp_bytes` / `_qp_locs` keys.
    """
    doc = Document()

    # Page setup: A4, 2 cm margins
    for section in doc.sections:
        section.top_margin    = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin   = Cm(2.0)
        section.right_margin  = Cm(2.0)

    # ── Group by Topic → Difficulty ───────────────────────────────────────────
    # First, normalize chapter names so visual differences don't fragment groups:
    #   "Chapter 1 : Number & Algebra"  →  "1. Number & Algebra"
    #   "Chapter 1: Number & Algebra"   →  "1. Number & Algebra"
    #   " 1.  Number & Algebra "        →  "1. Number & Algebra"
    def _normalize_topic(t: str) -> str:
        if not t or str(t).strip().lower() in ("error", "none", "nan", ""):
            return "Unclassified"
        s = str(t).strip()
        # Strip leading # markers
        s = re.sub(r'^#+\s*', '', s).strip()
        # "Chapter N : Foo" or "Chapter N: Foo" → "N. Foo"
        m = re.match(r"^\s*chapter\s+(\d+)\s*[:\.\-–]?\s*(.+?)\s*$",
                     s, re.IGNORECASE)
        if m:
            return f"{m.group(1)}. {m.group(2).strip()}"
        # "N : Foo" / "N: Foo" / "N. Foo" → "N. Foo"
        m = re.match(r"^\s*(\d+)\s*[\.:\-–]\s*(.+?)\s*$", s)
        if m:
            return f"{m.group(1)}. {m.group(2).strip()}"
        # Collapse multiple spaces
        return re.sub(r"\s+", " ", s)

    # Apply normalization in-place so later rendering uses the canonical name
    for q in questions:
        q["topic"] = _normalize_topic(q.get("topic"))

    grouped: dict[str, dict[str, list]] = {}
    for q in questions:
        t = q["topic"]
        d = q.get("difficulty") or "Unspecified"
        grouped.setdefault(t, {}).setdefault(d, []).append(q)

    # Sort chapters by their leading number (1., 2., …, 10.). Topics without
    # a leading number (e.g. "Unclassified") go last in alphabetical order.
    def _chapter_sort_key(topic: str):
        m = re.match(r"^\s*(\d+)\s*[\.\-:]", str(topic))
        if m:
            return (0, int(m.group(1)), str(topic))
        return (1, 0, str(topic).lower())

    sorted_topics = sorted(grouped.keys(), key=_chapter_sort_key)

    visual_qs   = [q for q in questions if q.get("needs_image")]
    total_imgs  = len(visual_qs)
    img_done    = 0

    is_first_chapter = True
    for topic in sorted_topics:
        diffs = grouped[topic]
        # ── Chapter heading: bold 20pt black, plain (no border, no color) ──
        h = doc.add_paragraph()
        # Page break before each new chapter (except the very first one)
        if not is_first_chapter:
            h.paragraph_format.page_break_before = True
        h.paragraph_format.space_before = Pt(0)
        h.paragraph_format.space_after  = Pt(8)
        _run(h, topic, bold=True, size_pt=20)

        sorted_diffs = sorted(
            diffs.keys(),
            key=lambda d: DIFF_ORDER.index(d) if d in DIFF_ORDER else 99
        )

        # Per-chapter display counter — restarts at 1 for every new chapter.
        # The original q["qn"] from Excel is still used internally for
        # cropping the QP image and matching the answer.
        display_counter = 0

        is_first_diff_in_chapter = True
        for diff in sorted_diffs:
            # Difficulty heading: bold + italic, 13pt
            dp = doc.add_paragraph()
            # Page break before a new difficulty within same chapter
            # (skip if it's the first difficulty — chapter heading already
            # triggered the page break)
            if not is_first_diff_in_chapter:
                dp.paragraph_format.page_break_before = True
            dp.paragraph_format.space_before = Pt(8)
            dp.paragraph_format.space_after  = Pt(6)
            _run(dp, f"— {diff} questions —",
                 bold=True, italic=True, size_pt=13)

            is_first_q_in_diff = True
            for q in diffs[diff]:
                q_num   = q["qn"]                     # internal — for QP/MS lookup
                display_counter += 1
                display_num = display_counter         # shown to the student
                found   = q.get("found", True)
                vis     = q.get("needs_image", False) and found
                answer  = q.get("answer", "Answer not found in uploaded Mark Scheme")
                ans_ok  = q.get("answerFound", False)
                text    = latex_to_text(q.get("text", ""))
                opts    = {L: latex_to_text(q.get(L, "")) for L in "ABCD"}
                topic_  = q.get("topic", "Unclassified")
                diff_   = q.get("difficulty", "")
                marks   = q.get("marks", 1)
                ref     = q.get("ref", "")
                quote   = q.get("quote", "")

                # ── Question header: bold 13pt ─────────────────────────────────
                qh = doc.add_paragraph()
                # Every question starts on a new page, except the very first
                # question in its difficulty group (the difficulty heading
                # paragraph already triggered the page break above).
                if not is_first_q_in_diff:
                    qh.paragraph_format.page_break_before = True
                qh.paragraph_format.space_before = Pt(0)
                qh.paragraph_format.space_after  = Pt(2)
                _run(qh, f"Question: {display_num}", bold=True, size_pt=13)

                # ── Meta line ─────────────────────────────────────────────────
                # **Level of question**: Easy  |  **Number of Marks: **1
                # CHANGE 8: Reference is NEVER shown in the Word output.
                # It is used only internally for QP/MS matching.
                mp = doc.add_paragraph()
                mp.paragraph_format.space_before = Pt(0)
                mp.paragraph_format.space_after  = Pt(2)
                _run(mp, "Level of question",  bold=True, size_pt=11)
                _run(mp, f": {diff_}  |  ",                size_pt=11)
                _run(mp, "Number of Marks: ",  bold=True, size_pt=11)
                _run(mp, f"{marks}",                       size_pt=11)

                # ── Chapter line ──────────────────────────────────────────────
                # **Chapter** :  [name]
                cp = doc.add_paragraph()
                cp.paragraph_format.space_before = Pt(0)
                cp.paragraph_format.space_after  = Pt(2)
                _run(cp, "Chapter", bold=True, size_pt=11)
                _run(cp, f" :  {topic_}",     size_pt=11)

                # ── Separator after meta info ─────────────────────────────────
                hr_after_meta = _hr(doc, color="BFBFBF")

                # ── Detect failed text extraction ──────────────────────────────
                def _looks_failed(stem: str, options: dict) -> bool:
                    """True if the extracted text is empty, placeholder, or
                    test data — i.e. anything that would look unprofessional
                    in the final Word file. When True, the build switches to
                    a cropped image from QP PDF instead.
                    """
                    s = (stem or "").strip()
                    if not s:
                        return True
                    sl = s.lower()
                    # Bracketed placeholders: [Q2 stem], [Q12], [stem]
                    if re.search(r"\[\s*q\s*\d+", sl):
                        return True
                    if re.search(r"\[\s*stem\s*\]", sl):
                        return True
                    # "Test stem for Q12", "test stem", "stem"
                    if re.search(r"\btest\s+stem\b", sl):
                        return True
                    if sl in ("stem", "[stem]", "n/a", "tbd",
                              "placeholder", "todo", "todo:"):
                        return True
                    # "Q12 stem", "Q5 stem"
                    if re.match(r"^q\s*\d+\s*stem\s*$", sl):
                        return True
                    # Generic placeholder words
                    if "placeholder" in sl:
                        return True

                    # Options validation
                    non_empty = [v for v in options.values() if v and v.strip()]
                    if not non_empty:
                        return True
                    # Single letters: A=A, B=B…
                    if all(v.strip().upper() in ("A", "B", "C", "D")
                           for v in non_empty):
                        return True
                    # "option A", "option B", "Option C"…
                    if all(re.match(r"^option\s*[a-d]$", v.strip(), re.I)
                           for v in non_empty):
                        return True
                    # Generic placeholder words in options
                    if any("placeholder" in v.lower() or
                           re.search(r"\[\s*[a-d]\s*\]", v.lower())
                           for v in non_empty):
                        return True
                    return False

                # If marked text-only but extraction is bad AND we can crop the
                # question from QP → switch to image fallback.
                if (found and not vis
                        and _looks_failed(text, opts)
                        and q_num in locations):
                    vis = True

                # ── Body: image OR text ────────────────────────────────────────
                # Structured mode always uses image crop — never plain text —
                # so questions with equations/diagrams are always clear.
                if is_structured and found and not vis:
                    vis = True   # force image for every found structured Q

                if not found:
                    np_ = doc.add_paragraph()
                    _run(np_, "Question not found on specified page - needs review",
                         italic=True, size_pt=11)

                elif vis:
                    if progress_cb:
                        progress_cb(img_done + 1, total_imgs or 1,
                                    f"Cropping Q{q_num} from QP…")
                    img_done += 1
                    # Use the per-row location (matched by Excel page number)
                    # so the right occurrence of Q# is cropped from the PDF.
                    q_loc = q.get("_loc")
                    img_bytes = crop_question_png(qp_bytes, q_loc) if q_loc else None
                    if img_bytes:
                        ip = doc.add_paragraph()
                        ip.paragraph_format.space_before = Pt(0)
                        ip.paragraph_format.space_after  = Pt(4)
                        ip.add_run().add_picture(io.BytesIO(img_bytes),
                                                 width=Cm(CONTENT_WIDTH_CM - 1.0))
                    elif text and not _looks_failed(text, opts):
                        # Crop failed but we have valid text — use it
                        for line in text.split("\n"):
                            if line.strip():
                                _run(doc.add_paragraph(), line, size_pt=11)
                        for L in "ABCD":
                            if opts[L]:
                                op = doc.add_paragraph()
                                op.paragraph_format.space_before = Pt(2)
                                op.paragraph_format.space_after  = Pt(2)
                                _run(op, f"{L}.  {opts[L]}", size_pt=11)
                    else:
                        # Both crop and text failed
                        np_ = doc.add_paragraph()
                        _run(np_,
                             f"Question Q{q_num} could not be extracted from QP — "
                             "please refer to the original Question Paper.",
                             italic=True, size_pt=11)

                else:
                    # Text-only question (extraction looks valid)
                    for line in text.split("\n"):
                        if line.strip():
                            tp = doc.add_paragraph()
                            tp.paragraph_format.space_before = Pt(0)
                            tp.paragraph_format.space_after  = Pt(2)
                            _run(tp, line, size_pt=11)
                    for L in "ABCD":
                        if opts[L]:
                            op = doc.add_paragraph()
                            op.paragraph_format.space_before = Pt(2)
                            op.paragraph_format.space_after  = Pt(2)
                            _run(op, f"{L}.  {opts[L]}", size_pt=11)

                # ── Student's Solution (bold 11pt) ─────────────────────────────
                if is_structured:
                    # Dynamic sol_box: calculate rows from QP image height
                    _q_loc = q.get("_loc")
                    _img_h_pt = 0
                    if _q_loc:
                        _img_h_pt = ((_q_loc.get("bottom_y", 0) - _q_loc.get("top_y", 0))
                                     / 72 * 72)  # keep as pt
                    _CW_PT_s = (CONTENT_WIDTH_CM - 1.0) * 28.35
                    _PAGE_H_s = 728.0; _HDR_s = 125; _SOL_LBL_s = 30
                    _SOL_ROW_s = 28; _SAFETY_s = 20; _IMG_THRESH_s = 380
                    # Estimate rendered image pt height from QP location
                    _qp_img_bytes = crop_question_png(qp_bytes, _q_loc) if _q_loc else None
                    _img_render_pt = 0
                    if _qp_img_bytes:
                        try:
                            _pil = Image.open(io.BytesIO(_qp_img_bytes))
                            _iw, _ih = _pil.size
                            _img_render_pt = (_ih / _iw) * _CW_PT_s
                        except Exception:
                            pass
                    _sol_separate = (_img_render_pt > _IMG_THRESH_s)
                    if _sol_separate:
                        _n_sol = 6
                    else:
                        _avail = _PAGE_H_s - _HDR_s - _img_render_pt - _SOL_LBL_s - _SAFETY_s
                        _n_sol = max(4, min(int(_avail / _SOL_ROW_s), 8))

                    if _sol_separate:
                        # Large question: sol on its own page
                        _pb2 = doc.add_paragraph()
                        _pb2.paragraph_format.page_break_before = True
                        _pb2.paragraph_format.space_before = Pt(0)
                        _pb2.paragraph_format.space_after  = Pt(0)

                    sl = doc.add_paragraph()
                    sl.paragraph_format.space_before = Pt(8 if not _sol_separate else 0)
                    sl.paragraph_format.space_after  = Pt(2)
                    sl.paragraph_format.keep_with_next = True
                    _run(sl, "Student's Solution:", bold=True, size_pt=11)
                    _add_solution_box(doc, n_lines=_n_sol)
                else:
                    sl = doc.add_paragraph()
                    sl.paragraph_format.space_before = Pt(8)
                    sl.paragraph_format.space_after  = Pt(2)
                    _run(sl, "Student's Solution:", bold=True, size_pt=11)
                    _add_solution_box(doc, n_lines=max(6, (marks or 1) * 2))

                # ── Page break → Answer from Mark Scheme on separate page ───────
                if is_structured:
                    _pb_ms = doc.add_paragraph()
                    _pb_ms.paragraph_format.page_break_before = True
                    _pb_ms.paragraph_format.space_before = Pt(0)
                    _pb_ms.paragraph_format.space_after  = Pt(0)

                # ── Answer from Mark Scheme ────────────────────────────────────
                ap = doc.add_paragraph()
                ap.paragraph_format.space_before = Pt(0)
                ap.paragraph_format.space_after  = Pt(4)
                ap.paragraph_format.keep_with_next = True
                _run(ap, "Answer from Mark Scheme:", bold=True, size_pt=11)

                # Structured mode: MS image (full worked solution) when present
                ms_img_bytes = q.get("_ms_image")
                if ms_img_bytes:
                    # ms_img_bytes can be raw bytes (short answer) or a list
                    # of bytes chunks (long answer split across pages).
                    if isinstance(ms_img_bytes, (bytes, bytearray)):
                        chunks_to_insert = [bytes(ms_img_bytes)]
                    else:
                        chunks_to_insert = list(ms_img_bytes)   # list of bytes

                    for part_idx, chunk_bytes in enumerate(chunks_to_insert):
                        # NO "continued" labels — just insert images sequentially
                        aip = doc.add_paragraph()
                        aip.paragraph_format.space_before = Pt(0)
                        aip.paragraph_format.space_after  = Pt(4)
                        # Wrap in a cantSplit table so Word won't break mid-image
                        _t = doc.add_table(rows=1, cols=1)
                        _t.autofit = False
                        _t.columns[0].width = Cm(CONTENT_WIDTH_CM - 0.5)
                        _cell = _t.rows[0].cells[0]
                        _trPr = _t.rows[0]._tr.get_or_add_trPr()
                        _cs = OxmlElement("w:cantSplit")
                        _cs.set(docx_qn("w:val"), "1")
                        _trPr.append(_cs)
                        _tcPr = _cell._tc.get_or_add_tcPr()
                        _tcB = OxmlElement("w:tcBorders")
                        for _side in ["top","left","bottom","right","insideH","insideV"]:
                            _b = OxmlElement(f"w:{_side}")
                            _b.set(docx_qn("w:val"), "nil")
                            _tcB.append(_b)
                        _tcPr.append(_tcB)
                        _tcMar = OxmlElement("w:tcMar")
                        for _s in ["top","left","bottom","right"]:
                            _m = OxmlElement(f"w:{_s}")
                            _m.set(docx_qn("w:w"), "0")
                            _m.set(docx_qn("w:type"), "dxa")
                            _tcMar.append(_m)
                        _tcPr.append(_tcMar)
                        _p2 = _cell.paragraphs[0]
                        _p2.paragraph_format.space_before = Pt(0)
                        _p2.paragraph_format.space_after  = Pt(4)
                        _p2.add_run().add_picture(
                            io.BytesIO(chunk_bytes),
                            width=Cm(CONTENT_WIDTH_CM - 0.5)
                        )
                else:
                    # No MS image available
                    av = doc.add_paragraph()
                    av.paragraph_format.space_before = Pt(0)
                    av.paragraph_format.space_after  = Pt(2)
                    if is_structured:
                        # Structured mode: always prefer image; if missing,
                        # show the Excel answer text (or a review notice).
                        ms_txt = str(answer or "").strip()
                        if ms_txt and ms_txt not in ("nan", "Answer not found - needs review"):
                            _run(av, ms_txt, size_pt=11)
                        else:
                            _run(av,
                                 "⚠ Mark Scheme image not available for this question — needs review",
                                 italic=True, size_pt=11)
                    else:
                        _run(av, answer, bold=True, size_pt=12)

                # ── Separator + Keep it up (after MS content) ──────────────────
                if quote:
                    _hr(doc, color="BFBFBF")
                    qup = doc.add_paragraph()
                    qup.paragraph_format.space_before = Pt(4)
                    qup.paragraph_format.space_after  = Pt(2)
                    _run(qup, "Keep it up", bold=True, size_pt=11)
                    _run(qup, f" : {quote}",       size_pt=11)
                elif not is_structured:
                    _hr(doc, color="BFBFBF")

                is_first_q_in_diff = False

            is_first_diff_in_chapter = False

        is_first_chapter = False

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN STREAMLIT FLOW
# ═══════════════════════════════════════════════════════════════════════════════
if st.button(
    "⚡ Extract & Generate Worksheet",
    type="primary",
    disabled=not (qp_file and ms_file and xl_file),
):
    client = anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])

    # ── File validation before reading ──────────────────────────────────────
    if not qp_file or not ms_file or not xl_file:
        st.warning("Please upload QP PDF, MS PDF, and Excel file first.")
        st.stop()

    qp_bytes = read_bytes(qp_file)
    ms_bytes = read_bytes(ms_file)

    if not qp_bytes or not ms_bytes:
        st.error("❌ Could not read one or more uploaded files. Please re-upload and try again.")
        st.stop()

    # ── Safety check: did the user accidentally upload the same file twice? ─
    if qp_bytes == ms_bytes:
        st.error(
            "❌ **The QP and MS files are identical!**\n\n"
            "You uploaded the same PDF in both the Question Paper and Mark "
            "Scheme slots. The worksheet would show the QP question in the "
            "'Answer from Mark Scheme' section — which is wrong.\n\n"
            "Please upload the actual **Mark Scheme PDF** (different from the "
            "QP) in the MS slot and try again."
        )
        st.stop()
    # Also catch near-identical files (same size + same first 1KB usually
    # means same source even after PDF re-export)
    if (abs(len(qp_bytes) - len(ms_bytes)) < 1024
            and qp_bytes[:1024] == ms_bytes[:1024]):
        st.warning(
            "⚠ The QP and MS files look very similar (same first bytes and "
            "almost identical size).  Are you sure you uploaded the correct "
            "Mark Scheme?  Continuing anyway, but the output may be wrong."
        )

    with st.status("Processing…", expanded=True) as status:

        # ── 1) Excel ─────────────────────────────────────────────────────────
        st.write("📊 Reading Excel (metadata only)…")
        try:
            xl_rows = parse_excel(xl_file)
        except Exception as e:
            st.error(f"Failed to read Excel: {e}")
            st.stop()
        if not xl_rows:
            st.error("No valid question rows found in Excel.")
            st.stop()

        st.write(f"✅ {len(xl_rows)} question rows in Excel")

        # ── Show how Excel columns were mapped (catches header mismatches) ──
        col_map = getattr(parse_excel, "last_column_map", {})
        if col_map:
            mapped_pairs = []
            unmapped = []
            for logical, (idx, actual) in col_map.items():
                if idx is None:
                    unmapped.append(logical)
                else:
                    mapped_pairs.append(f"**{logical}** → `{actual}`")
            with st.expander("🔍 Excel column mapping (click to verify)",
                             expanded=bool(unmapped)):
                st.write("Detected columns:")
                st.write(" · ".join(mapped_pairs))
                if unmapped:
                    st.error(
                        f"❌ Missing column(s): {', '.join(unmapped)}.  "
                        "Rename your Excel headers to one of the accepted "
                        "aliases (e.g. 'Page Number', 'PDF Page', 'Q#', "
                        "'Reference', 'Topic', 'Difficulty', 'Marks')."
                    )

        # Warn if many rows have page_num=0 (column mismatch or missing data)
        zero_page_count = sum(1 for r in xl_rows if not r.get("page_num"))
        if zero_page_count > len(xl_rows) * 0.1:   # >10% rows with no page
            st.warning(
                f"⚠ {zero_page_count}/{len(xl_rows)} rows have Page Number = 0.  "
                "This will cause poor matching with the QP. Check that the "
                "'Page Number' column in your Excel is populated, and that "
                "the header name is recognised (see column mapping above)."
            )

        # Show only TRUE duplicate removals
        dups = getattr(parse_excel, "last_duplicates", [])
        if dups:
            st.warning(
                f"⚠ Removed {len(dups)} truly-identical duplicate rows "
                "(same Q#, Reference, Page, Topic, and Difficulty)"
            )

        # ── 2) PyMuPDF coordinate detection ──────────────────────────────────
        st.write("📍 Locating questions in QP PDF…")
        try:
            locations = find_question_locations(qp_bytes)
        except Exception as e:
            st.error(f"Failed to scan QP PDF: {e}")
            st.stop()
        st.write(f"✅ Found {len(locations)} question marker(s) in QP across all pages")

        # Warn if Excel refers to pages beyond the QP PDF's actual size
        try:
            qp_pdf_pages = max(loc["page_idx"] for loc in locations.values()) + 1
        except (ValueError, KeyError):
            qp_pdf_pages = 0
        beyond_qp = [r for r in xl_rows
                     if r.get("page_num", 0) > qp_pdf_pages]
        if beyond_qp:
            sample_refs = sorted({r["ref"] for r in beyond_qp})[:5]
            st.warning(
                f"⚠ {len(beyond_qp)} Excel row(s) reference pages **beyond "
                f"the QP PDF's {qp_pdf_pages} pages**.  Those questions are "
                "from papers not included in the uploaded QP.\n\n"
                "Affected reference(s) (sample): "
                + ", ".join(f"`{r}`" for r in sample_refs)
                + (" …" if len(beyond_qp) > 5 else "")
                + "\n\nUpload a QP PDF that contains every paper referenced "
                "in your Excel — or remove those Excel rows."
            )

        # ── 2b) Detect Excel paper segments + per-segment page offsets ──────
        # If Excel page numbers don't line up with QP PDF page numbers
        # (e.g. Excel comes from a combined PDF with cover pages), we infer
        # a per-segment offset using each segment's Q1 location.
        # A new segment starts whenever:
        #   - the Reference string changes (different exam), OR
        #   - the Q# resets to 1 (next paper in the same session), OR
        #   - the Q# drops significantly (e.g. from Q12 back to Q3)
        excel_segments: list[list[int]] = []
        cur_seg: list[int] = []
        seg_max = 0
        seg_ref = None
        for i, r in enumerate(xl_rows):
            qn = r["qn"]
            ref = r.get("ref", "")
            new_segment = False
            if cur_seg:
                if ref != seg_ref:
                    new_segment = True
                elif qn == 1:
                    new_segment = True
                elif seg_max >= 10 and qn < seg_max - 5:
                    new_segment = True
            if new_segment:
                excel_segments.append(cur_seg)
                cur_seg = []
                seg_max = 0
            cur_seg.append(i)
            if qn > seg_max:
                seg_max = qn
            seg_ref = ref
        if cur_seg:
            excel_segments.append(cur_seg)

        seg_offsets = infer_segment_page_offsets(
            locations, xl_rows, excel_segments
        )
        # Build row → offset lookup
        row_offset = {}
        for seg_idx, seg_rows in enumerate(excel_segments):
            off = seg_offsets[seg_idx] if seg_idx < len(seg_offsets) else 0
            for ri in seg_rows:
                row_offset[ri] = off

        nonzero_offsets = [(i+1, o) for i, o in enumerate(seg_offsets) if o != 0]
        if nonzero_offsets:
            st.info(
                f"📐 Detected page offset(s) between Excel and QP PDF: "
                + ", ".join(f"segment {i} → {o:+d} pages"
                            for i, o in nonzero_offsets[:5])
                + (" …" if len(nonzero_offsets) > 5 else "")
            )

        # ── 3) Extract text + classify per-Excel-row using (qn, page) ───────
        # CRITICAL: We pass tuples of (q_num, excel_page) so the extractor can
        # disambiguate when the same Q# appears in multiple papers within one PDF.
        st.write("🔍 Extracting question text per row…")
        # Apply per-segment offset BEFORE handing pages to extractor
        row_tuples = []
        for i, r in enumerate(xl_rows):
            p = r.get("page_num", 0)
            off = row_offset.get(i, 0)
            if p and off:
                p = max(1, p + off)
            row_tuples.append((r["qn"], p))
        try:
            qp_data = classify_and_extract(
                None, "", row_tuples,
                qp_bytes=qp_bytes, locations=locations,
            )
        except Exception as e:
            st.error(f"QP extraction failed: {e}")
            st.stop()

        # qp_data is aligned 1:1 with xl_rows (same order)
        visual_n = sum(1 for q in qp_data if q.get("needs_image"))
        not_found_n = sum(1 for q in qp_data if not q.get("found"))
        st.write(f"✅ {visual_n} visual · {not_found_n} not found on specified page")

        is_structured = (mode == "Structured Mark Scheme Mode")

        # ── 4) Mark Scheme: branch by mode ───────────────────────────────────
        if is_structured:
            st.write("📐 Structured MS mode — locating worked solutions in MS PDF…")
            try:
                ms_struct_sections = find_ms_question_locations(ms_bytes)
            except Exception as e:
                st.error(f"Structured MS extraction failed: {e}")
                st.stop()
            st.write(
                f"✅ Found {len(ms_struct_sections)} MS section(s) with "
                f"{sum(len(s['questions']) for s in ms_struct_sections)} "
                "worked solutions total"
            )
            ms_sections = ms_struct_sections   # use shape compatible w/ rest
        else:
            st.write("🔑 MCQ mode — extracting answer letters from MS PDF…")
            try:
                ms_sections = extract_answers_pymupdf_per_section(ms_bytes)
            except Exception as e:
                st.error(f"MS extraction failed: {e}")
                st.stop()
            st.write(f"✅ Found {len(ms_sections)} answer section(s) in MS PDF")

        # Reuse the excel_segments computed earlier (with the same
        # Reference + Q-restart logic), so QP and MS matching stay aligned.
        segments = excel_segments

        st.write(f"📚 Using {len(segments)} paper segment(s) for MS matching")

        # ── Build row_idx → ms_section mapping: PURE POSITIONAL ────────────
        # Rule: QP paper i ↔ MS paper i (same ordinal position).
        # Algorithm:
        #   1. Sort Excel segments by their minimum QP page number (ascending).
        #   2. MS sections are already ordered by start_page from detection.
        #   3. Pair by index: sorted_segments[i] → ms_sections[i].
        # No paper code, no topic, no keyword used for matching.
        row_to_section: dict[int, dict | None] = {}
        row_to_section_idx: dict[int, int]     = {}

        def _seg_min_page(seg_rows):
            pgs = [xl_rows[ri].get("page_num", 0) for ri in seg_rows
                   if xl_rows[ri].get("page_num", 0) > 0]
            return min(pgs) if pgs else 999999

        # Sort segments by QP page order
        sorted_seg_indices = sorted(
            range(len(segments)),
            key=lambda i: _seg_min_page(segments[i])
        )
        seg_paper_codes = [None] * len(segments)   # kept for diagnostic table

        for pos, seg_idx in enumerate(sorted_seg_indices):
            seg_rows = segments[seg_idx]
            sec     = ms_sections[pos] if pos < len(ms_sections) else None
            ms_idx  = pos              if pos < len(ms_sections) else -1
            for ri in seg_rows:
                row_to_section[ri]     = sec
                row_to_section_idx[ri] = ms_idx

        unpaired_segs = []   # kept for diagnostic count

                # ── Diagnostic: paper / MS section matching table ──────────────────────
        # Safe rebuild of ms_code_to_section for diagnostic use only.
        # (The actual matching now uses the static QP-code map above.)
        try:
            ms_code_to_section = {}
            for _ms_i, _ms_sec in enumerate(ms_sections):
                _code = (_ms_sec.get("paper_code")
                         if isinstance(_ms_sec, dict) else None)
                if _code and _code not in ms_code_to_section:
                    ms_code_to_section[_code] = _ms_i
        except Exception:
            ms_code_to_section = {}

        try:
            paired_by_code = sum(1 for code in seg_paper_codes
                                  if code and code in ms_code_to_section)
        except Exception:
            paired_by_code = 0

        if is_structured:
            # Build per-segment matching table
            match_table = []
            for seg_idx, seg_rows in enumerate(segments):
                qp_code  = (seg_paper_codes[seg_idx]
                             if seg_idx < len(seg_paper_codes) else None)
                ms_sec_i = row_to_section_idx.get(seg_rows[0], -1)
                ms_sec   = row_to_section.get(seg_rows[0])
                ms_code  = (ms_sec.get("paper_code")
                             if isinstance(ms_sec, dict) else None)
                method   = "Positional"
                # Reference from first row of segment
                ref_str  = xl_rows[seg_rows[0]].get("ref", "") if seg_rows else ""
                match_table.append({
                    "Seg#":          seg_idx + 1,
                    "Excel Ref":     ref_str[:40],
                    "QP code":       qp_code or "—",
                    "MS section#":   ms_sec_i + 1 if ms_sec_i >= 0 else "—",
                    "MS code":       ms_code or "—",
                    "Method":        method,
                    "Questions":     len(seg_rows),
                    "MS Qs avail":   len(ms_sec.get("questions", {})) if ms_sec else 0,
                })
            with st.expander(
                f"📋 Paper ↔ MS Section Matching  "
                f"({paired_by_code} by code, "
                f"{len(unpaired_segs)} positional)",
                expanded=True,
            ):
                st.dataframe(match_table, use_container_width=True, hide_index=True)
                if paired_by_code == 0:
                    st.info(
                        "ℹ️ QP and MS use different code formats "
                        "(numeric vs alphanumeric). "
                        "All segments matched by position — this is normal "
                        "when QP and MS come from the same combined PDF."
                    )

        if len(ms_sections) != len(segments):
            ratio = len(ms_sections) / max(len(segments), 1)
            if ratio < 0.5 and is_structured:
                st.warning(
                    f"⚠ **MS PDF covers fewer papers than Excel needs.** "
                    f"Excel has {len(segments)} paper segment(s) but the MS "
                    f"PDF only contains {len(ms_sections)} section(s).\n\n"
                    "The app will attempt to match what it can using the "
                    "per-row topic-match validator. Rows whose MS topic "
                    "doesn't match the question will be flagged as "
                    "'MS answer mismatch' instead of being given a wrong "
                    "answer. Rows with no MS section will show "
                    "'Answer not found'.\n\n"
                    "For best results, upload an MS PDF that contains every "
                    "paper your Excel references."
                )
            elif len(ms_sections) != len(segments):
                st.warning(
                    f"⚠ Excel has {len(segments)} paper segment(s) but MS PDF "
                    f"yielded {len(ms_sections)} section(s). Some answers may be "
                    "missing or mismatched — review the validation table below."
                )

        # ── 5) Merge — every Excel row gets its own per-row extraction ──────
        st.write("🔗 Merging per row…")
        questions: list[dict] = []
        for i, (r, qp_q) in enumerate(zip(xl_rows, qp_data)):
            n = r["qn"]
            sec = row_to_section.get(i)
            sec_idx = row_to_section_idx.get(i, -1)
            found_q   = bool(qp_q) and qp_q.get("found") is not False

            # CHANGE 4 (Structured only): skip rows not found in QP when a
            # page number was provided — these are almost certainly Paper 2.
            # In MCQ mode we keep all rows so the user sees what's missing.
            if is_structured and not found_q and r.get("page_num", 0) > 0:
                continue

            # Special override: known wrong marks in Excel
            if r.get("qn") == 9 and r.get("page_num") == 147:
                r = dict(r); r["marks"] = 6

            # CHANGE 7 (Structured only): if marks=0, read "[Maximum mark: N]"
            # from the QP page. For MCQ mode, 0 simply falls back to 1.
            if r.get("marks", 0) == 0:
                if is_structured and qp_q.get("_loc"):
                    try:
                        _tmp = fitz.open(stream=qp_bytes, filetype="pdf")
                        _ptxt = _tmp[qp_q["_loc"]["page_idx"]].get_text()
                        _tmp.close()
                        _mm = re.search(
                            r'\[Maximum mark[:\s]+(\d+)\]', _ptxt, re.IGNORECASE
                        )
                        if _mm:
                            r = dict(r)       # shallow copy before mutating
                            r["marks"] = int(_mm.group(1))
                    except Exception:
                        pass
                # Still 0 → default to 1 (works for both modes)
                if r.get("marks", 0) == 0:
                    r = dict(r)
                    r["marks"] = 1
            # In Structured mode (Math, etc.) we always crop QP as image
            # because equations / diagrams / fractions can't be reliably
            # rendered as plain text.
            if is_structured:
                needs_img = found_q
            else:
                needs_img = found_q and qp_q.get("needs_image", False)

            sec_label = (f"MS section {sec_idx+1}"
                         + (f" (pp.{sec.get('start_page','?')}–"
                            f"{sec.get('end_page','?')})"
                            if sec else "")) if sec_idx >= 0 else "no section"

            if is_structured:
                # Structured: crop a per-question image from MS
                q_info = (sec["questions"].get(n) if sec and "questions" in sec
                          else None)
                ms_img = None
                ms_page_used = (q_info["page_idx"] + 1) if q_info else None
                ms_first_line = ""
                topic_match = True

                if q_info:
                    # FIX: Always use actual QP page text for topic check —
                    # the Excel topic string ("Chapter 1…") has zero overlap
                    # with MS keywords, causing 197 false mismatches.
                    try:
                        _doc = fitz.open(stream=ms_bytes, filetype="pdf")
                        ms_page_text = _doc[q_info["page_idx"]].get_text()
                        _doc.close()

                        # Always prefer real QP page text for topic comparison
                        qp_text_for_check = ""
                        if qp_q.get("_loc"):
                            try:
                                _doc2 = fitz.open(stream=qp_bytes, filetype="pdf")
                                qp_text_for_check = _doc2[qp_q["_loc"]["page_idx"]].get_text()
                                _doc2.close()
                            except Exception:
                                pass
                        # Last resort: Excel topic + quote
                        if not qp_text_for_check.strip():
                            qp_text_for_check = (r.get("topic", "") + " "
                                                  + (r.get("quote", "") or ""))

                        topic_match = _topics_match(qp_text_for_check, ms_page_text)

                        # Capture first non-trivial MS line for validation table
                        for ln in ms_page_text.split("\n"):
                            ln = ln.strip()
                            if len(ln) > 5 and not re.match(r"^[\-–\d\s/]+$", ln):
                                ms_first_line = ln[:80]
                                break
                    except Exception:
                        topic_match = True   # don't block on errors

                    # FIX: topic_match is a WARNING only — not a blocker.
                    # Requirement: "Topic يستخدم كتحذير فقط"
                    # Always crop the MS image when q_info is found.
                    ms_img = crop_ms_question_image(ms_bytes, q_info)

                ans_ok = ms_img is not None
                if ans_ok:
                    ans_text = ""
                elif q_info and not topic_match:
                    # Topic mismatch recorded for display — image still attempted
                    ans_text = "Answer not found - needs review"
                else:
                    ans_text = "Answer not found - needs review"
            else:
                # MCQ: extract a single letter
                ans = (sec["answers"].get(str(n), "") if sec else "")
                ans_ok = ans in ("A", "B", "C", "D")
                ans_text = ans if ans_ok else "Answer not found - needs review"
                ms_img = None
                ms_page_used = sec.get("start_page") if sec else None

            questions.append({
                **r,
                "_loc":        qp_q.get("_loc"),
                "_ms_section": sec_label,
                "_ms_image":   ms_img,           # only set in Structured mode
                "_ms_page":    ms_page_used,     # which MS PDF page was used
                "_ms_first":   ms_first_line if is_structured else "",
                "_topic_match": topic_match if is_structured else True,
                "_qp_page":    (qp_q["_loc"]["page_idx"] + 1
                                 if qp_q.get("_loc") else None),
                "_mode":       mode,
                "found":       found_q,
                "needs_image": needs_img,
                "text":        qp_q.get("text", "")  if found_q else "",
                "A":           qp_q.get("A", "")     if found_q else "",
                "B":           qp_q.get("B", "")     if found_q else "",
                "C":           qp_q.get("C", "")     if found_q else "",
                "D":           qp_q.get("D", "")     if found_q else "",
                "answer":      ans_text,
                "answerFound": ans_ok,
            })

        status.update(label="✅ Extraction done!", state="complete")

    # ── Validation summary ────────────────────────────────────────────────────
    st.subheader("Preview")
    total       = len(questions)
    found_qp    = sum(1 for q in questions if q["found"])
    found_ms    = sum(1 for q in questions if q["answerFound"])
    visual_cnt  = sum(1 for q in questions if q["needs_image"])

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total",         total)
    c2.metric("Found in QP",   found_qp)
    c3.metric("Answers found", found_ms)
    c4.metric("With visuals",  visual_cnt)

    st.dataframe(
        [{
            "Excel Row":  i + 2,
            "Q#":         q["qn"],
            "Reference":  q.get("ref", ""),
            "Excel Page": q.get("page_num", ""),
            "QP Page Used": q.get("_qp_page") or "—",
            "MS Page Used": q.get("_ms_page") or "—",
            "Mode":       "MCQ" if "MCQ" in (q.get("_mode") or "") else "Structured",
            "MS Section": q.get("_ms_section", ""),
            "Chapter":    q.get("topic", ""),
            "Difficulty": q.get("difficulty", ""),
            "Marks":      q.get("marks", ""),
            "QP Type":    "🖼 Image" if q["needs_image"] else "📝 Text",
            "MS Type":    ("🖼 Image" if q.get("_ms_image")
                            else "🔤 Letter" if q["answerFound"]
                            else "—"),
            "Topic Match": ("✅" if q.get("_topic_match", True)
                            else "❌ MISMATCH"),
            "First line of extracted":
                (q.get("text", "")[:60] + "…") if q.get("text") and len(q.get("text", "")) > 60
                else (q.get("text", "") or ("[image — see crop]" if q["needs_image"]
                      else "[not found]")),
            "First line of MS":   q.get("_ms_first", "")[:60],
            "Answer":     (q["answer"] if q["answerFound"] and not q.get("_ms_image")
                           else ("[image — see Word]" if q.get("_ms_image")
                                 else "—")),
            "Status":     ("✅" if q["found"] and q["answerFound"]
                           else ("⚠ MS missing" if q["found"]
                                 else ("⚠ QP missing" if q["answerFound"]
                                       else "❌ both missing"))),
        } for i, q in enumerate(questions)],
        use_container_width=True,
        hide_index=True,
    )

    # ── Unmatched rows table (problems that block a clean download) ──────────
    unmatched = []
    for i, q in enumerate(questions):
        problems = []
        if not q["found"]:
            if not q.get("page_num"):
                problems.append("no Page Number in Excel")
            else:
                problems.append(
                    f"Q{q['qn']} not located on page {q['page_num']} of QP"
                )
        if not q["answerFound"]:
            if q.get("_ms_image") is None and "Structured" in (q.get("_mode") or ""):
                problems.append(f"Q{q['qn']} not found in MS section")
            elif "MCQ" in (q.get("_mode") or ""):
                problems.append(f"Q{q['qn']} has no A/B/C/D answer in MS")
        if problems:
            unmatched.append({
                "Excel Row":      i + 2,
                "Reference":      q.get("ref", ""),
                "Q#":             q["qn"],
                "Page (Excel)":   q.get("page_num", 0),
                "Reason":         " · ".join(problems),
            })

    qp_match_pct = found_qp / max(total, 1)
    if unmatched:
        st.subheader(f"⚠ Unmatched rows ({len(unmatched)})")
        st.dataframe(unmatched, use_container_width=True, hide_index=True)

    # Gate the download when QP matching is clearly broken
    BAD_MATCH_THRESHOLD = 0.70    # need ≥ 70% rows located in QP
    matching_is_bad = qp_match_pct < BAD_MATCH_THRESHOLD

    # Also count topic-mismatches in Structured mode
    structured_rows = [q for q in questions
                       if "Structured" in (q.get("_mode") or "")
                       and q.get("_ms_page")]
    mismatch_rows = [q for q in structured_rows
                     if q.get("_topic_match") is False]
    mismatch_pct = (len(mismatch_rows) / max(len(structured_rows), 1)
                    if structured_rows else 0)

    # ── QP vs MS Sanity Check (first 10 questions) ───────────────────────────
    if is_structured:
        with st.expander("🔍 QP ↔ MS Sanity Check — first 10 questions", expanded=True):
            sanity_rows = []
            try:
                _doc_qp = fitz.open(stream=qp_bytes, filetype="pdf")
                _doc_ms = fitz.open(stream=ms_bytes, filetype="pdf")

                def _first_content(doc, pg_idx, skip_pat=None):
                    if pg_idx is None or pg_idx >= len(doc): return "—"
                    lines = (doc[pg_idx].get_text() or "").split("\n")
                    for l in lines:
                        ls = l.strip()
                        if len(ls) < 5: continue
                        if re.match(r"^(Turn over|Answers must|©|–\s*\d|M\d{2}/|N\d{2}/|"
                                    r"instructions|abbreviation|implied|method of mark)",
                                    ls, re.I): continue
                        return ls[:80]
                    return "—"

                shown = 0
                for q in questions:
                    if shown >= 10: break
                    qp_loc = q.get("_loc")
                    ms_img = q.get("_ms_image")
                    # Only include if both QP and MS are resolved
                    if not qp_loc: continue
                    qp_snip = _first_content(_doc_qp, qp_loc.get("page_idx"))
                    # MS page from question's assigned section
                    sec = row_to_section.get(questions.index(q))
                    ms_snip = "—"
                    if sec and isinstance(sec, dict):
                        qn = q.get("qn", 0)
                        q_info = sec.get("questions", {}).get(qn)
                        if q_info:
                            ms_snip = _first_content(_doc_ms, q_info.get("page_idx"))
                    sanity_rows.append({
                        "Q#": shown + 1,
                        "QP topic (first line)": qp_snip,
                        "MS answer (first line)": ms_snip,
                        "MS image": "✅" if ms_img else "⚠️ missing",
                    })
                    shown += 1

                _doc_qp.close(); _doc_ms.close()
            except Exception as _e:
                st.warning(f"Sanity check could not run: {_e}")

            if sanity_rows:
                st.dataframe(sanity_rows, use_container_width=True, hide_index=True)
                st.caption(
                    "Review QP topic vs MS answer — if they clearly differ "
                    "(e.g. QP is about trees but MS shows financial app), "
                    "the section pairing may be wrong. Contact support with the "
                    "session reference and question number."
                )

    # ── Build & Download ──────────────────────────────────────────────────────
    st.subheader("Download")

    if matching_is_bad:
        st.error(
            f"❌ Cannot generate worksheet: only {found_qp}/{total} "
            f"({qp_match_pct:.0%}) of rows were located in the QP PDF. "
            "Fix the matching before downloading."
        )
        st.info(
            "Most likely causes:\n"
            "- The QP PDF doesn't contain the papers your Excel rows refer "
            "to (e.g. Excel mixes Paper 1 + Paper 2 but only Paper 1 was "
            "uploaded). Look at **Unmatched rows** for which references "
            "are affected.\n"
            "- Excel **Page Number** column has wrong values or "
            "auto-converted dates. Review the **Excel column mapping** "
            "expander.\n"
            "- Wrong **Mode** selected for this subject."
        )
        st.stop()

    if mismatch_pct > 0.30 and len(structured_rows) >= 5:
        # CHANGE: topic mismatch is a WARNING only — never blocks download.
        # The static-map MS matching is reliable; high mismatch_pct usually
        # reflects keyword-overlap false-negatives in _topics_match, not
        # genuinely wrong answers.
        st.warning(
            f"⚠️ **Topic-match check flagged {len(mismatch_rows)}/{len(structured_rows)} rows.**\n\n"
            "This is a keyword-overlap heuristic and may over-report. "
            "Review the **Topic Match** column below — if the MS images look "
            "correct for their questions, you can proceed to download."
        )
    if qp_match_pct < 0.95:
        st.warning(
            f"⚠ {found_qp}/{total} rows matched ({qp_match_pct:.0%}). "
            f"{total - found_qp} unmatched rows will show "
            "'Question not found' in the worksheet. Review the **Unmatched "
            "rows** table above before downloading."
        )

    if visual_cnt:
        st.info(
            f"ℹ️ {visual_cnt} questions contain visual content — actual images "
            "will be cropped from the QP PDF and embedded in the Word file."
        )

    # ── Build Word document & store in session_state ────────────────────────
    ms_missing_n = sum(
        1 for q in questions
        if not q.get("_ms_image") and not q.get("answerFound")
    )

    _prog_bar  = st.progress(0)
    _prog_text = st.empty()

    def on_progress(done, total_, msg):
        _prog_bar.progress(done / max(total_, 1))
        _prog_text.text(msg)

    with st.spinner("Building Word document…"):
        try:
            _docx = build_word_document(
                questions, qp_bytes, locations,
                progress_cb=on_progress if visual_cnt else None,
            )
        except Exception as e:
            st.error(f"Word generation failed: {e}")
            st.stop()

    _prog_bar.progress(1.0)
    _prog_text.empty()

    # Persist in session_state — survives checkbox/rerun cycles
    st.session_state["ws_docx"]    = _docx
    st.session_state["ws_missing"] = ms_missing_n
    st.session_state["ws_stats"]   = (
        f"Total: {total} | Found in QP: {found_qp} | Answers found: {found_ms}"
    )
    st.session_state["ws_ready"]   = True
    st.success("✅ Worksheet built — scroll down to download.")


# ═══════════════════════════════════════════════════════════════════════════════
#  PERSISTENT DOWNLOAD — outside button block, survives all reruns
# ═══════════════════════════════════════════════════════════════════════════════
if st.session_state.get("ws_ready") and st.session_state.get("ws_docx"):
    st.divider()
    st.subheader("⬇️ Download Worksheet")

    _stats = st.session_state.get("ws_stats", "")
    if _stats:
        st.info(f"📊 {_stats}")

    _missing = st.session_state.get("ws_missing", 0)
    if _missing:
        st.warning(
            f"⚠️ {_missing} question(s) have no MS answer — "
            "they will show 'Answer not found' in the worksheet."
        )

    st.download_button(
        label="⬇️ Download Worksheet (.docx)",
        data=st.session_state["ws_docx"],
        file_name="IB_Worksheet.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        key="persistent_dl",
        type="primary",
    )

    if st.button("🔄 Start New Worksheet", key="reset_ws"):
        for _k in ["ws_docx", "ws_missing", "ws_stats", "ws_ready"]:
            st.session_state.pop(_k, None)
        st.rerun()
