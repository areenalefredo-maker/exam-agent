"""
IB Exam Worksheet Generator — Final Version
MCQ Mode + Structured Mark Scheme Mode
Based on: IB_Worksheet_Final (20) FINAL AGENT
"""
import streamlit as st
import anthropic
import openpyxl
import json
import re
import io
import base64
import datetime
import numpy as np

import fitz
from PIL import Image


def parse_page_range(raw_page: str, max_pages: int = 9999) -> tuple[int, int] | None:
    """Parse Excel page-number cell into (start_page, end_page).

    Handles:
      "91"          → (91, 91)    – plain page number
      "10-11"       → (10, 11)   – explicit range with dash
      "1011"        → (10, 11)   – merged-digit range
      "1920"        → (19, 20)   – merged-digit range
      "3032"        → (30, 32)   – merged-digit range
      "2026-10-11"  → (10, 11)   – Excel auto-converted date (month-day)
    Returns None when the value is unparseable.
    """
    if not raw_page:
        return None
    raw = str(raw_page).strip()

    # Explicit dash/en-dash range: "10-11", "10–11"
    m = re.match(r'^(\d+)\s*[-\u2013\u2014]\s*(\d+)$', raw)
    if m:
        s, e = int(m.group(1)), int(m.group(2))
        if 1 <= s and s <= e <= max_pages + 20:
            return (s, e)

    # Excel auto-converted date "YYYY-MM-DD" → take month and day
    m = re.match(r'^\d{4}-(\d{1,2})-(\d{1,2})$', raw)
    if m:
        s, e = int(m.group(1)), int(m.group(2))
        if 1 <= s <= max_pages and s <= e:
            return (s, e)

    # Plain integer
    m = re.match(r'^(\d+)$', raw)
    if m:
        n = int(m.group(1))
        if 1 <= n <= max_pages:
            return (n, n)
        # Attempt merged-range split: 1011 → 10-11, 1920 → 19-20, 3032 → 30-32
        s_raw = str(n)
        L = len(s_raw)
        best = None
        for split in range(1, L):
            s = int(s_raw[:split])
            e = int(s_raw[split:])
            if (1 <= s <= max_pages and s < e <= max_pages + 20
                    and 1 <= e - s <= 15):
                if best is None:
                    best = (s, e)
                else:
                    curr_diff = abs(split - (L - split))
                    prev_split = len(str(best[0]))
                    prev_diff  = abs(prev_split - (L - prev_split))
                    if curr_diff < prev_diff:
                        best = (s, e)
        return best  # None if no valid split found

    # Fallback: first integer in the cell
    m = re.search(r'\d+', raw)
    if m:
        n = int(m.group())
        if 1 <= n <= max_pages:
            return (n, n)
    return None

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
        "- Visual questions (graphs/diagrams/tables) → cropped image from QP\n"
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
        "Scheme is extracted as a cropped image from the MS PDF for each question."
    ),
)

st.divider()


# ═══════════════════════════════════════════════════════════════════════════════
#  SHARED HELPERS
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
    raise TypeError(f"Unsupported type: {type(uploaded_file)}")


def to_b64(data: bytes) -> str:
    return base64.standard_b64encode(data).decode()


def safe_json(text: str):
    text = re.sub(r"```json\s*", "", text)
    text = re.sub(r"```\s*", "", text)
    text = text.strip()
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


# ═══════════════════════════════════════════════════════════════════════════════
#  EXCEL PARSER
# ═══════════════════════════════════════════════════════════════════════════════
def parse_excel(uploaded_file) -> list[dict]:
    uploaded_file.seek(0)
    wb = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)
    ws = wb.active
    raw_headers = [c.value for c in next(ws.iter_rows(max_row=1))]
    headers = [str(h).strip() if h else "" for h in raw_headers]
    norm_headers = [re.sub(r"[\s_\-\.]+", " ", h.lower()).strip() for h in headers]

    def col_idx(*candidates):
        for name in candidates:
            target = re.sub(r"[\s_\-\.]+", " ", str(name).lower()).strip()
            for i, h in enumerate(norm_headers):
                if h == target:
                    return i
        for name in candidates:
            target = re.sub(r"[\s_\-\.]+", " ", str(name).lower()).strip()
            for i, h in enumerate(norm_headers):
                if target in h or h in target:
                    return i
        return None

    qn_idx   = col_idx("Question No.", "Question No", "Q No", "Q#", "Question Number", "Q", "QuestionNumber")
    page_idx = col_idx("Page Number", "Page", "page_number", "PDF Page", "PDF Page Number", "pdf_page", "QP Page", "Page No.", "Page No", "PageNum")
    top_idx  = col_idx("Topic", "Chapter", "Section", "Sub-topic", "Subtopic")
    dif_idx  = col_idx("Difficulty", "Level", "Hardness")
    mrk_idx  = col_idx("Marks", "Mark", "Points", "Score")
    qut_idx  = col_idx("Quote", "Note", "Motivational quote")
    ref_idx  = col_idx("Reference", "Ref", "Session", "Year", "Paper", "Exam Session", "Paper Reference")

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
        if not topic or topic.strip().lower() in ("error", "none", "n/a", "nan"):
            topic = "Unclassified"

        try:
            marks = int(float(cell_val(row, mrk_idx, "0")))
        except ValueError:
            marks = 0
        if marks < 0:
            marks = 0

        raw_page = ""
        if page_idx is not None:
            try:
                cell = list(row)[page_idx]
                v = cell.value
                if v is None:
                    raw_page = ""
                elif isinstance(v, datetime.datetime):
                    # Excel auto-converted "10-11" → date: recover month-day
                    raw_page = f"{v.month}-{v.day}"
                else:
                    raw_page = str(v).strip()
            except (IndexError, AttributeError):
                raw_page = ""

        page_num     = 0
        page_num_end = 0   # end page for multi-page questions
        if raw_page:
            pr = parse_page_range(raw_page)   # no max_pages at parse time
            if pr:
                page_num, page_num_end = pr
            else:
                # Could not interpret — keep 0 (will warn during generation)
                pass

        difficulty = cell_val(row, dif_idx, "Unspecified") or "Unspecified"
        ref        = cell_val(row, ref_idx, "")
        quote      = cell_val(row, qut_idx, "")

        ref = re.sub(r'\boctober\b', 'November', ref, flags=re.IGNORECASE)
        ref = re.sub(r'\boct\b',     'Nov',      ref, flags=re.IGNORECASE)

        # Paper 2 filter — only applies when caller enables it
        if getattr(parse_excel, '_filter_paper2', False):
            _P2_CODES = ('7205', '7210', '7215', '7310', '7315', '7320')
            if any(code in ref for code in _P2_CODES):
                continue
            if re.search(r'\bpaper[\s\-]?2\b|[/\-_](?:h|s)p2[/\-_]', ref, re.IGNORECASE):
                continue


        dedup_key = (q_num, ref.strip().lower(), page_num, topic.strip().lower(),
                     difficulty.strip().lower(), marks, quote.strip().lower())
        if dedup_key in seen_keys:
            duplicates.append((q_num, ref))
            continue
        seen_keys.add(dedup_key)

        rows.append({
            "qn": q_num, "page_num": page_num, "page_num_end": page_num_end,
            "topic": topic, "difficulty": difficulty, "marks": marks,
            "quote": quote, "ref": ref,
        })

    parse_excel.last_duplicates = duplicates
    return rows


# ═══════════════════════════════════════════════════════════════════════════════
#  QP LOCATION FINDER (used by MCQ mode)
# ═══════════════════════════════════════════════════════════════════════════════
def find_question_locations(pdf_bytes: bytes) -> dict:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    raw = []
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
    raw.sort(key=lambda e: (e[1], e[2]))

    doc2 = fitz.open(stream=pdf_bytes, filetype="pdf")
    content_bottom_per_page: dict = {}
    footer_pat = re.compile(
        r"(international\s+baccalaureate|turn\s+over|references?\s*:|"
        r"^\s*\d{4}\s*[\-\u2013]\s*\d{3,4}\s*$|"
        r"^\s*[\-\u2013]\s*\d{1,3}\s*[\-\u2013]\s*$)",
        re.IGNORECASE
    )
    bare_num_pat = re.compile(r"^\s*\d{1,3}\s*$")

    def is_writing_line_mcq(txt: str) -> bool:
        s = (txt or "").replace(" ", "").replace("\t", "")
        if len(s) < 8:
            return False
        return s.count(".") / max(len(s), 1) >= 0.8

    for pi in range(len(doc2)):
        page = doc2[pi]
        ph = page.rect.height
        td = page.get_text("dict")
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
                lines_info.append((bb[1], bb[3], full))

        def is_footer(y0, txt):
            if footer_pat.search(txt):
                return True
            if bare_num_pat.match(txt) and y0 > ph * 0.90:
                return True
            return False

        first_footer_y0 = ph * 0.92
        for y0, y1, txt in sorted(lines_info, key=lambda t: t[0]):
            if y0 < ph * 0.5:
                continue
            if is_footer(y0, txt):
                first_footer_y0 = min(first_footer_y0, y0)
                break

        first_writing_y0 = None
        for y0, y1, txt in sorted(lines_info, key=lambda t: t[0]):
            if is_writing_line_mcq(txt):
                first_writing_y0 = y0
                break

        upper_bound = first_footer_y0
        if first_writing_y0 is not None and first_writing_y0 < upper_bound:
            upper_bound = first_writing_y0

        content_bottom = 0
        for y0, y1, txt in lines_info:
            if is_footer(y0, txt) or is_writing_line_mcq(txt):
                continue
            if y1 < upper_bound - 2:
                if y1 > content_bottom:
                    content_bottom = y1

        if content_bottom == 0:
            content_bottom = ph * 0.88

        content_bottom_per_page[pi] = min(content_bottom + 8, upper_bound - 6)

    doc2.close()

    result = {}
    for i, (qn, page_idx, top_y, ph, pw, mrx) in enumerate(raw):
        bottom_y = content_bottom_per_page.get(page_idx, ph * 0.88)
        bottom_y = min(bottom_y, ph * 0.92)
        for j in range(i + 1, len(raw)):
            nq, np_idx, ntop, _, _, _ = raw[j]
            if np_idx == page_idx:
                bottom_y = ntop - 4
                break
            if np_idx > page_idx:
                break
        page_1based = page_idx + 1
        key = (page_1based, qn)
        if key in result:
            continue
        result[key] = {
            "page_idx": page_idx, "top_y": top_y, "bottom_y": bottom_y,
            "page_height": ph, "page_width": pw, "marker_right_x": mrx,
        }

    return result


def find_question_for_excel_row(locations, q_num, excel_page, tolerance=5):
    if excel_page and excel_page > 0:
        if (excel_page, q_num) in locations:
            return locations[(excel_page, q_num)]
        candidates = [
            (abs(p - excel_page), p)
            for (p, qn) in locations
            if qn == q_num and abs(p - excel_page) <= tolerance
        ]
        if candidates:
            candidates.sort()
            _, p = candidates[0]
            return locations[(p, q_num)]
        return None
    for (p, qn), loc in sorted(locations.items()):
        if qn == q_num:
            return loc
    return None


def infer_segment_page_offsets(locations, xl_rows, segments):
    offsets = []
    for seg_rows in segments:
        seg_offset = 0
        q1_rows = [xl_rows[ri] for ri in seg_rows if xl_rows[ri]["qn"] == 1]
        if q1_rows:
            excel_q1_page = q1_rows[0]["page_num"]
            if excel_q1_page > 0:
                q1_qp_pages = sorted(p for (p, qn) in locations if qn == 1)
                if q1_qp_pages:
                    nearest = min(q1_qp_pages, key=lambda p: abs(p - excel_q1_page))
                    if abs(nearest - excel_q1_page) < 20:
                        seg_offset = nearest - excel_q1_page
        offsets.append(seg_offset)
    return offsets


def crop_question_png(pdf_bytes: bytes, loc: dict, q_num: int = None, dpi: int = 200) -> bytes | None:
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
    mrx = loc.get("marker_right_x", 0)
    left_px = max(0, int((mrx + 6) * scale)) if mrx else 0
    pw = loc.get("page_width", 0) or (iw / scale)
    right_px = iw - int(pw * 0.05 * scale)
    if bottom_px <= top_px or right_px <= left_px:
        return None
    cropped = img.crop((left_px, top_px, right_px, bottom_px))
    buf = io.BytesIO()
    cropped.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
#  MCQ EXTRACTION FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════
def classify_and_extract(client, qp_b64, q_nums_or_rows, qp_bytes=None, locations=None):
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
        text_dict = page.get_text("dict", clip=clip)
        all_spans = []
        size_counts = {}
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
                    all_spans.append({"text": text, "size": sz, "x0": bb[0],
                                      "y0": bb[1], "y1": bb[3], "flags": span.get("flags", 0)})
                    if text.strip():
                        size_counts[sz] = size_counts.get(sz, 0) + len(text.strip())
        if not all_spans:
            results.append({"qn": q_num, "found": True, "needs_image": True,
                             "text": "", "A": "", "B": "", "C": "", "D": ""})
            continue
        body_size = max(size_counts.items(), key=lambda kv: kv[1])[0] if size_counts else 11.0
        SUBS = str.maketrans("0123456789+-=()n", "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₙ")
        SUPS = str.maketrans("0123456789+-=()n", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿ")
        unsafe_sub_super = False
        Y_BUCKET = 3
        line_buckets = {}
        for sp in all_spans:
            key = round(sp["y0"] / Y_BUCKET) * Y_BUCKET
            line_buckets.setdefault(key, []).append(sp)
        rendered_lines = []
        for key in sorted(line_buckets):
            spans_on_line = sorted(line_buckets[key], key=lambda s: s["x0"])
            full_size_spans = [s for s in spans_on_line if s["size"] >= body_size - 0.5]
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
                if is_smaller and sp["y1"] > main_y1 + 0.5:
                    if all(c in "0123456789+-=()n" for c in text.strip()):
                        parts.append(text.translate(SUBS))
                    else:
                        unsafe_sub_super = True
                        parts.append(text)
                elif is_smaller and sp["y0"] < main_y0 - 0.5:
                    if all(c in "0123456789+-=()n" for c in text.strip()):
                        parts.append(text.translate(SUPS))
                    else:
                        unsafe_sub_super = True
                        parts.append(text)
                else:
                    parts.append(text)
            rendered = "".join(parts)
            rendered = re.sub(r"[\u2009\u00a0]", " ", rendered).strip()
            if rendered:
                rendered_lines.append((key, spans_on_line[0]["x0"], rendered))
        opt_re = re.compile(r"^\s*([ABCD])\.\s*(.*)$")
        opt_markers = {}
        for y_key, x0, rendered in rendered_lines:
            m = opt_re.match(rendered)
            if m and x0 < 200:
                opt_markers[y_key] = m.group(1)
        needs_image = bool(unsafe_sub_super)
        stem_lines = []
        option_lines = {"A": [], "B": [], "C": [], "D": []}
        cur_opt = None
        for y_key, x0, rendered in rendered_lines:
            if y_key in opt_markers:
                cur_opt = opt_markers[y_key]
                m = opt_re.match(rendered)
                if m:
                    rest = m.group(2).strip()
                    if rest:
                        option_lines[cur_opt].append(rest)
                continue
            if cur_opt:
                option_lines[cur_opt].append(rendered)
            else:
                stem_lines.append(rendered)
        qn_pref = re.compile(r"^\d+\.\s*")
        if stem_lines and qn_pref.match(stem_lines[0]):
            stem_lines[0] = qn_pref.sub("", stem_lines[0]).strip()
        stem = "\n".join(stem_lines).strip()
        options = {L: " ".join(option_lines[L]).strip() for L in "ABCD"}

        def _has_chem_complexity(s):
            if not s:
                return False
            if re.search(r"[A-Z]{2,}", s) and re.search(r"\d", s):
                return True
            if re.search(r"[A-Z][a-z]?\d", s):
                return True
            if re.search(r"\b[A-Z][a-z]?\s+\d", s):
                return True
            if re.search(r"[=×]", s):
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
        non_empty_opts = [v for v in options.values() if v and len(v.strip()) >= 1]
        if len(non_empty_opts) < 4 or not stem:
            needs_image = True
        if needs_image:
            stem = ""
            options = {"A": "", "B": "", "C": "", "D": ""}
        results.append({
            "qn": q_num, "found": True, "needs_image": needs_image,
            "text": stem, "A": options["A"], "B": options["B"],
            "C": options["C"], "D": options["D"],
            "_loc": loc, "_excel_page": excel_page,
        })
    doc.close()
    return results


def _extract_qn_ans_pairs(text: str) -> list[tuple]:
    pattern = re.compile(
        r"(?:^|[\s\|])(?:Q\s*)?(\d{1,2})[\.\):]?\s*([ABCD])(?=$|[\s\.\,\;\|])"
    )
    pairs = []
    for m in pattern.finditer(text):
        qn = int(m.group(1))
        if 1 <= qn <= 99:
            pairs.append((qn, m.group(2)))
    return pairs


def extract_answers_pymupdf_per_section(ms_bytes: bytes) -> list[dict]:
    sections = []
    current = {"start_page": 1, "answers": {}, "max_qn": 0}
    try:
        doc = fitz.open(stream=ms_bytes, filetype="pdf")
    except Exception:
        return []
    for page_idx, page in enumerate(doc):
        text = page.get_text()
        page_pairs = _extract_qn_ans_pairs(text)
        if len(page_pairs) < 3:
            continue
        for qn, ans in page_pairs:
            qn_str = str(qn)
            is_boundary = (qn_str in current["answers"] or (qn == 1 and current["max_qn"] >= 5))
            if is_boundary and current["answers"]:
                sections.append({"start_page": current["start_page"], "end_page": page_idx + 1,
                                  "answers": current["answers"]})
                current = {"start_page": page_idx + 1, "answers": {}, "max_qn": 0}
            current["answers"][qn_str] = ans
            if qn > current["max_qn"]:
                current["max_qn"] = qn
    if current["answers"]:
        sections.append({"start_page": current["start_page"], "end_page": len(doc),
                          "answers": current["answers"]})
    doc.close()
    return [s for s in sections if s["answers"]]


def extract_answers_pymupdf(ms_bytes: bytes) -> dict:
    sections = extract_answers_pymupdf_per_section(ms_bytes)
    answers = {}
    for sec in sections:
        for q, a in sec["answers"].items():
            answers.setdefault(q, a)
    return answers


def extract_answers(client, ms_b64, q_nums, ms_bytes=None) -> dict:
    answers = {}
    if ms_bytes is not None:
        answers = extract_answers_pymupdf(ms_bytes)
    missing = [n for n in q_nums if str(n) not in answers]
    if missing and client is not None:
        nums = ", ".join(str(n) for n in missing)
        try:
            response = client.messages.create(
                model="claude-opus-4-5", max_tokens=1500,
                messages=[{"role": "user", "content": [
                    {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": ms_b64}},
                    {"type": "text", "text": (
                        f"Extract answers for Q{nums} from this IB mark scheme. "
                        'Return ONLY JSON: {"1":"D","2":"A",...}. '
                        'Use "NOT_FOUND" if absent.'
                    )},
                ]}]
            )
            raw = "".join(b.text for b in response.content if hasattr(b, "text"))
            data = safe_json(raw)
            if isinstance(data, dict):
                for k, v in data.items():
                    if k not in answers and v in ("A", "B", "C", "D"):
                        answers[k] = v
        except Exception:
            pass
    return answers


# ═══════════════════════════════════════════════════════════════════════════════
#  STRUCTURED MS — SECTION DETECTION
# ═══════════════════════════════════════════════════════════════════════════════
def _detect_paper_code(text: str) -> str | None:
    if not text:
        return None
    m = re.search(r'([MN])(\d{2})/\d/MATH\w+/([HS])P(\d)/\w+/(TZ\d|TZ0)', text)
    if m:
        session, year, level, paper, tz = m.groups()
        return f"{session}{year}/{level}P{paper}/{tz}"
    m = re.search(r'\b(\d{4})\s*[\-–]\s*(\d{4})\b', text)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    return None


def _build_page_to_paper_map(pdf_bytes: bytes) -> dict:
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


def find_ms_question_locations(ms_bytes: bytes) -> list[dict]:
    try:
        doc = fitz.open(stream=ms_bytes, filetype="pdf")
    except Exception:
        return []

    raw_markers = []
    pat_strict  = re.compile(r"^(\d{1,2})\.\s*$")
    pat_qa      = re.compile(r"^(\d{1,2})\.\s+\(?\w")
    pat_nodot   = re.compile(r"^(\d{1,2})\s*$")
    pat_qprefix = re.compile(r"^Q(\d{1,2})\.\s*(?:\(?\w|$)")

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
                if bb[0] > 80:
                    continue
                qn = None
                m = pat_strict.match(full_line) or pat_qa.match(full_line)
                if m:
                    qn = int(m.group(1))
                else:
                    mq = pat_qprefix.match(full_line)
                    if mq:
                        qn = int(mq.group(1))
                    else:
                        m2 = pat_nodot.match(full_line)
                        if m2 and bb[0] < 50:
                            font = spans[0].get("font", "")
                            size = spans[0].get("size", 0)
                            if ("Bold" in font or "bold" in font) and 9 <= size <= 14:
                                qn = int(m2.group(1))
                if qn is None or qn < 1 or qn > 30:
                    continue
                if re.match(rf"^{qn}\.\d", full_line):
                    continue
                marker_right_x = spans[0]["bbox"][2] if spans else (bb[0] + 20)
                raw_markers.append((qn, pi, bb[1], ph, marker_right_x))

    doc.close()
    if not raw_markers:
        return []

    sections = []
    current = {"questions": {}, "_order": [], "start_page": raw_markers[0][1] + 1, "max_qn": 0}
    for qn, pi, top_y, ph, mrx in raw_markers:
        if (qn in current["questions"] or (qn == 1 and current["max_qn"] >= 5)):
            if current["questions"]:
                sections.append(current)
                current = {"questions": {}, "_order": [], "start_page": pi + 1, "max_qn": 0}
        current["questions"][qn] = {
            "page_idx": pi, "top_y": top_y, "bottom_y": ph * 0.92,
            "end_page_idx": pi, "end_y": ph * 0.92,
            "page_height": ph, "marker_right_x": mrx,
        }
        current["_order"].append((qn, pi, top_y))
        if qn > current["max_qn"]:
            current["max_qn"] = qn

    if current["questions"]:
        sections.append(current)

    for sec in sections:
        order = sec["_order"]
        for idx, (qn, pi, top_y) in enumerate(order):
            if idx + 1 < len(order):
                next_qn, next_pi, next_top_y = order[idx + 1]
                if next_pi == pi:
                    sec["questions"][qn]["end_page_idx"] = pi
                    sec["questions"][qn]["end_y"] = next_top_y - 4
                else:
                    sec["questions"][qn]["end_page_idx"] = next_pi
                    sec["questions"][qn]["end_y"] = next_top_y - 4
        last_qn, last_pi, _ = order[-1]
        sec["end_page"] = sec["questions"][last_qn]["end_page_idx"] + 1
        sec.pop("_order", None)
        # Store max_qn for matching heuristic
        sec["max_qn"] = max(sec["questions"].keys()) if sec["questions"] else 0

    sections = [s for s in sections if len(s["questions"]) >= 5]

    instr_pat = re.compile(
        r"(instructions to examiners|abbreviations|marks? awarded for|"
        r"using the markscheme|method of marking|implied marks|misread|brackets in working)",
        re.IGNORECASE
    )
    try:
        doc2 = fitz.open(stream=ms_bytes, filetype="pdf")
        filtered = []
        for s in sections:
            start = s["start_page"] - 1
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

    page_codes = _build_page_to_paper_map(ms_bytes)
    for s in sections:
        first_page_idx = s["start_page"] - 1
        s["paper_code"] = page_codes.get(first_page_idx)

    return sections


# ═══════════════════════════════════════════════════════════════════════════════
#  STRUCTURED MODE — QP CROP (FINAL AGENT approach)
# ═══════════════════════════════════════════════════════════════════════════════
_S_DPI   = 150
_S_SCALE = _S_DPI / 72
_S_CW    = 16.0          # cm content width
_S_CW_PT = _S_CW * 28.35

# Continuation text patterns
_CONTINUES_PAT = re.compile(r'this question continues on the following page', re.I)
_CONTINUED_PAT = re.compile(r'question\s+\d+\s+continued', re.I)
_MS_CONT_PAT   = re.compile(r'question\s*\d*\s*continued\.?|this question continues', re.I)

_S_FOOTER_PAT  = re.compile(
    r"(M\d{2}/\d|N\d{2}/\d|\d{4}EP\d+|©\s*\d{4}|\bTurn over\b|\bPlease do not\b|"
    r"international\s+baccalaureate|\d{4}[\s\-–]+\d{4})", re.I)
_MS_FT  = re.compile(r"(M\d{2}/\d/|N\d{2}/\d/|©\s*\d{4}|international\s+baccalaureate)", re.I)
_MS_BN  = re.compile(r"^\s*\d{1,3}\s*$")
_MS_HDR = re.compile(r"(^\s*[–\-]\s*\d{1,3}\s*[–\-]\s*$|M\d{2}/\d/|N\d{2}/\d/)", re.I)


def _s_trim_bottom(img, thr=248):
    arr = np.array(img)
    rm  = arr.min(axis=(1, 2))
    dk  = np.where(rm < thr)[0]
    if len(dk) == 0:
        return img
    return img.crop((0, 0, img.width, min(int(dk[-1]) + 10, img.height)))


def _s_split_smart(img, max_h=1500, thr=250, win=250):
    img = _s_trim_bottom(img)
    w, h = img.size
    if h <= max_h:
        return [img]
    arr = np.array(img)
    rd  = (arr.min(axis=(1, 2)) < thr).astype(int)
    chunks = []
    prev = 0
    while prev < h:
        tgt = prev + max_h
        if tgt >= h:
            c = img.crop((0, prev, w, h))
            if c.height > 5:
                chunks.append(_s_trim_bottom(c))
            break
        lo = max(prev + 80, tgt - win)
        hi = min(h - 10, tgt + win)
        best = tgt
        bl = 0
        iw = False
        ws = 0
        for i, d in enumerate(rd[lo:hi]):
            ry = lo + i
            if d == 0:
                if not iw:
                    iw = True
                    ws = ry
            else:
                if iw:
                    rl = ry - ws
                    mid = (ws + ry) // 2
                    if rl > bl:
                        bl = rl
                        best = mid
                    iw = False
        if iw and (hi - ws) > bl:
            best = (ws + hi) // 2
        c = img.crop((0, prev, w, best))
        if c.height > 5:
            chunks.append(_s_trim_bottom(c))
        prev = best
    return chunks if chunks else [img]


def _s_is_writing_line(txt):
    if not txt:
        return False
    s = txt.replace(" ", "").replace("\t", "")
    if len(s) < 4:
        return False
    if s.count(".") / max(len(s), 1) >= 0.70:
        return True
    if len(s) >= 15 and s.count("_") / max(len(s), 1) >= 0.70:
        return True
    if len(s) >= 15 and s.count("-") / max(len(s), 1) >= 0.88:
        return True
    n_sp = sum(1 for c in s if ord(c) < 32 or ord(c) == 0xfffd or ord(c) == 0x08)
    if n_sp / max(len(s), 1) >= 0.50 and len(s) >= 5:
        return True
    n_na = sum(1 for c in s if ord(c) > 127)
    if n_na / max(len(s), 1) >= 0.70 and len(s) >= 10:
        return True
    return False


def _s_is_q_marker(txt, x0):
    """True only when txt is a QP question-number marker at x0 < 55pt."""
    if x0 > 55:
        return False
    m = re.match(r'^(\d{1,2})[.\s]', txt)
    if m and 1 <= int(m.group(1)) <= 20:
        return True
    return False


def _s_get_lines(page):
    r = []
    for block in page.get_text("dict").get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            txt = "".join(s.get("text", "") for s in spans).strip()
            if not txt:
                continue
            bb = line["bbox"]
            r.append((bb[0], bb[1], bb[3], txt))
    return sorted(r, key=lambda x: x[1])


def _s_find_answer_box_top(page, cs, se):
    pw = page.rect.width
    best = None
    for d in page.get_drawings():
        for item in d.get("items", []):
            if item[0] != "l":
                continue
            p1, p2 = item[1], item[2]
            if abs(p1.y - p2.y) > 2 or abs(p1.x - p2.x) < pw * 0.65:
                continue
            y = min(p1.y, p2.y)
            if y < cs - 2 or y > se:
                continue
            if best is None or y < best:
                best = y
    return best


def _crop_qp_page(qp_doc, pg_idx, qn, is_cont=False):
    page = qp_doc[pg_idx]
    ph = page.rect.height
    pw = page.rect.width
    lines = _s_get_lines(page)
    X_LEFT  = 30.0
    X_RIGHT = pw - 22.0

    if is_cont:
        hb = 0
        cb = 0
        for x0, y0, y1, txt in lines:
            if y0 > ph * 0.20:
                break
            if _S_FOOTER_PAT.search(txt) or re.match(r'^[\-–]\s*\d', txt):
                hb = max(hb, y1)
            if _CONTINUED_PAT.search(txt):
                cb = max(cb, y1)
        cs = max(hb + 4, cb + 4, ph * 0.07)
    else:
        qt = None
        qb = 0
        for x0, y0, y1, txt in lines:
            if not _s_is_q_marker(txt, x0):
                continue
            m = re.match(r'^(\d{1,2})[.\s]', txt)
            if m and int(m.group(1)) == qn and 1 <= qn <= 20 and y0 < ph * 0.88:
                qt = y0
                qb = y1
                break
        if qt is None:
            qt = ph * 0.07
            qb = ph * 0.12
        for x0, y0, y1, txt in lines:
            if y0 < qb - 2:
                continue
            if y0 > qb + 18:
                break
            if re.match(r'^\[\s*Maximum mark', txt, re.I):
                qb = max(qb, y1)
        cs = qb + 2

    nq = ph * 0.92
    if not is_cont:
        for x0, y0, y1, txt in lines:
            if y0 <= cs + 2:
                continue
            if not _s_is_q_marker(txt, x0):
                continue
            m = re.match(r'^(\d{1,2})[.\s]', txt)
            if m:
                n = int(m.group(1))
                if n > qn and 1 <= n <= 20 and y0 < ph * 0.88:
                    nq = y0
                    break

    cont_y = None
    for x0, y0, y1, txt in lines:
        if y0 < cs:
            continue
        if y0 > nq:
            break
        if _CONTINUES_PAT.search(txt):
            cont_y = y0
            break

    fw = None
    for x0, y0, y1, txt in lines:
        if y0 < cs - 2:
            continue
        if y0 >= (cont_y or nq):
            break
        if _s_is_writing_line(txt):
            fw = y0
            break

    se = cont_y or fw or nq
    ab = _s_find_answer_box_top(page, cs, se)

    if ab:
        ce = ab - 1
    else:
        lc = cs
        for x0, y0, y1, txt in lines:
            if y0 < cs - 2:
                continue
            if y0 >= se:
                break
            if _s_is_writing_line(txt):
                break
            if (_S_FOOTER_PAT.search(txt) or _CONTINUES_PAT.search(txt)
                    or _CONTINUED_PAT.search(txt)):
                continue
            lc = max(lc, y1)
        ce = min(lc + 16, se - 2)

    ce = min(ce, ph * 0.91)
    if ce <= cs + 5:
        return None, bool(cont_y)

    has_cont = bool(cont_y) or bool(_CONTINUES_PAT.search(page.get_text()))
    clip = fitz.Rect(X_LEFT, cs, X_RIGHT, ce)
    pix  = page.get_pixmap(matrix=fitz.Matrix(_S_SCALE, _S_SCALE), clip=clip, alpha=False)
    img  = Image.open(io.BytesIO(pix.tobytes("png")))
    return img, has_cont


def _crop_qp_full(qp_doc, pg, qn, pg_end: int = 0):
    """Crop QP question qn from page pg (through pg_end if it spans pages).

    pg_end:  last page of this question (inclusive).  0 → auto-detect via
             'continues on following page' text in the PDF.
    """
    if not pg or pg > len(qp_doc):
        return None

    img1, hc = _crop_qp_page(qp_doc, pg - 1, qn, False)
    if img1 is None:
        return None

    # ── If Excel told us the page range, use it directly ─────────────────────
    if pg_end and pg_end > pg:
        pieces = [img1]
        for pn in range(pg + 1, pg_end + 1):
            if pn > len(qp_doc):
                break
            ic, _ = _crop_qp_page(qp_doc, pn - 1, qn, True)
            if ic is not None:
                pieces.append(ic)
        tw = max(p.size[0] for p in pieces)
        th = sum(p.size[1] for p in pieces)
        c  = Image.new("RGB", (tw, th), "white")
        y  = 0
        for p in pieces:
            c.paste(p, (0, y))
            y += p.size[1]
        return c

    # ── Auto-detect via "continues on following page" ─────────────────────────
    if not hc:
        return img1
    pieces = [img1]
    pn = pg
    for _ in range(3):
        pn += 1
        if pn > len(qp_doc):
            break
        ic, hc2 = _crop_qp_page(qp_doc, pn - 1, qn, True)
        if ic is None:
            break
        pieces.append(ic)
        if not hc2:
            break
    tw = max(p.size[0] for p in pieces)
    th = sum(p.size[1] for p in pieces)
    c  = Image.new("RGB", (tw, th), "white")
    y  = 0
    for p in pieces:
        c.paste(p, (0, y))
        y += p.size[1]
    return c


# ═══════════════════════════════════════════════════════════════════════════════
#  STRUCTURED MODE — MS CROP (FINAL AGENT approach)
# ═══════════════════════════════════════════════════════════════════════════════
def _ms_hdr_bot(pg):
    ph = pg.rect.height
    hb = 0
    for block in pg.get_text("dict").get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            txt = "".join(s.get("text", "") for s in spans).strip()
            bb  = line["bbox"]
            if bb[1] > ph * 0.20:
                continue
            if _MS_HDR.search(txt) and bb[3] > hb:
                hb = bb[3]
            if _MS_CONT_PAT.search(txt) and bb[3] > hb:
                hb = bb[3]
    return hb


def _ms_ft_top(pg):
    ph = pg.rect.height
    fy = ph * 0.95
    for block in pg.get_text("dict").get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            txt = "".join(s.get("text", "") for s in spans).strip()
            bb  = line["bbox"]
            if bb[1] < ph * 0.85:
                continue
            if (_MS_FT.search(txt) or (_MS_BN.match(txt) and bb[1] > ph * 0.93)) and bb[1] < fy:
                fy = bb[1]
    return fy


def _last_ms_y(pg, top, hard):
    ph   = pg.rect.height
    last = top
    for block in pg.get_text("dict").get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            txt = "".join(s.get("text", "") for s in spans).strip()
            if not txt:
                continue
            bb = line["bbox"]
            if bb[1] < top - 2 or bb[1] >= hard - 2:
                continue
            if _MS_FT.search(txt) or (_MS_BN.match(txt) and bb[1] > ph * 0.93):
                continue
            if _MS_CONT_PAT.search(txt):
                continue
            if bb[3] > last:
                last = bb[3]
    return last


def _crop_ms_pieces(ms_doc, q_info):
    spi = q_info["page_idx"]
    epi = q_info.get("end_page_idx", spi)
    sy  = q_info["top_y"]
    ey  = q_info["end_y"]
    mrx = q_info.get("marker_right_x", 0)
    final = []
    for pi in range(spi, epi + 1):
        if pi >= len(ms_doc):
            break
        pg = ms_doc[pi]
        ph = pg.rect.height
        pw = pg.rect.width
        pix = pg.get_pixmap(matrix=fitz.Matrix(_S_SCALE, _S_SCALE), alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        iw, ih = img.size
        ft = _ms_ft_top(pg)
        pb = max(ph * 0.5, ft - 6)
        rx = iw - int(pw * 0.06 * _S_SCALE)
        lx = max(0, int((mrx + 6) * _S_SCALE)) if (pi == spi and mrx) else int(pw * 0.04 * _S_SCALE)
        hb = _ms_hdr_bot(pg)
        pt = max(ph * 0.06, hb + 6) if hb else ph * 0.06
        if pi == spi == epi:
            ty = sy
            by = min(ey, pb)
        elif pi == spi:
            ty = sy
            lc = _last_ms_y(pg, sy, pb)
            by = min(lc + 10, pb)
        elif pi == epi:
            ty = pt
            lc2 = _last_ms_y(pg, pt, min(ey, pb))
            by = min(lc2 + 10, ey, pb)
        else:
            ty = pt
            lc3 = _last_ms_y(pg, pt, pb)
            by = min(lc3 + 10, pb)
        tp = max(0, int(ty * _S_SCALE) - 6)
        bp = min(ih, int(by * _S_SCALE) + 6)
        if bp > tp + 20 and rx > lx:
            piece = img.crop((lx, tp, rx, bp))
            piece = _s_trim_bottom(piece)
            if piece.height > 10:
                final.extend(_s_split_smart(piece, 1500))
    return final if final else None


# ═══════════════════════════════════════════════════════════════════════════════
#  STRUCTURED MODE — WORD GENERATION (FINAL AGENT layout)
# ═══════════════════════════════════════════════════════════════════════════════
_PAGE_H_PT  = 728.0   # A4 content height (2 cm margins)
_HEADER_PT  = 125     # conservative header estimate
_SOL_LBL_PT = 30      # "Student's Solution:" label height
_SOL_ROW_PT = 28      # 28pt per sol-box row
_SAFETY_PT  = 20      # safety margin
_IMG_THRESH = 380     # pt — images taller than this → sol on separate page


def _sol_config(img):
    """Returns (n_rows, sol_on_separate_page)."""
    if img is None:
        return (6, False)
    img_pt = (img.height / img.width) * _S_CW_PT
    if img_pt > _IMG_THRESH:
        return (6, True)
    avail = _PAGE_H_PT - _HEADER_PT - img_pt - _SOL_LBL_PT - _SAFETY_PT
    n = max(4, min(int(avail / _SOL_ROW_PT), 8))
    return (n, False)


def _s_run(p, txt, bold=False, italic=False, size_pt=11, color=None):
    r = p.add_run(str(txt))
    r.bold       = bold
    r.italic     = italic
    r.font.size  = Pt(size_pt)
    r.font.name  = "Arial"
    if color:
        r.font.color.rgb = RGBColor(*bytes.fromhex(color))
    return r


def _s_hr(doc, color="CCCCCC"):
    p    = doc.add_paragraph()
    pBdr = OxmlElement("w:pBdr")
    bot  = OxmlElement("w:bottom")
    bot.set(docx_qn("w:val"),   "single")
    bot.set(docx_qn("w:sz"),    "6")
    bot.set(docx_qn("w:color"), color)
    pBdr.append(bot)
    pPr = p._p.get_or_add_pPr()
    idx = len(pPr)
    for j, c in enumerate(pPr):
        if c.tag.split('}')[-1] in ('spacing', 'ind', 'jc', 'rPr', 'sectPr'):
            idx = j
            break
    pPr.insert(idx, pBdr)


def _s_sol_box(doc, n):
    n = max(4, min(n, 8))
    t = doc.add_table(rows=n, cols=1)
    t.autofit = False
    for c in t.columns[0].cells:
        c.width = Cm(_S_CW)
    for row in t.rows:
        trPr = row._tr.get_or_add_trPr()
        trH  = OxmlElement("w:trHeight")
        trH.set(docx_qn("w:val"), "560")   # 28pt
        trPr.append(trH)
        cell = row.cells[0]
        tcPr = cell._tc.get_or_add_tcPr()
        tcB  = OxmlElement("w:tcBorders")
        for s, v, sz, col in [("top", "nil", "0", "auto"), ("left", "nil", "0", "auto"),
                               ("right", "nil", "0", "auto"), ("bottom", "single", "6", "BFBFBF")]:
            b = OxmlElement(f"w:{s}")
            b.set(docx_qn("w:val"), v)
            if v != "nil":
                b.set(docx_qn("w:sz"),    sz)
                b.set(docx_qn("w:color"), col)
            tcB.append(b)
        tcPr.append(tcB)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)


def _s_add_img(doc, pil_img):
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    buf.seek(0)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(4)
    p.add_run().add_picture(buf, width=Cm(_S_CW))


def _s_add_ms_img(doc, pil_img):
    """Add MS image wrapped in cantSplit table — no mid-image page breaks."""
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    buf.seek(0)
    t = doc.add_table(rows=1, cols=1)
    t.autofit = False
    t.columns[0].width = Cm(_S_CW)
    cell = t.rows[0].cells[0]
    trPr = t.rows[0]._tr.get_or_add_trPr()
    cs   = OxmlElement("w:cantSplit")
    cs.set(docx_qn("w:val"), "1")
    trPr.append(cs)
    tcPr = cell._tc.get_or_add_tcPr()
    tcB  = OxmlElement("w:tcBorders")
    for side in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        b = OxmlElement(f"w:{side}")
        b.set(docx_qn("w:val"), "nil")
        tcB.append(b)
    tcPr.append(tcB)
    tcMar = OxmlElement("w:tcMar")
    for s in ["top", "left", "bottom", "right"]:
        m = OxmlElement(f"w:{s}")
        m.set(docx_qn("w:w"),    "0")
        m.set(docx_qn("w:type"), "dxa")
        tcMar.append(m)
    tcPr.append(tcMar)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(4)
    p.add_run().add_picture(buf, width=Cm(_S_CW))


def _normalize_topic(t: str) -> str:
    if not t or str(t).strip().lower() in ("error", "none", "nan", ""):
        return "Unclassified"
    s = re.sub(r'^#+\s*', '', str(t).strip())
    m = re.match(r"^\s*chapter\s+(\d+)\s*[:\.\-–]?\s*(.+?)\s*$", s, re.IGNORECASE)
    if m:
        return f"{m.group(1)}. {m.group(2).strip()}"
    m = re.match(r"^\s*(\d+)\s*[\.:\-–]\s*(.+?)\s*$", s)
    if m:
        return f"{m.group(1)}. {m.group(2).strip()}"
    return re.sub(r"\s+", " ", s)


def _chapter_sort_key(topic: str):
    m = re.match(r"^\s*(\d+)\s*[\.\-:]", str(topic))
    if m:
        return (0, int(m.group(1)), str(topic))
    return (1, 0, str(topic).lower())


DIFF_ORDER = ["Easy", "Medium", "Hard"]


def build_structured_worksheet(
    xl_rows: list[dict],
    qp_bytes: bytes,
    ms_bytes: bytes,
    ms_sections: list[dict],
    row_to_section: dict,
    progress_cb=None,
) -> bytes:
    """Generate the structured worksheet using the FINAL AGENT layout."""

    qp_doc = fitz.open(stream=qp_bytes, filetype="pdf")
    ms_doc = fitz.open(stream=ms_bytes, filetype="pdf")

    doc = Document()
    for s in doc.sections:
        s.top_margin    = Cm(2)
        s.bottom_margin = Cm(2)
        s.left_margin   = Cm(2)
        s.right_margin  = Cm(2)

    # ── Group by normalized topic → difficulty (preserve Excel row order within) ──
    for r in xl_rows:
        r["_topic_norm"] = _normalize_topic(r.get("topic", ""))

    grouped = {}
    for ri, r in enumerate(xl_rows):
        t = r["_topic_norm"]
        d = r.get("difficulty") or "Unspecified"
        grouped.setdefault(t, {}).setdefault(d, []).append(ri)

    sorted_topics = sorted(grouped.keys(), key=_chapter_sort_key)

    total_qs = len(xl_rows)
    done_qs  = 0
    first    = True
    ws_q     = 0
    cur_t    = None
    cur_d    = None

    for topic in sorted_topics:
        diffs = grouped[topic]
        sorted_diffs = sorted(
            diffs.keys(),
            key=lambda d: DIFF_ORDER.index(d) if d in DIFF_ORDER else 99
        )

        for diff in sorted_diffs:
            ri_list = diffs[diff]
            for ri in ri_list:
                ws_q += 1
                r    = xl_rows[ri]
                pg   = r.get("page_num", 0)
                qn   = r.get("qn", 0)
                marks = r.get("marks", 1) or 1
                quote = r.get("quote", "") or ""
                if quote in ("nan",):
                    quote = ""

                # Special override: Q9 page 147
                if qn == 9 and pg == 147:
                    marks = 6

                # Crop QP (pass pg_end so multi-page questions are complete)
                pg_end = r.get("page_num_end", 0) or 0
                qp_img = _crop_qp_full(qp_doc, pg, qn, pg_end=pg_end)

                # Crop MS
                sec     = row_to_section.get(ri)
                ms_imgs = []
                if sec and isinstance(sec, dict):
                    q_info = sec.get("questions", {}).get(qn)
                    if q_info:
                        pieces = _crop_ms_pieces(ms_doc, q_info)
                        if pieces:
                            ms_imgs = pieces

                # sol config
                n_sol, sol_separate = _sol_config(qp_img)

                # ── Chapter heading (on same page as Q — no separate break) ──
                if topic != cur_t:
                    h = doc.add_paragraph()
                    if not first:
                        pass   # page break is on qh below
                    h.paragraph_format.space_before = Pt(0)
                    h.paragraph_format.space_after  = Pt(6)
                    _s_run(h, topic, bold=True, size_pt=20)
                    cur_t = topic
                    cur_d = None

                if diff != cur_d:
                    dp = doc.add_paragraph()
                    dp.paragraph_format.space_before = Pt(6)
                    dp.paragraph_format.space_after  = Pt(4)
                    _s_run(dp, f"— {diff} —", bold=True, italic=True, size_pt=13)
                    cur_d = diff

                # ── Question header (ONE page break per question) ──────────
                qh = doc.add_paragraph()
                if not first:
                    qh.paragraph_format.page_break_before = True
                    # Move topic/diff headings before qh if they were just added
                    # (they'll flow after qh page break)
                    # Restructure: insert topic/diff before qh in XML
                    if topic != xl_rows[ri_list[0] if ri == ri_list[0] else ri]["_topic_norm"] if ri != ri_list[0] else False:
                        pass
                qh.paragraph_format.space_before = Pt(0)
                qh.paragraph_format.space_after  = Pt(2)
                _s_run(qh, f"Question: {ws_q}", bold=True, size_pt=13)

                mp = doc.add_paragraph()
                mp.paragraph_format.space_before = Pt(0)
                mp.paragraph_format.space_after  = Pt(2)
                _s_run(mp, "Level of question", bold=True, size_pt=11)
                _s_run(mp, f": {diff}  |  ", size_pt=11)
                _s_run(mp, "Number of Marks: ", bold=True, size_pt=11)
                _s_run(mp, f"{marks}", size_pt=11)

                cp = doc.add_paragraph()
                cp.paragraph_format.space_before = Pt(0)
                cp.paragraph_format.space_after  = Pt(2)
                _s_run(cp, "Chapter", bold=True, size_pt=11)
                _s_run(cp, f" :  {topic}", size_pt=11)
                _s_hr(doc, "BFBFBF")

                # ── QP image ──────────────────────────────────────────────
                if qp_img:
                    _s_add_img(doc, qp_img)
                else:
                    p = doc.add_paragraph()
                    _s_run(p, "[Question image not available]", italic=True, size_pt=11)

                # ── Student's Solution ────────────────────────────────────
                if sol_separate:
                    pb2 = doc.add_paragraph()
                    pb2.paragraph_format.page_break_before = True
                    pb2.paragraph_format.space_before = Pt(0)
                    pb2.paragraph_format.space_after  = Pt(0)

                sl = doc.add_paragraph()
                sl.paragraph_format.space_before = Pt(0 if sol_separate else 8)
                sl.paragraph_format.space_after  = Pt(4)
                sl.paragraph_format.keep_with_next = True
                _s_run(sl, "Student's Solution:", bold=True, size_pt=11)
                _s_sol_box(doc, n_sol)

                # ── MS page (separate page) ───────────────────────────────
                pb = doc.add_paragraph()
                pb.paragraph_format.page_break_before = True
                pb.paragraph_format.space_before = Pt(0)
                pb.paragraph_format.space_after  = Pt(0)

                ap = doc.add_paragraph()
                ap.paragraph_format.space_before = Pt(0)
                ap.paragraph_format.space_after  = Pt(4)
                ap.paragraph_format.keep_with_next = True
                _s_run(ap, "Answer from Mark Scheme:", bold=True, size_pt=11)

                if ms_imgs:
                    for img in ms_imgs:
                        _s_add_ms_img(doc, img)
                else:
                    p = doc.add_paragraph()
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after  = Pt(2)
                    _s_run(p, "⚠ Mark Scheme image not available — needs review",
                           italic=True, size_pt=11)

                # ── Keep it up (after MS, with HR) ─────────────────────────
                if quote:
                    _s_hr(doc, "BFBFBF")
                    kp = doc.add_paragraph()
                    kp.paragraph_format.space_before = Pt(4)
                    kp.paragraph_format.space_after  = Pt(0)
                    _s_run(kp, "Keep it up", bold=True, size_pt=11)
                    _s_run(kp, f" : {quote}", size_pt=11)

                first   = False
                done_qs += 1
                if progress_cb:
                    progress_cb(done_qs, total_qs, f"Building Q{ws_q}…")

    qp_doc.close()
    ms_doc.close()

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
#  MCQ WORD BUILDER (unchanged from original)
# ═══════════════════════════════════════════════════════════════════════════════
CONTENT_WIDTH_CM = 17.0


def _set_cell_borders(cell, top="single", left="single", right="single",
                      bottom="single", color="999999",
                      bot_color=None, top_color=None, sz_top=8, sz_bot=8):
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
    r.bold       = bold
    r.italic     = italic
    r.font.size  = Pt(size_pt)
    r.font.name  = "Arial"
    if color:
        r.font.color.rgb = RGBColor(*bytes.fromhex(color))
    return r


def _insert_pBdr(p, pBdr):
    pPr = p._p.get_or_add_pPr()
    insert_idx = len(pPr)
    for i, child in enumerate(pPr):
        tag = child.tag.split('}')[-1]
        if tag in ('spacing', 'ind', 'jc', 'contextualSpacing', 'mirrorIndents', 'rPr', 'sectPr'):
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


def _add_solution_box(doc, n_lines: int = 4):
    table = doc.add_table(rows=n_lines, cols=1)
    table.autofit = False
    for cell in table.columns[0].cells:
        cell.width = Cm(CONTENT_WIDTH_CM)
    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        trH  = OxmlElement("w:trHeight")
        trH.set(docx_qn("w:val"), "510")
        trPr.append(trH)
        cell = row.cells[0]
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = OxmlElement("w:tcBorders")
        for side, val, sz, color in [
            ("top",    "nil",    "0", "auto"),
            ("left",   "nil",    "0", "auto"),
            ("right",  "nil",    "0", "auto"),
            ("bottom", "single", "4", "BFBFBF"),
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


def build_word_document(questions, qp_bytes=None, locations=None, progress_cb=None) -> bytes:
    """MCQ mode word builder."""
    doc = Document()
    for section in doc.sections:
        section.top_margin    = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin   = Cm(2.0)
        section.right_margin  = Cm(2.0)

    def _normalize_topic_mcq(t: str) -> str:
        if not t or str(t).strip().lower() in ("error", "none", "nan", ""):
            return "Unclassified"
        s = str(t).strip()
        s = re.sub(r'^#+\s*', '', s).strip()
        m = re.match(r"^\s*chapter\s+(\d+)\s*[:\.\-–]?\s*(.+?)\s*$", s, re.IGNORECASE)
        if m:
            return f"{m.group(1)}. {m.group(2).strip()}"
        m = re.match(r"^\s*(\d+)\s*[\.:\-–]\s*(.+?)\s*$", s)
        if m:
            return f"{m.group(1)}. {m.group(2).strip()}"
        return re.sub(r"\s+", " ", s)

    for q in questions:
        q["topic"] = _normalize_topic_mcq(q.get("topic"))

    grouped = {}
    for q in questions:
        t = q["topic"]
        d = q.get("difficulty") or "Unspecified"
        grouped.setdefault(t, {}).setdefault(d, []).append(q)

    sorted_topics = sorted(grouped.keys(), key=_chapter_sort_key)

    visual_qs  = [q for q in questions if q.get("needs_image")]
    total_imgs = len(visual_qs)
    img_done   = 0
    is_first_chapter = True
    display_counter  = 0

    for topic in sorted_topics:
        diffs = grouped[topic]
        h = doc.add_paragraph()
        if not is_first_chapter:
            h.paragraph_format.page_break_before = True
        h.paragraph_format.space_before = Pt(0)
        h.paragraph_format.space_after  = Pt(8)
        _run(h, topic, bold=True, size_pt=20)

        sorted_diffs = sorted(
            diffs.keys(),
            key=lambda d: DIFF_ORDER.index(d) if d in DIFF_ORDER else 99
        )

        is_first_diff_in_chapter = True
        for diff_ in sorted_diffs:
            qs_in_diff = diffs[diff_]
            dh = doc.add_paragraph()
            if not is_first_diff_in_chapter:
                dh.paragraph_format.page_break_before = True
            dh.paragraph_format.space_before = Pt(8)
            dh.paragraph_format.space_after  = Pt(6)
            _run(dh, f"— {diff_} —", bold=True, italic=True, size_pt=13)

            is_first_q_in_diff = True
            for q in qs_in_diff:
                display_counter += 1
                q_num  = q["qn"]
                marks  = q.get("marks", 1) or 1
                quote  = q.get("quote", "") or ""
                answer = q.get("answer", "")
                found  = q.get("found", False)
                vis    = q.get("needs_image", False)
                text   = q.get("text", "")
                opts   = {L: q.get(L, "") for L in "ABCD"}
                topic_ = q.get("topic", topic)

                qh = doc.add_paragraph()
                if not is_first_q_in_diff:
                    qh.paragraph_format.page_break_before = True
                qh.paragraph_format.space_before = Pt(0)
                qh.paragraph_format.space_after  = Pt(2)
                _run(qh, f"Question: {display_counter}", bold=True, size_pt=13)

                mp = doc.add_paragraph()
                mp.paragraph_format.space_before = Pt(0)
                mp.paragraph_format.space_after  = Pt(2)
                _run(mp, "Level of question",  bold=True, size_pt=11)
                _run(mp, f": {diff_}  |  ",                size_pt=11)
                _run(mp, "Number of Marks: ",  bold=True, size_pt=11)
                _run(mp, f"{marks}",                       size_pt=11)

                cp = doc.add_paragraph()
                cp.paragraph_format.space_before = Pt(0)
                cp.paragraph_format.space_after  = Pt(2)
                _run(cp, "Chapter", bold=True, size_pt=11)
                _run(cp, f" :  {topic_}",     size_pt=11)

                _hr(doc, color="BFBFBF")

                def _looks_failed(stem, options):
                    s = (stem or "").strip()
                    if not s:
                        return True
                    sl = s.lower()
                    if re.search(r"\[\s*q\s*\d+", sl): return True
                    if re.search(r"\[\s*stem\s*\]", sl): return True
                    if re.search(r"\btest\s+stem\b", sl): return True
                    if sl in ("stem", "[stem]", "n/a", "tbd", "placeholder", "todo", "todo:"):
                        return True
                    if re.match(r"^q\s*\d+\s*stem\s*$", sl): return True
                    non_empty = [v for v in options.values() if v and v.strip()]
                    if not non_empty: return True
                    if all(v.strip().upper() in ("A", "B", "C", "D") for v in non_empty): return True
                    if all(re.match(r"^option\s*[a-d]$", v.strip(), re.I) for v in non_empty): return True
                    return False

                if found and not vis and _looks_failed(text, opts) and q_num in locations:
                    vis = True

                if not found:
                    np_ = doc.add_paragraph()
                    _run(np_, "Question not found on specified page - needs review", italic=True, size_pt=11)
                elif vis:
                    if progress_cb:
                        progress_cb(img_done + 1, total_imgs or 1, f"Cropping Q{q_num} from QP…")
                    img_done += 1
                    q_loc = q.get("_loc")
                    img_bytes = crop_question_png(qp_bytes, q_loc) if q_loc else None
                    if img_bytes:
                        ip = doc.add_paragraph()
                        ip.paragraph_format.space_before = Pt(0)
                        ip.paragraph_format.space_after  = Pt(4)
                        ip.add_run().add_picture(io.BytesIO(img_bytes), width=Cm(CONTENT_WIDTH_CM - 1.0))
                    elif text and not _looks_failed(text, opts):
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
                        np_ = doc.add_paragraph()
                        _run(np_, f"Question Q{q_num} could not be extracted — refer to QP.", italic=True, size_pt=11)
                else:
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

                sl = doc.add_paragraph()
                sl.paragraph_format.space_before = Pt(8)
                sl.paragraph_format.space_after  = Pt(2)
                _run(sl, "Student's Solution:", bold=True, size_pt=11)
                _add_solution_box(doc, n_lines=max(6, (marks or 1) * 2))

                ap = doc.add_paragraph()
                ap.paragraph_format.space_before = Pt(8)
                ap.paragraph_format.space_after  = Pt(2)
                ap.paragraph_format.keep_with_next = True
                _run(ap, "Answer from Mark Scheme:", bold=True, size_pt=11)

                av = doc.add_paragraph()
                av.paragraph_format.space_before = Pt(0)
                av.paragraph_format.space_after  = Pt(2)
                _run(av, answer, bold=True, size_pt=12)

                if quote:
                    _hr(doc, color="BFBFBF")
                    qup = doc.add_paragraph()
                    qup.paragraph_format.space_before = Pt(2)
                    qup.paragraph_format.space_after  = Pt(2)
                    _run(qup, "Keep it up", bold=True, size_pt=11)
                    _run(qup, f" : {quote}",       size_pt=11)
                else:
                    _hr(doc, color="BFBFBF")

                is_first_q_in_diff = False

            is_first_diff_in_chapter = False
        is_first_chapter = False

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
#  TOPIC MATCH (for diagnostic only)
# ═══════════════════════════════════════════════════════════════════════════════
def _topic_keywords(text: str) -> set:
    if not text:
        return set()
    t = text.lower()
    boilerplate = re.compile(
        r"(maximum\s+mark|marks?\s*\]|international\s+baccalaureate|©\s*\d{4}|"
        r"turn\s+over|award\s+(a\d|m\d|r\d)|note\s*:|\b(m\d|a\d|r\d|ft|ag)\b|"
        r"calculator|gdc|\bsolve\b|\bfind\b|\bcalculate\b|\bdetermine\b|"
        r"\bthe\b|\bof\b|\ba\b|\bto\b|\bin\b|\bis\b|\bfor\b|\bbe\b|\bby\b|"
        r"\bwith\b|\bare\b|\band\b|\bor\b|\bif\b|\bnot\b)", re.IGNORECASE
    )
    t = boilerplate.sub(" ", t)
    tokens = re.findall(r"\b[a-z][a-z]{3,}\b", t)
    return set(tokens)


def _topics_match(qp_text: str, ms_text: str, threshold: float = 0.10) -> bool:
    qp_kw = _topic_keywords(qp_text)
    ms_kw = _topic_keywords(ms_text)
    if not qp_kw or not ms_kw:
        return True
    overlap = qp_kw & ms_kw
    smaller = min(len(qp_kw), len(ms_kw))
    if smaller == 0:
        return True
    return len(overlap) / smaller >= threshold


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN STREAMLIT FLOW
# ═══════════════════════════════════════════════════════════════════════════════
if st.button(
    "⚡ Extract & Generate Worksheet",
    type="primary",
    disabled=not (qp_file and ms_file and xl_file),
):
    client = anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])

    if not qp_file or not ms_file or not xl_file:
        st.warning("Please upload QP PDF, MS PDF, and Excel file first.")
        st.stop()

    qp_bytes = read_bytes(qp_file)
    ms_bytes = read_bytes(ms_file)

    if not qp_bytes or not ms_bytes:
        st.error("❌ Could not read uploaded files. Please re-upload and try again.")
        st.stop()

    if qp_bytes == ms_bytes:
        st.error("❌ QP and MS files are identical! Please upload the correct Mark Scheme.")
        st.stop()

    is_structured = (mode == "Structured Mark Scheme Mode")

    with st.status("Processing…", expanded=True) as status:

        # ── 1) Excel ──────────────────────────────────────────────────────────
        st.write("📊 Reading Excel (metadata only)…")
        try:
            parse_excel._filter_paper2 = st.session_state.get("filter_p2", False)
            xl_rows = parse_excel(xl_file)
        except Exception as e:
            st.error(f"Failed to read Excel: {e}")
            st.stop()
        if not xl_rows:
            _fp = st.session_state.get("filter_p2", False)
            if _fp:
                st.error(
                    "❌ No valid question rows found in Excel.\n\n"
                    "**Most likely reason:** All rows were removed because the "
                    "Reference column was detected as Paper 2.  \n"
                    "👉 **Uncheck the 'Skip Paper 2 rows' checkbox** above "
                    "and click Extract & Generate again."
                )
            else:
                st.error(
                    "❌ No valid question rows found in Excel.\n\n"
                    "Please check:\n"
                    "- Row 1 must be column headers\n"
                    "- There must be a Question Number column (Q No., Q#, etc.)\n"
                    "- Question numbers must be integers (1, 2, 3…)\n"
                    "- If this is a Paper 2 file, uncheck 'Skip Paper 2 rows'"
                )
            st.stop()
        st.write(f"✅ {len(xl_rows)} question rows in Excel")

        col_map = getattr(parse_excel, "last_column_map", {})
        if col_map:
            mapped_pairs = []
            unmapped = []
            for logical, (idx, actual) in col_map.items():
                if idx is None:
                    unmapped.append(logical)
                else:
                    mapped_pairs.append(f"**{logical}** → `{actual}`")
            with st.expander("🔍 Excel column mapping", expanded=bool(unmapped)):
                st.write(" · ".join(mapped_pairs))
                if unmapped:
                    st.error(f"❌ Missing column(s): {', '.join(unmapped)}")

        zero_page_count = sum(1 for r in xl_rows if not r.get("page_num"))
        range_count     = sum(1 for r in xl_rows
                              if r.get("page_num_end", 0) > r.get("page_num", 0))
        if zero_page_count > len(xl_rows) * 0.1:
            st.warning(f"⚠ {zero_page_count}/{len(xl_rows)} rows have Page Number = 0. "
                       "Check the Page Number column in your Excel.")
        if range_count:
            st.info(f"ℹ️ {range_count} question(s) span multiple pages "
                    f"(e.g. '1011' interpreted as pages 10–11) — "
                    "they will be cropped across all specified pages.")

        dups = getattr(parse_excel, "last_duplicates", [])
        if dups:
            st.warning(f"⚠ Removed {len(dups)} duplicate rows")

        # ── 2) QP locations ───────────────────────────────────────────────────
        st.write("📍 Locating questions in QP PDF…")
        try:
            locations = find_question_locations(qp_bytes)
        except Exception as e:
            st.error(f"Failed to scan QP PDF: {e}")
            st.stop()
        st.write(f"✅ Found {len(locations)} question marker(s) in QP")

        # Segment detection (for MS matching)
        excel_segments: list[list[int]] = []
        cur_seg: list[int] = []
        seg_max = 0
        seg_ref = None
        for i, r in enumerate(xl_rows):
            qn  = r["qn"]
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

        seg_offsets = infer_segment_page_offsets(locations, xl_rows, excel_segments)
        row_offset  = {}
        for seg_idx, seg_rows in enumerate(excel_segments):
            off = seg_offsets[seg_idx] if seg_idx < len(seg_offsets) else 0
            for ri in seg_rows:
                row_offset[ri] = off

        # ── 3) MS sections ─────────────────────────────────────────────────────
        if is_structured:
            st.write("📐 Locating worked solutions in MS PDF…")
            try:
                ms_sections = find_ms_question_locations(ms_bytes)
            except Exception as e:
                st.error(f"MS extraction failed: {e}")
                st.stop()
            st.write(f"✅ Found {len(ms_sections)} MS section(s)")
        else:
            st.write("🔑 MCQ mode — extracting answers from MS PDF…")
            try:
                ms_sections = extract_answers_pymupdf_per_section(ms_bytes)
            except Exception as e:
                st.error(f"MS extraction failed: {e}")
                st.stop()
            st.write(f"✅ Found {len(ms_sections)} answer section(s)")

        segments = excel_segments
        st.write(f"📚 Using {len(segments)} paper segment(s)")

        # ── 4) Row → MS section mapping ──────────────────────────────────────────
        #
        # Problem: MS PDFs for IB Math typically contain BOTH SL and HL sections
        # interleaved (SL section, then HL section, for each exam session).
        # Excel rows are HL-only (or a single paper type).
        # Pure positional mapping therefore pairs the wrong sections.
        #
        # Fix strategy:
        #   1. Determine the max question number per Excel segment (= paper type).
        #   2. Filter MS sections to only those whose max question number is
        #      "compatible" (within ±2) with the Excel segment max_qn.
        #   3. When counts differ (MS has SL+HL, Excel has HL only), pick only
        #      the compatible subset and map positionally within that subset.
        #   4. Fallback: if counts match exactly, use standard positional mapping.
        # ─────────────────────────────────────────────────────────────────────────

        row_to_section     = {}
        row_to_section_idx = {}

        def _seg_min_page(seg_rows_):
            pgs = [xl_rows[ri].get("page_num", 0) for ri in seg_rows_
                   if xl_rows[ri].get("page_num", 0) > 0]
            return min(pgs) if pgs else 999999

        def _seg_max_qn(seg_rows_):
            return max((xl_rows[ri].get("qn", 0) for ri in seg_rows_), default=0)

        sorted_seg_indices = sorted(range(len(segments)), key=lambda i: _seg_min_page(segments[i]))

        # Compute max_qn for each segment (= highest question number in that segment)
        seg_max_qns = [_seg_max_qn(segments[si]) for si in range(len(segments))]
        overall_excel_max_qn = max(seg_max_qns) if seg_max_qns else 13

        # ── Filter MS sections that are compatible with the Excel paper type ─────
        # Compatible = max question number within ±3 of the Excel segment max qn,
        # OR if the MS section covers more questions than the SL counterpart.
        # When the MS has twice as many sections as Excel segments (SL+HL combined),
        # pick only the sections whose max_qn >= (overall_excel_max_qn - 2).
        n_segs = len(segments)
        n_ms   = len(ms_sections)

        if n_ms >= 2 * n_segs:
            # MS has ~2× sections → contains both SL and HL (or TZ1+TZ2 with different counts)
            # Keep only the sections compatible with Excel's question count
            threshold = max(overall_excel_max_qn - 2, 10)
            compatible_ms = [
                (orig_idx, sec)
                for orig_idx, sec in enumerate(ms_sections)
                if isinstance(sec, dict) and sec.get("max_qn", len(sec.get("questions", {}))) >= threshold
            ]
            # Fallback: if that still doesn't give enough, use every other section
            if len(compatible_ms) < n_segs:
                # Take every second section starting from index 1 (the HL one in each pair)
                compatible_ms = [
                    (orig_idx, sec)
                    for orig_idx, sec in enumerate(ms_sections)
                    if orig_idx % 2 == 1
                ]
            # If still not enough, fall back to all sections
            if len(compatible_ms) < n_segs:
                compatible_ms = list(enumerate(ms_sections))
        else:
            # MS count roughly matches Excel segments → use all sections
            compatible_ms = list(enumerate(ms_sections))

        # ── Positional mapping within compatible MS sections ──────────────────────
        seg_paper_codes = [None] * len(segments)
        for pos, seg_idx in enumerate(sorted_seg_indices):
            seg_rows = segments[seg_idx]
            if pos < len(compatible_ms):
                orig_ms_idx, sec = compatible_ms[pos]
            else:
                orig_ms_idx, sec = -1, None
            ms_idx = orig_ms_idx
            for ri in seg_rows:
                row_to_section[ri]     = sec
                row_to_section_idx[ri] = ms_idx

        try:
            ms_code_to_section = {}
            for _ms_i, _ms_sec in enumerate(ms_sections):
                _code = (_ms_sec.get("paper_code") if isinstance(_ms_sec, dict) else None)
                if _code and _code not in ms_code_to_section:
                    ms_code_to_section[_code] = _ms_i
        except Exception:
            ms_code_to_section = {}

        paired_by_code = sum(1 for code in seg_paper_codes if code and code in ms_code_to_section)
        unpaired_segs  = []

        if is_structured:
            match_table = []
            for seg_idx, seg_rows in enumerate(segments):
                ms_sec_i = row_to_section_idx.get(seg_rows[0], -1)
                ms_sec   = row_to_section.get(seg_rows[0])
                ms_code  = (ms_sec.get("paper_code") if isinstance(ms_sec, dict) else None)
                ref_str  = xl_rows[seg_rows[0]].get("ref", "") if seg_rows else ""
                match_table.append({
                    "Seg#":        seg_idx + 1,
                    "Excel Ref":   ref_str[:40],
                    "MS section#": ms_sec_i + 1 if ms_sec_i >= 0 else "—",
                    "MS code":     ms_code or "—",
                    "Method":      "HL-filtered" if n_ms >= 2 * n_segs else "Positional",
                    "Questions":   len(seg_rows),
                    "MS Qs avail": len(ms_sec.get("questions", {})) if ms_sec else 0,
                })
            with st.expander(f"📋 Paper ↔ MS Section Matching ({paired_by_code} by code, {len(unpaired_segs)} positional)", expanded=True):
                st.dataframe(match_table, use_container_width=True, hide_index=True)

        if len(ms_sections) != len(segments):
            ratio = len(ms_sections) / max(len(segments), 1)
            if ratio < 0.5 and is_structured:
                st.warning(f"⚠ MS PDF covers fewer papers than Excel needs ({len(ms_sections)} vs {len(segments)} segments).")
            elif len(ms_sections) != len(segments):
                st.warning(f"⚠ Excel has {len(segments)} segment(s) but MS has {len(ms_sections)} section(s). Some answers may be missing.")

        # ── 5) QP extraction (for MCQ) + merge ───────────────────────────────
        if is_structured:
            # Structured: directly build worksheet from QP crops + MS crops
            # (no classify_and_extract needed)
            st.write("🔗 Merging rows (structured mode)…")

            # Apply marks override and filter
            clean_rows = []
            # Open QP once (avoids re-opening 226 times for marks lookup)
            try:
                _qp_doc_shared = fitz.open(stream=qp_bytes, filetype="pdf")
                _qp_n_pages_shared = len(_qp_doc_shared)
            except Exception:
                _qp_doc_shared = None
                _qp_n_pages_shared = 9999

            for i, r in enumerate(xl_rows):
                pg = r.get("page_num", 0)
                if pg <= 0:
                    continue
                r2 = dict(r)
                # Apply marks from QP if marks=0
                if r2.get("marks", 0) == 0 and _qp_doc_shared is not None:
                    try:
                        _ptxt = (_qp_doc_shared[pg - 1].get_text()
                                 if pg <= _qp_n_pages_shared else "")
                        _mm = re.search(r'\[Maximum mark[:\s]+(\d+)\]', _ptxt, re.IGNORECASE)
                        if _mm:
                            r2["marks"] = int(_mm.group(1))
                    except Exception:
                        pass
                if r2.get("marks", 0) == 0:
                    r2["marks"] = 1
                # Validate / clamp page_num_end against actual PDF length
                pg_end_raw = r2.get("page_num_end", 0) or 0
                if pg_end_raw > 0:
                    _qp_n_pages = _qp_n_pages_shared
                    if pg_end_raw > _qp_n_pages:
                        r2["page_num_end"] = _qp_n_pages
                    elif pg_end_raw < r2.get("page_num", 1):
                        r2["page_num_end"] = r2.get("page_num", 1)
                clean_rows.append((i, r2))

            # Build questions list for validation display
            questions = []
            for i, r2 in clean_rows:
                sec     = row_to_section.get(i)
                sec_idx = row_to_section_idx.get(i, -1)
                qn      = r2["qn"]
                q_info  = (sec.get("questions", {}).get(qn) if sec and "questions" in sec else None)
                ms_img  = None   # will crop during word build
                ans_ok  = q_info is not None
                questions.append({
                    **r2,
                    "_loc":         None,
                    "_ms_section":  f"MS section {sec_idx+1}" if sec_idx >= 0 else "no section",
                    "_ms_image":    ms_img,
                    "_ms_page":     (q_info["page_idx"] + 1 if q_info else None),
                    "_ms_first":    "",
                    "_topic_match": True,
                    "_qp_page":     r2.get("page_num"),
                    "_mode":        mode,
                    "found":        r2.get("page_num", 0) > 0,
                    "needs_image":  True,
                    "text":         "",
                    "A": "", "B": "", "C": "", "D": "",
                    "answer":       "" if ans_ok else "Answer not found - needs review",
                    "answerFound":  ans_ok,
                })

            if _qp_doc_shared is not None:
                try:
                    _qp_doc_shared.close()
                except Exception:
                    pass
            xl_rows_clean = [r2 for (i, r2) in clean_rows]

        else:
            # MCQ: classify_and_extract
            st.write("🔍 Extracting question text…")
            row_tuples = []
            for i, r in enumerate(xl_rows):
                p   = r.get("page_num", 0)
                off = row_offset.get(i, 0)
                if p and off:
                    p = max(1, p + off)
                row_tuples.append((r["qn"], p))
            try:
                qp_data = classify_and_extract(None, "", row_tuples, qp_bytes=qp_bytes, locations=locations)
            except Exception as e:
                st.error(f"QP extraction failed: {e}")
                st.stop()

            questions = []
            for i, (r, qp_q) in enumerate(zip(xl_rows, qp_data)):
                n       = r["qn"]
                sec     = row_to_section.get(i)
                sec_idx = row_to_section_idx.get(i, -1)
                found_q = bool(qp_q) and qp_q.get("found") is not False
                needs_img = found_q and qp_q.get("needs_image", False)

                r2 = dict(r)
                if r2.get("marks", 0) == 0:
                    r2["marks"] = 1

                ans = (sec["answers"].get(str(n), "") if sec and "answers" in sec else "")
                ans_ok   = ans in ("A", "B", "C", "D")
                ans_text = ans if ans_ok else "Answer not found - needs review"

                sec_label = (f"MS section {sec_idx+1}" if sec_idx >= 0 else "no section")

                questions.append({
                    **r2,
                    "_loc":         qp_q.get("_loc"),
                    "_ms_section":  sec_label,
                    "_ms_image":    None,
                    "_ms_page":     sec.get("start_page") if sec else None,
                    "_ms_first":    "",
                    "_topic_match": True,
                    "_qp_page":     (qp_q["_loc"]["page_idx"] + 1 if qp_q.get("_loc") else None),
                    "_mode":        mode,
                    "found":        found_q,
                    "needs_image":  needs_img,
                    "text":         qp_q.get("text", "") if found_q else "",
                    "A": qp_q.get("A", "") if found_q else "",
                    "B": qp_q.get("B", "") if found_q else "",
                    "C": qp_q.get("C", "") if found_q else "",
                    "D": qp_q.get("D", "") if found_q else "",
                    "answer":      ans_text,
                    "answerFound": ans_ok,
                })

        status.update(label="✅ Extraction done!", state="complete")

    # ── Validation summary ────────────────────────────────────────────────────
    st.subheader("Preview")
    total      = len(questions)
    found_qp   = sum(1 for q in questions if q["found"])
    found_ms   = sum(1 for q in questions if q["answerFound"])
    visual_cnt = sum(1 for q in questions if q["needs_image"])

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total",         total)
    c2.metric("Found in QP",   found_qp)
    c3.metric("Answers found", found_ms)
    c4.metric("With visuals",  visual_cnt)

    st.dataframe(
        [{
            "Excel Row":    i + 2,
            "Q#":           q["qn"],
            "Reference":    q.get("ref", ""),
            "Excel Page":   q.get("page_num", ""),
            "QP Page Used": q.get("_qp_page") or "—",
            "MS Page Used": q.get("_ms_page") or "—",
            "Mode":         "MCQ" if "MCQ" in (q.get("_mode") or "") else "Structured",
            "MS Section":   q.get("_ms_section", ""),
            "Chapter":      q.get("topic", ""),
            "Difficulty":   q.get("difficulty", ""),
            "Marks":        q.get("marks", ""),
            "Status":       ("✅" if q["found"] and q["answerFound"]
                             else ("⚠ MS missing" if q["found"]
                                   else ("⚠ QP missing" if q["answerFound"]
                                         else "❌ both missing"))),
        } for i, q in enumerate(questions)],
        use_container_width=True,
        hide_index=True,
    )

    unmatched = []
    for i, q in enumerate(questions):
        problems = []
        if not q["found"]:
            if not q.get("page_num"):
                problems.append("no Page Number in Excel")
            else:
                problems.append(f"Q{q['qn']} not located on page {q['page_num']} of QP")
        if not q["answerFound"] and not q.get("_ms_image"):
            if "Structured" in (q.get("_mode") or ""):
                problems.append(f"Q{q['qn']} not found in MS section")
            elif "MCQ" in (q.get("_mode") or ""):
                problems.append(f"Q{q['qn']} has no A/B/C/D answer")
        if problems:
            unmatched.append({
                "Excel Row": i + 2, "Reference": q.get("ref", ""),
                "Q#": q["qn"], "Page (Excel)": q.get("page_num", 0),
                "Reason": " · ".join(problems),
            })

    qp_match_pct    = found_qp / max(total, 1)
    matching_is_bad = qp_match_pct < 0.70

    if unmatched:
        st.subheader(f"⚠ Unmatched rows ({len(unmatched)})")
        st.dataframe(unmatched, use_container_width=True, hide_index=True)

    st.subheader("Download")
    if matching_is_bad:
        st.error(f"❌ Only {found_qp}/{total} ({qp_match_pct:.0%}) rows found in QP. Fix matching before downloading.")
        st.stop()

    if qp_match_pct < 0.95:
        st.warning(f"⚠ {found_qp}/{total} rows matched ({qp_match_pct:.0%}). {total - found_qp} unmatched will show 'Question not found'.")

    # ── Build Word ─────────────────────────────────────────────────────────────
    _prog_bar  = st.progress(0)
    _prog_text = st.empty()

    def on_progress(done, total_, msg):
        _prog_bar.progress(done / max(total_, 1))
        _prog_text.text(msg)

    with st.spinner("Building Word document…"):
        try:
            if is_structured:
                # Build using FINAL AGENT structured generator
                _docx = build_structured_worksheet(
                    xl_rows_clean,
                    qp_bytes,
                    ms_bytes,
                    ms_sections,
                    {i: row_to_section[orig_i]
                     for i, (orig_i, _) in enumerate(clean_rows)},
                    progress_cb=on_progress,
                )
            else:
                _docx = build_word_document(questions, qp_bytes, locations, progress_cb=on_progress)
        except Exception as e:
            st.error(f"Word generation failed: {e}")
            import traceback
            st.code(traceback.format_exc())
            st.stop()

    _prog_bar.progress(1.0)
    _prog_text.empty()

    ms_missing_n = sum(1 for q in questions if not q.get("_ms_image") and not q.get("answerFound"))

    st.session_state["ws_docx"]    = _docx
    st.session_state["ws_missing"] = ms_missing_n
    st.session_state["ws_stats"]   = (
        f"Total: {total} | Found in QP: {found_qp} | Answers found: {found_ms}"
    )
    st.session_state["ws_ready"]   = True
    st.success("✅ Worksheet built — scroll down to download.")


# ═══════════════════════════════════════════════════════════════════════════════
#  PERSISTENT DOWNLOAD
# ═══════════════════════════════════════════════════════════════════════════════
if st.session_state.get("ws_ready") and st.session_state.get("ws_docx"):
    st.divider()
    st.subheader("⬇️ Download Worksheet")

    _stats = st.session_state.get("ws_stats", "")
    if _stats:
        st.info(f"📊 {_stats}")

    _missing = st.session_state.get("ws_missing", 0)
    if _missing:
        st.warning(f"⚠️ {_missing} question(s) have no MS answer.")

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
