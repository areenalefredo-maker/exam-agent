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

import fitz                           # PyMuPDF — PDF inspection + rendering
from PIL import Image

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


# ═══════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
def read_bytes(uploaded_file) -> bytes:
    uploaded_file.seek(0)
    return uploaded_file.read()


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
    Sanitises: 'Error'/empty topic → 'Unclassified', marks 0/missing → 1.
    """
    uploaded_file.seek(0)
    wb = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)
    ws = wb.active
    headers = [str(c.value).strip() if c.value else "" for c in next(ws.iter_rows(max_row=1))]

    def col_idx(*names):
        for n in names:
            if n in headers:
                return headers.index(n)
        return None

    qn_idx   = col_idx("Question No.", "Question No", "Q No", "Q#", "Question Number", "Q") or 2
    page_idx = col_idx("Page Number", "Page", "page_number")
    top_idx  = col_idx("Topic", "Chapter", "topic", "chapter")
    dif_idx  = col_idx("Difficulty", "difficulty")
    mrk_idx  = col_idx("Marks", "marks")
    qut_idx  = col_idx("Quote", "quote")
    ref_idx  = col_idx("Reference", "ref")

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
    duplicates = []   # (qn, reference) pairs that were skipped

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
            marks = int(float(cell_val(row, mrk_idx, "1")))
        except ValueError:
            marks = 1
        if marks <= 0:
            marks = 1

        try:
            page_num = int(float(cell_val(row, page_idx, "0")))
        except ValueError:
            page_num = 0

        difficulty = cell_val(row, dif_idx, "Unspecified") or "Unspecified"
        ref        = cell_val(row, ref_idx, "")
        quote      = cell_val(row, qut_idx, "")

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

    result = {}
    for i, (qn, page_idx, top_y, ph, pw, mrx) in enumerate(raw):
        # bottom_y = next marker on SAME page, else page bottom
        bottom_y = ph * 0.94
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
    tolerance: int = 3,        # tolerate ±N pages of mismatch
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
    right_px = iw

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


def extract_answers_pymupdf_per_section(ms_bytes: bytes) -> list[dict]:
    """Extract answers from MS PDF, grouped by section.

    A 'section' = a contiguous range of MS pages that hold answers for ONE
    paper. We detect section boundaries by looking for the answer-grid restart
    (i.e. Q1 appearing again after a higher Q# was already seen).

    Returns a list of section dicts:
        [{"start_page": int_1based, "answers": {qn_str: 'A'|'B'|'C'|'D'}}, …]

    Each section's `answers` covers ONE paper's Q1..Q40 (or whatever range).
    The order of the list matches the document order — first paper's answers
    come first.
    """
    sections: list[dict] = []
    current: dict = {"start_page": 1, "answers": {}, "max_qn": 0}

    try:
        doc = fitz.open(stream=ms_bytes, filetype="pdf")
    except Exception:
        return []

    patterns = [
        re.compile(r"\b(\d{1,2})\.\s+([ABCD])(?![A-Za-z0-9])"),
        re.compile(r"(?:^|\s)(\d{1,2})\s{2,}([ABCD])(?![A-Za-z0-9])"),
        re.compile(r"\bQ\s*(\d{1,2})\s+([ABCD])(?![A-Za-z0-9])"),
        re.compile(r"\b(\d{1,2})\s*[:\-]\s*([ABCD])(?![A-Za-z0-9])"),
    ]

    for page_idx, page in enumerate(doc):
        text = page.get_text()
        page_pairs = []  # (qn, ans) found on this page, in reading order
        for line in text.split("\n"):
            line = line.strip()
            if not line:
                continue
            for pat in patterns:
                for m in pat.finditer(line):
                    qn  = int(m.group(1))
                    ans = m.group(2)
                    if 1 <= qn <= 99:
                        page_pairs.append((qn, ans))
                        break   # one match per line is enough

        for qn, ans in page_pairs:
            qn_str = str(qn)
            # Detect new paper: a lower Q# appearing after we already saw
            # a higher Q# in this section (e.g. Q1 after we recorded Q35)
            if (qn_str in current["answers"] or
                    (current["max_qn"] >= 10 and qn < current["max_qn"] - 5)):
                # Save current section, start a new one
                if current["answers"]:
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
    """Deterministic answer extraction from MS PDF using PyMuPDF + regex.
       IB mark schemes typically format answers in a grid like:
         1.  D       16.  C       31.  C
         2.  A       17.  D       32.  B
       This function scans every page line by line and captures
       (question_number, answer_letter) pairs.
       Returns: {qn_str: 'A'|'B'|'C'|'D'}
    """
    answers: dict = {}
    try:
        doc = fitz.open(stream=ms_bytes, filetype="pdf")
    except Exception:
        return {}

    # Patterns ordered from strict → permissive.  We try each on each line.
    patterns = [
        # "1.  D" — most common (with dot)
        re.compile(r"\b(\d{1,2})\.\s+([ABCD])(?![A-Za-z0-9])"),
        # "1   D" — space-separated grid (no dot)
        re.compile(r"(?:^|\s)(\d{1,2})\s{2,}([ABCD])(?![A-Za-z0-9])"),
        # "Q1   D"
        re.compile(r"\bQ\s*(\d{1,2})\s+([ABCD])(?![A-Za-z0-9])"),
        # "1: D"
        re.compile(r"\b(\d{1,2})\s*[:\-]\s*([ABCD])(?![A-Za-z0-9])"),
    ]

    for page in doc:
        text = page.get_text()
        for line in text.split("\n"):
            line = line.strip()
            if not line:
                continue
            for pat in patterns:
                for m in pat.finditer(line):
                    qn  = m.group(1)
                    ans = m.group(2)
                    if qn not in answers:
                        answers[qn] = ans

    doc.close()
    return answers


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


def _add_solution_box(doc, n_lines: int = 4):
    """Reference format: 4-row table, each row ~0.9 cm tall,
       only a thin bottom border (#BFBFBF) — produces 4 writing lines.
    """
    table = doc.add_table(rows=n_lines, cols=1)
    table.autofit = False
    for cell in table.columns[0].cells:
        cell.width = Cm(CONTENT_WIDTH_CM)

    for row in table.rows:
        # Row height = 510 twips (matches reference)
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
    grouped: dict[str, dict[str, list]] = {}
    for q in questions:
        t = q.get("topic") or "Unclassified"
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
                # **Level of question**: Easy  |  **Number of Marks: **1  |
                # **Reference:** (xxx)  |
                mp = doc.add_paragraph()
                mp.paragraph_format.space_before = Pt(0)
                mp.paragraph_format.space_after  = Pt(2)
                _run(mp, "Level of question",  bold=True, size_pt=11)
                _run(mp, f": {diff_}  |  ",                size_pt=11)
                _run(mp, "Number of Marks: ",  bold=True, size_pt=11)
                _run(mp, f"{marks}  |  ",                  size_pt=11)
                _run(mp, "Reference:",         bold=True, size_pt=11)
                _run(mp, f" {ref}  |",                     size_pt=11)

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
                sl = doc.add_paragraph()
                sl.paragraph_format.space_before = Pt(8)
                sl.paragraph_format.space_after  = Pt(2)
                _run(sl, "Student's Solution:", bold=True, size_pt=11)

                # 4 writing lines (matches reference exactly)
                _add_solution_box(doc, n_lines=4)

                # ── Answer from Mark Scheme ────────────────────────────────────
                ap = doc.add_paragraph()
                ap.paragraph_format.space_before = Pt(8)
                ap.paragraph_format.space_after  = Pt(2)
                _run(ap, "Answer from Mark Scheme:", bold=True, size_pt=11)

                av = doc.add_paragraph()
                av.paragraph_format.space_before = Pt(0)
                av.paragraph_format.space_after  = Pt(2)
                _run(av, answer, bold=True, size_pt=12)

                # ── Separator before Keep it up ────────────────────────────────
                _hr(doc, color="BFBFBF")

                # ── Keep it up: ────────────────────────────────────────────────
                if quote:
                    qup = doc.add_paragraph()
                    qup.paragraph_format.space_before = Pt(2)
                    qup.paragraph_format.space_after  = Pt(2)
                    _run(qup, "Keep it up", bold=True, size_pt=11)
                    _run(qup, f" : {quote}",       size_pt=11)

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

    qp_bytes = read_bytes(qp_file)
    ms_bytes = read_bytes(ms_file)

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

        # Show only TRUE duplicate removals (same Q# + same Reference +
        # same Page + same Topic + same Difficulty)
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

        # ── 3) Extract text + classify per-Excel-row using (qn, page) ───────
        # CRITICAL: We pass tuples of (q_num, excel_page) so the extractor can
        # disambiguate when the same Q# appears in multiple papers within one PDF.
        st.write("🔍 Extracting question text per row…")
        row_tuples = [(r["qn"], r.get("page_num", 0)) for r in xl_rows]
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

        # ── 4) Mark Scheme: extract per-section answers ──────────────────────
        # The MS PDF may contain multiple papers concatenated. Each section
        # holds answers for ONE paper (Q1..Q40). We detect Excel "segments"
        # (one segment = a contiguous run of rows belonging to the same paper)
        # and match Excel segments 1-to-1 with MS sections in document order.
        st.write("🔑 Extracting answers from Mark Scheme PDF…")
        try:
            ms_sections = extract_answers_pymupdf_per_section(ms_bytes)
        except Exception as e:
            st.error(f"MS extraction failed: {e}")
            st.stop()
        st.write(f"✅ Found {len(ms_sections)} answer section(s) in MS PDF")

        # Detect Excel paper segments: a new segment starts whenever Q#
        # drops back to 1 (or to a value much lower than the running max).
        # Each segment is a list of row INDICES in xl_rows.
        segments: list[list[int]] = []
        current_seg: list[int] = []
        max_qn = 0
        for i, r in enumerate(xl_rows):
            qn = r["qn"]
            # New segment if Q# resets to 1 OR drops significantly
            if current_seg and (qn == 1 or (max_qn >= 10 and qn < max_qn - 5)):
                segments.append(current_seg)
                current_seg = []
                max_qn = 0
            current_seg.append(i)
            if qn > max_qn:
                max_qn = qn
        if current_seg:
            segments.append(current_seg)

        st.write(f"📚 Detected {len(segments)} paper segment(s) in Excel "
                 f"(by Q# restart pattern)")

        # Build row_idx → ms_section mapping by segment order
        row_to_section: dict[int, dict | None] = {}
        for seg_idx, seg_rows in enumerate(segments):
            sec = ms_sections[seg_idx] if seg_idx < len(ms_sections) else None
            for ri in seg_rows:
                row_to_section[ri] = sec

        if len(ms_sections) != len(segments):
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
            ans = (sec["answers"].get(str(n), "") if sec else "")
            ans_ok = ans in ("A", "B", "C", "D")
            found_q   = bool(qp_q) and qp_q.get("found") is not False
            needs_img = found_q and qp_q.get("needs_image", False)

            questions.append({
                **r,
                "_loc":        qp_q.get("_loc"),       # per-row location
                "found":       found_q,
                "needs_image": needs_img,
                "text":        qp_q.get("text", "")  if found_q else "",
                "A":           qp_q.get("A", "")     if found_q else "",
                "B":           qp_q.get("B", "")     if found_q else "",
                "C":           qp_q.get("C", "")     if found_q else "",
                "D":           qp_q.get("D", "")     if found_q else "",
                "answer":      ans if ans_ok else "Answer not found - needs review",
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
            "Excel Row":  i + 2,   # +2 because header row + 1-based
            "Q#":         q["qn"],
            "Reference":  q.get("ref", ""),
            "Page":       q.get("page_num", ""),
            "Chapter":    q.get("topic", ""),
            "Difficulty": q.get("difficulty", ""),
            "Marks":      q.get("marks", ""),
            "Type":       "🖼 Image" if q["needs_image"] else "📝 Text",
            "First line of extracted":
                (q.get("text", "")[:60] + "…") if q.get("text") and len(q.get("text", "")) > 60
                else (q.get("text", "") or ("[image — see crop]" if q["needs_image"]
                      else "[not found]")),
            "Answer":     q["answer"] if q["answerFound"] else "—",
            "Status":     ("✅" if q["found"] and q["answerFound"]
                           else ("⚠ MS missing" if q["found"]
                                 else ("⚠ QP missing" if q["answerFound"]
                                       else "❌ both missing"))),
        } for i, q in enumerate(questions)],
        use_container_width=True,
        hide_index=True,
    )

    # ── Build & Download ──────────────────────────────────────────────────────
    st.subheader("Download")
    if visual_cnt:
        st.info(
            f"ℹ️ {visual_cnt} questions contain visual content — actual images "
            "will be cropped from the QP PDF and embedded in the Word file."
        )

    prog_bar  = st.progress(0)
    prog_text = st.empty()

    def on_progress(done, total_, msg):
        prog_bar.progress(done / max(total_, 1))
        prog_text.text(msg)

    with st.spinner("Building Word document…"):
        try:
            docx_bytes = build_word_document(
                questions, qp_bytes, locations,
                progress_cb=on_progress if visual_cnt else None,
            )
        except Exception as e:
            st.error(f"Word generation failed: {e}")
            st.stop()

    prog_bar.progress(1.0)
    prog_text.empty()

    st.download_button(
        label="⬇️ Download Worksheet (.docx)",
        data=docx_bytes,
        file_name="IB_Chemistry_Worksheet.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
