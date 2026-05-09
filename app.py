import streamlit as st
import anthropic
import openpyxl
import json
import re
import io
import base64
import fitz                          # PyMuPDF — PDF → image
from PIL import Image
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn as docx_qn   # renamed → never shadowed
from docx.oxml import OxmlElement

# ─── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Exam Worksheet Generator",
    page_icon="📄",
    layout="centered",
)

st.title("📄 Exam Worksheet Generator")
st.caption("Questions from QP · Answers from MS · Metadata from Excel")

with st.expander("ℹ️ Source rules", expanded=False):
    st.markdown("""
| Source | Used for |
|--------|----------|
| **QP PDF** | Question text, options, tables, diagrams, graphs — verbatim |
| **MS PDF** | Answers only |
| **Excel** | Question numbers, chapter, difficulty, marks, quotes |

- Questions with visuals → actual image cropped from QP page, embedded in Word
- Topic = "Error" or Marks = 0 in Excel → replaced with "Unclassified" / 1
""")

# ─── File uploads ─────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    qp_file = st.file_uploader("📋 Question Paper (QP)", type=["pdf"], key="qp")
    ms_file = st.file_uploader("✅ Mark Scheme (MS)",     type=["pdf"], key="ms")
with col2:
    xl_file = st.file_uploader("📊 Excel Sheet", type=["xlsx", "xls"], key="xl")

st.divider()


# ══════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def read_bytes(uploaded_file) -> bytes:
    uploaded_file.seek(0)
    return uploaded_file.read()


def to_b64(data: bytes) -> str:
    return base64.standard_b64encode(data).decode()


def safe_json(text: str):
    text = re.sub(r"```json\s*", "", text)
    text = re.sub(r"```\s*",     "", text)
    try:
        return json.loads(text.strip())
    except json.JSONDecodeError:
        return None


# ─── Excel parser ─────────────────────────────────────────────────────────────
def parse_excel(uploaded_file) -> list[dict]:
    """Return metadata rows from Excel.
    Sanitises 'Error' topics and zero marks.
    Never used for question text or answer.
    """
    uploaded_file.seek(0)
    wb = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)
    ws = wb.active
    headers = [str(c.value).strip() if c.value else "" for c in next(ws.iter_rows(max_row=1))]

    def col_idx(names):
        for name in names:
            if name in headers:
                return headers.index(name)
        return None

    qn_idx   = col_idx(["Question No.", "Question No", "Q No", "Q#", "Question Number", "Q"]) or 2
    page_idx = col_idx(["Page Number", "Page", "page_number"])
    top_idx  = col_idx(["Topic", "Chapter", "topic", "chapter"])
    dif_idx  = col_idx(["Difficulty", "difficulty"])
    mrk_idx  = col_idx(["Marks", "marks"])
    qut_idx  = col_idx(["Quote", "quote"])
    ref_idx  = col_idx(["Reference", "ref"])

    def cell_val(row, idx, fallback=""):
        if idx is None:
            return fallback
        try:
            v = list(row)[idx].value
            return str(v).strip() if v is not None else fallback
        except IndexError:
            return fallback

    rows = []
    for row in ws.iter_rows(min_row=2):
        raw_qn = list(row)[qn_idx].value
        try:
            q_num = int(float(str(raw_qn)))
        except (TypeError, ValueError):
            continue
        if q_num <= 0:
            continue

        topic = cell_val(row, top_idx)
        if not topic or topic.lower() in ("error", "none", "n/a", ""):
            topic = "Unclassified"

        try:
            marks = int(float(cell_val(row, mrk_idx, "1")))
            if marks <= 0:
                marks = 1
        except ValueError:
            marks = 1

        try:
            page_num = int(float(cell_val(row, page_idx, "0")))
        except ValueError:
            page_num = 0

        rows.append({
            "qn":         q_num,
            "page_num":   page_num,
            "topic":      topic,
            "difficulty": cell_val(row, dif_idx, ""),
            "marks":      marks,
            "quote":      cell_val(row, qut_idx, ""),
            "ref":        cell_val(row, ref_idx, ""),
        })
    return rows


# ─── Claude: extract questions from QP (verbatim) ────────────────────────────
def extract_questions_from_qp(
    client: anthropic.Anthropic,
    qp_b64: str,
    q_nums: list[int],
) -> list[dict]:
    nums_str = ", ".join(str(n) for n in q_nums)
    resp = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=8000,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {"type": "base64", "media_type": "application/pdf", "data": qp_b64},
                },
                {
                    "type": "text",
                    "text": f"""Extract ONLY question numbers {nums_str} from this IB exam paper.

STRICT RULES — never violate:
1. Copy text VERBATIM — do NOT paraphrase, shorten, or reword anything
2. Include the COMPLETE question stem and all 4 options (A, B, C, D) exactly as printed
3. For tables: represent as plain-text rows e.g.
   "| Header1 | Header2 |\\n| Val1 | Val2 |"
4. For equations/formulas: copy as-is (fractions, arrows, sub/superscripts)
5. If the question contains a DIAGRAM, GRAPH, IMAGE, or ORBITAL SHAPE that cannot
   be fully represented as text, set "hasVisual": true.
   Do NOT describe the visual — just flag it.
6. If a question number is NOT FOUND: set "found": false,
   "text": "Question not found in uploaded Question Paper"

Return ONLY this JSON array — no markdown, no preamble:
[{{
  "qn": <int>,
  "text": "<verbatim stem>",
  "A": "<option A verbatim>",
  "B": "<option B verbatim>",
  "C": "<option C verbatim>",
  "D": "<option D verbatim>",
  "hasVisual": <bool>,
  "found": <bool>
}}]""",
                },
            ],
        }],
    )
    raw = "".join(b.text for b in resp.content if hasattr(b, "text"))
    data = safe_json(raw)
    return data if isinstance(data, list) else []


# ─── Claude: extract answers from MS ─────────────────────────────────────────
def extract_answers_from_ms(
    client: anthropic.Anthropic,
    ms_b64: str,
    q_nums: list[int],
) -> dict:
    nums_str = ", ".join(str(n) for n in q_nums)
    resp = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=2000,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {"type": "base64", "media_type": "application/pdf", "data": ms_b64},
                },
                {
                    "type": "text",
                    "text": f"""Extract the correct answers for question numbers {nums_str}
from this IB mark scheme.

The mark scheme shows a grid like:
1. D    16. C    31. C
2. A    17. D    32. B

Return ONLY this JSON — no markdown, no explanation:
{{"1":"D","2":"A","3":"B",...}}

Use "NOT_FOUND" for any question not in the mark scheme.""",
                },
            ],
        }],
    )
    raw = "".join(b.text for b in resp.content if hasattr(b, "text"))
    data = safe_json(raw)
    return data if isinstance(data, dict) else {}


# ─── PDF page → PNG bytes (cached per session) ───────────────────────────────
_page_cache: dict = {}

def render_page_png(pdf_bytes: bytes, page_num: int, dpi: int = 180) -> bytes:
    """Render a 1-indexed PDF page to PNG bytes (cached within a run)."""
    key = (id(pdf_bytes), page_num)
    if key not in _page_cache:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        idx = page_num - 1
        if idx < 0 or idx >= len(doc):
            doc.close()
            return b""
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        pix = doc[idx].get_pixmap(matrix=mat, alpha=False)
        _page_cache[key] = pix.tobytes("png")
        doc.close()
    return _page_cache[key]


def crop_question_png(
    client: anthropic.Anthropic,
    pdf_bytes: bytes,
    page_num: int,
    q_num: int,
) -> bytes:
    """Return PNG bytes of the question area cropped from the QP page.
    Uses Claude to locate the bounding box; falls back to full page.
    Never returns a text description — always an image.
    """
    page_png = render_page_png(pdf_bytes, page_num)
    if not page_png:
        return b""

    img = Image.open(io.BytesIO(page_png))
    w, h = img.size

    try:
        resp = client.messages.create(
            model="claude-opus-4-5",
            max_tokens=300,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": to_b64(page_png),
                        },
                    },
                    {
                        "type": "text",
                        "text": (
                            f"This is page {page_num} of an IB exam paper "
                            f"({w}x{h} px).\n"
                            f"Find question number {q_num} and return its "
                            "bounding box in pixels, including ALL its content "
                            "(stem, options, table, diagram, graph).\n"
                            "Return ONLY JSON — no markdown:\n"
                            '{"top":<int>,"left":<int>,"bottom":<int>,"right":<int>}'
                        ),
                    },
                ],
            }],
        )
        raw    = "".join(b.text for b in resp.content if hasattr(b, "text"))
        coords = safe_json(raw)

        if isinstance(coords, dict) and all(
            k in coords for k in ("top", "left", "bottom", "right")
        ):
            pad    = 20
            left   = max(0, int(coords["left"])   - pad)
            top    = max(0, int(coords["top"])    - pad)
            right  = min(w, int(coords["right"])  + pad)
            bottom = min(h, int(coords["bottom"]) + pad)
            if right > left and bottom > top:
                cropped = img.crop((left, top, right, bottom))
                buf = io.BytesIO()
                cropped.save(buf, format="PNG")
                return buf.getvalue()
    except Exception:
        pass

    # Fallback: return the full rendered page
    return page_png


# ─── LaTeX → Unicode ──────────────────────────────────────────────────────────
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
        r"\leftrightarrow":"↔", r"\to":"→", r"\Delta":"Δ", r"\delta":"δ",
        r"\ominus":"⊖", r"\oplus":"⊕", r"\theta":"θ", r"\alpha":"α",
        r"\beta":"β",  r"\gamma":"γ",  r"\lambda":"λ", r"\mu":"μ",
        r"\pi":"π",    r"\sigma":"σ",  r"\cdot":"·",   r"\pm":"±",
        r"\geq":"≥",   r"\leq":"≤",    r"\neq":"≠",    r"\approx":"≈",
        r"\infty":"∞", r"\circ":"°",
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
    t = re.sub(r"\\\[([^\]]+)\\\]",lambda m: _math(m.group(1)), t)
    t = re.sub(r"\*\*([^*]+)\*\*", r"\1", t)
    t = re.sub(r"\*([^*]+)\*",     r"\1", t)
    return t


# ══════════════════════════════════════════════════════════════════════════════
#  WORD DOCUMENT BUILDER
# ══════════════════════════════════════════════════════════════════════════════

DIFF_ORDER = ["Easy", "Medium", "Hard"]


def _set_cell_borders(cell):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(docx_qn("w:val"),   "single")
        el.set(docx_qn("w:sz"),    "4")
        el.set(docx_qn("w:color"), "999999")
        tcPr.append(el)


def _run(para, text, *, bold=False, italic=False, color=None, size_pt=11):
    r = para.add_run(str(text))
    r.bold        = bold
    r.italic      = italic
    r.font.size   = Pt(size_pt)
    r.font.name   = "Arial"
    if color:
        r.font.color.rgb = RGBColor(*bytes.fromhex(color))
    return r


def _hr(doc):
    p   = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot  = OxmlElement("w:bottom")
    bot.set(docx_qn("w:val"),   "single")
    bot.set(docx_qn("w:sz"),    "4")
    bot.set(docx_qn("w:color"), "CCCCCC")
    pBdr.append(bot)
    pPr.append(pBdr)


def build_word_document(
    questions:   list[dict],
    qp_bytes:    bytes,
    client:      anthropic.Anthropic,
    progress_cb=None,
) -> bytes:
    """
    Build the Word worksheet.
    Visual questions → crop actual image from QP PDF page, embed in Word.
    No 'Visual content: …' text is ever written.
    """
    _page_cache.clear()

    doc = Document()
    for section in doc.sections:
        section.top_margin    = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # Title
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(tp, "IB Chemistry – Higher Level", bold=True, color="1F4E79", size_pt=18)
    sp = doc.add_paragraph()
    sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(sp, "Exam Worksheet", color="555555", size_pt=12)
    doc.add_paragraph()

    # Group by topic → difficulty
    grouped: dict[str, dict[str, list]] = {}
    for q in questions:
        t = q.get("topic") or "Unclassified"
        d = q.get("difficulty") or "Unspecified"
        grouped.setdefault(t, {}).setdefault(d, []).append(q)

    visual_qs    = [q for q in questions if q.get("hasVisual")]
    total_visual = len(visual_qs)
    visual_done  = 0

    for topic, diffs in grouped.items():
        h2 = doc.add_heading(topic, level=2)
        if h2.runs:
            h2.runs[0].font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

        sorted_diffs = sorted(
            diffs.keys(),
            key=lambda d: DIFF_ORDER.index(d) if d in DIFF_ORDER else 99,
        )

        for diff in sorted_diffs:
            dp = doc.add_paragraph()
            _run(dp, f"— {diff} questions —", italic=True, color="666666", size_pt=11)

            for q in diffs[diff]:
                q_num    = q["qn"]
                page_num = q.get("page_num", 0)
                stem     = latex_to_text(q.get("text", ""))
                opt_a    = latex_to_text(q.get("A", ""))
                opt_b    = latex_to_text(q.get("B", ""))
                opt_c    = latex_to_text(q.get("C", ""))
                opt_d    = latex_to_text(q.get("D", ""))
                answer   = q.get("answer",      "Answer not found in uploaded Mark Scheme")
                ans_ok   = q.get("answerFound", False)
                has_vis  = q.get("hasVisual",   False)
                quote    = q.get("quote",       "")
                topic_   = q.get("topic",       "")
                diff_    = q.get("difficulty",  "")
                marks    = q.get("marks",       1)
                ref      = q.get("ref",         "")

                # ── Question header ────────────────────────────────────────
                qh = doc.add_paragraph()
                _run(qh, f"Question: {q_num}", bold=True, color="1F4E79", size_pt=13)

                # ── Meta line ──────────────────────────────────────────────
                meta_p = doc.add_paragraph()
                _run(meta_p, "Level: ",     bold=True, size_pt=10)
                _run(meta_p, f"{diff_}   ",        size_pt=10, color="333333")
                _run(meta_p, "Marks: ",     bold=True, size_pt=10)
                _run(meta_p, f"{marks}   ",        size_pt=10, color="333333")
                _run(meta_p, "Ref: ",       bold=True, size_pt=10)
                _run(meta_p, ref,                  size_pt=10, color="333333")

                ch_p = doc.add_paragraph()
                _run(ch_p, "Chapter: ", bold=True, size_pt=10)
                _run(ch_p, topic_,             size_pt=10, color="333333")

                # ── Visual: crop real image from QP, embed in Word ─────────
                if has_vis:
                    visual_done += 1
                    if progress_cb:
                        progress_cb(
                            visual_done, total_visual,
                            f"Cropping image for Q{q_num} (page {page_num})…",
                        )
                    img_bytes = b""
                    if qp_bytes and page_num > 0:
                        img_bytes = crop_question_png(client, qp_bytes, page_num, q_num)
                    if img_bytes:
                        try:
                            doc.add_picture(io.BytesIO(img_bytes), width=Cm(14))
                        except Exception:
                            _run(
                                doc.add_paragraph(),
                                "[Image could not be embedded — see QP PDF]",
                                italic=True, color="888888", size_pt=10,
                            )
                    else:
                        _run(
                            doc.add_paragraph(),
                            "[Image not available — see QP PDF]",
                            italic=True, color="888888", size_pt=10,
                        )

                # ── Question stem text ─────────────────────────────────────
                for line in stem.split("\n"):
                    lp = doc.add_paragraph()
                    _run(lp, line, size_pt=11)

                # ── Options ────────────────────────────────────────────────
                for letter, opt_text in [
                    ("A", opt_a), ("B", opt_b), ("C", opt_c), ("D", opt_d)
                ]:
                    if opt_text:
                        op = doc.add_paragraph()
                        op.paragraph_format.left_indent = Cm(1.0)
                        _run(op, f"{letter}.  ", bold=True, size_pt=11)
                        _run(op, opt_text,              size_pt=11)

                # ── Student solution box ───────────────────────────────────
                sl = doc.add_paragraph()
                _run(sl, "Student's Solution:", bold=True, size_pt=11)

                tbl  = doc.add_table(rows=1, cols=1)
                tbl.style = "Table Grid"
                cell = tbl.cell(0, 0)
                _set_cell_borders(cell)
                ip   = cell.paragraphs[0]
                ip.paragraph_format.space_before = Pt(18)
                ip.paragraph_format.space_after  = Pt(18)
                ip.add_run(" ")
                doc.add_paragraph()

                # ── Answer from MS ─────────────────────────────────────────
                ap = doc.add_paragraph()
                _run(ap, "Answer from Mark Scheme:", bold=True, size_pt=11)
                av = doc.add_paragraph()
                _run(av, answer, bold=True,
                     color="C00000" if ans_ok else "888888", size_pt=14)

                # ── Motivational quote ─────────────────────────────────────
                if quote:
                    qp_ = doc.add_paragraph()
                    _run(qp_, "Keep it up:  ", bold=True, italic=True,
                         color="555555", size_pt=10)
                    _run(qp_, quote, italic=True, color="555555", size_pt=10)

                _hr(doc)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN STREAMLIT FLOW
# ══════════════════════════════════════════════════════════════════════════════

if st.button(
    "⚡ Extract & Generate Worksheet",
    type="primary",
    disabled=not (qp_file and ms_file and xl_file),
):
    client = anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])

    # Read QP bytes once — needed for both Claude extraction and image cropping
    qp_bytes = read_bytes(qp_file)
    ms_bytes = read_bytes(ms_file)

    with st.status("Processing…", expanded=True) as status:

        # 1. Excel → metadata only (never question text or answer)
        st.write("📊 Reading Excel (metadata only)…")
        try:
            xl_rows = parse_excel(xl_file)
        except Exception as e:
            st.error(f"Failed to read Excel: {e}")
            st.stop()

        q_nums = [r["qn"] for r in xl_rows]
        meta   = {r["qn"]: r for r in xl_rows}
        st.write(
            f"✅ {len(q_nums)} questions: "
            f"Q{', Q'.join(str(n) for n in q_nums[:6])}"
            f"{'…' if len(q_nums) > 6 else ''}"
        )

        # 2. Extract questions verbatim from QP PDF
        st.write(f"🔍 Extracting {len(q_nums)} questions from QP PDF (verbatim)…")
        try:
            qp_data = extract_questions_from_qp(client, to_b64(qp_bytes), q_nums)
        except Exception as e:
            st.error(f"QP extraction failed: {e}")
            st.stop()

        found_n = sum(1 for q in qp_data if q.get("found") is not False)
        st.write(f"✅ {found_n}/{len(q_nums)} questions extracted from QP")

        not_found_list = [
            n for n in q_nums
            if not any(
                int(q.get("qn", -1)) == n and q.get("found") is not False
                for q in qp_data
            )
        ]
        if not_found_list:
            st.warning(
                f"⚠ Not found in QP: Q{', Q'.join(str(n) for n in not_found_list)}"
            )

        # 3. Extract answers from MS PDF
        st.write("🔑 Extracting answers from Mark Scheme PDF…")
        try:
            ms_data = extract_answers_from_ms(client, to_b64(ms_bytes), q_nums)
        except Exception as e:
            st.error(f"MS extraction failed: {e}")
            st.stop()

        ans_n = sum(1 for v in ms_data.values() if v and v != "NOT_FOUND")
        st.write(f"✅ {ans_n} answers extracted from Mark Scheme")

        # 4. Merge
        st.write("🔗 Merging data…")
        questions: list[dict] = []
        for n in q_nums:
            qp_q        = next((q for q in qp_data if int(q.get("qn", -1)) == n), {})
            ans         = ms_data.get(str(n), "")
            not_found_q = qp_q.get("found") is False or not qp_q
            not_found_a = not ans or ans == "NOT_FOUND"
            m           = meta[n]

            questions.append({
                **m,
                "text":        (
                    "Question not found in uploaded Question Paper"
                    if not_found_q else qp_q.get("text", "")
                ),
                "A":           qp_q.get("A", "") if not not_found_q else "",
                "B":           qp_q.get("B", "") if not not_found_q else "",
                "C":           qp_q.get("C", "") if not not_found_q else "",
                "D":           qp_q.get("D", "") if not not_found_q else "",
                "hasVisual":   qp_q.get("hasVisual", False),
                "found":       not not_found_q,
                "answer":      (
                    "Answer not found in uploaded Mark Scheme"
                    if not_found_a else ans
                ),
                "answerFound": not not_found_a,
            })

        status.update(label="✅ Extraction done!", state="complete")

    # ── Preview ────────────────────────────────────────────────────────────────
    st.subheader("Preview")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total",         len(questions))
    c2.metric("Found in QP",   sum(1 for q in questions if q["found"]))
    c3.metric("Answers found", sum(1 for q in questions if q["answerFound"]))
    c4.metric("Has visuals",   sum(1 for q in questions if q["hasVisual"]))

    st.dataframe(
        [{
            "Q#":         q["qn"],
            "Page":       q.get("page_num", ""),
            "Chapter":    q.get("topic", ""),
            "Difficulty": q.get("difficulty", ""),
            "Marks":      q.get("marks", ""),
            "Answer":     q.get("answer", ""),
            "Found?":     "✅" if q["found"] else "❌",
            "Visual?":    "🖼" if q["hasVisual"] else "",
        } for q in questions],
        use_container_width=True,
        hide_index=True,
    )

    # ── Build Word ─────────────────────────────────────────────────────────────
    st.subheader("Download")
    visual_count = sum(1 for q in questions if q["hasVisual"])

    if visual_count:
        st.info(
            f"ℹ️ {visual_count} questions contain visual content (diagrams/graphs). "
            "Actual images will be cropped from the QP PDF and embedded in the Word file. "
            "This may take ~10–20 seconds."
        )

    prog_bar  = st.progress(0)
    prog_text = st.empty()

    def on_progress(done, total, msg):
        prog_bar.progress(done / max(total, 1))
        prog_text.text(msg)

    with st.spinner("Building Word document…"):
        try:
            docx_bytes = build_word_document(
                questions,
                qp_bytes,
                client,
                progress_cb=on_progress if visual_count else None,
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
