import streamlit as st
import anthropic
import openpyxl
import json
import re
import io
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import base64

# ─── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Exam Worksheet Generator",
    page_icon="📄",
    layout="centered"
)

st.title("📄 Exam Worksheet Generator")
st.caption("Extracts questions from QP · answers from MS · metadata from Excel")

# ─── Extraction rules reminder ────────────────────────────────────────────────
with st.expander("ℹ️ Source rules", expanded=False):
    st.markdown("""
| Source | Used for |
|--------|----------|
| **QP PDF** | Question text, options, tables, formulas — verbatim |
| **MS PDF** | Answers only |
| **Excel** | Question numbers, chapter, difficulty, marks, quotes |

- Excel is **never** used as source for question text or answers
- If question not found in QP → *"Question not found in uploaded Question Paper"*
- If answer not found in MS → *"Answer not found in uploaded Mark Scheme"*
""")

# ─── File uploads ─────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    qp_file = st.file_uploader("📋 Question Paper (QP)", type=["pdf"], key="qp")
    ms_file = st.file_uploader("✅ Mark Scheme (MS)", type=["pdf"], key="ms")
with col2:
    xl_file = st.file_uploader("📊 Excel Sheet", type=["xlsx", "xls"], key="xl")

st.divider()

# ─── Helpers ──────────────────────────────────────────────────────────────────

def file_to_b64(uploaded_file) -> str:
    uploaded_file.seek(0)
    return base64.standard_b64encode(uploaded_file.read()).decode("utf-8")


def parse_excel(uploaded_file):
    """Return list of dicts with question metadata from Excel.
    Only reads: question number, topic/chapter, difficulty, marks, quote.
    Never uses Excel as source for question text or answer.
    """
    uploaded_file.seek(0)
    wb = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)
    ws = wb.active

    headers = [str(c.value).strip() if c.value else "" for c in next(ws.iter_rows(max_row=1))]

    # Detect question-number column
    qn_col = None
    for h in ["Question No.", "Question No", "Q No", "Q#", "Question Number", "Q"]:
        if h in headers:
            qn_col = headers.index(h)
            break
    if qn_col is None:
        # Fall back: third column (index 2)
        qn_col = 2

    def get(row, name, fallback=""):
        try:
            idx = headers.index(name)
            v = list(row)[idx].value
            return str(v).strip() if v is not None else fallback
        except (ValueError, IndexError):
            return fallback

    rows = []
    for row in ws.iter_rows(min_row=2):
        raw_qn = list(row)[qn_col].value
        try:
            qn = int(float(str(raw_qn)))
        except (TypeError, ValueError):
            continue
        if qn <= 0:
            continue
        rows.append({
            "qn": qn,
            "topic":      get(row, "Topic")      or get(row, "Chapter") or "",
            "difficulty": get(row, "Difficulty") or "",
            "marks":      get(row, "Marks", "1") or "1",
            "quote":      get(row, "Quote")      or "",
            "ref":        get(row, "Reference")  or "",
        })
    return rows


def safe_json(text: str):
    text = re.sub(r"```json\s*", "", text)
    text = re.sub(r"```\s*", "", text)
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return None


def extract_questions_from_qp(client: anthropic.Anthropic, b64_pdf: str, q_nums: list[int]):
    """Call Claude with the QP PDF and return structured question data."""
    nums_str = ", ".join(str(n) for n in q_nums)
    response = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=8000,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": b64_pdf
                    }
                },
                {
                    "type": "text",
                    "text": f"""Extract ONLY these question numbers from the exam paper: {nums_str}

STRICT RULES — do NOT violate:
1. Copy text VERBATIM — never paraphrase, summarise, or shorten
2. Include complete question stem and all 4 options (A, B, C, D) exactly as printed
3. For tables: represent as plain-text rows e.g. "| Col1 | Col2 |\\n| Val1 | Val2 |"
4. Keep formulas/equations as written (e.g. fractions, arrows, sub/superscripts)
5. If question has a diagram/graph/image that cannot be text: set hasVisual:true and describe briefly in visualDesc
6. If a question number is NOT FOUND in the paper: set found:false and text:"Question not found in uploaded Question Paper"

Return ONLY this JSON array — no markdown, no preamble, nothing else:
[{{"qn":<int>,"text":"<verbatim stem>","A":"<option A>","B":"<option B>","C":"<option C>","D":"<option D>","hasVisual":<bool>,"visualDesc":"<desc or null>","found":<bool>}}]"""
                }
            ]
        }]
    )
    raw = "".join(b.text for b in response.content if hasattr(b, "text"))
    data = safe_json(raw)
    if not isinstance(data, list):
        return []
    return data


def extract_answers_from_ms(client: anthropic.Anthropic, b64_pdf: str, q_nums: list[int]):
    """Call Claude with the MS PDF and return {qn_str: letter} dict."""
    nums_str = ", ".join(str(n) for n in q_nums)
    response = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=2000,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": b64_pdf
                    }
                },
                {
                    "type": "text",
                    "text": f"""Extract the correct answers for these question numbers from the mark scheme: {nums_str}

The mark scheme typically shows a grid like:
1. D    16. C    31. C
2. A    17. D    32. B

Return ONLY this JSON object — no markdown, no explanation:
{{"1":"D","2":"A","3":"B",...}}

For any question not found use "NOT_FOUND"."""
                }
            ]
        }]
    )
    raw = "".join(b.text for b in response.content if hasattr(b, "text"))
    data = safe_json(raw)
    return data if isinstance(data, dict) else {}


# ─── LaTeX → Unicode ──────────────────────────────────────────────────────────
SUPS = {"0":"⁰","1":"¹","2":"²","3":"³","4":"⁴","5":"⁵","6":"⁶","7":"⁷","8":"⁸","9":"⁹",
        "+":"⁺","-":"⁻","n":"ⁿ","m":"ᵐ","x":"ˣ","a":"ᵃ","b":"ᵇ"}
SUBS = {"0":"₀","1":"₁","2":"₂","3":"₃","4":"₄","5":"₅","6":"₆","7":"₇","8":"₈","9":"₉",
        "+":"₊","-":"₋","n":"ₙ","e":"ₑ","r":"ᵣ","x":"ₓ","a":"ₐ","i":"ᵢ"}

def _to_sup(s): return "".join(SUPS.get(c, c) for c in s)
def _to_sub(s): return "".join(SUBS.get(c, c) for c in s)

def _convert_math(m: str) -> str:
    t = m
    t = re.sub(r"\\frac\{([^{}]+)\}\{([^{}]+)\}", r"(\1)/(\2)", t)
    replacements = {
        r"\times": "×", r"\rightarrow": "→", r"\rightleftharpoons": "⇌",
        r"\leftrightarrow": "↔", r"\to": "→", r"\Delta": "Δ", r"\delta": "δ",
        r"\ominus": "⊖", r"\oplus": "⊕", r"\theta": "θ", r"\alpha": "α",
        r"\beta": "β",  r"\gamma": "γ", r"\lambda": "λ", r"\mu": "μ",
        r"\pi": "π", r"\sigma": "σ", r"\cdot": "·", r"\pm": "±",
        r"\geq": "≥", r"\leq": "≤", r"\neq": "≠", r"\approx": "≈",
        r"\infty": "∞", r"\circ": "°",
    }
    for k, v in replacements.items():
        t = t.replace(k, v)
    t = re.sub(r"\^\{([^{}]{1,12})\}", lambda m: _to_sup(m.group(1)), t)
    t = re.sub(r"\^([a-zA-Z0-9+\-])",  lambda m: _to_sup(m.group(1)), t)
    t = re.sub(r"_\{([^{}]{1,12})\}",  lambda m: _to_sub(m.group(1)), t)
    t = re.sub(r"_([a-zA-Z0-9])",      lambda m: _to_sub(m.group(1)), t)
    t = re.sub(r"\\(?:text|mathrm|mbox)\{([^{}]+)\}", r"\1", t)
    t = re.sub(r"[{}]", "", t)
    return t.strip()

def latex_to_text(text: str) -> str:
    if not text:
        return ""
    t = str(text)
    t = re.sub(r"\$\$([^$]+)\$\$", lambda m: _convert_math(m.group(1)), t)
    t = re.sub(r"\$([^$\n]+)\$",   lambda m: _convert_math(m.group(1)), t)
    t = re.sub(r"\\\[([^\]]+)\\\]",lambda m: _convert_math(m.group(1)), t)
    t = re.sub(r"\*\*([^*]+)\*\*", r"\1", t)
    t = re.sub(r"\*([^*]+)\*",     r"\1", t)
    return t


# ─── Word document builder ────────────────────────────────────────────────────
DIFF_ORDER = ["Easy", "Medium", "Hard"]

def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), "999999")
        tcPr.append(el)

def add_run(para, text, bold=False, italic=False, color=None, size_pt=11):
    run = para.add_run(str(text))
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size_pt)
    run.font.name = "Arial"
    if color:
        run.font.color.rgb = RGBColor(*bytes.fromhex(color))
    return run

def build_word_document(questions: list[dict]) -> bytes:
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin    = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # Title
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(title_para, "IB Chemistry – Higher Level", bold=True,
            color="1F4E79", size_pt=18)
    sub_para = doc.add_paragraph()
    sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(sub_para, "Exam Worksheet", color="555555", size_pt=12)
    doc.add_paragraph()

    # Group by topic → difficulty
    grouped: dict[str, dict[str, list]] = {}
    for q in questions:
        t = q.get("topic") or "Other"
        d = q.get("difficulty") or "Unspecified"
        grouped.setdefault(t, {}).setdefault(d, []).append(q)

    for topic, diffs in grouped.items():
        # Chapter heading
        h2 = doc.add_heading(topic, level=2)
        h2.runs[0].font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

        sorted_diffs = sorted(diffs.keys(),
                              key=lambda d: DIFF_ORDER.index(d) if d in DIFF_ORDER else 99)
        for diff in sorted_diffs:
            # Difficulty sub-heading
            diff_para = doc.add_paragraph()
            add_run(diff_para, f"— {diff} questions —", italic=True, color="666666", size_pt=11)

            for q in diffs[diff]:
                qn       = q["qn"]
                stem     = latex_to_text(q.get("text", ""))
                opt_a    = latex_to_text(q.get("A", ""))
                opt_b    = latex_to_text(q.get("B", ""))
                opt_c    = latex_to_text(q.get("C", ""))
                opt_d    = latex_to_text(q.get("D", ""))
                answer   = q.get("answer", "Answer not found in uploaded Mark Scheme")
                ans_ok   = q.get("answerFound", False)
                has_vis  = q.get("hasVisual", False)
                vis_desc = q.get("visualDesc") or ""
                quote    = q.get("quote", "")
                topic_   = q.get("topic", "")
                diff_    = q.get("difficulty", "")
                marks    = q.get("marks", "1")
                ref      = q.get("ref", "")

                # Question header
                qh = doc.add_paragraph()
                add_run(qh, f"Question: {qn}", bold=True, color="1F4E79", size_pt=13)

                # Meta line
                meta = doc.add_paragraph()
                add_run(meta, "Level: ", bold=True, size_pt=10)
                add_run(meta, f"{diff_}   ", size_pt=10, color="333333")
                add_run(meta, "Marks: ", bold=True, size_pt=10)
                add_run(meta, f"{marks}   ", size_pt=10, color="333333")
                add_run(meta, "Ref: ", bold=True, size_pt=10)
                add_run(meta, ref, size_pt=10, color="333333")

                # Chapter line
                ch_para = doc.add_paragraph()
                add_run(ch_para, "Chapter: ", bold=True, size_pt=10)
                add_run(ch_para, topic_, size_pt=10, color="333333")

                # Visual warning
                if has_vis and vis_desc:
                    vis_para = doc.add_paragraph()
                    add_run(vis_para,
                            f"⚠ Visual content: {vis_desc}. Refer to the original exam paper.",
                            italic=True, color="888888", size_pt=10)

                # Question stem (may be multi-line)
                for line in stem.split("\n"):
                    p = doc.add_paragraph()
                    add_run(p, line, size_pt=11)

                # Options
                for letter, text in [("A", opt_a), ("B", opt_b),
                                      ("C", opt_c), ("D", opt_d)]:
                    if text:
                        op = doc.add_paragraph()
                        op.paragraph_format.left_indent = Cm(1.0)
                        add_run(op, f"{letter}.  ", bold=True, size_pt=11)
                        add_run(op, text, size_pt=11)

                # Student solution box
                sol_label = doc.add_paragraph()
                add_run(sol_label, "Student's Solution:", bold=True, size_pt=11)

                tbl = doc.add_table(rows=1, cols=1)
                tbl.style = "Table Grid"
                cell = tbl.cell(0, 0)
                cell.width = Cm(15)
                set_cell_border(cell)
                inner = cell.paragraphs[0]
                inner.add_run(" " * 5)
                inner.paragraph_format.space_before = Pt(18)
                inner.paragraph_format.space_after  = Pt(18)
                doc.add_paragraph()

                # Answer
                ans_label = doc.add_paragraph()
                add_run(ans_label, "Answer from Mark Scheme:", bold=True, size_pt=11)

                ans_para = doc.add_paragraph()
                color = "C00000" if ans_ok else "888888"
                add_run(ans_para, answer, bold=True, color=color, size_pt=14)

                # Quote
                if quote:
                    qp_ = doc.add_paragraph()
                    add_run(qp_, "Keep it up:  ", bold=True, italic=True,
                            color="555555", size_pt=10)
                    add_run(qp_, quote, italic=True, color="555555", size_pt=10)

                # Horizontal rule via bottom border on empty paragraph
                hr = doc.add_paragraph()
                pPr = hr._p.get_or_add_pPr()
                pBdr = OxmlElement("w:pBdr")
                bottom = OxmlElement("w:bottom")
                bottom.set(qn("w:val"), "single")
                bottom.set(qn("w:sz"), "4")
                bottom.set(qn("w:color"), "CCCCCC")
                pBdr.append(bottom)
                pPr.append(pBdr)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ─── Main generate button ─────────────────────────────────────────────────────
if st.button("⚡ Extract & Generate Worksheet", type="primary",
             disabled=not (qp_file and ms_file and xl_file)):

    client = anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])

    with st.status("Processing…", expanded=True) as status:

        # 1. Parse Excel (metadata only)
        st.write("📊 Reading Excel — extracting metadata only…")
        try:
            xl_rows = parse_excel(xl_file)
        except Exception as e:
            st.error(f"Failed to read Excel: {e}")
            st.stop()

        q_nums = [r["qn"] for r in xl_rows]
        meta   = {r["qn"]: r for r in xl_rows}
        st.write(f"✅ Found **{len(q_nums)}** questions: Q{', Q'.join(str(n) for n in q_nums[:6])}{'…' if len(q_nums) > 6 else ''}")

        # 2. Load PDFs as base64
        st.write("📄 Loading PDF files…")
        qp_b64 = file_to_b64(qp_file)
        ms_b64 = file_to_b64(ms_file)
        st.write("✅ PDFs loaded")

        # 3. Extract questions from QP PDF — verbatim, never from Excel
        st.write(f"🔍 Extracting **{len(q_nums)}** questions from QP PDF…")
        try:
            qp_data = extract_questions_from_qp(client, qp_b64, q_nums)
        except Exception as e:
            st.error(f"QP extraction failed: {e}")
            st.stop()

        found_count = sum(1 for q in qp_data if q.get("found") is not False)
        st.write(f"✅ Extracted **{found_count}/{len(q_nums)}** questions from QP")

        if found_count < len(q_nums):
            not_found = [n for n in q_nums
                         if not any(int(q.get("qn", -1)) == n and q.get("found") is not False
                                    for q in qp_data)]
            st.warning(f"⚠ Not found in QP: Q{', Q'.join(str(n) for n in not_found)}")

        # 4. Extract answers from MS PDF — never from Excel
        st.write("🔑 Extracting answers from Mark Scheme PDF…")
        try:
            ms_data = extract_answers_from_ms(client, ms_b64, q_nums)
        except Exception as e:
            st.error(f"MS extraction failed: {e}")
            st.stop()

        ans_count = sum(1 for v in ms_data.values() if v and v != "NOT_FOUND")
        st.write(f"✅ Extracted **{ans_count}** answers from Mark Scheme")

        # 5. Merge
        st.write("🔗 Merging data…")
        questions = []
        for n in q_nums:
            qp_q = next((q for q in qp_data if int(q.get("qn", -1)) == n), {})
            ans  = ms_data.get(str(n), "")
            not_found_q   = qp_q.get("found") is False or not qp_q
            not_found_ans = not ans or ans == "NOT_FOUND"
            questions.append({
                **meta[n],
                "text":       "Question not found in uploaded Question Paper" if not_found_q else qp_q.get("text", ""),
                "A":          qp_q.get("A", "") if not not_found_q else "",
                "B":          qp_q.get("B", "") if not not_found_q else "",
                "C":          qp_q.get("C", "") if not not_found_q else "",
                "D":          qp_q.get("D", "") if not not_found_q else "",
                "hasVisual":  qp_q.get("hasVisual", False),
                "visualDesc": qp_q.get("visualDesc"),
                "found":      not not_found_q,
                "answer":     "Answer not found in uploaded Mark Scheme" if not_found_ans else ans,
                "answerFound": not not_found_ans,
            })

        status.update(label="✅ Done!", state="complete")

    # ─── Preview table ─────────────────────────────────────────────────────
    st.subheader("Preview")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total",        len(questions))
    c2.metric("Found in QP",  sum(1 for q in questions if q["found"]))
    c3.metric("Answers found",sum(1 for q in questions if q["answerFound"]))
    c4.metric("Has visuals",  sum(1 for q in questions if q["hasVisual"]))

    preview_rows = [{
        "Q#":        q["qn"],
        "Chapter":   q.get("topic",""),
        "Difficulty":q.get("difficulty",""),
        "Marks":     q.get("marks",""),
        "Answer":    q.get("answer",""),
        "Found?":    "✅" if q["found"] else "❌",
    } for q in questions]
    st.dataframe(preview_rows, use_container_width=True, hide_index=True)

    # ─── Generate & download Word ──────────────────────────────────────────
    st.subheader("Download")
    with st.spinner("Building Word document…"):
        try:
            docx_bytes = build_word_document(questions)
        except Exception as e:
            st.error(f"Word generation failed: {e}")
            st.stop()

    st.download_button(
        label="⬇️ Download Worksheet (.docx)",
        data=docx_bytes,
        file_name="IB_Chemistry_Worksheet.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
