"""
app.py — IB Chemistry Question Extractor
========================================
Reads an Excel question bank + a combined QP PDF,
matches each row via composite key (Reference + Page Number),
handles "Extraction Failed" rows by cropping PDF pages as images,
and exports a formatted Word (.docx) document.

Composite key logic
-------------------
Reference alone is NOT unique — May 2021 and May 2022 each have two
separate papers (TZ1 / TZ2) sharing the same date string.
We disambiguate by grouping consecutive Q#=1 rows within the same
Reference and assigning an auto-detected paper sub-index.
The actual PDF page number (Page Number column) is the ground truth
for which page to crop when a question needs an image fallback.
"""

import io
import re
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

# ── optional heavy deps (imported lazily so the preview works even if missing)
try:
    import fitz  # PyMuPDF  — pip install pymupdf
    PYMUPDF_OK = True
except ImportError:
    PYMUPDF_OK = False

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    DOCX_OK = True
except ImportError:
    DOCX_OK = False


# ════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ════════════════════════════════════════════════════════════════════════════

EXTRACTION_FAILED_MARKER = "Extraction Failed"

# Colour palette for Word document headers
COLOUR_HEADER_BG  = "2E4057"   # dark navy  — section header
COLOUR_META_BG    = "F0F4F8"   # light grey — metadata row
COLOUR_LABEL_TEXT = "2E4057"   # navy text

# PDF crop resolution (DPI) when rasterising a page
PDF_DPI = 150


# ════════════════════════════════════════════════════════════════════════════
# EXCEL HELPERS
# ════════════════════════════════════════════════════════════════════════════

# Canonical column name aliases → internal key
COLUMN_ALIASES = {
    "question no.":   "q_num",
    "question no":    "q_num",
    "q#":             "q_num",
    "question number":"q_num",
    "question text":  "q_text",
    "question":       "q_text",
    "reference":      "reference",
    "topic":          "topic",
    "difficulty":     "difficulty",
    "marks":          "marks",
    "mark scheme":    "ms",
    "markscheme":     "ms",
    "mark_scheme":    "ms",
    "answer":         "ms",
    "page number":    "page_num",
    "page no":        "page_num",
    "page":           "page_num",
    "file name":      "file_name",
    "paper":          "paper",
    "session":        "session",
    "year":           "year",
    "quote":          "quote",
}


def normalise_columns(df: pd.DataFrame) -> dict:
    """
    Return a mapping {internal_key: actual_column_name} for every
    column we recognise.  Unknown columns are kept but not mapped.
    """
    mapping = {}
    for col in df.columns:
        alias = col.strip().lower()
        if alias in COLUMN_ALIASES:
            key = COLUMN_ALIASES[alias]
            mapping[key] = col
    return mapping


def load_excel(file) -> tuple[pd.DataFrame, dict]:
    """Read the uploaded Excel and return (dataframe, column_mapping)."""
    df = pd.read_excel(file, dtype=str)
    df = df.fillna("")
    mapping = normalise_columns(df)
    return df, mapping


# ════════════════════════════════════════════════════════════════════════════
# PAPER IDENTIFICATION  (composite-key logic)
# ════════════════════════════════════════════════════════════════════════════

def assign_paper_id(df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    """
    Add a 'paper_id' column that uniquely identifies each exam paper.

    Strategy
    --------
    Within each Reference group, every time Question No. resets to "1"
    we start a new sub-paper.  This works because the rows are already
    ordered by paper then by question number.

    The resulting paper_id looks like:
        "(Friday 14 May 2021)__P1"   ← first  paper in that date
        "(Friday 14 May 2021)__P2"   ← second paper in that date
        " (Wednesday 10 November 2021)__P1"

    The composite key used for matching is:
        (reference, paper_id, q_num)
    """
    ref_col = mapping.get("reference", None)
    q_col   = mapping.get("q_num",     None)
    page_col = mapping.get("page_num", None)

    paper_ids = []
    paper_counters: dict[str, int] = {}
    prev_q: dict[str, int] = {}

    for _, row in df.iterrows():
        ref   = row[ref_col].strip()  if ref_col  else "UNKNOWN_REF"
        q_raw = row[q_col].strip()    if q_col    else "0"
        try:
            q_int = int(float(q_raw))
        except ValueError:
            q_int = 0

        # Start a new sub-paper when Q# resets to 1
        if ref not in paper_counters:
            paper_counters[ref] = 1
            prev_q[ref] = q_int
        else:
            # Q# went back to a small value → new paper detected
            if q_int < prev_q.get(ref, 999):
                paper_counters[ref] += 1
            prev_q[ref] = q_int

        pid = f"{ref}__P{paper_counters[ref]}"
        paper_ids.append(pid)

    df = df.copy()
    df["paper_id"] = paper_ids
    return df


def build_paper_page_ranges(df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    """
    For each paper_id, compute min_page and max_page from the Page Number column.
    Returns a summary dataframe.
    """
    page_col = mapping.get("page_num", None)
    if not page_col:
        return pd.DataFrame()

    df2 = df.copy()
    df2["_page_int"] = pd.to_numeric(df2[page_col], errors="coerce")

    summary = (
        df2.groupby("paper_id")
        .agg(
            reference   = (mapping.get("reference","paper_id"), "first"),
            min_page    = ("_page_int", "min"),
            max_page    = ("_page_int", "max"),
            num_questions = ("_page_int", "count"),
        )
        .reset_index()
    )
    return summary


# ════════════════════════════════════════════════════════════════════════════
# PDF HELPERS
# ════════════════════════════════════════════════════════════════════════════

def pdf_page_to_image_bytes(pdf_bytes: bytes, page_index: int, dpi: int = PDF_DPI) -> bytes | None:
    """
    Rasterise one page of a PDF (0-based index) → PNG bytes.
    Returns None if PyMuPDF is unavailable or the page index is out of range.
    """
    if not PYMUPDF_OK:
        return None
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if page_index < 0 or page_index >= doc.page_count:
            return None
        page = doc[page_index]
        mat  = fitz.Matrix(dpi / 72, dpi / 72)
        pix  = page.get_pixmap(matrix=mat, alpha=False)
        return pix.tobytes("png")
    except Exception:
        return None


def extract_text_from_pdf_page(pdf_bytes: bytes, page_index: int) -> str:
    """Extract plain text from a single PDF page (0-based index)."""
    if not PYMUPDF_OK:
        return ""
    try:
        doc  = fitz.open(stream=pdf_bytes, filetype="pdf")
        if page_index < 0 or page_index >= doc.page_count:
            return ""
        return doc[page_index].get_text("text").strip()
    except Exception:
        return ""


# ════════════════════════════════════════════════════════════════════════════
# WORD DOCUMENT BUILDER
# ════════════════════════════════════════════════════════════════════════════

def hex_to_rgb(hex_str: str):
    """'2E4057' → RGBColor(0x2E, 0x40, 0x57)"""
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _set_cell_bg(cell, hex_colour: str):
    """Apply a solid background fill to a table cell."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_colour.upper())
    tcPr.append(shd)


def _set_cell_borders(cell, colour: str = "CCCCCC"):
    """Add thin borders on all four sides of a table cell."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"),   "single")
        el.set(qn("w:sz"),    "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), colour)
        tcBorders.append(el)
    tcPr.append(tcBorders)


def add_meta_table(doc: "Document", meta: dict):
    """
    Render a two-column metadata table (Label | Value) with light-grey background.
    """
    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"

    col_widths = [Inches(1.8), Inches(5.0)]

    for label, value in meta.items():
        row = table.add_row()
        for i, text in enumerate((label, str(value))):
            cell = row.cells[i]
            cell.width = col_widths[i]
            _set_cell_bg(cell, COLOUR_META_BG)
            _set_cell_borders(cell)
            p = cell.paragraphs[0]
            run = p.add_run(text)
            run.font.size = Pt(10)
            if i == 0:
                run.bold = True
                run.font.color.rgb = hex_to_rgb(COLOUR_LABEL_TEXT)

    doc.add_paragraph()  # spacer


def add_section_heading(doc: "Document", title: str):
    """Dark-navy full-width heading row for each question block."""
    table = doc.add_table(rows=1, cols=1)
    cell  = table.cell(0, 0)
    _set_cell_bg(cell, COLOUR_HEADER_BG)
    _set_cell_borders(cell, "FFFFFF")
    p   = cell.paragraphs[0]
    run = p.add_run(title)
    run.bold = True
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    run.font.size = Pt(11)
    doc.add_paragraph()  # spacer


def add_labelled_block(doc: "Document", label: str, content: str, is_image: bool = False,
                       image_bytes: bytes | None = None):
    """
    Add a bold label paragraph followed by the content (text or image).
    """
    lbl = doc.add_paragraph()
    run = lbl.add_run(f"{label}:")
    run.bold = True
    run.font.color.rgb = hex_to_rgb(COLOUR_LABEL_TEXT)
    run.font.size = Pt(10)

    if is_image and image_bytes:
        img_para = doc.add_paragraph()
        run_img  = img_para.add_run()
        run_img.add_picture(io.BytesIO(image_bytes), width=Inches(6.0))
        note = doc.add_paragraph()
        note_run = note.add_run("⚠ Extraction failed — question shown as cropped PDF image.")
        note_run.italic = True
        note_run.font.size = Pt(9)
        note_run.font.color.rgb = RGBColor(0xCC, 0x44, 0x00)
    elif content:
        txt = doc.add_paragraph(content)
        txt.runs[0].font.size = Pt(10) if txt.runs else None
    else:
        txt = doc.add_paragraph("(not available)")
        txt.runs[0].italic = True

    doc.add_paragraph()  # spacer after block


def build_word_document(
    df: pd.DataFrame,
    mapping: dict,
    pdf_bytes: bytes | None,
) -> bytes:
    """
    Build and return the .docx file as bytes.

    For each row in df:
    - If Question Text contains EXTRACTION_FAILED_MARKER and pdf_bytes is given,
      crop the relevant PDF page as a PNG image instead.
    - Mark Scheme is taken directly from the Mark Scheme column in Excel.
    """
    if not DOCX_OK:
        raise ImportError("python-docx is not installed.  Run: pip install python-docx")

    doc = Document()

    # ── Document title
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title_p.add_run("IB Chemistry — Question Bank")
    tr.bold      = True
    tr.font.size = Pt(16)
    tr.font.color.rgb = hex_to_rgb(COLOUR_HEADER_BG)
    doc.add_paragraph()

    # ── Column references
    ref_col   = mapping.get("reference")
    q_col     = mapping.get("q_num")
    qtext_col = mapping.get("q_text")
    ms_col    = mapping.get("ms")
    topic_col = mapping.get("topic")
    diff_col  = mapping.get("difficulty")
    page_col  = mapping.get("page_num")
    marks_col = mapping.get("marks")

    for idx, row in df.iterrows():
        q_num    = row.get(q_col,     "?")   if q_col     else "?"
        ref      = row.get(ref_col,   "")    if ref_col   else ""
        paper_id = row.get("paper_id","")
        topic    = row.get(topic_col, "")    if topic_col else ""
        diff     = row.get(diff_col,  "")    if diff_col  else ""
        marks    = row.get(marks_col, "")    if marks_col else ""
        page_raw = row.get(page_col,  "")    if page_col  else ""
        q_text   = row.get(qtext_col, "")    if qtext_col else ""
        ms_text  = row.get(ms_col,    "")    if ms_col    else ""

        # Human-readable paper label (P1, P2 …)
        paper_label = paper_id.split("__")[-1] if "__" in paper_id else paper_id

        # ── Section heading
        heading = f"Q{q_num}  ·  {ref.strip()}  ·  {paper_label}"
        add_section_heading(doc, heading)

        # ── Metadata block
        meta = {}
        meta["Reference"]      = ref.strip()
        meta["Paper"]          = paper_label
        if page_raw:
            meta["PDF Page"]   = page_raw
        if topic:
            meta["Topic"]      = topic
        if diff:
            meta["Difficulty"] = diff
        if marks:
            meta["Marks"]      = marks
        add_meta_table(doc, meta)

        # ── Question block
        failed = EXTRACTION_FAILED_MARKER in q_text

        if failed and pdf_bytes and PYMUPDF_OK:
            # Crop the PDF page → image
            try:
                page_index = int(float(page_raw)) - 1  # Excel stores 1-based page numbers
            except (ValueError, TypeError):
                page_index = -1
            img_bytes = pdf_page_to_image_bytes(pdf_bytes, page_index)
            add_labelled_block(doc, "Question", "", is_image=True, image_bytes=img_bytes)
        elif failed:
            # PyMuPDF not available or no PDF — show placeholder
            add_labelled_block(doc, "Question",
                "⚠ Extraction failed.  Please supply the QP PDF to embed the page image.")
        else:
            add_labelled_block(doc, "Question", q_text)

        # ── Mark Scheme block
        add_labelled_block(doc, "Mark Scheme / Answer", ms_text)

        # ── Page break between questions (except the last)
        if idx < len(df) - 1:
            doc.add_page_break()

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════════════
# STREAMLIT UI
# ════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="IB Chem — Question Extractor",
    page_icon="🧪",
    layout="wide",
)

st.title("🧪 IB Chemistry — Question Extractor")
st.caption(
    "Upload your Excel question bank and (optionally) the combined QP PDF. "
    "The app uses **Reference + Page Number** as a composite key so that "
    "Q1–Q40 repeated across multiple exam papers are handled correctly."
)

# ── Dependency warnings
if not PYMUPDF_OK:
    st.warning("⚠ PyMuPDF not installed — PDF image fallback is disabled.  "
               "Install with: `pip install pymupdf`", icon="⚠")
if not DOCX_OK:
    st.warning("⚠ python-docx not installed — Word export is disabled.  "
               "Install with: `pip install python-docx`", icon="⚠")

st.divider()

# ════════════════════════════════════════════════════════════════════════════
# STEP 1 — FILE UPLOAD
# ════════════════════════════════════════════════════════════════════════════
st.header("📁 Step 1 — Upload Files")

col_xl, col_pdf = st.columns(2)

with col_xl:
    uploaded_excel = st.file_uploader(
        "Excel Question Bank (.xlsx)",
        type=["xlsx", "xls"],
        key="excel_upload",
    )

with col_pdf:
    uploaded_pdf = st.file_uploader(
        "Combined QP PDF (optional — used for failed extractions)",
        type=["pdf"],
        key="pdf_upload",
    )

if not uploaded_excel:
    st.info("👆 Upload the Excel file to continue.")
    st.stop()

# ════════════════════════════════════════════════════════════════════════════
# LOAD & ANALYSE EXCEL
# ════════════════════════════════════════════════════════════════════════════

with st.spinner("Reading Excel …"):
    df_raw, mapping = load_excel(uploaded_excel)
    df = assign_paper_id(df_raw, mapping)
    summary = build_paper_page_ranges(df, mapping)

pdf_bytes: bytes | None = uploaded_pdf.read() if uploaded_pdf else None

# ════════════════════════════════════════════════════════════════════════════
# STEP 2 — PREVIEW / ANALYSIS
# ════════════════════════════════════════════════════════════════════════════
st.header("📊 Step 2 — Excel Analysis Preview")

ref_col   = mapping.get("reference")
q_col     = mapping.get("q_num")
page_col  = mapping.get("page_num")
qtext_col = mapping.get("q_text")
ms_col    = mapping.get("ms")

# ── KPI cards
total_rows       = len(df)
q_col_vals       = df[q_col].astype(str) if q_col else pd.Series(dtype=str)
dup_q_count      = int(q_col_vals.duplicated().sum()) if q_col else 0
ref_col_vals     = df[ref_col].astype(str) if ref_col else pd.Series(dtype=str)
dup_ref_count    = int(ref_col_vals.duplicated().sum()) if ref_col else 0
failed_count     = int(df[qtext_col].str.contains(EXTRACTION_FAILED_MARKER, na=False).sum()) \
                   if qtext_col else 0
num_papers       = df["paper_id"].nunique()

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Total Rows",         total_rows)
k2.metric("Exam Papers Detected", num_papers)
k3.metric("Duplicate Q#",       dup_q_count)
k4.metric("Duplicate Reference", dup_ref_count)
k5.metric("Extraction Failures", failed_count)

st.divider()

# ── Composite key columns used
st.subheader("🔑 Composite Key Columns Used for Matching")
used_cols = {k: v for k, v in mapping.items()
             if k in ("reference", "q_num", "page_num", "paper", "session", "year")}
key_df = pd.DataFrame(
    [(k, v) for k, v in used_cols.items()],
    columns=["Internal Key", "Excel Column Name"],
)
st.dataframe(key_df, use_container_width=True, hide_index=True)

# ── All detected columns
with st.expander("All recognised columns"):
    all_col_df = pd.DataFrame(
        [(k, v) for k, v in mapping.items()],
        columns=["Internal Key", "Excel Column"],
    )
    st.dataframe(all_col_df, use_container_width=True, hide_index=True)

st.divider()

# ── Paper summary table
st.subheader("📄 Detected Exam Papers & Page Ranges")
st.caption(
    "Papers with the same Reference are disambiguated by Q# reset detection. "
    "P1 = first paper found under that Reference, P2 = second, etc."
)
if not summary.empty:
    st.dataframe(
        summary.rename(columns={
            "paper_id":      "Paper ID (composite key)",
            "reference":     "Reference",
            "min_page":      "First PDF Page",
            "max_page":      "Last PDF Page",
            "num_questions": "Questions",
        }),
        use_container_width=True,
        hide_index=True,
    )
else:
    st.warning("Could not compute page ranges — 'Page Number' column not found.")

st.divider()

# ── Composite key examples
st.subheader("🔍 Composite Key Examples")
ex_cols = [c for c in [ref_col, q_col, page_col, "paper_id"] if c and c in df.columns]
if ex_cols:
    st.dataframe(
        df[ex_cols].head(12).rename(columns={
            ref_col:   "Reference",
            q_col:     "Q#",
            page_col:  "Page Number",
            "paper_id":"Paper ID",
        }),
        use_container_width=True,
        hide_index=True,
    )

st.divider()

# ── Extraction failures detail
if failed_count:
    with st.expander(f"⚠ {failed_count} rows with 'Extraction Failed' (will use PDF image fallback)"):
        fail_mask  = df[qtext_col].str.contains(EXTRACTION_FAILED_MARKER, na=False) if qtext_col else pd.Series(False, index=df.index)
        fail_rows  = df[fail_mask]
        show_cols  = [c for c in [q_col, ref_col, page_col, "paper_id"] if c and c in df.columns]
        st.dataframe(fail_rows[show_cols].rename(columns={
            q_col:     "Q#",
            ref_col:   "Reference",
            page_col:  "Page Number",
            "paper_id":"Paper ID",
        }), use_container_width=True, hide_index=True)

        if pdf_bytes:
            st.success("✅ QP PDF uploaded — images will be embedded for these rows.")
        else:
            st.warning("No QP PDF uploaded — failed rows will show a text placeholder.")

# ════════════════════════════════════════════════════════════════════════════
# STEP 3 — GENERATE WORD DOCUMENT
# ════════════════════════════════════════════════════════════════════════════
st.header("📝 Step 3 — Generate Word Document")

# ── Paper filter (optional)
all_paper_ids = sorted(df["paper_id"].unique().tolist())
selected_papers = st.multiselect(
    "Select which exam papers to include (leave empty = all)",
    options=all_paper_ids,
    default=[],
    key="paper_filter",
)

if selected_papers:
    df_export = df[df["paper_id"].isin(selected_papers)].copy()
else:
    df_export = df.copy()

st.caption(f"Questions to export: **{len(df_export)}**")

generate_btn = st.button(
    "🚀 Generate Word Document",
    type="primary",
    disabled=not DOCX_OK,
)

if generate_btn:
    if not DOCX_OK:
        st.error("python-docx is not installed.")
    else:
        with st.spinner(f"Building Word document for {len(df_export)} questions …"):
            try:
                docx_bytes = build_word_document(df_export, mapping, pdf_bytes)
                st.success("✅ Word document generated successfully!")
                st.download_button(
                    label="⬇ Download Word Document (.docx)",
                    data=docx_bytes,
                    file_name="IB_Chemistry_Questions.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            except Exception as e:
                st.error(f"Error generating document: {e}")
                st.exception(e)

# ════════════════════════════════════════════════════════════════════════════
# SIDEBAR — requirements reminder
# ════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.header("ℹ️ Requirements")
    st.markdown("""
**Python packages:**
```
pip install streamlit pandas openpyxl pymupdf python-docx
```

**Excel columns recognised:**
| Column | Used for |
|--------|---------|
| `Reference` | Paper date string |
| `Question No.` | Q# (composite key) |
| `Page Number` | PDF page to crop |
| `Question Text` | Question body |
| `Mark Scheme` | Answer |
| `Topic` | Metadata |
| `Difficulty` | Metadata |
| `Marks` | Metadata |

**Composite Key Logic:**
> `Reference` + `Q# reset detection` + `Page Number`

Same Reference date can appear in TZ1 and TZ2 — we detect new papers automatically when Q# resets to 1.
""")
