"""
app.py — Streamlit web app for converting supplier PDF quotes into a
Purchase Order Lines Excel template. Supports Haymans, Cetnaj, and
Ideal Electrical layouts (auto-detected).
"""

import hashlib
import io
from pathlib import Path

import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from quote_to_excel import parse_quote_pdf


# --- Constants matching the Purchase Order Lines template ----------------

PO_HEADERS = [
    "Job Code\n(*text 20)",
    "Part No on catalogue line\n(text 20)",
    "Activity Code\n(*text 10)",
    "Work Centre\n(*text 10)",
    "Description\n(text 100)",
    "Details\n(text 2000)",
    "Quantity\n(*number)",
    "Unit\n(text 10)",
    "Unit Cost\n(*number)",
]
PO_COLUMN_WIDTHS = {
    "A": 15.14, "B": 19.0, "C": 10.29, "D": 9.86,
    "E": 58.86, "F": 4.0,  "G": 10.71, "H": 7.71, "I": 10.71,
}
HEADER_FILL = PatternFill(start_color="FF305496", end_color="FF305496", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFFFF")
DEFAULT_WORK_CENTRE = "WC004"


# --- Page setup ----------------------------------------------------------

st.set_page_config(
    page_title="Quote PDF → Purchase Order",
    page_icon="📄",
    layout="centered",
)

st.title("📄 Quote PDF → Purchase Order")
st.write(
    "Upload a supplier PDF quote. "
    "The app extracts line items and fills the Purchase Order Lines template."
)


# --- Cached PDF parsing --------------------------------------------------

@st.cache_data(show_spinner=False)
def parse_pdf_bytes(pdf_bytes: bytes, _file_hash: str):
    """Parse a supplier quote PDF from raw bytes.

    Detection of supplier and routing to the right parser happens in
    quote_to_excel.parse_quote_pdf. We just pass the bytes through.
    """
    return parse_quote_pdf(pdf_bytes)


def build_po_xlsx(items, job_code: str, activity_code: str) -> bytes:
    """Write items into the Purchase Order Lines template layout in memory."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Header row — matches the template's blue fill + white bold font
    ws.append(PO_HEADERS)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 45

    # Data rows
    for it in items:
        # Haymans quotes the price as "Unit Price per N units" (e.g. $479 per
        # 100m). Divide by 'per' so the value written into the template is the
        # true per-single-unit cost. For the vast majority of rows per == 1,
        # so this is a no-op.
        unit_cost = it["unit_price"] / it["per"]
        ws.append([
            job_code or None,                # A — Job Code
            it["part"] or None,              # B — Part No (None for freight)
            activity_code or None,           # C — Activity Code
            DEFAULT_WORK_CENTRE,             # D — Work Centre
            it["description"],               # E — Description
            None,                            # F — Details (left blank)
            it["qty"],                       # G — Quantity
            it["uom"] or None,               # H — Unit (None for freight)
            unit_cost,                       # I — Unit Cost
        ])

    # Column widths from the template
    for col, width in PO_COLUMN_WIDTHS.items():
        ws.column_dimensions[col].width = width

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --- UI ------------------------------------------------------------------

col1, col2 = st.columns(2)
with col1:
    job_code = st.text_input(
        "Job Code (optional)",
        value="",
        max_chars=20,
        help="Leave blank if you'd rather fill it in Excel.",
    )
with col2:
    activity_code = st.text_input(
        "Activity Code (optional)",
        value="",
        max_chars=10,
        help="Leave blank if you'd rather fill it in Excel.",
    )

uploaded = st.file_uploader(
    "Choose a quote PDF",
    type=["pdf"],
    accept_multiple_files=False,
)

if uploaded is None:
    st.info("👆 Drag a PDF here or click to browse.")
    st.stop()


pdf_bytes = uploaded.getvalue()
file_hash = hashlib.sha1(pdf_bytes).hexdigest()

with st.spinner("Reading PDF…"):
    try:
        items = parse_pdf_bytes(pdf_bytes, file_hash)
    except Exception as exc:
        st.error(f"Couldn't parse that PDF: {exc}")
        st.stop()

if not items:
    st.warning(
        "No line items found. "
        "If this is from a new supplier with a different layout, "
        "the parser may need new rules."
    )
    st.stop()


st.success(f"Found **{len(items)}** line item(s).")

# Subtotal — lets the user cross-check against the PDF's 'Total Excl GST'
subtotal = sum(it["qty"] * it["unit_price"] / it["per"] for it in items)
col_a, col_b = st.columns(2)
col_a.metric("Line items", len(items))
col_b.metric("Subtotal (excl GST)", f"${subtotal:,.2f}")
st.caption(
    "💡 Compare the subtotal above with the **Total Excl GST** at the bottom "
    "of the PDF to confirm every line was captured correctly."
)

# Show a preview of what'll go into the template
st.dataframe(
    [
        {
            "Job Code": job_code or "",
            "Part No": it["part"],
            "Activity Code": activity_code or "",
            "Work Centre": DEFAULT_WORK_CENTRE,
            "Description": it["description"],
            "Quantity": it["qty"],
            "Unit": it["uom"],
            "Unit Cost": it["unit_price"] / it["per"],
        }
        for it in items
    ],
    use_container_width=True,
    hide_index=True,
)

# Build PO xlsx in memory
xlsx_bytes = build_po_xlsx(items, job_code.strip(), activity_code.strip())

# Filename: <original-pdf-name>-PO.xlsx
out_name = Path(uploaded.name).stem + "-PO.xlsx"

st.download_button(
    label="⬇️  Download Purchase Order Excel",
    data=xlsx_bytes,
    file_name=out_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

st.caption(
    f"Work Centre is set to **{DEFAULT_WORK_CENTRE}** for every row. "
    "Job Code, Activity Code, and Details can be edited in Excel after download."
)