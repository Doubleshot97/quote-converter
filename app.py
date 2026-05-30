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

# --- Constants matching the Catalogue Lines template --------------------

CATALOGUE_HEADERS = [
    "Catalogue Part No\n(text 20)",
    "Supplier Part No\n(text 30)",
    "Description\n(*text 100)",
    "Activity Code\n(*text 10)",
    "Specification\n(text 2000)",
    "Favourite\n(integer)",
    "Promotional\n(integer)",
    "Cost Rate\n(number)",
    "Sell Rate\n(number)",
    "Unit\n(text 10)",
    "Negotiated % Disct\n(number)",
    "Negotiated Date\n(date)",
    "Last Invoice No\n(text 10)",
    "Last Invoice Cost\n(number)",
    "Negotiated Date\n(date)",
    "Bill Type\n(text 2)",
    "Group Code\n(text 10)",
    "SubCategory Code\n(text 10)",
    "Category Code\n(text 10)",
    "Inventory Code\n(text 20)",
    "Inactive\n(integer)",
    "Supplier\n(integer)",
    "Currency Code\n(*text 10)",
]
CATALOGUE_COLUMN_WIDTHS = {
    "A": 20.57, "B": 20.71, "C": 58.86, "D": 10.29, "E": 40.71,
    "F": 10.71, "G": 10.71, "H": 10.71, "I": 10.71, "J": 7.71,
    "K": 14.71, "L": 12.71, "M": 11.71, "N": 12.71, "O": 12.71,
    "P": 7.71,  "Q": 9.14,  "R": 13.86, "S": 11.14, "T": 11.57,
    "U": 10.71, "V": 10.71, "W": 11.0,
}
DEFAULT_CURRENCY = "AUD"
SUPPLIER_PART_PREFIX_LEN = 3  # strip first 3 chars of part no for supplier ref

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


def build_catalogue_xlsx(items, activity_code: str) -> bytes:
    """Write items into the Catalogue Lines template layout in memory.

    Columns filled per item:
      A — Catalogue Part No  (raw part number)
      B — Supplier Part No   (part number with first 3 chars stripped)
      C — Description
      D — Activity Code      (optional, from UI)
      H — Cost Rate          (= unit_price / per — true per-unit cost)
      J — Unit               (UOM)
      W — Currency Code      (always 'AUD')

    All other columns are left blank for the user to fill in if needed.
    Items without a part number (e.g. freight rows from Ideal) are skipped
    — the catalogue is for physical stocked items, freight isn't one.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Header row — same styling as the PO file
    ws.append(CATALOGUE_HEADERS)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 45

    # Data rows
    for it in items:
        part = it["part"]
        if not part:
            # Skip freight rows etc. — they don't belong in a parts catalogue
            continue
        supplier_part = part[SUPPLIER_PART_PREFIX_LEN:] if len(part) > SUPPLIER_PART_PREFIX_LEN else ""
        unit_cost = it["unit_price"] / it["per"]
        # 23 columns total (A-W); only A, B, C, D, H, J, W are populated.
        row = [None] * 23
        row[0] = part                       # A — Catalogue Part No
        row[1] = supplier_part or None      # B — Supplier Part No (None if part ≤3 chars)
        row[2] = it["description"] or None  # C — Description
        row[3] = activity_code or None      # D — Activity Code
        row[7] = unit_cost                  # H — Cost Rate
        row[9] = it["uom"] or None          # J — Unit
        row[22] = DEFAULT_CURRENCY          # W — Currency Code
        ws.append(row)

    # Column widths
    for col, width in CATALOGUE_COLUMN_WIDTHS.items():
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

# Build both xlsx files in memory
po_xlsx_bytes = build_po_xlsx(items, job_code.strip(), activity_code.strip())
catalogue_xlsx_bytes = build_catalogue_xlsx(items, activity_code.strip())

# Filenames mirror the input PDF
pdf_stem = Path(uploaded.name).stem
po_out_name = pdf_stem + "-PO.xlsx"
catalogue_out_name = pdf_stem + "-Catalogue.xlsx"

# Two download buttons side-by-side
dl_col1, dl_col2 = st.columns(2)
with dl_col1:
    st.download_button(
        label="⬇️  Purchase Order",
        data=po_xlsx_bytes,
        file_name=po_out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
with dl_col2:
    st.download_button(
        label="⬇️  Catalogue Lines",
        data=catalogue_xlsx_bytes,
        file_name=catalogue_out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.caption(
    f"Work Centre is set to **{DEFAULT_WORK_CENTRE}** for every row in the PO file. "
    "Job Code, Activity Code, and Details can be edited in Excel after download."
)
