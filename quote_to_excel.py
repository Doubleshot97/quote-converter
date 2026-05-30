"""
quote_to_excel.py
-----------------
Convert a supplier PDF quote into a list of line items, then optionally
into an Excel file. Multiple supplier layouts are supported:

  * Haymans / Cetnaj (similar layouts, columns:
      Line# | Part | Qty | UOM | UnitPrice | Per | GST | LineValue)
  * Ideal Electrical (columns:
      Line# | Part | Description | Qty | UOM | UnitPrice | GST | LineValue)

A dispatcher (`parse_quote_pdf`) detects the supplier from a header
signature in the PDF text and routes to the right parser. New suppliers
can be added by writing another parser and adding a detection rule.

Usage:
    python quote_to_excel.py <input.pdf> [output.xlsx]
"""

import re
import sys
from pathlib import Path

import pdfplumber
from openpyxl import Workbook


# =========================================================================
# SHARED HELPERS
# =========================================================================

def _clean_description(text: str) -> str:
    """Tidy a raw description line.

    - Strips a leading '#' or '# ' marker (used by Haymans on some items).
    - Collapses any double spaces left behind.
    """
    text = text.lstrip()
    if text.startswith("#"):
        text = text[1:].lstrip()
    return re.sub(r"\s{2,}", " ", text).strip()


def _to_number(s: str):
    """Convert a string like '3,287.6500' to a float, stripping commas."""
    return float(s.replace(",", ""))


def _extract_lines(pdf_source):
    """Pull all PDF text and return it as a flat list of lines.

    Accepts a Path/string filename, raw bytes, or a file-like object.

    x_tolerance=2 (default is 3) makes pdfplumber more eager to insert
    spaces between characters, which fixes "ApecCutting" → "Apec Cutting"
    on Haymans, "U/A20MM" → "U/A 20MM" on Ideal, etc.
    """
    import io as _io
    if isinstance(pdf_source, (bytes, bytearray)):
        pdf_source = _io.BytesIO(pdf_source)
    lines = []
    with pdfplumber.open(pdf_source) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=2) or ""
            lines.extend(text.splitlines())
    return lines


# =========================================================================
# HAYMANS / CETNAJ PARSER
# =========================================================================
# Columns: Line# | Part | Qty | UOM | UnitPrice | Per | GST | LineValue
# Each item header is one PDF line; description is on the next line(s).

# Matches a line-item header row, e.g.:
#   "5 FLU1630-2FC 1.000 EA 3,287.6500 1 328.77 3,287.65"
#   "5 #CCGPDM025E-SS 12.000 ea 50.2700 1 60.32 603.24"  (# prefix, lowercase uom)
HAYMANS_ITEM_RE = re.compile(
    r"""^\s*
        (?P<line_no>\d+)\s+              # line number (e.g. 5, 10, 15)
        \#?                              # optional leading '#' marker (stripped)
        (?P<part>[A-Z0-9][A-Z0-9\-/.]*)\s+   # part number (alphanumeric + dashes)
        (?P<qty>[\d,]+\.\d+)\s+          # quantity (e.g. 1.000)
        (?P<uom>[A-Za-z]{1,4})\s+        # unit of measure
        (?P<unit_price>[\d,]+\.\d+)\s+   # unit price
        (?P<per>\d+)\s+                  # per
        (?P<gst>[\d,]+\.\d+)\s+          # GST amount
        (?P<line_value>[\d,]+\.\d+)\s*$  # line value
    """,
    re.VERBOSE,
)

# Lines that end description collection entirely (footer / totals block)
HAYMANS_END_RE = re.compile(
    r"""^(
        Specially\s+ordered    # disclaimer line
        |are\s+non-returnable  # disclaimer line
        |Total\s+(Excl|incl)   # totals
        |GST\b
    )""",
    re.VERBOSE | re.IGNORECASE,
)

# Lines that are warehouse / delivery / packaging metadata — skip but
# keep scanning for the next item.
HAYMANS_SKIP_RE = re.compile(
    r"""^(
        EX\s+\w+                          # "EX SYDNEY", "EX BRISBANE", etc.
        |DELIVERY\b                       # "DELIVERY 1 WEEK"
        |FROM\s+(PACEMENT|PLACEMENT)      # "FROM PACEMENT OF" (supplier typo)
        |AN\s+ORDER\b                     # "AN ORDER"
        |\d+\s*X\s*\d+\s*M\w*\s+          # "1 X 30M COIL", "1 X 30M DRUM",
            (COIL|DRUM|ROLL|REEL|PACK|BOX|LENGTH|LENGTHS)\b
    )""",
    re.VERBOSE | re.IGNORECASE,
)


def _parse_haymans(lines):
    """Parser for Haymans / Cetnaj layout."""
    items = []
    i = 0
    while i < len(lines):
        m = HAYMANS_ITEM_RE.match(lines[i])
        if not m:
            i += 1
            continue

        item = {
            "part": m.group("part"),
            "qty": _to_number(m.group("qty")),
            "uom": m.group("uom").upper(),
            "unit_price": _to_number(m.group("unit_price")),
            "per": int(m.group("per")),
            "description": "",
        }

        description = ""
        j = i + 1
        while j < len(lines):
            nxt = lines[j].strip()
            if not nxt:
                j += 1
                continue
            if HAYMANS_END_RE.match(nxt):
                j = len(lines)
                break
            if HAYMANS_ITEM_RE.match(lines[j]):
                break
            if HAYMANS_SKIP_RE.match(nxt):
                j += 1
                continue
            if not description:
                description = _clean_description(nxt)
            j += 1

        item["description"] = description
        items.append(item)
        i = j

    return items


# =========================================================================
# IDEAL ELECTRICAL PARSER
# =========================================================================
# Columns: Line# | Part | Description | Qty | UOM | UnitPrice | GST | LineValue
# Description sits inline; a freight line "Charges Freight <amount>" may
# follow the items; delivery instructions appear as free-text lines.

# Matches an Ideal line item. Two shapes occur in practice:
#   Shape A (inline description):
#     "1 CCG0564-0-VX A2EX VX BARRIER GLAND U/A 1 EAC 60.0000 6.00 60.00"
#     "1 CBL164CXLPEBK CABLE XLPE 16MM 4C BLACKSHEATH 50 M 17.3100 86.55 865.50"
#   Shape B (no inline description, all on continuation lines):
#     "10 APECBLAPECBLCBL1.510PR# 60 EAC 9.3600 56.16 561.60"
#
# Some rows are prefixed with "* " (Ideal's marker for non-returnable items)
# and some part codes carry trailing '#' or '*' annotation symbols:
#     "* 10 BURB1008S### CHANNEL NUT LONG M10 S/S 140 EAC 2.3000 32.20 322.00"
#     "* 13 BURB2000A*# ALUMINUM STRUT CHANNEL 41X41 6M 2 EAC 104.9200 20.98 209.84"
# We absorb the leading * and the trailing #/* into optional groups and strip
# them so the final output shows clean part codes.
IDEAL_ITEM_RE = re.compile(
    r"""^\s*
        \*?\s*                                       # optional '*' non-return marker
        (?P<line_no>\d+)\s+                          # line number
        (?P<part>[A-Z0-9][A-Z0-9\-/.]*)              # core part number
        [\#\*]*                                      # optional trailing #/* annotations
        \s+
        (?:(?P<description>.+?)\s+)?                 # description (optional)
        (?P<qty>\d+)\s+                              # quantity
        (?P<uom>[A-Za-z]{1,4})\s+                    # UOM (1-4 letters)
        (?P<unit_price>[\d,]+\.\d+)\s+               # unit price
        (?P<gst>[\d,]+\.\d+)\s+                      # GST amount
        (?P<line_value>[\d,]+\.\d+)\s*$              # line value
    """,
    re.VERBOSE,
)

# Matches the freight charge line: "Charges Freight 55.00"
IDEAL_FREIGHT_RE = re.compile(
    r"""^\s*Charges\s+Freight\s+(?P<amount>[\d,]+\.\d+)\s*$""",
    re.VERBOSE | re.IGNORECASE,
)

# The real end-of-quote marker is the totals row: "ORDER TOTAL <amount>" or
# "TOTAL AMOUNT <amount>". On multi-page Ideal quotes the disclaimer block
# repeats at the bottom of every page, so we can't use disclaimer phrases
# as end-markers — only the totals line.
IDEAL_END_RE = re.compile(
    r"""(
        ORDER\s+TOTAL\s+[\d,]+\.\d+
        |TOTAL\s+AMOUNT\s+[\d,]+\.\d+
        |GST\s+AMOUNT\s+[\d,]+\.\d+
    )""",
    re.VERBOSE | re.IGNORECASE,
)

# Disclaimer / page-footer lines to ignore so they don't get appended to the
# previous item's description. (Distinct from end markers, which stop scan.)
IDEAL_FOOTER_RE = re.compile(
    r"""^(
        Important\s+Messages
        |Division\s+of
        |\*\s+Indicates
        |\+\s+Indicates
        |We\s+have\s+updated
        |By\s+accepting
        |which\s+can\s+be
        |we\s+have\s+specifically
        |you\s+agree\s+to
        |documents\s+you
        |immediately\s+if
        |\*\*\s+Due\s+to
        |this\s+quote\s+is
        |regardless\s+of
        |a\s+line-note
    )""",
    re.VERBOSE | re.IGNORECASE,
)

# Free-text delivery/dispatch instructions to discard. These typically
# appear between line items and the freight line. We treat any line
# starting with DELIVER/SHIP/PICK UP/COLLECT as a non-item.
IDEAL_INSTRUCTION_RE = re.compile(
    r"^\s*(DELIVER|SHIP|PICK\s*UP|COLLECT|URGENT|ATTN|PLEASE)\b",
    re.IGNORECASE,
)

# Sales/stock notes between items that aren't part of any description:
#   "NO STOCK - ETA 23/6"
#   "SEE ALT BELOW - EX STOCK"
#   "EVERYTHING IS EX STK - ALT ( 12C ) 1.5 4 PR - ( 16/6 ) - NS"
#   "Most of BUR order MTO - ETA 3 weeks"
#   "**Stocked Option**", "**Stocked option**"
#   "alt offered"
#   "COLD GALVANISING TOUCH UP"  (free-text instructional, varies)
IDEAL_NOTE_RE = re.compile(
    r"""^\s*(
        NO\s+STOCK
        |SEE\s+ALT
        |EVERYTHING\s+IS
        |EX\s+STOCK
        |ETA\s+\d
        |ALT\s+BELOW
        |Most\s+of                  # "Most of BUR order MTO..."
        |\*+\s*Stocked              # "**Stocked Option**"
        |alt\s+offered
        |MTO\s+-\s+ETA
    )""",
    re.VERBOSE | re.IGNORECASE,
)


def _parse_ideal(lines):
    """Parser for Ideal Electrical layout."""
    items = []
    i = 0
    last_item_idx = None  # index in `items` of the most recently added item
    in_page_header = False  # True while traversing repeated page-header text

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Stop scanning once we hit the totals row (could be embedded in a
        # longer line like "Important Messages... ORDER TOTAL 6,811.21").
        if IDEAL_END_RE.search(stripped):
            break

        # Repeated page headers begin with "Ideal Electrical Suppliers" and
        # end with the column-header row that contains "PART NUMBER".
        # While inside a page header, ignore everything (don't append to
        # previous items' descriptions).
        if re.match(r"^\s*Ideal\s+Electrical\s+Suppliers", stripped, re.I):
            in_page_header = True
        if in_page_header:
            # The column-header line is the last line of the page header.
            # After it, the next line is a real item.
            if re.search(r"\bPART\s+NUMBER\b", stripped, re.I):
                in_page_header = False
            i += 1
            continue

        # Try a freight line first — it doesn't match the item pattern
        fm = IDEAL_FREIGHT_RE.match(stripped)
        if fm:
            items.append({
                "part": "",                  # no part number
                "qty": 1,
                "uom": "",                   # no UOM
                "unit_price": _to_number(fm.group("amount")),
                "per": 1,
                "description": "Freight",
            })
            last_item_idx = len(items) - 1
            i += 1
            continue

        # Try a standard line item
        m = IDEAL_ITEM_RE.match(line)
        if m:
            # Description may be absent on the item line (sits on continuation
            # lines below). Default to empty string in that case.
            raw_desc = m.group("description") or ""
            items.append({
                "part": m.group("part"),
                "qty": int(m.group("qty")),
                "uom": m.group("uom").upper(),
                "unit_price": _to_number(m.group("unit_price")),
                "per": 1,
                "description": _clean_description(raw_desc),
            })
            last_item_idx = len(items) - 1
            i += 1
            continue

        # Not an item line. If we have an item in progress, the line might
        # be either (a) a description continuation, or (b) noise to discard.
        if (
            last_item_idx is not None
            and stripped
            and not IDEAL_INSTRUCTION_RE.match(stripped)
            and not IDEAL_NOTE_RE.match(stripped)
            and not IDEAL_FOOTER_RE.match(stripped)
        ):
            # Description continuations are short technical fragments — cable
            # specs like "4C+E 500M 450/750V", "PVC/L/PVC CABLE". Reject lines
            # that look like page headers/footers, column headers, or have
            # dollar amounts.
            is_short = len(stripped) <= 40
            looks_like_continuation = (
                is_short
                and not re.search(r"\$", stripped)
                and not re.search(
                    r"\b(QTY|UOM|PRICE|GST|TOTAL|PAGE|QUOTE|DATE|EXPIRY"
                    r"|SALESPERSON|CONTACT|CUSTOMER|ORIGINAL|EMAIL|TEL"
                    r"|A\.B\.N|TELEPHONE|FAX|SHIP\s+TO|PART\s+NUMBER"
                    r"|LINE\s+VALUE|ITEM\s+DESC|DELIVER|MOLENDINAR|BARNETT"
                    r"|Ideal\s+Electrical|Suppliers|ADDITIONAL\s+INSTR)\b",
                    stripped, re.IGNORECASE)
            )
            if looks_like_continuation:
                # Append to the previous item's description
                prev = items[last_item_idx]
                merged = (prev["description"] + " " + _clean_description(stripped)).strip()
                prev["description"] = re.sub(r"\s{2,}", " ", merged)

        i += 1

    return items


# =========================================================================
# DISPATCHER
# =========================================================================

# A "signature" is a regex that, if found anywhere in the PDF text,
# identifies which supplier it is.
SUPPLIER_SIGNATURES = [
    ("ideal",   re.compile(r"Ideal\s+Electrical|idealelectrical\.com", re.IGNORECASE)),
    ("haymans", re.compile(r"Haymans\s+Electrical|mmem\.com\.au", re.IGNORECASE)),
    ("cetnaj",  re.compile(r"\bCetnaj\b|cetnaj\.com", re.IGNORECASE)),
]


def _detect_supplier(lines):
    """Return a short supplier key based on which signature matches."""
    blob = "\n".join(lines)
    for key, sig in SUPPLIER_SIGNATURES:
        if sig.search(blob):
            return key
    return None


def parse_quote_pdf(pdf_source):
    """Top-level entry point.

    Detects supplier and dispatches to the matching parser. Returns a
    list of dicts with keys: part, description, qty, uom, unit_price, per.

    Accepts a Path/string filename, raw bytes, or a file-like object.

    If no supplier is detected, falls back to the Haymans parser (covers
    Cetnaj and likely other look-alike layouts).
    """
    lines = _extract_lines(pdf_source)

    supplier = _detect_supplier(lines)

    if supplier == "ideal":
        return _parse_ideal(lines)

    # Haymans, Cetnaj, and unknown fallback all use the same parser
    return _parse_haymans(lines)


# =========================================================================
# LEGACY EXCEL WRITER (used by CLI, not by the web app)
# =========================================================================

def write_excel(items, output_path):
    """Write items to xlsx in the two-rows-per-item layout (legacy)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Quote"

    ws.append(["Part Number Description", "Qty", "UOM", "Unit Price", "Per"])
    for item in items:
        ws.append([
            item["part"],
            item["qty"],
            item["uom"],
            item["unit_price"],
            item["per"],
        ])
        ws.append([item["description"], None, None, None, None])

    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 6
    ws.column_dimensions["C"].width = 6
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 6

    wb.save(output_path)


# =========================================================================
# BACK-COMPAT EXPORTS for app.py
# =========================================================================
# The Streamlit app imports these names directly. They originally pointed
# at the only parser (Haymans-style). Now they still work as the Haymans
# parser regexes, but the app should also call parse_quote_pdf for proper
# multi-supplier dispatch.
ITEM_HEADER_RE = HAYMANS_ITEM_RE
END_OF_ITEMS_RE = HAYMANS_END_RE
WAREHOUSE_OR_DELIVERY_RE = HAYMANS_SKIP_RE


def main():
    if len(sys.argv) < 2:
        print("Usage: python quote_to_excel.py <input.pdf> [output.xlsx]")
        sys.exit(1)

    in_path = Path(sys.argv[1])
    out_path = Path(sys.argv[2]) if len(sys.argv) >= 3 else in_path.with_suffix(".xlsx")

    items = parse_quote_pdf(in_path)
    if not items:
        print(f"No line items found in {in_path}.")
        sys.exit(2)

    write_excel(items, out_path)
    print(f"Wrote {len(items)} item(s) to {out_path}")
    for it in items:
        print(f"  - {it['part']:<20} {it['qty']:>6} {it['uom']:<4} "
              f"{it['unit_price']:>10.4f}  {it['description']}")


if __name__ == "__main__":
    main()