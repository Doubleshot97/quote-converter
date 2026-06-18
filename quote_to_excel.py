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
        (?P<part>[A-Z0-9][A-Z0-9\-/.+]*)\s+  # part number (alphanumeric, -, /, ., +)
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

# Haymans freight charge row, e.g.:  "Freight 2.50 25.00"
# Two numbers: GST amount, line value (= the freight cost).
HAYMANS_FREIGHT_RE = re.compile(
    r"""^\s*Freight\s+
        (?P<gst>[\d,]+\.\d+)\s+
        (?P<line_value>[\d,]+\.\d+)\s*$""",
    re.VERBOSE | re.IGNORECASE,
)


def _parse_haymans(lines):
    """Parser for Haymans / Cetnaj layout."""
    items = []
    i = 0
    while i < len(lines):
        # Haymans freight rows appear as a standalone line: "Freight 2.50 25.00"
        # No part number, no qty/UOM column — just GST + line value. Add as a
        # freight item with no Part No so the catalogue output skips it but
        # the PO output includes it.
        fm = HAYMANS_FREIGHT_RE.match(lines[i])
        if fm:
            items.append({
                "part": "",
                "qty": 1,
                "uom": "",
                "unit_price": _to_number(fm.group("line_value")),
                "per": 1,
                "description": "Freight",
            })
            i += 1
            continue

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
            if HAYMANS_FREIGHT_RE.match(lines[j]):
                # Stop here so the outer loop captures the freight row.
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
# PROCESS SYSTEMS / valvesonline.com.au PARSER
# =========================================================================
# Columns: QTY | Code | Name | Unit Price | Subtotal
#   Prices are prefixed with '$'.
# Multi-line names continue on plain lines below; we keep only the FIRST
# line of each description (per user spec — Description field has a char
# limit, technical detail lines aren't wanted).
# Freight appears in the totals area as "Star Track Road: $30.00" (or a
# similar courier label). Captured as a freight item with no part number.

PS_ITEM_RE = re.compile(
    r"""^\s*
        (?P<qty>\d+)\s+                          # quantity (whole number)
        (?P<part>[A-Z][A-Z0-9\-/.]*)\s+          # part code (starts with letter)
        (?P<description>.+?)\s+                  # description
        \$(?P<unit_price>[\d,]+\.\d+)\s+         # unit price ($-prefixed)
        \$(?P<line_value>[\d,]+\.\d+)\s*$        # line value ($-prefixed)
    """,
    re.VERBOSE,
)

# Freight row in the totals area, e.g. "Star Track Road: $30.00".
PS_FREIGHT_RE = re.compile(
    r"""^\s*(?P<label>[A-Za-z][A-Za-z\s]+?):\s*\$(?P<amount>[\d,]+\.\d+)\s*$""",
    re.VERBOSE,
)

# Labels that look like the freight pattern but are actually totals — skip
# them (don't capture as freight, just stop processing).
PS_TOTALS_RE = re.compile(
    r"""^\s*(Product\s+Subtotal|Sub\s*Total|GST|Grand\s+Total|Total)\b""",
    re.VERBOSE | re.IGNORECASE,
)


def _parse_process_systems(lines):
    """Parser for Process Systems / valvesonline.com.au layout."""
    items = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # Stop on the totals block.
        if PS_TOTALS_RE.match(stripped):
            break

        # Regular line item
        m = PS_ITEM_RE.match(line)
        if m:
            items.append({
                "part": m.group("part"),
                "qty": int(m.group("qty")),
                "uom": "EA",   # Process Systems quotes don't have a UOM column
                "unit_price": _to_number(m.group("unit_price")),
                "per": 1,
                "description": _clean_description(m.group("description")),
            })
            continue

        # Freight (courier-named row), e.g. "Star Track Road: $30.00"
        fm = PS_FREIGHT_RE.match(stripped)
        if fm and not PS_TOTALS_RE.match(stripped):
            items.append({
                "part": "",
                "qty": 1,
                "uom": "",
                "unit_price": _to_number(fm.group("amount")),
                "per": 1,
                "description": "Freight",
            })
            continue

        # Continuation lines (multi-line Names) and all other noise are
        # silently ignored — we only keep the first line of each description.

    return items


# =========================================================================
# APS INDUSTRIAL PARSER
# =========================================================================
# Columns: Line | Product Code | Description | Quantity | Unit Price | Total Price
# Product codes are 7-10 digit numbers (e.g. 3239724, 1991920000).
# Quantity always followed by " EA " (always EA for APS).
# Multi-line descriptions wrap onto the NEXT line(s) AFTER the price row,
# before the "Plant:" / "Confirmed Quantity:" / "Date Available:" metadata.
# Section headers like "24VDC DISTRIBUTION GROUP 1 FUSE/LINK TER" appear
# between items without a line number — they don't match the item regex
# so they're naturally skipped.

APS_ITEM_RE = re.compile(
    r"""^\s*
        (?P<line_no>\d+)\s+                  # line number
        (?P<part>\d{7,10})\s+                # APS product code (7-10 digits)
        (?P<desc1>.+?)\s+                    # description part 1
        (?P<qty>\d+)\s+EA\s+                 # quantity (always EA)
        (?P<unit_price>[\d,]+\.\d+)\s+       # unit price
        (?P<total>[\d,]+\.\d+)\s*$           # total price
    """,
    re.VERBOSE,
)

# Lines that end description continuation — metadata, page footers, headers
APS_END_CONTINUATION_RE = re.compile(
    r"""^(
        Plant:
        |Confirmed\s+Quantity
        |Unconfirmed\s+Quantity
        |Date\s+Available
        |Sales\s+Note
        |Subtotal\b
        |GST\s
        |Total\s
        |APS\s+Industrial
        |VIC/TAS|SA:|NSW:|WA:|QLD:
        |Quotation:
        |Page:
        |Line\s+Product\s+Code
        |Direct\s+Deposit
        |Please\s+quote
        |All\s+Sales
        |\*Available
    )""",
    re.VERBOSE | re.IGNORECASE,
)


def _parse_aps(lines):
    """Parser for APS Industrial layout."""
    items = []
    i = 0
    while i < len(lines):
        m = APS_ITEM_RE.match(lines[i])
        if not m:
            i += 1
            continue

        desc_parts = [m.group("desc1").strip()]

        # Description may wrap onto subsequent lines BEFORE metadata starts.
        j = i + 1
        while j < len(lines):
            nxt = lines[j].strip()
            if not nxt:
                j += 1
                continue
            if APS_END_CONTINUATION_RE.match(nxt):
                break
            if APS_ITEM_RE.match(lines[j]):
                break
            desc_parts.append(nxt)
            j += 1

        # Join, normalise whitespace, then fix word-break artifacts where the
        # PDF extractor put a space after a hyphen: "Push- In" → "Push-In"
        description = " ".join(desc_parts)
        description = re.sub(r"\s+", " ", description).strip()
        description = re.sub(r"([A-Za-z0-9])-\s+([A-Z])", r"\1-\2", description)

        items.append({
            "part": m.group("part"),
            "qty": int(m.group("qty")),
            "uom": "EA",
            "unit_price": _to_number(m.group("unit_price")),
            "per": 1,
            "description": _clean_description(description),
        })
        i = j

    return items


# =========================================================================
# IPD GROUP PARSER
# =========================================================================
# Columns: ITEM# | DESCRIPTION | QTY | UNIT PRICE | TOTAL
# No line numbers. Items start with a part code (e.g. "SFD1-20-50-275-A",
# "QS20.241", "CN-FF-90-3"). Prices are $-prefixed.
# Multi-line descriptions wrap AFTER the price line, like APS.
# Footer markers: "IN STOCK SUBJECT TO SALE.", "Grand Total - ex GST".

# Part codes can be split by pdfplumber on hyphen boundaries (e.g.
# "CN-FF-90 -3"). Pre-fix at the start of each line, then match.
_IPD_PART_BREAK_RE = re.compile(
    r"^(\s*[A-Z0-9]+(?:[\-./][A-Z0-9]+)*)\s+-(\d+|[A-Z]+)(\s)"
)


def _fix_ipd_part_breaks(line):
    """Glue back together part codes that pdfplumber split on hyphen.

    Conservative: only fires at the start of the line, only when the
    candidate token is followed by `-<digits/letters><space>`. Won't
    affect mid-line text like "100 -240 VAC".
    """
    prev = None
    while prev != line:
        prev = line
        line = _IPD_PART_BREAK_RE.sub(r"\1-\2\3", line)
    return line


IPD_ITEM_RE = re.compile(
    r"""^\s*
        (?P<part>[A-Z][A-Z0-9\-./]*)\s+       # part code (letters/digits/-/./)
        (?P<desc>.+?)\s+                      # description (lazy)
        (?P<qty>\d+)\s+                       # quantity
        \$(?P<unit_price>[\d,]+\.\d+)\s+      # unit price $X.XX
        \$(?P<total>[\d,]+\.\d+)\s*$          # total $X.XX
    """,
    re.VERBOSE,
)

# Lines that end description continuation — footers/page elements/notes
IPD_END_CONTINUATION_RE = re.compile(
    r"""^(
        IN\s+STOCK
        |Grand\s+Total
        |Page:
        |Copyright
        |All\s+rights
        |Power\s+Distribution
        |Safety\s+&\s+Hazardous
        |Quote\s+No:
        |Terms\s+and\s+Conditions
        |\d+\.\s+(Pricing|Delivery|Validity|Trading)
    )""",
    re.VERBOSE | re.IGNORECASE,
)


def _parse_ipd(lines):
    """Parser for IPD Group layout."""
    items = []
    i = 0
    while i < len(lines):
        line = _fix_ipd_part_breaks(lines[i])
        m = IPD_ITEM_RE.match(line)
        if not m:
            i += 1
            continue

        desc_parts = [m.group("desc").strip()]

        # Look ahead for description continuation (wraps AFTER the price line).
        # Stops on next item, footer markers, or blank.
        j = i + 1
        while j < len(lines):
            nxt = lines[j].strip()
            if not nxt:
                j += 1
                continue
            if IPD_END_CONTINUATION_RE.match(nxt):
                break
            # Try matching as a new item line (with the part-break fix applied)
            if IPD_ITEM_RE.match(_fix_ipd_part_breaks(lines[j])):
                break
            desc_parts.append(nxt)
            j += 1

        description = " ".join(desc_parts)
        description = re.sub(r"\s+", " ", description).strip()

        items.append({
            "part": m.group("part"),
            "qty": int(m.group("qty")),
            "uom": "EA",   # IPD quotes don't show UOM — assume each
            "unit_price": _to_number(m.group("unit_price")),
            "per": 1,
            "description": _clean_description(description),
        })
        i = j

    return items


# =========================================================================
# PHOENIX CONTACT PARSER
# =========================================================================
# Columns: Item | Material | Description | Quantity QU* | Net Price | PU* | Amount
# Material codes are 6-8 digits (e.g. 1348516, 3044076, 800886).
# Unit is "PCE" — always.
# Per (PU*) is captured to support quotes where it might not be 1, though
# it's always 1 in the sample.
# Description is taken from the first line only — the long technical text
# below (Feed-through terminal block, nom. voltage: ...) is excluded per
# user spec.
# Each item has lots of metadata after it: a repeated material code on its
# own line, "Customer material:", "Sales unit:", "Min. order quant:",
# "Partial packaging:", "Delivery time:" — all of which must be skipped.

PC_ITEM_RE = re.compile(
    r"""^\s*
        (?P<line_no>\d+)\s+                  # line number (10, 20, 30, ...)
        (?P<part>\d{6,8})\s+                 # material code (6-8 digits)
        (?P<desc>.+?)\s+                     # description (lazy)
        (?P<qty>\d+)\s+                      # quantity
        (?P<unit>[A-Z]{2,4})\s+              # unit (PCE)
        (?P<unit_price>[\d,]+\.\d+)\s+       # unit price (Net Price)
        (?P<per>\d+)\s+                      # PU* (Price Unit)
        (?P<amount>[\d,]+\.\d+)\s*$          # amount
    """,
    re.VERBOSE,
)

# Phoenix Contact freight appears in the footer with European number
# formatting (comma is the decimal separator):
#   "Freight costs 20,00"
# Treated as a generic freight item — no part number, "Freight" description,
# qty 1, $20.00 unit cost.
PC_FREIGHT_RE = re.compile(
    r"""^\s*Freight\s+costs\s+(?P<amount>[\d.]+,\d{2})\s*$""",
    re.VERBOSE | re.IGNORECASE,
)


def _european_to_number(s: str) -> float:
    """Convert '4.291,87' or '20,00' (European format) to a float.

    Dots are thousand separators, commas are decimal separators.
    """
    # Remove the thousand separators (dots), replace the decimal comma.
    return float(s.replace(".", "").replace(",", "."))


def _parse_phoenix_contact(lines):
    """Parser for Phoenix Contact layout."""
    items = []
    for line in lines:
        m = PC_ITEM_RE.match(line)
        if m:
            items.append({
                "part": m.group("part"),
                "qty": int(m.group("qty")),
                "uom": m.group("unit"),  # always 'PCE'
                "unit_price": _to_number(m.group("unit_price")),
                "per": int(m.group("per")),  # supports per != 1 in future quotes
                "description": _clean_description(m.group("desc")),
            })
            continue

        # Check for the freight row in the totals footer
        fm = PC_FREIGHT_RE.match(line.strip())
        if fm:
            items.append({
                "part": "",
                "qty": 1,
                "uom": "",
                "unit_price": _european_to_number(fm.group("amount")),
                "per": 1,
                "description": "Freight",
            })
    return items


# =========================================================================
# MECHTRIC PARSER
# =========================================================================
# Columns: Qty | Item | Description | UOM | Net Price | Total
# Item codes are uppercase alphanumeric with hyphens/dots (e.g.
# SFD1-20-50-275-A, SL36-G, CN-FF-90-3, DRL-24V480W1EN).
# UOM is lowercase "ea" in the PDF — normalised to "EA".
# Prices are $-prefixed. No "Per" concept.
#
# Description wraps onto the line below the price-bearing line, like APS
# and IPD. Stock-availability notes ("ex brisbane", "16 ex sydney",
# "6 ex factory 7-10 days") appear between items and should be discarded.
#
# Freight is "<qty> Freight Out Freight Out <uom> $<price> $<total>" —
# captured as a freight item with no part number.

# Item header line
MECHTRIC_ITEM_RE = re.compile(
    r"""^\s*
        (?P<qty>\d+)\s+                          # quantity
        (?P<part>[A-Z][A-Za-z0-9\-./]*)\s+       # part code
        (?P<desc>.+?)\s+                         # description fragment (lazy)
        (?P<unit>[a-z]{2,4})\s+                  # unit (lowercase in PDF)
        \$(?P<unit_price>[\d,]+\.\d+)\s+         # unit price
        \$(?P<total>[\d,]+\.\d+)\s*$             # total
    """,
    re.VERBOSE,
)

# Freight row — match BEFORE the general item regex so "Freight Out" doesn't
# get parsed as a part code.
MECHTRIC_FREIGHT_RE = re.compile(
    r"""^\s*
        (?P<qty>\d+)\s+
        Freight\s+Out\s+Freight\s+Out\s+
        (?P<unit>[a-z]{2,4})\s+
        \$(?P<unit_price>[\d,]+\.\d+)\s+
        \$(?P<total>[\d,]+\.\d+)\s*$
    """,
    re.VERBOSE | re.IGNORECASE,
)

# Stock-availability notes to filter out:
#   "ex brisbane", "ex sydney"
#   "16 ex sydney", "6 ex factory 7-10 days"
MECHTRIC_STOCK_NOTE_RE = re.compile(
    r"^\s*(\d+\s+)?ex\s+\w+",
    re.IGNORECASE,
)

# Lines that end an item's description block (footer / totals)
MECHTRIC_END_RE = re.compile(
    r"""^\s*(
        Please\s+Note
        |Bank\s+Account
        |Subtotal
        |GST\s+Total
        |\bTotal\b\s+\$
        |Credit\s+Card
        |Cheques:
        |Enquiries:
        |ANZ\s+EFT
        |Please\s+use\s+Invoice
    )""",
    re.VERBOSE | re.IGNORECASE,
)


def _parse_mechtric(lines):
    """Parser for Mechtric layout."""
    items = []
    i = 0
    while i < len(lines):
        line = lines[i]

        # Stop on footer/totals block
        if MECHTRIC_END_RE.match(line):
            break

        # Try freight first (more specific pattern)
        fm = MECHTRIC_FREIGHT_RE.match(line)
        if fm:
            items.append({
                "part": "",
                "qty": int(fm.group("qty")),
                "uom": "",
                "unit_price": _to_number(fm.group("unit_price")),
                "per": 1,
                "description": "Freight",
            })
            i += 1
            continue

        # Regular item line
        m = MECHTRIC_ITEM_RE.match(line)
        if not m:
            i += 1
            continue

        desc_parts = [m.group("desc").strip()]

        # Look ahead for description continuation lines, until next item,
        # freight, footer, or another stock note we can't recognise as such.
        j = i + 1
        while j < len(lines):
            nxt = lines[j].strip()
            if not nxt:
                j += 1
                continue
            if MECHTRIC_END_RE.match(nxt):
                break
            if MECHTRIC_FREIGHT_RE.match(lines[j]):
                break
            if MECHTRIC_ITEM_RE.match(lines[j]):
                break
            if MECHTRIC_STOCK_NOTE_RE.match(nxt):
                # Stock note — skip, but keep scanning for more continuation
                j += 1
                continue
            desc_parts.append(nxt)
            j += 1

        description = " ".join(desc_parts)
        description = re.sub(r"\s+", " ", description).strip()

        items.append({
            "part": m.group("part"),
            "qty": int(m.group("qty")),
            "uom": m.group("unit").upper(),  # 'ea' → 'EA'
            "unit_price": _to_number(m.group("unit_price")),
            "per": 1,
            "description": _clean_description(description),
        })
        i = j

    return items


# =========================================================================
# NHP ELECTRICAL ENGINEERING PARSER
# =========================================================================
# Columns: Item | Product | Description | Qty | Net (Ea) | Ext Net
# Item numbers have a decimal format (1.00, 3.50, 6.50, 6.75, ...).
# Each item fits on one line.
# Some parts have a "†" non-returnable marker after the part code.
# Prices can have comma thousand separators (e.g. 2,423.05).
#
# NHP displays the unit price rounded to 2 decimals, but the underlying
# value can have more precision (observed: 22 × $16.71 = $367.62, but the
# PDF says $367.54 for the line — so the real unit price is ~16.7064).
# To make every line reconcile to the PDF's Ext Net column, the true unit
# price is computed as line_total / qty rather than reading the displayed
# unit price directly. This guarantees the subtotal matches the PDF exactly.

NHP_ITEM_RE = re.compile(
    r"""^\s*
        (?P<item_no>\d+\.\d{2})\s+               # item number (1.00, 3.50, etc.)
        (?P<part>[A-Z0-9][A-Z0-9]*)              # part code (may start with digit
                                                 #   e.g. 9404, 5534007424VDC)
        (?:\s+†)?                                # optional non-returnable marker
        \s+
        (?P<desc>.+?)\s+                         # description (lazy)
        (?P<qty>\d+)\s+EA\s+                     # quantity (always EA)
        (?P<unit_price>[\d,]+\.\d+)\s+           # displayed unit price
        \$(?P<line_total>[\d,]+\.\d+)\s*$        # Ext Net (authoritative)
    """,
    re.VERBOSE,
)


def _clean_nhp_watermark(line: str) -> str:
    """Strip stray single-letter watermark fragments from a line.

    NHP "DRAFT" watermarked PDFs leak single letters (D, R, A, F, T) into
    extracted text — usually between description words and the qty. Pattern
    observed: "... C CURVE T1 EA ..." should be "... C CURVE 1 EA ...".

    We only strip a single uppercase letter that appears immediately before
    a digit followed by " EA " — this is conservative enough not to corrupt
    legitimate descriptions like "1NO" or "C/O".
    """
    return re.sub(r"\s[DRAFT](?=\d+\s+EA\s+)", " ", line)


def _parse_nhp(lines):
    """Parser for NHP Electrical Engineering layout."""
    items = []
    for line in lines:
        # Pre-clean stray DRAFT watermark fragments
        cleaned = _clean_nhp_watermark(line)
        m = NHP_ITEM_RE.match(cleaned)
        if not m:
            continue

        qty = int(m.group("qty"))
        # Use Ext Net / qty to recover the true unit price — the displayed
        # unit price is rounded for presentation and can be a cent or two
        # off from the underlying value.
        line_total = _to_number(m.group("line_total"))
        true_unit_price = line_total / qty if qty else 0.0

        items.append({
            "part": m.group("part"),
            "qty": qty,
            "uom": "EA",
            "unit_price": true_unit_price,
            "per": 1,
            "description": _clean_description(m.group("desc")),
        })
    return items


# =========================================================================
# DORE ELECTRICS PARSER
# =========================================================================
# Columns: Stock Code | Description | Quantity | Unit | Price | Disc% | Total Ex GST
# Stock codes can start with letters OR digits (e.g. SM202, 165E24, E/NFEETBLK).
# Unit is "PIEC" or "PCE" — both normalised to "EA".
# Most lines have a Disc% column; a few don't (e.g. when there's no discount).
# The displayed Price is BEFORE discount — true unit cost = line_total / qty.
# This matches what the user prefers and avoids any rounding drift from the
# Price × (1 - Disc%) math.
#
# "C/P F/O" rows (Credit Pickup / Freight Out admin note) appear with no qty
# or price columns and are correctly ignored by the regex.

# Item with discount: <part> <desc> <qty> <unit> <price> <disc> <line_total>
DORE_ITEM_DISC_RE = re.compile(
    r"""^\s*
        (?P<part>[A-Z0-9][A-Z0-9/]*)\s+          # stock code (may start with digit)
        (?P<desc>.+?)\s+                          # description (lazy)
        (?P<qty>\d+\.\d+)\s+                      # quantity (e.g. 2.00)
        (?P<unit>P[IC]E[C]?)\s+                   # PIEC or PCE
        (?P<price>[\d,]+\.\d+)\s+                 # displayed price (pre-disc)
        (?P<disc>\d+\.\d+)\s+                     # discount %
        (?P<line_total>[\d,]+\.\d+)\s*$           # total ex GST
    """,
    re.VERBOSE,
)

# Item without discount column — same as above minus the disc field
DORE_ITEM_NODISC_RE = re.compile(
    r"""^\s*
        (?P<part>[A-Z0-9][A-Z0-9/]*)\s+
        (?P<desc>.+?)\s+
        (?P<qty>\d+\.\d+)\s+
        (?P<unit>P[IC]E[C]?)\s+
        (?P<price>[\d,]+\.\d+)\s+
        (?P<line_total>[\d,]+\.\d+)\s*$
    """,
    re.VERBOSE,
)


def _parse_dore(lines):
    """Parser for Dore Electrics layout."""
    items = []
    for line in lines:
        # Try with-discount first (more specific). If it doesn't match,
        # fall back to the no-discount form. The order matters because the
        # no-discount regex would also match a discount line — leaving the
        # disc% inside the description.
        m = DORE_ITEM_DISC_RE.match(line) or DORE_ITEM_NODISC_RE.match(line)
        if not m:
            continue

        qty = float(m.group("qty"))
        line_total = _to_number(m.group("line_total"))
        # True unit cost = line total / qty. Avoids rounding drift from the
        # displayed Price × (1 - Disc%) math, and matches the convention used
        # for NHP (where displayed unit prices were also rounded).
        true_unit_price = line_total / qty if qty else 0.0

        items.append({
            "part": m.group("part"),
            "qty": int(qty) if qty == int(qty) else qty,
            "uom": "EA",  # normalise PIEC/PCE → EA
            "unit_price": true_unit_price,
            "per": 1,
            "description": _clean_description(m.group("desc")),
        })
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
    ("process_systems",
                re.compile(r"Process\s+Systems|valvesonline\.com", re.IGNORECASE)),
    ("aps",     re.compile(r"APS\s+Industrial|apsindustrial\.com", re.IGNORECASE)),
    ("ipd",     re.compile(r"IPD\s+Group|www\.ipd\.com\.au", re.IGNORECASE)),
    ("phoenix_contact",
                re.compile(r"Phoenix\s+Contact|phoenixcontact\.com", re.IGNORECASE)),
    ("mechtric",
                re.compile(r"Mechtric\s+Pty|mechtric\.com", re.IGNORECASE)),
    ("nhp",     re.compile(r"NHP\s+Electrical|nhp\.com\.au", re.IGNORECASE)),
    ("dore",    re.compile(r"Dore\s+Electrics|doreelectrics\.com", re.IGNORECASE)),
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
    list of dicts with keys: part, description, qty, uom, unit_price, per,
    supplier. The 'supplier' key is one of: 'haymans', 'cetnaj', 'ideal',
    or None when the supplier couldn't be identified (unknown layout, but
    parsed via the Haymans fallback).

    Accepts a Path/string filename, raw bytes, or a file-like object.

    If no supplier is detected, falls back to the Haymans parser (covers
    Cetnaj and likely other look-alike layouts).
    """
    lines = _extract_lines(pdf_source)

    supplier = _detect_supplier(lines)

    if supplier == "ideal":
        items = _parse_ideal(lines)
    elif supplier == "process_systems":
        items = _parse_process_systems(lines)
    elif supplier == "aps":
        items = _parse_aps(lines)
    elif supplier == "ipd":
        items = _parse_ipd(lines)
    elif supplier == "phoenix_contact":
        items = _parse_phoenix_contact(lines)
    elif supplier == "mechtric":
        items = _parse_mechtric(lines)
    elif supplier == "nhp":
        items = _parse_nhp(lines)
    elif supplier == "dore":
        items = _parse_dore(lines)
    else:
        # Haymans, Cetnaj, and unknown fallback all use the same parser.
        items = _parse_haymans(lines)

    # Tag each item with the detected supplier so downstream code (the
    # catalogue builder) can apply supplier-specific rules.
    for it in items:
        it["supplier"] = supplier

    return items


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
