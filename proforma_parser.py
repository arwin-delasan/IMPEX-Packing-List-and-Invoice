"""Extract per-product-code Unit Price from a SAWO Pro-Forma Invoice PDF.

Two distinct Pro-Forma templates exist in practice, and this module
supports both, picking between them by what it actually finds in the file
rather than asking the caller:

1. *Odoo-generated* - each item line's Description column starts with the
   product code in brackets, e.g. "[SCA-90NS-P-C] Scandia 3H.E 9,0kW NS
   Premium 3P-1P". Fee lines (Bank Charge, Freight Costs, Export
   Processing Fee, ...) have no bracketed code and are skipped - they
   aren't shipped product, so they never appear on the packing list this
   price gets matched against.

2. *Bare-code* (e.g. "proforma invoice.pdf", SPACORP FZCO / Jebel Ali) -
   the product code is its own left-most column with no brackets, the
   columns are PRODUCT CODE / DESCRIPTION / QTY / UM / UNIT PRICE / TOTAL
   PRICE (no CBM or Weight columns at all), and prices are in EUR rather
   than USD. Verified against that file: 106 item rows, every one of them
   left-aligned at the same x0 and carrying exactly two decimal values.

Only Unit Price is extracted (that's all generate_inv.py needs). In the
bare-code template it's taken as the *leftmost* decimal on the row rather
than by an x-position anchor: the euro glyph renders as its own word and
splits the Total Price column mid-number (e.g. "€ 6 61.12" for 661.12,
"€ 1 ,190.40" for 1190.40), so an x-anchor near the right of the row can
land on a fragment. Column order puts Unit Price before Total Price, and
Quantity is an integer that never matches a two-decimal pattern, so
"leftmost decimal" is exact. Checked against all 106 rows of that file:
unit price x quantity reproduces the printed line total on every row, and
the leftmost-decimal pick agrees with an x-anchored pick on every row.
"""

import re

import pdfplumber

from pdf_parser import ExtractionError, cluster_rows, fail, parse_number

BRACKETED_CODE_RE = re.compile(r"^\[(.+)\]$")
DECIMAL_RE = re.compile(r"-?[\d,]+\.\d+")

# Bare-code template: a two-decimal money value (8.98, 1,190.40). Stricter
# than DECIMAL_RE on purpose - it must not match a measurement embedded in
# a description, e.g. the "0.6" of "Cozy Tank Humidifier Cylindrical 0.6L".
MONEY_RE = re.compile(r"-?[\d,]*\.\d\d")

# A bare product code: starts alphanumeric, then alphanumerics and the
# separators SAWO codes actually use - "221-THD", "730-4SGD-M",
# "LP15-001-EU", "STN-45-1/2-DFP-X", "STP-BTN-2.0", "TRD-90/120NS-G-P".
BARE_CODE_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9./-]*$")

# Item rows put the code in the left-most column; the Description column
# starts well to the right of it (x0 138 in the reference file, vs 31 for
# the codes), as do the page's own header/footer lines (x0 54). Anything
# starting right of this isn't an item row's first token.
CODE_COLUMN_MAX_X = 100

# x0 anchors (points from the page's left edge) for the *Odoo* template's
# numeric columns - same nearest-anchor approach as
# pdf_parser.GRAND_TOTAL_ANCHORS. Only used to tell Unit Price apart from
# CBM/Weight/Amount on the same row; Quantity is excluded already since
# it's an integer (no decimal point) and never matches DECIMAL_RE. Not
# used by the bare-code template, which has no CBM/Weight columns at all.
COLUMN_ANCHORS = {
    "cbm": 365,
    "weight": 410,
    "unit_price": 450,
    "amount": 520,
}


def detect_currency(pages_text):
    """Return the ISO code for whichever currency symbol the document
    actually uses. generate_inv.py prints this on the Grand Total row, so
    a EUR-priced proforma isn't silently labelled USD."""
    if "\u20ac" in pages_text:
        return "EUR"
    return "USD"


def _rows(page):
    # Not use_text_flow: in the Odoo template this PDF's content stream
    # interleaves a wrapped description's continuation line *between* the
    # first line and that row's numeric columns, which defeats
    # cluster_rows' sequential top-proximity grouping. Default
    # extract_words sorts by (top, x0) instead, which reconstructs correct
    # visual rows regardless of stream order.
    return cluster_rows(page.extract_words(keep_blank_chars=False, x_tolerance=1.5))


def _parse_bracketed(pages):
    """Odoo template: code is bracketed inside the Description column."""
    prices = {}
    for rows in pages:
        for row in rows:
            row_words = row["words"]
            m = BRACKETED_CODE_RE.match(row_words[0]["text"])
            if not m:
                continue
            decimal_words = [w for w in row_words if DECIMAL_RE.fullmatch(w["text"])]
            if not decimal_words:
                continue
            nearest = min(decimal_words, key=lambda w: abs(w["x0"] - COLUMN_ANCHORS["unit_price"]))
            prices[m.group(1)] = parse_number(nearest["text"])
    return prices


def _parse_bare_code(pages):
    """Bare-code template: code is its own left-most column, Unit Price is
    the leftmost two-decimal value on the row (see module docstring)."""
    prices = {}
    for rows in pages:
        for row in rows:
            row_words = row["words"]
            first = row_words[0]
            if first["x0"] > CODE_COLUMN_MAX_X or not BARE_CODE_RE.match(first["text"]):
                continue
            money = [w for w in row_words if MONEY_RE.fullmatch(w["text"])]
            # Two values minimum (Unit Price and Total Price). This is
            # what separates a real item row from page furniture that
            # happens to start with a word-shaped token ("SAWO INC.,").
            if len(money) < 2:
                continue
            prices[first["text"]] = parse_number(money[0]["text"])
    return prices


def parse_proforma(pdf_path):
    """Return ({product_code: unit_price}, currency) for the Pro-Forma
    Invoice PDF at pdf_path."""
    with pdfplumber.open(pdf_path) as pdf:
        if not pdf.pages or len(pdf.pages[0].chars) < 20:
            fail("No text layer found — file may be a scan, aborting.")
        pages = [_rows(page) for page in pdf.pages]
        currency = detect_currency("\n".join(page.extract_text() or "" for page in pdf.pages))

    prices = _parse_bracketed(pages)
    if not prices:
        # No bracketed codes anywhere - this is the bare-code template
        # rather than a malformed Odoo one.
        prices = _parse_bare_code(pages)

    if not prices:
        fail(
            "Could not find any product-code line items in the Pro-Forma Invoice PDF "
            "(tried both the bracketed-code and bare-code layouts)."
        )
    return prices, currency


def parse_proforma_prices(pdf_path):
    """Prices only - kept for callers that don't care about currency."""
    prices, _currency = parse_proforma(pdf_path)
    return prices
