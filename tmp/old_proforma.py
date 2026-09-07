"""Extract per-product-code Unit Price from a SAWO Pro-Forma Invoice PDF.

Odoo-generated: each item line's Description column starts with the
product code in brackets, e.g. "[SCA-90NS-P-C] Scandia 3H.E 9,0kW NS
Premium 3P-1P" - unambiguous, so unlike pdf_parser.py's packing-list
extraction this doesn't need any UM-suffix disambiguation. Fee lines
(Bank Charge, Freight Costs, Export Processing Fee, etc.) have no
bracketed code and are skipped - they aren't shipped product, so they
never appear on the packing list this price gets matched against.

Only Unit Price is extracted (that's all generate_inv.py needs); CBM,
Weight, Taxes, and Amount are read only well enough to be told apart
from Unit Price, not stored.
"""

import re

import pdfplumber

from pdf_parser import ExtractionError, cluster_rows, fail, parse_number

CODE_RE = re.compile(r"^\[(.+)\]$")
DECIMAL_RE = re.compile(r"-?[\d,]+\.\d+")

# x0 anchors (points from the page's left edge) for the proforma's
# numeric columns - same nearest-anchor approach as
# pdf_parser.GRAND_TOTAL_ANCHORS. Only used to tell Unit Price apart
# from CBM/Weight/Amount on the same row; Quantity is excluded already
# since it's an integer (no decimal point) and never matches DECIMAL_RE.
COLUMN_ANCHORS = {
    "cbm": 365,
    "weight": 410,
    "unit_price": 450,
    "amount": 520,
}


def parse_proforma_prices(pdf_path):
    """Return {product_code: unit_price} for every bracketed-code line
    item in the Pro-Forma Invoice PDF at pdf_path."""
    prices = {}
    with pdfplumber.open(pdf_path) as pdf:
        if not pdf.pages or len(pdf.pages[0].chars) < 20:
            fail("No text layer found — file may be a scan, aborting.")
        for page in pdf.pages:
            # Not use_text_flow: this PDF's content stream interleaves a
            # wrapped description's continuation line *between* the first
            # line and that row's numeric columns, which defeats
            # cluster_rows' sequential top-proximity grouping. Default
            # extract_words sorts by (top, x0) instead, which reconstructs
            # correct visual rows regardless of stream order.
            words = page.extract_words(keep_blank_chars=False, x_tolerance=1.5)
            for row in cluster_rows(words):
                row_words = row["words"]
                m = CODE_RE.match(row_words[0]["text"])
                if not m:
                    continue
                code = m.group(1)
                decimal_words = [w for w in row_words if DECIMAL_RE.fullmatch(w["text"])]
                if not decimal_words:
                    continue
                nearest = min(decimal_words, key=lambda w: abs(w["x0"] - COLUMN_ANCHORS["unit_price"]))
                prices[code] = parse_number(nearest["text"])

    if not prices:
        fail("Could not find any bracketed product-code line items in the Pro-Forma Invoice PDF.")
    return prices
