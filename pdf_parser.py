"""Extract header fields and line items from a SAWO, INC. packing-list PDF.

Geometry-based (word x-position/stream order), not whitespace-splitting -
see classify_item_row for why. Shared by generate_pl1.py.
"""

import re
from datetime import datetime

import pdfplumber

UM_CODES = ("BX", "UN", "PC", "RL", "ST", "BOT", "PL", "PK", "SET")

# Qty is the right-most column; every other token in an item row ends by
# x1 ~533 (CBM), so this cleanly isolates it - see classify_item_row.
QTY_COLUMN_MIN_X = 540

GRAND_TOTAL_ANCHORS = {
    "packages": 365,
    "net_weight": 420,
    "gross_weight": 465,
    "cbm": 513,
    "qty": 550,
}


class ExtractionError(Exception):
    """Raised when the PDF does not match the expected layout."""


def fail(msg):
    raise ExtractionError(msg)


def reformat_date(text):
    """Convert the PDF's "12-Aug-2026" style date to "08/12/2026" (M/D/Y)."""
    try:
        return datetime.strptime(text, "%d-%b-%Y").strftime("%m/%d/%Y")
    except ValueError:
        fail(f"Could not parse date {text!r} (expected format like '12-Aug-2026').")


# ---------------------------------------------------------------------------
# PDF header block extraction
# ---------------------------------------------------------------------------

def parse_header(text):
    lines = text.split("\n")

    def find_line(label):
        for l in lines:
            if label in l:
                return l
        fail(f"Header label not found in PDF: {label!r}")

    l1 = find_line("Sold To:")
    m1 = re.search(r"Sold To:\s*(.*?)\s+Date:\s*(.*)", l1)
    if not m1:
        fail(f"Could not parse Sold To / Date line: {l1!r}")

    idx = lines.index(l1)
    if idx + 7 >= len(lines):
        fail("PDF header block is shorter than expected.")
    l2, l3, l4, l5, l6, l7, l8 = lines[idx + 1: idx + 8]

    m2 = re.search(r"^(.*?)\s*Invoice No\.:\s*(.*)$", l2)
    m3 = re.search(r"^(.*?)\s*Container No\.:\s*(.*)$", l3)
    m4 = re.search(r"^(.*?)\s*Seal No\.:\s*(.*)$", l4)
    m5 = re.search(r"Destination:\s*(.*)$", l5)
    m6 = re.search(r"^Tel\. No\.:\s*(.*?)\s+Payment Terms:\s*(.*)$", l6)
    m7 = re.search(r"^Fax No\.:\s*(.*)$", l7)
    m8 = re.search(r"^ATTN\.:\s*(.*)$", l8)

    for m, src in ((m2, l2), (m3, l3), (m4, l4), (m5, l5), (m6, l6), (m7, l7), (m8, l8)):
        if not m:
            fail(f"Could not parse header line: {src!r}")

    address_lines = [g.strip() for g in (m2.group(1), m3.group(1), m4.group(1)) if g.strip()]

    return {
        "sold_to_company": m1.group(1).strip(),
        "sold_to_address": address_lines,
        "date": reformat_date(m1.group(2).strip()),
        "invoice_no": m2.group(2).strip(),
        "container_no": m3.group(2).strip(),
        "seal_no": m4.group(2).strip(),
        "destination": m5.group(1).strip(),
        "tel_no": m6.group(1).strip(),
        "payment_terms": m6.group(2).strip(),
        "fax_no": m7.group(1).strip(),
        "attn": m8.group(1).strip(),
    }


# ---------------------------------------------------------------------------
# PDF line-item table extraction (geometry-based, not whitespace splitting)
# ---------------------------------------------------------------------------

def parse_number(text):
    """Parse a PDF number that may use a thousands separator, e.g. "1,065.00"."""
    return float(text.replace(",", ""))


def cluster_rows(words, tol=2.5):
    """Group words (already top-to-bottom ordered) into visual rows,
    preserving each word's original stream order within the row."""
    rows = []
    for w in words:
        if rows and abs(w["top"] - rows[-1]["top"]) <= tol:
            rows[-1]["words"].append(w)
        else:
            rows.append({"top": w["top"], "words": [w]})
    return rows


def classify_item_row(words):
    """Classify one row's words into item fields.

    Columns are identified by their fixed stream order (Packages, External
    Code, Qty, Description..., UM, Gross Weight, CBM, Net Weight) rather than
    by x-position ranges, because long descriptions in this PDF visually
    overlap the UM/Packages columns when they don't wrap to a second line.
    The one exception is the Qty column, which is found by x-position -
    stream order alone can't say where a multi-token External Code ends
    and Qty begins, and nothing else in the row reaches that far right.

    Net Weight, Gross Weight, and CBM are always the row's last three
    tokens, and the UM code is always the token immediately before them -
    so the UM/description boundary is found by parsing from the *end* of
    the row backward, not by scanning forward for the first UM-looking
    token. Forward-scanning previously misfired on description text that
    merely *contains* a UM code as a substring (e.g. a model number like
    "1616RL", which ends in "RL" - a valid UM code - but sits at the start
    of the description, not the end). The UM code can also render fused
    onto the last description word (e.g. "RoomUN"); that is split back
    apart here.
    """
    texts = [w["text"] for w in words]
    n = len(texts)
    i = 0

    try:
        if texts[i] == "Part":
            # "Part of Crate", "Part of Box", etc. - the container word varies.
            if texts[i + 1] != "of":
                fail(f"Unexpected 'Part' token in row: {texts}")
            packages = f"Part of {texts[i + 2]}"
            i += 3
        else:
            packages = int(texts[i].replace(",", ""))
            i += 1

        # An External Code can contain a literal space (e.g.
        # "STP-GASKET 6H.E."), which extract_words returns as two
        # separate tokens - so the code is not always exactly one token.
        # Assuming it was made Qty read the code's second half instead
        # and fail the whole extraction. Qty is the only field rendered
        # in the far-right Qty column, so it is located by x-position
        # and everything between Packages and it is the code.
        qty_idx = next((j for j in range(i, n) if words[j]["x0"] >= QTY_COLUMN_MIN_X), None)
        if qty_idx is None:
            fail(f"Could not locate a Qty column value in row: {texts}")
        if qty_idx == i:
            fail(f"Row has no External Code between Packages and Qty: {texts}")

        external_code = " ".join(texts[i:qty_idx])
        qty = int(texts[qty_idx].replace(",", ""))
        i = qty_idx + 1
    except (ValueError, IndexError):
        fail(f"Could not parse External Code / Qty / Packages in row: {texts}")

    if n - i < 4:
        fail(f"Row too short to contain UM/Net Weight/Gross Weight/CBM: {texts}")

    try:
        net_weight = parse_number(texts[n - 3])
        gross_weight = parse_number(texts[n - 2])
        cbm = parse_number(texts[n - 1])
    except ValueError:
        fail(f"Could not parse Net Weight / Gross Weight / CBM in row: {texts}")

    um_token = texts[n - 4]
    if um_token in UM_CODES:
        um = um_token
        desc_words = texts[i:n - 4]
    else:
        suffix = next((c for c in UM_CODES if um_token.endswith(c) and len(um_token) > len(c)), None)
        if suffix is not None:
            um = suffix
            desc_words = texts[i:n - 4] + [um_token[: -len(suffix)]]
        else:
            # An unrecognized UM code (not in UM_CODES, e.g. a new
            # shipping-unit abbreviation we haven't seen before) shouldn't
            # block the whole packing list from generating - leave it
            # blank for a human to fill in/investigate rather than failing
            # the entire extraction over one row.
            um = ""
            desc_words = texts[i:n - 4]

    return {
        "external_code": external_code,
        "description": " ".join(desc_words).strip(),
        "um": um,
        "packages": packages,
        "net_weight": net_weight,
        "gross_weight": gross_weight,
        "cbm": cbm,
        "qty": qty,
    }


def parse_grand_total(words):
    numeric_words = [w for w in words if re.fullmatch(r"-?[\d,]*\.?\d+", w["text"])]
    result = {}
    for w in numeric_words:
        key = min(GRAND_TOTAL_ANCHORS, key=lambda k: abs(GRAND_TOTAL_ANCHORS[k] - w["x0"]))
        result[key] = parse_number(w["text"])
    missing = set(GRAND_TOTAL_ANCHORS) - result.keys()
    if missing:
        fail(f"Grand Total row missing values for: {sorted(missing)}")
    result["qty"] = int(result["qty"])
    result["packages"] = int(result["packages"])
    return result


def _is_item_row_start(words):
    first_word = words[0]["text"]
    return first_word == "Part" or bool(re.fullmatch(r"-?[\d,]+", first_word))


def parse_items_on_page(page):
    """Return (item_row_word_lists, grand_total_words_or_None) for one page.

    Multi-page packing lists repeat the item-table header ("External Code
    / Description / UM / ...") on every page. If a page has no such header
    (verified: this PDF always repeats it, but a stray page shouldn't be
    mis-scanned), it's treated as having no items rather than scanning
    from the top of the page - otherwise page furniture starting with a
    number (e.g. a street address like "16635 KEELE STREET") could be
    mistaken for an item row, since the only real signal for "is this an
    item row" is its first word looking like a Packages value.
    """
    words = page.extract_words(use_text_flow=True, keep_blank_chars=False, x_tolerance=1.5)
    rows = cluster_rows(words)

    header_idx = grand_idx = None
    for idx, row in enumerate(rows):
        texts = [w["text"] for w in row["words"]]
        if header_idx is None and "External" in texts and "Code" in texts:
            header_idx = idx
        if "Grand" in texts and "Total" in texts:
            grand_idx = idx
            break

    grand_total_words = rows[grand_idx]["words"] if grand_idx is not None else None

    if header_idx is None:
        return [], grand_total_words

    # The column header wraps onto a second PDF line (e.g. "Number of
    # Packages" / "Net Weight" continue below "External Code Description
    # UM..."); skip forward to the first row that actually starts a data
    # row.
    start_idx = header_idx + 1
    while start_idx < len(rows) and not _is_item_row_start(rows[start_idx]["words"]):
        start_idx += 1

    end_idx = grand_idx if grand_idx is not None else len(rows)
    item_rows = [row["words"] for row in rows[start_idx:end_idx] if _is_item_row_start(row["words"])]
    return item_rows, grand_total_words


def parse_items(pdf):
    """Collect item rows across every page until the 'Grand Total' row is
    found - packing lists with enough items span multiple PDF pages, each
    repeating the item-table header but sharing one Grand Total row on the
    final page."""
    all_item_rows = []
    grand_total_words = None

    for page in pdf.pages:
        item_rows, page_grand_total_words = parse_items_on_page(page)
        all_item_rows.extend(item_rows)
        if page_grand_total_words is not None:
            grand_total_words = page_grand_total_words
            break

    if not all_item_rows:
        fail("Could not locate item table header row ('External Code') in PDF.")
    if grand_total_words is None:
        fail("Could not locate 'Grand Total' row in PDF.")

    items = [classify_item_row(words) for words in all_item_rows]
    grand_total = parse_grand_total(grand_total_words)
    return items, grand_total


def extract_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        if len(page.chars) < 20:
            fail("No text layer found — file may be a scan, aborting.")
        header = parse_header(page.extract_text() or "")
        items, grand_total = parse_items(pdf)
    return header, items, grand_total
