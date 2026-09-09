"""Extract header fields and line items from a SAWO, INC. packing-list PDF.

Geometry-based (word x-position/stream order), not whitespace-splitting -
see classify_item_row for why. Shared by generate_pl1.py.
"""

import re
from datetime import datetime

import pdfplumber

UM_CODES = ("BX", "UN", "PC", "RL", "ST", "BOT", "PL", "PK", "SET")

# Right-hand column labels, and the left-hand ones that share their
# lines. Header fields are found by locating these labels rather than by
# counting lines - see parse_header.
RIGHT_LABELS = [
    ("date", "Date:"),
    ("invoice_no", "Invoice No.:"),
    ("container_no", "Container No.:"),
    ("seal_no", "Seal No.:"),
    ("destination", "Destination:"),
    ("payment_terms", "Payment Terms:"),
]

LEFT_LABELS = [
    ("sold_to_company", "Sold To:"),
    ("tel_no", "Tel. No.:"),
    ("fax_no", "Fax No.:"),
    ("attn", "ATTN.:"),
]


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
    """Read the header block by locating each field's own label, rather
    than by counting lines from "Sold To:".

    The previous version took a fixed slice of exactly seven lines after
    "Sold To:" and required each one to match its own regex, so any
    variation in the block's shape - a customer with two or four address
    lines instead of three, a missing Fax No., an extra field - shifted
    every following line and failed the whole extraction. Labels are
    stable where line offsets aren't, so each field is found wherever it
    actually sits, and the optional ones (Fax No., ATTN., Payment Terms)
    are simply left blank when absent instead of aborting.
    """
    lines = text.splitlines()

    start = next((i for i, l in enumerate(lines) if "Sold To:" in l), None)
    if start is None:
        fail("Header label not found in PDF: 'Sold To:'")
    # The item table's own column header ends the block - the same
    # boundary parse_items_on_page() keys off, and exact regardless of
    # how tall the header happens to be.
    end = next(
        (i for i in range(start + 1, len(lines)) if "External" in lines[i] and "Code" in lines[i]),
        len(lines),
    )

    right_fields = {}
    right_offsets = {}
    left_parts = []
    for offset, line in enumerate(lines[start:end]):
        rest = line
        for key, label in RIGHT_LABELS:
            pos = rest.find(label)
            if pos != -1:
                right_fields[key] = rest[pos + len(label):].strip()
                right_offsets[key] = offset
                rest = rest[:pos]
                break
        left_parts.append((offset, rest.strip()))

    left_fields = {}
    address_candidates = []
    for offset, part in left_parts:
        label_match = next(((k, l) for k, l in LEFT_LABELS if part.startswith(l)), None)
        if label_match:
            key, label = label_match
            left_fields[key] = part[len(label):].strip()
        elif part:
            address_candidates.append((offset, part))

    for key in ("date", "invoice_no", "container_no", "seal_no", "destination"):
        if key not in right_fields:
            fail(f"Header label not found in PDF: {dict(RIGHT_LABELS)[key]!r}")
    for key in ("sold_to_company", "tel_no"):
        if key not in left_fields:
            fail(f"Header label not found in PDF: {dict(LEFT_LABELS)[key]!r}")

    # The address occupies the unlabelled left-hand text above the
    # Destination line - however many lines that turns out to be.
    address_lines = [p for offset, p in address_candidates if offset < right_offsets["destination"]]

    return {
        "sold_to_company": left_fields["sold_to_company"],
        "sold_to_address": address_lines,
        "date": reformat_date(right_fields["date"]),
        "invoice_no": right_fields["invoice_no"],
        "container_no": right_fields["container_no"],
        "seal_no": right_fields["seal_no"],
        "destination": right_fields["destination"],
        "tel_no": left_fields["tel_no"],
        "payment_terms": right_fields.get("payment_terms", ""),
        "fax_no": left_fields.get("fax_no", ""),
        "attn": left_fields.get("attn", ""),
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


def _texts(row):
    return [w["text"] for w in row["words"]]


def classify_item_row(words, qty_min_x):
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
        qty_idx = next((j for j in range(i, n) if words[j]["x0"] >= qty_min_x), None)
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


def column_anchors(page):
    """Measure the numeric columns' x centres off the item table's own
    column header row, instead of hard-coding them.

    parse_grand_total() reads that row by nearest-anchor matching, which
    always returns *some* column - so hard-coded anchors don't fail on a
    differently sized or differently margined rendering of this
    template, they silently map values onto the wrong fields. Anchors
    measured from the header that's actually on the page stay correct at
    any page size or scale.

    Returns None when the header row isn't on this page (a continuation
    page carrying no item table), leaving the caller to reuse the last
    page's measurements.
    """
    rows = cluster_rows(page.extract_words(keep_blank_chars=False, x_tolerance=1.5))
    header = next(
        (r for r in rows if "External" in _texts(r) and "Code" in _texts(r)),
        None,
    )
    if header is None:
        return None

    # "Number of / Packages", "Net / Weight" and "Gross / Weight" wrap
    # onto a second header line, which carries those three columns' real
    # centres; the single-word CBM and Qty stay on the first.
    wrapped = next(
        (r for r in rows if r["top"] > header["top"] and _texts(r)[:1] == ["Packages"]),
        None,
    )

    def centre(row, text, occurrence=0):
        if row is None:
            return None
        hits = [w for w in row["words"] if w["text"] == text]
        if len(hits) <= occurrence:
            return None
        return (hits[occurrence]["x0"] + hits[occurrence]["x1"]) / 2

    anchors = {
        "packages": centre(wrapped, "Packages") or centre(header, "Number"),
        "net_weight": centre(wrapped, "Weight", 0) or centre(header, "Net"),
        "gross_weight": centre(wrapped, "Weight", 1) or centre(header, "Gross"),
        "cbm": centre(header, "CBM"),
        "qty": centre(header, "Qty"),
    }
    if any(v is None for v in anchors.values()):
        return None
    return anchors


def qty_column_min_x(anchors):
    """Left edge of the Qty column, midway between the CBM and Qty header
    centres. Qty is the only field rendered that far right (see
    classify_item_row); measured against every sample, CBM values reach
    x0 518 at most and Qty values start at x0 547, so the midpoint sits
    well inside the gap - and it tracks the page instead of assuming
    one."""
    return (anchors["cbm"] + anchors["qty"]) / 2


def parse_grand_total(words, anchors):
    numeric_words = [w for w in words if re.fullmatch(r"-?[\d,]*\.?\d+", w["text"])]
    result = {}
    for w in numeric_words:
        key = min(anchors, key=lambda k: abs(anchors[k] - w["x0"]))
        result[key] = parse_number(w["text"])
    missing = set(anchors) - result.keys()
    if missing:
        fail(f"Grand Total row missing values for: {sorted(missing)}")
    result["qty"] = int(result["qty"])
    result["packages"] = int(result["packages"])
    return result


def _is_item_row_start(words):
    first_word = words[0]["text"]
    return first_word == "Part" or bool(re.fullmatch(r"-?[\d,]+", first_word))


def parse_items_on_page(page):
    """Return (item_row_word_lists, grand_total_words_or_None, anchors_or_None)
    for one page.

    Multi-page packing lists repeat the item-table header ("External Code
    / Description / UM / ...") on every page. If a page has no such header
    (verified: this PDF always repeats it, but a stray page shouldn't be
    mis-scanned), it's treated as having no items rather than scanning
    from the top of the page - otherwise page furniture starting with a
    number (e.g. a street address like "16635 KEELE STREET") could be
    mistaken for an item row, since the only real signal for "is this an
    item row" is its first word looking like a Packages value.
    """
    anchors = column_anchors(page)
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
        return [], grand_total_words, anchors

    # The column header wraps onto a second PDF line (e.g. "Number of
    # Packages" / "Net Weight" continue below "External Code Description
    # UM..."); skip forward to the first row that actually starts a data
    # row.
    start_idx = header_idx + 1
    while start_idx < len(rows) and not _is_item_row_start(rows[start_idx]["words"]):
        start_idx += 1

    end_idx = grand_idx if grand_idx is not None else len(rows)
    item_rows = [row["words"] for row in rows[start_idx:end_idx] if _is_item_row_start(row["words"])]
    return item_rows, grand_total_words, anchors


def parse_items(pdf):
    """Collect item rows across every page until the 'Grand Total' row is
    found - packing lists with enough items span multiple PDF pages, each
    repeating the item-table header but sharing one Grand Total row on the
    final page."""
    all_item_rows = []  # (page number, row words)
    grand_total_words = None
    grand_total_page = None
    anchors = None

    for page_no, page in enumerate(pdf.pages, start=1):
        item_rows, page_grand_total_words, page_anchors = parse_items_on_page(page)
        if page_anchors is not None:
            anchors = page_anchors
        all_item_rows.extend((page_no, words) for words in item_rows)
        if page_grand_total_words is not None:
            grand_total_words = page_grand_total_words
            grand_total_page = page_no
            break

    if not all_item_rows:
        fail("Could not locate item table header row ('External Code') in PDF.")
    if grand_total_words is None:
        fail("Could not locate 'Grand Total' row in PDF.")
    if anchors is None:
        fail("Could not measure the item table's column positions from its header row.")

    # Where a row fails to parse, say which page and how far down it sits -
    # the token dump alone doesn't locate it in a 176-row, multi-page list.
    qty_min_x = qty_column_min_x(anchors)
    items = []
    for page_no, words in all_item_rows:
        try:
            items.append(classify_item_row(words, qty_min_x))
        except ExtractionError as e:
            fail(f"{e} (page {page_no}, row top {words[0]['top']:.0f})")

    try:
        grand_total = parse_grand_total(grand_total_words, anchors)
    except ExtractionError as e:
        fail(f"{e} (page {grand_total_page})")
    return items, grand_total


def extract_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        if len(page.chars) < 20:
            fail("No text layer found — file may be a scan, aborting.")
        header = parse_header(page.extract_text() or "")
        items, grand_total = parse_items(pdf)
    return header, items, grand_total
