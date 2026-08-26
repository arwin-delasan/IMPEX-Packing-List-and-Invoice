"""Extract header fields and line items from an Odoo-native packing-list PDF.

A second, distinct template from the one pdf_parser.py handles - same
"SAWO INC." packing list in spirit, but rendered by Odoo's own report
engine, which lays out both the header block and the item table
differently:

- Header: the company name can wrap onto its own line before the address
  (instead of sitting right after "Sold To:"), and a multi-value field
  like Container No. can span several visual lines that interleave with
  the neighboring column when the page is linearized to plain text. Since
  pdf_parser.py's parse_header() is a regex-on-text-lines approach that
  assumes a fixed line layout, it breaks on this template. Here, header
  fields are found by word x-position instead: a left column (Sold To /
  address / Tel / Fax / ATTN) and a right column (Date / Invoice No. /
  Container No. / Seal No. / Destination / Payment Terms), with any
  unlabeled continuation row assigned to whichever label row it sits
  closest to vertically.
- Item table: each row's tokens stream in the same order as the visual
  column headers (External Code, Description, UM, Packages, Net Weight,
  Gross Weight, CBM, Qty) - unlike pdf_parser.py's template, where the
  stream order doesn't match the visual header order at all. That makes
  the row parse simpler: Code is always the first token, and the last 4
  tokens are always Net/Gross/CBM/Qty, with Packages and UM immediately
  before those (Packages can be "Part of <container>", same as
  pdf_parser.py, which shifts UM one slot further back).

classify()/classify_items() and everything downstream of extract_pdf()
are unaware of which parser produced the data - both return the same
dict/list shapes, and the GUI/CLI simply choose which parser to call.
"""

import re

import pdfplumber

from pdf_parser import cluster_rows, fail, parse_number, reformat_date

LEFT_COLUMN_MAX_X = 300
RIGHT_COLUMN_MIN_X = 300
HEADER_BLOCK_MIN_TOP = 100  # above this is the letterhead/title banner, not header fields

RIGHT_COLUMN_LABELS = [
    ("date", ["Date:"]),
    ("invoice_no", ["Invoice", "No.:"]),
    ("container_no", ["Container", "No.:"]),
    ("seal_no", ["Seal", "No.:"]),
    ("destination", ["Destination:"]),
    ("payment_terms", ["Payment", "Terms:"]),
]


# ---------------------------------------------------------------------------
# PDF header block extraction (geometry-based)
# ---------------------------------------------------------------------------

def parse_header(page):
    # The header block's height varies with how long the address/ATTN
    # content is, so a fixed top cutoff isn't safe - a short header
    # (fewer address lines) lets the item table start higher up the page,
    # and its column header ("Number of / Net / Gross / ... ") words can
    # fall inside a fixed window and get mistaken for a continuation line
    # of whatever label sits closest (this really happened: it silently
    # appended "Number of Net Gross CBM Qty" onto a Destination value).
    # Finding the item table's own header row and cutting off there,
    # same way parse_items_on_page() locates it, is exact regardless of
    # how tall the header block is.
    all_rows = cluster_rows(page.extract_words(keep_blank_chars=False, x_tolerance=1.5))
    item_table_row = next(
        (r for r in all_rows if "External" in [w["text"] for w in r["words"]] and "Code" in [w["text"] for w in r["words"]]),
        None,
    )
    max_top = item_table_row["top"] if item_table_row is not None else float("inf")
    if item_table_row is not None:
        # The item table's own header wraps onto an extra line just above
        # "External Code ..." (the first half of two-line column headers
        # like "Number of\nPackages") - fold that line's top in too, or
        # its words get mistaken for a continuation of whatever header
        # field sits just above it (this really happened: "Number of Net
        # Gross" got appended onto a Destination value).
        wrapped_row = next(
            (r for r in all_rows if r["top"] < item_table_row["top"] and r["words"][0]["text"] == "Number"),
            None,
        )
        if wrapped_row is not None:
            max_top = wrapped_row["top"]

    words = [w for row in all_rows for w in row["words"] if HEADER_BLOCK_MIN_TOP < w["top"] < max_top]
    left_rows = cluster_rows([w for w in words if w["x0"] < LEFT_COLUMN_MAX_X])
    right_rows = cluster_rows([w for w in words if w["x0"] >= RIGHT_COLUMN_MIN_X])

    def row_value(row, skip):
        return " ".join(w["text"] for w in row["words"][skip:]).strip()

    def find_left_row(prefix_tokens):
        return next(
            (i for i, r in enumerate(left_rows) if [w["text"] for w in r["words"][:len(prefix_tokens)]] == prefix_tokens),
            None,
        )

    # --- left column: Sold To / address / Tel / Fax / ATTN ---
    sold_to_idx = find_left_row(["Sold", "To:"])
    if sold_to_idx is None:
        fail("Header label not found in PDF: 'Sold To:'")
    sold_to_company = row_value(left_rows[sold_to_idx], 2)

    tel_idx = find_left_row(["Tel.", "No.:"])
    if tel_idx is None:
        fail("Header label not found in PDF: 'Tel. No.:'")
    address_lines = [row_value(r, 0) for r in left_rows[sold_to_idx + 1:tel_idx]]
    tel_no = row_value(left_rows[tel_idx], 2)

    fax_idx = find_left_row(["Fax", "No.:"])
    fax_no = row_value(left_rows[fax_idx], 2) if fax_idx is not None else ""

    attn_idx = next((i for i, r in enumerate(left_rows) if r["words"][0]["text"].startswith("ATTN")), None)
    attn = row_value(left_rows[attn_idx], 1) if attn_idx is not None else ""

    # --- right column: Date / Invoice No. / Container No. / Seal No. / Destination / Payment Terms ---
    labeled = []
    orphans = []
    for row in right_rows:
        texts = [w["text"] for w in row["words"]]
        match = next((entry for entry in RIGHT_COLUMN_LABELS if texts[:len(entry[1])] == entry[1]), None)
        if match:
            key, pattern = match
            labeled.append({"key": key, "top": row["top"], "tokens": texts[len(pattern):], "extra_rows": []})
        else:
            orphans.append({"top": row["top"], "tokens": texts})

    if not labeled:
        fail("Could not locate Date/Invoice No./Container No./Seal No./Destination labels in PDF header.")

    # An orphan row is a continuation line of a multi-line value (e.g. a
    # Container No. holding several package numbers) - it can land either
    # above or below its label's own row, so "nearest by vertical
    # distance" is used rather than "next row after the label".
    for orphan in orphans:
        nearest = min(labeled, key=lambda entry: abs(entry["top"] - orphan["top"]))
        nearest["extra_rows"].append(orphan)

    right_fields = {}
    for entry in labeled:
        rows_in_order = sorted([entry] + entry["extra_rows"], key=lambda r: r["top"])
        right_fields[entry["key"]] = " ".join(" ".join(r["tokens"]) for r in rows_in_order if r["tokens"]).strip()

    for required in ("date", "invoice_no", "container_no", "seal_no", "destination"):
        if required not in right_fields:
            fail(f"Header label not found in PDF: {required!r}")

    return {
        "sold_to_company": sold_to_company,
        "sold_to_address": address_lines,
        "date": reformat_date(right_fields["date"]),
        "invoice_no": right_fields["invoice_no"],
        "container_no": right_fields["container_no"],
        "seal_no": right_fields["seal_no"],
        "destination": right_fields["destination"],
        "tel_no": tel_no,
        "payment_terms": right_fields.get("payment_terms", ""),
        "fax_no": fax_no,
        "attn": attn,
    }


# ---------------------------------------------------------------------------
# PDF line-item table extraction (Code-first stream order)
# ---------------------------------------------------------------------------

def classify_item_row(words):
    """Classify one row's words into item fields. Code is always the
    first token; Net Weight, Gross Weight, and CBM are always the last 3
    tokens before Qty, with UM and Packages immediately before that -
    parsed from the end backward for the same reason pdf_parser.py does:
    it's the only boundary that's reliably fixed regardless of how long
    the description is."""
    texts = [w["text"] for w in words]
    n = len(texts)

    if n < 7:
        fail(f"Row too short to contain Code/UM/Packages/Net/Gross/CBM/Qty: {texts}")

    external_code = texts[0]

    try:
        qty = int(texts[n - 1].replace(",", ""))
        cbm = parse_number(texts[n - 2])
        gross_weight = parse_number(texts[n - 3])
        net_weight = parse_number(texts[n - 4])
    except ValueError:
        fail(f"Could not parse Net Weight / Gross Weight / CBM / Qty in row: {texts}")

    if texts[n - 7:n - 4] and texts[n - 7] == "Part" and texts[n - 6] == "of":
        packages = f"Part of {texts[n - 5]}"
        um_idx = n - 8
    else:
        try:
            packages = int(texts[n - 5].replace(",", ""))
        except ValueError:
            fail(f"Could not parse Packages in row: {texts}")
        um_idx = n - 6

    if um_idx < 1:
        fail(f"Row too short to contain a Description between Code and UM: {texts}")

    # Unlike pdf_parser.py's template, UM's position here is already
    # fixed by the Net/Gross/CBM/Qty/Packages boundary found above (proven
    # by the Grand Total row's totals matching exactly), so whatever token
    # sits there simply *is* the UM value - no need to validate it against
    # a fixed vocabulary (this template spells units out in full, e.g.
    # "Piece", not the abbreviations UM_CODES was built for).
    um = texts[um_idx]
    desc_words = texts[1:um_idx]

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
    """Unlike pdf_parser.py's anchor-based version, this template's
    Grand Total row tokens already stream in column order (Packages, Net,
    Gross, CBM, Qty), so the last 5 numeric tokens can be taken directly."""
    texts = [w["text"] for w in words]
    numeric = [t for t in texts if re.fullmatch(r"-?[\d,]*\.?\d+", t)]
    if len(numeric) < 5:
        fail(f"Grand Total row missing expected values: {texts}")
    packages, net_weight, gross_weight, cbm, qty = numeric[-5:]
    return {
        "packages": int(packages.replace(",", "")),
        "net_weight": parse_number(net_weight),
        "gross_weight": parse_number(gross_weight),
        "cbm": parse_number(cbm),
        "qty": int(qty.replace(",", "")),
    }


def _is_item_row_start(words):
    # Code-first rows have no reliable marker at the *start* (the code
    # can be almost anything) - Qty at the *end* is the reliable one.
    return bool(re.fullmatch(r"-?[\d,]+", words[-1]["text"]))


def parse_items_on_page(page):
    """Return (item_row_word_lists, grand_total_words_or_None) for one page.

    Mirrors pdf_parser.py's approach (repeated per-page header, shared
    Grand Total row on the final page) - see its parse_items_on_page for
    the fuller rationale. Not use_text_flow, for the same reason
    proforma_parser.py and this module's own parse_header() aren't: a
    value that wraps onto an extra line can land between that line's
    "anchor" row and the rest of its own row in content-stream order,
    which defeats cluster_rows' sequential top-proximity grouping.
    Default extract_words sorts by (top, x0) instead, which groups rows
    correctly regardless of stream order."""
    words = page.extract_words(keep_blank_chars=False, x_tolerance=1.5)
    rows = cluster_rows(words)

    header_idx = grand_idx = None
    for idx, row in enumerate(rows):
        texts = [w["text"] for w in row["words"]]
        if header_idx is None and "External" in texts and "Code" in texts:
            header_idx = idx
        if "Grand" in texts and ("Total" in texts or "Total:" in texts):
            grand_idx = idx
            break

    grand_total_words = rows[grand_idx]["words"] if grand_idx is not None else None

    if header_idx is None:
        return [], grand_total_words

    start_idx = header_idx + 1
    while start_idx < len(rows) and not _is_item_row_start(rows[start_idx]["words"]):
        start_idx += 1

    end_idx = grand_idx if grand_idx is not None else len(rows)

    # A product code too long for its column wraps onto an extra line
    # (e.g. "LED-SUPPLY-60-" / "V2"), which renders as its own row: a
    # single word, sitting just below the row it belongs to and
    # left-aligned to the same x0 as that row's code. Left unmerged, this
    # breaks two things at once - the fragment row itself doesn't end in
    # a number so it's dropped as "not an item row", *and* the row that
    # was split loses its second half. Reattach it onto the code before
    # classifying.
    item_rows = []
    for row in rows[start_idx:end_idx]:
        if _is_item_row_start(row["words"]):
            item_rows.append(row["words"])
            continue
        if len(row["words"]) == 1 and item_rows:
            orphan = row["words"][0]
            code_word = item_rows[-1][0]
            if abs(orphan["x0"] - code_word["x0"]) < 5 and abs(orphan["top"] - code_word["top"]) < 15:
                item_rows[-1][0] = {**code_word, "text": code_word["text"] + orphan["text"]}

    return item_rows, grand_total_words


def parse_items(pdf):
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
        header = parse_header(page)
        items, grand_total = parse_items(pdf)
    return header, items, grand_total
