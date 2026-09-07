#!/usr/bin/env python3
"""Generate an Invoice (matching the historical "INV" format) from a SAWO
packing-list PDF.

Same extraction/classification pipeline as generate_pl1.py (same PDF, same
category/subtype grouping via categorize.py) - only the rendering differs:
Quantity / UM / Unit Price / Total Price columns instead of Packages / Net
Weight / Gross Weight / CBM, and Sub-Totals / Grand Total sum price instead
of weight.

Unit Price has no data source yet (not on the PDF, and deliberately not
pulled from Odoo - see project history) - left blank for the user to fill
in by hand; Total Price is a live formula (=UnitPrice*Qty per row, summed
for Sub-Totals/Grand Total) that resolves once they do.

Usage:
    python generate_inv.py <input.pdf> <output.xlsx>
"""

import math
import os
import sys

from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.styles import Border

from pdf_parser import ExtractionError
from pdf_parser import extract_pdf as custom_extract_pdf
from odoo_pdf_parser import extract_pdf as odoo_extract_pdf
from generate_pl1 import (
    GenerationError,
    ROW_HEIGHT_PER_LINE,
    add_signature_block,
    build_blocks,
    classify_items,
    copy_style,
    derive_so_numbers,
    um_warnings,
    validate,
)
from odoo_client import OdooError, prompt_login
from overrides import OverridesError
from proforma_parser import parse_proforma
from app_paths import app_path
from pricelist import resolve_bella_vivo_path

INV_REFERENCE_PATH = app_path("INV - Copy.xlsx")
SHEET_NAME = "summary"

# Bella Vivo gets its own dedicated price list (2268 product codes, one
# price each - verified no duplicates). Other customers' pricing comes
# from a separate quote/proforma PDF instead (not wired up yet).
# Fallback only. The edition actually in force is resolved at call time
# by pricelist.resolve_bella_vivo_path(), so a new edition can be adopted
# without editing this file - see pricelist.py.
BELLA_VIVO_PRICELIST_PATH = app_path("BELLA VIVO PRICELIST-0125.xlsx")
BELLA_VIVO_SHEET_NAME = "data"
BELLA_VIVO_HEADER_ROW = 2  # "PRODUCT CATEGORY" / "PRODUCT CODE" / "BV PRICE USD"

# Bella Vivo is who gets billed (the "Invoice:" field), not who the goods
# ship to - Consignee/Address stay whatever the PDF actually says.
BELLA_VIVO_BILLING = {
    "invoice_to_company": "BELLA VIVO LTD.",
    "invoice_to_address": ["38th Floor, Tower One, Lippo Centre", "89 Queensway, Hong Kong"],
}


def apply_bella_vivo_billing(header):
    """Return a copy of `header` with Bella Vivo's billing info added
    under the "Invoice:" field - Consignee/Address (the real ship-to) and
    everything else (Date, Destination, SO numbers) stay as extracted
    from the PDF, untouched."""
    return {**header, **BELLA_VIVO_BILLING}


def load_bella_vivo_prices(path=None):
    """Return {code: price} for every row in the Bella Vivo price list.

    With no path, uses whichever edition is currently in force (see
    pricelist.resolve_bella_vivo_path). Callers that need to know *which*
    file that was - to show it to the user, or to validate a candidate
    file before adopting it - should pass the path explicitly.
    Product codes are stored as a mix of int (e.g. 304) and str (e.g.
    "221-THD") in the source file - normalized to str here so they match
    the string-typed external_code values extracted from the PDF."""
    if path is None:
        path, _source = resolve_bella_vivo_path()
    wb = load_workbook(path, data_only=True)
    ws = wb[BELLA_VIVO_SHEET_NAME]
    prices = {}
    for r in range(BELLA_VIVO_HEADER_ROW + 1, ws.max_row + 1):
        code = ws.cell(r, 2).value
        price = ws.cell(r, 3).value
        if code is None or price is None:
            continue
        prices[str(code).strip()] = price
    return prices

# Column layout (1-indexed) - narrower than PL1's: no Packages/Net/Gross/CBM,
# Quantity moves up next to the description, Unit Price/Total Price added.
COL_CODE, COL_NAME, COL_QTY, COL_UM, COL_UNIT_PRICE, COL_TOTAL_PRICE = range(1, 7)

# Reference rows to copy styles from (found by inspecting INV - Copy.xlsx,
# converted from the historical INV - Copy.xls via LibreOffice headless).
REF_HEADER_ROW = 16      # "Product Code" / "Product Name" / ... row
REF_CATEGORY_ROW = 18
REF_ITEM_ROW = 19
REF_SUBTOTAL_ROW = 26
REF_SUBTYPE_ROW = 28
REF_GRANDTOTAL_ROW = 99


def lookup_price(price_lookup, code):
    """Find `code`'s price, falling back to a case-insensitive match.

    The same product is not always spelled with the same case in the two
    documents - the China/SO 6773 packing list prints "Decal-4L" while
    that shipment's Pro-Forma Invoice prints "DECAL-4L". Matched
    case-sensitively, a real, known price is silently dropped and the
    Unit Price cell ships blank. Codes are otherwise identical strings,
    so folding case is safe: no two distinct SAWO codes differ only by
    capitalization (checked across the Bella Vivo price list's 2268
    codes - folding case yields no collisions).
    """
    price = price_lookup.get(code)
    if price is not None:
        return price
    folded = code.casefold()
    for known, known_price in price_lookup.items():
        if known.casefold() == folded:
            return known_price
    return None


def build_invoice_workbook(header, blocks, output_path, price_lookup=None, currency="USD"):
    """price_lookup: optional {code: price} dict (e.g. from
    load_bella_vivo_prices()). When given, Unit Price is filled in for any
    matching code; codes with no match are noted in the returned warnings
    list and left blank, same as when price_lookup is None entirely.

    currency: the code printed on the Grand Total row. Not always USD -
    a Pro-Forma Invoice can be priced in EUR (see proforma_parser.
    detect_currency), and printing those figures under a "USD" label
    would misstate the invoice total."""
    warnings = []
    ref_wb = load_workbook(INV_REFERENCE_PATH)
    ref_ws = ref_wb[SHEET_NAME]

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    # --- copy the static title + header-field block (rows 1-15) verbatim ---
    # Unlike PL1 - Copy.xlsx, every header-block value cell in this
    # reference is already blank (label cells only), so there's nothing to
    # strip before writing real values - just copy whatever's there.
    # Exception: the "Invoice:" label (row 3, col A) only makes sense when
    # there's a Bella Vivo bill-to value to put next to it - for a
    # proforma-priced invoice there's no such value, so the whole left
    # column below it (Consignee:/Address:/Contact Person:/Contact No.:)
    # shifts up by the 4 rows "Invoice:" would've taken, preserving the
    # same spacing between them instead of leaving a gap where "Invoice:"
    # used to be.
    # Consignee/Address start at a fixed row (7/8 normally, or 3/4 for a
    # proforma invoice, which has no "Invoice:" value so Consignee moves
    # up to replace it) - matching the reference exactly for the normal
    # 3-line-address case. Everything below that is dynamic, not fixed,
    # so it holds for any address length: Contact Person always sits
    # exactly 1 blank row after wherever the address actually ends, and
    # the item table header always sits exactly 1 blank row below Contact
    # No., wherever that ends up.
    bella_vivo_invoice = bool(header.get("invoice_to_company"))
    shift = 0 if bella_vivo_invoice else 4
    consignee_row = 7 - shift
    address_row = 8 - shift
    contact_person_row = address_row + len(header["sold_to_address"]) + 1
    contact_no_row = contact_person_row + 1
    item_header_row = contact_no_row + 2

    # {source row in the reference -> destination row in the output};
    # "Invoice:" (row 3) only gets replanted when there's a Bella Vivo
    # value for it - otherwise it's just dropped.
    repositioned = {7: consignee_row, 8: address_row, 12: contact_person_row, 13: contact_no_row}
    if bella_vivo_invoice:
        repositioned[3] = 3
    else:
        repositioned[3] = None

    reposition_skip = {(src, COL_CODE) for src in repositioned}
    for r in range(1, REF_HEADER_ROW):
        for c in range(1, ref_ws.max_column + 1):
            if (r, c) in reposition_skip:
                continue
            src = ref_ws.cell(row=r, column=c)
            if src.value not in (None, ""):
                dst = ws.cell(row=r, column=c, value=src.value)
                copy_style(src, dst)
    for src_row, dst_row in repositioned.items():
        if dst_row is None:
            continue
        src = ref_ws.cell(row=src_row, column=COL_CODE)
        copy_style(src, ws.cell(row=dst_row, column=COL_CODE, value=src.value))
    for merge_range in ref_ws.merged_cells.ranges:
        if merge_range.max_row < REF_HEADER_ROW:
            ws.merge_cells(str(merge_range))
    for col_letter, dim in ref_ws.column_dimensions.items():
        if dim.width:
            ws.column_dimensions[col_letter].width = dim.width

    # --- overwrite the header fields we can actually source ---
    # (Invoice number, Reference Number, Mode of Shipment, Freight Terms,
    # Container No., Seal No., Contact Person, Contact No. have no PDF/Odoo
    # source - stay blank, same fields PL1 leaves blank.)
    ws.cell(row=consignee_row, column=COL_NAME, value=header["sold_to_company"])
    for i, line in enumerate(header["sold_to_address"], start=1):
        ws.cell(row=address_row + i - 1, column=COL_NAME, value=line)
    if header.get("invoice_to_company"):
        ws.cell(row=3, column=COL_NAME, value=header["invoice_to_company"])
        for i, line in enumerate(header.get("invoice_to_address", []), start=1):
            ws.cell(row=3 + i, column=COL_NAME, value=line)
    ws.cell(row=4, column=6, value=header["date"])
    ws.cell(row=6, column=6, value=header["destination"])
    ws.cell(row=10, column=6, value=derive_so_numbers(header))

    # --- copy the item-table header row (Product Code / Name / ...) ---
    for c in range(1, COL_TOTAL_PRICE + 1):
        src = ref_ws.cell(row=REF_HEADER_ROW, column=c)
        dst = ws.cell(row=item_header_row, column=c, value=src.value)
        copy_style(src, dst)

    def style_row(dst_row, ref_row, values=None):
        values = values or {}
        for c in range(1, COL_TOTAL_PRICE + 1):
            cell = ws.cell(row=dst_row, column=c, value=values.get(c))
            copy_style(ref_ws.cell(row=ref_row, column=c), cell)

    name_col_width = ws.column_dimensions[get_column_letter(COL_NAME)].width or 40

    row = item_header_row + 1
    style_row(row, REF_HEADER_ROW + 1)  # blank spacer row between the column header and the first block
    ws.row_dimensions[row].height = 3
    # No border on this row's left/right outer edges (mirrors generate_pl1.py).
    left_cell = ws.cell(row=row, column=COL_CODE)
    left_cell.border = Border(top=left_cell.border.top, bottom=left_cell.border.bottom, right=left_cell.border.right)
    right_cell = ws.cell(row=row, column=COL_TOTAL_PRICE)
    right_cell.border = Border(top=right_cell.border.top, bottom=right_cell.border.bottom, left=right_cell.border.left)
    row += 1

    subtotal_rows = []
    prev_category = None
    unit_price_col = get_column_letter(COL_UNIT_PRICE)
    qty_col = get_column_letter(COL_QTY)

    for category, subtype, subchunks in blocks:
        if category != prev_category:
            style_row(row, REF_CATEGORY_ROW, {COL_NAME: f"***{category}***"})
            row += 1
            prev_category = category

        if subtype:
            style_row(row, REF_SUBTYPE_ROW, {COL_CODE: subtype})
            row += 1

        for _subgroup, chunk_items in subchunks:
            if _subgroup and " :: " in _subgroup:
                sub_label = _subgroup.split(" :: ", 1)[1]
                style_row(row, REF_SUBTYPE_ROW, {COL_CODE: sub_label})
                row += 1
            first_item_row = row
            for it in chunk_items:
                vals = {
                    COL_CODE: it["external_code"], COL_NAME: it["name"],
                    COL_QTY: it["qty"], COL_UM: it["um"],
                    COL_TOTAL_PRICE: f"={unit_price_col}{row}*{qty_col}{row}",
                }
                if price_lookup is not None:
                    price = lookup_price(price_lookup, it["external_code"])
                    if price is not None:
                        vals[COL_UNIT_PRICE] = price
                    else:
                        warnings.append(f"{it['external_code']}: not found in price list, Unit Price left blank")
                style_row(row, REF_ITEM_ROW, vals)
                line_count = max(1, math.ceil(len(it["name"]) / name_col_width))
                if line_count > 1:
                    ws.row_dimensions[row].height = ROW_HEIGHT_PER_LINE * line_count
                row += 1
            last_item_row = row - 1

            subtotal_vals = {COL_NAME: "Sub-Total:"}
            for c in (COL_QTY, COL_TOTAL_PRICE):
                col_letter = get_column_letter(c)
                subtotal_vals[c] = f"=SUM({col_letter}{first_item_row}:{col_letter}{last_item_row})"
            style_row(row, REF_SUBTOTAL_ROW, subtotal_vals)
            subtotal_rows.append(row)
            row += 1

    # --- "Nothing Follows" marker, blank spacer, then Grand Total ---
    style_row(row, 97, {COL_NAME: "*** Nothing Follows ***"})
    row += 2  # matches the reference: one blank row before Grand Total

    total_price_col = get_column_letter(COL_TOTAL_PRICE)
    qty_formula = "=SUM(" + ",".join(f"{qty_col}{r}" for r in subtotal_rows) + ")"
    price_formula = "=SUM(" + ",".join(f"{total_price_col}{r}" for r in subtotal_rows) + ")"
    style_row(row, REF_GRANDTOTAL_ROW, {
        COL_NAME: "GRAND TOTAL:", COL_QTY: qty_formula,
        COL_UNIT_PRICE: currency, COL_TOTAL_PRICE: price_formula,
    })
    row += 1
    row += 2  # spacer before the signature block, mirroring generate_pl1.py

    # Signature block: same floating "S A W O, I N C. / Authorized
    # Signatory" shape as the packing list (see add_signature_block in
    # generate_pl1.py), centered against this document's own item table
    # (columns A:F) rather than the packing list's wider one.
    sig_anchor_row = row
    row += 4

    # Print area/setup, matching the reference's own sheet1.xml.
    ws.print_area = f"A1:F{row}"
    ws.print_title_rows = f"1:{item_header_row + 1}"
    ws.page_margins = PageMargins(left=0.2, right=0.2, top=1.4, bottom=0.5, header=0.45, footer=0.51)
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=False)
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.scale = 100
    ws.page_setup.orientation = "portrait"
    ws.oddHeader.right.text = "Page &P of &N"

    wb.save(output_path)
    add_signature_block(ws, output_path, sig_anchor_row, get_column_letter(COL_TOTAL_PRICE))
    return warnings


USAGE = "Usage: python generate_inv.py [--odoo] [--bella-vivo | --proforma <quote.pdf>] <input.pdf> <output.xlsx>"


def main():
    argv = list(sys.argv[1:])
    odoo_pdf = "--odoo" in argv
    if odoo_pdf:
        argv.remove("--odoo")

    bella_vivo = "--bella-vivo" in argv
    if bella_vivo:
        argv.remove("--bella-vivo")

    proforma_path = None
    if "--proforma" in argv:
        idx = argv.index("--proforma")
        if idx + 1 >= len(argv):
            print(USAGE, file=sys.stderr)
            sys.exit(1)
        proforma_path = argv[idx + 1]
        del argv[idx:idx + 2]

    if len(argv) != 2:
        print(USAGE, file=sys.stderr)
        sys.exit(1)

    pdf_path, output_path = argv
    extract_pdf = odoo_extract_pdf if odoo_pdf else custom_extract_pdf

    try:
        header, items, grand_total = extract_pdf(pdf_path)
        print(f"✅ Extracted header block and {len(items)} line items from PDF.")
        for w in um_warnings(items):
            print(f"WARNING: {w}")

        odoo = prompt_login()
        print("✅ Connected to Odoo.")

        classified, warnings = classify_items(items, odoo)
        for w in warnings:
            print(f"WARNING: {w}")
        print(f"✅ Classified {len(classified)} items ({len(warnings)} warning(s)).")

        blocks = build_blocks(classified)
        print(f"✅ Grouped into {len(blocks)} category/subtype block(s).")

        currency = "USD"
        if bella_vivo:
            pricelist_path, source = resolve_bella_vivo_path()
            price_lookup = load_bella_vivo_prices(pricelist_path)
            print(f"✅ Price list ({source}): {os.path.basename(pricelist_path)} "
                  f"- {len(price_lookup)} product code(s).")
            header = apply_bella_vivo_billing(header)
        elif proforma_path:
            price_lookup, currency = parse_proforma(proforma_path)
            print(f"✅ Read {len(price_lookup)} unit price(s) from the Pro-Forma Invoice ({currency}).")
        else:
            price_lookup = None
        price_warnings = build_invoice_workbook(
            header, blocks, output_path, price_lookup=price_lookup, currency=currency,
        )
        for w in price_warnings:
            print(f"WARNING: {w}")
        print(f"✅ Wrote output workbook: {output_path}")
    except (ExtractionError, GenerationError, OdooError, OverridesError) as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)

    totals_warnings = validate(items, grand_total)
    if not totals_warnings:
        print("✅ Validation passed: extracted totals reconcile with PDF Grand Total.")
    else:
        for w in totals_warnings:
            print(f"WARNING: {w}")
        print("Validation FAILED — see WARNING(s) above.", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
