#!/usr/bin/env python3
"""Regression check: parse every sample packing list and prove it still reads.

Touching the extractors is the riskiest edit in this project - the
templates only differ in ways that are invisible until a row lands in
the wrong column, and there are six real Email-template files and three
Odoo ones sitting next to this script that between them cover the
awkward cases (multi-page lists, "Part of Crate" rows, descriptions that
overflow into the numeric columns, a product code containing a space).
Running them all takes a second and answers the question that actually
matters after a parser change: did I break any of the other formats?

Two things are checked per file:

1. Totals reconcile - the summed line items match the PDF's own Grand
   Total row. This is the real test. A misparse that shifts a column
   still produces numbers, and this is what catches it.
2. The snapshot still matches - detected template, item count and totals
   are compared against check_parsers.json, so a change that alters
   *what* is extracted has to be looked at and re-blessed deliberately
   rather than passing silently.

    python check_parsers.py           # check against the snapshot
    python check_parsers.py --bless   # re-record it after an intended change

The PDFs themselves are gitignored (they're customer documents), so this
runs against whatever samples are in the working folder and skips the
snapshot entries it can't find.
"""

import glob
import json
import sys

from app_paths import app_path
from generate_pl1 import validate
from packing_list_pdf import extract_pdf
from pdf_parser import ExtractionError

SNAPSHOT_PATH = app_path("check_parsers.json")


def summarise(pdf_path):
    header, items, grand_total, template = extract_pdf(pdf_path)
    return {
        "template": template,
        "items": len(items),
        "sold_to_company": header["sold_to_company"],
        "grand_total": grand_total,
    }


def main():
    bless = "--bless" in sys.argv[1:]

    try:
        with open(SNAPSHOT_PATH, encoding="utf-8") as f:
            snapshot = json.load(f)
    except FileNotFoundError:
        snapshot = {}

    pdfs = sorted(glob.glob(app_path("*.pdf")))
    recorded = {}
    failures = []

    for pdf_path in pdfs:
        name = pdf_path.rsplit("\\", 1)[-1].rsplit("/", 1)[-1]
        try:
            header, items, grand_total, template = extract_pdf(pdf_path)
        except ExtractionError:
            # Not a packing list (pro-forma invoices and the scanned file
            # live in the same folder) - only a file the snapshot claims
            # used to parse is a regression.
            if name in snapshot:
                failures.append(f"{name}: no longer parses as a packing list")
            continue

        current = {
            "template": template,
            "items": len(items),
            "sold_to_company": header["sold_to_company"],
            "grand_total": grand_total,
        }
        recorded[name] = current

        totals_warnings = validate(items, grand_total)
        for w in totals_warnings:
            failures.append(f"{name}: {w}")

        expected = snapshot.get(name)
        if expected is None:
            print(f"NEW      {name}: {template}, {len(items)} items (not in snapshot)")
        elif expected != current:
            failures.append(f"{name}: snapshot mismatch\n    was {expected}\n    now {current}")
        else:
            status = "OK" if not totals_warnings else "TOTALS"
            print(f"{status:<8} {name}: {template}, {len(items)} items")

    missing = [n for n in snapshot if n not in recorded]
    for n in missing:
        print(f"SKIPPED  {n}: not in this folder")

    if bless:
        with open(SNAPSHOT_PATH, "w", encoding="utf-8") as f:
            json.dump(recorded, f, indent=2, sort_keys=True)
            f.write("\n")
        print(f"\nSnapshot re-recorded: {len(recorded)} file(s) -> {SNAPSHOT_PATH}")
        return 0

    if failures:
        print("\nFAILED:")
        for f in failures:
            print(f"  {f}")
        return 1

    print(f"\nAll {len(recorded)} packing list(s) parse and reconcile.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
