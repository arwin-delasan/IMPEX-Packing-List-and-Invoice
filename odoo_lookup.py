#!/usr/bin/env python3
"""Look up product(s) in Odoo by product code.

Usage:
    python odoo_lookup.py <code> [code2 code3 ...]
"""

import sys

from odoo_client import OdooClient, OdooError
from categorize import classify


def main():
    if len(sys.argv) < 2:
        print("Usage: python odoo_lookup.py <code> [code2 code3 ...]", file=sys.stderr)
        sys.exit(1)

    codes = sys.argv[1:]

    try:
        odoo = OdooClient()
        products = odoo.lookup_products_by_code(codes)
    except OdooError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)

    for code in codes:
        rec = products.get(code)
        print(f"=== {code} ===")
        if not rec:
            print("  NOT FOUND in Odoo")
            print()
            continue

        categ_path = rec["categ_id"][1] if rec["categ_id"] else None
        category, subtype = classify(code, rec["name"], categ_path)

        print(f"  Name:      {rec['name']}")
        print(f"  Category:  {categ_path or '(none)'}")
        print(f"  Packing:   {category or '(unresolved)'}"
              + (f" / {subtype}" if subtype else ""))
        print()


if __name__ == "__main__":
    main()
