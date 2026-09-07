import sys, os, json, pdfplumber
sys.path.insert(0, os.getcwd())
import odoo_pdf_parser as O
from generate_pl1 import validate
from categorize import classify

pdf = pdfplumber.open('China 09-04-2026 Packing list REV.1.pdf')
print("=== HEADER ===")
try:
    print(json.dumps(O.parse_header(pdf.pages[0]), indent=1))
except Exception as e:
    print("HEADER FAIL:", e)

items=[]; gtw=None
for page in pdf.pages:
    rows,g = O.parse_items_on_page(page)
    for w in rows:
        try: items.append(O.classify_item_row(w))
        except Exception: pass
    if g is not None: gtw=g; break
print("\n=== GRAND TOTAL ===")
gt={"packages":442,"net_weight":3707.01,"gross_weight":4456.19,"cbm":30.932,"qty":1898}; print("hardcoded from PDF:",gt)
print("\n=== VALIDATE (with 8 rows already dropped) ===")
for w in validate(items, gt): print(" ", w)

print("\n=== CLASSIFY (no Odoo; code+PDF desc only) ===")
un=0
for it in items:
    c,s = classify(it["external_code"], it["description"])
    if c is None:
        un+=1; print("  UNCLASSIFIED:", it["external_code"], "|", it["description"][:55])
print("unclassified:",un,"of",len(items))
