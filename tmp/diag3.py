import sys, os, json; sys.path.insert(0, os.getcwd())
from pdf_parser import extract_pdf
from generate_pl1 import validate, um_warnings
from categorize import classify

h,items,gt = extract_pdf('China 09-04-2026 Packing list REV.1.pdf')
print("=== HEADER ==="); print(json.dumps(h,indent=1))
print("\n=== VALIDATE ===")
w=validate(items,gt); print("  PASS" if not w else "\n".join("  "+x for x in w))
print("\n=== UM WARNINGS ==="); print(um_warnings(items) or "  none")
print("\n=== UM values seen ===", sorted({i['um'] for i in items}))
print("\n=== suspicious codes (len>18 or lowercase run) ===")
for i in items:
    c=i["external_code"]
    if len(c)>18 or any(ch.islower() for ch in c[1:]): print("  ",repr(c),"|",i["description"][:45])
print("\n=== the 8 rows that broke the Odoo parser ===")
for i in items:
    if i["external_code"] in ("741-4SGD-3","LED-FRAME-1") or i["external_code"].startswith("STN-"):
        print(f"  {i['external_code']:<20} um={i['um']:<4} pkg={i['packages']} net={i['net_weight']} gross={i['gross_weight']} cbm={i['cbm']} qty={i['qty']}")
        print(f"      desc={i['description']!r}")
print("\n=== unclassified (no Odoo) ===")
un=[i for i in items if classify(i["external_code"],i["description"])[0] is None]
for i in un: print("  ",i["external_code"],"|",i["description"][:55])
print("count:",len(un),"of",len(items))
