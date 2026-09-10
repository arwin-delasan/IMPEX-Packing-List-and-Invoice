# SAWO Packing List & Invoice Generator

Turns a SAWO packing-list PDF into the categorized "PL1" packing list and the
matching commercial invoice, as Excel workbooks.

The PDF is read once. Each item is looked up in Odoo and sorted into product
categories by code/name rules (`categorize.py`), then both outputs render from
that same extracted data.

## Using the app

1. **Choose Packing List PDF...** - extracts, looks up in Odoo, classifies.
   Two PDF layouts are supported (the emailed template and Odoo's own report);
   the right reader is picked automatically.
2. **Review & Correct Classification...** - fix anything mis-categorized. A
   correction is saved per product code and reused on every later shipment.
3. **Generate Packing List...** - asks FCL or LCL. FCL totals the packages with
   a formula and labels them "packages". LCL asks how many boxes, pallets and
   crates, and that breakdown replaces the package total outright.
4. **Generate Invoice...** - asks the price source: the Bella Vivo price list,
   or a pro-forma invoice PDF for other customers.

Each finished workbook opens automatically.

## Running from source

```
pip install -r requirements.txt
python packing_list_gui.py
```

Command line, one output each:

```
python generate_pl1.py <input.pdf> <output.xlsx>
python generate_inv.py [--bella-vivo | --proforma <quote.pdf>] <input.pdf> <output.xlsx>
```

`python check_parsers.py` re-parses every sample PDF in this folder. Run it
after touching any extractor.

## Files it reads at runtime

These sit next to the app (or next to the .exe) and are **not** bundled into the
build:

| File | Purpose |
| --- | --- |
| `.env` | Odoo server/database, plus the shared paths for corrections and price lists |
| `PL1 - Copy.xlsx` | Style reference the packing list is built from |
| `INV - Copy.xlsx` | Style reference for the invoice |
| `BELLA VIVO PRICELIST-0125.xlsx` | Local fallback price list |
| `overrides.json` | Local fallback for saved classification corrections |

Corrections and price lists normally live on the shared folder named in `.env`,
so one upload reaches every install. The local copies are only used when that
share is unreachable, and the app says so in its log when it falls back.

Odoo username and password are typed at startup, never stored in `.env`.

## Building the .exe

```
python -m PyInstaller --noconfirm "SAWO Packing List Invoice Generator.spec"
```

Output lands in `dist/SAWO Packing List Invoice Generator/`. PyInstaller wipes
that folder first, so copy the five runtime files above back in beside the .exe
before testing the build.

## Updating the deployed app

Users run the Desktop shortcut "Packing List and Invoice Gen.lnk", pointing at:

```
\\172.16.0.4\Automation\IMPEX\dist\SAWO Packing List Invoice Generator\
```

To update, copy over just two things from your fresh build:

- `SAWO Packing List Invoice Generator.exe`
- the `_internal` folder

Leave the deployed `.env` and the three xlsx files alone - they are the live
configuration. Everyone must have the app closed during the copy, or Windows
will refuse to replace the running .exe.

Build from a clean working tree, so the commit you have checked out is a true
record of what shipped.
