"""Derive a packing-list Category / Subtype for a SAWO product from its
product code, name, and (if available) Odoo category path.

Classification runs in three layers, most-reliable first:

1. Code-prefix / in-code-keyword rules (SR.., NR.., GUARD, COLLAR, etc.) -
   validated exactly against a real historical packing list (34/34 match).
2. Odoo category-path rules (ODOO_CATEGORY_MAP) - Odoo's product.category
   tree turned out to be far more reliable at the top-level branch (e.g.
   "Heaters", "Sauna Rooms", "Spare Parts") than at fine subcategory
   assignment, so this layer maps whole branches by path *prefix*, mapping
   the more specific paths first. Built from the full 76-category tree
   under "All / Finished Goods" (see result of _dump_categories.py).
3. Name-keyword fallback, for the many products Odoo hasn't categorized
   beyond the bare "Finished Goods" root at all.

Any code that doesn't match a rule returns (None, None) - callers should
treat that as "needs a human to look at it", not silently bucket it.
"""

import re

# The only parent categories - fixed set, fixed order. Packaging and
# Samples used to be separate top-level categories too; folded into Sauna
# Spare Parts / Sauna Other Item respectively per explicit instruction to
# keep only these 7 as parent categories (User Interface stays top-level -
# matches the historical reference file, PL1 - Copy.xlsx, which prints
# "User Interface" as its own header).
CATEGORY_ORDER = [
    "Sauna Room",
    "Sauna Equipment",
    "User Interface",
    "Sauna Power Controller",
    "Sauna Accessories",
    "Sauna Spare Parts",
    "Sauna Other Item",
]

UNCATEGORIZED = "(Uncategorized)"

MATERIAL_WORDS = ["Stainless", "Soapstone", "Wooden"]

# subtypes that should be prefixed with the detected material (e.g. "Wooden
# Pail" vs "Stainless Pail") rather than used as a bare label.
_MATERIAL_SUBTYPES = {"Pail", "Ladle", "Backrest"}

# sentinel subtype: resolved by guess_thermo_type() rather than used as-is,
# since the "Thermo-Hygrometers" Odoo category (and the "thermometer" /
# "hygrometer" keywords) actually cover three distinct product types.
_THERMO_SENTINEL = "Thermometer/Hygrometer"


def guess_material(name):
    n = name.lower()
    found = [(n.index(m.lower()), m) for m in MATERIAL_WORDS if m.lower() in n]
    if found:
        return min(found)[1]  # leftmost-mentioned material wins
    return "Wooden"


def guess_thermo_type(name):
    """Distinguish a plain Thermometer from a plain Hygrometer from a
    combined Thermo-Hygrometer, since Odoo lumps all three under one
    "Thermo-Hygrometers" category and a blanket label would mislabel a
    single-function thermometer as a combined instrument."""
    n = name.lower()
    has_hygro = "hygrometer" in n
    has_thermo = "thermo" in n  # matches "thermo", "thermometer", "thermo°f-..."
    if has_hygro and has_thermo:
        return "Sauna Thermo-Hygrometer"
    if has_hygro:
        return "Sauna Hygrometer"
    return "Sauna Thermometer"


def _resolve_subtype(name, subtype):
    if subtype == _THERMO_SENTINEL:
        return guess_thermo_type(name)
    if subtype in _MATERIAL_SUBTYPES:
        return f"{guess_material(name)} {subtype}"
    return subtype


# Odoo category path (product.category.complete_name) -> (PL1 category,
# subtype). Checked as a PREFIX match against the product's full path, most
# specific entries FIRST so e.g. ".../Doors & Handles/Glass Doors" wins over
# the more general ".../Doors & Handles" catch-all beneath it. Order matters.
#
# Deliberately NOT mapped: the bare "All / Finished Goods" root (too
# generic to mean anything) and "All / Finished Goods / Controls" (splits
# three ways in practice - User Interface / Power Controller / Spare Parts
# Control Unit - depending on the specific product, not the category).
ODOO_CATEGORY_MAP = [
    ("All / Finished Goods / Spare Parts", "Sauna Spare Parts", None),
    ("All / Finished Goods / Sauna Rooms", "Sauna Room", None),
    ("All / Finished Goods / Heaters", "Sauna Equipment", "Sauna Heater"),
    ("All / Finished Goods / Steam", "Sauna Equipment", "Sauna Heater"),
    ("All / Finished Goods / Controls / Additionals", "Sauna Spare Parts", "Control Additional"),
    ("All / Finished Goods / Marketing Material", "Sauna Accessories", None),
    ("All / Finished Goods / Packaging", "Sauna Spare Parts", "Packaging"),
    ("All / Finished Goods / Panels & Timbers", "Sauna Accessories", "Wood Panels"),
    ("All / Finished Goods / Samples", "Sauna Other Item", "Samples"),

    # --- Accessories subtree (most specific first) ---
    ("All / Finished Goods / Accessories / Doors & Handles / Glass Doors", "Sauna Accessories", "Glass Door"),
    ("All / Finished Goods / Accessories / Doors & Handles / Glass Window", "Sauna Accessories", "Glass Window"),
    ("All / Finished Goods / Accessories / Doors & Handles / Wooden Doors", "Sauna Accessories", "Sauna Door"),
    ("All / Finished Goods / Accessories / Doors & Handles / Wooden Handles", "Sauna Accessories", "Door Handle"),
    ("All / Finished Goods / Accessories / Doors & Handles", "Sauna Accessories", "Door"),
    ("All / Finished Goods / Accessories / Guards & Collars", "Sauna Accessories", "Heater Guard"),
    ("All / Finished Goods / Accessories / Head-& Backrests", "Sauna Accessories", "Backrest"),
    ("All / Finished Goods / Accessories / Lights & Covers", "Sauna Accessories", "Sauna Light Cover"),
    ("All / Finished Goods / Accessories / Benches", "Sauna Accessories", "Sauna Bench"),
    ("All / Finished Goods / Accessories / Clocks", "Sauna Accessories", "Wooden Clock"),
    ("All / Finished Goods / Accessories / Aroma Oil", "Sauna Other Item", "Aroma Scents"),
    ("All / Finished Goods / Accessories / Decal", "Sauna Accessories", "Decal"),
    ("All / Finished Goods / Accessories / Infrared", "Sauna Accessories", "Infrared"),
    ("All / Finished Goods / Accessories / Miscellaneous / Signs", "Sauna Accessories", "Sign"),
    ("All / Finished Goods / Accessories / Miscellaneous / Speakers", "Sauna Accessories", "Speaker"),
    ("All / Finished Goods / Accessories / Miscellaneous", "Sauna Accessories", None),
    ("All / Finished Goods / Accessories / Soapstone", "Sauna Accessories", "Soapstone"),
    ("All / Finished Goods / Accessories / Pails", "Sauna Accessories", "Pail"),
    ("All / Finished Goods / Accessories / Ladles", "Sauna Accessories", "Ladle"),
    ("All / Finished Goods / Accessories / Sandtimers", "Sauna Accessories", "Sandtimer"),
    ("All / Finished Goods / Accessories / Thermo-Hygrometers", "Sauna Accessories", _THERMO_SENTINEL),
    ("All / Finished Goods / Accessories / Controls", "Sauna Accessories", "Controls"),
    ("All / Finished Goods / Accessories / Heaters", "Sauna Accessories", "Heater Accessory"),
    ("All / Finished Goods / Accessories", "Sauna Accessories", None),
]

# Heater product-line/brand names, always the first word of the product
# name (e.g. "Aries Wall 3H.E 9,0kW NS Premium"). Pulled from every distinct
# first-word among products actually filed in Odoo's Heaters category tree,
# then checked catalog-wide to drop anything that also appears as the first
# word of an unrelated (non-heater) product. "Holl's" was dropped - it's a
# gift-line brand spanning pails/ladles/benches/etc, not heater-specific.
# "Steam" was dropped too - it collides with Glass Doors/Controls products;
# steam equipment is matched by the more specific "steam gen" rule above.
_HEATER_BRAND_WORDS = {
    "nordex", "tower", "aries", "scandia", "sawo30", "mini", "super",
    "savonia", "cirrus", "krios", "cumulus", "taurus", "cubos",
    "scandifire", "phoenix", "minidragon", "fiberjungle", "nimbus",
    "helius", "orion", "tpi",
}

_NAME_RULES = [
    ("sandtimer", "Sauna Accessories", "Sandtimer"),
    ("hook rack", "Sauna Accessories", "Wooden Rack"),
    ("hygrometer", "Sauna Accessories", _THERMO_SENTINEL),
    ("thermometer", "Sauna Accessories", _THERMO_SENTINEL),
    ("backrest", "Sauna Accessories", "Backrest"),
    ("display", "Sauna Accessories", "Display"),
    ("scent", "Sauna Accessories", "Scent"),
    ("curve light", "Sauna Accessories", "Sauna Light Cover"),
    ("light cover", "Sauna Accessories", "Sauna Light Cover"),
    ("power controller", "Sauna Power Controller", None),
    ("user interface", "User Interface", None),
    ("control unit", "Sauna Spare Parts", "Control Unit"),
    ("glass door", "Sauna Accessories", "Glass Door"),
    ("door", "Sauna Accessories", "Door"),
    ("glass window", "Sauna Accessories", "Glass Window"),
    ("handle", "Sauna Accessories", "Door Handle"),
    ("towel", "Sauna Accessories", "Wood Cloth Hanger"),
    ("bench", "Sauna Accessories", "Sauna Bench"),
    ("clock", "Sauna Accessories", "Wooden Clock"),
    ("light", "Sauna Accessories", "Sauna Light Cover"),
    ("trendline", "Sauna Equipment", "Sauna Heater"),
]


def classify(code, name, categ_path=None):
    """Return (category, subtype) for a product, or (None, None) if no
    rule matches. subtype may be None even when category is resolved
    (e.g. Sauna Rooms have no subtype)."""
    c = code.upper()
    n = name.lower()

    # Accessory keyword overrides, checked before family rules since e.g.
    # "NRNS-GUARD-W-D" starts like a Nordex heater code but is an accessory.
    if "GUARD" in c:
        return "Sauna Accessories", f"{guess_material(name)} Heater Guard"
    if "COLLAR" in c:
        return "Sauna Accessories", f"{guess_material(name)} Integration Collar"

    if re.match(r"^SR\d", c):
        return "Sauna Room", None
    if re.match(r"^(NR[A-Z]*)-?\d", c):  # NRM-36.., NRNS-80.., NR-45.., NRX-.. Nordex heater models
        return "Sauna Equipment", "Sauna Heater"
    if re.match(r"^TH\d", c):  # TH5-90.. Tower Heater model
        return "Sauna Equipment", "Sauna Heater"
    if re.match(r"^SW\d", c):  # SW3-60.. Wall heater model
        return "Sauna Equipment", "Sauna Heater"
    if c.startswith("INC-"):
        return "User Interface", None
    if c.startswith("INP-"):
        return "Sauna Power Controller", None
    if c.startswith(("INT-", "INS-")):
        return "User Interface", None
    if c.startswith("INN-IH"):
        return "Sauna Accessories", "Interface Holder"
    if c.startswith("LED-SUPPLY"):
        return "Sauna Spare Parts", "Control Unit"
    if c.startswith("HIM"):
        return "Sauna Accessories", "Himalayan Salt"
    if re.match(r"^LP\d", c):
        if "wire" in n or "splitter" in n:
            return "Sauna Spare Parts", "Wire / Cable"
        if "transformer" in n or "plug" in n:
            return "Sauna Spare Parts", "Transformer"
        return "Sauna Spare Parts", None
    if c.startswith("SP02-") and "moisture paper" not in n:
        # SP02- is used exclusively for raw panel/timber stock (295 Odoo
        # SKUs checked; the one exception, "Moisture Paper", is excluded
        # above and caught by the moisture-paper rule below instead). Many
        # of these sit at Odoo's bare "Finished Goods" root with no
        # category, or get miscategorized under generic "Spare Parts" -
        # the product code is a far more reliable signal here than either.
        # Wood Panels is a subtype under Sauna Accessories, not its own
        # top-level category.
        return "Sauna Accessories", "Wood Panels"

    # High-priority name rules: checked before Odoo's own category, since
    # Odoo frequently files these under a generic bucket (or not at all)
    # regardless of what the product actually is.
    if "heater guard" in n:  # e.g. "Wooden Heater Guard for DRFT12" with no GUARD in the code
        return "Sauna Accessories", f"{guess_material(name)} Heater Guard"
    if "heater" in n:  # heaters always sit under Sauna Equipment
        return "Sauna Equipment", "Sauna Heater"
    if "steam gen" in n:  # e.g. "Steam Gen. Next Series 12,0kW 3P Australia"
        # Steam generators are heaters - Sauna Equipment / Sauna Heater,
        # same as any other heater, not a category of their own. Matching
        # "steam gen" specifically (not bare "steam") avoids sweeping in
        # unrelated products like "Steamwater" pails, "Steamshot" ladles,
        # or the "Protective Steam Head Cover" spare part.
        return "Sauna Equipment", "Sauna Heater"
    if n.startswith("spare parts for"):
        # e.g. "Spare Parts for SR02-45620192 - Panel, Horizontal & Vertical
        # Panel, Spruce" - checked catalog-wide (119 SKUs): always a Sauna
        # Spare Parts item regardless of what it's a spare part *for*.
        return "Sauna Spare Parts", None
    if "heating element" in n:
        return "Sauna Spare Parts", None
    if "contactor unit" in n:  # e.g. "Saunova 2.0 Contactor Unit" (not always INP- coded)
        return "Sauna Power Controller", None
    if "sauna room" in n:  # catches room codes that don't use the SR.. prefix
        return "Sauna Room", None
    if "accessory set" in n:
        return "Sauna Accessories", "Wooden Pail"
    if "pail shower" in n:
        return "Sauna Accessories", "Pail Shower"
    if "wooden cover" in n:
        return "Sauna Accessories", "Wooden Pail Cover"
    if "ladle" in n:
        # Odoo sometimes misfiles ladles under an unrelated category (e.g.
        # "Sandtimers" for one real SKU) - trust the name over Odoo here.
        return "Sauna Accessories", f"{guess_material(name)} Ladle"
    if "pail" in n and "pail-look" not in n and "pail look" not in n:
        # "Pail-look" is a clock shape descriptor (e.g. "Wooden Clock
        # Pail-look..."), not an actual pail - let it fall through to the
        # "clock" rule below instead.
        return "Sauna Accessories", f"{guess_material(name)} Pail"
    if "cube door" in n:
        return "Sauna Accessories", "Glass Door"
    if "sensor" in n and "holder" in n:
        return "Sauna Accessories", "Interface Holder"
    if "display stand" in n:
        return "Sauna Accessories", "Wooden Display Stand"
    if "sauna stone" in n:
        return "Sauna Other Item", "Sauna Stones"
    if "sauna guidelines" in n:
        return "Sauna Other Item", None
    if "signage" in n:
        return "Sauna Other Item", None
    if "aroma oil" in n:
        return "Sauna Other Item", "Aroma Scents"
    if "moisture paper" in n:
        return "Sauna Other Item", None
    if n.startswith("protective box") or "protective steam head cover" in n:
        # e.g. "Protective Box - SAV-120N", "Protective Box - STN-75-C1/3AS",
        # "Protective Steam Head Cover" - shipping/protective boxes for
        # heaters and steam generators. All grouped under one "Steam Head
        # Cover" subtype header historically, regardless of which specific
        # heater/generator model they're for.
        return "Sauna Spare Parts", "Steam Head Cover"
    if "carton box" in n or "carboard box" in n:
        # e.g. "Carton Box for ECOT 3H.E", "Carboard Box (Cirrus-2HE)" -
        # heater shipping cartons. Matches the "Packaging" category their
        # already-categorized sibling SKUs overwhelmingly use in Odoo
        # (folded under Sauna Spare Parts - see CATEGORY_ORDER).
        return "Sauna Spare Parts", "Packaging"
    if n.startswith("tag for"):
        # e.g. "Tag for Aries Round", "Tag for STN" - matches the
        # "Marketing Material" category most of these already have in
        # Odoo (mapped to Sauna Accessories - see ODOO_CATEGORY_MAP).
        return "Sauna Accessories", None
    if (
        "reflector" in n
        or "reflection sheet" in n
        or "rock container" in n
        or "rock containe" in n  # typo in some real Odoo SKUs, e.g. "Rock Containe Assy 6HE"
        or "catch pan" in n
        or "wire set" in n
        or "bottom cover" in n
        or "top cover" in n
        or "side cover" in n
    ):
        # Heater/steam-generator internal spare parts (drip pans, wiring
        # harnesses, stone covers, reflectors...). Each phrase checked
        # catalog-wide: consistently Spare Parts wherever Odoo has them
        # categorized at all.
        return "Sauna Spare Parts", None
    if "small glass" in n:
        # e.g. "Door 700x2040mm with Small Glass Window Clear, Cedar Left" -
        # Odoo files these under "Wooden Doors", but they're Glass Doors.
        return "Sauna Accessories", "Glass Door"
    if "floor mat" in n:
        return "Sauna Accessories", "Wooden Floormat"
    if "pillow" in n:
        return "Sauna Accessories", "Wooden Pillow"
    if "ventilation" in n:
        return "Sauna Accessories", "Sauna Air Ventilation Louver"
    if re.match(r"^R-\d", c):
        # Stone-named items were already routed to Sauna Stones above; the
        # rest of this code family (cups, coolers, candle holders) is the
        # Aroma Cup line.
        return "Sauna Accessories", "Aroma Cup"

    if categ_path:
        for prefix, categ, sub in ODOO_CATEGORY_MAP:
            if categ_path.startswith(prefix):
                return categ, _resolve_subtype(name, sub)

    for kw, categ, sub in _NAME_RULES:
        if re.search(rf"\b{re.escape(kw)}\b", n):  # word-boundary: "door" shouldn't match "outdoor"
            return categ, _resolve_subtype(name, sub)

    # Last-resort fallback: recognize a heater by its product-line/brand
    # name even when the code prefix is nonstandard AND Odoo's own category
    # is missing or wrong - both real, observed failure modes (e.g. Odoo
    # had "Mini 3,6kW NB Premium Australia" filed at the bare "Finished
    # Goods" root with no category, and a "Mini 2,3kW..." heater filed
    # under "Wooden Cover" by mistake). Mined from every product actually
    # in Odoo's Heaters category tree: these are the brand names heater
    # model names always start with, regardless of code/category. Checked
    # dead last, since it's a coarser signal than everything above it -
    # e.g. a heater spare part like "Box - TH12-210NS-P Tower Heater..."
    # doesn't start with the brand name, so it's not caught here (it's
    # correctly caught earlier as a Spare Part instead, if categorized).
    first_word = n.split()[0].rstrip(",.") if n.split() else ""
    if first_word in _HEATER_BRAND_WORDS and "cover" not in n and "assy" not in n:
        # Excludes a handful of real Odoo SKUs like "Aries Tower Heater 3HE
        # - Outer Cover" or "Super Nordex Floor Standing Outer Cover Assy" -
        # spare/replacement covers named after the heater they fit, not the
        # heater itself.
        return "Sauna Equipment", "Sauna Heater"

    return None, None


def classify_subgroup(category, subtype, code, name):
    """Key used to decide where Sub-Total rows break *within* one printed
    subtype header/category, for products that shouldn't all be summed
    together under one Sub-Total despite sharing a header:

    - "Wire / Cable" covers both actual wire (LP15-005/006) and three-way
      splitters (LP15-002); same header, separate Sub-Totals.
    - "Sauna Other Item" collects a grab-bag of otherwise-unrelated
      products with no specific subtype (Signage, Sauna Guidelines, ...).
      Each distinct product code gets its own Sub-Total there - grouped
      only with other rows of the *same* code (repeated shipment lines for
      one product), never merged with a different product just because
      neither has a subtype.
    Defaults to the subtype itself (one Sub-Total per header, the normal
    case for every other category, e.g. Sauna Room stays one block)."""
    if subtype is None and category == "Sauna Other Item":
        return code
    if subtype == "Wire / Cable" and "splitter" in name.lower():
        return f"{subtype} :: Splitter"
    return subtype


def category_sort_key(category):
    try:
        return CATEGORY_ORDER.index(category)
    except ValueError:
        return len(CATEGORY_ORDER)  # unknown/uncategorized sorts last


# Subtypes listed here sort first within their category, in this order;
# any subtype not listed keeps the default first-appearance-in-PDF order,
# placed after all listed ones. E.g. "Sauna Heater" is the core Sauna
# Equipment item, so it shouldn't be bumped out of first place just
# because a display stand happened to appear earlier in the PDF.
SUBTYPE_ORDER = {
    "Sauna Equipment": ["Sauna Heater"],
}


def subtype_sort_key(category, subtype):
    try:
        return SUBTYPE_ORDER.get(category, []).index(subtype)
    except ValueError:
        return len(SUBTYPE_ORDER.get(category, []))
