r"""Persistent per-product-code classification corrections.

classify() in categorize.py guesses each product's Category/Subtype from
its code/name, since Odoo's own categorization has proven unreliable
throughout this project. This module lets a human correct a wrong guess
once and have it remembered - the corrected code is never misclassified
again on a future run.

Stored as a flat {code: entry} dict in a JSON file. Deliberately per-code
only (no name-pattern/regex rules) - a bad pattern could silently
misclassify a whole family of unrelated products (this project has hit
that exact failure mode more than once, e.g. "door" matching inside
"Outdoor"), whereas a per-code entry can only ever affect the one code
it was written for.

By default the file lives next to this script, so it's local to one
machine. To make it a shared single source of truth across everyone's
install (e.g. a LAN path like \\172.16.0.4\Shared\overrides.json), add
an OVERRIDES_PATH line to .env - same file/format already used for the
Odoo connection settings. Every install pointed at the same path reads
and writes the same file, so a correction one person saves is picked up
by everyone else's next run. There's no locking, so two people saving a
correction at the exact same moment could clobber each other's write
(last one wins) - a real but low-probability risk for occasional,
one-at-a-time corrections; not an issue worth a distributed lock for at
this scale.
"""

import json
import os
from datetime import datetime

from app_paths import app_path


def _env_value(key, env_path=None):
    """Read one KEY=value line from .env, or None if the file/key is
    missing. Mirrors odoo_client.py's _load_env() parsing (# comments
    stripped, surrounding quotes stripped) without importing an
    Odoo-specific module for a generic concern."""
    env_path = env_path or app_path(".env")
    if not os.path.exists(env_path):
        return None
    with open(env_path, encoding="utf-8-sig") as f:  # -sig: strip a UTF-8 BOM if present
        for line in f:
            line = line.split("#", 1)[0].strip()
            if not line or "=" not in line:
                continue
            k, v = line.split("=", 1)
            if k.strip() == key:
                return v.strip().strip('"').strip("'")
    return None


def _resolve_overrides_path():
    path = _env_value("OVERRIDES_PATH") or app_path("overrides.json")
    if os.path.isdir(path):
        # A shared *folder* was given (e.g. "\\172.16.0.4\Shared") rather
        # than a full file path - store the file inside it instead of
        # failing. Windows raises PermissionError (not IsADirectoryError)
        # when you try to open() a directory as a file, which is what led
        # here in the first place.
        path = os.path.join(path, "overrides.json")
    return path


OVERRIDES_PATH = _resolve_overrides_path()


class OverridesError(Exception):
    """Raised when the overrides file exists but can't be read as JSON."""


def load_overrides(path=OVERRIDES_PATH):
    """Return {code: entry}, or {} if the file doesn't exist yet (first
    run - nothing has been corrected)."""
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError) as e:
        # Fail loud rather than silently discarding saved corrections -
        # matches ExtractionError/GenerationError's fail-fast philosophy
        # elsewhere in this codebase.
        raise OverridesError(f"Could not read {path}: {e}")


def save_overrides(overrides, path=OVERRIDES_PATH):
    """Atomic write: json.dump to a temp file, then os.replace() over the
    real path, so a crash mid-write can never corrupt existing corrections."""
    tmp_path = f"{path}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(overrides, f, indent=2, ensure_ascii=False, sort_keys=True)
    os.replace(tmp_path, path)


def get_override(code, overrides):
    """Pure lookup against an already-loaded dict - classify_items() loads
    the file once per run, not once per line item."""
    return overrides.get(code)


def set_overrides_bulk(codes, category, subtype, product_names, originals, path=OVERRIDES_PATH):
    """Upsert one entry per code (already deduped by the caller) in a
    single load/save round trip, and return the updated dict.

    product_names: {code: product name}, for audit/display in the Manage
    Corrections view.
    originals: {code: (original_category, original_subtype)} - the
    category/subtype this code had *before* this correction (whatever
    classify() produced, or a prior override's value if it was already
    overridden once).

    `category` doesn't have to be one of CATEGORY_ORDER - a brand new
    category is allowed, for cases the existing rules/categories don't
    cover yet. build_blocks()/category_sort_key() already sort anything
    not in CATEGORY_ORDER after the known ones, so a new category still
    prints its own header/Sub-Total correctly, just always last."""
    if not category or not category.strip():
        raise ValueError("Category can't be blank.")

    overrides = load_overrides(path)
    now = datetime.now().isoformat(timespec="seconds")
    for code in codes:
        orig_category, orig_subtype = originals.get(code, (None, None))
        overrides[code] = {
            "category": category,
            "subtype": subtype,
            "original_category": orig_category,
            "original_subtype": orig_subtype,
            "product_name": product_names.get(code, ""),
            "corrected_at": now,
        }
    save_overrides(overrides, path)
    return overrides


def delete_override(code, path=OVERRIDES_PATH):
    """Remove one saved correction, if present. Returns the updated dict."""
    overrides = load_overrides(path)
    overrides.pop(code, None)
    save_overrides(overrides, path)
    return overrides
