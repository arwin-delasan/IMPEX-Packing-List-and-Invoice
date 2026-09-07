r"""Swappable Bella Vivo price list: which .xlsx is currently in force.

The Bella Vivo price list is reissued periodically (the original was
"BELLA VIVO PRICELIST-0125.xlsx" - 0125 being its January 2025 edition).
Hard-coding that filename meant a new edition could only be adopted by
editing the source, so this module keeps the *active* file's name in a
small JSON pointer instead, and lets the GUI install a new edition.

Adopting a new edition never deletes or overwrites the old one - only the
pointer moves. The superseded .xlsx stays on disk, so an invoice priced
last month can still be reconciled against the list it was actually
priced from, and a bad edition can be rolled back by pointing at the
previous file again.

WHERE IT LIVES
--------------
Same shared-folder model as overrides.py, and for the same reason: more
than one person runs this app. If each install kept its own price list,
one person adopting a new edition would leave everyone else quoting the
superseded prices with no indication anything had changed - the failure
would surface as a wrong invoice total, not as an error. Pointing
PRICELIST_PATH at the same LAN folder as OVERRIDES_PATH means one upload
reaches everyone on their next run. PRICELIST_PATH is optional: with no
such line, the folder already configured as OVERRIDES_PATH is used, so an
install that shares corrections shares price lists too without anyone
having to remember a second .env edit on every machine.

If neither is set, or the share is unreachable (laptop off the network,
VPN down), resolve_bella_vivo_path() falls back to whatever
.xlsx sits next to the app - the original bundled behavior - rather than
failing the invoice outright. It reports which of the two it used so the
GUI can say so out loud; silently pricing from a stale local copy is
exactly the failure this module exists to prevent.
"""

import json
import os
import shutil
from datetime import datetime

from app_paths import app_path, env_value

POINTER_FILENAME = "pricelist.json"
BELLA_VIVO_KEY = "bella_vivo"

# The edition shipped alongside the app, used until a newer one is
# installed (and as the fallback when the shared folder is unreachable).
DEFAULT_BELLA_VIVO_FILENAME = "BELLA VIVO PRICELIST-0125.xlsx"


class PriceListError(Exception):
    """Raised when a price list file or its pointer can't be read, or when
    a file being installed doesn't look like a Bella Vivo price list."""


def shared_dir():
    """The folder holding the pointer and the price list files, or None
    when PRICELIST_PATH isn't configured. Accepts either a folder or a
    full path to the pointer file, mirroring overrides.py's tolerance for
    both - a user pasting a LAN path is as likely to give one as the
    other."""
    path = env_value("PRICELIST_PATH")
    if not path:
        # .env is per-machine and not in version control, so requiring a
        # PRICELIST_PATH line on every install would mean one forgotten
        # edit silently puts that machine back on its own local copy -
        # the precise failure the shared folder exists to prevent. Every
        # install that shares corrections has already been pointed at the
        # team folder via OVERRIDES_PATH, so reuse it: the price list and
        # the corrections belong in the same place anyway. Callers still
        # report which file was used, so this is never silent.
        path = env_value("OVERRIDES_PATH")
    if not path:
        return None
    if path.lower().endswith(".json"):
        path = os.path.dirname(path)
    return path or None


def _pointer_path(directory):
    return os.path.join(directory, POINTER_FILENAME)


def _read_pointer(directory):
    path = _pointer_path(directory)
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError) as e:
        raise PriceListError(f"Could not read {path}: {e}")


def _write_pointer(directory, data):
    """Atomic write, same approach as overrides.save_overrides(): write a
    temp file then os.replace() over the real one, so a crash or a dropped
    network connection mid-write can't leave a half-written pointer that
    makes the price list unresolvable for everyone."""
    tmp_path = _pointer_path(directory) + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False, sort_keys=True)
    os.replace(tmp_path, _pointer_path(directory))


def resolve_bella_vivo_path():
    """Return (path, source) for the price list currently in force.

    source is "shared" when it came from the configured shared folder and
    "local" when it fell back to the file next to the app - the caller is
    expected to surface that distinction to the user.
    """
    directory = shared_dir()
    if directory:
        try:
            entry = _read_pointer(directory).get(BELLA_VIVO_KEY)
            if entry:
                candidate = os.path.join(directory, entry["filename"])
                if os.path.exists(candidate):
                    return candidate, "shared"
        except (PriceListError, OSError, KeyError, TypeError):
            # Unreachable share, unreadable pointer, or a pointer written
            # by a newer version with a shape this one doesn't understand
            # - all mean "fall back to local", never "fail the invoice".
            pass
    return app_path(DEFAULT_BELLA_VIVO_FILENAME), "local"


def active_entry():
    """The pointer's record for the active price list (filename, when it
    was installed, by whom), or None if no edition has been installed
    yet - i.e. everyone is still on the bundled default."""
    directory = shared_dir()
    if not directory:
        return None
    try:
        return _read_pointer(directory).get(BELLA_VIVO_KEY)
    except PriceListError:
        return None


def validate_pricelist(path, load_prices):
    """Check a candidate .xlsx really is a Bella Vivo price list before
    anything is copied or the pointer is moved.

    load_prices is generate_inv.load_bella_vivo_prices, passed in rather
    than imported to keep this module free of a circular import. A file
    that opens but yields no codes (wrong sheet, wrong columns, an export
    with a shifted header row) would otherwise be adopted happily and
    then silently blank out every Unit Price on the next invoice.
    Returns the parsed {code: price} so the caller can show a count.
    """
    try:
        prices = load_prices(path)
    except Exception as e:
        raise PriceListError(
            f"Doesn't look like a Bella Vivo price list - couldn't read it: {e}"
        )
    if not prices:
        raise PriceListError(
            "Doesn't look like a Bella Vivo price list - no product codes found. "
            'Expected a sheet named "data" with codes in column B and prices in column C.'
        )
    return prices


def install_bella_vivo(src_path, load_prices, installed_by=""):
    """Validate src_path, copy it into the shared folder, and point the
    active price list at it. Returns (destination_path, price_count).

    The copy keeps the uploaded file's own name (editions are already
    dated in it, e.g. "-0125"), but never overwrites an existing file: a
    name that's already taken gets a numeric suffix instead, since
    clobbering a superseded edition would destroy the audit trail this
    module exists to preserve.
    """
    directory = shared_dir()
    if not directory:
        raise PriceListError(
            "No shared folder is configured. Add a PRICELIST_PATH (or "
            "OVERRIDES_PATH) line to .env pointing at the folder everyone's "
            "install should read from."
        )
    if not os.path.isdir(directory):
        raise PriceListError(
            f"Shared price list folder is not reachable:\n{directory}\n\n"
            "Check the network connection, then try again."
        )

    prices = validate_pricelist(src_path, load_prices)

    basename = os.path.basename(src_path)
    stem, ext = os.path.splitext(basename)
    dest = os.path.join(directory, basename)
    counter = 2
    while os.path.exists(dest) and not os.path.samefile(src_path, dest):
        dest = os.path.join(directory, f"{stem} ({counter}){ext}")
        counter += 1

    if not (os.path.exists(dest) and os.path.samefile(src_path, dest)):
        shutil.copy2(src_path, dest)

    pointer = _read_pointer(directory)
    previous = pointer.get(BELLA_VIVO_KEY)
    pointer[BELLA_VIVO_KEY] = {
        "filename": os.path.basename(dest),
        "installed_at": datetime.now().isoformat(timespec="seconds"),
        "installed_by": installed_by,
        "product_count": len(prices),
        # Kept so the superseded edition is still named somewhere even
        # after the pointer has moved on.
        "supersedes": previous.get("filename") if previous else None,
    }
    _write_pointer(directory, pointer)
    return dest, len(prices)
