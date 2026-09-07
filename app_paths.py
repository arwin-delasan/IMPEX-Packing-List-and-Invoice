"""Resolves file paths relative to where the app itself lives, not the
current working directory.

Every reference file this app reads by a bare relative name (.env,
"PL1 - Copy.xlsx", "INV - Copy.xlsx", "BELLA VIVO PRICELIST-0125.xlsx",
the default "overrides.json") is written and tested assuming that name
resolves against the project folder - true when running `python
generate_pl1.py ...` from this directory, but not once the app is
packaged into a standalone .exe (PyInstaller): the working directory
there is wherever the user double-clicked from (Desktop, Downloads,
...), not the folder the .exe and its support files live in. app_dir()
gives the right base in both cases.
"""

import os
import sys


def app_dir():
    if getattr(sys, "frozen", False):  # running as a PyInstaller-built .exe
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def app_path(*parts):
    return os.path.join(app_dir(), *parts)


def env_value(key, env_path=None):
    """Read one KEY=value line from .env, or None if the file/key is
    missing. Lives here rather than in any one feature module because
    several of them (Odoo settings, the overrides path, the Bella Vivo
    price list path) need the same tiny parser. Strips # comments,
    surrounding quotes, and a UTF-8 BOM."""
    env_path = env_path or app_path(".env")
    if not os.path.exists(env_path):
        return None
    with open(env_path, encoding="utf-8-sig") as f:
        for line in f:
            line = line.split("#", 1)[0].strip()
            if not line or "=" not in line:
                continue
            k, v = line.split("=", 1)
            if k.strip() == key:
                return v.strip().strip('"').strip("'")
    return None
