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
