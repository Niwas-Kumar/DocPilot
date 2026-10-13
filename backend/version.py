# -*- coding: utf-8 -*-
"""The one place DocPilot's version lives.

The status window, the exe's Properties > Details, the installer and the
in-app updater all read it from here, so a release is: bump VERSION, run
tools/build_installer.py, upload the setup and update the manifest.

VERSION must stay dotted numbers ("2.1.0"): the updater compares it number by
number with the manifest's. LABEL is only what people see.
"""

VERSION = "2.1.3"
LABEL = VERSION + "-beta"

# The update manifest: a Google Doc (or a plain text / JSON file on Drive)
# shared as "Anyone with the link can view". Its link never changes - each
# release only edits its text. Empty = updates switched off.
UPDATE_MANIFEST = ("https://docs.google.com/document/d/"
                   "1wFuTqYSiIPhss0QE8tI5tYY1emX3Oluv7t0CJPxVB_8/edit")


def parse(version):
    """'2.1.0' -> (2, 1, 0); anything after the numbers ('-beta') is ignored."""
    import re
    nums = re.findall(r"\d+", (version or "").split("-")[0])
    return tuple(int(n) for n in nums) or (0,)
