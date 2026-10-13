# -*- coding: utf-8 -*-
"""Write the Windows VERSIONINFO resource stamped into DocPilot.exe.

Explorer's Properties > Details tab and Add/Remove Programs both read this;
without it the exe shows blank publisher fields and looks unsigned junk.
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from backend.version import LABEL, VERSION as _VERSION, parse  # noqa: E402

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "file_version.txt")
VERSION = (parse(_VERSION) + (0, 0, 0, 0))[:4]
DISPLAY = LABEL

TEMPLATE = """VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={v}, prodvers={v},
    mask=0x3f, flags=0x0, OS=0x40004, fileType=0x1, subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable('040904B0', [
        StringStruct('CompanyName', 'Physics Wallah - OnlyIAS'),
        StringStruct('FileDescription', 'DocPilot - UPSC Document Studio'),
        StringStruct('FileVersion', '{d}'),
        StringStruct('InternalName', 'DocPilot'),
        StringStruct('LegalCopyright', 'Copyright (C) Physics Wallah'),
        StringStruct('OriginalFilename', 'DocPilot.exe'),
        StringStruct('ProductName', 'DocPilot'),
        StringStruct('ProductVersion', '{d}')])
    ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"""


def main():
    with open(OUT, "w", encoding="utf-8") as fh:
        fh.write(TEMPLATE.format(v=VERSION, d=DISPLAY))
    print(f"wrote {OUT}")


if __name__ == "__main__":
    main()
