# -*- coding: utf-8 -*-
"""Build DocPilot-Setup-<version>.exe end to end.

    python tools/build_installer.py            # icon + app bundle + installer
    python tools/build_installer.py --skip-app # installer only, reuse dist/

Leaves the finished setup at dist/DocPilot-Setup-<version>.exe and prints the
lines to paste into the update notice on Drive (see backend/updater.py).
The version comes from backend/version.py - bump it there first.
"""
import hashlib
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from backend.version import LABEL, VERSION  # noqa: E402

BUNDLE = os.path.join(ROOT, "dist", f"DocPilot-{LABEL}")

# Inno Setup installs per-user or machine-wide depending on who ran it, so
# look in both before giving up
ISCC_CANDIDATES = [
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Inno Setup 6", "ISCC.exe"),
    r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    r"C:\Program Files\Inno Setup 6\ISCC.exe",
]


def find_iscc():
    for path in ISCC_CANDIDATES:
        if path and os.path.isfile(path):
            return path
    sys.exit("Inno Setup 6 not found. Install it from https://jrsoftware.org/isdl.php\n"
             "(the per-user option is enough - no admin needed)")


def run(cmd, **kw):
    print(f"\n>>> {' '.join(str(c) for c in cmd)}")
    if subprocess.call(cmd, cwd=ROOT, **kw) != 0:
        sys.exit(f"failed: {cmd[0]}")


def main():
    skip_app = "--skip-app" in sys.argv

    run([sys.executable, "tools/make_icon.py"])
    run([sys.executable, "tools/version_info.py"])

    if not skip_app:
        run([sys.executable, "-m", "PyInstaller", "docpilot.spec", "--noconfirm",
             "--workpath", "build_tmp", "--distpath", "dist"])

    if not os.path.isdir(BUNDLE):
        sys.exit(f"app bundle missing: {BUNDLE}\nrun without --skip-app first")

    num = ".".join((VERSION.split(".") + ["0", "0", "0"])[:4])
    run([find_iscc(), f"/DAppVersion={LABEL}", f"/DAppVersionNum={num}",
         "installer/DocPilot.iss"])

    setup = os.path.join(ROOT, "dist", f"DocPilot-Setup-{LABEL}.exe")
    if not os.path.isfile(setup):
        sys.exit("compiler reported success but the setup exe is missing")
    sha = hashlib.sha256()
    with open(setup, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            sha.update(chunk)
    print(f"\nOK  {setup}  ({os.path.getsize(setup) / 1e6:.0f} MB)")
    print("\nTo publish: upload that file to Drive, share it as 'Anyone with the "
          "link can view',\nthen set the update notice (Google Doc) to:\n")
    print(f"version: {VERSION}")
    print("installer: <the uploaded file's Drive link>")
    print(f"sha256: {sha.hexdigest()}")
    print("notes: <what is new, one line each>")


if __name__ == "__main__":
    main()
