# -*- coding: utf-8 -*-
"""Build DocPilot-Setup-<version>.exe end to end.

    python tools/build_installer.py            # icon + app bundle + installer
    python tools/build_installer.py --skip-app # installer only, reuse dist/

Leaves the finished setup at dist/DocPilot-Setup-2.0-beta.exe.
"""
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BUNDLE = os.path.join(ROOT, "dist", "DocPilot-2.0-beta")

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

    run([find_iscc(), "installer/DocPilot.iss"])

    setup = os.path.join(ROOT, "dist", "DocPilot-Setup-2.0-beta.exe")
    if os.path.isfile(setup):
        print(f"\nOK  {setup}  ({os.path.getsize(setup) / 1e6:.0f} MB)")
    else:
        sys.exit("compiler reported success but the setup exe is missing")


if __name__ == "__main__":
    main()
