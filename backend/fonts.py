# -*- coding: utf-8 -*-
"""Install the open-licence fonts DocPilot ships with, per user, at startup.

Why this exists: the Hindi Daily CA design needs **Baloo 2** (topic banners,
index heading) and **Rozha One** (cover title); the Hindi Daily Editorial
covers add **Karma** (date) and **Bakbak One** (index). Those were installed by hand
on the development machine, so the layout was correct there and wrong on the
client's PC - the desktop build carried neither the fonts nor the installer
script, Word substituted Nirmala UI, and the cover title grew 20pt taller and
printed straight through the PW logo.

Rules this module follows:

* **Never fight an existing installation.** If the family is already known to
  Windows - system-wide (HKLM, e.g. an IT-managed rollout) or per user (HKCU,
  e.g. someone ran `tools\\install_fonts.py`) - it is left completely alone.
  A newer or licensed copy on the user's machine always wins.
* **No admin rights.** Files go to the per-user font directory and are
  registered under HKCU, which is the supported per-user font install on
  Windows 10+.
* **Usable immediately.** The font is also added to the running session and a
  WM_FONTCHANGE broadcast tells other apps; DocPilot starts a fresh Word
  process per job, so the very first job after startup already sees it.
* **Never fatal.** Any failure downgrades to "not installed" - the pipelines
  already fall back to Nirmala UI and say so in the job notes. Formatting a
  booklet must never fail because a font could not be registered.

Commercial faces are deliberately NOT bundled: **Adobe Devanagari** (the
official Hindi body face, ships with Adobe CC) and **Utsaah** (a Microsoft
font, part of the Windows "Hindi Supplemental Fonts" feature).
"""
import os
import shutil
import sys
import threading

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FONT_DIR = os.path.join(ROOT, "assets", "fonts")
REG_KEY = r"Software\Microsoft\Windows NT\CurrentVersion\Fonts"

# file -> (family Word matches on, value name to register under)
# The registry name follows the Windows convention: the family alone for a
# Regular face, "Family Style" for the others.
BUNDLED = [
    ("Baloo2-Regular.ttf", "Baloo 2", "Baloo 2 (TrueType)"),
    ("Baloo2-Bold.ttf", "Baloo 2", "Baloo 2 Bold (TrueType)"),
    ("RozhaOne-Regular.ttf", "Rozha One", "Rozha One (TrueType)"),
    # Daily Editorial Analysis (Hindi) covers: date + index faces (2026-09-21)
    ("Karma-Regular.ttf", "Karma", "Karma (TrueType)"),
    ("Karma-Bold.ttf", "Karma", "Karma Bold (TrueType)"),
    ("BakbakOne-Regular.ttf", "Bakbak One", "Bakbak One (TrueType)"),
]

_lock = threading.Lock()
_done = None            # cached report, so startup only ever does this once


def installed_families():
    """Lower-cased font titles Windows knows about, machine-wide and per user."""
    import re
    import winreg
    names = set()
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            key = winreg.OpenKey(root, REG_KEY)
        except OSError:
            continue
        try:
            i = 0
            while True:
                try:
                    value_name, _, _ = winreg.EnumValue(key, i)
                except OSError:
                    break
                i += 1
                label = re.sub(r"\s*\((TrueType|OpenType|VGA res)\)\s*$", "",
                               value_name, flags=re.I)
                for part in label.split("&"):
                    names.add(part.strip().lower())
        finally:
            key.Close()
    return names


def _available(family, names):
    low = family.lower()
    return any(n == low or n.startswith(low + " ") for n in names)


def _register_session(path):
    """Make the font usable by processes started from now on, without a logout.

    `AddFontResourceW` loads it for this process; the WM_FONTCHANGE broadcast
    asks everyone else to refresh. Sent with a timeout because a broadcast to
    a hung window would otherwise block startup.
    """
    try:
        import ctypes
        from ctypes import wintypes
        ctypes.windll.gdi32.AddFontResourceW(ctypes.c_wchar_p(path))
        ctypes.windll.user32.SendMessageTimeoutW(
            wintypes.HWND(0xFFFF),      # HWND_BROADCAST
            wintypes.UINT(0x001D),      # WM_FONTCHANGE
            wintypes.WPARAM(0), wintypes.LPARAM(0),
            wintypes.UINT(0x0002),      # SMTO_ABORTIFHUNG
            wintypes.UINT(1000), None)
    except Exception:
        pass                            # cosmetic: the registry entry still stands


def _install_one(filename, value_name, dest_dir):
    """Copy + register one face. Returns 'installed' or raises."""
    import winreg
    src = os.path.join(FONT_DIR, filename)
    dest = os.path.join(dest_dir, filename)
    if not os.path.exists(src):
        raise FileNotFoundError(src)
    os.makedirs(dest_dir, exist_ok=True)
    if not os.path.exists(dest) or os.path.getsize(dest) != os.path.getsize(src):
        try:
            shutil.copyfile(src, dest)
        except OSError:
            # already there and loaded by another process - the existing file
            # is the same font, so registering it is still the right move
            if not os.path.exists(dest):
                raise
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_KEY) as key:
        winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, dest)
    _register_session(dest)
    return dest


def ensure_fonts(force=False):
    """Install the bundled fonts that are missing. Idempotent and safe to call
    on every start. Returns [(family, status, detail)] where status is one of
    'present' (already on the machine - untouched), 'installed', 'failed'."""
    global _done
    with _lock:
        if _done is not None and not force:
            return _done
        report = []
        if not sys.platform.startswith("win"):
            _done = [("(all)", "skipped", "not Windows")]
            return _done
        try:
            names = installed_families()
        except Exception as exc:
            names = set()
            report.append(("(registry)", "failed", f"{type(exc).__name__}: {exc}"))
        dest_dir = os.path.join(os.environ.get("LOCALAPPDATA", ""),
                                "Microsoft", "Windows", "Fonts")
        seen = set()
        for filename, family, value_name in BUNDLED:
            try:
                if _available(family, names):
                    if family not in seen:
                        report.append((family, "present", "already installed - left alone"))
                        seen.add(family)
                    continue
                dest = _install_one(filename, value_name, dest_dir)
                report.append((family, "installed", dest))
            except Exception as exc:
                report.append((family, "failed", f"{type(exc).__name__}: {exc}"))
        _done = report
        return report


def summary(report=None):
    """One line for the startup log / job notes."""
    report = report if report is not None else ensure_fonts()
    installed = sorted({f for f, s, _ in report if s == "installed"})
    present = sorted({f for f, s, _ in report if s == "present"})
    failed = sorted({f for f, s, _ in report if s == "failed"})
    bits = []
    if installed:
        bits.append("installed " + ", ".join(installed))
    if present:
        bits.append("already present " + ", ".join(present))
    if failed:
        bits.append("could not install " + ", ".join(failed))
    return "Bundled fonts: " + ("; ".join(bits) if bits else "nothing to do") + "."


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    for row in ensure_fonts(force=True):
        print(f"  {row[0]:12} {row[1]:10} {row[2]}")
    print(summary())
