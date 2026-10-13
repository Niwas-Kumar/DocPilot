# -*- coding: utf-8 -*-
"""Install the free fonts DocPilot needs (per-user, no admin required).

Only fonts with an open licence are fetched. Adobe Devanagari and Swiss721 BT
are commercial and are NOT downloaded - see the notes printed at the end.

    python tools\\install_fonts.py
"""
import os
import sys
import winreg

sys.stdout.reconfigure(encoding="utf-8")
import requests

GH_API = "https://api.github.com/repos/google/fonts/contents/{path}"
FONT_DIR = os.path.join(os.environ["LOCALAPPDATA"], "Microsoft", "Windows", "Fonts")
REG_KEY = r"Software\Microsoft\Windows NT\CurrentVersion\Fonts"

# family -> path inside the google/fonts repo
WANTED = {
    "Baloo 2": "ofl/baloo2",        # Hindi CA topic banners + cover index heading
    "Rozha One": "ofl/rozhaone",    # Hindi CA cover title (2026-08-31)
}


def list_repo_files(path):
    r = requests.get(GH_API.format(path=path), timeout=45,
                     headers={"Accept": "application/vnd.github+json"})
    r.raise_for_status()
    return [f for f in r.json() if f["name"].lower().endswith((".ttf", ".otf"))]


def already_installed(family):
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        w = win32com.client.DispatchEx("Word.Application")
        w.Visible = False
        names = {w.FontNames.Item(i) for i in range(1, w.FontNames.Count + 1)}
        w.Quit()
        pythoncom.CoUninitialize()
        return family in names
    except Exception:
        return False


def install_font_file(path, title=None):
    """Register a font for the current user."""
    name = os.path.basename(path)
    if title is None:
        try:
            from fontTools.ttLib import TTFont
            f = TTFont(path, fontNumber=0, lazy=True)
            title = f["name"].getDebugName(4) or f["name"].getDebugName(1)
            f.close()
        except Exception:
            title = os.path.splitext(name)[0]

    suffix = " (TrueType)" if name.lower().endswith(".ttf") else " (OpenType)"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_KEY) as key:
        winreg.SetValueEx(key, title + suffix, 0, winreg.REG_SZ, path)
    return title


def _set_names(font, family, subfamily):
    """Rewrite the name table so Word sees a normal Regular/Bold pair."""
    full = f"{family} {subfamily}" if subfamily != "Regular" else family
    ps = full.replace(" ", "")
    values = {1: family, 2: subfamily, 3: f"{ps};DocPilot", 4: full, 6: ps,
              16: family, 17: subfamily}
    name = font["name"]
    for rec in list(name.names):
        if rec.nameID in values:
            name.setName(values[rec.nameID], rec.nameID, rec.platformID,
                         rec.platEncID, rec.langID)
    for nid, val in values.items():
        if not name.getDebugName(nid):
            name.setName(val, nid, 3, 1, 0x409)

    bold = subfamily == "Bold"
    os2 = font["OS/2"]
    os2.fsSelection = (os2.fsSelection & ~0b1100001) | (0b100000 if bold else 0b1000000)
    os2.usWeightClass = 700 if bold else 400
    font["head"].macStyle = 1 if bold else 0


def make_static_instances(var_path, family, weights):
    """Variable fonts render at their default instance in Word no matter which
    named family is asked for, so bold never appears. Cutting real static
    instances out of the variable font fixes that.
    """
    from fontTools.ttLib import TTFont
    from fontTools.varLib import instancer

    made = []
    for subfamily, wght in weights.items():
        out = os.path.join(FONT_DIR, f"{family.replace(' ', '')}-{subfamily}.ttf")
        f = TTFont(var_path)
        inst = instancer.instantiateVariableFont(f, {"wght": wght}, updateFontNames=False)
        _set_names(inst, family, subfamily)
        inst.save(out)
        inst.close()
        f.close()
        title = f"{family} {subfamily}" if subfamily != "Regular" else family
        install_font_file(out, title)
        made.append((title, out))
    return made


def main():
    os.makedirs(FONT_DIR, exist_ok=True)
    installed, skipped = [], []

    for family, repo_path in WANTED.items():
        if already_installed(family):
            print(f"  {family}: already installed")
            skipped.append(family)
            continue
        print(f"  {family}: fetching {repo_path} ...")
        try:
            files = list_repo_files(repo_path)
        except Exception as e:
            print(f"     could not reach GitHub: {e}")
            continue
        if not files:
            print("     no font files found in that folder")
            continue
        for f in files:
            dest = os.path.join(FONT_DIR, f["name"])
            if not os.path.exists(dest):
                data = requests.get(f["download_url"], timeout=120).content
                with open(dest, "wb") as fh:
                    fh.write(data)
            is_variable = "[" in f["name"]
            if is_variable:
                try:
                    for title, path in make_static_instances(
                            dest, family, {"Regular": 400, "Bold": 700}):
                        print(f"     cut static  \"{title}\"  ({os.path.getsize(path)//1024} KB)")
                        installed.append(title)
                    continue
                except Exception as e:
                    print(f"     could not cut static instances ({e}); "
                          f"registering the variable font as-is")
            title = install_font_file(dest)
            print(f"     installed {f['name']}  ->  \"{title}\"  ({os.path.getsize(dest)//1024} KB)")
            installed.append(title)

    print()
    print("Installed:", installed or "nothing new")
    print()
    print("NOT installed (commercial licences - obtain them through your own"
          " Adobe / Bitstream licence if you want an exact typographic match):")
    print("  - Adobe Devanagari : official body face of the Hindi Daily CA."
          " Ships with Adobe Creative Cloud.")
    print("  - Swiss721 BT      : lettering inside the Hindi footer buttons."
          " Not needed - those buttons are embedded as images.")
    print("  - Utsaah           : index and date on the Hindi CA cover. A"
          " Microsoft font (Windows 'Hindi Supplemental Fonts' feature);"
          " Nirmala UI is used when it is absent.")
    print()
    print("Word picks up new fonts on next launch; DocPilot starts a fresh Word"
          " process per job, so no restart is needed.")


if __name__ == "__main__":
    main()
