# -*- coding: utf-8 -*-
"""In-app updates from Google Drive.

A release is a new DocPilot-Setup-<version>.exe uploaded to Drive, plus a
small manifest whose link is baked into the app (version.UPDATE_MANIFEST).
The manifest is a Google Doc - easy to edit in the browser - holding lines
like:

    version: 2.1.1
    installer: https://drive.google.com/file/d/<id>/view
    sha256: 3f7a...            (optional, printed by tools/build_installer.py)
    notes: Prelims Question Booklet in English and Hindi
    Any other line is added to the notes.

(A JSON object with the same keys works too.)

Updating = download that setup, then hand over to it: DocPilot exits, a
hidden PowerShell waits for the process to end - the installer refuses to
run while DocPilot holds its AppMutex - and starts the setup silently with
/RELAUNCH=1, which reopens DocPilot when it is done. The installer is
per-user, so none of this needs admin rights.
"""
import hashlib
import html
import json
import os
import re
import subprocess
import sys
import tempfile
import threading

import requests

from . import version
from .drive import UA, parse_drive_link, DriveError

KEYS = ("version", "installer", "sha256", "notes")
DOWNLOAD_DIR = os.path.join(tempfile.gettempdir(), "DocPilot-update")


class UpdateError(Exception):
    pass


# --------------------------------------------------------------- manifest
def _manifest_url(link):
    """The plain-text form of the manifest link."""
    try:
        kind, fid = parse_drive_link(link)
    except DriveError:
        return link                       # any other URL is fetched as it is
    if kind == "gdoc":
        return f"https://docs.google.com/document/d/{fid}/export?format=txt"
    return f"https://drive.usercontent.google.com/download?id={fid}&export=download&confirm=t"


def parse_manifest(text):
    text = text.lstrip("﻿").strip()
    try:
        data = json.loads(text)
        if isinstance(data, dict):
            return {k: str(data.get(k, "")).strip() for k in KEYS}
    except ValueError:
        pass
    out = {k: "" for k in KEYS}
    notes = []
    for line in text.splitlines():
        m = re.match(r"^\s*([A-Za-z0-9_]+)\s*[:=]\s*(.*)$", line)
        key = m.group(1).lower() if m else None
        if key in KEYS and key != "notes":
            out[key] = m.group(2).strip()
        elif key == "notes":
            notes.append(m.group(2).strip())
        elif line.strip():
            notes.append(line.strip())
    out["notes"] = "\n".join(n for n in notes if n)
    return out


def check():
    """What the app knows about updates, for the UI."""
    info = {"current": version.LABEL, "configured": bool(version.UPDATE_MANIFEST),
            "frozen": bool(getattr(sys, "frozen", False)),
            "available": False, "latest": None, "notes": "", "error": None}
    if not version.UPDATE_MANIFEST:
        return info
    try:
        r = requests.get(_manifest_url(version.UPDATE_MANIFEST), headers=UA,
                         timeout=20)
        r.raise_for_status()
        if "text/html" in r.headers.get("Content-Type", ""):
            raise UpdateError("the update notice is not shared as 'Anyone with "
                              "the link can view'")
        m = parse_manifest(r.content.decode("utf-8", "replace"))
    except (requests.RequestException, UpdateError) as e:
        info["error"] = f"Could not check for updates: {e}"
        return info
    if not m["version"] or not m["installer"]:
        info["error"] = "The update notice has no version or installer link."
        return info
    info.update(latest=m["version"], notes=m["notes"],
                available=version.parse(m["version"]) > version.parse(version.VERSION))
    with _lock:
        _state["manifest"] = m
    return info


# --------------------------------------------------------------- download
_lock = threading.Lock()
_state = {"state": "idle",          # idle|downloading|ready|installing|error
          "progress": 0, "error": None, "path": None, "manifest": None}


def status():
    with _lock:
        return {k: _state[k] for k in ("state", "progress", "error")}


def _set(**kw):
    with _lock:
        _state.update(kw)


def _open_drive_file(link, sess):
    """A streaming response for a Drive file of any size.

    Files over ~100 MB answer with a "can't scan for viruses" page whose form
    carries the real download URL and its one-time fields; it is submitted
    as is. confirm=t usually skips that page in the first place."""
    try:
        kind, fid = parse_drive_link(link)
    except DriveError as e:
        raise UpdateError(str(e))
    if kind != "file":
        raise UpdateError("The installer link must point to the uploaded .exe "
                          "file, not to a Google Doc.")
    r = sess.get("https://drive.usercontent.google.com/download",
                 params={"id": fid, "export": "download", "confirm": "t"},
                 headers=UA, stream=True, timeout=60)
    not_shared = UpdateError("The installer on Drive is not shared as 'Anyone "
                             "with the link can view' - Drive asked for a "
                             "sign-in instead of handing it over.")
    if "accounts.google.com" in r.url or "/signin" in r.url:
        raise not_shared
    if "text/html" in r.headers.get("Content-Type", ""):
        page = r.text
        form = re.search(r'<form[^>]+action="([^"]+)"', page)
        fields = dict(re.findall(r'<input[^>]+name="([^"]+)"[^>]+value="([^"]*)"',
                                 page))
        # a restricted file answers with a sign-in form, not the virus-scan one
        if form and not html.unescape(form.group(1)).startswith("https://"):
            raise not_shared
        if not form or not fields:
            raise UpdateError("Drive did not hand over the installer - share it "
                              "as 'Anyone with the link can view'.")
        r = sess.get(html.unescape(form.group(1)), params=fields, headers=UA,
                     stream=True, timeout=60)
    if r.status_code != 200 or "text/html" in r.headers.get("Content-Type", ""):
        raise UpdateError(f"Drive refused the installer download (HTTP "
                          f"{r.status_code}). Share it as 'Anyone with the "
                          f"link can view'.")
    return r


def _download(m):
    try:
        os.makedirs(DOWNLOAD_DIR, exist_ok=True)
        path = os.path.join(DOWNLOAD_DIR, f"DocPilot-Setup-{m['version']}.exe")
        part = path + ".part"
        r = _open_drive_file(m["installer"], requests.Session())
        total = int(r.headers.get("Content-Length") or 0)
        got = 0
        sha = hashlib.sha256()
        with open(part, "wb") as fh:
            for chunk in r.iter_content(1024 * 512):
                fh.write(chunk)
                sha.update(chunk)
                got += len(chunk)
                if total:
                    _set(progress=min(99, int(got * 100 / total)))
        with open(part, "rb") as fh:
            if fh.read(2) != b"MZ":
                raise UpdateError("The downloaded installer is not a Windows "
                                  "program - check the installer link.")
        want = m.get("sha256", "").lower()
        if want and sha.hexdigest() != want:
            raise UpdateError("The downloaded installer does not match the "
                              "update notice's sha256 - download damaged, or "
                              "the notice points to the wrong file.")
        os.replace(part, path)
        _set(state="ready", progress=100, path=path)
    except Exception as e:                       # noqa: BLE001 - shown to user
        _set(state="error", error=str(e) or type(e).__name__)


def start_download():
    with _lock:
        if _state["state"] in ("downloading", "installing"):
            return
        m = _state["manifest"]
    if m is None:
        info = check()
        if info["error"]:
            raise UpdateError(info["error"])
        with _lock:
            m = _state["manifest"]
    if version.parse(m["version"]) <= version.parse(version.VERSION):
        raise UpdateError("DocPilot is already up to date.")
    _set(state="downloading", progress=0, error=None, path=None)
    threading.Thread(target=_download, args=(m,), daemon=True).start()


# ----------------------------------------------------------------- install
def install():
    """Hand over to the downloaded installer and quit DocPilot."""
    if not getattr(sys, "frozen", False):
        raise UpdateError("This copy runs from source code - updates install "
                          "only into the installed DocPilot app.")
    with _lock:
        path = _state["path"]
        if _state["state"] != "ready" or not path or not os.path.isfile(path):
            raise UpdateError("The update has not been downloaded yet.")
        _state["state"] = "installing"

    def ps_quote(s):
        return "'" + s.replace("'", "''") + "'"

    # Setup's own log beside the download: the only record of a silent
    # install that failed after DocPilot had already closed
    log_arg = ps_quote('/LOG="%s"' % os.path.join(DOWNLOAD_DIR, "setup.log"))
    script = (f"Wait-Process -Id {os.getpid()} -ErrorAction SilentlyContinue; "
              f"Start-Process -FilePath {ps_quote(path)} -ArgumentList "
              f"'/SILENT','/SUPPRESSMSGBOXES','/NORESTART','/RELAUNCH=1',"
              f"{log_arg}")
    # A new process group with no window - NOT DETACHED_PROCESS: PowerShell
    # started without a console does nothing at all, so 2.1.0 closed for the
    # update and never ran the installer. Its own group keeps it alive when
    # DocPilot exits a moment later.
    flags = 0x00000200 | 0x08000000                # new group, no window
    subprocess.Popen(["powershell", "-NoProfile", "-NonInteractive",
                      "-WindowStyle", "Hidden", "-Command", script],
                     creationflags=flags, close_fds=True)
    # answer the request first, then let go of the files the installer replaces
    threading.Timer(1.5, os._exit, args=(0,)).start()
