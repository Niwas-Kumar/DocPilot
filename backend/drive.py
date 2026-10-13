# -*- coding: utf-8 -*-
"""Google Drive / Google Docs shared-link downloader (no API keys).

Supported link shapes:
- https://docs.google.com/document/d/<ID>/...      (native Google Doc -> export .docx)
- https://drive.google.com/file/d/<ID>/...          (uploaded .docx file)
- https://drive.google.com/open?id=<ID>
- https://drive.google.com/uc?id=<ID>...
"""
import os
import re
import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) DocPilot/1.0"}


class DriveError(Exception):
    pass


def parse_drive_link(url: str):
    """Returns (kind, file_id) where kind is 'gdoc' or 'file'."""
    url = url.strip()
    m = re.search(r"docs\.google\.com/document/d/([\w-]{10,})", url)
    if m:
        return "gdoc", m.group(1)
    m = re.search(r"drive\.google\.com/file/d/([\w-]{10,})", url)
    if m:
        return "file", m.group(1)
    m = re.search(r"[?&]id=([\w-]{10,})", url)
    if m:
        return "file", m.group(1)
    raise DriveError("Not a recognizable Google Drive / Docs link")


def download_docx(url: str, dest_dir: str, cancelled=lambda: False) -> str:
    """Download the shared file as .docx into dest_dir; returns local path."""
    kind, fid = parse_drive_link(url)
    os.makedirs(dest_dir, exist_ok=True)

    if kind == "gdoc":
        dl = f"https://docs.google.com/document/d/{fid}/export?format=docx"
    else:
        dl = f"https://drive.google.com/uc?export=download&id={fid}"

    sess = requests.Session()
    r = sess.get(dl, headers=UA, stream=True, timeout=60)
    # big-file interstitial ("can't scan for viruses") -> confirm token
    if "text/html" in r.headers.get("Content-Type", "") and kind == "file":
        m = re.search(r'confirm=([\w-]+)', r.text)
        token = m.group(1) if m else "t"
        r = sess.get(f"https://drive.google.com/uc?export=download&id={fid}&confirm={token}",
                     headers=UA, stream=True, timeout=60)
    if r.status_code != 200:
        raise DriveError(
            f"Download failed (HTTP {r.status_code}). The link must be shared as "
            "'Anyone with the link can view' - links restricted to the physicswallah.org "
            "domain or to specific people cannot be fetched without login.")
    ctype = r.headers.get("Content-Type", "")
    if "text/html" in ctype:
        raise DriveError(
            "Drive returned a sign-in page instead of the file. The link is restricted "
            "(domain-only or specific people). Change sharing to 'Anyone with the link "
            "can view', or download the file and use Local Upload instead.")

    # filename from headers if present
    name = None
    cd = r.headers.get("Content-Disposition", "")
    m = re.search(r"filename\*=UTF-8''([^;]+)", cd) or re.search(r'filename="([^"]+)"', cd)
    if m:
        from urllib.parse import unquote
        name = os.path.basename(unquote(m.group(1)))
    if not name:
        name = fid + ".docx"
    elif not name.lower().endswith(".docx"):
        name += ".docx"

    path = os.path.join(dest_dir, _sanitize(name))
    with open(path, "wb") as f:
        for chunk in r.iter_content(1024 * 256):
            if cancelled():
                f.close()
                os.remove(path)
                raise InterruptedError("stopped")
            f.write(chunk)
    if os.path.getsize(path) < 2000:
        raise DriveError("Downloaded file is suspiciously small - check link permissions.")
    return path


def _sanitize(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", name).strip() or "input.docx"
