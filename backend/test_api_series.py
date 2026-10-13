# -*- coding: utf-8 -*-
"""API tests for the Test Series categories (server must be on :8930)."""
import os
import sys
import time
import traceback

sys.stdout.reconfigure(encoding="utf-8")
import requests

BASE = "http://127.0.0.1:8930"
S = os.environ.get("DOCPILOT_SAMPLES") or os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop", "UPSC sample")
PRE_EN = os.path.join(S, "Prelims Test Series (Answer Booklet English)", "Input English.docx")
QB_EN = os.path.join(S, "Prelims Test Series (Question Booklet Bilingual)", "Input English (1).docx")
QB_HI = os.path.join(S, "Prelims Test Series (Question Booklet Bilingual)", "Input Hindi (1).docx")
MAINS = os.path.join(S, "Mains Test Series (Answer Booklet Hindi)", "Input SDM 2719 Hind  (1).docx")

PASS, FAIL = [], []


def check(name, fn):
    try:
        fn()
        PASS.append(name)
        print(f"  PASS  {name}")
    except Exception as e:
        FAIL.append((name, e))
        print(f"  FAIL  {name}: {type(e).__name__}: {e}")
        traceback.print_exc(limit=2)


def post(cat, path=None, path2=None, **data):
    files = {}
    handles = []
    if path:
        h = open(path, "rb")
        handles.append(h)
        files["file"] = (os.path.basename(path), h)
    if path2:
        h = open(path2, "rb")
        handles.append(h)
        files["file2"] = (os.path.basename(path2), h)
    try:
        return requests.post(BASE + "/api/jobs",
                             data={"category_id": cat, **{k: str(v) for k, v in data.items()}},
                             files=files or None)
    finally:
        for h in handles:
            h.close()


def wait(jid, timeout=900):
    t0 = time.time()
    while time.time() - t0 < timeout:
        j = next((x for x in requests.get(BASE + "/api/jobs").json() if x["id"] == jid), None)
        if j and j["status"] in ("done", "error", "stopped"):
            return j
        time.sleep(2)
    raise TimeoutError("job did not settle")


print("== categories metadata ==")


def _meta():
    cats = {c["id"]: c for c in requests.get(BASE + "/api/categories").json()}
    assert cats["prelims_ab_english"]["pages"] == "start"
    assert cats["prelims_ab_english"]["files"] == 1
    assert cats["prelims_ab_english"]["output"] == "docx"
    assert cats["prelims_qb_bilingual"]["pages"] == "start_end"
    assert cats["prelims_qb_bilingual"]["files"] == 2
    assert cats["daily_ca_english"]["output"] == "pdf"
    assert cats["daily_ca_english"]["pages"] == "none"
    ready = [c for c in cats.values() if c["ready"]]
    assert len(ready) == 6, [c["id"] for c in ready]
check("category metadata drives the UI correctly", _meta)


print("== page validation ==")


def _bad_pages():
    r = post("prelims_ab_english", PRE_EN, start_page=0)
    assert r.status_code == 400 and "1 or greater" in r.text, r.text[:120]
    r = post("prelims_ab_english", PRE_EN, start_page="abc")
    assert r.status_code == 400 and "whole number" in r.text, r.text[:120]
    r = post("prelims_qb_bilingual", QB_EN, QB_HI, start_page=20, end_page=5)
    assert r.status_code == 400 and "before the start" in r.text, r.text[:120]
check("invalid page numbers rejected with a clear reason", _bad_pages)


def _missing_second_file():
    r = post("prelims_qb_bilingual", QB_EN, start_page=1)
    assert r.status_code == 400 and "two files" in r.text, r.text[:160]
check("bilingual without the Hindi file -> 400", _missing_second_file)


def _pending_category():
    r = post("mains_ab_english", PRE_EN, start_page=1)
    assert r.status_code == 400 and "not configured" in r.text, r.text[:120]
check("category without a sample stays disabled", _pending_category)


def _pages_ignored_when_not_applicable():
    """Daily CA takes no pages; stray values must not reach the frozen pipeline."""
    r = post("daily_ca_english", PRE_EN, start_page=99)
    assert r.status_code == 200, r.text[:160]
    j = r.json()
    assert j["start_page"] is None, j
    requests.post(BASE + f"/api/jobs/{j['id']}/stop")
    requests.delete(BASE + f"/api/jobs/{j['id']}")
check("page fields ignored for categories that do not use them", _pages_ignored_when_not_applicable)


print("== end to end ==")


def _prelims_download():
    r = post("prelims_ab_english", PRE_EN, start_page=23)
    r.raise_for_status()
    j = wait(r.json()["id"])
    assert j["status"] == "done", j.get("error")
    assert j["start_page"] == 23, j
    blob = requests.get(BASE + f"/api/jobs/{j['id']}/download")
    assert blob.status_code == 200
    assert blob.content[:2] == b"PK", "should be a .docx (zip) not a PDF"
    ct = blob.headers.get("content-type", "")
    assert "wordprocessingml" in ct, ct
    requests.delete(BASE + f"/api/jobs/{j['id']}")
check("Prelims booklet downloads as a real .docx", _prelims_download)


def _mains_notes_surface():
    r = post("mains_ab_hindi", MAINS)
    r.raise_for_status()
    j = wait(r.json()["id"])
    assert j["status"] == "done", j.get("error")
    notes = " ".join(j.get("notes") or [])
    assert "not in the source file" in notes, notes
    requests.delete(BASE + f"/api/jobs/{j['id']}")
check("Mains job reports what it deliberately left out", _mains_notes_surface)


def _stop_long_job():
    """The bilingual job takes ~2 min - a good target for Stop."""
    r = post("prelims_qb_bilingual", QB_EN, QB_HI, start_page=1, end_page=22)
    r.raise_for_status()
    jid = r.json()["id"]
    t0 = time.time()
    while time.time() - t0 < 120:
        j = next(x for x in requests.get(BASE + "/api/jobs").json() if x["id"] == jid)
        if j["status"] == "running" and j["progress"] >= 28:
            break
        time.sleep(1)
    assert j["status"] == "running", j
    requests.post(BASE + f"/api/jobs/{jid}/stop").raise_for_status()
    j = wait(jid, timeout=300)
    assert j["status"] == "stopped", j["status"]
    requests.delete(BASE + f"/api/jobs/{jid}")
check("Stop interrupts a long bilingual job", _stop_long_job)


def _queue_is_serial():
    ids = []
    for _ in range(2):
        r = post("prelims_ab_english", PRE_EN, start_page=23)
        r.raise_for_status()
        ids.append(r.json()["id"])
    jobs = [wait(i, timeout=600) for i in ids]
    assert all(j["status"] == "done" for j in jobs), [j["status"] for j in jobs]
    spans = sorted((j["started"], j["finished"]) for j in jobs)
    for (s1, f1), (s2, f2) in zip(spans, spans[1:]):
        assert s2 >= f1 - 0.5, "jobs overlapped - Word COM not serialized"
    for i in ids:
        requests.delete(BASE + f"/api/jobs/{i}")
check("two booklet jobs run one after the other", _queue_is_serial)


print()
print(f"{len(PASS)} passed, {len(FAIL)} failed")
if FAIL:
    print("FAILURES:")
    for n, e in FAIL:
        print(" -", n, "->", e)
    sys.exit(1)
