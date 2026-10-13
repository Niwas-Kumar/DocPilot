# -*- coding: utf-8 -*-
"""API / queue / stop behavior tests against the live server."""
import os
import sys
import time
import traceback

sys.stdout.reconfigure(encoding="utf-8")
import requests
import fitz

BASE = "http://127.0.0.1:8930"
SAMPLES = os.environ.get("DOCPILOT_SAMPLES") or os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop", "UPSC sample")
EN_SAMPLE = os.path.join(SAMPLES, "Daily Current Affairs Summary (English)", "17th August 2026.docx")
EDGE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "work", "tests", "edge_inputs")

PASS, FAIL = [], []


def check(name, fn):
    try:
        fn()
        PASS.append(name)
        print(f"  PASS  {name}")
    except Exception as e:
        FAIL.append((name, e))
        print(f"  FAIL  {name}: {type(e).__name__}: {e}")
        traceback.print_exc(limit=3)


def submit_file(path, cat="daily_ca_english", rename=None):
    with open(path, "rb") as f:
        r = requests.post(f"{BASE}/api/jobs", data={"category_id": cat},
                          files={"file": (rename or os.path.basename(path), f)})
    return r


def wait(jid, timeout=120, until=("done", "error", "stopped")):
    t0 = time.time()
    while time.time() - t0 < timeout:
        j = next((x for x in requests.get(f"{BASE}/api/jobs").json() if x["id"] == jid), None)
        if j and j["status"] in until:
            return j
        time.sleep(1.5)
    raise TimeoutError(f"job {jid} did not settle")


print("== 3. API & queue ==")

check("server is up", lambda: requests.get(f"{BASE}/api/categories", timeout=5).raise_for_status())


def _bad_inputs():
    r = requests.post(f"{BASE}/api/jobs", data={"category_id": "nope"}, files={"file": ("x.docx", b"1")})
    assert r.status_code == 400, r.status_code
    r = requests.post(f"{BASE}/api/jobs", data={"category_id": "daily_ca_english"})
    assert r.status_code == 400
    r = requests.post(f"{BASE}/api/jobs", data={"category_id": "daily_ca_english"},
                      files={"file": ("notes.txt", b"hello")})
    assert r.status_code == 400 and "docx" in r.json()["detail"]
    # a category with no sample pair yet must stay disabled
    r = requests.post(f"{BASE}/api/jobs", data={"category_id": "mains_ab_english"},
                      files={"file": ("x.docx", b"1")})
    assert r.status_code == 400 and "not configured" in r.json()["detail"], r.text[:120]
check("invalid submissions rejected", _bad_inputs)


def _corrupt_docx():
    """A .docx that is not really a docx must fail cleanly, not hang."""
    r = requests.post(f"{BASE}/api/jobs", data={"category_id": "prelims_ab_english",
                                                "start_page": "23"},
                      files={"file": ("broken.docx", b"this is not a zip archive")})
    assert r.status_code == 200, r.text[:160]
    jid = r.json()["id"]
    j = wait(jid, timeout=180)
    assert j["status"] == "error", j["status"]
    assert j["error"], "an error message should be reported"
    requests.delete(f"{BASE}/api/jobs/{jid}")
check("corrupt .docx fails cleanly with a message", _corrupt_docx)


def _queue_of_three():
    ids = []
    for i, name in enumerate(("queue test A.docx", "queue test B.docx", "queue test C.docx")):
        r = submit_file(os.path.join(EDGE, "minimal.docx"), rename=name)
        r.raise_for_status()
        ids.append(r.json()["id"])
    jobs = [wait(j, timeout=180) for j in ids]
    assert all(j["status"] == "done" for j in jobs), [j["status"] for j in jobs]
    # serialized: no two jobs overlap in run window
    spans = sorted((j["started"], j["finished"]) for j in jobs)
    for (s1, f1), (s2, f2) in zip(spans, spans[1:]):
        assert s2 >= f1 - 0.5, "jobs overlapped - Word COM not serialized!"
    for j in ids:
        requests.delete(f"{BASE}/api/jobs/{j}")
check("3-job queue runs serialized, all complete", _queue_of_three)


def _stop_mid_run():
    r = submit_file(EN_SAMPLE)  # ~10s job
    jid = r.json()["id"]
    # wait until running, then stop
    t0 = time.time()
    while time.time() - t0 < 60:
        j = next(x for x in requests.get(f"{BASE}/api/jobs").json() if x["id"] == jid)
        if j["status"] == "running":
            break
        time.sleep(0.4)
    assert j["status"] == "running", "never started"
    # removing a RUNNING job must be blocked
    assert requests.delete(f"{BASE}/api/jobs/{jid}").status_code == 400
    requests.post(f"{BASE}/api/jobs/{jid}/stop").raise_for_status()
    j = wait(jid, timeout=60)
    assert j["status"] == "stopped", j["status"]
    # now removable
    assert requests.delete(f"{BASE}/api/jobs/{jid}").json()["ok"] is True
check("stop mid-run + remove-running blocked", _stop_mid_run)


def _download_gates():
    r = submit_file(os.path.join(EDGE, "minimal.docx"), rename="gate test.docx")
    jid = r.json()["id"]
    dr = requests.get(f"{BASE}/api/jobs/{jid}/download")
    assert dr.status_code == 404, "download before done should 404"
    j = wait(jid)
    assert j["status"] == "done"
    dr = requests.get(f"{BASE}/api/jobs/{jid}/download")
    assert dr.status_code == 200 and dr.content[:4] == b"%PDF"
    assert "gate%20test_CA%20Analysis_UPSC_English.pdf" in dr.headers.get("content-disposition", "")
    requests.delete(f"{BASE}/api/jobs/{jid}")
check("download gated until done, correct name", _download_gates)


def _special_filename():
    r = submit_file(os.path.join(EDGE, "minimal.docx"), rename="Test & File's (v2) हिंदी.docx")
    assert r.status_code == 200, r.text
    j = wait(r.json()["id"])
    assert j["status"] == "done", j.get("error")
    assert "Test & File's (v2) हिंदी" in j["output_name"]
    requests.delete(f"{BASE}/api/jobs/{j['id']}")
check("special characters + Hindi in filename", _special_filename)


def _real_restricted_link():
    # a REAL private Google Doc (not link-shared), e.g.
    # DOCPILOT_RESTRICTED_DOC=https://docs.google.com/document/d/<id>/edit
    r = requests.post(f"{BASE}/api/jobs", data={
        "category_id": "daily_ca_english",
        "link": os.environ["DOCPILOT_RESTRICTED_DOC"]})
    r.raise_for_status()
    j = wait(r.json()["id"], timeout=60)
    assert j["status"] == "error", j["status"]
    assert "restricted" in j["error"] or "sign-in" in j["error"], j["error"]
    requests.delete(f"{BASE}/api/jobs/{j['id']}")
if os.environ.get("DOCPILOT_RESTRICTED_DOC"):
    check("real restricted Drive link -> clean actionable error", _real_restricted_link)
else:
    print("  SKIP  real restricted Drive link (set DOCPILOT_RESTRICTED_DOC)")


def _garbage_link():
    r = requests.post(f"{BASE}/api/jobs", data={
        "category_id": "daily_ca_english", "link": "https://example.com/file.docx"})
    j = wait(r.json()["id"], timeout=30)
    assert j["status"] == "error" and "recognizable" in j["error"]
    requests.delete(f"{BASE}/api/jobs/{j['id']}")
check("non-Drive link -> clean error", _garbage_link)

print()
print(f"API section: {len(PASS)} passed, {len(FAIL)} failed")
if FAIL:
    print("FAILURES:")
    for n, e in FAIL:
        print(" -", n, "->", e)
    sys.exit(1)
