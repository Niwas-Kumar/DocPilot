# -*- coding: utf-8 -*-
"""CLI smoke test: run the Daily CA pipeline on the official sample input."""
import sys, os, time
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from backend.pipeline.daily_ca import process

SAMPLES = os.environ.get("DOCPILOT_SAMPLES") or os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop", "UPSC sample")
SAMPLE = os.path.join(SAMPLES, "Daily Current Affairs Summary (English)", "17th August 2026.docx")
WORK = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "work", "clitest")

t0 = time.time()
out = process(SAMPLE, WORK, variant="English",
              progress=lambda s, p: print(f"[{p:3d}%] {s}", flush=True))
print(f"OK in {time.time()-t0:.1f}s -> {out}")
