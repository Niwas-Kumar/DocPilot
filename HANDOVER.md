# DocPilot — handover

DocPilot turns a raw typed UPSC document (.docx) into the finished, branded PW
OnlyIAS deliverable. It runs entirely on the user's own Windows PC: a small web
UI in the browser, a job queue, and **Microsoft Word** doing the page layout
through COM automation.

Start with **README.md** (what each category does, the design rules, the
measurements) and **AGENTS.md** (the hand-off notes: where the project stands,
what has already bitten, and the rules to keep). This file is only the
five-minute version.

---

## 1. What you need

| | |
|---|---|
| Windows | 10 or 11 (Word COM is Windows-only) |
| Microsoft Word | installed and activated — it renders every PDF |
| Python | 3.13 (3.11+ works) |
| Node.js | 18+ — **only** if you want to change the UI |

```bat
pip install -r backend\requirements.txt
```

Fonts need no action: `assets/fonts/` ships Baloo 2, Rozha One, Karma and
Bakbak One (all SIL Open Font License, licences included) and
`backend/fonts.py` installs them per user at startup, skipping any already on
the machine. The one face that cannot ship is **Adobe Devanagari** (commercial,
part of Adobe Creative Cloud): without it the Hindi booklets use Baloo 2 and
run ~18% longer. Install it if you want to match the printed copy line for line.

## 2. Run it

```bat
run.bat                 ::  http://localhost:8930  (opens the browser)
```
or `python -m uvicorn backend.app:app --port 8930`.

**Restart the server after every code change** — the pipelines are imported
once per process.

## 3. Build the desktop app

```bat
python tools\build_installer.py          :: icon -> PyInstaller -> Inno Setup, ~10 min
python tools\build_installer.py --skip-app   :: installer only, reuses dist\
```
Output: `dist\DocPilot-Setup-1.0-beta.exe`. Needs Inno Setup 6 (free). The
installer is **per-user on purpose** — the app writes its job scratch inside
its own folder, so Program Files would break it. See README → Installer.

## 4. Tests

```bat
python backend\test_suite.py     :: 33 checks, no server needed
python backend\test_series.py    :: 17 checks, Test Series pipelines
python backend\test_dea.py       :: 17 variant-input cases, Daily Editorial
python backend\test_api.py       :: needs the server running
python backend\test_api_series.py
```
`test_dea.py` and `test_series.py` read the sample files (see below). Everything
that renders goes through Word, so a full run takes a while — `test_dea.py`
alone is ~25 minutes because each case is a real booklet.

## 5. The sample files (ground truth) — SHIPPED SEPARATELY

`DocPilot-samples-<date>.zip` holds the folder `UPSC sample`, with, for every
category, the **raw input** the typists produce and the **published output** PW
actually releases. Every measurement in the code was taken from those PDFs, and
the tests read them.

The samples are **not in the repo** (they are client material). Put them at:

```
C:\Users\<you>\OneDrive\Desktop\UPSC sample\
```

or anywhere else and point the tests and tools at it:

```bat
set DOCPILOT_SAMPLES=D:\path\to\UPSC sample
```

## 6. Where things live

```
backend/app.py            FastAPI: /api/categories, /api/jobs, downloads
backend/categories.py     THE REGISTRY - add a category here first
backend/jobs.py           one-at-a-time queue (Word cannot be shared)
backend/pipeline/         one module per category + shared helpers
  styler/template/postprocess ....... Daily CA English (+ Xylem reuses them)
  daily_ca_hindi .................... Daily CA Hindi
  fitter ............................ measured pagination for Daily CA
  dea_* / combo_* / cover_kit ....... the four Daily Editorial editions
  cover*, convert, ooxml, docx_tools  covers, Word COM, OOXML helpers
assets/                   the client's artwork, templates and fonts
tools/                    asset/icon/font/installer builders (dev only)
frontend/src              React UI (frontend/dist is the built copy that ships)
```

## 7. House rules (please keep)

1. **Format, never invent.** If the published sample shows something the raw
   file does not contain, leave the gap and report it in the job notes. The
   user fills it in.
2. **Do not disturb a working category.** Each one was signed off against a
   published sample. New work goes in new modules that import the old helpers
   read-only (that is how the Daily Editorial editions were added).
3. **Measure, don't guess.** Every constant in the code came from measuring the
   published PDF. When something looks wrong, measure both and compare.
4. Keep README.md and AGENTS.md up to date — they are the project's memory.

## 8. Not included in this zip

Rebuildable or machine-local only: `work/` (job scratch), `dist/`,
`build_tmp/`, `frontend/node_modules/`, `__pycache__/`, server logs, and the
`.git` history. The code lives in the private GitHub repo
`Sahild191/DocPilot` — ask the owner for access if you want the history.
