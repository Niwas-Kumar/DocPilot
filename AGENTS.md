# AGENTS.md — DocPilot

Working notes for AI agents and developers on this project.

## What this is

Local web tool converting raw typed UPSC docx (Word/Google Docs) into branded
OnlyIAS PDFs per category. React frontend + FastAPI backend + Word COM engine.
Fully offline, zero API keys, deterministic (no LLM in the pipeline).

## HANDOFF — read this first (state as of 2026-09-21)

Written so another agent can continue cold. Nothing here is committed: the last
commit is `898f4b5`; everything after it lives in the working tree only. The
client's standing rules are local-first (never push or deploy without being
asked, each time) and **format, never invent**.

### Where the project stands

- **Ten live categories**: Daily CA English, Daily CA Hindi, Daily CA English
  Xylem, the four **Daily Editorial Analysis** editions (added 2026-09-21: DEA
  English, DEA Hindi, Daily CA + DEA English, Daily CA + DEA Hindi - see the
  section below), and the three Prelims booklets. All three Mains entries are gated off
  (`ready: False`) and show "template pending"; `mains_ab.py` still works and
  is one flag away from returning.
- **All three Daily CA editions have a front page** (added 2026-08-31), each
  with the same contract: a filled form block is used verbatim, an empty one
  is read from the raw file. Hindi additionally prints measured page numbers
  in its index and requires the date to be typed.
- **Tests**: `test_suite.py` 33/33, `test_series.py` 17/17 (2026-09-08).
  The two API suites need a running server.
- **Delivered build**: `E:\Apps Setup\DocPilot-1.0-beta\` + `.zip`. Rebuild
  command is in the README; strip `_internal\work\` before handing it over.
  Since 2026-09-08 the build **carries its own fonts** (`assets/fonts/`,
  installed per user at startup by `backend/fonts.py`), so the target PC needs
  nothing — that gap caused a real client bug, see the 2026-09-08 section.
- **Dev server**: `python -m uvicorn backend.app:app --port 8930`. Restart it
  after every code change.

### Ground truth lives outside the repo

`%DOCPILOT_SAMPLES%\<Category>\` holds the raw
`Input *.docx` and the published `Output *.docx|.pdf` for each category. These
files ARE the specification — measure against them, never against memory.

**The client replaces them without warning.** On 2026-08-27 all three Prelims
samples were replaced and `tools\make_templates.py` still pointed at the old
names, so every template on disk had been built from the wrong design and every
Prelims output was subtly wrong. When a sample changes: update the path in
`JOBS`, rerun `python tools\make_templates.py <name>.docx`, then re-verify.
`backend\test_series.py` also hard-codes sample paths — update those too.

### How a job flows

`app.py` validates the category (a non-ready one is refused with HTTP 400) and
creates a `Job`. One worker thread drains the queue — Word COM is single-flight,
so parallelism is not an option. `jobs._pipeline(name)` maps the category to a
module through an **explicit whitelist**; `inspect.signature` then decides which
extras that module's `process()` gets (`start_page`, `end_page`, `second_path`,
`cover_date`, `cover_index`, `cover_shorts`). `process()` returns either a path
or `(path, notes)`; stage strings starting with "Cover page added" / "No cover
page" / "Side box" are also collected into the job's notes for the UI.

**Adding a category means touching three places**: `categories.py` (registry),
`jobs._pipeline()` (whitelist — forget it and the job fails only at submit time
with "Unknown pipeline"), and the pipeline module itself. `test_suite.py` now
resolves every ready category, which catches exactly that omission.

### What each pipeline does now

- **Hindi body font**: chain is `Adobe Devanagari -> Baloo 2 -> Nirmala UI`.
  The first is the sample's own face (commercial, auto-detected via Word's
  `FontNames` so Creative-Cloud activations count); the second ships in
  `assets/fonts/`. Headings are Baloo 2 and now provably match the sample
  (1,633 chars against the published 1,617).
- `daily_ca.py` (English) and `daily_ca_hindi.py` — style, then `fitter.py`
  measures real Word pages and fixes pagination, then brand the PDF, then copy
  the fitted DOCX beside it as the Word companion. English then runs
  `cover.py` (2026-08-31): fills the official cover template (date / index /
  shorts — each from the form if typed, else extracted from the raw file) and
  prepends the page to both deliverables. Separate stage on purpose; the body
  pipeline is not touched. Hindi got the same on 2026-08-31 (`cover_hindi.py`):
  date typed (required), topics = `Heading 2` outside tables, shorts = the
  2-column table after the `Heading 1` bar, and **page numbers asked of Word
  on the fitted docx** (table-cell `Range.Information(3)`), then +1 for the
  cover.
- `daily_ca_xylem.py` (2026-08-31) — the OnlyIAS × Xylem IAS edition of the
  English booklet. Imports styler/fitter/cover from the English pipeline,
  then `furnish()` swaps the page furniture for section 2 of the client's
  Xylem cover template (green border, 42.5pt top, footer strip with a PAGE
  field from 2, no watermark/links/promo). Its own fitter spec
  (`_XylemSpec`: content 42.5..710.5) and `cover_xylem.py` (English cover
  machinery on the Xylem template, box 300.8pt). Nothing shared was edited.
  It also owns a **side-box collision guard** (`box_collisions` →
  `unfloat_colliding`, run before fitting on a cheap render and again on the
  fitted PDF): a box another box overprints, or that body text runs through,
  goes back into the flow at full width and the document is re-fitted. No
  PDF stamping stage at all — the page numbers are Word's own field, which
  is why the Xylem Word companion shows them and the English one does not.
- `prelims_ab.py` — answer grid + solutions from one file; title block and
  header code parsed from the FILE NAME; `dt.make_pdf_companion()` renders the
  PDF beside the DOCX.
- `prelims_qb.py` — bilingual booklet: template keeps the official cover /
  instruction pages, questions are poured into a `[[QUESTIONS]]` sentinel,
  page pairs are verified by rendering both languages.
- `fonts.py` (2026-09-08) — not a pipeline, but it runs before all of them:
  a FastAPI startup hook installs the bundled OFL fonts per user. See its
  contract below.
- `mains_ab.py` — unchanged; its category `mains_ab_hindi` was gated off on
  2026-08-31 at the client's request (`ready: False` in `categories.py`,
  shows "template pending" like the other two Mains entries). One flag to
  re-enable; the API and queue both refuse non-ready categories.

### Verification recipes (what "done" means)

Run these from a scratch script; they are how every claim above was checked.

- **Page fill / dead bands**: render the PDF, and per page take the lowest text
  bottom in the MAIN column (`line bbox x0 < 310`, plus images) and compare to
  the content bottom. A side box filling the right edge must not mask an empty
  text column — that mistake cost a whole debugging cycle.
- **Prelims QB pairing**: for each page pair, extract leading `\d+[.)]` numbers
  with a moving window (statement numbers restart inside questions, so an
  unbounded regex is useless) and assert Hindi page N and English page N+1 hold
  the same set. Target: 0 misaligned, reaches Q100.
- **Answer grid**: compare all 25×8 cells against the official booklet's grid.
- **Fonts actually used**: read them from the rendered PDF's spans, not from
  the DOCX and never from a whole-document tally — see the trap below.

### Danger zones (each of these has already bitten)

1. **Restart the server after any code change.** Python imports once; a stale
   uvicorn silently serves old behaviour. Check `(Get-Process -Id <pid>).StartTime`
   against the file's `LastWriteTime` before believing a bug report.
2. **`\b` does not fire next to `_`.** Production names look like
   `Test 08_A_PTS 2708_Indian Polity Part II_...`. Use `(?<![A-Za-z0-9])` and
   `(?!\d)` guards.
3. **`keepNext` inside a table cell** stops Word breaking that row across
   pages, so a tall table jumps whole to the next page and leaves a dead band.
   Never add keeps inside `w:tc`.
4. **A whole-document font tally is misleading** in bilingual files - Hindi
   dominates it. The QB's English pages are Times New Roman 11pt, its Hindi
   pages Nirmala UI 10pt; forcing one font on both cost a third of the density.
5. **Positioned tables and paragraph-anchored floats drag their anchor text**
   to the next page when they do not fit, and they cannot split. That is the
   cause of most "why is this page half empty" reports.
6. **Measurement probes must paginate like the real document** — same template,
   same insertion point, offset derived from the document rather than assumed.

### If you continue the work

Highest-value open items, in order:

1. ~~Prelims QB density~~ - done 2026-10-12 (column-fill layout, 44 pages
   = official). Remaining look difference: 7 English pages snake one
   question into the right column where the official hand-tuned to the half
   point to keep it whole.
2. Daily CA Hindi: give its floats `allowOverlap="0"` (the English styler has
   it, the Hindi pipeline builds its own anchors) so two illustrations cannot
   stack.
3. ~~Speed: reuse one Word instance across a job's renders.~~ **Tried and
   measured 2026-10-11 - do not redo.** 4 page lookups + 4 exports took 80 s
   with a fresh Word each and 92 s with one shared Word: Word starts in a
   second or two here, the time is the layout/export itself. A shared Word
   also intermittently reports "Object is not connected to server" after
   `Range.Information` lookups, which needs recovery code. Reverted from
   `restoration_points/before_word_session_reuse_2026-10-11_211851_IST`.
   Real speed levers are fewer renders (e.g. measured-height packing for the
   QB), not cheaper ones.
4. Mains AB English and Mains QB Bilingual are unbuilt — they need sample pairs
   from the client. Mains AB Hindi is built but gated off (client, 2026-08-31).
5. ~~The side-box collision guard lives only in the Xylem stage.~~ Done
   2026-10-12 for every English edition: see "Side boxes printed on top of
   each other" at the end of this file.
6. `fitter.py`'s note says "tightened by up to half a point" but the squeeze
   can take two half-point steps (11 → 10pt). Wording only; behaviour is as
   designed on 2026-08-26.

### Delivery

`python -m PyInstaller docpilot.spec --noconfirm --distpath dist --workpath build_tmp`
produces `dist\DocPilot-1.0-beta\` (onedir; the spec still carries the "beta"
name). Ship the folder or a zip of it to `E:\Apps Setup\`. Before zipping,
delete `_internal\work\` — the frozen app writes its job scratch there, so a
build you smoke-tested carries your test outputs. Fonts are not bundled: the
target PC needs `python tools\install_fonts.py` once (Baloo 2, Rozha One).
Verify a build by launching the exe (it picks the next free port from 8930)
and running one real job through it, not by trusting the file list.

## Daily Editorial Analysis — 4 categories (2026-09-21) — what it taught

**Delivered 2026-09-21:** `E:\Apps Setup\DocPilot-Setup-1.0-beta.exe` (120 MB,
SHA-256 2CD5CBF2…2CEB1B87) carries the four editorial categories + the web/app
line fix. Verified installed (upgrade over the previous copy): DEA English,
DEA Hindi and CA + DEA English jobs run through the installed app, QR decode +
Karma/Bakbak One working frozen. The 2026-09-09 installer is in
`E:\Apps Setup\previous-versions\`. `DocPilot-1.0-beta.zip` there is the OLD
pre-installer portable build (2026-09-08) - stale.



Modules: `dea_common.py` (reading the raw files: days, QR, dates, roles, the
promo placer, the docx merge), `dea_english.py`, `dea_hindi.py`,
`combo_english.py`, `combo_hindi.py`, `cover_kit.py`. Artwork: `assets/dea/`
(regenerate with `tools/make_dea_assets.py` only when new published samples
arrive). Fonts added to `backend/fonts.py`: **Karma**, **Bakbak One** (OFL).
New deps: `segno` (QR render), `zxing-cpp` (QR decode) - in requirements.

**Hard rule kept: no existing pipeline file was edited.** The new modules
import helpers from `styler.py`, `daily_ca_hindi.py`, `postprocess.py`,
`cover.py`, `cover_hindi.py`, `fitter.py` read-only. The combined editions
call the Daily CA stages themselves (restyle → `fit_document` →
`finalize_pdf` / `_stamp_page`) - the same calls `daily_ca*.process` make,
minus their covers. If you change a Daily CA helper, the combined editions
change with it: run `backend/test_dea.py` too.

Traps, each of which bit during the build:

- **The raw editorial files have no heading styles.** Titles are plain
  paragraphs in red bold 14-18pt; some Hindi section headings are not even
  bold (`ONOE (...) के संबंध में चिंताएँ:`). Roles come from `dea_common`
  (`is_topic_title`, `is_gs_line`, `is_heading` - emphasis OR structure,
  `split_leadin`). Bold is `w:b` OR `w:bCs` (Devanagari = complex script).
- **Multi-day files**: split at date lines (`looks_like_date_line` - a date
  and NOTHING else, so "Published: 17 August 2026" stays text), printed
  oldest first, each later day from a new page. Find a day's first block
  AFTER styling - styling replaces tables and drops empty paragraphs.
- **`insert_pdf` drops the page `/Group`.** The InDesign pages blend in
  `/DeviceCMYK`; without the group their navy renders (8,11,80) instead of
  (27,22,77). Use `dea_common.append_page_keeping_group` (closing page) and
  the extractor's `copy_group`.
- **Word hides overflow in a text box, and a clipped last line still yields
  a partial tick** - so a check that counts ticks passes. Covers are
  therefore MEASURED first: each list rendered alone in a tall unframed box,
  its real bottom read, THEN the frame drawn to that height.
- **Exact line spacing puts the extra leading above the glyphs** (6.8pt on
  the combined English cover). Lists use single spacing + space-after.
- **Word rounds letter-spacing to 1/20pt**: `cover_kit.fit_spacing` aims 3pt
  inside the line and rounds the step up.
- The published English DEA pitch wobbles 13.8-14.4pt (median 14.1): keep
  the styler's "at least 14pt"; only the margins (77.2/65pt) and the button
  height (734.7pt) are DEA-specific.
- `assets/template/promo/pw_webapp_line.png` WAS fully transparent (the
  Daily CA English closing page printed no web/app line). Replaced 2026-09-21
  with the client's go-ahead: same 1060x145 transparent-PNG format, keyed
  from the published line; verified by a before/after render that differs
  only inside the line's box. DEA keeps its own `assets/dea/en_webapp_line.png`.
- The Hindi sample's raw file holds the ENGLISH video's QR; its published
  cover used a Hindi link from outside the file. Hence the Video link block,
  and a job note naming the link printed.

Verification: `backend/test_dea.py` (17 variant cases) + the four samples
through `process()` compared page by page with the published PDFs (covers
within ~1pt; English DEA body page-for-page 8/8).

## Client bugs fixed 2026-09-08 (Hindi CA) — what they taught

Three items came back from the client on a 4-September Hindi booklet. Two
were real; the third was a misreading of their own design.

- **The delivered exe shipped no fonts and no installer.** `install_fonts.py`
  registers fonts **per user on the machine it runs on**, and `docpilot.spec`
  bundled only `assets` + `frontend/dist` — so Rozha One and Baloo 2 existed
  on the dev box and nowhere else. Proven, not guessed: rendering the cover
  with the title font forced to the fallback reproduced the client's
  screenshot pixel for pixel (title y=71.2 against 95.2, overlapping the logo
  at 90.9). **Never rely on a font a build does not carry.** Fonts now live in
  `assets/fonts/` and `backend/fonts.py` installs them at startup.
- **`lineRule="exact"` with a line shorter than the face needs overflows
  UPWARDS.** The fallback title used 1062 twips for a 50pt face, so the glyphs
  climbed 11pt above their own box and through the logo. Both title variants
  now use measured line boxes (1500 for Rozha One, 1600/48pt/72% for the
  fallback → y=94.9 vs 95.2). Any fallback path needs its own measurement,
  not the primary's numbers scaled by eye.
- **Google Docs "Tabs" exports write the tab name into the body**, usually
  followed by a hard page break — "Tab 1" alone on page 2 with the first topic
  pushed to page 3. `_cleanup_source` already stripped `w:pageBreakBefore`
  but **not a `<w:br w:type="page"/>` run**, which is what survived.
  The scan removes the artifact paragraph (and its break) but only *before
  the first real content*, so a genuine "Tab 2" inside an article is safe.
  **It lives in `ooxml.drop_google_tab_frontmatter` and BOTH stylers call it**
  (2026-09-09): English and Xylem read the same raw file as each other, so the
  artifact reaches them too, and one implementation stops the two from
  drifting. Each styler passes its own `style_of` callable; the scan stops at
  the first table, picture, heading/title, or ordinary text.
- **"Use Baloo 2 for the content" contradicts the client's own booklet.**
  Measured from their published PDF: Adobe Devanagari 11.5pt for 33,853 body
  chars, Baloo 2 for 1,429 heading chars. The request came from seeing the
  Nirmala UI fallback. Switching would cost ~3 pages (Baloo 2's natural line
  is 18.1pt at 11.5pt vs Nirmala UI's 15.3pt, and `_reset_ppr` uses
  `atLeast`) and legibility (1,601 glyphs vs 5,025 for the same 128
  Devanagari codepoints).

  **Both of those estimates were wrong, and the pushback was wrong too.**
  Measured on the real booklet, Baloo 2 as the body costs **nothing**: 14
  pages either way, and a *tighter* median page-bottom gap than Nirmala UI
  (19.8pt vs 23.7pt). Telling the client to go and install a commercial font
  was handing my problem back to them. The chain is now
  `Adobe Devanagari -> Baloo 2 -> Nirmala UI`: an exact match where the
  licensed face exists, and otherwise the face the client actually asked
  for, shipped inside the app so nothing has to be installed by hand.
  **Estimate the cost, then measure it before you refuse a request.**

  Do not chase a free substitute: Adobe Devanagari cannot be redistributed,
  the earlier benchmark found none that closes the page gap (Nirmala UI 14pp,
  Noto Sans Devanagari 14pp, Baloo 2 15pp, official 12), and the embedded
  copy in the client's PDF is a WinAnsi CFF **subset** — measuring text
  density from it is invalid (extraction counts glyphs, not characters, so
  it reads ~59% denser than it is) and extracting it would be a licence
  breach. Detection asks **Word** first (`Application.FontNames`), which is
  what makes a Creative-Cloud-activated font count even though it never
  reaches the font registry.

### The font installer's contract (`backend/fonts.py`)

Idempotent by design, because it runs on every start: a family already known
to Windows (HKLM **or** HKCU) is skipped untouched — a licensed or newer copy
on the user's machine always wins. Missing ones are copied to the per-user
font directory, registered under HKCU (no admin), and loaded into the running
session with `AddFontResourceW` + a `WM_FONTCHANGE` broadcast so the very
first job's fresh Word process sees them. Every failure degrades to "not
installed": the pipelines already fall back and say so in the job notes, and
formatting must never fail over a font.

## Layout + robustness round (2026-09-08) — what it taught

- **The Word companion was unusable as an editable file and nobody had
  measured it.** The fitter's per-topic squeeze left **8 body sizes and 8
  line heights** in one document, with **zero page breaks**. So typing
  anywhere reflowed the rest of the booklet and text near a boundary picked
  up a neighbour's size - the client's "editing disturbs the format". The
  fix the client asked for (every topic on a new page) also fixes this,
  because the squeeze rules all describe *a topic that continues overleaf*
  and simply stop firing: the body then keeps its designed size and each
  topic is internally uniform. `_break_between_topics` + `topic_page_break`
  on the FitSpec.
- **A table cannot carry `w:pageBreakBefore`** - Word has no such property
  for tables. The Hindi topics are banner TABLES, so a hairline (2pt, no
  spacing) spacer paragraph carries the break in front of them.
- **Two boundaries can sit back to back** (English opens with a "Topics"
  banner immediately followed by the first topic title). Breaking between
  them strands the banner alone on a page - only break where real content
  separates the two.
- **Cost is real and cannot be squeezed away**: 14->17 (EN), 14->18 (HI),
  13->15 (Xylem). The end-of-topic tails are half-page sized, far beyond
  what half a point of leading recovers, so "don't waste space" and "every
  topic on a new page" genuinely conflict. The published samples pack topics
  together (English 2 of 4 topics start mid-page, Hindi 3 of 4), so this is
  a deliberate divergence from them, on the client's instruction.
- **`docx_to_pdf` had no retry, and Word fails transiently.** One hiccup
  ("Word did not produce the PDF output") threw away a whole job - up to a
  quarter of an hour for the Prelims booklet. Proven transient: the exact
  file that failed converted in 10.5s on its own. Now 3 attempts, an orphan
  WINWORD killed between them (a half-dead one poisons every following
  DispatchEx), a stale part-written PDF removed first, and a final error
  that names the file.
- **The Prelims booklet burned ~32 rounds x 2 renders before admitting
  defeat.** It only ever fails when one question is taller than a page, and
  that is knowable up front: one render per language with every question in
  its own group. Failing in **2 renders instead of 54**, and the message now
  names the question ("Q57 (English)") and points at the two real causes -
  an oversized table/image, or a numbering slip that made one question
  swallow the next ones.

## Daily CA English Xylem (2026-08-31) — what it taught

- **A "front page only" difference wasn't.** The published Xylem pages also
  have a green border, no watermark/header logo, a boxed footer strip, no
  links and no promo page. Diff the whole sample, not the page you were
  pointed at.
- **The client's cover template carried the body furniture too** (section
  2: border, margins, footer strip with the PAGE field in Impact). Rendered
  with a filler page it matched the published pages to 0.1pt - carry parts
  across, never rebuild what the client already drew.
- **Carry only `w:type="default"` header/footer references.** The
  template's section 2 also names a first-page pair: a 7 MB leftover
  watermark and the English link buttons. With `titlePg` removed Word never
  shows them, but they travelled into every booklet until filtered.
- The same raw file feeds two categories; the English extractor
  (`cover.extract_from_raw`) is reused as-is for the Xylem cover.
- **A carried part's shapes resolve THEME references against the destination
  too** (2026-09-08, client-reported). The footer strip's rounded box and
  page-number ring carry no line width of their own - they say
  `<a:lnRef idx="2">`, and the template's theme defines slot 2 as 1.00pt
  where the raw typed files define it as 2.00pt. The strip therefore printed
  at double weight, visible against the sample. `_pin_line_widths()` now
  writes the template theme's real width onto each carried `<a:ln>`, so the
  strip is immune to whatever theme a raw file arrives with. Same lesson as
  the cover styles, one layer deeper: **carrying a part carries neither the
  styles nor the theme it depends on.**
- Verifying image *positions and pixel sizes* is not enough - that is what
  let this through the first time. Compare the **rendered** geometry AND the
  stroke widths (`page.get_drawings()[...]["width"]`).
- New editions = new module importing the shared blocks + a furniture
  swap. No flag inside the English pipeline, so the tested one stays
  byte-for-byte what it was.
- **Two side boxes too close = Word draws them on top of each other and
  flows the second one's anchor paragraph through it** (`tblOverlap=never`
  does not save you). Visible in the very first render, before any fitter
  pass; the fitter's tightening then makes it worse. The English fitter
  has no rule for "box pushed by a box". `daily_ca_xylem.box_collisions()`
  reads the PDF: left border segments at x~320 are the boxes (contiguous
  pieces = one box, overlapping pieces = two boxes), a box starting inside
  another or crossed by a line that begins left of it is a collision; the
  later box is unfloated and the document re-fitted.
- **`jobs._pipeline()` is an explicit whitelist, not a dynamic import.** A
  new category passes every local test and still fails at submit time with
  `Unknown pipeline: …` until its branch is added there. The suite now
  resolves every ready category through it.

## Daily CA Hindi front page (2026-08-31) — what it taught

- **The client's Word template and the published PDF disagree.** The PDF is
  an InDesign export (Letter, Rozha One / Utsaah, one red "विषय सूची", ✓ and
  ➢ bullets, indigo 090C4F); the template is a 612x866pt Word sketch with
  Nirmala UI, "मुख्य/प्रारंभिक परीक्षा" headings and pure navy. The PDF is
  the spec; the template is used for artwork only, and every object is
  re-anchored to the PAGE at measured coordinates. Collapse the body to one
  host paragraph holding all drawings so nothing can spill to page 2.
- **PDF text extraction of Devanagari is not usable for matching** - Word's
  export reorders matras and splits conjuncts (only 3 of 10 headings matched
  by skeleton). Page numbers come from Word COM on the fitted docx instead:
  ~10 s, exact, and cross-checked against the PDF page count.
- **The raw Hindi file has no date anywhere** (only the file name has day +
  month, no year). The Date block is mandatory; nothing is completed.
- **`_body_pr(anchor, **attrs)` with an attr called `anchor`** - a kwarg
  collision that only shows at call time. Name positional params oddly.
- Wingdings 2 `F050` is the thin tick, Wingdings `F0D8` the arrow - verified
  by rendering the candidates, not by the PDF's ToUnicode ("3", "¾").
- The watermark in the sample is the logo drawn at ~9.4% opacity (measured
  from the render: grey 231 on white). It ships pre-baked as an RGBA PNG.
- Nirmala UI is ~15% wider than Utsaah, so a short that fits one line in
  the sample may wrap here; the fit loop absorbs it. Install Utsaah on the
  client's PC for the exact look.
- **`test_suite.pdf_ok` used to assert `"2" in page1_text`.** With a cover
  in front, page 1 has no number and the check only passed on the sample
  because "2026" contains a 2. It now finds the page whose number line is
  exactly "2" within the first two pages. Substring tests on page text hide
  exactly this kind of regression.

## Daily CA English front page (2026-08-31) — what the template taught

- **Leave the template's two continuous sections exactly as they are.**
  Rendered untouched it reproduces the official cover to the point (box top
  360.0pt, buttons 715.1pt, no logo, no page border). Deleting the break
  paragraph - "both sections look identical" - put the page on section 2's
  furniture: logo header, blue page border, everything 5pt lower. Word takes
  a multi-section page's furniture from the FIRST section on that page.
- **`w:type` says how ITS OWN section starts.** Setting nextPage on the
  cover's closing section (meant to push the body to page 2) sent the
  text-box paragraph to page 2 instead. The body's section has no type,
  which already means next page.
- **Copied paragraphs resolve against the DESTINATION's styles.** In the
  Word companion the index lost its ✓/▪ (the numbering lives in the
  `ListParagraph`/`ListParagraph2` styles, absent in the body), the pitch
  tightened (Arial/276 docDefaults) and the date pill dropped 27pt
  (`Heading1` differs). `_carry_styles` copies the whole basedOn chains
  under `cov_` ids with the cover's docDefaults folded into the roots, and
  repoints every cover paragraph before anything is copied.
- **Word HIDES text that overflows a fixed text box** - it is not on the
  page to measure. "Lowest line <= floor" passed with the 6th short missing;
  the fit check counts the bullet lines too.
- **The index box is the official's 303.6pt, not the template's 279.9pt.**
  The client enlarged it by hand; at the template height the 17-Aug index
  would shrink to 14pt while the published cover shows 16pt.
- **Word writes every text box twice** (a `wps:` branch and a VML fallback).
  Fill every `w:txbxContent`, not the first one found, or the two disagree.
- **The raw file's news-table cells are styled `Heading 1` too.** Extracting
  topics must skip paragraphs inside `w:tc`, or the index lists the shorts.
- **The template footers carry a PAGE field; the official cover prints none.**
  Strip the field runs (begin…end) and keep the button drawings.
- **Never save a PyMuPDF/python-docx result over its own input.** The
  pipeline rewrites its deliverable in place; write to a temp file and
  `os.replace`. Same for the Word companion.
- **Re-homing header/footer parts by linking the original part collides on
  partname** (cover images are `image2/3/5.png` like the body's own) and
  produces duplicate zip entries. Clone each part under a fresh partname and
  re-create its relationships.
- **The index-box floor follows the official box, not the anchor
  arithmetic.** The official's 11-line index bottoms out at 656.6pt inside
  a 303.6pt box (floor 660); the template-derived 632pt made the shrink fire
  on an index that fit.
- **Bash heredocs turn a backslash-n escape in Python source into a real
  newline** when the string is written out to a JSX file — every "fix" that
  rejoined the placeholder literal re-split it, and the in-process check read
  the same broken text so it reported success. (This very bullet was mangled
  the same way when first written.) Use the Edit tool, which passes through no
  shell, for any source containing escape sequences.

## Prelims rebuild against the corrected samples (2026-08-27)

The client replaced all three Prelims sample pairs; the old ones were wrong.
`tools/make_templates.py` still pointed at the OLD file names, so every
template on disk had been built from the wrong source. Repointed and rebuilt.
What that round taught, all of it measured:

- **`\b` does not fire next to `_`.** Production files are named
  `Test 08_A_PTS 2708_Indian Polity Part II_...`, so `\bPTS 2708\b` and
  `\btest 08\b` never matched and every booklet printed `[CODE]` and
  `[TEST NAME]`. Use `(?<![A-Za-z0-9])` / `(?!\d)` guards. Cost three bugs in
  two files.
- **The Question Booklet keeps its front and back matter.** The official opens
  with an English cover + Test Instructions and closes with the same in Hindi.
  `carve_body()` keeps them and drops a `[[QUESTIONS]]` sentinel between; the
  pipeline pours questions in there and rewrites the test number / subject.
  The sentinel MUST sit after the section break that closes the title band -
  put it before, and the questions render in that 1-column section (the
  booklet went to 94 pages).
- **A probe must paginate like the real booklet.** `_write_probe` had been
  wiping the template and assuming a fixed 3-paragraph offset; with front
  matter present that mapping is nonsense and the fit loop cannot converge.
  It now finds the same sentinel and derives the offset from the document.
- **Per-language typography, measured from the official PDF's embedded
  fonts**: English pages are Times New Roman 11pt, Hindi pages Nirmala UI
  10pt. A whole-document font tally is misleading - Hindi dominates it. TNR
  is far more compact for Latin (8,711 chars/page against 5,995), so forcing
  one font on both languages cost a third of the density.
- **Leading is tracked per group AND per language** (`line_map[(gi, lang)]`),
  because the two papers run out of page at different questions - tightening
  only the language that spilled is what the official's DTP operator did (its
  leadings vary 245/247/259/288 page to page).
- **Do not seed the grouping from where each language naturally breaks**: the
  two languages break at different questions, so the agreed grouping is their
  intersection (3 per page), nothing ever spills, and nothing is ever
  tightened. Seed at the official's 5-per-page and let the loop earn it.
- Phase 1 of the fit loop settles ~one page pair per round, so its budget has
  to cover one round per pair; running out is what leaves pages carrying a
  single question (phase 2 splits in place instead of repacking). QB takes
  ~15 min because every round is two real Word renders.

## Golden rules for the Test Series categories (added 2026-08-24)

0. **Daily CA English design is frozen.** `styler.py`, `template.py`,
   `ooxml.py`, `postprocess.py` are signed off by the client. New categories get
   their own modules and share only `convert.py` and `docx_tools.py`.
   Exceptions carved out 2026-08-26 on the client's explicit mark-up:
   (a) `daily_ca.py` now calls `fitter.py` — a measured pagination pass that
   runs AFTER the styler and never changes the design itself (rule block
   below); (b) `_style_infobox` floats the blue boxes as right-hand SIDE
   BOXES (`_make_sidebox`: `w:tblpPr` right/margin/text-anchored, 4800 dxa,
   `w:tblOverlap never`; w:tblpPr must sit FIRST in tblPr or Word ignores
   it). Boxes >1800 chars stay full-width — a floating table cannot split
   across pages.
1. **Format, never invent.** If the published sample has text the raw file does
   not (Mains cover page, syllabus box, difficulty ratings, word limits, test
   titles), leave the gap and add a job note. The client was explicit.
2. **python-docx silently skips `w:sdt`.** Google Docs exports hide content in
   content controls — the Mains sample hides 5 of 6 tables and a third of its
   paragraphs. Always walk the body with `docx_tools.iter_block_items()` /
   `block_elements()`, or call `strip_sdt()` first.
3. **A tab is `<w:tab/>`, not text.** `"1.<tab>Which..."` reads back as
   `"1.Which..."`, so question regexes must not require whitespace after the
   number. Cost a real debugging cycle.
4. **Never walk every paragraph over COM.** `docx_tools.locate_pages()` binary
   searches (page numbers are monotonic): 100s → 10s on a 1,200-paragraph file.
5. **The two languages of a bilingual paper do NOT paginate alike** (65 vs 71
   pages in the sample). One page range cannot serve both — apply per file and
   fall back to a full scan, reporting it.

## The fitter (Daily CA English + Hindi, added 2026-08-26)

The client marked up a generated PDF by pen: pages must not end in dead space
while their topic continues overleaf ("Space", "better it starts here"), a
topic's last line must not spill onto the next page ("last line adjust in
above page"), a topic head must never strand at a page bottom ("next page"),
and the font may be decreased "at a decent level" to make this work.
`backend/pipeline/fitter.py` implements exactly that:

- **Static:** delete empty paragraphs; strip source keepNext/keepLines; add
  keep-with-next to red H1s, green Subject lines, blue headings and info-box
  header bars (keepNext must be inserted at the TOP of `w:pPr` or Word
  ignores it).
- **Measured loop** (export via Word → PyMuPDF, max 8 passes): squeeze a
  topic by one half-point step when its page ends >30pt short mid-topic (W),
  when a page opens with the stray end of a structure — detected by indent
  return or a bold section lead below other text (O), or when a topic tail
  occupies <110pt of its last page (T). When a page break is caused by a
  paragraph-anchored float that does not fit (Word moves the ANCHOR paragraph
  to the next page), squeezing cannot help — the float's run is moved into
  the following paragraph instead (F), and structural moves always settle
  before any squeezing starts.
- **Best-state selection:** every pass is scored (dead space + orphan/strand
  penalties + page count) and the best measured layout is restored at the
  end, so a change that reshuffles Word's breaks for the worse is discarded.
- Squeezes scale run sizes by one half-point per step (10pt floor: English
  11→10.5→10, Hindi 11.5→11→10.5→10) and line/after spacing proportionally;
  titles, banner tables and any run ≥14pt are never touched.
- Determinism: same input → same passes → same output. The pass adds ~60-90s
  (one Word export per measure cycle).
- **Floating side boxes push like pictures** (rule B): a positioned table
  that does not fit a page's remainder cannot split, so Word moves it AND
  its anchor text to the next page, leaving a dead band. The fitter detects
  this by matching the next page's first line against each box's header-bar
  / anchor-paragraph text, then shifts the box one body paragraph later in
  the flow (positioning properties untouched). Block indexes change after a
  move — blocks/ranges are rebuilt and the round re-measures immediately.
- **Editions are FitSpecs** (`ENGLISH_SPEC`, `HINDI_SPEC` in fitter.py):
  content box, XML topic-boundary detector (English: red 16pt H1 paragraph /
  24pt banner; Hindi: ≤2×2 banner table containing a CB202F-filled cell —
  the ≥3-col grids also have a red header row, excluded by column count),
  PDF boundary detector (English: red 15-17pt or 23-26pt span; Hindi:
  12.2-13.9pt span in CB202F or white), furniture filter (English: full-page
  watermark; Hindi: the 448×128 centred watermark INSIDE the content area),
  `strip_keeps` (True for English — the raw export's keeps fight pagination;
  False for Hindi — its styler sets keepNext deliberately), and `max_steps`.
  A new edition = a new FitSpec, not a fork of the loop.

## Golden rules

1. **The official sample is ground truth.** Everything in
   `backend/pipeline/styler.py` was reverse-engineered from
   `%DOCPILOT_SAMPLES%\Daily Current Affairs Summary\`
   (input docx + official PDF). The full spec with exact colors/rects/fonts lives
   in `docs/TEMPLATE_SPEC_DailyCA.md` and `assets/template/style_spec.json`.
   Before changing styling, re-check against the sample PDF.
2. **Word COM is single-flight** — `convert.py` holds a global lock; the job queue
   has exactly one worker. Do not parallelize conversions.
3. **Test loop**: `python backend\cli_test.py` runs the sample through the whole
   pipeline (~10 s) → render pages with PyMuPDF and eyeball against the sample.
   API test: `scratchpad` script pattern — submit → poll → download.
4. **Google-Docs-export defects are real**: the styler's `_cleanup_source` +
   in-table `pStyle` strip + `_rebuild_news_table` exist because the raw export
   (a) marks rows `cantSplit`, (b) carries `*Change` revision junk, (c) uses
   negative `tblInd`, and (d) heading styles inside cells stop Word from breaking
   rows across pages (bisect-verified — removing table `pStyle`s dropped 16→14
   pages). Don't remove these steps.

## Key facts (measured, don't re-derive)

- Page: US Letter 612×792pt; margins L57 R52 T76 B78pt; border thin-thick-thin
  `#0070C0` inset 24pt (top/left `thinThickSmallGap`, bottom/right `thickThinSmallGap` —
  Word's naming is direction-relative and was verified empirically).
- Colors: topic `#C00000` 16pt bold; subject `#00B050`; headings/borders `#0070C0`;
  banner fill `#BDD7EE` (Cambria Bold 24pt); page-number circle `#0070C0`
  (center 306,741.8 r11.5, Segoe UI Bold 10 white, **numbering starts at 2**).
- Bullets: Wingdings — lvl0 ❖ (F076) ind 420/420 twips, lvl1 ➢ (F0D8) ind 860/440.
- Body: TNR 11pt justified, line `atLeast` 280tw, after 40tw; cs font Kokila.
- Assets in `assets/template/` were extracted from the sample PDF **with alpha
  masks merged via PIL** (raw `extract_image` loses the PNG SMask → black boxes).
- Promo page is REBUILT in `postprocess.py` (not appended verbatim) so its page
  number continues the sequence; geometry measured from sample page 14.

## Run / build

- Server: `run.bat` → http://localhost:8930 (uvicorn `backend.app:app`).
- Frontend rebuild: `cd frontend && npm run build` (dev: `npm run dev`, proxies
  `/api` to :8930). Theme override for screenshots: `/?theme=dark|light`.
- Screenshots: headless Chrome CLI works —
  `chrome --headless=new --virtual-time-budget=6000 --screenshot=... URL`.
- Desktop build: `python -m PyInstaller docpilot.spec --noconfirm --distpath dist
  --workpath build_tmp` → `dist\DocPilot-1.0-beta\`; see **Delivery** in the
  HANDOFF section before shipping it.
- **Installer (2026-09-09): `python tools/build_installer.py`** →
  `dist\DocPilot-Setup-1.0-beta.exe` (~118 MB). Does icon → version resource →
  PyInstaller → Inno Setup. `--skip-app` reuses `dist\` and takes ~3 min.
  Needs Inno Setup 6; the script finds it per-user or machine-wide and prints
  the download link if it is missing.
- **Upgrades wipe `_internal` first (2026-10-12, `CurStepChanged` in
  `DocPilot.iss`).** Inno never deletes files an older build left behind, so
  installing the 3.11 build over the 3.13 one (21 Sept) gave a mixed tree:
  the app started, listened on 8930, but every request died in uvicorn with
  `httptools has no attribute HttpRequestParser` and the browser showed
  "loading" for ever. `_internal\work` (job scratch) is parked and restored.
  Diagnose this class of fault with `DocPilot.exe --console`.
- The installer is **per-user on purpose** (`PrivilegesRequired=lowest`). Not
  cosmetic: the app writes `_internal\work\jobs` inside its own folder, so the
  install directory must stay writable — Program Files would break every job,
  and a UAC prompt would block clients without admin. Leave it alone.
- The desktop build is **windowed** (`console=False`) and `tkinter` is
  deliberately NOT in the spec's `excludes` - it is the runner UI. Two traps
  that follow: windowed PyInstaller sets `sys.stdout`/`sys.stderr` to `None`,
  so a stray `print()` anywhere in the backend raises (the launcher installs a
  ring-buffer stand-in), and a crash has nowhere to surface (the window shows
  the buffer's tail instead).
- The window is per-monitor DPI aware and multiplies every coordinate and font
  size by the measured scale. Do not add raw pixel constants to `draw()` -
  wrap them in `y()`. This machine runs at 150%, so an unscaled constant looks
  right here and wrong nowhere else.
- **Never turn on `CloseApplications` in the .iss.** It hands the job to Windows
  Restart Manager, which only *asks* an app to shut itself down - a request a
  Python/Tk process never answers - so Setup stalls and then shows "Setup was
  unable to automatically close all applications" (the client hit exactly this).
  `force` is worse: it terminates whatever RM believes holds a file, up to and
  including the user's browser. The fix in place is `AppMutex` - the launcher
  publishes `DocPilot_PW_OnlyIAS_Running` and the installer declares the same
  name, so Setup names the running app and asks for it to be closed. Nothing
  but our own files live in `{app}`, so RM has nothing to contribute.
- **Never test an installer through the Bash tool - use PowerShell.** Inno's
  outer exe spawns a `.tmp` child that inherits the pipe, so bash blocks
  forever after Setup has already finished. The same install that "hung" for
  600s in bash completes in **1.9s** from PowerShell. Two hours were lost
  chasing this as an installer bug; it is a shell artifact.
- Two Setups must never run at once. A silent verification install racing the
  user's interactive one locks `{app}` and produces the same RM dialog.
- Icon: `tools/make_icon.py` → `assets/icon/docpilot.ico`. Drawn as vector
  polygons, not traced, because the shell renders it at 16 px on the taskbar
  where a downscaled screenshot is unreadable. Regenerate rather than hand-edit.
- `dist/` and `build_tmp/` are gitignored as `/dist/` and `/build_tmp/` —
  **anchored**. An unanchored `dist/` also matches `frontend/dist/` and silently
  drops the built UI the server serves, while `index.html` still points at it.

## Adding a new category

1. Get an input+output sample pair from the user, put under `UPSC sample\<cat>\`.
2. Reverse-engineer like Daily CA (see the analysis workflow in
   `docs/TEMPLATE_SPEC_DailyCA.md`): pymupdf spans/drawings/images for the spec,
   docx XML for input anatomy.
3. Implement `backend/pipeline/<category>.py` (reuse styler pieces), flip
   `ready: True` in `backend/categories.py` with its `variant`/naming rule.
4. **Add its branch to `jobs._pipeline()`** — that resolver is an explicit
   whitelist, so without this the category passes every local test and then
   fails at submit time with "Unknown pipeline". Add the module to
   `hiddenimports` in `docpilot.spec` too, or the frozen build breaks the
   same way.
5. Run `test_suite.py`: its category-resolution check covers step 4.

A new *edition* of an existing category (same raw file, different branding —
Xylem is the worked example) is best built as its own module that imports the
existing styler / fitter / cover and swaps only what differs. No flags inside
the signed-off pipeline, so its output stays byte-for-byte what it was.

## Per-category notes

- **daily_ca (English)** — the frozen design: `styler.py` (typography, bullets,
  info-boxes → right-docked side boxes, news table) + `template.py` (page
  furniture) + `postprocess.py` (circle numbers from 2, footer links, promo
  page). `fitter.py` runs between styling and branding. `cover.py` prepends the
  front page to both deliverables. Do not add flags here for other editions.
- **daily_ca_hindi** — its own design module (banner tables, news bar, floats),
  `HINDI_SPEC` in the fitter, `cover_hindi.py` for the front page. The cover's
  index page numbers come from Word COM on the fitted docx, not from PDF text
  (Devanagari extraction is scrambled and unusable for matching).
- **daily_ca_xylem** — OnlyIAS × Xylem IAS edition of the same English raw
  file. Imports the English blocks; `furnish()` carries the body furniture from
  section 2 of `assets/cover/daily_ca_xylem_cover.docx`; `box_collisions()` /
  `unfloat_colliding()` guard the side boxes; `cover_xylem.py` fills the Xylem
  cover template. No `postprocess` stage — numbers are Word's PAGE field.
- **prelims_ab** — template `assets/templates/prelims_ab.docx` (section 1 =
  1 column for title + 25×8 answer grid, section 2 = 2 columns for solutions).
  Input has a flat key then `N. Text Solution:` blocks; output is `N. Ans: (c)`
  + `Exp:`. The published booklets drop the `Source:` citation boxes and keep the
  SuperHint boxes (76 boxes in, 19 out) — `DROP_SOURCE_BOXES` controls this.
  Boxes copied from the A4 source must be re-fitted to the narrow column.
  Title block = constant `SERIES_TITLE` + `Test NN: <subject>` parsed from the
  input FILE NAME (same source as the `PTS 2708` header code); placeholders and
  a job note when the name carries neither. Writes a PDF companion beside the
  DOCX via `dt.make_pdf_companion()`.
  Verified 2026-08-27: 42 pages (EN) / 40 (HI), answer grid identical to the
  official sample, 0 mismatches.
- **prelims_qb** — Hindi page then English page for the same questions. The
  template keeps the official front matter (English cover + Test Instructions +
  title band) and back matter (the same in Hindi); questions are poured into
  the `[[QUESTIONS]]` sentinel and `_fill_test_identity()` rewrites the test
  number / subject inside those pages. Group sizes are *measured* (probe render
  + `pages_of_paragraphs`), leading is tracked per (group, language), and the
  fit loop runs dense-repack then split-in-place so it always terminates with
  the pairing intact. Also writes a PDF companion — for this category the PDF
  is what the client prints. Verified 2026-08-27: 52 pages / 25 page pairs
  against the official 44 / 21, **0 misaligned pairs**, all 100 questions;
  ~15 min because every fit round is two real Word renders.
- **mains_ab** — maps flat paragraphs onto the template's named styles
  (`Questions`, `AboveSpacing`, `Bullet-1`, `Shadding`, `Normal-DK`) and converts
  underline emphasis to the `Boxcolour` character style. Strips the literal
  `प्रश्न N.` prefix because the `Questions` style auto-numbers.
- **Templates** are built by `tools/make_templates.py` from the official outputs:
  body stripped (`strip_body`), orphan media purged, header/footer references
  copied onto every retained section (Word inherits them from the *previous*
  section, so deleting earlier sections silently loses all branding). The
  bilingual QB instead uses `carve_body()`, which KEEPS the front and back
  matter and leaves a `[[QUESTIONS]]` sentinel in between. `JOBS` hard-codes the
  sample paths — repoint and rerun whenever the client replaces a sample.

## Testing — 68 tests (33+17 re-verified locally 2026-09-08)

Start the server first for the two API suites.

| Suite | Covers | Count |
|---|---|---|
| `python backend\test_suite.py` | Daily CA pipeline edge cases + Drive-link logic (mocked gdoc/interstitial/restricted/cancel), the three front pages (extraction, dates, template fill — no Word needed), the Xylem furniture swap + side-box detector, the bundled-font installer, the Google-tabs scan for BOTH stylers, topic page breaks, the Word retry, the Prelims oversize message, and a guard that every ready category resolves to a pipeline | 33 |
| `python backend\test_api.py` | API basics: invalid input, corrupt docx, queue serialization, stop, download gating, real restricted-link | 9 |
| `python backend\test_series.py` | Test Series pipelines: sdt walk, page slicing, answer grid **vs the official booklet**, wrong start page, empty/no-solution files, bilingual matching, Mains styles, numbering integrity | 17 |
| `python backend\test_api_series.py` | Category metadata, page validation, missing second file, disabled categories, docx download, stop on a long job | 9 |

Regression guard: the old pixel-identical baseline for Daily CA English ended
on 2026-08-26 — the client's marked-up sample asked for packed pages, so the
fitter deliberately changes pagination (17-Aug sample: 15 → 13 content pages).
The guard now is behavioural: every topic H1 must sit mid-page or at a page
top WITH body after it, no page (except the last content page) may end more
than ~40pt short while its topic continues, and no page may open with a stray
last line of the previous structure. The design (colors/fonts/banners) must
still match `style_spec.json` — that part stays frozen.

Two stale-test traps to remember: the Hindi sample was renamed with an `Input `
prefix, and enabling a category breaks any test that asserted it returns
"not configured" — use `mains_ab_english` for that check.

## Fonts

`python tools\install_fonts.py` installs **Baloo 2** per-user (no admin).

The load-bearing detail: Baloo 2 ships from Google Fonts as a **variable** font,
and **Word renders a variable font at its default instance regardless of the
named weight requested** — asking for "Baloo 2" + bold, or even for the
"Baloo 2 SemiBold" family, still came out at Regular. The installer therefore
uses `fontTools.varLib.instancer` to cut real static Regular (400) and Bold (700)
instances and rewrites the name table (IDs 1/2/3/4/6/16/17, `OS/2.fsSelection`,
`OS/2.usWeightClass`, `head.macStyle`). After that plain "Baloo 2" + bold behaves
normally and the Hindi topic banners match the printed booklet.

Deliberately **not** installed (commercial — do not source them from unofficial
mirrors): **Adobe Devanagari** (official Hindi body face, ships with Adobe CC)
and **Swiss721 BT** (only appears inside the footer button artwork, which is
embedded as images).

Body-face benchmark on the Hindi CA sample — Nirmala UI 14 pages, Noto Sans
Devanagari 14, Baloo 2 15, official 12 (11 content + a promo page omitted
at the time of the benchmark; standalone Hindi now appends it). No free face
closes the gap; only Adobe Devanagari does.

## Bugs found and fixed while building the Test Series (keep these in mind)

1. **Stripped templates lost all branding.** Word inherits header/footer
   references from the *previous* section; deleting earlier sections leaves the
   survivor with nothing to inherit. `make_templates.py` now copies the refs onto
   every retained `sectPr`.
2. **Orphan media bloat.** Stripping the body leaves the image *relationships*
   behind, so nothing is purged and the template stays 5.6 MB. Must drop unused
   image rels first, then unreferenced `word/media/*` (5.6 MB → 122 KB).
3. **Sub-statements read as questions.** A hard-typed `1. First Past the Post…`
   inside a question would restart numbering. `split_questions` now requires
   question numbers to increase.
4. **Copied boxes overflowed the 2-column page** — tables carried their A4 fixed
   grid; they need `tblW=5000 pct` + cell widths reset + the old `tblGrid` dropped.
5. **Banner merged into the first topic** and later got styled as a heading,
   because a paragraph inserted at the top inherits the next one's style.
6. **Mains read the file twice** (an extra Word pagination pass) because the
   sdt-flatten happened between two `slice_blocks` calls.

## Fixes applied 2026-08-25 (deep-QA round 3 - Prelims/Bilingual)

A page-by-page QA agent diff against the official booklets surfaced these; all
fixed and verified:

1. **Bullets printed as decimal numbers.** Copied paragraphs kept their
   `w:numId`, which is undefined in the destination - Word silently substitutes
   decimal numbering, so ❖ lists printed "1./2." inside answers (actively
   misleading in ranking questions). `docx_tools._copy_numbering_def` now
   carries `w:num`+`w:abstractNum` across with remapped ids; `prelims_ab`
   additionally remaps source *bullet* lists onto the template's own ❖
   definition (numId 3), matching the published look.
2. **Hindi answers rendered centre-aligned at 11pt.** Copied paragraphs carried
   `pStyle=optionindentation`, which the TEMPLATE's styles.xml defines as
   jc=center. All body paragraphs are now flattened to direct formatting
   (official does the same), Devanagari forced to Nirmala UI 10pt on EVERY run -
   runs without `rPr` (the inter-word space runs) otherwise stay TNR 11pt.
3. **Images ran off the paper** (one 468pt picture ended 205pt past the page
   edge). All pictures are capped at 218pt - the official's own maximum.
4. **Option hanging indents collapsed in the bilingual booklet** - continuation
   lines started 70pt left of their text. The official sets tab stops
   426/851/1276 on ~1490 of 1494 paragraphs; we now do the same, with stems at
   432/432 and options 851/851.
5. **pPr children are ORDER-SENSITIVE.** Writing `jc` before `tabs/spacing/ind`
   makes Word silently DROP the later properties (same class of bug as the
   macro keepNext incident). `_normalize_qb` rebuilds pPr in schema order:
   numPr, tabs, spacing, ind, jc, rPr.
6. **Bilingual pairing now verified at BOTH ends.** Checking only each
   question's first paragraph let a group's tail spill unnoticed (probe and
   final agreed - both wrong). The measurer now also checks each group's last
   paragraph; a spilling group first has its leading tightened stepwise
   (271→255→240 - the official's own per-page trick, applied per group so it
   cannot drift) and is split only at the measured page turn. Result: 56 pages,
   0 misaligned pairs, 0 near-empty pages (official: 41 pages, drifts after 8).
   Page breaks are `pageBreakBefore` on the group's first paragraph - explicit
   break paragraphs stranded a fully blank page.
7. Duplicated `उत्तर: X` labels (5 in Hindi) - the label is often SPLIT ACROSS
   RUNS, so per-run matching missed it; now matched on joined text and consumed
   char-by-char across runs.
8. Answer-key grid: 456-twip rows, vertically centred, letters regular (were
   bold 10pt). Title lines and header code are explicit placeholders
   (`[TEST SERIES]`, `[TEST NAME]`, `[CODE]` - "[TEST CODE]" clipped in the
   header's green tab).

## Fixes applied 2026-08-25 (client-reported round 2)

1. **Mains cover page now generated.** The client counted the missing first page
   as a bug, so the "format only" line moved: the tool builds the cover LAYOUT
   (title band `DEEAF6`, syllabus box `FFF2CC`, 3-column question index with
   `D9EAD3` header, page break) and fills what the source provides - the paper
   code parsed from the file name and the real question stems. Facts only a
   human knows are obvious `[SERIES / YEAR]`-style placeholders, and the job
   notes say so. Page count now 7 = official.
2. **Prelims images fixed** (was 10/13 EN, 13/13 HI dangling). Both Prelims
   pipelines now use `copy_block_with_media`.
3. **JFIF-less JPEGs.** The Hindi paper's pictures are valid JPEGs that start
   `FF D8 FF DB` (no APP0 marker). python-docx's sniffer REFUSES them, so
   `get_or_add_image` raises and the old code silently dropped the picture.
   `docx_tools._adopt_image_raw` now adds the blob as a raw OPC part (deduped
   by sha1) and relates it manually. Word renders these fine - verified.
4. **`PTS 2708` header replaced at build time** - with a code parsed from the
   input file name when present, else `[TEST CODE]`. The text lives inside a
   TEXTBOX in the header (plus its mc:AlternateContent fallback), so
   `section.header.paragraphs` never sees it - the raw header part XML is
   walked instead. The advisory lands in the job notes.

## Mains Answer Booklet - fixes applied 2026-08-24

1. **Pictures were dangling.** `append_block` deep-copies the drawing XML but
   `r:embed` still names a relationship id from the *source* document, so Word
   showed broken-image placeholders. `docx_tools.copy_block_with_media()` now
   pulls the image part across and rewrites the id. Mains uses it; the other
   pipelines still call plain `append_block` (see Open items).
2. **PYQ box headings normalised** to `UPSC के विगत वर्षों के प्रश्न (PYQs)`.
   The raw files are inconsistent - one question says "Related UPSC PYQs" in
   English. Client approved the normalisation.
3. **Example lines split out.** In the raw file some `उदाहरण:` lines are
   `<w:br/>` continuations inside their parent bullet. The break sits at the END
   of its run, so the example text starts in the NEXT run - reading `runs[ri:]`
   picks up the tail of the current sentence and never matches. Shadding count
   now 20/20 against the official.
4. **Heading detection now uses boldness, not a keyword list.** These files come
   out of Google Docs with `w:b=1` in docDefaults, so body text carries an
   explicit `b=0` and headings simply inherit. A keyword list missed the ones
   phrased as questions. `Above Spacing` now matches the official 15/15.
5. Stray `उत्तर :` label dropped (carries no content; the official drops it).
6. PYQ box gets the pale-green `E2EFD9` fill; PYQ and demand box headings are
   centred, matching the printed booklet.

Known remaining difference: one bold standalone line
(`संयुक्त परिवार की कार्यात्मक प्रासंगिकता की समाप्ति`) becomes a heading here
but is body text in the official - the official looks inconsistent on that one.

## Open items

### Prelims client feedback on Test 13 (2026-10-11)

Client PDF "PDF generator feedback of solution doc" - six Answer Booklet
issues and a Question Booklet that would not build. Reproduced on the real
Test 13 files (Google Docs links in the feedback; English + Hindi, each holds
questions, key and solutions). Changes only in `prelims_ab.py`,
`prelims_qb.py` and two Test-Series-only helpers in `docx_tools.py`.
Pre-change snapshot: `restoration_points/before_prelims_fixes_2026-10-11_160534_IST`.

- **QB "taller than one page" for Q58-Q100 was a counting bug.** Word's
  `Document.Paragraphs` lists one end-of-row mark per table ROW (verified:
  a 5-row table = 17 paragraphs, 12 `w:p`). Every index after Q57's table
  pointed a few paragraphs early. `dt.word_paragraph_count` counts `w:p` +
  `w:tr` (text-box paragraphs excluded); used by `_emit_group_blocks`,
  `_paragraph_offset` and `dt._para_index_to_block_index`. The last one also
  made a typed **start page** land ~22 paragraphs late on Test 13, silently
  dropping answer-key lines 1-21. Now: 100 questions, 21 page pairs, 0
  misaligned, 44 pages, 10.8 min.
- **Answer line** (`_emit_ans_line`): "N." + real `<w:tab/>` + "Ans: (x)" on
  the body's 446/446 hanging indent + keepNext, as the official booklet. A
  tab typed as text fell to Word's default 0.5" stop.
- **Title**: line 1 cancels the Title style's 886-twip indent; line 2 sits
  between two 4472C4 rules (official `pBdr`). Unknown test name = ruled
  BLANK line to type into (client's choice), not "[TEST NAME]".
- **Lists**: a direct `w:ind` puts every list on the official grid -
  level 0 900/450, nested 1276/376, in a box 360 per level. Nested bullets
  print ➢ (own list, `_arrow_num_id`); the template's level 1 is "–".
  Box lists keep the raw file's own marker (client's choice).
- **Justified line ending in a manual break** is stretched by Word ("PW
  ONLYIAS SUPERHINT" spread across the box): `w:doNotExpandShiftReturn` in
  the booklet's compat settings. Trailing `<w:br/>` stripped from copied
  paragraphs (they left a blank line under list items and box headers).
- **Widows**: the template's docDefaults say `widowControl=0`; every
  paragraph now sets it, like the official. pPr is re-sorted to schema
  order (`_order_ppr`) - the same trap that bit `_normalize_qb`.
- **Image quality**: Word re-encodes every JPEG on PDF export whatever the
  settings say (tested `doNotAutoCompressPictures`: no effect).
  `dt.restore_original_jpegs` swaps the raw file's own bytes back into the
  exported PDF, matched by pixel size + content; both Prelims PDFs now carry
  the source pictures byte-identical. The raw Test 13 picture is itself
  only 369x270 px - sharper than that needs a better image in the raw file.
- **QB tables ran across the column rule / off the paper** (Q57, Q62, Q85):
  the raw A4 grid (e.g. 3960+3960) was copied as is. `_fit_table_to_column`
  sets the table to the column (4672 twips, autofit so a wordy column still
  gets the room - the raw files set no cell widths), un-floats it, left-
  aligns cells (justified narrow cells stretched "3.      Fiat") and gives
  Hindi tables Nirmala UI 10pt. Runs inside `_emit_group_blocks`, so the
  probe measures the same table the booklet prints.
- **QB printed the sample's "PTS 2708" header on every page** of any test;
  it now uses `prelims_ab._replace_header_code` (file name, else [CODE]).
- **The page range is only a hint for the QB**: start page 23 (past the
  questions) and start page 1 gave byte-identical booklets on Test 13 - the
  range-starved file is rescanned whole. It just costs a few minutes.
- **Where "Test 08: Indian Polity (Part II)" comes from**: NOT the raw file
  (the raw `Input English.docx` never names the test). The DTP operator
  typed it; DocPilot reads it from a production file name
  (`Test 08_A_PTS 2708_Indian Polity Part II_...`). Run under that name,
  the raw file gives the published booklet's 42 pages, title, header code
  and grid exactly.
- **QB matched to the Test 08 sample's FORMAT** (client: the page count
  follows the content, 44 is not a target). Snapshot first:
  `restoration_points/before_bilingual_sample_match_2026-10-11_192101_IST`.
  (1) Hindi back cover forced onto its own page (`_break_before_back_matter`;
  only continuous section breaks separated it, so it ran on under Q98-100).
  (2) Packing pass at the end of `measure_groups`: merges neighbouring
  groups (up to 6 a page, the official's fullest), each merge verified by
  the same two renders, plus a final full check that falls back to the
  pre-packing grouping if anything spills. Test 08: 52 -> 48 pages, lone
  pairs Q71/Q72 gone, 0 mismatched. It can only merge FORWARD; Q52 stays
  alone because absorbing it needs boundaries moved back across several
  groups (a measured-height packing redesign would, if ever wanted).
  (3) "(Part II)" restored in the cover / band (`_fill_test_identity`).
  (4) Hindi bold: the raw Hindi paper bolds whole stems and closing lines,
  the official prints them regular - `_unbold_whole_line` drops bold from
  paragraphs >=80% bold (question number kept), partial bold ('कथन I:',
  key terms) stays. Hindi bold 3011 -> 981 chars (official 591).
  Hindi font stays **Nirmala UI 10pt** (measured from the official; the
  client confirmed - Baloo 2 is the Daily CA face only).
- **QB format diff vs the official Test 08 booklet (2026-10-11, same raw
  inputs byte-for-byte)** - stems/options/fonts/columns already matched to
  the point; four real differences fixed in `prelims_qb.py` only
  (snapshot `restoration_points/before_qb_format_match_2026-10-11_222343_IST`):
  (1) `_strip_trailing_breaks` now walks the paragraph's CONTENT from the
  end - the raw stems end in a `<w:br/>` inside the same run as the text,
  which the run-by-run check missed (64 blank lines vs the official's 1).
  (2) Auto-numbered statement lists (347 paragraphs) become typed labels on
  the official grid: number at 432 (78.3 pt), text at 851 (99.5 pt), the
  option-text column. `_list_labels` renders the label exactly as Word
  does - verified against Word's `ListFormat.ListString` on all 837 list
  paragraphs of both Test 08 and both Test 13 papers, 0 mismatches.
  Bullets and unrendered formats stay real lists. (3) `QUESTION_GAP` = 170
  twips before every question not opening its page (official 8.8-10.4 pt
  above a question, ours was 4.5). (4) English pages forced to Times New
  Roman like Hindi pages to Nirmala UI (Q14 (b) printed in Gungsuh);
  symbol faces exempt. All inside `_normalize_qb` / `_emit_group_blocks`,
  so the page-fit probe measures exactly what is printed. Placeholders for
  test no. / subject / code stay (client: core rule unchanged).
- **QB column-fill layout (2026-10-12; client: "see how symmetric the
  official is" - Q42 sat alone on a page pair).** Snapshot first:
  `restoration_points/before_qb_column_fill_2026-10-11_232033_IST`.
  The official keeps every question WHOLE inside one column and tunes the
  leading per COLUMN (Q41-42 at 329 beside Q43-45 at 250) so both columns
  end at the foot (~726pt on every page); on its densest columns it also
  drops paragraph spacing (Hindi 0/0 + 140 above a question, English
  10/10 + 110, instead of 30/30 + 170). `measure_groups` was rewritten:
  (1) every question rendered alone three times (240, 330, 240 tight) ->
  height = a + b*L/240 per question, exact to ~1pt against real pages;
  (2) DP over page pairs: fewest pages, then the most even fill
  (sum of L*^2), each language choosing its own column split; tight
  spacing only where normal cannot reach 245; snaking (no column break)
  only for a question taller than a column or where two whole columns
  cannot fit; (3) render-verify loop brackets each column's leading
  (lo = verified fit, hi = seen to spill) and aims the column at 731pt.
  Columns are separated by a 1pt column-break paragraph (a break inside
  the stem would start the column on a continuation line). Test 08:
  **44 pages / 21 pairs = official**, 0 mismatched, column bottoms median
  728pt, 8.6 min (was 50 pages, lone Q42/Q52/Q100, 31 min).
  Traps found on the way: the probe must finish like build() (back cover on
  its own page, sentinel removed) or the last question measures the cover;
  "reaches the right column" is NOT "taller than a column" (Word balances
  the last page before a continuous break); the three empty section-break
  paragraphs after the questions need 1pt height or a full last page
  pushes them onto a blank page; a Word that hangs on Open keeps the probe
  locked, so the retry kills orphans and renames the probe. Also: a bare
  "67." paragraph followed by its text paragraph is joined (official),
  blank paragraphs closing a question are dropped. Line-spacing specs carry
  tight as a suffix: "250t".
- **Never run a verification job while the app runs one**: two Word
  automation sessions collided ("AttributeError: <unknown>.Range" in
  `pages_of_paragraphs`, which has no retry). The app queue serialises only
  its own jobs.

### Standalone Daily CA Hindi closing ads (2026-10-09)

`daily_ca_hindi.finalize_pdf` now always appends a dedicated closing ad page
before the cover is prepended. It uses the existing English course/book promo
artwork with Hindi header, watermark, grey frame, footer buttons and continuing
black-box page number. The contact row stays above the Hindi footer at 722 pt.
Body layout and the combined/editorial pipelines are unchanged. Like Daily CA
English, the standalone editable Word companion does not carry this PDF-only
ad page. Regression checks are in `backend.test_dea_promos.DailyCAHindiPromoTests`.
Verification: all 15 promo/combined-layout tests pass. Rebuilt the retained
standalone sample from job `9337a5cf3cc8`: 19 pages, closing number 19, both ads,
working footer links, and the original 18 pages identical in text and rendered
pixels. Preview: `output/pdf/Daily CA Hindi With Ads.pdf`. Source app restarted
at `http://127.0.0.1:8930` after the fix.
Pre-change source snapshot: `restoration_points/before_hindi_ca_ads_2026-10-09_152336_IST`.

### Combined Hindi reference formatting update (2026-10-09)

The combined Hindi edition now owns its layout in `combo_hindi_layout.py`;
do not send it through the standalone CA fitter's topic-per-page rules.
The user's reference is `Downloads\Daily CA+Daily Editorial_17 Aug_UPSC_Hindi.pdf`.
Raw inputs are also retained under `DocPilot-samples-2026-10-08/UPSC sample/`.

Google Docs table row minima, cell padding and cell border overrides were
surviving the old styler and inflating pagination. Combined layout removes
those overrides, shades Shorts, pairs secondary portrait infographics,
keeps wide artwork in the flow and uses continuous article layout.
Baloo 2 fallback is calibrated at 9pt body, 10pt headings, 15.2pt leading;
CA copy additionally uses 92% character width. Adobe Devanagari retains its
template size when installed. This matches structure and page count, not
the licensed font's letter shapes. Standalone category defaults are unchanged.

Word bookmarks measure actual section endings. Foundation goes before
Shorts; Srijan after CA and earlier editorial days. There is no final DEA
Foundation ad in this category. Ads respect the footer and may be omitted
when another input has insufficient space. Word ads must be anchored outside
table cells or Word clips them. Frames, circular PAGE fields, badges, ads,
cover and closing artwork now also appear in the editable companion.
QR decoding support is required explicitly rather than silently leaving
cover-only codes inside the body.

Verification: 13 tests in `backend.test_dea_promos` and
`backend.test_combo_hindi_layout`; actual app job `aaf8b8465f91` yields 19
pages, CA 2-12, DEA 13-18, ads 8/12/16, with no detected image/text overlap.
The companion was re-exported by Word: same 19 pages and all three ads;
editable body text was checked against both styled parts. A separate
single-day editorial input yields 18 pages with continuous numbering,
correct cover and no overlaps. The pre-change backup is
`restoration_points/before_formatting_2026-10-09_134715_IST/Current.zip`.

Superseded by the HANDOFF section at the top of this file (2026-09-01) — keep
that one current; the list below is the older backlog.

- ~~Prelims dangling pictures~~ — fixed 2026-08-25, see above.
- Daily CA Hindi: built and live; needs **Adobe Devanagari** (Adobe CC licence)
  to match the official page count — falls back to Nirmala UI, ~3 pages longer.
- Mains Answer Booklet English + Mains QB Bilingual: deferred by the client.
- Drive link limitation: PW files are domain-restricted; keyless download needs
  'Anyone with the link'. A Google-login flow (like Bilingual QC's) would lift
  this — feature, not bug.
- Page count runs 1 page over the hand-made sample (15 vs 14) because designers
  hand-tighten images; acceptable per current scope. (Superseded for Daily CA
  English by the 2026-08-26 fitter: that sample now renders 13 content pages.)
- **Daily CA fitter: three page-break faults fixed (2026-10-12).** Found on
  the 17-Aug CA+DEA English booklet made by the installed app. Snapshot:
  `restoration_points/before_combo_pagebreak_fixes_2026-10-12_113246_IST`.
  (1) A page carrying ONE line ("News Source: Indian Express", booklet p9):
  `_orphan_pages` reads the indent of the second top line, so a one-line
  page was invisible, and the scorer kept the 15-page layout with it.
  `_sliver_pages` (content under ORPHAN_LIMIT, no topic start) now feeds
  rule O and costs PENALTY_SLIVER; the owning topic is squeezed and the
  line comes back. (2) A News-in-Shorts row split right after its title
  (Indonesia Earthquake, p14/15): `_stranded_rows` (red 11pt bold title in
  the bottom ROW_STRAND of a page, row continuing overleaf, row shorter
  than a page) -> `_cant_split_row` puts `w:cantSplit` back on THAT row
  only (the styler strips them all so tall rows can break). Structural,
  one per cycle, judged by score like a box move; PENALTY_ROW outweighs the
  band it leaves behind. (3) Rule T never looked at the LAST topic (no
  boundary after it): its tail is now the last page. Result on 17-Aug: CA
  half 15 -> 14 pages, no one-line page, no split row head. What remains
  are 3-line topic tails (Forest fires, shorts): the body is already at the
  10pt floor, so no rule can pull them back - that is the "topic opens a
  new page" trade, not a bug. Cover: the combined cover lists CA topics +
  editorials only, no shorts - matches the published sample, by design.
- **Daily CA Hindi: input row minima dropped (2026-10-12).** The 17-Aug raw
  file (Google Docs export) pins every row of its 3-column challenges grid
  with `w:trHeight atLeast` 1095-1650tw, so the nine-row table filled page 4
  on its own with inch-tall rows, against the published sample's natural
  rows. `_style_table` now strips `trHeight` from input tables (the combo
  pipeline's `_content_tables` already did). The pipeline's own banner /
  spacer tables keep their heights (built after the styling loop's
  snapshot). 19 -> 18 pages on the sample. English raw files carry no such
  minima. Unchanged and still open: "Word paginates to N but the PDF has
  N+1; index printed without page numbers" on this sample - pre-dates
  today's changes (A/B-verified with the snapshot fitter).
- **Topics flow again - `fitter.TOPIC_PAGE_BREAK = False` (2026-10-12).**
  The user compared the tool's Daily CA Hindi (19 pages, half-empty tails)
  with the published 17-Aug sample (13 pages, topics running on) and chose
  the sample's page breaks for BOTH editions, knowing it undoes the 8 Sept
  "every topic opens a new page" rule. One module constant drives
  `_EnglishSpec`, `_HindiSpec` and (by inheritance) `_XylemSpec`; set it
  back to True to return to the 8 Sept layout - nothing else to touch
  (`test_suite._topic_page_breaks` follows the constant). With it off the
  W/O/T squeeze rules and today's sliver / row rules all have work again.
  Measured: CA+DEA English 23 -> 21 pages (published 22), Daily CA Hindi
  18 -> 15 (published 13, Baloo 2 fallback), no sliver page, no split row
  head on either.
- **Hindi GS line keeps to the next paragraph (2026-10-12).** Once topics
  flowed, the first flowing run left the NITI Aayog banner + its GS line
  alone at the foot of p4 with the topic's text overleaf: the keepNext
  chain ran banner -> GS line and stopped. `_style_gs_line` now sets
  keepNext, so banner -> GS -> first heading -> body move together. Hindi
  sample settles at 16 pages, no stranded head. The installer shipped on
  2026-10-12 12:19 carries everything above.
- **Bold headings inside tables print black (2026-10-12, client).** English
  styler only (`styler._style_cell_paragraph`, used by Daily CA English,
  Xylem, Daily Editorial English and both halves of CA+DEA English): a
  fully-bold cell paragraph and a split "Context:" lead-in in a cell now
  take `CELL_HEAD_COLOR` = 000000 instead of BLUE. Headings outside tables
  stay blue; News-in-Shorts left-cell titles stay red; info-box header bars
  stay white-on-blue. Hindi cells were already ink (231F20). Revert: set
  `CELL_HEAD_COLOR = BLUE`. Snapshot:
  `restoration_points/before_table_bold_black_2026-10-12_123435_IST`.
- **Side boxes printed on top of each other (2026-10-12).** Daily CA
  English 17-Aug p6: "Regulation of Major & Minor Minerals" overprinted
  "About DMF". Not the input (two paragraphs between them): rule F's
  `_move_box_later` slid the first box one paragraph at a time until it
  sat directly before the second; two floats with no text between them
  share an anchor and a position, and Word draws one over the other
  (`tblOverlap never` does not prevent it). Surfaced once topics flowed.
  `_move_box_later` now refuses a shift that lands the box right in front
  of a table; `_separate_stacked_boxes` (static pass) unfloats the second
  of any adjacent pair an input brings in. Verified: no stacked floats,
  p6 clean (Major & Minor full width, DMF beside the text).
- **Version 2.0-beta (2026-10-12).** Bumped in `tools/version_info.py`,
  `desktop_launcher.py`, `docpilot.spec`, `installer/DocPilot.iss`
  (AppVersion, SrcDir, VersionInfoVersion), `tools/build_installer.py`,
  `backend/app.py`. AppId unchanged, so it upgrades a 1.0-beta install in
  place (verified: Apps list shows "DocPilot 2.0-beta", exe ProductVersion
  2.0-beta, app answers). Output: `dist\DocPilot-Setup-2.0-beta.exe`.
  Older notes above that say 1.0-beta describe earlier deliveries.
