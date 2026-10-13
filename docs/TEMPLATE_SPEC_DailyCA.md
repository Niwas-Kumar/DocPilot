# DocPilot — Daily Current Affairs Summary: Template Specification

Reverse-engineered from the sample pair (2026-08-19):

- **Input**: `17th August 2026.docx` (raw typed file — plain Calibri, A4, no header/footer, no colors)
- **Output**: `17th August 2026_CA Analysis_UPSC_English_Abhijeet Pratap.pdf` (14 pages, US Letter, produced by Microsoft Word for Microsoft 365, author "Graphics Designer")

Content is **1:1 identical** between input and output — the transformation is purely
deterministic re-styling + template application (no AI/rewriting needed).

## Output filename convention

`{input_stem}_CA Analysis_UPSC_English_Abhijeet Pratap.pdf`

## Page template (every content page, 612×792 pt Letter)

| Element | Detail |
|---|---|
| Background watermark | `assets/template/watermark_page_background.jpeg` (1699×2199) — cream paper texture with faint PW ONLYIAS logo, full-bleed behind text |
| Page border | Thin–thick–thin triple rule, `#0070C0`, all 4 sides (x: 30–582, y: 24–768). Word page-border style `thinThickThinSmallGap`-family |
| Header logo | `header_logo_pw_onlyias.png` centered top, rect ≈ (223.5, 32.4)–(393.5, 74.0) |
| Footer left | `btn_join_whatsapp.png` rect ≈ (57, 727)–(150, 750), hyperlink `https://whatsapp.com/channel/0029Van2ybPIt5rqKRSCzK14` |
| Footer center | Blue circle `#0070C0` (~23 pt Ø, center ≈ (306, 745)) with white **Segoe UI Bold 10 pt** page number. **Numbering starts at 2** (runs 2..15 on the 14-page sample) |
| Footer right | `btn_join_telegram.png` rect ≈ (466.6, 727)–(560, 750), hyperlink `https://t.me/pw_upsc` |
| Margins (text) | left ≈ 57 pt, right edge ≈ 560 pt, first baseline ≈ y 132 |

## Content styling rules (input → output mapping)

| Input construct | Output style |
|---|---|
| Whole document | Calibri → **Times New Roman 11 pt**, justified; **bold runs preserved verbatim** |
| First page start | Light-blue banner `#BDD7EE`, rect (57, 76)–(560, 114), centered bold black "**Topics**" (~20 pt, Cambria Bold) — banner text is template-added, not in input |
| `Heading 1` style (topic title) | Times New Roman **Bold 16 pt `#C00000`** (dark red). Sample has 4 topics |
| `Subject: …` line (first line after topic) | Bold 11 pt **`#00B050`** (green) |
| Non-list fully-bold paragraph ending `:` or `?` (e.g. "More On News:", "Causes of Forest Fires:", "Way Forward:") | Bold 11 pt **`#0070C0`** (blue) section heading |
| `Context: <body…>` inline lead-in | Split: "Context:" becomes its own blue heading line; body continues as normal paragraph |
| List item `ilvl=0` (left indent 36 pt) | ❖ bullet (Wingdings), black, hanging indent (bullet at 57, text at 78) |
| List item `ilvl=1` (left indent 72 pt) | ➢ bullet, black, nested indent |
| 1×1 table (info box, e.g. "About District Mineral Foundation (DMF):") | Boxed callout: `#0070C0` border; first paragraph → **header bar** filled `#0070C0`, white bold text (title left; trailing "(UPSC CSE Prelims 20xx)" part goes right/second line); body rows transparent (paper texture shows through) with ❖/➢ bullets |
| `Title` style ("News In Shorts: 17 August 2026") | Same banner treatment as "Topics" (light-blue `#BDD7EE` band) |
| 6×2 table (News In Shorts) | Bordered 2-col table `#0070C0` grid: left column = **red bold `#C00000`** centered topic name (+ optional image below); right column = blue "Context:" / "About X:" headings + ❖ bullets |
| Images (3 floating `wrapSquare` anchors + 2 inline in table cells) | Preserved at same anchor points; floating images keep square wrap (text flows around, typically right-aligned ~45% width) |
| Hyperlinks (4 in input) | Preserved (blue underline) |
| Final page | Static **promo page** appended (`assets/template/promo_last_page.pdf` or rebuilt from `assets/template/promo/*`): PYQ-32-Years banner, SRIJAN batch banner, customer-service icon, "PW Web/App: https://smart.link/7wwosivoicgd4" + Telegram links, and continues the page-number sequence |

## Fonts observed in output PDF

Times New Roman (body/headings), Segoe UI Bold (page number), Cambria Bold (banner),
Wingdings (❖ bullets), **Kokila** (Devanagari glyphs, e.g. गदा in Cyber Kushti item) — all standard Windows fonts.

## Input file anatomy (what the parser must handle)

- Styles actually used: `Heading 1` (topic titles, already red FF0000 18 pt in input), `normal` (everything else), `Title` (News In Shorts banner line)
- Lists: many different `numId`s but only `ilvl` 0/1 matter (map to ❖/➢)
- 5 tables: 4× 1×1 info boxes + 1× 6×2 News-In-Shorts; **no cell shading in input** — all box styling is applied by the tool
- 5 media images (PNG, some multi-MB); 3 floating (`wp:anchor` + `wrapSquare`), 2 inline inside table cells
- Input page size A4; output is US Letter

## Asset inventory (extracted, in `assets/template/`)

- `watermark_page_background.jpeg` — full-page watermark/texture
- `header_logo_pw_onlyias.png` — header logo (472×115)
- `btn_join_whatsapp.png`, `btn_join_telegram.png` — footer buttons
- `promo_last_page.pdf` — last promo page lifted verbatim from sample
- `promo/` — individual promo-page images (`promo_pyq32`, `promo_srijan`, `customer_service_icon`, `pw_webapp_line`)
- `style_spec.json` — machine-readable version of this spec

## Open questions

1. Category list contains "Daily Current Affairs Summary" twice — Hindi variant? (Output name embeds `_English_`.)
2. No samples yet for the Mains/Prelims Test Series booklet categories — need input+output pairs to spec their templates.
3. "Link" input = Google Drive share link? (Or OneDrive/direct URL?)
4. Is "Abhijeet Pratap" fixed in the filename, or a per-user/per-run setting?
5. Page numbers start at **2** — intentional (cover page exists elsewhere)? Replicate as-is?
