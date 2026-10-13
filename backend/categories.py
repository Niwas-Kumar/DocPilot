# -*- coding: utf-8 -*-
"""Category registry.

`inputs` tells the UI what to ask for:
    pages      'none' | 'start' | 'start_end'
    files      1 | 2
`output` is the file the user downloads: 'pdf' or 'docx'.
"""

CATEGORIES = [
    {
        "id": "daily_ca_english",
        "label": "Daily Current Affairs Summary (English)",
        "ready": True,
        "variant": "English",
        "pipeline": "daily_ca",
        "pages": "none",
        "files": 1,
        "output": "pdf",
        "cover": True,
        "hint": "Raw typed English Daily CA document. The front page's date, "
                "index and News-in-Shorts are read from the file; type into a "
                "block only to override it.",
    },
    {
        "id": "daily_ca_hindi",
        "label": "Daily Current Affairs Summary (Hindi)",
        "ready": True,
        "variant": "Hindi",
        "pipeline": "daily_ca_hindi",
        "pages": "none",
        "files": 1,
        "output": "pdf",
        "cover": True,
        "cover_date_required": True,
        "hint": "Raw typed Hindi Daily CA document. Front page: type the Date "
                "(the Hindi file carries no date line); the index and "
                "संक्षेप में समाचार are read from the file with their page "
                "numbers - type into a block only to override it. Headings use "
                "Baloo 2, which DocPilot installs itself; for a body that "
                "matches the published booklet exactly, install Adobe "
                "Devanagari (Adobe Creative Cloud) - every job reports which "
                "fonts it actually used.",
    },
    {
        "id": "daily_ca_english_xylem",
        "label": "Daily Current Affairs Summary (English Xylem)",
        "ready": True,
        "variant": "English",
        "pipeline": "daily_ca_xylem",
        "pages": "none",
        "files": 1,
        "output": "pdf",
        "cover": True,
        "hint": "The same raw English Daily CA document, published as the "
                "PW OnlyIAS x Xylem IAS edition: Xylem front page, green page "
                "border, OnlyIAS / Xylem footer strip, no promo page. The "
                "front page's date, index and News-in-Shorts are read from "
                "the file; type into a block only to override it.",
    },
    # ---- Daily Editorial Analysis (added 2026-09-21). Optional keys, read
    # only by these entries (every other category keeps its behaviour):
    #   cover_fields  which front-page blocks the UI shows; default
    #                 ["date", "index", "shorts"]
    #   slots         the two file slots' labels for files == 2
    {
        "id": "daily_editorial_english",
        "label": "Daily Editorial Analysis (English)",
        "ready": True,
        "variant": "English",
        "pipeline": "dea_english",
        "pages": "none",
        "files": 1,
        "output": "pdf",
        "cover": True,
        "cover_date_required": True,
        "cover_fields": ["date", "index", "qr"],
        "hint": "Raw typed English Daily Editorial document. The topics and the "
                "video QR come from the file; the Date is required (the English "
                "file carries none). Paste a video link to replace the file's QR.",
    },
    {
        "id": "daily_editorial_hindi",
        "label": "Daily Editorial Analysis (Hindi)",
        "ready": True,
        "variant": "Hindi",
        "pipeline": "dea_hindi",
        "pages": "none",
        "files": 1,
        "output": "pdf",
        "cover": True,
        "cover_fields": ["date", "index", "qr"],
        "hint": "Raw typed Hindi Daily Editorial document. Date, topics (with "
                "page numbers) and the video QR come from the file - type into a "
                "block only to override it. A file holding several days is "
                "printed oldest first, each day from a new page.",
    },
    {
        "id": "daily_ca_editorial_english",
        "label": "Daily Current Affairs + Daily Editorial Analysis (English)",
        "ready": True,
        "variant": "English",
        "pipeline": "combo_english",
        "pages": "none",
        "files": 2,
        "slots": ["Daily Current Affairs file", "Daily Editorial Analysis file"],
        "output": "pdf",
        "cover": True,
        "cover_fields": ["date", "index", "index2"],
        "hint": "Upload the raw Daily CA file and the raw Daily Editorial file: "
                "one booklet - the Daily CA edition, then the editorials - under "
                "one front page listing both. Date and both indexes come from "
                "the files.",
    },
    {
        "id": "daily_ca_editorial_hindi",
        "label": "Daily Current Affairs + Daily Editorial Analysis (Hindi)",
        "ready": True,
        "variant": "Hindi",
        "pipeline": "combo_hindi",
        "pages": "none",
        "files": 2,
        "slots": ["Daily Current Affairs file", "Daily Editorial Analysis file"],
        "output": "pdf",
        "cover": True,
        "cover_date_required": True,
        "cover_fields": ["date", "index", "index2"],
        "hint": "Upload the raw Hindi Daily CA file and the raw Hindi Daily "
                "Editorial file: one booklet with section badges and one front "
                "page carrying both indexes with page numbers. Type the Date "
                "(the Hindi CA file carries none).",
    },
    {
        "id": "prelims_ab_english",
        "label": "Prelims Test Series (Answer Booklet English)",
        "ready": True,
        "variant": "English",
        "pipeline": "prelims_ab",
        "pages": "start",
        "files": 1,
        "output": "docx",
        "hint": "Give the page where the ANSWER KEY starts (the 1.(C) list). "
                "Everything before it - the question paper - is skipped.",
    },
    {
        "id": "prelims_ab_hindi",
        "label": "Prelims Test Series (Answer Booklet Hindi)",
        "ready": True,
        "variant": "Hindi",
        "pipeline": "prelims_ab",
        "pages": "start",
        "files": 1,
        "output": "docx",
        "hint": "Give the page where the ANSWER KEY starts (the 1.(C) list).",
    },
    {
        "id": "prelims_qb_bilingual",
        "label": "Prelims Test Series (Question Booklet Bilingual)",
        "ready": True,
        "variant": "Bilingual",
        "pipeline": "prelims_qb",
        "pages": "start_end",
        "files": 2,
        "output": "docx",
        "hint": "Upload the English paper and the Hindi paper. Hindi is placed "
                "first, then the same questions in English on the next page.",
    },
    {
        "id": "mains_ab_hindi",
        "label": "Mains Test Series (Answer Booklet Hindi)",
        "ready": False,          # gated off 2026-08-31 (client); pipeline kept intact
        "variant": "Hindi",
        "pipeline": "mains_ab",
        "pages": "start",
        "files": 1,
        "output": "docx",
        "hint": "Model answers are formatted as-is. The cover page, syllabus box "
                "and word limits are not in the raw file, so they are not added.",
    },
    {
        "id": "mains_ab_english",
        "label": "Mains Test Series (Answer Booklet English)",
        "ready": False,
        "variant": "English",
        "pipeline": None,
        "pages": "start",
        "files": 1,
        "output": "docx",
        "hint": "",
    },
    {
        "id": "mains_qb_bilingual",
        "label": "Mains Test Series (Question Booklet Bilingual)",
        "ready": False,
        "variant": "Bilingual",
        "pipeline": None,
        "pages": "start_end",
        "files": 2,
        "output": "docx",
        "hint": "",
    },
]


def get_category(cid):
    for c in CATEGORIES:
        if c["id"] == cid:
            return c
    return None
