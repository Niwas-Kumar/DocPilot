# -*- coding: utf-8 -*-
"""Promo placement regressions; run without Word or a client PDF.

    python -m unittest backend.test_dea_promos -v
"""
import unittest
import tempfile
from pathlib import Path
from unittest.mock import patch

import fitz

from backend.pipeline import combo_hindi_layout as B
from backend.pipeline import dea_common as C
from backend.pipeline import dea_hindi as D
from backend.pipeline import daily_ca_hindi as H


class DailyCAHindiPromoTests(unittest.TestCase):
    def test_finalizer_always_appends_ads_with_continuing_number(self):
        for body_pages, first_number in ((1, 2), (3, 7)):
            with self.subTest(body_pages=body_pages, first_number=first_number):
                with tempfile.TemporaryDirectory() as folder:
                    source = Path(folder) / "body.pdf"
                    output = Path(folder) / "final.pdf"
                    with fitz.open() as doc:
                        for index in range(body_pages):
                            page = doc.new_page(width=612, height=792)
                            # A full last page must still receive a separate ad page.
                            page.insert_text((40, 700), f"Content {index + 1}")
                        doc.save(source)
                    H.finalize_pdf(str(source), str(output), first_number)
                    with fitz.open(output) as doc:
                        self.assertEqual(len(doc), body_pages + 1)
                        for index in range(body_pages):
                            self.assertIn(f"Content {index + 1}", doc[index].get_text())
                        closing = doc[-1]
                        self.assertEqual(closing.get_text().strip(),
                                         str(first_number + body_pages))
                        widths = {info["width"] for info in closing.get_image_info()}
                        self.assertTrue({1396, 1600}.issubset(widths))

    def test_closing_ads_clear_header_and_footer_and_keep_links(self):
        with fitz.open() as doc:
            page = H._build_promo_page(doc, 19)
            promos = [fitz.Rect(info["bbox"]) for info in page.get_image_info()
                      if info["width"] in {1396, 1600, 140, 1060}]
            self.assertEqual(len(promos), 4)
            header_bottom = H.LOGO_RECT[1] + H.LOGO_RECT[3]
            footer_top = min(button[2] for button in H.BUTTONS)
            for rect in promos:
                self.assertGreater(rect.y0, header_bottom)
                self.assertLess(rect.y1, footer_top)
            for index, rect in enumerate(promos):
                for other in promos[index + 1:]:
                    self.assertFalse(rect.intersects(other))
            # Reloading materialises links added to a new page.
            page = doc.reload_page(page)
            uris = {link.get("uri") for link in page.get_links()}
            self.assertTrue({button[5] for button in H.BUTTONS}.issubset(uris))


class PromoPlacementTests(unittest.TestCase):
    def make_document(self, pages=1):
        doc = fitz.open()
        self.addCleanup(lambda: None if doc.is_closed else doc.close())
        for _ in range(pages):
            doc.new_page(width=612, height=792)
        return doc

    def place(self, page, bottom):
        return C.place_banner(page, D.SRIJAN, bottom, D.PAGE_FLOOR,
                              D.SRIJAN_W, D.SRIJAN_CX,
                              D.BANNER_MIN_SCALE, D.BANNER_GAP,
                              min_top=D.BANNER_TOP)

    def test_empty_page_reserves_header_and_border(self):
        page = self.make_document()[0]
        H._stamp_page(page, 2)
        # Centre watermark is page furniture, not article content.
        page.insert_image(fitz.Rect(83.7, 331.9, 531.6, 460.2),
                          filename=str(Path(H.ASSETS) / "watermark_center.png"),
                          keep_proportion=False)
        rect = self.place(page, C.content_bottom(page, D.PAGE_FLOOR))
        self.assertAlmostEqual(rect.y0, D.BANNER_TOP)
        self.assertGreater(rect.y0, H.FRAME_INNER.y0)
        self.assertLessEqual(rect.y1, D.PAGE_FLOOR)
        self.assertAlmostEqual(rect.width, D.SRIJAN_W)

    def test_content_gap_and_normal_size_are_preserved(self):
        page = self.make_document()[0]
        page.insert_text((40, 300), "Last article line", fontsize=12)
        bottom = C.content_bottom(page, D.PAGE_FLOOR)
        rect = self.place(page, bottom)
        self.assertAlmostEqual(rect.y0, bottom + D.BANNER_GAP)
        self.assertAlmostEqual(rect.width, D.SRIJAN_W)
        self.assertAlmostEqual((rect.x0 + rect.x1) / 2, D.SRIJAN_CX)

    def test_short_space_scales_ad_above_minimum(self):
        page = self.make_document()[0]
        rect = self.place(page, 530)
        self.assertAlmostEqual(rect.y1, D.PAGE_FLOOR)
        self.assertGreaterEqual(rect.width / D.SRIJAN_W, D.BANNER_MIN_SCALE)
        self.assertLess(rect.width, D.SRIJAN_W)

    def test_full_page_does_not_receive_an_ad(self):
        page = self.make_document()[0]
        self.assertIsNone(self.place(page, 650))
        self.assertEqual(page.get_image_info(), [])

    def test_shared_default_keeps_existing_english_placement(self):
        page = self.make_document()[0]
        rect = C.place_banner(page, D.SRIJAN, 0, D.PAGE_FLOOR,
                              D.SRIJAN_W, D.SRIJAN_CX)
        self.assertAlmostEqual(rect.y0, 12)

    def test_combined_rejects_negative_and_repeated_day_endings(self):
        doc = self.make_document(5)  # two CA pages, three editorial pages
        notes = B.fill_promos(doc, 2, {'articles': 0, 'ca': 1},
                             {'day0': -1, 'day1': 0, 'day2': 0, 'day3': 2, 'day4': 99})
        self.assertEqual([len(p.get_image_info()) for p in doc], [1, 1, 1, 0, 0])
        self.assertEqual(len(notes), 3)
        self.assertEqual(doc[0].get_image_info()[0]["width"], 3000)  # FOUNDATION before Shorts
        for page in doc:
            for info in page.get_image_info():
                self.assertGreaterEqual(info["bbox"][1], D.BANNER_TOP - 0.001)
                self.assertLessEqual(info["bbox"][3], D.PAGE_FLOOR + 0.001)

    def test_combined_keeps_one_ad_per_valid_day_ending(self):
        doc = self.make_document(5)
        B.fill_promos(doc, 2, {'ca': 1}, {'day0': 0, 'day1': 1})
        self.assertEqual([len(p.get_image_info()) for p in doc], [0, 1, 1, 1, 0])

    def test_standalone_hindi_deduplicates_day_endings(self):
        doc = self.make_document(3)
        D.stamp_and_fill(doc, 2, day_end_pages=[-1, 0, 0, 2])
        self.assertEqual([len(p.get_image_info()) for p in doc], [1, 0, 1])
        self.assertGreaterEqual(doc[0].get_image_info()[0]["bbox"][1], D.BANNER_TOP - 0.001)

    def test_body_lookup_does_not_emit_negative_or_duplicate_offsets(self):
        doc = self.make_document(4)
        info = {"days": [object()] * 4, "topics": ["First", "Later"],
                "day_first_topic": [0, 0, 1, 1]}
        with patch.object(D, "restyle", return_value=info), \
             patch.object(D, "docx_to_pdf"), \
             patch.object(D.fitz, "open", return_value=doc), \
             patch.object(D, "lookup_pages", return_value=({"First": 1, "Later": 3}, "")):
            _, _, result = D.build_body("input.docx", ".", [], {})
        self.assertEqual(result["day_end_pages"], [1])


if __name__ == "__main__":
    unittest.main()
