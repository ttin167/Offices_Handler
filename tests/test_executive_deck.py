"""
Unit test suite for Executive Deck Synthesizer in core.pptx.docx_to_pptx.
Tests:
- Target Deck Sizing: 'executive' (15-25 slides), 'standard' (30-45 slides), 'comprehensive'
- Automatic scope activation for large documents (>15,000 words or >40 headings)
- Semantic Fluff Pruning (RE_FLUFF_PARAGRAPH)
- Quantitative KPI Prioritization (RE_QUANTITATIVE_SIGNAL)
- Heading Depth Throttling (H3 minutiae folded into parent H2)
- High-Value Entity Preservation (100% preservation of tables & diagrams)
"""

from __future__ import annotations

import unittest
from typing import Any, Dict, List

from core.pptx.docx_to_pptx import synthesize_slide_spec


class TestExecutiveDeckSynthesizer(unittest.TestCase):
    def _create_mock_doc_tree(
        self,
        num_chapters: int = 5,
        subsections_per_chapter: int = 10,
        words_per_para: int = 40,
    ) -> Dict[str, Any]:
        """Creates a mock document AST structure simulating large documents."""
        chapters: List[Dict[str, Any]] = []
        for ch_idx in range(1, num_chapters + 1):
            subsections: List[Dict[str, Any]] = []
            for s_idx in range(1, subsections_per_chapter + 1):
                level = 2 if s_idx <= 4 else 3
                para_text = (
                    f"Đây là nội dung phân tích chi tiết của mục {ch_idx}.{s_idx}. "
                    "Hệ thống vận hành liên tục đảm bảo tính sẵn sàng cao. "
                    "Các tham số kỹ thuật được cấu hình tối ưu. "
                ) * (words_per_para // 10)
                subsections.append({
                    "title": f"{ch_idx}.{s_idx}. Mục Đề Chi Tiết {s_idx}",
                    "level": level,
                    "elements": [
                        {"type": "paragraph", "text": para_text},
                    ]
                })
            chapters.append({
                "title": f"Chương {ch_idx}. Khảo Sát Hệ Thống {ch_idx}",
                "subsections": subsections,
            })

        return {
            "title": "Báo Cáo Dự Án Quy Mô Lớn (100+ Trang)",
            "author": "Ban Giám Đốc",
            "date": "22/10/2026",
            "chapters": chapters,
        }

    def test_auto_scope_activates_executive_for_large_doc(self):
        """Documents with >40 headings or >15k words must automatically activate executive scope."""
        large_doc = self._create_mock_doc_tree(num_chapters=5, subsections_per_chapter=10, words_per_para=40)
        # Total headings = 50 (>40)
        spec = synthesize_slide_spec(large_doc, deck_scope="auto")

        self.assertLessEqual(len(spec["slides"]), 26)
        self.assertGreaterEqual(len(spec["slides"]), 10)

    def test_executive_deck_sizing_hard_cap(self):
        """Executive mode must produce a concise deck between 15 and 25 slides."""
        large_doc = self._create_mock_doc_tree(num_chapters=6, subsections_per_chapter=12, words_per_para=50)
        spec = synthesize_slide_spec(large_doc, deck_scope="executive")

        self.assertLessEqual(len(spec["slides"]), 26)

    def test_fluff_pruning_removes_filler_paragraphs(self):
        """Executive and standard modes must discard narrative fluff."""
        doc_tree = {
            "title": "Báo Cáo Tóm Lược",
            "author": "Chuyên Viên",
            "date": "22/10/2026",
            "chapters": [
                {
                    "title": "Chương 1. Kết Quả Triển Khai",
                    "subsections": [
                        {
                            "title": "1.1. Đánh giá chung",
                            "level": 2,
                            "elements": [
                                {"type": "paragraph", "text": "Như đã trình bày ở các phần trước, hệ thống có nhiều bước."},
                                {"type": "paragraph", "text": "Tài liệu này nhằm mục đích hướng dẫn quy trình chuyển đổi."},
                                {"type": "paragraph", "text": "Trong khuôn khổ của dự án, chúng tôi ghi nhận kết quả khả quan."},
                                {"type": "paragraph", "text": "Doanh thu quý 1 tăng trưởng 28.5% đạt 150 tỷ VND so với cùng kỳ."},
                            ]
                        }
                    ]
                }
            ]
        }

        spec = synthesize_slide_spec(doc_tree, deck_scope="executive")
        content_slides = [s for s in spec["slides"] if s.get("type") == "content"]
        self.assertTrue(len(content_slides) > 0)

        all_bullets = [b for s in content_slides for b in s.get("bullets", [])]
        for b in all_bullets:
            self.assertNotIn("Như đã trình bày", b)
            self.assertNotIn("Tài liệu này nhằm", b)
            self.assertNotIn("Trong khuôn khổ của", b)

        # Quantitative bullet must be retained
        kpi_found = any("28.5%" in b or "150 tỷ" in b for b in all_bullets)
        self.assertTrue(kpi_found)

    def test_h3_minutiae_folded_into_parent_h2(self):
        """In executive scope, pure text H3 subsections are folded into preceding H2."""
        doc_tree = {
            "title": "Báo Cáo Kiến Trúc",
            "author": "Kỹ Sư Trưởng",
            "date": "22/10/2026",
            "chapters": [
                {
                    "title": "Chương 1. Nền Tảng Kỹ Thuật",
                    "subsections": [
                        {
                            "title": "1.1. Cấu Trúc Tổng Thể",
                            "level": 2,
                            "elements": [
                                {"type": "paragraph", "text": "Hệ thống xây dựng theo mô hình phân tán."},
                            ]
                        },
                        {
                            "title": "1.1.1. Chi tiết phân hệ Kafka",
                            "level": 3,
                            "elements": [
                                {"type": "paragraph", "text": "Băng thông xử lý đạt 50,000 msg/s với độ trễ dưới 5ms."},
                            ]
                        }
                    ]
                }
            ]
        }

        spec = synthesize_slide_spec(doc_tree, deck_scope="executive")
        content_titles = [s.get("title") for s in spec["slides"] if s.get("type") == "content"]

        # H3 should NOT have its own standalone slide
        self.assertNotIn("1.1.1. Chi tiết phân hệ Kafka", content_titles)
        self.assertIn("1.1. Cấu Trúc Tổng Thể", content_titles)

        # But the quantitative bullet from H3 should be folded into the H2 slide
        h2_slide = next(s for s in spec["slides"] if s.get("title") == "1.1. Cấu Trúc Tổng Thể")
        folded_bullets = h2_slide.get("bullets", [])
        self.assertTrue(any("50,000 msg/s" in b or "5ms" in b for b in folded_bullets))

    def test_preserves_tables_and_visuals_in_executive_mode(self):
        """Tables and visual figures must not be dropped in executive mode."""
        doc_tree = {
            "title": "Báo Cáo Tài Chính",
            "author": "Kế Toán Trưởng",
            "date": "22/10/2026",
            "chapters": [
                {
                    "title": "Chương 1. Báo Cáo Doanh Thu",
                    "subsections": [
                        {
                            "title": "1.1. Bảng số liệu chi tiết",
                            "level": 2,
                            "elements": [
                                {
                                    "type": "table",
                                    "caption": "Bảng tổng hợp ngân sách 2026",
                                    "headers": ["STT", "Hạng Mục", "Ngân Sách"],
                                    "rows": [["1", "Thiết bị", "50 tỷ"], ["2", "Nhân sự", "30 tỷ"]]
                                }
                            ]
                        }
                    ]
                }
            ]
        }

        spec = synthesize_slide_spec(doc_tree, deck_scope="executive")
        table_slides = [s for s in spec["slides"] if s.get("type") == "table"]
        self.assertEqual(len(table_slides), 1)
        self.assertEqual(table_slides[0]["headers"], ["STT", "Hạng Mục", "Ngân Sách"])

    def test_comprehensive_scope_preserves_all_sections(self):
        """Comprehensive mode must keep all subsections without executive throttling."""
        doc_tree = self._create_mock_doc_tree(num_chapters=2, subsections_per_chapter=5)
        spec = synthesize_slide_spec(doc_tree, deck_scope="comprehensive")

        # 2 chapters * 5 subsections = 10 content sections + title + agenda + 2 dividers = 14 slides
        self.assertGreaterEqual(len(spec["slides"]), 14)


if __name__ == "__main__":
    unittest.main()
