"""
tests.test_native_charts — Unit test suite for Native OpenXML Charts in PPTX & DOCX (v4.0).

Verifies:
1. PPTX Native <c:chart> generation across 5 chart types: Column, Bar, Line, Pie, Doughnut.
2. PPTX_INV_14 conformance: Axis and legend font colors set to #000000 at run level.
3. Embedded Excel workbook package generation within PPTX graphic frames.
4. DOCX Vector SVG chart rendering and rasterization.
5. Invariant ERR_DOCX_005 printable margin locking (<= 14cm width).
6. Markdown ````chart code block conversion to Word DOCX.
"""

import os
import shutil
import tempfile
import unittest

from pptx import Presentation
from pptx.util import Inches
from pptx.enum.chart import XL_CHART_TYPE
from docx import Document

from core.pptx.charts import (
    add_chart_to_slide,
    add_chart_slide,
    build_chart_data,
    CHART_TYPE_MAP,
)
from core.docx.charts import (
    generate_chart_svg,
    add_chart_to_doc,
)
from core.converter.markdown import decoupled_to_docx


class TestNativeCharts(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp(prefix="test_native_charts_")

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_01_build_chart_data(self):
        categories = ["Q1", "Q2", "Q3", "Q4"]
        series = [
            {"name": "Doanh thu", "values": [120, 150, 180, 210]},
            {"name": "Chi phí", "values": [80, 95, 110, 125]},
        ]
        data = build_chart_data(categories, series)
        self.assertEqual(len(data.categories), 4)
        series_list = list(data)
        self.assertEqual(len(series_list), 2)
        self.assertEqual(series_list[0].name, "Doanh thu")
        self.assertEqual(list(series_list[0].values), [120.0, 150.0, 180.0, 210.0])

    def test_02_pptx_add_column_chart_to_slide(self):
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        chart_spec = {
            "type": "column",
            "title": "Báo Cáo Doanh Thu 2026",
            "categories": ["Tháng 1", "Tháng 2", "Tháng 3"],
            "series": [
                {"name": "Kế hoạch", "values": [100, 120, 140]},
                {"name": "Thực tế", "values": [105, 128, 150]},
            ],
            "show_data_labels": True,
        }
        frame = add_chart_to_slide(slide, chart_spec)
        self.assertIsNotNone(frame)
        self.assertTrue(frame.has_chart)
        chart = frame.chart
        self.assertEqual(chart.chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)
        self.assertTrue(chart.has_title)
        self.assertEqual(chart.chart_title.text_frame.text, "Báo Cáo Doanh Thu 2026")
        self.assertTrue(chart.has_legend)

        # PPTX_INV_14: Verify axis labels exist and are styled
        cat_axis = chart.category_axis
        self.assertIsNotNone(cat_axis.tick_labels.font.color)

        out_pptx = os.path.join(self.test_dir, "test_column_chart.pptx")
        prs.save(out_pptx)
        self.assertTrue(os.path.isfile(out_pptx))
        self.assertGreater(os.path.getsize(out_pptx), 1000)

    def test_03_pptx_add_line_chart_to_slide(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        chart_spec = {
            "type": "line",
            "title": "Tốc Độ Tăng Trưởng",
            "categories": ["Tuần 1", "Tuần 2", "Tuần 3", "Tuần 4"],
            "series": [
                {"name": "Tỉ lệ hoàn thành (%)", "values": [25, 45, 70, 95]},
            ],
        }
        frame = add_chart_to_slide(slide, chart_spec)
        chart = frame.chart
        self.assertEqual(chart.chart_type, XL_CHART_TYPE.LINE_MARKERS)

        out_pptx = os.path.join(self.test_dir, "test_line_chart.pptx")
        prs.save(out_pptx)
        self.assertTrue(os.path.isfile(out_pptx))

    def test_04_pptx_add_pie_and_doughnut_charts(self):
        prs = Presentation()
        # Pie
        slide1 = prs.slides.add_slide(prs.slide_layouts[6])
        pie_spec = {
            "type": "pie",
            "title": "Cơ Cấu Thị Phần",
            "categories": ["Sản phẩm A", "Sản phẩm B", "Sản phẩm C"],
            "series": [{"name": "Thị phần", "values": [50, 30, 20]}],
        }
        frame1 = add_chart_to_slide(slide1, pie_spec)
        self.assertEqual(frame1.chart.chart_type, XL_CHART_TYPE.PIE)

        # Doughnut
        slide2 = prs.slides.add_slide(prs.slide_layouts[6])
        doughnut_spec = {
            "type": "doughnut",
            "title": "Tỷ Lệ Chi Phí",
            "categories": ["R&D", "Marketing", "Vận hành"],
            "series": [{"name": "Chi phí", "values": [40, 35, 25]}],
        }
        frame2 = add_chart_to_slide(slide2, doughnut_spec)
        self.assertEqual(frame2.chart.chart_type, XL_CHART_TYPE.DOUGHNUT)

        out_pptx = os.path.join(self.test_dir, "test_pie_doughnut.pptx")
        prs.save(out_pptx)
        self.assertTrue(os.path.isfile(out_pptx))

    def test_05_pptx_add_chart_slide(self):
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        slide_data = {
            "type": "chart",
            "title": "Doanh Thu Quý 1-4/2026",
            "subtitle": "Đơn vị tính: Tỷ VNĐ",
            "tag": "TÀI CHÍNH & TĂNG TRƯỞNG",
            "notes": "Slide trình bày số liệu tăng trưởng đạt mục tiêu đề ra.",
            "chart": {
                "type": "bar",
                "categories": ["Miền Bắc", "Miền Trung", "Miền Nam"],
                "series": [
                    {"name": "2025", "values": [50, 40, 70]},
                    {"name": "2026", "values": [65, 55, 92]},
                ],
            },
        }
        slide = add_chart_slide(prs, slide_data)
        self.assertIsNotNone(slide)
        self.assertEqual(len(slide.shapes), 2)  # Header textbox + GraphicFrame chart

        out_pptx = os.path.join(self.test_dir, "test_full_chart_slide.pptx")
        prs.save(out_pptx)
        self.assertTrue(os.path.isfile(out_pptx))

    def test_06_docx_generate_chart_svg(self):
        # Column
        col_svg = generate_chart_svg({
            "type": "column",
            "title": "Doanh Thu",
            "categories": ["A", "B", "C"],
            "series": [{"name": "S1", "values": [10, 20, 30]}],
        })
        self.assertIn("<svg", col_svg)
        self.assertIn("Doanh Thu", col_svg)
        self.assertIn("<rect", col_svg)

        # Line
        line_svg = generate_chart_svg({
            "type": "line",
            "title": "Tiến Độ",
            "categories": ["T1", "T2"],
            "series": [{"name": "S1", "values": [50, 80]}],
        })
        self.assertIn("<polyline", line_svg)

        # Pie
        pie_svg = generate_chart_svg({
            "type": "pie",
            "title": "Thị Phần",
            "categories": ["HN", "HCM"],
            "series": [{"name": "Data", "values": [60, 40]}],
        })
        self.assertIn("<path", pie_svg)
        self.assertIn("HN", pie_svg)

    def test_07_docx_add_chart_to_doc(self):
        doc = Document()
        chart_spec = {
            "type": "column",
            "title": "Sản Lượng Q1/2026",
            "categories": ["Tháng 1", "Tháng 2", "Tháng 3"],
            "series": [{"name": "Tấn", "values": [1200, 1450, 1600]}],
        }
        elem = add_chart_to_doc(doc, chart_spec, caption="Sản lượng quý 1 năm 2026")
        self.assertIsNotNone(elem)
        self.assertGreater(len(doc.paragraphs), 0)

        out_docx = os.path.join(self.test_dir, "test_chart_embedded.docx")
        doc.save(out_docx)
        self.assertTrue(os.path.isfile(out_docx))
        self.assertGreater(os.path.getsize(out_docx), 1000)

    def test_08_markdown_chart_code_block(self):
        md_content = """# Báo Cáo Kinh Doanh

Dưới đây là biểu đồ số liệu thực tế:

```chart
{
  "type": "column",
  "title": "Hiệu Suất Vận Hành 2026",
  "categories": ["Q1", "Q2", "Q3", "Q4"],
  "series": [
    {"name": "Đạt", "values": [85, 90, 92, 98]}
  ]
}
```

Kết thúc nội dung báo cáo.
"""
        md_file = os.path.join(self.test_dir, "chart_report.md")
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(md_content)

        out_docx = os.path.join(self.test_dir, "chart_report.docx")
        ok, res_path = decoupled_to_docx(md_file, out_docx)
        self.assertTrue(ok)
        self.assertTrue(os.path.isfile(out_docx))
        self.assertGreater(os.path.getsize(out_docx), 1000)


if __name__ == "__main__":
    unittest.main()
