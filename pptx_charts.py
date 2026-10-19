"""
core.pptx.charts — Native OpenXML Presentation Chart Engine (v4.0).

Supports:
1. Native <c:chart> generation across 5 chart types:
   - Column (Clustered)
   - Bar (Horizontal Clustered)
   - Line (with data markers)
   - Pie (2D)
   - Doughnut
2. Embedded Excel workbook data package (embedded_xlsx) enabling right-click "Edit Data" in PowerPoint.
3. Strict PPTX_INV_14 conformance:
   - Run-level typography styling (<a:rPr sz="1000" b="0"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:rPr>).
   - Zero color locks, enabling full user color freedom.
   - Widescreen 16:9 responsive placement.
"""

from __future__ import annotations

import os
from typing import Any, Dict, List, Optional, Tuple, Union

import pptx
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import (
    XL_CHART_TYPE,
    XL_LEGEND_POSITION,
    XL_DATA_LABEL_POSITION,
)
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

CHART_TYPE_MAP: Dict[str, XL_CHART_TYPE] = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "column_clustered": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "bar_horizontal": XL_CHART_TYPE.BAR_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "line_markers": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
}

LEGEND_POS_MAP: Dict[str, XL_LEGEND_POSITION] = {
    "bottom": XL_LEGEND_POSITION.BOTTOM,
    "top": XL_LEGEND_POSITION.TOP,
    "right": XL_LEGEND_POSITION.RIGHT,
    "left": XL_LEGEND_POSITION.LEFT,
    "corner": XL_LEGEND_POSITION.CORNER,
}


def build_chart_data(categories: List[str], series_list: List[Dict[str, Any]]) -> CategoryChartData:
    """
    Constructs CategoryChartData with categories and numerical series values.
    """
    chart_data = CategoryChartData()
    chart_data.categories = [str(c) for c in categories]

    for idx, s in enumerate(series_list):
        name = str(s.get("name", f"Series {idx + 1}"))
        raw_vals = s.get("values", [])
        numeric_vals = []
        for v in raw_vals:
            try:
                numeric_vals.append(float(v))
            except (ValueError, TypeError):
                numeric_vals.append(0.0)
        chart_data.add_series(name, tuple(numeric_vals))

    return chart_data


def add_chart_to_slide(
    slide: Any,
    chart_spec: Dict[str, Any],
    theme: Optional[Dict[str, Any]] = None,
    left: Optional[float] = None,
    top: Optional[float] = None,
    width: Optional[float] = None,
    height: Optional[float] = None,
) -> Any:
    """
    Injects a native OpenXML <c:chart> into a PowerPoint slide:
    - Automatically builds embedded Excel workbook data.
    - Applies pure black typography (#000000) for axis labels and numbers (PPTX_INV_14).
    - Enforces Widescreen 16:9 safe boundary alignment.
    """
    raw_type = str(chart_spec.get("type", "column")).lower().strip()
    chart_type = CHART_TYPE_MAP.get(raw_type, XL_CHART_TYPE.COLUMN_CLUSTERED)

    categories = chart_spec.get("categories", ["Cat 1", "Cat 2", "Cat 3"])
    series_list = chart_spec.get("series", [{"name": "Series 1", "values": [10, 20, 30]}])
    chart_data = build_chart_data(categories, series_list)

    # Coordinates (inches)
    pos_left = Inches(float(left if left is not None else chart_spec.get("left", 1.0)))
    pos_top = Inches(float(top if top is not None else chart_spec.get("top", 1.8)))
    pos_w = Inches(float(width if width is not None else chart_spec.get("width", 11.333)))
    pos_h = Inches(float(height if height is not None else chart_spec.get("height", 4.9)))

    # Create Chart Shape (Generates <c:chart> + embedded_xlsx)
    graphic_frame = slide.shapes.add_chart(
        chart_type,
        pos_left,
        pos_top,
        pos_w,
        pos_h,
        chart_data,
    )
    chart = graphic_frame.chart

    # 1. Title
    title_text = chart_spec.get("title")
    if title_text:
        chart.has_title = True
        chart.chart_title.text_frame.text = str(title_text)
        p = chart.chart_title.text_frame.paragraphs[0]
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 0, 0)
    else:
        chart.has_title = False

    # 2. Legend Configuration
    legend_pos = str(chart_spec.get("legend_position", "bottom")).lower()
    if legend_pos in ("none", "false", "off") or chart_spec.get("has_legend") is False:
        chart.has_legend = False
    else:
        chart.has_legend = True
        chart.legend.position = LEGEND_POS_MAP.get(legend_pos, XL_LEGEND_POSITION.BOTTOM)
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(10)
        chart.legend.font.color.rgb = RGBColor(0, 0, 0)

    # 3. Data Labels (Values/Percentages on bars/columns/slices)
    show_labels = bool(chart_spec.get("show_data_labels", False))
    if chart.plots:
        plot = chart.plots[0]
        plot.has_data_labels = show_labels
        if show_labels:
            data_labels = plot.data_labels
            data_labels.font.size = Pt(9.5)
            data_labels.font.color.rgb = RGBColor(0, 0, 0)

    # 4. Axes Typography Formatting (Category Axis & Value Axis)
    # Pie and Doughnut charts do not have standard X/Y axes
    if raw_type not in ("pie", "doughnut"):
        try:
            cat_axis = chart.category_axis
            cat_axis.has_major_gridlines = False
            cat_axis.tick_labels.font.size = Pt(10)
            cat_axis.tick_labels.font.color.rgb = RGBColor(0, 0, 0)
        except Exception:
            pass

        try:
            val_axis = chart.value_axis
            val_axis.has_major_gridlines = True
            val_axis.tick_labels.font.size = Pt(9.5)
            val_axis.tick_labels.font.color.rgb = RGBColor(0, 0, 0)
        except Exception:
            pass

    return graphic_frame


def add_chart_slide(
    prs: Any,
    slide_data: Dict[str, Any],
    theme: Optional[Dict[str, Any]] = None,
) -> Any:
    """
    Creates a full 16:9 Widescreen slide dedicated to a Native OpenXML Chart:
    - Title, Subtitle, and Category Tracker Tag.
    - Centered high-resolution <c:chart> element.
    - Bottom notes / source attribution text.
    """
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    th = theme or {}
    font_heading = th.get("font_heading", "Calibri")
    font_body = th.get("font_body", "Calibri")
    accent_color = th.get("accent_color", RGBColor(37, 99, 235))

    # 1. Slide Header (Category Tag + Title + Subtitle)
    title_text = slide_data.get("title", "Biểu Đồ Số Liệu")
    subtitle_text = slide_data.get("subtitle", "")
    tag_text = slide_data.get("tag") or slide_data.get("section", "DATA & METRICS")

    header_box = slide.shapes.add_textbox(Inches(1.0), Inches(0.5), Inches(11.333), Inches(1.2))
    tf = header_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    # Tag paragraph
    p_tag = tf.paragraphs[0]
    p_tag.text = str(tag_text).upper()
    p_tag.font.name = font_heading
    p_tag.font.size = Pt(10)
    p_tag.font.bold = True
    p_tag.font.color.rgb = accent_color

    # Title paragraph
    p_title = tf.add_paragraph()
    p_title.text = str(title_text)
    p_title.font.name = font_heading
    p_title.font.size = Pt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = RGBColor(0, 0, 0)

    # Subtitle paragraph
    if subtitle_text:
        p_sub = tf.add_paragraph()
        p_sub.text = str(subtitle_text)
        p_sub.font.name = font_body
        p_sub.font.size = Pt(12)
        p_sub.font.color.rgb = RGBColor(71, 85, 105)

    # 2. Extract and render Chart Spec
    chart_spec = slide_data.get("chart") or slide_data
    add_chart_to_slide(slide, chart_spec, theme=th, left=1.0, top=1.8, width=11.333, height=5.0)

    # 3. Presenter Notes (PPTX_INV_08)
    notes = slide_data.get("notes")
    if notes:
        slide.notes_slide.notes_text_frame.text = str(notes)

    return slide
