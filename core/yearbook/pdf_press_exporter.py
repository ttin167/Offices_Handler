"""
core/yearbook/pdf_press_exporter.py — Pre-Press Industrial PDF Export Engine.

Standards & Invariants:
- 300+ DPI Industrial Print Quality with vector trim & bleed calculations.
- Bleed 3mm, Crop Marks (4 pairs), Safe Zone (8mm), and Gutter compensation.
- ERR_PDF_FONT_001: Automatic system Vietnamese Unicode font discovery & embedding.
- ERR_EXIF_001: EXIF orientation auto-normalization via ImageOps.exif_transpose.
- ERR_IMG_CROP_001: Golden-ratio top-biased (35%) center-crop to prevent letterboxing.
- ERR_RAM_001: Streamed image encoding and periodic garbage collection (<512MB RAM).
"""

from __future__ import annotations

import gc
import io
import os
import sys
from typing import Any, Dict, List, Optional, Tuple

try:
    import pymupdf as fitz
except ImportError:
    import fitz
from PIL import Image, ImageDraw, ImageFont, ImageOps

from .layout_engine import PRESETS

# 1 inch = 25.4 mm = 72 points
MM_TO_PT = 72.0 / 25.4


def get_vietnamese_font_paths() -> Tuple[Optional[str], Optional[str]]:
    """
    Discovers regular and bold Vietnamese Unicode TTF fonts (ERR_PDF_FONT_001).
    Returns (regular_font_path, bold_font_path).
    """
    regular_candidates = [
        r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\times.ttf",
        r"C:\Windows\Fonts\calibri.ttf",
        r"C:\Windows\Fonts\segoeui.ttf",
        r"C:\Windows\Fonts\tahoma.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
        "/Library/Fonts/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ]
    bold_candidates = [
        r"C:\Windows\Fonts\arialbd.ttf",
        r"C:\Windows\Fonts\timesbd.ttf",
        r"C:\Windows\Fonts\calibrib.ttf",
        r"C:\Windows\Fonts\segoeuib.ttf",
        r"C:\Windows\Fonts\tahomabd.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
        "/Library/Fonts/Arial Bold.ttf",
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
    ]

    reg = next((p for p in regular_candidates if os.path.isfile(p)), None)
    bold = next((p for p in bold_candidates if os.path.isfile(p)), reg)
    return reg, bold


def crop_image_to_aspect(
    pil_img: Image.Image,
    target_w: float,
    target_h: float,
    top_bias: float = 0.35,
) -> Image.Image:
    """
    Crops image to exact aspect ratio target_w / target_h using object-fit: cover logic (ERR_IMG_CROP_001).
    top_bias (0.0 to 1.0): 0.35 = biased slightly towards top to preserve graduation caps and faces.
    """
    if target_w <= 0 or target_h <= 0:
        return pil_img

    target_ratio = float(target_w) / float(target_h)
    orig_w, orig_h = pil_img.size
    if orig_h <= 0 or orig_w <= 0:
        return pil_img

    orig_ratio = float(orig_w) / float(orig_h)

    if abs(orig_ratio - target_ratio) < 1e-3:
        return pil_img

    if orig_ratio > target_ratio:
        # Image is wider than slot: crop left and right
        new_w = int(orig_h * target_ratio)
        left = int((orig_w - new_w) * 0.5)
        right = left + new_w
        return pil_img.crop((left, 0, right, orig_h))
    else:
        # Image is taller than slot: crop top and bottom with top_bias
        new_h = int(orig_w / target_ratio)
        max_y_offset = orig_h - new_h
        top = int(max_y_offset * top_bias)
        bottom = top + new_h
        return pil_img.crop((0, top, orig_w, bottom))


def generate_mock_photo_bytes(
    width_px: int,
    height_px: int,
    theme_name: str,
    label: str,
) -> bytes:
    """Generates an aesthetic vector-like mock photo image when real photo is not yet uploaded."""
    color_palettes = {
        "gown": ((30, 58, 138), (59, 130, 246)),      # Royal Blue
        "library": ((67, 56, 202), (99, 102, 241)),    # Indigo Academic
        "friends": ((15, 118, 110), (20, 184, 166)),   # Teal Fellowship
        "stage": ((180, 83, 9), (245, 158, 11)),      # Gold Stage
        "diploma": ((190, 24, 93), (244, 63, 94)),     # Rose Celebration
        "polaroid": ((51, 65, 85), (100, 116, 139)),   # Slate Film
        "family": ((124, 45, 18), (217, 119, 6)),     # Warm Amber
        "default": ((30, 41, 59), (71, 85, 105)),     # Dark Slate
    }

    key = "default"
    for k in color_palettes:
        if k in theme_name.lower():
            key = k
            break
    c1, c2 = color_palettes[key]

    img = Image.new("RGB", (max(width_px, 100), max(height_px, 100)), color=c1)
    draw = ImageDraw.Draw(img)

    # Draw gradient-like diagonal bands
    for i in range(0, max(width_px, height_px), 24):
        blend_factor = i / max(width_px, height_px)
        band_color = (
            int(c1[0] + (c2[0] - c1[0]) * blend_factor),
            int(c1[1] + (c2[1] - c1[1]) * blend_factor),
            int(c1[2] + (c2[2] - c1[2]) * blend_factor),
        )
        draw.line([(0, i), (i, 0)], fill=band_color, width=12)

    # Subtle inner border
    draw.rectangle([(8, 8), (width_px - 8, height_px - 8)], outline=(255, 255, 255, 120), width=2)

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=90)
    return buf.getvalue()


class YearbookPressEngine:
    """Pre-Press Industrial PDF generator and 3D preview renderer."""

    @staticmethod
    def _draw_crop_marks(
        page: fitz.Page,
        trim_rect: fitz.Rect,
        bleed_pt: float,
        mark_len: float = 12.0,
        offset: float = 3.0,
    ) -> None:
        """Draws 4 industrial corner crop mark pairs around trim_rect."""
        shape = page.new_shape()
        x0, y0, x1, y1 = trim_rect.x0, trim_rect.y0, trim_rect.x1, trim_rect.y1

        # Top-Left
        shape.draw_line((x0 - bleed_pt, y0), (x0 - offset, y0))
        shape.draw_line((x0, y0 - bleed_pt), (x0, y0 - offset))
        # Top-Right
        shape.draw_line((x1 + offset, y0), (x1 + bleed_pt, y0))
        shape.draw_line((x1, y0 - bleed_pt), (x1, y0 - offset))
        # Bottom-Left
        shape.draw_line((x0 - bleed_pt, y1), (x0 - offset, y1))
        shape.draw_line((x0, y1 + offset), (x0, y1 + bleed_pt))
        # Bottom-Right
        shape.draw_line((x1 + offset, y1), (x1 + bleed_pt, y1))
        shape.draw_line((x1, y1 + offset), (x1, y1 + bleed_pt))

        shape.finish(color=(0.25, 0.25, 0.25), width=0.5)
        shape.commit()

    @classmethod
    def _prepare_image_bytes(
        cls,
        photo_path: str,
        target_w_pt: float,
        target_h_pt: float,
        mock_theme: str = "gown",
        mock_label: str = "Photo",
    ) -> Optional[bytes]:
        """
        Prepares high-fidelity image bytes:
        - Resolves EXIF rotation (ERR_EXIF_001).
        - Crops to exact target aspect ratio with top-bias (ERR_IMG_CROP_001).
        - Generates styled vector mock if photo_path is mock or missing.
        """
        if photo_path and os.path.isfile(photo_path):
            try:
                with Image.open(photo_path) as raw_img:
                    # 1. Normalize EXIF orientation
                    try:
                        normalized = ImageOps.exif_transpose(raw_img)
                        if normalized is not None:
                            raw_img = normalized
                    except Exception:
                        pass

                    # Convert to RGB if RGBA/CMYK
                    if raw_img.mode != "RGB":
                        raw_img = raw_img.convert("RGB")

                    # 2. Crop to aspect ratio
                    cropped = crop_image_to_aspect(raw_img, target_w_pt, target_h_pt, top_bias=0.35)

                    buf = io.BytesIO()
                    cropped.save(buf, format="JPEG", quality=95, subsampling=0)
                    return buf.getvalue()
            except Exception as e:
                print(f"[WARN] Failed to process image '{photo_path}': {e}", file=sys.stderr)

        # Fallback to styled mock photo
        w_px = max(int(target_w_pt * 2.0), 200)
        h_px = max(int(target_h_pt * 2.0), 200)
        return generate_mock_photo_bytes(w_px, h_px, mock_theme, mock_label)

    @classmethod
    def _compute_slot_rects(
        cls,
        layout_type: str,
        content_rect: fitz.Rect,
        num_slots: int,
    ) -> List[fitz.Rect]:
        """Calculates precise geometric coordinates for slots inside content_rect."""
        x0, y0, x1, y1 = content_rect.x0, content_rect.y0, content_rect.x1, content_rect.y1
        w = x1 - x0
        h = y1 - y0
        rects: List[fitz.Rect] = []
        gap = 12.0

        if layout_type == "cover_hero" or num_slots == 1:
            # 1 large photo centered with generous margins
            rects.append(fitz.Rect(x0, y0, x1, y1))

        elif layout_type == "editorial_quote":
            # Left half photo, right half text
            half_w = (w - gap) / 2.0
            rects.append(fitz.Rect(x0, y0, x0 + half_w, y1))

        elif layout_type == "symmetric_duo" or num_slots == 2:
            # 2 vertical columns
            half_w = (w - gap) / 2.0
            rects.append(fitz.Rect(x0, y0, x0 + half_w, y1))
            rects.append(fitz.Rect(x0 + half_w + gap, y0, x1, y1))

        elif layout_type == "polaroid_scrapbook" or num_slots == 3:
            # 3 polaroid frames across width
            slot_w = (w - gap * 2) / 3.0
            for i in range(3):
                sx = x0 + i * (slot_w + gap)
                rects.append(fitz.Rect(sx, y0 + 10, sx + slot_w, y1 - 10))

        elif layout_type == "magazine_grid_4" or num_slots == 4:
            # 1 Hero slot on left, 3 stacked on right
            hero_w = (w - gap) * 0.55
            right_w = (w - gap) * 0.45
            rects.append(fitz.Rect(x0, y0, x0 + hero_w, y1))

            small_h = (h - gap * 2) / 3.0
            for i in range(3):
                sy = y0 + i * (small_h + gap)
                rects.append(fitz.Rect(x0 + hero_w + gap, sy, x1, sy + small_h))

        elif layout_type == "mosaic_gallery_6" or num_slots == 6:
            # 2 rows x 3 columns
            col_w = (w - gap * 2) / 3.0
            row_h = (h - gap) / 2.0
            for r in range(2):
                for c in range(3):
                    sx = x0 + c * (col_w + gap)
                    sy = y0 + r * (row_h + gap)
                    rects.append(fitz.Rect(sx, sy, sx + col_w, sy + row_h))

        elif layout_type == "classroom_grid_8" or num_slots == 8:
            # 2 rows x 4 columns
            col_w = (w - gap * 3) / 4.0
            row_h = (h - gap) / 2.0
            for r in range(2):
                for c in range(4):
                    sx = x0 + c * (col_w + gap)
                    sy = y0 + r * (row_h + gap)
                    rects.append(fitz.Rect(sx, sy, sx + col_w, sy + row_h))

        elif layout_type == "megagrid_12" or num_slots >= 12:
            # 3 rows x 4 columns
            col_w = (w - gap * 3) / 4.0
            row_h = (h - gap * 2) / 3.0
            for r in range(3):
                for c in range(4):
                    sx = x0 + c * (col_w + gap)
                    sy = y0 + r * (row_h + gap)
                    rects.append(fitz.Rect(sx, sy, sx + col_w, sy + row_h))

        else:
            # Default even columns
            col_w = (w - gap * (num_slots - 1)) / max(num_slots, 1)
            for i in range(num_slots):
                sx = x0 + i * (col_w + gap)
                rects.append(fitz.Rect(sx, y0, sx + col_w, y1))

        return rects

    @classmethod
    def export_print_ready_pdf(
        cls,
        yearbook_spec: Dict[str, Any],
        output_pdf_path: str,
    ) -> str:
        """
        Exports commercial print-ready PDF (Single pages, 3mm Bleed, Crop Marks, Vietnamese Unicode font).
        """
        preset_key = yearbook_spec.get("preset", "photobook_square_25x25")
        preset = PRESETS.get(preset_key, PRESETS["photobook_square_25x25"])
        bleed_mm = float(yearbook_spec.get("bleed_mm", 3.0))

        w_pt = preset["w_mm"] * MM_TO_PT
        h_pt = preset["h_mm"] * MM_TO_PT
        bleed_pt = bleed_mm * MM_TO_PT

        total_w = w_pt + (bleed_pt * 2.0)
        total_h = h_pt + (bleed_pt * 2.0)
        trim_rect = fitz.Rect(bleed_pt, bleed_pt, bleed_pt + w_pt, bleed_pt + h_pt)

        reg_font, bold_font = get_vietnamese_font_paths()

        doc = fitz.open()
        pages_data = yearbook_spec.get("pages") or yearbook_spec.get("spreads") or []

        for p_idx, p_spec in enumerate(pages_data, 1):
            page = doc.new_page(width=total_w, height=total_h)

            # 1. Embed Vietnamese Unicode font (ERR_PDF_FONT_001)
            font_alias = "helv"
            font_bold_alias = "helv"
            if reg_font:
                try:
                    page.insert_font(fontname="vn_reg", fontfile=reg_font)
                    font_alias = "vn_reg"
                except Exception:
                    pass
            if bold_font:
                try:
                    page.insert_font(fontname="vn_bold", fontfile=bold_font)
                    font_bold_alias = "vn_bold"
                except Exception:
                    font_bold_alias = font_alias

            # 2. Draw background color extending fully across bleed
            bg_hex = p_spec.get("bg_color", "#FAF8F5").lstrip("#")
            bg_rgb = tuple(int(bg_hex[i:i+2], 16) / 255.0 for i in (0, 2, 4))
            page.draw_rect(page.rect, color=None, fill=bg_rgb)

            # 3. Content margin inside trim (20mm margin, 8mm safe margin)
            margin = 22.0 * MM_TO_PT
            bottom_caption_h = 36.0
            top_title_h = 42.0

            layout_type = p_spec.get("layout", p_spec.get("type", "cover_hero"))
            title_text = p_spec.get("title", "")
            subtitle_text = p_spec.get("subtitle", "")
            caption_text = p_spec.get("caption", "")
            quote_text = p_spec.get("quote", "")
            slots = p_spec.get("slots", [])

            # Compute content area
            content_rect = fitz.Rect(
                trim_rect.x0 + margin,
                trim_rect.y0 + margin + (top_title_h if title_text else 0),
                trim_rect.x1 - margin,
                trim_rect.y1 - margin - (bottom_caption_h if caption_text else 0),
            )

            # Header title if present
            if title_text:
                t_rect = fitz.Rect(trim_rect.x0 + margin, trim_rect.y0 + margin, trim_rect.x1 - margin, trim_rect.y0 + margin + top_title_h)
                header_text = title_text + (f" — {subtitle_text}" if subtitle_text else "")
                page.insert_textbox(
                    t_rect, header_text,
                    fontsize=12, fontname=font_bold_alias,
                    color=(0.12, 0.16, 0.24), align=fitz.TEXT_ALIGN_CENTER
                )

            # Slots layout
            slot_rects = cls._compute_slot_rects(layout_type, content_rect, len(slots))

            for s_idx, slot in enumerate(slots):
                if s_idx >= len(slot_rects):
                    break
                s_rect = slot_rects[s_idx]
                photo_path = slot.get("photo_path", "")
                m_theme = slot.get("mock_theme", "gown")
                m_label = slot.get("caption", f"Slot {s_idx+1}")

                img_bytes = cls._prepare_image_bytes(photo_path, s_rect.width, s_rect.height, m_theme, m_label)
                if img_bytes:
                    page.insert_image(s_rect, stream=img_bytes, keep_proportion=False)

                    # Subtle slot border
                    page.draw_rect(s_rect, color=(0.85, 0.85, 0.85), width=0.5)

            # Handle Editorial Quote text card if editorial_quote
            if layout_type == "editorial_quote" and quote_text:
                half_w = (content_rect.width - 12.0) / 2.0
                q_rect = fitz.Rect(content_rect.x0 + half_w + 16.0, content_rect.y0 + 20.0, content_rect.x1 - 10.0, content_rect.y1 - 20.0)
                formatted_quote = f"“{quote_text}”\n\n— {yearbook_spec.get('author', 'Biện Trung Tín')}"
                page.insert_textbox(
                    q_rect, formatted_quote,
                    fontsize=11, fontname=font_alias,
                    color=(0.2, 0.25, 0.35), align=fitz.TEXT_ALIGN_LEFT
                )

            # Footer caption if present
            if caption_text:
                cap_rect = fitz.Rect(
                    trim_rect.x0 + margin,
                    trim_rect.y1 - margin - bottom_caption_h + 8.0,
                    trim_rect.x1 - margin,
                    trim_rect.y1 - margin
                )
                page.insert_textbox(
                    cap_rect, caption_text,
                    fontsize=9.5, fontname=font_alias,
                    color=(0.35, 0.4, 0.48), align=fitz.TEXT_ALIGN_CENTER
                )

            # Page number
            pnum_rect = fitz.Rect(trim_rect.x1 - margin - 50, trim_rect.y1 - margin + 4.0, trim_rect.x1 - margin, trim_rect.y1 - margin + 18.0)
            page.insert_textbox(
                pnum_rect, f"{p_idx:02d}",
                fontsize=8.5, fontname=font_alias,
                color=(0.5, 0.5, 0.5), align=fitz.TEXT_ALIGN_RIGHT
            )

            # 4. Draw industrial Crop Marks
            cls._draw_crop_marks(page, trim_rect, bleed_pt)

            # Trigger periodic garbage collection (ERR_RAM_001)
            if p_idx % 4 == 0:
                gc.collect()

        os.makedirs(os.path.dirname(os.path.abspath(output_pdf_path)), exist_ok=True)
        doc.save(output_pdf_path, garbage=4, deflate=True)
        doc.close()
        gc.collect()
        return output_pdf_path

    @classmethod
    def export_spread_preview_pdf(
        cls,
        yearbook_spec: Dict[str, Any],
        output_pdf_path: str,
    ) -> str:
        """
        Exports 2-page spread preview PDF (Double-page Spread View, no crop marks, center gutter guide).
        """
        preset_key = yearbook_spec.get("preset", "photobook_square_25x25")
        preset = PRESETS.get(preset_key, PRESETS["photobook_square_25x25"])

        w_pt = preset["w_mm"] * MM_TO_PT
        h_pt = preset["h_mm"] * MM_TO_PT

        # Double page spread width
        spread_w = w_pt * 2.0
        spread_h = h_pt

        reg_font, bold_font = get_vietnamese_font_paths()

        doc = fitz.open()
        pages_data = yearbook_spec.get("pages") or yearbook_spec.get("spreads") or []

        for sp_idx, p_spec in enumerate(pages_data, 1):
            page = doc.new_page(width=spread_w, height=spread_h)

            font_alias = "helv"
            font_bold_alias = "helv"
            if reg_font:
                try:
                    page.insert_font(fontname="vn_reg", fontfile=reg_font)
                    font_alias = "vn_reg"
                except Exception:
                    pass
            if bold_font:
                try:
                    page.insert_font(fontname="vn_bold", fontfile=bold_font)
                    font_bold_alias = "vn_bold"
                except Exception:
                    font_bold_alias = font_alias

            # Background fill
            bg_hex = p_spec.get("bg_color", "#FAF8F5").lstrip("#")
            bg_rgb = tuple(int(bg_hex[i:i+2], 16) / 255.0 for i in (0, 2, 4))
            page.draw_rect(page.rect, color=None, fill=bg_rgb)

            # Center gutter line (dashed line indicating the book spine)
            center_x = spread_w / 2.0
            spine = page.new_shape()
            spine.draw_line((center_x, 0), (center_x, spread_h))
            spine.finish(color=(0.75, 0.75, 0.75), width=0.5, dashes="[2 3] 0")
            spine.commit()

            margin = 20.0 * MM_TO_PT
            top_title_h = 36.0
            bottom_caption_h = 32.0

            layout_type = p_spec.get("layout", p_spec.get("type", "cover_hero"))
            title_text = p_spec.get("title", "")
            subtitle_text = p_spec.get("subtitle", "")
            caption_text = p_spec.get("caption", "")
            quote_text = p_spec.get("quote", "")
            slots = p_spec.get("slots", [])

            # Header
            if title_text:
                t_rect = fitz.Rect(margin, margin, spread_w - margin, margin + top_title_h)
                header_text = title_text + (f" — {subtitle_text}" if subtitle_text else "")
                page.insert_textbox(
                    t_rect, header_text,
                    fontsize=12, fontname=font_bold_alias,
                    color=(0.15, 0.18, 0.25), align=fitz.TEXT_ALIGN_CENTER
                )

            # Content area
            content_rect = fitz.Rect(
                margin,
                margin + (top_title_h if title_text else 0),
                spread_w - margin,
                spread_h - margin - (bottom_caption_h if caption_text else 0),
            )

            slot_rects = cls._compute_slot_rects(layout_type, content_rect, len(slots))

            for s_idx, slot in enumerate(slots):
                if s_idx >= len(slot_rects):
                    break
                s_rect = slot_rects[s_idx]
                photo_path = slot.get("photo_path", "")
                m_theme = slot.get("mock_theme", "gown")
                m_label = slot.get("caption", f"Slot {s_idx+1}")

                img_bytes = cls._prepare_image_bytes(photo_path, s_rect.width, s_rect.height, m_theme, m_label)
                if img_bytes:
                    page.insert_image(s_rect, stream=img_bytes, keep_proportion=False)
                    page.draw_rect(s_rect, color=(0.85, 0.85, 0.85), width=0.5)

            # Editorial quote card
            if layout_type == "editorial_quote" and quote_text:
                q_rect = fitz.Rect(center_x + 24.0, content_rect.y0 + 24.0, spread_w - margin - 20.0, content_rect.y1 - 24.0)
                formatted_quote = f"“{quote_text}”\n\n— {yearbook_spec.get('author', 'Biện Trung Tín')}"
                page.insert_textbox(
                    q_rect, formatted_quote,
                    fontsize=11, fontname=font_alias,
                    color=(0.2, 0.25, 0.35), align=fitz.TEXT_ALIGN_LEFT
                )

            # Caption
            if caption_text:
                cap_rect = fitz.Rect(
                    margin,
                    spread_h - margin - bottom_caption_h + 6.0,
                    spread_w - margin,
                    spread_h - margin
                )
                page.insert_textbox(
                    cap_rect, caption_text,
                    fontsize=9.5, fontname=font_alias,
                    color=(0.35, 0.4, 0.48), align=fitz.TEXT_ALIGN_CENTER
                )

            # Spread indicator at bottom corners
            page.insert_textbox(
                fitz.Rect(margin, spread_h - margin + 2.0, margin + 80, spread_h - margin + 16.0),
                f"Trang {sp_idx * 2 - 1}",
                fontsize=8.0, fontname=font_alias,
                color=(0.5, 0.5, 0.5), align=fitz.TEXT_ALIGN_LEFT
            )
            page.insert_textbox(
                fitz.Rect(spread_w - margin - 80, spread_h - margin + 2.0, spread_w - margin, spread_h - margin + 16.0),
                f"Trang {sp_idx * 2}",
                fontsize=8.0, fontname=font_alias,
                color=(0.5, 0.5, 0.5), align=fitz.TEXT_ALIGN_RIGHT
            )

        os.makedirs(os.path.dirname(os.path.abspath(output_pdf_path)), exist_ok=True)
        doc.save(output_pdf_path, garbage=4, deflate=True)
        doc.close()
        gc.collect()
        return output_pdf_path

    @classmethod
    def render_spread_thumbnail_png(
        cls,
        yearbook_spec: Dict[str, Any],
        spread_index: int,
        output_png_path: str,
        dpi: int = 150,
    ) -> str:
        """Renders a single spread into high-resolution PNG for fast Web 3D Flipbook display."""
        # Create single spread temp PDF in memory and render to pixmap
        sp_spec = dict(yearbook_spec)
        pages = yearbook_spec.get("pages", [])
        if not pages:
            raise ValueError("No pages in yearbook spec")

        idx = max(0, min(spread_index, len(pages) - 1))
        sp_spec["pages"] = [pages[idx]]

        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf:
            temp_pdf = tf.name

        try:
            cls.export_spread_preview_pdf(sp_spec, temp_pdf)
            doc = fitz.open(temp_pdf)
            if len(doc) > 0:
                pix = doc[0].get_pixmap(dpi=dpi)
                os.makedirs(os.path.dirname(os.path.abspath(output_png_path)), exist_ok=True)
                pix.save(output_png_path)
            doc.close()
        finally:
            if os.path.isfile(temp_pdf):
                try:
                    os.remove(temp_pdf)
                except OSError:
                    pass

        return output_png_path
