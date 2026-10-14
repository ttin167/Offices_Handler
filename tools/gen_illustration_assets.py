"""
gen_illustration_assets.py — Generates publication-grade illustration assets (unDraw / Storyset aesthetic)
for Antigravity Office Studio (Presentation & Document Hub).
"""

import os
import math
from PIL import Image, ImageDraw, ImageFont

ASSETS_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "diagram_assets", "illustrations")
os.makedirs(ASSETS_DIR, exist_ok=True)

WIDTH, HEIGHT = 1200, 800

def create_base_canvas(bg_color=(15, 23, 42)):
    im = Image.new("RGBA", (WIDTH, HEIGHT), bg_color)
    draw = ImageDraw.Draw(im)
    # Subtle radial glow in background
    for r in range(400, 0, -20):
        alpha = int(25 * (1.0 - r / 400.0))
        glow_color = (56, 189, 248, alpha)
        draw.ellipse([WIDTH // 2 - r, HEIGHT // 2 - r, WIDTH // 2 + r, HEIGHT // 2 + r], fill=glow_color)
    return im, draw

def save_illustration(im, filename):
    out_path = os.path.join(ASSETS_DIR, filename)
    im.convert("RGB").save(out_path, "PNG", dpi=(300, 300), optimize=True)
    print(f"[OK] Generated illustration: {out_path} ({os.path.getsize(out_path)} bytes)")

def gen_cloud_architecture():
    im, draw = create_base_canvas((11, 15, 25))
    # Grid dots
    for x in range(40, WIDTH, 60):
        for y in range(40, HEIGHT, 60):
            draw.point((x, y), fill=(255, 255, 255, 25))

    # Center Cloud Shape
    cx, cy = WIDTH // 2, HEIGHT // 2 - 20
    # Outer Glow
    draw.ellipse([cx - 260, cy - 140, cx + 260, cy + 180], fill=(56, 189, 248, 30))
    # Cloud bubbles
    draw.ellipse([cx - 160, cy - 60, cx + 20, cy + 100], fill=(30, 41, 59, 230), outline=(56, 189, 248), width=3)
    draw.ellipse([cx - 60, cy - 130, cx + 160, cy + 80], fill=(30, 41, 59, 230), outline=(56, 189, 248), width=3)
    draw.ellipse([cx + 40, cy - 60, cx + 190, cy + 100], fill=(30, 41, 59, 230), outline=(56, 189, 248), width=3)
    draw.rounded_rectangle([cx - 150, cy + 10, cx + 180, cy + 100], radius=30, fill=(30, 41, 59, 255), outline=(56, 189, 248), width=3)

    # Server Racks inside / below
    racks = [(-280, 180), (0, 180), (280, 180)]
    for rx, ry in racks:
        x1, y1 = cx + rx - 90, cy + ry - 40
        x2, y2 = cx + rx + 90, cy + ry + 40
        # Connecting dashed lines to cloud
        draw.line([(cx + rx, cy + ry - 40), (cx + rx // 2, cy + 100)], fill=(56, 189, 248, 180), width=2)
        # Server box
        draw.rounded_rectangle([x1, y1, x2, y2], radius=12, fill=(15, 23, 42), outline=(99, 102, 241), width=2)
        # Server LEDs
        draw.ellipse([x1 + 15, y1 + 18, x1 + 25, y1 + 28], fill=(16, 185, 129))
        draw.ellipse([x1 + 35, y1 + 18, x1 + 45, y1 + 28], fill=(56, 189, 248))
        draw.line([x1 + 60, y1 + 23, x2 - 20, y1 + 23], fill=(100, 116, 139), width=3)
        draw.line([x1 + 60, y1 + 33, x2 - 50, y1 + 33], fill=(71, 85, 105), width=2)

    save_illustration(im, "cloud_architecture.png")

def gen_data_analytics():
    im, draw = create_base_canvas((15, 23, 42))
    # Main Dashboard Card
    cx, cy = WIDTH // 2, HEIGHT // 2
    bx1, by1 = cx - 380, cy - 240
    bx2, by2 = cx + 380, cy + 240
    draw.rounded_rectangle([bx1, by1, bx2, by2], radius=24, fill=(30, 41, 59), outline=(71, 85, 105), width=2)
    # Header bar
    draw.rounded_rectangle([bx1, by1, bx2, by1 + 50], radius=24, fill=(15, 23, 42))
    draw.ellipse([bx1 + 25, by1 + 18, bx1 + 39, by1 + 32], fill=(239, 68, 68))
    draw.ellipse([bx1 + 47, by1 + 18, bx1 + 61, by1 + 32], fill=(245, 158, 11))
    draw.ellipse([bx1 + 69, by1 + 18, bx1 + 83, by1 + 32], fill=(16, 185, 129))

    # Metric Cards Top
    for i in range(3):
        mx1 = bx1 + 30 + i * 240
        my1 = by1 + 70
        mx2 = mx1 + 220
        my2 = my1 + 90
        draw.rounded_rectangle([mx1, my1, mx2, my2], radius=12, fill=(15, 23, 42), outline=(56, 189, 248, 80), width=1)
        draw.line([mx1 + 20, my1 + 25, mx1 + 90, my1 + 25], fill=(148, 163, 184), width=4)
        draw.line([mx1 + 20, my1 + 55, mx1 + 140, my1 + 55], fill=(56, 189, 248) if i==0 else (16, 185, 129), width=8)

    # Bar Chart Area
    ch_x1 = bx1 + 30
    ch_y1 = by1 + 185
    ch_x2 = bx1 + 420
    ch_y2 = by2 - 30
    draw.rounded_rectangle([ch_x1, ch_y1, ch_x2, ch_y2], radius=16, fill=(15, 23, 42))
    bars = [80, 140, 110, 190, 160, 230]
    for idx, bh in enumerate(bars):
        bx = ch_x1 + 35 + idx * 60
        by = ch_y2 - 25
        draw.rounded_rectangle([bx, by - bh, bx + 36, by], radius=6, fill=(99, 102, 241) if idx % 2 == 0 else (56, 189, 248))

    # Line Chart Area
    ln_x1 = bx1 + 470
    ln_y1 = by1 + 185
    ln_x2 = bx2 - 30
    ln_y2 = by2 - 30
    draw.rounded_rectangle([ln_x1, ln_y1, ln_x2, ln_y2], radius=16, fill=(15, 23, 42))
    pts = [
        (ln_x1 + 30, ln_y2 - 50),
        (ln_x1 + 80, ln_y2 - 90),
        (ln_x1 + 130, ln_y2 - 60),
        (ln_x1 + 180, ln_y2 - 160),
        (ln_x1 + 225, ln_y2 - 130),
        (ln_x2 - 30, ln_y2 - 210),
    ]
    draw.line(pts, fill=(16, 185, 129), width=4)
    for px, py in pts:
        draw.ellipse([px - 6, py - 6, px + 6, py + 6], fill=(255, 255, 255), outline=(16, 185, 129), width=3)

    save_illustration(im, "data_analytics.png")

def gen_security_shield():
    im, draw = create_base_canvas((8, 12, 20))
    cx, cy = WIDTH // 2, HEIGHT // 2

    # Shield Outer Boundary
    shield_pts = [
        (cx, cy - 230),
        (cx + 190, cy - 140),
        (cx + 170, cy + 90),
        (cx, cy + 240),
        (cx - 170, cy + 90),
        (cx - 190, cy - 140),
    ]
    # Shield glow
    draw.polygon(shield_pts, fill=(16, 185, 129, 35))
    draw.polygon(shield_pts, outline=(16, 185, 129), width=4)

    # Inner Shield Card
    inner_pts = [
        (cx, cy - 200),
        (cx + 160, cy - 120),
        (cx + 140, cy + 70),
        (cx, cy + 200),
        (cx - 140, cy + 70),
        (cx - 160, cy - 120),
    ]
    draw.polygon(inner_pts, fill=(15, 23, 42), outline=(56, 189, 248), width=3)

    # Padlock / Checkmark in Center
    draw.ellipse([cx - 45, cy - 70, cx + 45, cy + 20], fill=(16, 185, 129, 40))
    # Padlock arch
    draw.arc([cx - 38, cy - 75, cx + 38, cy + 5], 180, 0, fill=(56, 189, 248), width=10)
    # Padlock body
    draw.rounded_rectangle([cx - 50, cy - 20, cx + 50, cy + 65], radius=14, fill=(16, 185, 129), outline=(255, 255, 255), width=2)
    # Keyhole
    draw.ellipse([cx - 10, cy + 5, cx + 10, cy + 25], fill=(15, 23, 42))
    draw.polygon([(cx - 7, cy + 20), (cx + 7, cy + 20), (cx + 10, cy + 45), (cx - 10, cy + 45)], fill=(15, 23, 42))

    save_illustration(im, "security_shield.png")

def gen_team_collaboration():
    im, draw = create_base_canvas((15, 23, 42))
    cx, cy = WIDTH // 2, HEIGHT // 2

    # 3 User Nodes with connecting sync ring
    radius = 190
    draw.ellipse([cx - radius, cy - radius, cx + radius, cy + radius], outline=(56, 189, 248, 70), width=3)

    positions = [
        (cx, cy - radius, (56, 189, 248)),
        (cx - int(radius * 0.866), cy + int(radius * 0.5), (99, 102, 241)),
        (cx + int(radius * 0.866), cy + int(radius * 0.5), (16, 185, 129)),
    ]

    for px, py, col in positions:
        # Avatar card
        draw.ellipse([px - 65, py - 65, px + 65, py + 65], fill=(30, 41, 59), outline=col, width=3)
        # Head
        draw.ellipse([px - 22, py - 35, px + 22, py + 9], fill=col)
        # Body
        draw.pieslice([px - 42, py + 12, px + 42, py + 65], 180, 0, fill=col)

    # Central Collaboration Hub Document
    dx1, dy1 = cx - 75, cy - 90
    dx2, dy2 = cx + 75, cy + 90
    draw.rounded_rectangle([dx1, dy1, dx2, dy2], radius=14, fill=(15, 23, 42), outline=(255, 255, 255), width=2)
    # Document lines
    for line_y in range(dy1 + 25, dy2 - 20, 18):
        draw.line([dx1 + 20, line_y, dx2 - 20, line_y], fill=(56, 189, 248), width=3)

    save_illustration(im, "team_collaboration.png")

def gen_workflow_automation():
    im, draw = create_base_canvas((11, 15, 25))
    cx, cy = WIDTH // 2, HEIGHT // 2

    # Pipeline Steps (DOCX -> AST -> PPTX -> PDF)
    steps = [
        (-330, "DOCX", (59, 130, 246)),
        (-110, "PARSER", (99, 102, 241)),
        (110, "ENGINE", (16, 185, 129)),
        (330, "16:9 PPTX", (245, 158, 11)),
    ]

    for i, (sx, label, col) in enumerate(steps):
        bx1 = cx + sx - 85
        by1 = cy - 60
        bx2 = cx + sx + 85
        by2 = cy + 60

        draw.rounded_rectangle([bx1, by1, bx2, by2], radius=16, fill=(30, 41, 59), outline=col, width=3)
        # Inner bar
        draw.line([bx1 + 25, by1 + 25, bx2 - 25, by1 + 25], fill=col, width=6)
        draw.line([bx1 + 25, by1 + 45, bx2 - 45, by1 + 45], fill=(148, 163, 184), width=4)

        # Arrow to next step
        if i < len(steps) - 1:
            next_x = cx + steps[i + 1][0] - 85
            mid_x = (bx2 + next_x) // 2
            draw.line([bx2 + 10, cy, next_x - 10, cy], fill=(56, 189, 248), width=3)
            draw.polygon([(next_x - 10, cy), (next_x - 22, cy - 7), (next_x - 22, cy + 7)], fill=(56, 189, 248))

    save_illustration(im, "workflow_automation.png")

if __name__ == "__main__":
    gen_cloud_architecture()
    gen_data_analytics()
    gen_security_shield()
    gen_team_collaboration()
    gen_workflow_automation()
    print("[SUCCESS] All 5 illustration assets generated.")
