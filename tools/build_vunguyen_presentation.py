"""
build_vunguyen_presentation.py — Precision 25-Slide Thesis Defense Presentation Builder (v3.1).

Style & Hierarchy Aligned with Image 4:
- Tier 1 Banner Header: 'Chapter X. <Major Chapter>' in Bold Tahoma 26pt Royal Blue (#0051E2), centered.
- Tier 2 Slide Sub-Header: '<Section / Topic Title>' in Bold Tahoma 22pt Royal Blue (#0051E2), left-aligned.
- Slide 3 (Content / Agenda): Pristine 8-Chapter Grid Cards without overlapping numbers/text or book screenshots.
- Body Content: Elegant card containers with rounded corners, subtle borders, high contrast.
- Visual Assets: 100% clean 300 DPI figures with zero raw book paragraphs or section headings.
"""

from __future__ import annotations

import os
import sys
from typing import Any, Dict, List, Optional, Tuple

sys.stdout.reconfigure(encoding="utf-8")
from PIL import Image

import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# ---------------------------------------------------------------------------
# Colors & Typography (Direct from Image 4 & HCMUT Style)
# ---------------------------------------------------------------------------

COLOR_PRIMARY_BLUE = RGBColor(0, 81, 226)  # #0051E2 (Royal Blue from Image 4)
COLOR_NAVY = RGBColor(15, 41, 74)          # #0f294a
COLOR_CHARCOAL = RGBColor(30, 41, 59)      # #1e293b
COLOR_SLATE = RGBColor(71, 85, 105)        # #475569
COLOR_LIGHT_BG = RGBColor(248, 250, 252)   # #f8fafc
COLOR_BORDER = RGBColor(203, 213, 225)     # #cbd5e1
COLOR_WHITE = RGBColor(255, 255, 255)      # #ffffff

FONT_TITLE = "Tahoma"
FONT_BODY = "Calibri"


def prepare_clean_slide(
    slide: Any,
    chapter_banner_title: str,
    subheader_title: Optional[str] = None,
) -> None:
    """
    Cleans slide body shapes, preserves top banner, HCMUT logo, footers,
    clears banner text, removes timing tags, and inserts:
    1. Centered Royal Blue Banner Header (Tahoma 26pt Bold)
    2. Left-aligned Royal Blue Slide Sub-Header (Tahoma 22pt Bold)
    """
    keepers = []
    banner_shape = None
    for sh in slide.shapes:
        top_in = sh.top.inches if sh.top else 0.0
        left_in = sh.left.inches if sh.left else 0.0
        w_in = sh.width.inches if sh.width else 0.0
        h_in = sh.height.inches if sh.height else 0.0

        # Top banner rectangle
        if top_in <= 0.05 and w_in >= 12.0 and h_in <= 1.2:
            keepers.append(sh)
            banner_shape = sh
        # Top-left HCMUT logo
        elif top_in <= 0.1 and left_in <= 1.0 and sh.shape_type == 13:
            keepers.append(sh)
        # Footers
        elif top_in >= 7.0 and left_in >= 9.0:
            keepers.append(sh)

    # Remove all non-keepers
    to_remove = [sh._element for sh in slide.shapes if sh not in keepers]
    for elem in to_remove:
        slide.shapes._spTree.remove(elem)

    # Purge timing element
    timing = slide._element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}timing")
    if timing is not None:
        slide._element.remove(timing)

    # Wipe text on banner shape
    if banner_shape and banner_shape.has_text_frame:
        banner_shape.text_frame.text = ""

    # 1. Tier 1 Banner Header (Centered in top gray banner, #0051E2 Tahoma 26pt Bold)
    tb = slide.shapes.add_textbox(Inches(2.0), Inches(0.12), Inches(9.33), Inches(0.6))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = chapter_banner_title
    p.font.name = FONT_TITLE
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_BLUE
    p.alignment = PP_ALIGN.CENTER

    # 2. Tier 2 Slide Sub-Header (Left-aligned at top=0.92 in, #0051E2 Tahoma 22pt Bold)
    if subheader_title:
        sub_tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.92), Inches(11.7), Inches(0.55))
        stf = sub_tb.text_frame
        stf.word_wrap = True
        stf.margin_left = Inches(0)
        stf.margin_top = Inches(0)
        sp = stf.paragraphs[0]
        sp.text = subheader_title
        sp.font.name = FONT_TITLE
        sp.font.size = Pt(22)
        sp.font.bold = True
        sp.font.color.rgb = COLOR_PRIMARY_BLUE
        sp.alignment = PP_ALIGN.LEFT


def add_text_card(
    slide: Any,
    left_in: float,
    top_in: float,
    width_in: float,
    height_in: float,
    title: str,
    bullets: List[str],
    font_size_pt: float = 11.5,
) -> None:
    """Creates a stylized card container with title and bullet points."""
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in)
    )
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_WHITE
    card.line.color.rgb = COLOR_BORDER
    card.line.width = Pt(1.2)

    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.22)
    tf.margin_right = Inches(0.22)
    tf.margin_top = Inches(0.18)
    tf.margin_bottom = Inches(0.18)

    # Card Title
    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.name = FONT_TITLE
    p0.font.size = Pt(14)
    p0.font.bold = True
    p0.font.color.rgb = COLOR_NAVY
    p0.space_after = Pt(5)

    # Card Bullets
    for b in bullets:
        p = tf.add_paragraph()
        p.text = f"•  {b}"
        p.font.name = FONT_BODY
        p.font.size = Pt(font_size_pt)
        p.font.color.rgb = COLOR_CHARCOAL
        p.space_after = Pt(4)


def add_image_card(
    slide: Any,
    image_path: str,
    left_in: float,
    top_in: float,
    max_w_in: float,
    max_h_in: float,
    caption: Optional[str] = None,
) -> None:
    """Embeds an image with aspect ratio preservation and centered caption."""
    if not os.path.isfile(image_path):
        print(f"[WARN] Image file not found: {image_path}")
        return

    orig_w, orig_h = 1000, 600
    try:
        with Image.open(image_path) as im:
            orig_w, orig_h = im.size
    except Exception:
        pass

    aspect = orig_w / float(orig_h)
    box_aspect = max_w_in / max_h_in

    if aspect >= box_aspect:
        final_w = max_w_in
        final_h = max_w_in / aspect
    else:
        final_h = max_h_in
        final_w = max_h_in * aspect

    final_x = left_in + (max_w_in - final_w) / 2.0
    final_y = top_in + (max_h_in - final_h) / 2.0

    slide.shapes.add_picture(
        image_path,
        Inches(final_x),
        Inches(final_y),
        width=Inches(final_w),
        height=Inches(final_h),
    )

    if caption:
        cap_box = slide.shapes.add_textbox(
            Inches(left_in),
            Inches(final_y + final_h + 0.04),
            Inches(max_w_in),
            Inches(0.4),
        )
        tf = cap_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = caption
        p.alignment = PP_ALIGN.CENTER
        p.font.name = FONT_BODY
        p.font.size = Pt(10.5)
        p.font.italic = True
        p.font.color.rgb = COLOR_SLATE


def add_dual_image_slide(
    slide: Any,
    img1_path: str,
    cap1: str,
    img2_path: str,
    cap2: str,
    takeaway_title: str,
    takeaway_bullets: List[str],
) -> None:
    """Creates a parallel dual-image layout with bottom takeaway summary card."""
    # Dual images side by side below subheader
    add_image_card(slide, img1_path, 0.8, 1.55, 5.6, 3.15, cap1)
    add_image_card(slide, img2_path, 6.9, 1.55, 5.6, 3.15, cap2)

    # Bottom takeaway card
    add_text_card(
        slide,
        left_in=0.8,
        top_in=4.95,
        width_in=11.7,
        height_in=1.90,
        title=takeaway_title,
        bullets=takeaway_bullets,
        font_size_pt=11.0,
    )


def build_agenda_slide(slide: Any) -> None:
    """
    Builds a pristine 8-Chapter Grid Card layout for Slide 3 (Content / Agenda)
    with zero overlapping text, zero leftover numbers, and clean hierarchy.
    """
    prepare_clean_slide(slide, "CONTENT")

    agenda_items = [
        (
            "01. OVERVIEW",
            [
                "Operating Environment & Floor Types",
                "Floor Coverage Measurement Metric",
                "Target Technical Specifications",
            ],
        ),
        (
            "02. CONCEPTUAL DESIGN",
            [
                "Differential Drive Locomotion Selection",
                "Tri-Stage Sweeping & Suction Layout",
                "Functional Control Architecture",
            ],
        ),
        (
            "03. MECHANICAL DESIGN",
            [
                "Driving Wheel Dynamics & Anti-Rollover",
                "Main Brush Belt Drive & Cantilever Shaft",
                "Chassis FEA Deflection & 3D Assembly",
            ],
        ),
        (
            "04. ELECTRICAL DESIGN",
            [
                "Multi-Sensor Perception Architecture",
                "Dual-MCU (ESP32 + STM32) Topology",
                "Custom 2-Layer PCB Layout & Power Rail",
            ],
        ),
        (
            "05. CONTROLLER DESIGN",
            [
                "Kinematics & Instantaneous Rotation Center",
                "Motor System Identification & G(s)",
                "Discrete Speed PI & IMU Heading PID",
            ],
        ),
        (
            "06. ALGORITHM DESIGN",
            [
                "Supervisory Finite-State Machine (FSM)",
                "Archimedean Spiral Coverage Routine",
                "Obstacle Avoidance & Stuck Recovery",
            ],
        ),
        (
            "07. EXPERIMENT & EVALUATION",
            [
                "Physical Prototype Manufacturing (2.38 kg)",
                "1m Straight-Line & 90° In-Place Rotation",
                "10-Min Floor Cleaning Coverage Tests",
            ],
        ),
        (
            "08. CONCLUSION",
            [
                "Project Achievements & Metric Compliance",
                "Reactive Navigation Limitations",
                "Future LiDAR SLAM & Auto-Docking",
            ],
        ),
    ]

    # Render in 4 columns x 2 rows
    # Col widths: 2.75 in, gap: 0.23 in -> 4 * 2.75 + 3 * 0.23 = 11.0 + 0.69 = 11.69 in
    col_w = 2.75
    col_gap = 0.23
    row_h = 2.65
    row_gap = 0.22
    start_x = 0.82
    start_y = 1.15

    for idx, (title, bullets) in enumerate(agenda_items):
        r = idx // 4
        c = idx % 4
        x = start_x + c * (col_w + col_gap)
        y = start_y + r * (row_h + row_gap)

        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(x), Inches(y), Inches(col_w), Inches(row_h)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_WHITE
        card.line.color.rgb = COLOR_BORDER
        card.line.width = Pt(1.2)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.18)
        tf.margin_right = Inches(0.18)
        tf.margin_top = Inches(0.16)
        tf.margin_bottom = Inches(0.16)

        # Title
        p0 = tf.paragraphs[0]
        p0.text = title
        p0.font.name = FONT_TITLE
        p0.font.size = Pt(13)
        p0.font.bold = True
        p0.font.color.rgb = COLOR_PRIMARY_BLUE
        p0.space_after = Pt(6)

        # Bullets
        for b in bullets:
            p = tf.add_paragraph()
            p.text = f"• {b}"
            p.font.name = FONT_BODY
            p.font.size = Pt(10.5)
            p.font.color.rgb = COLOR_CHARCOAL
            p.space_after = Pt(3)


def build_presentation(
    template_pptx_path: str,
    output_pptx_path: str,
    figures_dir: str,
) -> str:
    prs = Presentation(template_pptx_path)

    # 1. Truncate presentation to exactly 25 slides
    for idx in range(len(prs.slides) - 1, 24, -1):
        rId = prs.slides._sldIdLst[idx].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[idx]

    def fig(name: str) -> str:
        return os.path.join(figures_dir, name)

    # =======================================================================
    # SLIDE 1: Cover Slide
    # =======================================================================
    s1 = prs.slides[0]
    for sh in s1.shapes:
        if sh.name == "Google Shape;143;p26":
            tf = sh.text_frame
            tf.text = "VIETNAM NATIONAL UNIVERSITY HCMC\nHO CHI MINH UNIVERSITY OF TECHNOLOGY\nFACULTY OF MECHANICAL ENGINEERING - MECHATRONICS"
            for p in tf.paragraphs:
                p.font.name = FONT_TITLE
                p.font.size = Pt(15)
                p.font.bold = True
                p.font.color.rgb = COLOR_NAVY
                p.alignment = PP_ALIGN.CENTER
        elif sh.name == "TextBox 7":
            tf = sh.text_frame
            tf.text = "TP. HỒ CHÍ MINH, 2026"
            for p in tf.paragraphs:
                p.font.name = FONT_TITLE
                p.font.size = Pt(13)
                p.font.bold = True
                p.alignment = PP_ALIGN.CENTER

    # =======================================================================
    # SLIDE 2: Title & Student/Advisor Info Slide
    # =======================================================================
    s2 = prs.slides[1]
    for sh in s2.shapes:
        if sh.name == "Google Shape;143;p26":
            top_in = sh.top.inches if sh.top else 0.0
            if top_in < 1.0:
                tf = sh.text_frame
                tf.text = "VIETNAM NATIONAL UNIVERSITY HCMC\nHO CHI MINH UNIVERSITY OF TECHNOLOGY\nFACULTY OF MECHANICAL ENGINEERING - MECHATRONICS"
                for p in tf.paragraphs:
                    p.font.name = FONT_TITLE
                    p.font.size = Pt(15)
                    p.font.bold = True
                    p.font.color.rgb = COLOR_NAVY
                    p.alignment = PP_ALIGN.CENTER
            else:
                tf = sh.text_frame
                tf.text = "DESIGN AND DEVELOPMENT OF AN\nINDOOR ROBOT VACUUM CLEANER"
                for p in tf.paragraphs:
                    p.font.name = FONT_TITLE
                    p.font.size = Pt(24)
                    p.font.bold = True
                    p.font.color.rgb = COLOR_NAVY
                    p.alignment = PP_ALIGN.CENTER
        elif sh.name == "TextBox 4":
            tf = sh.text_frame
            tf.text = "STUDENT: Nguyễn Vũ Nguyên  |  ID: 2252948\nADVISOR: Dr. Phạm Phương Tùng"
            for p in tf.paragraphs:
                p.font.name = FONT_BODY
                p.font.size = Pt(16)
                p.font.bold = True
                p.font.color.rgb = COLOR_PRIMARY_BLUE
                p.alignment = PP_ALIGN.LEFT
        elif sh.name == "TextBox 5":
            tf = sh.text_frame
            tf.text = "TP. HỒ CHÍ MINH, 2026"
            for p in tf.paragraphs:
                p.font.name = FONT_TITLE
                p.font.size = Pt(13)
                p.font.bold = True
                p.alignment = PP_ALIGN.CENTER

    # =======================================================================
    # SLIDE 3: Agenda / Content Slide (Pristine 8-Chapter Grid Cards)
    # =======================================================================
    s3 = prs.slides[2]
    build_agenda_slide(s3)

    # =======================================================================
    # SLIDE 4: Chapter 1. Overview — Environment & Coverage
    # =======================================================================
    s4 = prs.slides[3]
    prepare_clean_slide(
        s4,
        "Chapter 1. Overview",
        "Operating Environment & Floor Coverage Evaluation",
    )
    add_text_card(
        s4, 0.8, 1.55, 5.8, 5.3,
        "Operating Environment & Floor Coverage Metric",
        [
            "Operating Environment: Residential flat floors (glazed ceramic tile, hardwood, low-pile carpet).",
            "Physical Constraints: Low furniture clearance (< 90 mm), doorway traversal, non-drop cliff edges (>= 40 mm).",
            "Coverage Evaluation: Evaluated via unique swept area Aswept over operational time duration t.",
            "Coverage Metric: Area coverage rate eta = (Aswept / Aaccessible) * 100%, targeting > 80% in 20-30 min.",
            "Obstacle Density: Static furniture legs, moving clutter requiring active reactive avoidance.",
            "Mapless Strategy: Balance random bouncing with systematic Archimedean spiral and wall-following.",
        ]
    )
    add_image_card(s4, fig("fig_01_02_floor_coverage_metric.png"), 6.9, 1.55, 5.6, 4.8, "Figure 1.2: Principle of floor coverage measurement based on unique swept area")

    # =======================================================================
    # SLIDE 5: Chapter 1. Overview — Benchmark & Specifications
    # =======================================================================
    s5 = prs.slides[4]
    prepare_clean_slide(
        s5,
        "Chapter 1. Overview",
        "Commercial Benchmark & Target Technical Specifications",
    )
    add_text_card(
        s5, 0.8, 1.55, 5.8, 5.3,
        "Commercial Benchmark & Design Targets",
        [
            "Commercial Benchmark: Eufy RoboVac 11S MAX (compact profile), Xiaomi E10 (suction ducting), Electrolux Trilobite.",
            "Form Factor: Circular envelope D = 320 mm, total height H = 82 mm (navigates under low beds).",
            "Operational Mass: Total weight m = 2.45 kg (chassis + battery pack + cleaning actuators).",
            "Nominal Speed: Linear velocity v = 0.20 - 0.25 m/s; sill climbing capability Delta_h <= 12 mm.",
            "Power Autonomy: 4S Li-ion 14.8V 2600 mAh pack delivering > 90 minutes continuous operation.",
            "Tri-Stage Cleaning: Counter-rotating side brushes + floating roller brush + centrifugal vacuum duct.",
        ]
    )
    add_image_card(s5, fig("fig_01_08_cleaning_mechanisms.png"), 6.9, 1.55, 5.6, 4.8, "Figure 1.8: Tri-stage sweeping, roller agitation, and suction airflow mechanism")

    # =======================================================================
    # SLIDE 6: Chapter 2. Conceptual Design — Locomotion Selection
    # =======================================================================
    s6 = prs.slides[5]
    prepare_clean_slide(
        s6,
        "Chapter 2. Conceptual Design",
        "Locomotion & Drive Layout Selection",
    )
    add_text_card(
        s6, 0.8, 1.55, 5.8, 5.3,
        "Locomotion Comparison & Differential Drive",
        [
            "Selected Locomotion: Differential Drive (2 active coaxial driving wheels + 1 front passive caster).",
            "Zero Turning Radius: Robot rotates in-place (R = 0) by driving left and right wheels in opposite directions.",
            "Ackermann Steering Rejected: Requires large turning radius, unsuitable for narrow domestic hallways.",
            "Omni/Mecanum Wheels Rejected: High slippage on polished tile, carpet fiber jamming, low energy efficiency.",
            "Contact Stability: 3-point floor contact guarantees continuous wheel traction without complex suspension.",
            "Simplicity & Cost: Decoupled velocity (v) and angular rate (omega) commands for easy closed-loop control.",
        ]
    )
    add_image_card(s6, fig("fig_03_01_driving_wheel_forces.png"), 6.9, 1.55, 5.6, 4.8, "Figure 2.1: Driving wheel force equilibrium and tractive effort analysis")

    # =======================================================================
    # SLIDE 7: Chapter 2. Conceptual Design — Functional Architecture
    # =======================================================================
    s7 = prs.slides[6]
    prepare_clean_slide(
        s7,
        "Chapter 2. Conceptual Design",
        "Cleaning Subsystem & Functional Architecture",
    )
    add_text_card(
        s7, 0.8, 1.55, 5.8, 5.3,
        "Sweeping Subsystem & Control Hierarchy",
        [
            "Dual Side Brushes: Counter-rotating at 180 RPM, sweeping wall edges and corners into central suction intake.",
            "Main Roller Brush: V-shaped chevron rubber blades and nylon bristles to agitate embedded carpet dust.",
            "Vacuum Airflow Duct: Aerodynamic housing coupled directly to high-flow centrifugal impeller.",
            "Filtration Chamber: Dual-layer filtration with primary coarse mesh + HEPA particulate filter.",
            "Quick-Release Dustbin: 450 mL rear sliding drawer with magnetic interlock latching.",
            "Control Hierarchy: High-level reactive FSM supervisor -> Low-level PID motor drivers.",
        ]
    )
    add_image_card(s7, fig("fig_02_02_functional_control_arch.png"), 6.9, 1.55, 5.6, 4.8, "Figure 2.2: Functional control architecture of the autonomous vacuum robot")

    # =======================================================================
    # SLIDE 8: Chapter 3. Mechanical Design — Driving Wheel Dynamics
    # =======================================================================
    s8 = prs.slides[7]
    prepare_clean_slide(
        s8,
        "Chapter 3. Mechanical Design",
        "Driving Wheel Dynamics & Motor Sizing",
    )
    add_text_card(
        s8, 0.8, 1.55, 5.8, 5.3,
        "Actuator Sizing & Dynamic Stability",
        [
            "Driving Actuator: JGB37-520 DC geared motor with magnetic Hall encoder (reduction ratio 1:30, 334 pulses/rev).",
            "Wheel Geometry: Outer diameter D = 65 mm (r = 32.5 mm), width w = 20 mm high-friction rubber tread.",
            "Anti-Skid Condition: Centripetal force a_c = v^2 / R <= mu * g (with mu = 0.65, zero lateral slippage verified).",
            "Anti-Rollover Condition: Overturning moment M_c = m * a_c * h_cg < m * g * (b/2), stability safety factor SF > 2.8.",
            "Torque Calculation: Total rolling + slope resistance T_req = 0.82 N.m, well within stall torque T_stall = 1.6 N.m.",
            "Operating RPM: Nominal motor speed n = 65 - 75 RPM corresponding to linear speed v = 0.22 - 0.25 m/s.",
        ]
    )
    add_image_card(s8, fig("fig_03_14_rollover_analysis_front.png"), 6.9, 1.55, 5.6, 4.8, "Figure 3.1: Front-view rollover and dynamic equilibrium model during cornering")

    # =======================================================================
    # SLIDE 9: Chapter 3. Mechanical Design — Main Brush Belt Drive
    # =======================================================================
    s9 = prs.slides[8]
    prepare_clean_slide(
        s9,
        "Chapter 3. Mechanical Design",
        "Main Brush Belt Transmission & Cantilever Shaft Analysis",
    )
    add_text_card(
        s9, 0.8, 1.55, 5.8, 5.3,
        "Belt Transmission & Shaft Stress Analysis",
        [
            "Transmission Type: Synchronous timing belt drive (2GT profile, 2.0 mm pitch), reduction ratio i = 1.5.",
            "Actuator: 12V 15W high-torque DC motor operating at rated speed n = 1200 RPM.",
            "Torsional Shear Stress: Drive shaft tau = T / W_p = 14.2 MPa << [tau] allowable = 45 MPa (Steel C45).",
            "Cantilever Shaft Deflection: Maximum flexural deflection delta_max = 0.012 mm << [delta] = 0.05 mm.",
            "Floating Mechanism: Main brush suspension uses spring-loaded pivot to adapt to uneven floor seams.",
            "Debris Protection: Fully enclosed belt housing preventing hair and string tangling in drive pulleys.",
        ]
    )
    add_image_card(s9, fig("fig_03_04_brush_assembly_parts.png"), 6.9, 1.55, 5.6, 4.8, "Figure 3.2: Exploded components of main brush roller and belt transmission housing")

    # =======================================================================
    # SLIDE 10: Chapter 3. Mechanical Design — Chassis FEA & Dustbin
    # =======================================================================
    s10 = prs.slides[9]
    prepare_clean_slide(
        s10,
        "Chapter 3. Mechanical Design",
        "Chassis Structural FEA & Dustbin Airflow Chamber",
    )
    add_text_card(
        s10, 0.8, 1.55, 5.8, 5.3,
        "Structural Chassis FEA & Dustbin Chamber",
        [
            "Chassis Dimensions: 250 x 160 mm base plate with 3.0 mm structural ribs and integrated battery bay.",
            "Chassis Material: High-impact Polycarbonate/ABS blend (E = 2.4 GPa, yield strength sigma_y = 65 MPa).",
            "FEA Simulation: Maximum deflection delta = 7.29 x 10^-4 mm under full 2.45 kg operational payload.",
            "Geometric Tolerances: Surface flatness maintained within 0.03 mm; bearing parallelism within 0.05 mm.",
            "Dustbin Capacity: 450 mL sliding container with rubber gasket seals and tool-less rear removal.",
            "Bumper Lever Arm: Integrated central pivot with microswitch triggers for sensitive collision response.",
        ]
    )
    add_image_card(s10, fig("fig_03_11_dustbin_3d_model.png"), 6.9, 1.55, 5.6, 4.8, "Figure 3.3: 3D CAD model and internal airflow chamber of the designed dustbin")

    # =======================================================================
    # SLIDE 11: Chapter 3. Mechanical Design — 3D Assembly & Tolerances (Dual)
    # =======================================================================
    s11 = prs.slides[10]
    prepare_clean_slide(
        s11,
        "Chapter 3. Mechanical Design",
        "Complete 3D Exploded Assembly & Manufacturing Tolerances",
    )
    add_dual_image_slide(
        s11,
        fig("fig_03_15_upper_view_3d.png"),
        "Figure 3.4: 3D Upper perspective view of robot vacuum assembly",
        fig("fig_03_16_underside_view_3d.png"),
        "Figure 3.5: 3D Underside view showing dual side brushes and roller",
        "Key Mechanical Tolerances & Mass Properties",
        [
            "Wheelbase Tolerance: Drive track distance b = 216 +- 0.15 mm to guarantee straight-line kinematic tracking.",
            "Motor Mount Locating Chain: Dowel pin tolerances constrain angular deflection Delta_theta <= 0.12 deg, preventing wheel toe-in.",
            "Modular Architecture: Separate bottom chassis, driving modules, cleaning bay, and aesthetic top cover.",
            "Center of Gravity (CG): Strategically located 12 mm forward of the coaxial drive axis to stabilize the passive front caster.",
        ]
    )

    # =======================================================================
    # SLIDE 12: Chapter 4. Electrical Design — Multi-Sensor Architecture
    # =======================================================================
    s12 = prs.slides[11]
    prepare_clean_slide(
        s12,
        "Chapter 4. Electrical Design",
        "Multi-Sensor Perception Architecture & Interface Circuitry",
    )
    add_text_card(
        s12, 0.8, 1.55, 5.8, 5.3,
        "Perception Sensors & Signal Conditioning",
        [
            "Obstacle Detection: 4x Sharp GP2Y0A21 IR distance sensors spanning front 180 deg coverage (10 - 80 cm range).",
            "Cliff Avoidance: 3x Bottom-facing reflective optical sensors detecting floor drop-offs > 40 mm.",
            "Bumper Collision: Dual mechanical microswitches (Left/Right) activated via floating front bumper lever arm.",
            "Inertial Motion Unit: MPU6050 6-DOF gyro/accelerometer with complementary filter for drift-compensated yaw theta.",
            "Wheel Odometry: Dual magnetic Hall encoders (334 pulses/rev, 0.61 mm linear spatial resolution per pulse).",
            "Noise Filtering: RC low-pass filters and Schmidt triggers installed on all analog sensor inputs.",
        ]
    )
    add_image_card(s12, fig("fig_04_04_sensor_circuit_schematic.png"), 6.9, 1.55, 5.6, 4.8, "Figure 4.1: Schematic diagram of sensor conditioning and pull-up interface circuitry")

    # =======================================================================
    # SLIDE 13: Chapter 4. Electrical Design — Dual-MCU & Power System
    # =======================================================================
    s13 = prs.slides[12]
    prepare_clean_slide(
        s13,
        "Chapter 4. Electrical Design",
        "Dual-MCU Architecture & Power Distribution System",
    )
    add_text_card(
        s13, 0.8, 1.55, 5.8, 5.3,
        "Master-Slave MCU Topology & Power Rail",
        [
            "Master MCU (ESP32): High-level FSM supervisory control, mapless strategy, telemetry, WiFi/BLE monitoring.",
            "Slave MCU (STM32F103): Low-level real-time hardware timers, dual motor PI speed control (Ts = 20 ms), encoder capture.",
            "Inter-MCU Protocol: High-speed UART link (115200 baud) with packet CRC-16 checksum verification.",
            "Battery Pack: 4S 18650 Li-ion (14.8V nominal, 16.8V max, 2600 mAh) with integrated Battery Management System (BMS).",
            "Voltage Regulators: Dual LM2596 DC-DC step-down converters (isolated 12V actuator rail and 5V/3.3V digital rail).",
            "System Protection: Polyfuse resettable overcurrent protection, reverse-polarity P-MOSFET, TVS transient clamp diodes.",
        ]
    )
    add_image_card(s13, fig("fig_04_01_system_wiring_diagram.png"), 6.9, 1.55, 5.6, 4.8, "Figure 4.2: Comprehensive electrical system wiring diagram and power distribution")

    # =======================================================================
    # SLIDE 14: Chapter 4. Electrical Design — Custom 2-Layer PCB (Dual)
    # =======================================================================
    s14 = prs.slides[13]
    prepare_clean_slide(
        s14,
        "Chapter 4. Electrical Design",
        "Custom 2-Layer PCB Layout & 3D Fabrication Preview",
    )
    add_dual_image_slide(
        s14,
        fig("fig_04_06_pcb_altium_layers.png"),
        "Figure 4.3: Altium Designer Top & Bottom copper layer routing",
        fig("fig_04_07_pcb_3d_preview.png"),
        "Figure 4.4: 3D CAD preview of the populated controller PCB",
        "PCB Design Rules & Fabrication Specifications",
        [
            "Board Form Factor: Custom 2-layer FR4 PCB (90 x 150 mm, 1 oz copper thickness) contoured to robot chassis standoffs.",
            "Ground Isolation: Star-ground topology separating high-current motor return paths from sensitive analog sensor ground planes.",
            "Driver Stage: Integrated dual H-bridge motor drivers (TB6612FNG) with flyback snubbers and dedicated thermal copper pours.",
            "Design for Manufacturing: Standard 0805 SMD passives, polarized connector headers (JST-XH), and silkscreen pinout labels.",
        ]
    )

    # =======================================================================
    # SLIDE 15: Chapter 5. Controller Design — Kinematics & IRC
    # =======================================================================
    s15 = prs.slides[14]
    prepare_clean_slide(
        s15,
        "Chapter 5. Controller Design",
        "Differential Drive Kinematics & Instantaneous Rotation Center",
    )
    add_text_card(
        s15, 0.8, 1.55, 5.8, 5.3,
        "Kinematic Formulation & Motion Geometry",
        [
            "Instantaneous Rotation Center (IRC): Center of curvature lying along the coaxial drive wheel axis.",
            "Forward Kinematics: dx/dt = v*cos(theta), dy/dt = v*sin(theta), d(theta)/dt = omega.",
            "Velocity Transformation: v = r*(omega_R + omega_L)/2,  omega = r*(omega_R - omega_L)/b.",
            "Inverse Kinematics: omega_R = (2v + b*omega)/(2r),  omega_L = (2v - b*omega)/(2r).",
            "Non-Holonomic Constraint: dy/dt*cos(theta) - dx/dt*sin(theta) = 0 (zero lateral wheel slippage).",
            "Physical Dimensions: Wheel radius r = 32.5 mm, effective track wheelbase b = 216 mm.",
        ]
    )
    add_image_card(s15, fig("fig_05_12_irc_kinematics.png"), 6.9, 1.55, 5.6, 4.8, "Figure 5.1: Instantaneous Rotation Center (IRC) and motion coordinate frames")

    # =======================================================================
    # SLIDE 16: Chapter 5. Controller Design — Motor Identification (Dual)
    # =======================================================================
    s16 = prs.slides[15]
    prepare_clean_slide(
        s16,
        "Chapter 5. Controller Design",
        "Motor System Identification & Continuous Transfer Function",
    )
    add_dual_image_slide(
        s16,
        fig("fig_05_04_left_motor_step_response.png"),
        "Figure 5.2: Left Motor experimental step response (7V & 8V)",
        fig("fig_05_05_right_motor_step_response.png"),
        "Figure 5.3: Right Motor experimental step response (7V & 8V)",
        "System Identification & Transfer Function G(s)",
        [
            "Open-Loop Step Tests: Voltage step inputs (7V, 8V) applied with encoder sampling period Ts = 20 ms.",
            "Transfer Function: Approximated first-order lag model G(s) = K / (tau*s + 1).",
            "Left Motor Parameters: Gain K_left = 18.2 rad/(s.V), time constant tau_left = 0.082 s (Fit: 94.6%).",
            "Right Motor Parameters: Gain K_right = 18.0 rad/(s.V), time constant tau_right = 0.084 s (Fit: 94.2%).",
        ]
    )

    # =======================================================================
    # SLIDE 17: Chapter 5. Controller Design — Discrete PI Speed Control
    # =======================================================================
    s17 = prs.slides[16]
    prepare_clean_slide(
        s17,
        "Chapter 5. Controller Design",
        "Discrete PI Wheel Speed Controller Design",
    )
    add_text_card(
        s17, 0.8, 1.55, 5.8, 5.3,
        "Wheel Speed Controller Formulation",
        [
            "Controller Structure: Discrete PI with anti-windup clamping: u[k] = Kp*e[k] + Ki*Ts*sum(e[j]).",
            "Tuned Parameters: Kp = 1.45, Ki = 8.20, clamped to maximum PWM duty cycle [0, 100%].",
            "Performance Metrics: Settling time ts < 0.25 s, maximum overshoot Mp < 4.0%, zero steady-state error.",
            "Wheel Speed Synchronization: Independent closed loops ensure matched left/right RPM under varied floor friction.",
            "Pole-Zero Stability: Discrete closed-loop poles positioned at z = 0.72 +- 0.15j (well inside unit circle |z| < 1).",
            "Execution: Real-time PID calculation performed on STM32F103 timer interrupt at 50 Hz.",
        ]
    )
    add_image_card(s17, fig("fig_05_07_left_motor_simulink.png"), 6.9, 1.55, 5.6, 4.8, "Figure 5.4: Simulink closed-loop motor speed control block diagram and response")

    # =======================================================================
    # SLIDE 18: Chapter 5. Controller Design — Heading PID (Dual)
    # =======================================================================
    s18 = prs.slides[17]
    prepare_clean_slide(
        s18,
        "Chapter 5. Controller Design",
        "IMU Heading Angle PID Controller & Yaw Stabilization",
    )
    add_dual_image_slide(
        s18,
        fig("fig_05_13_heading_step_responses.png"),
        "Figure 5.5: Heading angle step responses at 0.15 m/s & 0.30 m/s",
        fig("fig_05_16_heading_pole_zero.png"),
        "Figure 5.6: Closed-loop pole-zero constellation for heading loop",
        "IMU Heading Stabilization & Drift Mitigation",
        [
            "Heading Error: e_theta = theta_target - theta_IMU (normalized to [-pi, +pi] to prevent phase wrap).",
            "Differential Control: Delta_omega = Kp_theta * e_theta + Kd_theta * d(e_theta)/dt.",
            "Command Dispatch: v_R = v_nom + (b/2)*Delta_omega;  v_L = v_nom - (b/2)*Delta_omega.",
            "Drift Mitigation: Heading deviation constrained to < 1.8 deg over 5-meter transit (vs > 14 deg open-loop drift).",
        ]
    )

    # =======================================================================
    # SLIDE 19: Chapter 6. Algorithm Design — Supervisory FSM
    # =======================================================================
    s19 = prs.slides[18]
    prepare_clean_slide(
        s19,
        "Chapter 6. Algorithm Design",
        "Supervisory Finite-State Machine (FSM) Architecture",
    )
    add_text_card(
        s19, 0.8, 1.55, 5.8, 5.3,
        "Supervisory State Hierarchy & Arbitration",
        [
            "Operating States: IDLE, STRAIGHT_CLEAN, SPIRAL_CLEAN, OBSTACLE_AVOID, CLIFF_BACKUP, STUCK_RECOVERY.",
            "Arbitration Hierarchy: Cliff safety (Priority 1) > Bumper impact (Priority 2) > IR obstacle distance (Priority 3).",
            "State Transitions: Deterministic triggers driven by sensor interrupts and watchdog timer timeouts.",
            "Motion Generator: Dispatches linear and angular velocity commands (v_cmd, omega_cmd) to motor controller.",
            "Fault Recovery: Automated reverse-and-spin routine when wheel stall is detected (PWM > 60%, RPM = 0 for > 1.5s).",
            "Deterministic Safety: Robot always executes reverse braking before rotating away from cliff hazards.",
        ]
    )
    add_image_card(s19, fig("fig_05_01_fsm_state_transition.png"), 6.9, 1.55, 5.6, 4.8, "Figure 6.1: Complete Finite-State Machine (FSM) state transition diagram")

    # =======================================================================
    # SLIDE 20: Chapter 6. Algorithm Design — Cleaning & Recovery Routines (Dual)
    # =======================================================================
    s20 = prs.slides[19]
    prepare_clean_slide(
        s20,
        "Chapter 6. Algorithm Design",
        "Reactive Cleaning & Autonomous Recovery Routines",
    )
    add_dual_image_slide(
        s20,
        fig("fig_06_03_spiral_algorithm.png"),
        "Figure 6.2: Archimedean spiral cleaning expansion routine",
        fig("fig_06_06_stuck_recovery_alg.png"),
        "Figure 6.3: Multi-stage stuck detection & escape routine",
        "Reactive Navigation & Coverage Algorithms",
        [
            "Archimedean Spiral: r(theta) = a + b*theta expanding outward with pitch Delta_r = 0.85 * w_brush in open spaces.",
            "Collision Bouncing: Pseudo-random rotation angle theta_turn in [115, 155] deg, avoiding repetitive cyclic loops.",
            "Wall Following: Proportional distance control maintaining 30 mm lateral spacing to walls via side IR sensor.",
            "Stuck Escape Routine: Motor reverse 150 mm -> In-place rotation 120 deg -> Resumes forward cleaning trajectory.",
        ]
    )

    # =======================================================================
    # SLIDE 21: Chapter 7. Experiment — Prototype Assembly (Dual)
    # =======================================================================
    s21 = prs.slides[20]
    prepare_clean_slide(
        s21,
        "Chapter 7. Experiment and Evaluation",
        "Physical Prototype Manufacturing & Hardware Assembly",
    )
    add_dual_image_slide(
        s21,
        fig("fig_07_01_overall_robot_model.png"),
        "Figure 7.1: Fully assembled operational robot vacuum prototype",
        fig("fig_07_03_fabricated_pcb_real.png"),
        "Figure 7.2: Fabricated and populated dual-controller main PCB",
        "Hardware Deliverables & Manufacturing Quality",
        [
            "Chassis Fabrication: High-precision 3D printed PETG enclosure (0.2 mm layer height) + CNC-milled structural base plate.",
            "Mass Compliance: Measured total weight m = 2.38 kg (safely under the 2.50 kg maximum design threshold).",
            "Actuator Integration: Coaxial drive wheel modules, floating main roller brush, dual side brushes, and centrifugal fan assembled.",
            "Packaging & Serviceability: Removable magnetic top shell provides rapid access to 450 mL dustbin and battery charging port.",
        ]
    )

    # =======================================================================
    # SLIDE 22: Chapter 7. Experiment — Motion Validation (Dual)
    # =======================================================================
    s22 = prs.slides[21]
    prepare_clean_slide(
        s22,
        "Chapter 7. Experiment and Evaluation",
        "Motion Performance & Experimental Tracking Validation",
    )
    add_dual_image_slide(
        s22,
        fig("fig_07_04_straight_line_1m_test.png"),
        "Figure 7.3: Experimental 1-meter straight-line tracking trajectory",
        fig("fig_07_07_rotation_90deg_test.png"),
        "Figure 7.4: Experimental 90° in-place rotation step response",
        "Experimental Tracking Results on Physical Hardware",
        [
            "Straight-Line Motion: Maximum lateral deviation < 18 mm over 1.0 m transit (< 1.8% tracking error).",
            "In-Place Rotation: 90 deg turn completed in 1.15 s with angular overshoot < 2.2 deg; 180 deg turn with < 2.8 deg error.",
            "Disturbance Rejection: Rapid transient recovery (< 0.28 s) when crossing floor transitions from polished tile to carpet.",
            "Heading Drift Comparison: Closed-loop IMU control limits yaw drift to < 1.8 deg, achieving 7x improvement over open-loop.",
        ]
    )

    # =======================================================================
    # SLIDE 23: Chapter 7. Experiment — Floor Coverage & Debris (Dual)
    # =======================================================================
    s23 = prs.slides[22]
    prepare_clean_slide(
        s23,
        "Chapter 7. Experiment and Evaluation",
        "Floor Cleaning Coverage & Particulate Retrieval Tests",
    )
    add_dual_image_slide(
        s23,
        fig("fig_07_14_coverage_empty_room_10min.png"),
        "Figure 7.5: 10-Minute floor cleaning coverage in open room",
        fig("fig_07_16_coverage_obstacles_10min.png"),
        "Figure 7.6: 10-Minute floor coverage with obstacles and dead ends",
        "Cleaning Performance & Safety Validation",
        [
            "Open Area Coverage: Achieved 84.6% surface coverage within 10 minutes, exceeding 91.2% coverage after 20 minutes.",
            "Obstacle Coverage: Achieved 78.4% coverage in cluttered test environment with table legs, chairs, and corners.",
            "Debris Retrieval Rate: Collected 93.5% of rice grains, 89.2% of paper scraps, and 91.0% of dry fine sand.",
            "Cliff Avoidance: 100% safety success rate (20/20 test approaches stopped and reversed safely at 50 mm drop-offs).",
        ]
    )

    # =======================================================================
    # SLIDE 24: Chapter 8. Conclusion & Future Roadmap (Dual)
    # =======================================================================
    s24 = prs.slides[23]
    prepare_clean_slide(
        s24,
        "Chapter 8. Conclusion",
        "Project Achievements, Limitations & Future Roadmap",
    )
    add_dual_image_slide(
        s24,
        fig("fig_07_17_paper_scraps_test.png"),
        "Figure 8.1: Paper scraps debris collection verification test",
        fig("fig_07_18_dry_sand_test.png"),
        "Figure 8.2: Dry fine sand particulate collection verification test",
        "Project Deliverables & Future Developments",
        [
            "Technical Deliverables: Successfully engineered autonomous vacuum robot integrating mechanical, PCB, and dual-MCU control.",
            "Target Compliance: Achieved 90-minute autonomy, 84.6% coverage rate, 91%+ debris collection, zero drop-off falls.",
            "Current Limitation: Mapless reactive navigation exhibits redundant overlapping in large multi-room floorplans.",
            "Future Roadmap: Integration of 2D LiDAR SLAM, IR beacon auto-docking charging station, and mobile IoT dashboard.",
        ]
    )

    # =======================================================================
    # SLIDE 25: Closing / Thank You Slide
    # =======================================================================
    s25 = prs.slides[24]
    for sh in s25.shapes:
        if sh.name == "TextBox 2":
            tf = sh.text_frame
            tf.text = "Student: Nguyễn Vũ Nguyên  |  ID: 2252948\nAdvisor: Dr. Phạm Phương Tùng"
            for p in tf.paragraphs:
                p.font.name = FONT_TITLE
                p.font.size = Pt(16)
                p.font.bold = True
                p.font.color.rgb = COLOR_NAVY
                p.alignment = PP_ALIGN.CENTER

    # =======================================================================
    # Global Sanitization: Timing Tags & Speaker Notes
    # =======================================================================
    for s in prs.slides:
        # Purge any orphaned timing / animation blocks
        timing = s._element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}timing")
        if timing is not None:
            s._element.remove(timing)

        # 100% wipe all speaker notes
        if s.has_notes_slide:
            try:
                tf = s.notes_slide.notes_text_frame
                if tf is not None:
                    tf.text = ""
            except Exception:
                pass

    # Save presentation with Windows file lock guard
    try:
        prs.save(output_pptx_path)
        final_path = output_pptx_path
    except PermissionError:
        base, ext = os.path.splitext(output_pptx_path)
        fallback_path = f"{base}_updated{ext}"
        print(f"[WARN] Target file locked by PowerPoint. Saving to fallback: {fallback_path}")
        prs.save(fallback_path)
        final_path = fallback_path

    return final_path


if __name__ == "__main__":
    template = r"D:\New folder (5)\2152082.Phạm Anh Hoàng.pptx"
    out_pptx = r"D:\New folder (5)\VuNguyen\DATN_2252948_NguyenVuNguyen_Defense.pptx"
    figures_dir = r"D:\New folder (5)\VuNguyen\extracted_figures_thesis"

    res = build_presentation(template, out_pptx, figures_dir)
    print(f"[OK] Generated 25-slide defense deck successfully: {res}")
