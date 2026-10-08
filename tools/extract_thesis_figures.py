"""
extract_thesis_figures.py — Precision 300 DPI Figure Extractor for Thesis Defense Presentation.

Extracts all key technical figures from:
'D:\\New folder (5)\\VuNguyen\\Nguyễn Vũ Nguyên_2252948_Thuyetminh.pdf'
and saves them into:
'D:\\New folder (5)\\VuNguyen\\extracted_figures_thesis\\'

Tightly bounded clips: 100% clean diagrams with ZERO surrounding body paragraphs or book headings.
"""

from __future__ import annotations

import os
import re
import sys
from typing import Dict, List, Optional, Tuple

sys.stdout.reconfigure(encoding="utf-8")
import fitz  # PyMuPDF


def extract_figure_clip(
    doc: fitz.Document,
    page_num_1based: int,
    output_filename: str,
    output_dir: str,
    top_y: float,
    bottom_y: float,
    left_x: float = 40.0,
    right_x: Optional[float] = None,
    dpi: int = 300,
) -> Optional[str]:
    """Crops and saves an exact rectangular region from a PDF page at specified DPI."""
    page_idx = page_num_1based - 1
    if page_idx < 0 or page_idx >= len(doc):
        print(f"[WARN] Page {page_num_1based} out of range.")
        return None

    page = doc[page_idx]
    page_w = page.rect.width
    r_x = page_w - 40.0 if right_x is None else right_x

    clip_rect = fitz.Rect(left_x, top_y, r_x, bottom_y)
    pix = page.get_pixmap(dpi=dpi, clip=clip_rect)

    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, output_filename)
    pix.save(out_path)
    print(f"[OK] Saved {output_filename} ({pix.width}x{pix.height}) from P.{page_num_1based} [y: {top_y:.0f}-{bottom_y:.0f}]")
    return out_path


def main() -> None:
    pdf_path = r"D:\New folder (5)\VuNguyen\Nguyễn Vũ Nguyên_2252948_Thuyetminh.pdf"
    out_dir = r"D:\New folder (5)\VuNguyen\extracted_figures_thesis"
    doc = fitz.open(pdf_path)

    # Master registry of tightly-bounded figure regions:
    # (page_1based, filename, top_y, bottom_y)
    figures_to_extract = [
        # Chapter 1
        (15, "fig_01_02_floor_coverage_metric.png", 366.0, 626.0),
        (21, "fig_01_08_cleaning_mechanisms.png", 145.0, 652.0),
        (23, "fig_01_11_electrical_overview_block.png", 140.0, 663.0),
        (28, "fig_01_16_general_operating_process.png", 285.0, 730.0),

        # Chapter 2
        (37, "fig_02_01_vacuum_electrical_block.png", 70.0, 270.0),
        (39, "fig_02_02_functional_control_arch.png", 220.0, 600.0),
        (40, "fig_02_03_hardware_control_arch.png", 150.0, 745.0),
        (41, "fig_02_05_hybrid_coverage_process.png", 250.0, 528.0),

        # Chapter 3
        (43, "fig_03_01_driving_wheel_forces.png", 92.0, 298.0),
        (43, "fig_03_02_robot_forces.png", 350.0, 660.0),
        (46, "fig_03_04_brush_assembly_parts.png", 240.0, 704.0),
        (51, "fig_03_08_cantilever_shaft_fbd.png", 395.0, 746.0),
        (54, "fig_03_11_dustbin_3d_model.png", 80.0, 270.0),
        (56, "fig_03_14_rollover_analysis_front.png", 248.0, 429.0),
        (61, "fig_03_15_upper_view_3d.png", 166.0, 415.0),
        (61, "fig_03_16_underside_view_3d.png", 492.0, 741.0),

        # Chapter 4
        (70, "fig_04_01_system_wiring_diagram.png", 366.0, 686.0),
        (72, "fig_04_04_sensor_circuit_schematic.png", 272.0, 418.0),
        (73, "fig_04_06_pcb_altium_layers.png", 270.0, 453.0),
        (73, "fig_04_07_pcb_3d_preview.png", 474.0, 655.0),

        # Chapter 5
        (75, "fig_05_01_fsm_state_transition.png", 85.0, 328.0),
        (78, "fig_05_03_rpm_pwm_relationship.png", 180.0, 712.0),
        (79, "fig_05_04_left_motor_step_response.png", 85.0, 368.0),
        (79, "fig_05_05_right_motor_step_response.png", 390.0, 735.0),
        (80, "fig_05_07_left_motor_simulink.png", 545.0, 746.0),
        (81, "fig_05_08_left_motor_pole_zero.png", 220.0, 728.0),
        (84, "fig_05_12_irc_kinematics.png", 125.0, 623.0),
        (87, "fig_05_13_heading_step_responses.png", 85.0, 394.0),
        (90, "fig_05_16_heading_pole_zero.png", 85.0, 507.0),

        # Chapter 6
        (92, "fig_06_01_main_algorithm_flow.png", 248.0, 660.0),
        (93, "fig_06_03_spiral_algorithm.png", 322.0, 716.0),
        (94, "fig_06_05_obstacle_avoidance_alg.png", 230.0, 638.0),
        (94, "fig_06_06_stuck_recovery_alg.png", 230.0, 620.0),

        # Chapter 7
        (95, "fig_07_01_overall_robot_model.png", 146.0, 318.0),
        (95, "fig_07_02_brush_assembly_real.png", 490.0, 683.0),
        (96, "fig_07_03_fabricated_pcb_real.png", 120.0, 322.0),
        (97, "fig_07_04_straight_line_1m_test.png", 355.0, 716.0),
        (99, "fig_07_07_rotation_90deg_test.png", 85.0, 228.0),
        (99, "fig_07_08_rotation_180deg_test.png", 270.0, 422.0),
        (100, "fig_07_09_obstacle_sensor_response.png", 100.0, 300.0),
        (100, "fig_07_10_cliff_sensor_response.png", 325.0, 725.0),
        (104, "fig_07_14_coverage_empty_room_10min.png", 96.0, 305.0),
        (105, "fig_07_16_coverage_obstacles_10min.png", 292.0, 475.0),
        (106, "fig_07_17_paper_scraps_test.png", 85.0, 492.0),
        (106, "fig_07_18_dry_sand_test.png", 515.0, 744.0),
    ]

    print(f"Extracting {len(figures_to_extract)} tightly-bounded clean figures...")
    success = 0
    for pno, fname, top_y, bot_y in figures_to_extract:
        res = extract_figure_clip(doc, pno, fname, out_dir, top_y, bot_y, dpi=300)
        if res and os.path.isfile(res):
            success += 1

    print(f"\n[DONE] Successfully extracted {success}/{len(figures_to_extract)} clean figures.")


if __name__ == "__main__":
    main()
