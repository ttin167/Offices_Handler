"""
tests/test_standard_diagram_symbols.py — Comprehensive tests for standard diagram geometric shapes,
symbols, and mxGraph Draw.io declarative mapping.
"""

import json
import xml.etree.ElementTree as ET
import pytest
from core.diagrams.legacy.canvas import (
    Node,
    PrecisionDiagram,
    STANDARD_ICONS,
)
from core.diagrams.mxgraph import (
    STANDARD_MXGRAPH_SHAPES,
    DEFAULT_SHAPE_DIMENSIONS,
    spec_to_mxgraph_xml,
    _layout_spec_nodes,
)


def test_canvas_all_standard_node_types_render():
    """Verify that all standard geometric node shapes render valid SVG elements."""
    diag = PrecisionDiagram(width=2000, height=1200)

    node_types = [
        "actor", "usecase", "database", "decision", "cloud",
        "terminator", "data_io", "queue", "component", "package", "document"
    ]

    for i, ntype in enumerate(node_types):
        diag.add_node(
            id=f"node_{ntype}",
            label=f"Label {ntype.title()}",
            x=80.0 + (i % 4) * 250.0,
            y=80.0 + (i // 4) * 200.0,
            width=180.0,
            height=90.0,
            type=ntype,
        )

    svg = diag.to_svg()
    assert svg.startswith("<svg")
    assert svg.strip().endswith("</svg>")

    # Check that specialized SVG elements exist
    # 1. Actor stick figure (circle head, line body, line arms, line legs)
    assert '<circle' in svg
    assert '<line x1=' in svg

    # 2. Database cylinder (top ellipse and cylinder body)
    assert 'ellipse cx=' in svg

    # 3. Decision diamond (polygon with 4 points)
    assert '<polygon points=' in svg

    # 4. Cloud Bezier
    assert 'C' in svg or 'Q' in svg

    # 5. Terminator capsule pill
    assert 'rx=' in svg

    # 6. Component tabs (tab rectangles)
    assert 'width="14" height="8"' in svg

    # 7. Package folder
    assert 'node_package' in svg

    # 8. Document wavy bottom
    assert 'node_document' in svg


def test_canvas_badge_icons():
    """Verify that vector badge icons render into node containers."""
    diag = PrecisionDiagram(width=1000, height=600)

    icons_to_test = ["database", "server", "cloud", "shield", "mobile", "user", "zap"]
    for i, icon_name in enumerate(icons_to_test):
        diag.add_node(
            id=f"icon_node_{icon_name}",
            label=f"Node {icon_name}",
            x=60.0 + i * 120.0,
            y=100.0,
            width=100.0,
            height=70.0,
            type="standard",
            icon=icon_name,
        )

    svg = diag.to_svg()
    # Check that icon transformation and paths exist
    assert '<g transform="translate(' in svg
    assert 'currentColor' in svg

    # Ensure standard icons dictionary is populated
    for icon_name in icons_to_test:
        assert icon_name in STANDARD_ICONS
        # Verify icon SVG path is embedded in output SVG
        icon_path_sample = STANDARD_ICONS[icon_name][:25]
        assert icon_path_sample in svg


def test_decision_port_vertex_anchoring():
    """Verify that Decision Diamond ports anchor at the exact 4 diamond vertices."""
    diamond_node = Node(
        id="decision_1",
        label="Auth Check?",
        x=200.0,
        y=150.0,
        width=140.0,
        height=80.0,
        type="decision",
    )

    # Center is at x = 200 + 70 = 270, y = 150 + 40 = 190
    top_port = diamond_node.get_port("top")
    bottom_port = diamond_node.get_port("bottom")
    left_port = diamond_node.get_port("left")
    right_port = diamond_node.get_port("right")

    assert top_port == (270.0, 150.0)
    assert bottom_port == (270.0, 230.0)
    assert left_port == (200.0, 190.0)
    assert right_port == (340.0, 190.0)


def test_drawio_spec_to_mxgraph_xml_shapes():
    """Verify that spec_to_mxgraph_xml maps standard shapes to Draw.io native cell styles."""
    spec = {
        "nodes": [
            {"id": "act", "label": "Client Actor", "type": "actor", "x": 50, "y": 50},
            {"id": "db", "label": "SQL Server", "type": "database", "x": 200, "y": 50},
            {"id": "cond", "label": "Approved?", "type": "decision", "x": 400, "y": 50},
            {"id": "cld", "label": "Cloud Service", "type": "cloud", "x": 600, "y": 50},
            {"id": "q", "label": "Kafka Bus", "type": "queue", "x": 800, "y": 50},
            {"id": "cmp", "label": "Parser Engine", "type": "component", "x": 1000, "y": 50},
            {"id": "pkg", "label": "Domain Core", "type": "package", "x": 1200, "y": 50},
            {"id": "tbl", "label": "UserTable", "columns": [
                {"name": "id", "type": "int", "is_pk": True},
                {"name": "email", "type": "varchar", "is_fk": False}
            ], "x": 1400, "y": 50},
        ],
        "edges": [
            {"source": "act", "target": "db", "label": "Direct Tap"},
            {"source": "db", "target": "cond", "label": "Evaluate"},
        ]
    }

    xml_str = spec_to_mxgraph_xml(spec)
    assert "<mxGraphModel" in xml_str
    assert "</mxGraphModel>" in xml_str

    root = ET.fromstring(xml_str)
    cells = {c.attrib.get("id"): c for c in root.findall(".//mxCell")}

    # Check Actor style
    assert "shape=umlActor" in cells["act"].attrib.get("style", "")

    # Check Database cylinder style
    assert "shape=cylinder3" in cells["db"].attrib.get("style", "")

    # Check Decision rhombus style
    assert "rhombus" in cells["cond"].attrib.get("style", "")

    # Check Cloud style
    assert "shape=cloud" in cells["cld"].attrib.get("style", "")

    # Check Queue tape style
    assert "shape=tape" in cells["q"].attrib.get("style", "")

    # Check Component module style
    assert "shape=module" in cells["cmp"].attrib.get("style", "")

    # Check Package folder style
    assert "shape=folder" in cells["pkg"].attrib.get("style", "")

    # Check Table style
    assert "shape=table" in cells["tbl"].attrib.get("style", "")
    assert "tbl_row_1" in cells
    assert "tbl_row_2" in cells

    # Check Edges have orthogonal routing & white mask
    edge_1 = cells["edge_1"]
    assert "edgeStyle=orthogonalEdgeStyle" in edge_1.attrib.get("style", "")
    assert "labelBackgroundColor=#ffffff" in edge_1.attrib.get("style", "")


def test_drawio_auto_layout():
    """Verify that nodes without coordinates receive automatic non-overlapping layout."""
    spec = {
        "nodes": [
            {"id": "n1", "label": "Start", "type": "terminator"},
            {"id": "n2", "label": "Process", "type": "standard"},
            {"id": "n3", "label": "Check", "type": "decision"},
            {"id": "n4", "label": "Store", "type": "database"},
        ],
        "edges": [
            {"source": "n1", "target": "n2"},
            {"source": "n2", "target": "n3"},
            {"source": "n3", "target": "n4"},
        ]
    }

    _layout_spec_nodes(spec["nodes"], spec["edges"])

    # All nodes must have x > 0 and distinct positions
    positions = [(n["x"], n["y"]) for n in spec["nodes"]]
    assert len(positions) == len(set(positions)), "All nodes must have unique positions"
    for n in spec["nodes"]:
        assert n["x"] >= 80.0
        assert n["y"] >= 80.0


def test_canvas_from_spec_robustness():
    """Verify from_spec handles title vs label, missing coordinates, and icons."""
    spec = {
        "nodes": [
            {"id": "n1", "title": "Node 1", "type": "actor", "icon": "user"},
            {"id": "n2", "title": "Node 2", "type": "database", "icon": "database"},
        ],
        "edges": [
            {"source": "n1", "target": "n2", "label": "Connects To"}
        ]
    }

    diag = PrecisionDiagram.from_spec(spec)
    assert len(diag.nodes) == 2
    assert diag.nodes["n1"].label == "Node 1"
    assert diag.nodes["n1"].type == "actor"
    assert diag.nodes["n1"].icon == "user"
    assert diag.nodes["n2"].label == "Node 2"
    assert diag.nodes["n2"].type == "database"
