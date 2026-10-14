"""
tools/test_phase2_deck.py — Generates a Phase 2 Showcase Presentation Deck.
Tests:
- Cyber Tech Dark & Minimalist Nordic Themes
- Desktop macOS Browser Mockup Frame (Traffic lights, address bar, aspect locked viewport)
- Smartphone Mockup Frame (Dynamic Island, rounded chassis, home indicator)
- Split Screen Mockup (Card left + Browser mockup right)
- Vector Illustration Asset Alias Resolving
"""

import os
import sys
import json

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pptx_writer import write_pptx_from_spec
from pptx_engine import init_presentation

spec = {
    "theme": "cyber_tech_dark",
    "slides": [
        {
            "type": "title",
            "title": "Antigravity Cloud Infrastructure",
            "subtitle": "Phase 2 Presentation Engine & Device Mockup Showcase",
            "presenter": "Antigravity Engineering Core",
            "date": "2026-10-09",
        },
        {
            "type": "mockup_browser",
            "title": "Universal Cloud Management Console",
            "subtitle": "Native Vector DrawingML Browser Frame (macOS Style)",
            "image_path": "cloud",
            "url": "https://console.antigravity.io/dashboard",
            "caption": "Figure 1: High-Availability Microservices Topology",
        },
        {
            "type": "mockup_mobile",
            "title": "Mobile Zero-Trust Authenticator",
            "subtitle": "Native DrawingML Smartphone Chassis with Dynamic Island",
            "image_path": "security_shield",
            "caption": "Figure 2: Real-time MFA & Security Guard Screen",
        },
        {
            "type": "split_mockup",
            "title": "Live Predictive Analytics Dashboard",
            "subtitle": "Nordic Minimalist Theme Split Layout",
            "theme": "minimalist_nordic",
            "bullets": [
                ("Stream Processing", "50,000 telemetry events per second processed with sub-millisecond jitter."),
                ("Automated Insights", "Real-time anomaly detection with statistical variance scoring."),
                ("Seamless Scaling", "Instant zero-downtime horizontal scaling across Kubernetes nodes."),
            ],
            "image_path": "data_analytics",
            "mockup": "browser",
            "url": "https://analytics.antigravity.io/metrics",
            "caption": "Figure 3: Production Metrics Stream (300 DPI Vector Viewport)",
        },
    ]
}

out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".web_outputs")
os.makedirs(out_dir, exist_ok=True)
out_file = os.path.join(out_dir, "phase2_showcase_deck.pptx")

result = write_pptx_from_spec(spec, out_file)
print(f"[SUCCESS] Phase 2 Showcase Deck created successfully at: {result}")
