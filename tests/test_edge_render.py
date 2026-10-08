import os
import subprocess
import tempfile
import unittest
from PIL import Image


class TestEdgeRender(unittest.TestCase):
    def test_edge_screenshot(self):
        svg_content = """<svg xmlns="http://www.w3.org/2000/svg" width="400" height="200" viewBox="0 0 400 200">
  <rect width="100%" height="100%" fill="#ffffff"/>
  <rect x="50" y="50" width="300" height="100" fill="#ffffff" stroke="#000000" stroke-width="2.5" rx="4"/>
  <text x="200" y="105" font-family="Segoe UI, Arial" font-size="16" font-weight="bold" fill="#000000" text-anchor="middle">Test Precision Engine</text>
</svg>"""

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_svg = os.path.join(tmpdir, "test_edge.svg")
            tmp_png = os.path.join(tmpdir, "test_edge.png")

            with open(tmp_svg, "w", encoding="utf-8") as f:
                f.write(svg_content)

            edge_candidates = [
                r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            ]
            edge_path = next((p for p in edge_candidates if os.path.isfile(p)), None)
            if not edge_path:
                self.skipTest("msedge.exe not found on system")

            cmd = [
                edge_path,
                "--headless",
                "--disable-gpu",
                "--hide-scrollbars",
                "--force-device-scale-factor=3",
                "--window-size=400,200",
                f"--screenshot={tmp_png}",
                tmp_svg,
            ]
            res = subprocess.run(cmd, capture_output=True, text=True)
            self.assertEqual(res.returncode, 0)
            self.assertTrue(os.path.isfile(tmp_png))
            with Image.open(tmp_png) as im:
                self.assertGreater(im.size[0], 0)


if __name__ == "__main__":
    unittest.main()

