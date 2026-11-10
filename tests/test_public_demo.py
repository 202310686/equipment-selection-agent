"""Verify that the public example works without private data or model credentials."""
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

import openpyxl

ROOT = Path(__file__).resolve().parents[1]


class PublicDemoTest(unittest.TestCase):
    def test_export_preserves_input_and_matches_synthetic_camera(self):
        with tempfile.TemporaryDirectory() as tmp:
            # The public demo must work even when a Windows shell defaults to GBK.
            env = dict(os.environ, PYTHONUTF8="0", PYTHONIOENCODING="gbk")
            result = subprocess.run(
                [sys.executable, str(ROOT / "examples" / "demo.py"), "--output", tmp],
                capture_output=True, text=True, encoding="utf-8", env=env, timeout=60,
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            path = Path(tmp) / "demo_比单表.xlsx"
            self.assertTrue(path.is_file())
            wb = openpyxl.load_workbook(path, data_only=True)
            try:
                ws = wb.active
                rows = list(ws.values)
                self.assertEqual(rows[1][:4], (1, "会议摄像机", "4K，12倍光学变焦", 1))
                model_col = rows[0].index("我司型号")
                self.assertEqual(rows[1][model_col], "HT-DEMO-CAM4K")
                self.assertIsNone(rows[2][model_col])
                self.assertIn("自检", result.stdout)
            finally:
                wb.close()


if __name__ == "__main__":
    unittest.main()
