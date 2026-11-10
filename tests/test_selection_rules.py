"""Regression checks for business mistakes described in the project record."""
import importlib.util
import json
from pathlib import Path
import sys
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import kb_tool as K


class SelectionRulesTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        data = json.loads((ROOT / "examples" / "catalog.json").read_text(encoding="utf-8"))
        cls.camera4k, cls.camera2k = data[0]["subcategories"][0]["products"]

    def test_4k_requirement_rejects_2k_camera(self):
        gaps = K.param_gaps("会议摄像机", "4K，12倍光学变焦", self.camera2k)
        self.assertTrue(any(level == "❌" and "4K" in reason for level, reason in gaps))

    def test_4k_requirement_accepts_4k_camera(self):
        self.assertEqual(K.param_gaps("会议摄像机", "4K，12倍光学变焦", self.camera4k), [])

    def test_stopped_product_is_excluded_from_candidates(self):
        stopped = dict(self.camera4k, available=False)
        # The original loader adds searchable text and category metadata.
        stopped.update(_text="4K会议摄像机 12倍光学变焦", _cat="会议摄像机/会议摄像机")
        with patch.object(K, "_KB", [stopped]):
            candidates, _ = K.rank("会议摄像机", "4K", kb=[stopped])
        self.assertEqual(candidates, [])

    def test_door_access_is_out_of_scope(self):
        system, ours, _ = K.system_of("门禁控制器", "门禁系统")
        self.assertFalse(ours)
        self.assertIn("门禁", system)

    def test_wireless_link_types_are_distinct(self):
        self.assertNotEqual(K.link_of("无线会议/U段系列"), K.link_of("无线会议/5G WiFi系列"))


class SelfcheckFailureTest(unittest.TestCase):
    def test_checker_crash_is_not_reported_as_success(self):
        spec = importlib.util.spec_from_file_location("public_selfcheck", ROOT / "plugins" / "fyys-selfcheck" / "__init__.py")
        plugin = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(plugin)
        import subprocess
        failure = subprocess.CompletedProcess([], 2, "", "checker failed")
        with patch.object(plugin.subprocess, "run", return_value=failure):
            result = plugin._check("fictional.xlsx", False)
        self.assertEqual(result[0], "ERR")
        self.assertIn("没跑成", plugin._note("fictional.xlsx", result))


if __name__ == "__main__":
    unittest.main()
