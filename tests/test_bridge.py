from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from bridge import DesktopBridge


class FakeWindow:
    def __init__(self) -> None:
        self.scripts: list[str] = []

    def evaluate_js(self, script: str) -> None:
        self.scripts.append(script)


class DesktopBridgeTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.project = self.root / "project"
        self.project.mkdir()
        self.output = self.root / "output"
        self.summary = self.root / "summary.xlsx"
        self.manual = self.root / "manual.xlsx"
        self.route = self.root / "route.xlsx"
        for path in (self.summary, self.manual, self.route):
            path.write_bytes(b"test")
        self.detail = self.root / "detail"
        self.disease = self.root / "disease"
        self.detail.mkdir()
        self.disease.mkdir()
        self.cq_template = self.root / "重庆项目报告模板.docx"
        self.gd_template = self.root / "广东项目第五章模板.docx"
        self.cq_template.write_bytes(b"template")
        self.gd_template.write_bytes(b"template")
        self.bridge = DesktopBridge()

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def templates(self) -> dict[str, Path]:
        return {
            "重庆项目报告模板": self.cq_template,
            "广东项目第五章模板": self.gd_template,
        }

    def cq_payload(self) -> dict:
        return {
            "template": "cq",
            "values": {
                "projectPath": str(self.project),
                "summaryPath": str(self.summary),
                "detailPath": str(self.detail),
                "diseasePath": str(self.disease),
                "outputPath": str(self.output),
            },
        }

    def test_validate_chongqing_payload_creates_output(self) -> None:
        with patch.object(self.bridge, "_template_paths", return_value=self.templates()):
            self.bridge._validate_payload(self.cq_payload())
        self.assertTrue(self.output.is_dir())

    def test_start_run_validation_error_releases_lock(self) -> None:
        payload = self.cq_payload()
        payload["values"]["summaryPath"] = ""
        with patch.object(self.bridge, "_template_paths", return_value=self.templates()):
            result = self.bridge.start_run(payload)
        self.assertFalse(result["ok"])
        self.assertIn("分段汇总表", result["error"])
        self.assertFalse(self.bridge._running)

    def test_guangdong_specialized_folders_are_optional(self) -> None:
        payload = {
            "template": "gd",
            "values": {
                "projectPath": str(self.project),
                "markingPath": "",
                "guardrailPath": "",
                "manualPath": str(self.manual),
                "routePath": str(self.route),
                "outputPath": str(self.output),
                "markingThreshold": "7",
                "heightThreshold": "5",
                "boltThreshold": "5",
            },
        }
        with patch.object(self.bridge, "_template_paths", return_value=self.templates()):
            self.bridge._validate_payload(payload)
        self.assertTrue(self.output.is_dir())

    def test_missing_specialized_template_fails_guangdong_validation(self) -> None:
        templates = self.templates()
        templates["广东项目第五章模板"] = self.root / "missing-gd.docx"
        scalar = {
            "template": "gd",
            "values": {
                "projectPath": str(self.project),
                "markingPath": "",
                "guardrailPath": "",
                "manualPath": str(self.manual),
                "routePath": str(self.route),
                "outputPath": str(self.output),
                "markingThreshold": "7",
                "heightThreshold": "5",
                "boltThreshold": "5",
            },
        }
        with patch.object(self.bridge, "_template_paths", return_value=templates):
            with self.assertRaisesRegex(FileNotFoundError, "templates 文件夹"):
                self.bridge._validate_payload(scalar)

    def test_worker_emits_complete_and_releases_running_state(self) -> None:
        window = FakeWindow()
        self.bridge.attach_window(window)
        self.bridge._running = True
        with patch.object(self.bridge, "_run_chongqing"):
            self.bridge._run_worker(self.cq_payload())
        self.assertFalse(self.bridge._running)
        events = [json.loads(script.split("desktopEvents(", 1)[1][:-1]) for script in window.scripts]
        self.assertEqual(events[0]["status"], "running")
        self.assertEqual(events[-1]["status"], "complete")
        self.assertEqual(events[-1]["progress"], 100)

    def test_progress_mapping(self) -> None:
        self.assertEqual(self.bridge._progress_from_log("正在扫描资料"), (34, 1))
        self.assertEqual(self.bridge._progress_from_log("图表已生成"), (72, 2))
        self.assertEqual(self.bridge._progress_from_log("已保存文件"), (92, 3))


if __name__ == "__main__":
    unittest.main()
