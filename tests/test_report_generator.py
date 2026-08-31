from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document
from docx.oxml.ns import qn

from bridge import DesktopBridge
from backend import minimal_docx, report_engine as engine


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

    def test_resource_template_path_uses_application_root(self) -> None:
        self.assertEqual(
            engine.resource_template_path("x.docx"),
            engine.application_root() / "templates" / "x.docx",
        )


def build_demo_segments():
    return [
        {
            "grade": "一级",
            "manager": "重庆交通行政执法总队",
            "start": 2264000.0,
            "end": 2265000.0,
            "mileage": 1.0,
        },
        {
            "grade": "一级",
            "manager": "重庆交通行政执法总队",
            "start": 2265000.0,
            "end": 2266000.0,
            "mileage": 1.0,
        },
    ]


def build_demo_height_records():
    records = []
    for index, (start, end) in enumerate([(2264000.0, 2265000.0), (2265000.0, 2266000.0)]):
        for offset in (10, 20, 30):
            records.append({
                "file": "交安设施现场检测明细.xlsx",
                "direction": "上行",
                "station": start + offset * 10,
                "raw_station": start + offset * 10,
                "electronic_station": start + offset * 10,
                "basis": "电子修正桩号",
                "kind": "二波",
                "height": 580.0 + (offset * 2 % 40),
                "segment": index,
            })
        for offset in (40, 50):
            records.append({
                "file": "交安设施现场检测明细.xlsx",
                "direction": "上行",
                "station": start + offset * 10,
                "raw_station": start + offset * 10,
                "electronic_station": start + offset * 10,
                "basis": "电子修正桩号",
                "kind": "三波",
                "height": 677.0 + (offset % 20),
                "segment": index,
            })
    return records


def build_demo_bolt_records():
    return [
        {
            "file": "交安设施现场检测明细.xlsx",
            "direction": "下行",
            "station": 2264010.0,
            "raw_station": 2264010.0,
            "electronic_station": 2264010.0,
            "basis": "原始桩号",
            "segment": 0,
            "splice": 80,
            "splice_missing": 3,
            "connection": 120,
            "connection_missing": 1,
        }
    ]


def build_demo_stats(segments, records):
    stats = []
    for index in range(len(segments)):
        kinds = {"二波": [], "三波": []}
        bins = {"二波": [0] * 5, "三波": [0] * 5}
        for record in records:
            if record["segment"] != index:
                continue
            kinds[record["kind"]].append(record["height"])
            bin_index = report_engine_bin(record)
            bins[record["kind"]][bin_index] += 1
        types = {}
        for kind in ("二波", "三波"):
            heights = kinds[kind]
            pcts = [round(v * 100 / len(heights), 2) if heights else 0 for v in bins[kind]]
            types[kind] = {
                "count": len(heights),
                "bins": bins[kind],
                "pcts": pcts,
                "pass": pcts[2] if heights else 0,
            }
        stats.append({"segment": segments[index], "types": types})
    return stats


def report_engine_bin(record):
    kind = record["kind"]
    if kind == "二波":
        limits = (560, 580, 620, 640)
    else:
        limits = (657, 677, 717, 737)
    height = record["height"]
    if height < limits[0]:
        return 0
    if height < limits[1]:
        return 1
    if height <= limits[2]:
        return 2
    if height <= limits[3]:
        return 3
    return 4


class MinimalDocxTests(unittest.TestCase):
    def test_generates_report_without_template(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir)
            config = engine.Config(
                project_dir=Path(temp_dir),
                summary_xlsx=Path(temp_dir) / "summary.xlsx",
                detail_dir=Path(temp_dir),
                template_docx=Path(temp_dir) / "missing.docx",
                output_dir=output,
                disease_dir=None,
            )
            segments = build_demo_segments()
            records = build_demo_height_records()
            bolt_records = build_demo_bolt_records()
            height_stats = build_demo_stats(segments, records)
            bolt_stats = engine.make_bolt_stats(segments, bolt_records)
            messages = []
            config.out_docx.unlink(missing_ok=True)
            result = minimal_docx.run(
                config,
                segments,
                height_stats,
                records,
                bolt_stats,
                bolt_records,
                None,
                log=messages.append,
            )
            self.assertTrue(result.is_file())
            self.assertGreater(result.stat().st_size, 12000)
            self.assertTrue(any("程序化" in message for message in messages))

    def test_engine_routes_to_minimal_without_template(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            output = root / "output"
            config = engine.Config(
                project_dir=root,
                summary_xlsx=root / "summary.xlsx",
                detail_dir=root,
                template_docx=root / "missing.docx",
                output_dir=output,
                disease_dir=None,
            )
            segments = build_demo_segments()
            records = build_demo_height_records()
            bolt_records = build_demo_bolt_records()
            height_stats = build_demo_stats(segments, records)
            bolt_stats = engine.make_bolt_stats(segments, bolt_records)
            messages = []
            engine.make_docx(
                config,
                segments,
                height_stats=height_stats,
                height_records=records,
                bolt_stats=bolt_stats,
                bolt_records=bolt_records,
                disease_image_index=None,
                log=messages.append,
                require_template=False,
            )
            self.assertTrue(config.out_docx.is_file())
            self.assertGreater(config.out_docx.stat().st_size, 12000)
            self.assertTrue(any("程序化" in message for message in messages))


class GuangdongBusinessRegressionTests(unittest.TestCase):
    def test_scanner_preserves_bridge_guardrail_note(self) -> None:
        scanner = engine.GuangdongInputScanner(".")
        note_row = scanner._convert(
            "height",
            {
                "地市": "佛山市",
                "路线编号": "G1",
                "方向": "上行",
                "电子修正桩号": "K1+000",
                "检测区段": "K1+000～K2+000",
                "护栏类型": "二波",
                "梁板中心高度(mm)": "",
                "备注标记": "桥梁地段",
            },
            Path("guardrail.xlsx"),
            "Sheet1",
        )
        self.assertIsNotNone(note_row)
        self.assertEqual(note_row["guardrail_note"], "桥梁地段")

    def test_writer_handles_single_wave_and_bridge_only_segments(self) -> None:
        def height(segment: str, kind: str, value: float) -> dict:
            return {
                "category": "高速公路",
                "route": "G1",
                "direction": "上行",
                "segment": segment,
                "guardrail_type": kind,
                "height": value,
            }

        def note(segment: str) -> dict:
            return {
                "category": "高速公路",
                "route": "G1",
                "direction": "上行",
                "segment": segment,
                "guardrail_note": "桥梁地段",
            }

        bundle = {
            "marking": [],
            "height": [
                height("K1+000～K2+000", "二波", 600),
                height("K2+000～K3+000", "三波", 700),
            ],
            "bolt": [],
            "notes": [note("K3+000～K4+000")],
            "comparison_detail": [],
            "weak_segments": [],
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "template.docx"
            Document().save(template)
            output = engine.GuangdongChapterWriter.write(
                "佛山市",
                bundle,
                root / "output",
                template,
                {"marking": 7, "height": 5, "bolt": 5},
            )
            document = Document(output)
            text = "\n".join(paragraph.text for paragraph in document.paragraphs)

        self.assertIn("其中二波护栏", text)
        self.assertIn("其中三波护栏", text)
        self.assertEqual(text.count("当前区段为桥梁路段，无有效检测点位。"), 2)
        self.assertNotIn("区段内无护栏", text)
        self.assertNotIn("共检测0个有效点，其中。", text)

    def test_add_table_emits_expected_physical_and_logical_grid(self) -> None:
        document = Document()
        headers = [
            "路线",
            "护栏类型",
            "护栏位置",
            "方向",
            "桩号范围",
            "人工复核螺栓缺失数量",
            "拼接螺栓缺失数量",
            "自动化螺栓缺失数量",
            "连接螺栓缺失数量",
            "备注",
        ]
        try:
            table = engine.GuangdongChapterWriter._add_table(
                document,
                headers,
                [["G1", "二波", "左侧", "上行", "K1+000～K2+000", 1, 2, ""]],
                merge={
                    5: ("人工复核螺栓缺失数量", 2),
                    7: ("自动化螺栓缺失数量", 2),
                },
            )
        except TypeError as exc:
            self.fail(f"_add_table 尚不支持两级合并表头：{exc}")

        top_cells = table._tbl.findall("./w:tr", table._tbl.nsmap)[0].findall("./w:tc", table._tbl.nsmap)
        spans = []
        for cell in top_cells:
            span = cell.find("./w:tcPr/w:gridSpan", cell.nsmap)
            spans.append(int(span.get(qn("w:val"))) if span is not None else 1)
        grid_columns = table._tbl.findall("./w:tblGrid/w:gridCol", table._tbl.nsmap)
        self.assertEqual(len(top_cells), 8)
        self.assertEqual(sum(spans), 10)
        self.assertEqual(len(grid_columns), 10)
        self.assertEqual(spans, [1, 1, 1, 1, 1, 2, 2, 1])

    def test_comparison_analysis_uses_average_without_min_max_range(self) -> None:
        document = Document()
        detail = [
            {
                "indicator": "height",
                "absolute_difference": 1.0,
                "relative_deviation": None,
                "within_threshold": True,
            },
            {
                "indicator": "height",
                "absolute_difference": 3.0,
                "relative_deviation": None,
                "within_threshold": False,
            },
        ]
        engine.GuangdongChapterWriter._comparison_table(
            document,
            detail,
            {"marking": 7, "height": 5, "bolt": 5},
        )
        text = "\n".join(paragraph.text for paragraph in document.paragraphs)
        self.assertIn("平均偏差", text)
        self.assertIn("一致性占比", text)
        self.assertNotIn("1.00～3.00", text)


if __name__ == "__main__":
    unittest.main()
