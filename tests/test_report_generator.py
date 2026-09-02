from __future__ import annotations

import json
import os
import tempfile
import unittest
import zipfile
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
        self.cq_template = self.root / "重庆项目报告模板.md"
        self.gd_template = self.root / "广东项目第五章模板.md"
        self.cq_template.write_text("# 重庆模板\n", encoding="utf-8")
        self.gd_template.write_text("# 广东模板\n", encoding="utf-8")
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

    def test_native_window_reference_is_private_from_js_api(self) -> None:
        window = FakeWindow()
        self.bridge.attach_window(window)
        self.assertNotIn("window", vars(self.bridge))
        self.assertIs(self.bridge._window, window)

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
        templates["广东项目第五章模板"] = self.root / "missing-gd.md"
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
            engine.resource_template_path("x.md"),
            engine.application_root() / "templates" / "x.md",
        )

    def test_guangdong_run_uses_resource_template_for_bundles(self) -> None:
        values = {
            "projectPath": str(self.project),
            "markingPath": "",
            "guardrailPath": "",
            "manualPath": str(self.manual),
            "routePath": str(self.route),
            "outputPath": str(self.output),
            "markingThreshold": "7",
            "heightThreshold": "5",
            "boltThreshold": "5",
        }
        scanned = {
            "marking": [{"city": "佛山市", "route": "G1"}],
            "height": [],
            "bolt": [],
            "notes": [],
            "issues": [],
        }
        expected_template = self.root / "resource-template.md"
        expected_template.write_text("# 模板", encoding="utf-8")
        route_index = type("RouteIndex", (), {"mapping": {}})()
        with (
            patch.object(engine.RouteCategoryIndex, "from_file", return_value=route_index),
            patch.object(engine.GuangdongInputScanner, "scan", return_value=scanned),
            patch.object(engine.ManualAutoComparator, "read_file", return_value=([], [])),
            patch.object(engine.GuangdongBatchRunner, "build_bundles", return_value={}),
            patch.object(engine.GuangdongBatchRunner, "run_bundles", return_value={}) as run_bundles,
            patch.object(engine, "resource_template_path", return_value=expected_template) as resource_path,
        ):
            self.bridge._run_guangdong(values)

        resource_path.assert_called_once_with("广东项目第五章模板.md")
        self.assertEqual(run_bundles.call_args.args[2], expected_template)


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


class UpdaterTests(unittest.TestCase):
    def test_updater_core_validation(self) -> None:
        from updater import is_newer, parse_version, select_installer_asset, validate_download

        self.assertEqual(parse_version("v1.2.3"), (1, 2, 3))
        self.assertTrue(is_newer((1, 2, 4), (1, 2, 3)))
        self.assertFalse(is_newer((1, 2, 3), (1, 2, 3)))
        asset = select_installer_asset({
            "assets": [{
                "name": "report-generator-Setup.exe",
                "browser_download_url": "https://github.com/a/b/releases/download/v1/报告生成工具-Setup.exe",
                "size": 2,
            }],
        })
        self.assertEqual(asset["name"], "report-generator-Setup.exe")
        with self.assertRaises(ValueError):
            validate_download(b"not-an-exe", 11)


    def test_updater_rejects_untrusted_assets_and_checks_release(self) -> None:
        from updater import check_for_update, select_installer_asset

        release = {
            "tag_name": "v0.2.0",
            "assets": [{
                "name": "report-generator-Setup.exe",
                "browser_download_url": "https://objects.githubusercontent.com/a/b.exe",
                "size": 2,
            }],
        }
        result = check_for_update("0.1.0", lambda: release)
        self.assertTrue(result["update_available"])
        self.assertEqual(result["latest_version"], "0.2.0")

        malicious = dict(release)
        malicious["assets"] = [{
            "name": "report-generator-Setup.exe",
            "browser_download_url": "http://example.com/update.exe",
            "size": 2,
        }]
        with self.assertRaises(ValueError):
            select_installer_asset(malicious)

    def test_updater_launcher_does_not_use_shell(self) -> None:
        from updater import launch_installer

        with patch("updater.subprocess.Popen") as popen:
            launch_installer(Path("C:/Temp/update.exe"))
        popen.assert_called_once_with([str(Path("C:/Temp/update.exe"))], close_fds=True)


class MinimalDocxTests(unittest.TestCase):
    def test_generates_report_without_template(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir)
            skeleton = Path(temp_dir) / "重庆项目报告模板.md"
            skeleton.write_text(
                "# 2026年普通公路国省道交通安全设施自动化检测报告\n\n"
                "## 1.1 项目概况\n\n"
                "## 3.1 G210线整体情况\n\n"
                "## 5.1 G210线整体情况\n\n"
                "## 6 结论与建议\n",
                encoding="utf-8",
            )
            config = engine.Config(
                project_dir=Path(temp_dir),
                summary_xlsx=Path(temp_dir) / "summary.xlsx",
                detail_dir=Path(temp_dir),
                template_docx=skeleton,
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
                skeleton_md=skeleton,
            )
            self.assertTrue(result.is_file())
            self.assertGreater(result.stat().st_size, 12000)
            self.assertTrue(any("Markdown" in message or "程序化" in message for message in messages))

    def test_engine_routes_to_minimal_without_template(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            output = root / "output"
            skeleton = root / "重庆项目报告模板.md"
            skeleton.write_text(
                "# 2026年普通公路国省道交通安全设施自动化检测报告\n\n"
                "## 1.1 项目概况\n\n"
                "## 3.1 G210线整体情况\n\n"
                "## 5.1 G210线整体情况\n\n"
                "## 6 结论与建议\n",
                encoding="utf-8",
            )
            config = engine.Config(
                project_dir=root,
                summary_xlsx=root / "summary.xlsx",
                detail_dir=root,
                template_docx=skeleton,
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
            self.assertTrue(any("Markdown" in message or "程序化" in message for message in messages))

    def test_generates_report_from_markdown_skeleton(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            skeleton = root / "重庆项目报告模板.md"
            skeleton.write_text(
                "# 2026年普通公路国省道交通安全设施自动化检测报告\n\n"
                "## 1.1 项目概况\n\n"
                "本报告依据委托单位提供的数据生成。\n\n"
                "## 3.1 G210线整体情况\n\n"
                "### 3.2 G210线K2264+000~K2265+000段（样例）\n\n"
                "## 5 结论与建议\n\n"
                "## 6 建议\n\n"
                "标志牌安装情况检查合格。\n",
                encoding="utf-8",
            )
            config = engine.Config(
                project_dir=root,
                summary_xlsx=root / "summary.xlsx",
                detail_dir=root,
                template_docx=skeleton,
                output_dir=root / "output",
                disease_dir=None,
            )
            segments = build_demo_segments()
            records = build_demo_height_records()
            bolt_records = build_demo_bolt_records()
            height_stats = build_demo_stats(segments, records)
            bolt_stats = engine.make_bolt_stats(segments, bolt_records)
            result = minimal_docx.run(
                config,
                segments,
                height_stats,
                records,
                bolt_stats,
                bolt_records,
                None,
                skeleton_md=skeleton,
            )
            self.assertTrue(result.is_file())
            document = Document(result)
            text = "\n".join(paragraph.text for paragraph in document.paragraphs)
            self.assertIn("自动化检测报告", text)
            self.assertIn("本报告依据委托单位提供的数据生成。", text)
            self.assertIn("波形梁护栏横梁中心高度检测结果", text)
            self.assertIn("结论与建议", text)
            self.assertNotIn("（样例）", text)


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
            template = root / "template.md"
            template.write_text(
                "# 五、交通安全设施技术状况检测评价情况\n\n"
                "本章为{{地市}}交通安全设施技术状况检测评价内容。\n",
                encoding="utf-8",
            )
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

    def test_writer_deduplicates_height_and_note_segments_with_suffix(self) -> None:
        bundle = {
            "marking": [],
            "height": [{
                "category": "高速公路",
                "route": "G1",
                "direction": "上行",
                "segment": "K1+000～K2+000",
                "guardrail_type": "二波",
                "height": 600,
            }],
            "bolt": [],
            "notes": [{
                "category": "高速公路",
                "route": "G1",
                "direction": "上行",
                "segment": "K1+000～K2+000段",
                "guardrail_note": "桥梁地段",
            }],
            "comparison_detail": [],
            "weak_segments": [],
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "template.md"
            template.write_text(
                "# 五、交通安全设施技术状况检测评价情况\n\n"
                "本章为{{地市}}交通安全设施技术状况检测评价内容。\n",
                encoding="utf-8",
            )
            output = engine.GuangdongChapterWriter.write(
                "佛山市",
                bundle,
                root / "output",
                template,
                {"marking": 7, "height": 5, "bolt": 5},
            )
            document = Document(output)
            headings = []
            height_text = []
            in_height_section = False
            for paragraph in document.paragraphs:
                if paragraph.text == "（2）护栏中心高度":
                    in_height_section = True
                elif paragraph.text == "（3）螺栓安装情况":
                    in_height_section = False
                elif in_height_section:
                    height_text.append(paragraph.text)
                    if (
                        paragraph.style.name == "Heading 5"
                        and "K1+000～K2+000段" in paragraph.text
                    ):
                        headings.append(paragraph.text)

        self.assertEqual(headings, ["a. K1+000～K2+000段"])
        self.assertNotIn("当前区段为桥梁路段，无有效检测点位。", "\n".join(height_text))

    def test_writer_accepts_markdown_template_with_city_placeholder(self) -> None:
        bundle = {
            "marking": [],
            "height": [{
                "category": "高速公路",
                "route": "G1",
                "direction": "上行",
                "segment": "K1+000～K2+000",
                "guardrail_type": "二波",
                "height": 600,
            }],
            "bolt": [],
            "notes": [],
            "comparison_detail": [],
            "weak_segments": [],
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "template.md"
            template.write_text(
                "# 五、交通安全设施技术状况检测评价情况\n\n"
                "本章为{{地市}}交通安全设施技术状况检测评价内容。\n\n"
                "## （一）高速公路交安设施技术状况\n\n"
                "### 1.沿线设施技术状况TCI\n",
                encoding="utf-8",
            )
            output = engine.GuangdongChapterWriter.write(
                "佛山市",
                bundle,
                root / "output",
                template,
                {"marking": 7, "height": 5, "bolt": 5},
            )
            document = Document(output)
            text = "\n".join(paragraph.text for paragraph in document.paragraphs)

        self.assertIn("本章为佛山市交通安全设施技术状况检测评价内容。", text)
        self.assertIn("五、交通安全设施技术状况检测评价情况", text)
        self.assertIn("（2）护栏中心高度", text)
        self.assertIn("其中二波护栏", text)
        self.assertNotIn("{{地市}}", text)

    def test_comparison_table_emits_expected_two_level_bolt_header_xml(self) -> None:
        document = Document()
        engine.GuangdongChapterWriter._comparison_table(
            document,
            [{
                "indicator": "bolt",
                "route": "G1",
                "gtype": "二波",
                "position": "左侧",
                "direction": "上行",
                "segment": "K1+000～K2+000",
                "msplice": 1,
                "mconn": 2,
                "asplice": 1,
                "aconn": 2,
                "remark": "",
            }],
            {"marking": 7, "height": 5, "bolt": 5},
        )
        table = document.tables[1]
        xml_rows = table._tbl.findall("./w:tr", table._tbl.nsmap)
        cell_counts = [len(row.findall("./w:tc", table._tbl.nsmap)) for row in xml_rows]
        top_cells = xml_rows[0].findall("./w:tc", table._tbl.nsmap)
        spans = []
        for cell in top_cells:
            span = cell.find("./w:tcPr/w:gridSpan", cell.nsmap)
            spans.append(int(span.get(qn("w:val"))) if span is not None else 1)
        grid_columns = table._tbl.findall("./w:tblGrid/w:gridCol", table._tbl.nsmap)

        self.assertEqual(cell_counts, [8, 10, 10])
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


REAL_E2E_ENV = "REPORT_E2E_REAL"
_REAL_E2E_PATHS = {
    "route_xlsx": ("REPORT_E2E_ROUTE_XLSX", "file"),
    "manual_xlsx": ("REPORT_E2E_MANUAL_XLSX", "file"),
    "template": ("REPORT_E2E_TEMPLATE_DOCX", "file"),
    "foshan_marking": ("REPORT_E2E_FOSHAN_MARKING_DIR", "dir"),
    "foshan_guardrail": ("REPORT_E2E_FOSHAN_GUARDRAIL_DIR", "dir"),
    "zhuhai_marking": ("REPORT_E2E_ZHUHAI_MARKING_DIR", "dir"),
    "zhuhai_guardrail": ("REPORT_E2E_ZHUHAI_GUARDRAIL_DIR", "dir"),
}


def _required_real_path(key: str) -> Path:
    env_name, kind = _REAL_E2E_PATHS[key]
    value = os.environ.get(env_name, "").strip()
    if not value:
        raise AssertionError(f"真实数据测试必须设置环境变量：{env_name}")
    path = Path(value)
    valid = path.is_file() if kind == "file" else path.is_dir()
    if not valid:
        raise AssertionError(f"{env_name} 路径不存在或类型错误：{path}")
    return path


def _real_e2e_inputs() -> dict[str, Path]:
    return {key: _required_real_path(key) for key in _REAL_E2E_PATHS}


class RealDataConfigurationTests(unittest.TestCase):
    def test_real_e2e_requires_input_environment_variables(self) -> None:
        env = {REAL_E2E_ENV: "1"}
        env.update({env_name: "" for env_name, _ in _REAL_E2E_PATHS.values()})
        with patch.dict(os.environ, env, clear=False):
            with self.assertRaisesRegex(AssertionError, "REPORT_E2E_ROUTE_XLSX"):
                _real_e2e_inputs()


def _scan_city(city: dict, route_index) -> dict:
    """真实数据扫描：与 run_guangdong_project 中标线/护栏分别扫描的同款逻辑。"""
    scanned = {"height": [], "bolt": [], "marking": [], "notes": [], "issues": []}
    for key, kinds in (("marking", ("marking",)), ("guardrail", ("height", "bolt", "notes"))):
        root = city.get(key)
        if not root or not root.is_dir():
            continue
        partial = engine.GuangdongInputScanner(root, route_index).scan()
        for kind in kinds:
            scanned[kind].extend(partial.get(kind, []))
        scanned["issues"].extend(partial.get("issues", []))
    return scanned


def _run_real_pipeline(
    cities: list[dict],
    artifact_root: Path,
    route_xlsx: Path,
    manual_xlsx: Path,
    template: Path,
) -> dict:
    """使用运行环境提供的真实输入：直接拼装 scan+build_bundles+run_bundles。"""
    artifact_root.mkdir(parents=True, exist_ok=True)
    route_index = engine.RouteCategoryIndex.from_file(route_xlsx)
    manual_records, manual_issues = engine.ManualAutoComparator.read_file(manual_xlsx)
    thresholds = {"marking": 7, "height": 5, "bolt": 5}

    result = {"success": [], "failed": {}, "warnings": []}
    for city in cities:
        # 每个城市独立 scan/build/run，避免跨城市路线记录进入同一个 bundle。
        scanned = _scan_city(city, route_index)
        scanned["issues"] = list(manual_issues) + scanned["issues"]
        bundles = engine.GuangdongBatchRunner.build_bundles(scanned, route_index, manual_records, thresholds)
        partial = engine.GuangdongBatchRunner.run_bundles(
            bundles, artifact_root, template, thresholds
        )
        result["success"].extend(partial["success"])
        result["failed"].update(partial["failed"])
        result["warnings"].extend(partial["warnings"])
    return result


def _collect_docx_text(docx_path: Path) -> str:
    document = Document(str(docx_path))
    return "\n".join(paragraph.text for paragraph in document.paragraphs)


def _collect_segment_order(text: str, marker: str) -> list[str]:
    """按指标汇总表分段提取 a. 区段标题中的桩号范围。"""
    import re

    summary_markers = (
        "标线逆反射区段汇总表",
        "护栏中心高度区段汇总表",
        "护栏螺栓缺失区段汇总表",
    )
    aliases = {
        "标线逆反射区段": "标线逆反射区段汇总表",
        "护栏中心高度": "护栏中心高度区段汇总表",
        "螺栓": "护栏螺栓缺失区段汇总表",
    }
    wanted = aliases.get(marker, marker)
    if wanted not in summary_markers:
        raise ValueError(f"未知指标汇总表 marker：{marker}")

    title = re.compile(r"^\s*[a-z]\.\s+")
    pattern = re.compile(r"K\+?\d+(?:\.\d+)?[~～]K\+?\d+(?:\.\d+)?")
    result: list[str] = []
    active = False
    for line in text.splitlines():
        found = next((item for item in summary_markers if item in line), None)
        if found is not None:
            active = found == wanted
            continue
        if not active or not title.match(line):
            continue
        match = pattern.search(line)
        if match:
            result.append(match.group(0))
    return result


@unittest.skipUnless(os.environ.get(REAL_E2E_ENV) == "1", f"set {REAL_E2E_ENV}=1 to run")
class RealDataE2ETests(unittest.TestCase):
    """真实佛山/珠海端到端回归；产物仅落在 tests/artifacts/real-data/，不进 git。"""

    @classmethod
    def setUpClass(cls) -> None:
        paths = _real_e2e_inputs()
        cls.route_xlsx = paths["route_xlsx"]
        cls.manual_xlsx = paths["manual_xlsx"]
        cls.template = paths["template"]
        cls.foshan = {
            "name": "佛山",
            "marking": paths["foshan_marking"],
            "guardrail": paths["foshan_guardrail"],
        }
        cls.zhuhai = {
            "name": "珠海",
            "marking": paths["zhuhai_marking"],
            "guardrail": paths["zhuhai_guardrail"],
        }

    def _run(self, cities: list[dict], sub: str) -> dict:
        artifact_root = Path(__file__).resolve().parent / "artifacts" / "real-data" / sub
        if artifact_root.exists():
            # 清空旧产物，避免上轮 run 的 docx 干扰本次断言。
            import shutil

            shutil.rmtree(artifact_root, ignore_errors=True)
        return _run_real_pipeline(cities, artifact_root, self.route_xlsx, self.manual_xlsx, self.template)

    def test_foshan_minimal_integration_runs_end_to_end(self) -> None:
        result = self._run([self.foshan], "foshan")
        self.assertIn("佛山市", result["success"])
        foshan_dir = Path(__file__).resolve().parent / "artifacts" / "real-data" / "foshan" / "佛山市"
        docx_path = next(foshan_dir.glob("*第五部分.docx"), None)
        self.assertIsNotNone(docx_path, f"未生成佛山 docx：{foshan_dir}")
        self._assert_docx_complete(docx_path, "佛山")
        text = _collect_docx_text(docx_path)
        self.assertEqual(text.count("区段内无护栏"), 0)
        self.assertEqual(text.count("共检测0个有效点，其中。"), 0)

    def test_foshan_and_zhuhai_full_e2e_artifact_checks(self) -> None:
        result = self._run([self.foshan, self.zhuhai], "foshan-zhuhai")
        self.assertIn("佛山市", result["success"])
        self.assertIn("珠海市", result["success"], msg=f"珠海失败：{result['failed']}")
        self.assertFalse(result["failed"], f"运行失败：{result['failed']}")

        artifact_root = Path(__file__).resolve().parent / "artifacts" / "real-data" / "foshan-zhuhai"
        foshan_docx = next((artifact_root / "佛山市").glob("*第五部分.docx"), None)
        zhuhai_docx = next((artifact_root / "珠海市").glob("*第五部分.docx"), None)
        self.assertIsNotNone(foshan_docx, "佛山 docx 缺失")
        self.assertIsNotNone(zhuhai_docx, "珠海 docx 缺失")
        self._assert_docx_complete(foshan_docx, "佛山")
        self._assert_docx_complete(zhuhai_docx, "珠海")

        # 仅抽取检查中用得到的城市子串，避免佛山段落污染珠海或反之。
        foshan_text = _collect_docx_text(foshan_docx)
        zhuhai_text = _collect_docx_text(zhuhai_docx)

        for label, text in (("佛山", foshan_text), ("珠海", zhuhai_text)):
            self.assertEqual(text.count("区段内无护栏"), 0, f"{label}出现旧句：区段内无护栏")
            self.assertEqual(text.count("共检测0个有效点，其中。"), 0, f"{label}出现残句：共检测0个有效点，其中。")

        # 标线/高度/螺栓三章节共同区段顺序一致：抽取各章节首次出现的区段链，比较相等。
        for label, text in (("佛山", foshan_text), ("珠海", zhuhai_text)):
            for marker in ("标线逆反射区段", "护栏中心高度", "螺栓"):
                segments = _collect_segment_order(text, marker)
                self.assertGreater(len(segments), 1, f"{label}/{marker} 区段数不足，无法比序")
            marking = _collect_segment_order(text, "标线逆反射区段")
            height = _collect_segment_order(text, "护栏中心高度")
            bolt = _collect_segment_order(text, "螺栓")
            common = [s for s in marking if s in set(height) & set(bolt)]
            self.assertGreater(len(common), 1, f"{label} 共同区段数={len(common)}，无法比序")
            height_in_common = [s for s in height if s in set(common)]
            bolt_in_common = [s for s in bolt if s in set(common)]
            self.assertEqual(common, height_in_common, f"{label} 标线/高度共同区段相对顺序不一致")
            self.assertEqual(common, bolt_in_common, f"{label} 标线/螺栓共同区段相对顺序不一致")

        # 抽查珠海 docx 的一个螺栓对比表：物理 w:tc 8/10/10、两个父表头 gridSpan=2。
        bolt_table_index, bolt_table = self._find_bolt_comparison_table(zhuhai_docx)
        with zipfile.ZipFile(zhuhai_docx) as archive:
            xml = archive.read(f"word/document.xml")
        self._assert_two_level_bolt_header(
            xml, bolt_table_index, label=f"珠海 螺栓对比表 #{bolt_table_index}"
        )

    def _find_bolt_comparison_table(self, docx_path: Path):
        document = Document(str(docx_path))
        for index, table in enumerate(document.tables):
            text = "\n".join(cell.text or "" for row in table.rows for cell in row.cells)
            rows = table._tbl.findall("./w:tr", table._tbl.nsmap)
            cell_counts = [len(row.findall("./w:tc", table._tbl.nsmap)) for row in rows]
            if "拼接螺栓" in text and cell_counts[:3] == [8, 10, 10]:
                return index, table
        self.fail(f"未在 {docx_path} 中找到物理列为8/10/10的螺栓对比表")

    def _assert_docx_complete(self, docx_path: Path, label: str) -> None:
        with zipfile.ZipFile(docx_path) as archive:
            self.assertIsNone(archive.testzip(), f"{label} docx ZIP 存在损坏成员")

    def _assert_two_level_bolt_header(self, document_xml: bytes, table_index: int, label: str) -> None:
        from xml.etree import ElementTree as ET

        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        root = ET.fromstring(document_xml)
        tables = root.findall(f".//{{{ns['w']}}}tbl")
        if not 0 <= table_index < len(tables):
            self.fail(f"{label} XML 表格索引越界：{table_index}/{len(tables)}")
        table = tables[table_index]
        text = "".join(t.text or "" for t in table.iter(f"{{{ns['w']}}}t"))
        self.assertIn("拼接螺栓", text, f"{label} XML 表格内容不符")
        rows = table.findall(f"{{{ns['w']}}}tr")
        cell_counts = [len(r.findall(f"{{{ns['w']}}}tc")) for r in rows]
        self.assertGreaterEqual(len(cell_counts), 3, f"{label} 行数不足：{cell_counts}")
        self.assertEqual(cell_counts[:3], [8, 10, 10], f"{label} 物理列不符：{cell_counts}")
        self.assertTrue(all(count == 10 for count in cell_counts[2:]), f"{label} 数据行物理列不符：{cell_counts}")
        grid = table.find(f"{{{ns['w']}}}tblGrid")
        grid_columns = [] if grid is None else grid.findall(f"{{{ns['w']}}}gridCol")
        self.assertEqual(len(grid_columns), 10, f"{label} 物理网格列不符：{len(grid_columns)}")
        top = rows[0].findall(f"{{{ns['w']}}}tc")
        spans = []
        for cell in top:
            span = cell.find(f"{{{ns['w']}}}tcPr/{{{ns['w']}}}gridSpan")
            spans.append(int(span.get(f"{{{ns['w']}}}val")) if span is not None else 1)
        self.assertEqual(spans.count(2), 2, f"{label} 父表头 gridSpan=2 应有2个：{spans}")
        self.assertEqual(len(spans), 8, f"{label} 父表头数量不符：{spans}")


if __name__ == "__main__":
    unittest.main()
