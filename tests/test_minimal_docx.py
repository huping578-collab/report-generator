from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from backend import minimal_docx, report_engine as engine


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
    def test_generates_report_without_template(self):
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

    def test_engine_routes_to_minimal_without_template(self):
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


if __name__ == "__main__":
    unittest.main()
