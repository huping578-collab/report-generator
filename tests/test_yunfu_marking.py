"""云浮宽表输入、共享区段和报告/Excel持久回归（不复制参考统计数）。"""
from collections import Counter
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from docx import Document

import openpyxl
from backend import report_engine as engine

HEADERS = ["地市", "路线编号", "检测方向", "管养单位", "桩号",
           "主车道左侧标线逆反亮度系数", "主车道右侧标线逆反亮度系数",
           "左侧标线逆反射目标值", "右侧标线逆反射目标值", "计算区间"]
TEMPLATE = Path(__file__).resolve().parents[1] / "templates/广东项目第五章模板.md"
REAL_INPUT = Path("C:/文件/工作工具台/广东报告数据/标线数据/云浮市标线统计.xlsx")


def wide_row(station="1000+000.0-1000+020.0", **changes):
    row = dict(zip(HEADERS, ["云浮", "G234", "上行", "甲管养单位", station, 40, 90, 50, 80, 20]))
    row.update(changes)
    return row


def scan_rows(rows, headers=HEADERS):
    with TemporaryDirectory() as tmp:
        path = Path(tmp) / "宽表.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(headers)
        for row in rows:
            ws.append([row.get(key) for key in headers])
        wb.save(path)
        wb.close()
        return engine.GuangdongInputScanner(tmp)._process_file(path)


class YunfuMarkingTest(unittest.TestCase):
    def test_wide_input_preserves_each_side_and_source_coordinates(self):
        rows = [wide_row(), wide_row("1000+020.0-1000+040.0", **{HEADERS[5]: None}),
                wide_row("1000+040.0-1000+060.0", **{HEADERS[6]: None, "检测方向": "下行"})]
        scanned = scan_rows(rows)
        self.assertEqual(scanned["issues"], [])
        got = [(r.get("source_row"), r.get("source_range"), r["station_m"], r.get("end_m"),
                r["direction"], r["marking_position"], r["value"], r["target"], r["manager"])
               for r in scanned["marking"]]
        expected = []
        for row_no, row in enumerate(rows, 2):
            a, b = row["桩号"].split("-")
            start, end = [float(x.split("+")[0]) * 1000 + float(x.split("+")[1]) for x in (a, b)]
            for vi, ti, side in ((5, 7, "左侧标线"), (6, 8, "右侧标线")):
                if row[HEADERS[vi]] is not None:
                    expected.append((row_no, row["桩号"], start, end, row["检测方向"], side,
                                     row[HEADERS[vi]], row[HEADERS[ti]], row["管养单位"]))
        self.assertEqual(got, expected)

    def test_invalid_metadata_and_numbers_are_reported_without_fabrication(self):
        for field in ("地市", "路线编号", "检测方向", "管养单位", "桩号"):
            with self.subTest(field=field):
                result = scan_rows([wide_row(**{field: " "})])
                self.assertFalse(result["marking"])
                self.assertTrue(result["issues"])
                self.assertIn("第2行", "".join(result["issues"]))
        for column in (HEADERS[5], HEADERS[7]):
            for value in ("NaN", "inf", "-inf", -1, "not-a-number"):
                with self.subTest(column=column, value=value):
                    result = scan_rows([wide_row(**{column: value})])
                    self.assertEqual([r["marking_position"] for r in result["marking"]], ["右侧标线"])
                    self.assertTrue(result["issues"])
        result = scan_rows([wide_row(**{HEADERS[7]: None})])
        self.assertEqual([r["target"] for r in result["marking"]], [80])
        self.assertTrue(result["issues"])
        for value in (float("nan"), float("inf"), -float("inf")):
            self.assertIsNone(engine._float(value))

    def test_city_labels_differing_only_by_shi_suffix_merge_into_one_bundle(self):
        rows = [wide_row(), wide_row(**{"地市": "云浮市", "路线编号": "G2518"})]
        scanned = scan_rows(rows)
        route_index = engine.RouteCategoryIndex({("云浮", "G234"): "普通国省道", ("云浮市", "G2518"): "高速公路"})
        bundles = engine.GuangdongBatchRunner.build_bundles(scanned, route_index)
        self.assertEqual(list(bundles), ["云浮市"])
        self.assertEqual(sorted({r["route"] for r in bundles["云浮市"]["marking"]}), ["G234", "G2518"])
        self.assertTrue(all(r["city"] == "云浮市" for r in bundles["云浮市"]["marking"]))

    def test_legacy_single_value_keeps_segment_and_default_target(self):
        scanner = engine.GuangdongInputScanner(".", engine.RouteCategoryIndex({("云浮", "G234"): "普通国省道"}))
        row = {"计算区间": "K1+000～K1+020", "逆反亮度系数": 88}
        record = scanner._convert("marking", row, Path("G234下行标线3K1K2.csv"), "CSV")
        self.assertEqual((record["city"], record["direction"], record["marking_position"], record["target"]),
                         ("云浮", "下行", "标线3", 80))
        self.assertEqual(record["segment"], "G234下行K1-K2")  # T2k：桩号区间连接符统一 `-`
        self.assertEqual(record["source_range"], row["计算区间"])
        bundle = engine.GuangdongBatchRunner.build_bundles({"marking": [record]}, scanner.route_index)["云浮市"]
        self.assertEqual(bundle["marking"][0]["segment"], record["segment"])
        row["逆反亮度系数目标值"] = -1
        with self.assertRaises(ValueError):
            scanner._convert("marking", row, Path("G234下行标线3K1K2.csv"), "CSV")

    def test_only_connected_ranges_share_report_segment(self):
        rows = [wide_row(), wide_row("1000+020.0-1000+040.0", **{"计算区间": 10}),
                wide_row("1000+035.0-1000+060.0"), wide_row("1000+080.0-1000+100.0"),
                wide_row(**{"管养单位": "乙管养单位"}), wide_row(**{"路线编号": "S265"}),
                wide_row("1000+020.0-1000+000.0", **{"检测方向": "下行"}),
                wide_row("1000+040.0-1000+020.0", **{"检测方向": "下行"}),
                wide_row(**{"地市": "肇庆"})]
        scanned = scan_rows(rows)
        route_index = engine.RouteCategoryIndex({(r["地市"], r["路线编号"]): "普通国省道" for r in rows})
        bundles = engine.GuangdongBatchRunner.build_bundles(scanned, route_index)
        marking = bundles["云浮市"]["marking"]
        groups = engine.GuangdongStatistics.marking_segment_pair_summary(marking)
        self.assertEqual(len(groups), 5)
        self.assertEqual(len(engine.GuangdongStatistics.marking_segment_summary(marking)), 10)
        first = [r for r in marking if r["source_row"] in (2, 3, 4)]
        self.assertEqual(len({r["segment"] for r in first}), 1)
        gap = next(r for r in marking if r["source_row"] == 5)
        self.assertNotEqual(first[0]["segment"], gap["segment"])
        self.assertEqual({r["source_range"] for r in first}, {r["桩号"] for r in rows[:3]})
        down = [r for r in marking if r["direction"] == "下行"]
        self.assertEqual(len({r["segment"] for r in down}), 1)
        self.assertEqual(down[0]["station_m"], 1000020)
        self.assertEqual(down[0]["end_m"], 1000000)
        self.assertEqual(len(marking), 16)
        all_rows = marking + bundles["肇庆市"]["marking"]
        self.assertEqual(len(engine.GuangdongStatistics.marking_segment_pair_summary(all_rows)), 6)
        self.assertEqual(len(scanned["marking"]), len(all_rows))

    def test_explicit_left_right_names_never_swap(self):
        self.assertEqual(engine.marking_side_names(["右侧标线", "左侧标线"]),
                         {"左侧标线": "左侧标线", "右侧标线": "右侧标线"})
        self.assertEqual(engine.marking_side_names(["右侧标线"]), {"右侧标线": "右侧标线"})
        self.assertEqual(engine.marking_side_names(["标线2", "标线3"]),
                         {"标线2": "左侧标线", "标线3": "右侧标线"})
        self.assertEqual(engine.marking_side_names(["标线1", "标线2"]),
                         {"标线1": "左侧标线", "标线2": "右侧标线"})

    def test_word_charts_use_gd03_charts_and_excel_export_keeps_groups_isolated(self):
        rows = []
        for changes in ({}, {"管养单位": "乙管养单位"}, {"路线编号": "S265"}, {"检测方向": "下行"}):
            rows.extend([wide_row(**changes), wide_row("1000+020.0-1000+040.0", **changes)])
        rows.append(wide_row("1000+080.0-1000+100.0", **{HEADERS[5]: None}))
        scanned = scan_rows(rows)
        route_index = engine.RouteCategoryIndex({("云浮", r["路线编号"]): "普通国省道" for r in rows})
        bundle = engine.GuangdongBatchRunner.build_bundles(scanned, route_index)["云浮市"]
        with TemporaryDirectory() as tmp:
            images = engine.guangdong_report_images(bundle, tmp)
            self.assertEqual(len(images), 9)  # 4 identities x 2 sides + right-only gap（Excel 逐区段图不变）
            out = engine.GuangdongChapterWriter.write("云浮", bundle, tmp, TEMPLATE, {"marking": 5, "height": 5, "bolt": 5})
            doc = Document(out)
            chart_dir = Path(tmp) / "云浮" / "charts"
            # P4-Chart-Native：统计图不再出 PNG 嵌 Word，改 Word 原生图表。
            # 这里仍要求 docx 内不含任何 GD03 统计图 PNG（照片/骨架图不受影响）。
            self.assertEqual(list(chart_dir.glob("*gd03_*.png")), [],
                             "GD03 统计图不应再生成 PNG（已改为 Word 原生图表）")
            for shape in doc.inline_shapes:
                self.assertIsNone(shape._inline.graphic.graphicData.pic,
                                 "docx 内不应再有 PNG 嵌图（统计图已改为原生图表）")
            table = next(t for t in doc.tables if "起止桩号" in [c.text for c in t.rows[0].cells])
            headers = [c.text for c in table.rows[0].cells]
            self.assertIn("管养单位", headers)
            self.assertLess(headers.index("左侧合格率(%)"), headers.index("右侧合格率(%)"))
            self.assertEqual(len(table.rows), 4)  # 表头 + 3 个“路线—管养单位”100m 单元
            for cells in table.rows[1:]:
                record = dict(zip(headers, (c.text for c in cells.cells)))
                self.assertNotIn("None", [c.text for c in cells.cells])
                self.assertEqual(record["左侧合格率(%)"], "0.0")
                self.assertEqual(record["右侧合格率(%)"], "100.0")
                self.assertEqual(record["总体合格率(%)"], "50.0")
            wb = openpyxl.load_workbook(engine.write_guangdong_chart_workbook(bundle, tmp))
            ws = wb["标线逆反射"]
            self.assertEqual(len(ws._charts), 9)
            exported = []
            header = None
            for values in ws.iter_rows(values_only=True):
                if values[0] == "桩号":
                    header = list(values)
                elif values[0]:
                    exported.append(dict(zip(header, values)))
            def key(r):
                return (r["路线"], r["方向"], r["管养单位"], r["标线位置"], r["源行号"],
                        r["源桩号范围"], r["逆反射亮度系数"], r["目标值"])
            expected = [(r["route"], r["direction"], r["manager"], r["marking_position"], r["source_row"],
                         r["source_range"], r["value"], r["target"]) for r in bundle["marking"]]
            self.assertEqual(Counter(map(key, exported)), Counter(expected))
            # 每张图的实测/目标系列引用各自数据块，且没有互相覆盖。
            from openpyxl.utils.cell import range_boundaries
            count = 0
            for chart in ws._charts:
                self.assertEqual(len(chart.series), 2)
                ref = chart.series[0].val.numRef.f.split("!")[1]
                x1, y1, x2, y2 = range_boundaries(ref)
                self.assertEqual(x1, x2)
                points = [ws.cell(y, x1).value for y in range(y1, y2 + 1)]
                self.assertTrue(all(isinstance(v, (float, int)) for v in points))
                count += len(points)
            self.assertEqual(count, len(bundle["marking"]))
            wb.close()

    def test_continuous_marking_weak_uses_cell_end_and_length(self):
        def cell(offset):
            return {"route": "G999", "direction": "上行", "marking_position": "左侧标线",
                    "station_m": 1000.0 + 20 * offset, "end_m": 1020.0 + 20 * offset,
                    "value": 10.0, "target": 80.0}

        rows = [cell(i) for i in range(6)]  # 6 个 20m 单元，连续区间恰为 120m
        weak = engine.GuangdongStatistics.continuous_marking_weak(rows, minimum_length_m=120)
        self.assertEqual(weak, [{"route": "G999", "direction": "上行", "marking_position": "左侧标线",
                                 "start_m": 1000.0, "end_m": 1120.0}])
        # 断点拆段：两段各 40m，均不足阈值
        broken = [cell(0), cell(1), dict(cell(2), value=90.0), cell(3), cell(4)]
        self.assertEqual(engine.GuangdongStatistics.continuous_marking_weak(broken, minimum_length_m=120), [])
        # 下行逆向桩号：起点取区间低端、止点取区间高端
        reverse = [{"route": "G998", "direction": "下行", "marking_position": "右侧标线",
                    "station_m": 2000.0 - 20 * i, "end_m": 2000.0 - 20 * i - 20.0,
                    "value": 10.0, "target": 80.0} for i in range(6)]
        weak = engine.GuangdongStatistics.continuous_marking_weak(reverse, minimum_length_m=120)
        self.assertEqual([(w["start_m"], w["end_m"]) for w in weak], [(1880.0, 2000.0)])

    def test_bolt_over5_runs_uses_per_km_keys(self):
        def bolt(km, splice_missing, connection_missing):
            return {"route": "G999", "direction": "上行", "manager": "甲管养单位",
                    "station_m": km * 1000.0, "splice": 100.0, "splice_missing": float(splice_missing),
                    "connection": 100.0, "connection_missing": float(connection_missing), "outline": 0.0}

        rows = [bolt(1000, 10, 5), bolt(1001, 8, 4), bolt(1002, 0, 0)]
        runs = engine.GuangdongStatistics.bolt_over5_runs(rows)
        self.assertEqual(len(runs), 1)
        run = runs[0]
        self.assertEqual((run["start_m"], run["end_m"], run["length_km"]), (1000000.0, 1002000.0, 2))
        self.assertEqual(run["missing"], 27.0)
        self.assertAlmostEqual(run["rate"], 27.0 / (215.0 + 212.0), places=12)

    def test_height_over10_runs_buckets_per_kind(self):
        """R-02：同公里内两波与三波分别判定，不得因混合稀释而漏报（S51 K37 口径）。"""
        rows = [{"route": "S51", "direction": "下行", "manager": "甲管养单位", "station_m": 37000.0 + index,
                 "guardrail_type": "二波", "height": 600.0} for index in range(131)]
        rows += [{"route": "S51", "direction": "下行", "manager": "甲管养单位", "station_m": 37000.0 + index,
                  "guardrail_type": "三波", "height": 850.0 if index < 5 else 700.0} for index in range(20)]
        runs = engine.GuangdongStatistics.height_over10_runs(rows)
        self.assertEqual(len(runs), 1)
        run = runs[0]
        self.assertEqual((run["route"], run["direction"], run["kind"]), ("S51", "下行", "三波"))
        self.assertEqual((run["count"], run["over_count"]), (20, 5))
        self.assertAlmostEqual(run["over_ratio"], 0.25)
        self.assertEqual((run["start_m"], run["end_m"]), (37000.0, 38000.0))

    def test_km_runs_splits_at_missing_km_not_across_the_gap(self):
        """P1a：km 不相邻即断开，不把无检测数据的公里拼成假连续段（brief-v3 §5）。"""
        items = [{"km": km, "overall_rate": 0.1} for km in (10, 11, 12, 15, 16, 20)]
        runs = engine.GuangdongStatistics._marking_runs(items, threshold=0.8)
        self.assertEqual([[i["km"] for i in run] for run in runs], [[10, 11, 12], [15, 16], [20]])


class TypicalPerKmTests(unittest.TestCase):
    """P1a：③典型状况不佳路段统计表逐公里列明（brief-v3 §5 用户修正）。"""

    @staticmethod
    def _marking(km, left_ok, right_ok, direction="下行", route="G94", manager="甲管养单位"):
        """一个公里：每侧 10 组 × 5 条 20 m 记录 = 50 个点位；left_ok/right_ok 为该侧合格组数。

        P4f2-A 标线改**点级**口径后，点数 = 组数 × 5，合格率 = 5·ok/50 = ok/10（与旧单元口径同值），
        所以本组用例断言的合格率全部不变，只有计数字段名 `unit_count` → `point_count`（20 → 100）。
        """
        rows = []
        for side, ok in (("左侧标线", left_ok), ("右侧标线", right_ok)):
            for index in range(10):
                value = 100.0 if index < ok else 10.0
                for record in range(5):
                    station = km * 1000.0 + index * 100.0 + record * 20.0
                    rows.append({"route": route, "direction": direction, "manager": manager,
                                 "station_m": station, "end_m": station + 20.0,
                                 "marking_position": side, "value": value, "target": 80.0})
        return rows

    @staticmethod
    def _height(km, qualified, total, gtype="二波", direction="下行", route="S32", manager="甲管养单位"):
        """一个公里、一个波形梁类型的一组检测点；合格点取该类型的设计高度。"""
        good, bad = engine.HEIGHT_DESIGN[gtype], engine.HEIGHT_DESIGN[gtype] - 200.0
        return [{"route": route, "direction": direction, "manager": manager,
                 "station_m": km * 1000.0 + index, "guardrail_type": gtype,
                 "height": good if index < qualified else bad}
                for index in range(total)]

    def test_zhuhai_g94_ten_km_rows_not_one_merged_row(self):
        """验收：珠海 G94 K407+000-K417+000 在标线③表恰 10 行 K407-K408 … K416-K417。"""
        rows = [row for km in range(407, 417)
                for row in self._marking(km, left_ok=(km - 407) % 5, right_ok=(km - 406) % 5)]
        typical = engine.GuangdongStatistics.marking_typical_segments(rows, limit=3)
        self.assertEqual(len(typical), 1)
        item = typical[0]
        self.assertEqual((item["start_m"], item["end_m"], item["length_km"]), (407000.0, 417000.0, 10))
        self.assertEqual([km["km"] for km in item["km_items"]], list(range(407, 417)))
        # ③表每公里一行，桩号逐公里、总体/左侧/右侧均为该公里原始单元复算值
        cls = engine.GuangdongChapterWriter
        table_rows = [[item["route"], cls._g03_km_manager(rows, item, km["km"]), cls._g03_km_range(km["km"]),
                       cls._g03_num(km["overall_rate"]), cls._g03_num(km["left_rate"]),
                       cls._g03_num(km["right_rate"])]
                      for km in item["km_items"]]
        self.assertEqual(len(table_rows), 10)
        self.assertEqual([r[2] for r in table_rows],
                         [f"K{km}-K{km + 1}" for km in range(407, 417)])
        for km, row in zip(item["km_items"], table_rows):
            self.assertEqual(row[0], "G94")
            self.assertEqual(row[1], "甲管养单位")
            # 该公里左侧 (km-407)%5 个单元合格、右侧 (km-406)%5 个，共 20 个单元
            left, right = (km["km"] - 407) % 5, (km["km"] - 406) % 5
            self.assertEqual(km["point_count"], 100)
            self.assertAlmostEqual(km["left_rate"], left / 10)
            self.assertAlmostEqual(km["right_rate"], right / 10)
            # 总体 =（左侧率 + 右侧率）/2，两侧先按 1 位小数取整（引擎 _gd03_overall 口径）
            overall = (round(left / 10 * 100, 1) + round(right / 10 * 100, 1)) / 200.0
            self.assertAlmostEqual(km["overall_rate"], overall)
            self.assertEqual(row[3], f"{overall * 100:.1f}")
            self.assertEqual(row[4], f"{left / 10 * 100:.1f}")
            self.assertEqual(row[5], f"{right / 10 * 100:.1f}")
        # 不是段平均填充：段平均 20.0，而 10 行是逐公里各自的复算值（5/15/25/35 循环）
        self.assertEqual(cls._g03_num(item["average"]), "20.0")
        self.assertEqual([r[3] for r in table_rows],
                         ["5.0", "15.0", "25.0", "35.0", "20.0"] * 2)
        self.assertGreater(len({r[3] for r in table_rows}), 1)

    def test_reverse_direction_and_one_km_are_handled(self):
        """逆桩号（下行）与单公里段：每段行数 = 有检测数据的 km_items 条数。"""
        rows = []
        for km in (30, 29, 28):  # 下行原始顺序递减
            rows += self._marking(km, 1, 1, direction="下行", route="G999")
        typical = engine.GuangdongStatistics.marking_typical_segments(rows, limit=3)
        self.assertEqual([[km["km"] for km in item["km_items"]] for item in typical], [[28, 29, 30]])
        self.assertEqual((typical[0]["start_m"], typical[0]["end_m"]), (28000.0, 31000.0))
        rows += self._marking(10, 0, 0, direction="上行", route="G999")
        typical = engine.GuangdongStatistics.marking_typical_segments(rows, limit=3)
        one = [item for item in typical if item["direction"] == "上行"][0]
        self.assertEqual([km["km"] for km in one["km_items"]], [10])
        self.assertEqual((one["start_m"], one["end_m"], one["length_km"]), (10000.0, 11000.0, 1))

    def test_missing_side_prints_dash_not_segment_average(self):
        """缺失分侧：某公里只有左侧数据 → 右侧 `—`，不拿段平均或另一侧回填。"""
        rows = [row for km in (5, 6) for row in self._marking(km, 5, 5)]
        rows = [row for row in rows if not (int(row["station_m"] // 1000) == 6
                                            and row["marking_position"] == "右侧标线")]
        typical = engine.GuangdongStatistics.marking_typical_segments(rows, limit=3)
        item = typical[0]
        self.assertEqual([km["km"] for km in item["km_items"]], [5, 6])
        cls = engine.GuangdongChapterWriter
        self.assertEqual(cls._g03_num(item["km_items"][1]["right_rate"]), "—")
        self.assertEqual(cls._g03_num(item["km_items"][0]["right_rate"]), "50.0")
        # 该公里只有左侧单元 → 总体按左侧实际值算（0.5），右侧列 `—`
        self.assertAlmostEqual(item["km_items"][1]["overall_rate"], 0.5)
        self.assertEqual(cls._g03_num(item["km_items"][1]["overall_rate"]), "50.0")
        self.assertEqual(cls._g03_num(item["km_items"][1]["right_rate"]), "—")

    def test_height_typical_per_km_two_and_three_wave_recomputed(self):
        """高度③表逐公里：两波/三波按该公里 raw 明细分别重算，不复制段平均。"""
        rows = []
        for km in (40, 41, 42):
            rows += self._height(km, qualified=2, total=10, gtype="二波")
            rows += self._height(km, qualified=1 if km == 40 else 5, total=10, gtype="三波")
        typical = engine.GuangdongStatistics.height_typical_segments(rows, limit=3)
        item = typical[0]
        self.assertEqual([km["km"] for km in item["km_items"]], [40, 41, 42])
        cls = engine.GuangdongChapterWriter
        rendered = [[cls._g03_km_range(km["km"]), cls._g03_cell(km["rate"]),
                     *[cls._g03_cell(rate) for rate in cls._g03_km_kinds(rows, item, km["km"])]]
                    for km in item["km_items"]]
        self.assertEqual(len(rendered), 3)
        # km40: 二波 2/10、三波 1/10 → 总体 3/20；km41/42: 二波 2/10、三波 5/10 → 总体 7/20
        self.assertEqual(rendered[0], ["K40-K41", "15.00", "20.00", "10.00"])
        self.assertEqual(rendered[1], ["K41-K42", "35.00", "20.00", "50.00"])
        self.assertEqual(rendered[2], ["K42-K43", "35.00", "20.00", "50.00"])
        # 逐公里三波率确实不同（10% vs 50%）→ 段平均无法冒充任一公里
        segment_two = engine.gd_height_segment_kinds(rows, item)[0]
        segment_three = engine.gd_height_segment_kinds(rows, item)[1]
        self.assertEqual(cls._g03_cell(segment_two), "20.00")
        self.assertEqual(cls._g03_cell(segment_three), "36.67")  # 11/30，段级合并
        self.assertNotIn(cls._g03_cell(segment_three), {r[3] for r in rendered})
        # 只有二波的公里 → 三波 `/`
        rows2 = [row for km in (50, 51) for row in self._height(km, 1, 10, gtype="二波", route="S9")]
        item2 = engine.GuangdongStatistics.height_typical_segments(rows2, limit=3)[0]
        self.assertEqual(cls._g03_cell(cls._g03_km_kinds(rows2, item2, 50)[1]), "/")

    def test_km_gap_is_not_bridged_into_one_fake_continuous_segment(self):
        """有缺口时不跨缺口假装连续：返回段的 km 必连续，③表行数 = 该段 km_items 条数。"""
        rows = [row for km in (60, 61, 62, 66, 67)  # 63-65 无检测数据
                for row in self._marking(km, 1, 1, route="G777")]
        typical = engine.GuangdongStatistics.marking_typical_segments(rows, limit=3)
        self.assertTrue(typical)
        for item in typical:
            kms = [km["km"] for km in item["km_items"]]
            # km 连续（不跨 63-65 缺口），且段跨度 = 首末公里
            self.assertEqual(kms, list(range(kms[0], kms[-1] + 1)))
            self.assertEqual((item["start_m"], item["end_m"]),
                             (kms[0] * 1000.0, (kms[-1] + 1) * 1000.0))
            # ③表行数 = 有检测数据的 km_items 条数（无数据的公里不补行、不填均值）
            cls = engine.GuangdongChapterWriter
            table_rows = [cls._g03_km_range(km["km"]) for km in item["km_items"]]
            self.assertEqual(len(table_rows), len(kms))
        self.assertNotIn([60, 61, 62, 66, 67], [[km["km"] for km in i["km_items"]] for i in typical])

    def test_bolt_over5_list_is_per_km_but_severe_clusters_stay_physical(self):
        """螺栓③：整公里清单逐公里拆；严重缺失簇（物理区段）不强拆。"""
        def bolt(km, splice_missing, connection_missing, station=None):
            return {"route": "G888", "direction": "下行", "manager": "甲管养单位",
                    "station_m": km * 1000.0 + 10 if station is None else station,
                    "splice": 100.0, "splice_missing": float(splice_missing),
                    "connection": 100.0, "connection_missing": float(connection_missing), "outline": 0.0}

        # 三个公里缺失强度不同 → 段合并率与逐公里率必然不同
        rows = [bolt(90, 20, 10), bolt(91, 5, 2), bolt(92, 0, 0)]
        runs = engine.GuangdongStatistics.bolt_over5_runs(rows)
        # km92 缺失率 0 不入清单 → 清单只剩 km90-91
        self.assertEqual([km["km"] for km in runs[0]["km_items"]], [90, 91])
        cls = engine.GuangdongChapterWriter
        table_rows = [[cls._g03_km_range(km["km"]), cls._g03_num(km["total_rate"], 2),
                       cls._g03_num(km["splice_rate"], 2), cls._g03_num(km["conn_rate"], 2)]
                      for km in runs[0]["km_items"]]
        self.assertEqual([r[0] for r in table_rows], ["K90-K91", "K91-K92"])
        # 逐公里值取该公里 raw 计数：km90 30/230、km91 7/207，均≠段合并率 37/437
        self.assertEqual(table_rows[0], ["K90-K91", "13.04", "16.67", "9.09"])
        self.assertEqual(table_rows[1], ["K91-K92", "3.38", "4.76", "1.96"])
        self.assertNotIn(f"{runs[0]['rate'] * 100:.2f}", {r[1] for r in table_rows})
        # 严重缺失簇表仍按物理区段（K154+070-K154+090 型）整段列示：缺失率 > 50%
        clustered = [bolt(154, 200, 100, station=154070.0 + 5 * i) for i in range(5)]
        clusters = engine.GuangdongStatistics.bolt_severe_clusters(clustered)
        self.assertEqual(len(clusters), 1)
        self.assertEqual((clusters[0]["start_m"], clusters[0]["end_m"]), (154070.0, 154090.0))
        self.assertEqual((clusters[0]["km"], clusters[0]["points"]), (154, 5))


if __name__ == "__main__":
    unittest.main()
