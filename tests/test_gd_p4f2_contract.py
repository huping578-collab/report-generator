# -*- coding: utf-8 -*-
"""P4f2 契约测试：review-p4 P2-3 要求的 4 条数据契约 + 原生图表载体不变。

* ``P4f2DataContractTests``  —— ①表闭合 / 清单表整桩号 / ②表分行 / 4 条删除文案 0 残留。
* ``P4f2MarkingPointTests`` —— 标线点级口径、区间点数比、整公里长连续窗。
* ``P4f2NativeChartTests``  —— 每张统计图仍是原生图表（部件数 == 图题数、内嵌 xlsx 相等、a:blip 锚点 0）。

跑法：`.venv/Scripts/python.exe -m pytest tests/test_gd_p4f2_contract.py -q`
数据源是**合成的最小 bundle**（不依赖真实数据目录），因此 CI 可跑。
"""
import math
import re
import sys
import unittest
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

from docx import Document
from lxml import etree  # noqa: E402  (Pyright 对 C 扩展子模块报假阳性，运行时正常)

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from backend import report_engine as engine                                    # noqa: E402
from backend.report_engine import GuangdongChapterWriter, GuangdongStatistics  # noqa: E402

TEMPLATE = Path(__file__).resolve().parents[1] / "templates" / "广东项目第五章模板.md"
BANNED = ("未落段", "共识别整公里偏差超过10 cm", "明细见交安设施统计图表工作簿", "100米计算单元")
STATION_M = re.compile(r"^K\d+\+\d+")
WHOLE_STATION = re.compile(r"^K(\d+)-K(\d+)$")
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"


def _bundle() -> dict:
    """最小可渲染 bundle：两条高速明细（其中一条桩号落在路线表区间外 → 未落段）+
    一条普通国省道明细；route_segments 模拟项目路线表（`inspection` 的一级源）。"""
    marking = [
        # S14 上行 K1-K11（落在路线表区间内），左侧 3 点合格、右侧 2 点合格
        {"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K11+000",
         "marking_position": "左侧标线", "value": value, "target": 80.0, "city": "测试市",
         "manager": "甲管养单位", "station_m": station, "end_m": station + 20.0}
        for station, value in ((1000.0, 90.0), (1020.0, 90.0), (1040.0, 90.0), (1060.0, 10.0))
    ] + [
        {"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K11+000",
         "marking_position": "右侧标线", "value": value, "target": 80.0, "city": "测试市",
         "manager": "甲管养单位", "station_m": station, "end_m": station + 20.0}
        for station, value in ((1100.0, 90.0), (1120.0, 90.0), (1140.0, 10.0), (1160.0, 10.0))
    ] + [
        # K90-K95 落在路线表区间**之外** → manager 被覆写成内部哨兵（未落段）
        {"category": "高速公路", "route": "S99", "direction": "下行", "segment": "K90+000～K95+000",
         "marking_position": "左侧标线", "value": 10.0, "target": 80.0, "city": "测试市",
         "manager": "乙管养单位", "station_m": 90000.0, "end_m": 90020.0},
        # 普通国省道
        {"category": "普通国省道", "route": "S120", "direction": "下行", "segment": "K6+000～K6+001",
         "marking_position": "右侧标线", "value": 40.0, "target": 80.0, "city": "测试市",
         "manager": "东莞市公路事务中心", "station_m": 6000.0, "end_m": 6001.0},
    ]
    height = [
        {"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K11+000",
         "guardrail_type": "二波", "height": 600.0, "city": "测试市", "manager": "甲管养单位",
         "station_m": 1000.0, "end_m": 1001.0},
        {"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K11+000",
         "guardrail_type": "二波", "height": 700.0, "city": "测试市", "manager": "甲管养单位",
         "station_m": 2000.0, "end_m": 2001.0},
        {"category": "高速公路", "route": "S99", "direction": "下行", "segment": "K90+000～K95+000",
         "guardrail_type": "三波", "height": 500.0, "city": "测试市", "manager": "乙管养单位",
         "station_m": 90000.0, "end_m": 90001.0},
        {"category": "普通国省道", "route": "S120", "direction": "下行", "segment": "K6+000～K6+001",
         "guardrail_type": "二波", "height": 600.0, "city": "测试市", "manager": "东莞市公路事务中心",
         "station_m": 6000.0, "end_m": 6001.0},
    ]
    bolt = [
        {"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K11+000",
         "bolt_type": "拼接螺栓", "missing": 2, "expected": 100, "city": "测试市",
         "manager": "甲管养单位", "station_m": 1000.0, "end_m": 1001.0},
        {"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K11+000",
         "bolt_type": "拼接螺栓", "missing": 0, "expected": 100, "city": "测试市",
         "manager": "甲管养单位", "station_m": 2000.0, "end_m": 2001.0},
        {"category": "高速公路", "route": "S99", "direction": "下行", "segment": "K90+000～K95+000",
         "bolt_type": "拼接螺栓", "missing": 3, "expected": 100, "city": "测试市",
         "manager": "乙管养单位", "station_m": 90000.0, "end_m": 90001.0},
        {"category": "普通国省道", "route": "S120", "direction": "下行", "segment": "K6+000～K6+001",
         "bolt_type": "拼接螺栓", "missing": 1, "expected": 100, "city": "测试市",
         "manager": "东莞市公路事务中心", "station_m": 6000.0, "end_m": 6001.0},
    ]
    return {"city": "测试市", "marking": marking, "height": height, "bolt": bolt,
            "notes": [], "comparison_detail": [], "weak_segments": [],
            "route_segments": [
                # S14 上行：里程 10 km，非省交通集团；起止桩号非整数 → 清单表须 floor/ceil
                {"city": "测试", "route": "S14", "direction": "上行", "category": "高速公路",
                 "start": 0.986, "end": 10.004, "length": 9.018, "manager": "甲管养单位",
                 "owner": engine.GD_OWNER_NON_PROVINCIAL, "route_name": "测试高速"},
                # S120 下行：普通国省道 1 km
                {"city": "测试", "route": "S120", "direction": "下行", "category": "普通国省道",
                 "start": 6.0, "end": 7.0, "length": 1.0, "manager": "东莞市公路事务中心",
                 "owner": None, "route_name": "测试省道"},
            ]}


class _DocFixture(unittest.TestCase):
    """整份报告只生成一次（一次 ≈ 数秒），子类共用。"""

    @classmethod
    def setUpClass(cls) -> None:
        cls._tmp = TemporaryDirectory()
        out = Path(cls._tmp.name) / "out"
        cls.path = engine.GuangdongChapterWriter.write(
            "测试市", _bundle(), out, TEMPLATE, {"marking": 5, "height": 5, "bolt": 5})
        cls.doc = Document(cls.path)
        cls.text = "\n".join(p.text for p in cls.doc.paragraphs)
        cls.table_text = "\n".join(c.text for t in cls.doc.tables for r in t.rows for c in r.cells)

    @classmethod
    def tearDownClass(cls) -> None:
        cls._tmp.cleanup()

    def tables_with_head(self, *needles):
        out = []
        for t in self.doc.tables:
            head = [c.text.strip() for c in t.rows[0].cells]
            if all(n in head for n in needles):
                out.append(t)
        return out


class P4f2DataContractTests(_DocFixture):
    """review-p4 P2-3：①表闭合 / 清单表整桩号 / ②表分行 / 删除文案 0 残留。"""

    def test_city_table_highway_km_equals_route_table_segment_sum(self) -> None:
        """契约①：①表「全市抽检高速」里程 == 路线表该市该类逐段求和。"""
        table = self.tables_with_head("管理单位", "抽检里程(km)")[0]
        rows = {r.cells[0].text.strip(): r.cells[1].text.strip() for r in table.rows[1:]}
        segments = [s for s in _bundle()["route_segments"] if s["category"] == "高速公路"]
        expected = sum(float(s["length"]) for s in segments)
        self.assertAlmostEqual(float(rows["测试市抽检高速"]), expected, delta=0.5)
        # 非省 + 省 == 全市（P1j 最大余数法，三值整数恒等）
        self.assertEqual(int(float(rows["非省交通集团"])) + int(float(rows["省交通集团"])),
                         int(float(rows["测试市抽检高速"])))
        # 全省行 4 格留空
        blank = next(r for r in table.rows[1:] if r.cells[0].text.strip() == "全省抽检高速")
        self.assertEqual([c.text.strip() for c in blank.cells[1:]], ["", "", "", ""])

    def test_inspection_list_uses_floor_ceil_whole_station(self) -> None:
        """契约②：清单表逐行起止桩号 = 起点 floor、终点 ceil（0.986/10.004 → K0-K11）。"""
        table = self.tables_with_head("类型", "路线编号", "起点桩号", "终点桩号")[0]
        checked = 0
        for row in table.rows[1:]:
            cells = [c.text.strip() for c in row.cells]
            start, end = cells[4], cells[5]
            if start == "—":
                continue
            low, high = float(start), float(end)
            self.assertEqual(cells[4], f"{math.floor(low):g}",
                             f"起点未 floor：{cells}")
            self.assertEqual(cells[5], f"{math.ceil(high):g}",
                             f"终点未 ceil：{cells}")
            checked += 1
        self.assertGreater(checked, 0, "清单表没有可比对的行")

    def test_route_table_splits_by_manager_per_route_and_direction(self) -> None:
        """契约③：同路线同方向不同管养单位逐组分行（无合并成一行）。"""
        rows = [{"route": "G1", "direction": "上行", "manager": "甲", "station_m": 1000.0,
                 "value": 90.0, "target": 80.0, "marking_position": "左侧标线"},
                {"route": "G1", "direction": "上行", "manager": "乙", "station_m": 3000.0,
                 "value": 10.0, "target": 80.0, "marking_position": "左侧标线"}]
        units = GuangdongStatistics.marking_route_units(rows)
        self.assertEqual([(u["route"], u["manager"]) for u in units],
                         [("G1", "甲"), ("G1", "乙")])
        self.assertEqual([u["left_rate"] for u in units], [1.0, 0.0])

    def test_deleted_text_has_zero_residue(self) -> None:
        """契约④：本轮 4 条删除文案 0 残留（段落 + 表格单元，两处都查）。"""
        for phrase in BANNED:
            with self.subTest(phrase=phrase):
                self.assertEqual(self.text.count(phrase), 0, f"段落残留 {phrase!r}")
                self.assertEqual(self.table_text.count(phrase), 0, f"表格残留 {phrase!r}")

    def test_unassigned_row_is_kept_but_rendered_as_dash(self) -> None:
        """P4f2-D：未落段的行**保留**，管养单位显示为「—」，不泄露内部占位符。"""
        self.assertEqual(engine.manager_display(engine.GD_UNASSIGNED), "—")
        self.assertEqual(engine.GD_UNASSIGNED, "未落段")       # 引擎内部哨兵仍在（分组/过滤要用）
        hits = [row for t in self.doc.tables for row in t.rows
                if any(engine.GD_UNASSIGNED in c.text for c in row.cells)]
        self.assertEqual(hits, [], "产物里仍有未落段字样")


class P4f2RouteTableContractTests(_DocFixture):
    """P4f2-B/E/F：②图合计最左、②表整桩号同源、差额脚注。"""

    def test_route_charts_put_city_total_first(self) -> None:
        """P4f2-B：②图首类是「XX市合计」，值与①表该市行同源（不因换位而变）。

        合计项沿用①表该市行的键名（`left`/`right`/`overall`），逐行项用 `left_rate`/`right_rate`/
        `overall_rate`；取值函数写成 `.get(x_rate, .get(x))` 才能同时吃下这两种键 ——
        与引擎三个调用点逐字一致。
        """
        pick = lambda key: (lambda r: r.get(f"{key}_rate", r.get(key)))      # noqa: E731
        units = [{"route": "G1", "manager": "甲", "left_rate": 0.6, "right_rate": 0.4, "overall_rate": 0.5},
                 {"route": "G2", "manager": "乙", "left_rate": 0.2, "right_rate": 0.1, "overall_rate": 0.15}]
        totals = {"left": 0.55, "right": 0.45, "overall": 0.5}      # ①表该市行
        cats, series = GuangdongChapterWriter._g03_route_chart(
            "测试市", units, [("左侧标线", pick("left"), "#4F81BD"),
                              ("总体", pick("overall"), "#9BBB59")], totals)
        self.assertEqual(cats[0], "测试市合计")
        self.assertEqual(cats[1:], ["G1甲", "G2乙"])
        self.assertEqual(series[0][1], [0.55, 0.6, 0.2])            # 合计在前，逐行在后
        self.assertEqual(series[1][1], [0.5, 0.5, 0.15])
        for _name, values, _color in series:                       # 长度守恒：系列 == 类目
            self.assertEqual(len(values), len(cats))

    def test_route_table_station_uses_inspection_range_not_record_span(self) -> None:
        """P4f2-F：②表起止桩号取清单表区间（floor/ceil），不再用检测记录 min/max。"""
        inspection = [{"category": "高速公路", "route": "G1", "direction": "上行", "manager": "甲",
                       "start": 74.2, "end": 75.9, "length_km": 1.7}]
        got = engine.gd_inspection_range(inspection, category="高速公路", route="G1", manager="甲")
        self.assertEqual(got, (74200.0, 75900.0))
        self.assertEqual(GuangdongChapterWriter._gd03_whole_range(*got), "K74-K76")
        # 无匹配行 → None（不是 (None, None)，否则调用点的 `or` 回退会被真值元组吃掉）
        self.assertIsNone(engine.gd_inspection_range(inspection, route="G999"))
        # 文档里所有 ②/③表 起止桩号都必须是整桩号（无米数）。
        # 「连续缺失严重区段」表带 `跨度(m)` 列 —— 那里米级桩号是**正确**的（物理缺失点位置），排除。
        for table in self.doc.tables:
            head = [c.text.strip() for c in table.rows[0].cells]
            if len(head) > 3 and head[3] in ("起止桩号", "检测范围") and "跨度(m)" not in head:
                for row in table.rows[1:]:
                    self.assertIsNone(STATION_M.match(row.cells[3].text.strip()),
                                      f"②表起止桩号带米数：{row.cells[3].text!r}")

    def test_gap_note_discloses_difference_and_is_silent_when_zero(self) -> None:
        """P4f2-E：有差额逐市披露（全市/已纳入/未纳入 + ±0.5 km 取整差）；差额 0 不出脚注。"""
        inspection = [
            {"category": "高速公路", "route": "G1", "direction": "上行", "manager": "甲",
             "start": 0.0, "end": 10.0, "length_km": 10.0},
            {"category": "高速公路", "route": "G2", "direction": "下行", "manager": "乙",
             "start": 0.0, "end": 5.0, "length_km": 5.0},
        ]
        units = [{"route": "G1", "direction": "上行", "manager": "甲"}]      # G2 未被②表覆盖
        rows = [{"manager": engine.GD_UNASSIGNED}]
        note = engine.gd_route_gap_note(inspection, units, rows, "高速公路")
        self.assertIsNotNone(note)
        assert note is not None                                # 收窄 Optional[str]
        for needle in ("全市抽检高速公路 15 km", "已纳入分段统计 10 km", "另有 5 km",
                       "1 条记录", "差额已计入表5-2 全市值", "±0.5 km 取整差"):
            self.assertIn(needle, note)
        # 覆盖完整 → 无脚注
        full = [{"route": "G1", "direction": "上行", "manager": "甲"},
                {"route": "G2", "direction": "下行", "manager": "乙"}]
        self.assertIsNone(engine.gd_route_gap_note(inspection, full, [], "高速公路"))


class P4f2MarkingPointTests(unittest.TestCase):
    """P4f2-A：标线点级判定、区间点数比、整公里长连续窗；旧 100 m 单元口径已废除。"""

    def _rows(self):
        # 4 点：3 合格 1 不合格 → 点级合格率 0.75（单元均值口径下会是 1.0）
        return [{"route": "G1", "direction": "上行", "manager": "甲", "marking_position": "左侧标线",
                 "value": value, "target": 80.0, "station_m": 1000.0 + 20.0 * i, "end_m": 1020.0 + 20.0 * i}
                for i, value in enumerate((100.0, 100.0, 100.0, 10.0))]

    def test_single_record_level_qualification_not_unit_average(self) -> None:
        points = GuangdongStatistics.marking_points(self._rows())
        self.assertEqual(len(points), 4)                      # 4 条记录 = 4 个点位
        self.assertEqual([p["qualified"] for p in points], [True, True, True, False])
        overall = GuangdongStatistics.marking_overall(points)
        self.assertEqual(overall["point_count"], 4)
        self.assertEqual(overall["left_points"], 4)
        self.assertAlmostEqual(overall["left_rate"], 0.75)    # 3/4，不是单元均值 1.0
        self.assertIsNone(overall["right_rate"])              # 缺右侧不拿左侧回填
        self.assertAlmostEqual(overall["overall_rate"], 0.75)

    def test_rate_denominator_is_actual_points_not_theoretical(self) -> None:
        """分母 = 区间实际点位数：3 点 2 合格 → 2/3，不按 5 条/单元的理论点补齐成 10 点。"""
        rows = self._rows()[:3]                               # 只剩 3 点（原本 4 点）
        rows[2] = dict(rows[2], value=10.0)                  # → 100 / 100 / 10
        rates = GuangdongStatistics.marking_side_rates(GuangdongStatistics.marking_points(rows))
        self.assertEqual(rates["左侧标线"]["point_count"], 3)
        self.assertEqual(rates["左侧标线"]["qualified_count"], 2)
        self.assertAlmostEqual(rates["左侧标线"]["qualified_rate"], 2.0 / 3.0)

    def test_overall_column_is_mean_of_two_one_decimal_sides(self) -> None:
        """甲方模板口径：总体 =（左 + 右）/2，两侧**各自先按 1 位小数取整**（百分比取整后再平均）。

        44.4% 与 55.6% → 50.0%；44.4% 与 55.5% → 49.95%（先各自取整，再相加除 2）。
        """
        self.assertAlmostEqual(engine._gd03_overall(0.444, 0.556), 0.5)
        self.assertAlmostEqual(engine._gd03_overall(0.444, 0.555), (44.4 + 55.5) / 200.0)
        # 缺一侧时取另一侧，不做 (x+0)/2 稀释
        self.assertAlmostEqual(engine._gd03_overall(0.444, None), 0.444)
        self.assertIsNone(engine._gd03_overall(None, None))

    def test_long_runs_use_whole_km_window_and_min_three_km(self) -> None:
        """P4f2-A：长连续窗以**整公里**为单位；连续坏公里 ≥ 3 km 才输出。"""
        rows = []
        for km in range(10, 16):                              # 6 个连续整公里，90% 不合格
            for i in range(10):
                rows.append({"route": "G1", "direction": "上行", "manager": "甲",
                             "marking_position": "左侧标线", "value": 10.0, "target": 80.0,
                             "station_m": km * 1000.0 + i * 100.0, "end_m": km * 1000.0 + i * 100.0 + 20.0})
        for km in range(30, 33):                              # 只有 3 km…刚好够
            rows.append({"route": "G2", "direction": "上行", "manager": "甲",
                         "marking_position": "左侧标线", "value": 10.0, "target": 80.0,
                         "station_m": km * 1000.0, "end_m": km * 1000.0 + 20.0})
        for km in range(50, 52):                              # 只有 2 km → 不输出
            rows.append({"route": "G3", "direction": "上行", "manager": "甲",
                         "marking_position": "左侧标线", "value": 10.0, "target": 80.0,
                         "station_m": km * 1000.0, "end_m": km * 1000.0 + 20.0})
        runs = {r["route"]: r for r in GuangdongStatistics.marking_long_runs(rows)}
        self.assertEqual(runs["G1"]["length_km"], 6)
        self.assertEqual((runs["G1"]["start_m"], runs["G1"]["end_m"]), (10000.0, 16000.0))
        self.assertEqual(runs["G1"]["point_count"], 60)
        self.assertEqual(runs["G2"]["length_km"], 3)
        self.assertNotIn("G3", runs)                          # 2 km < 3 km 门槛

    def test_legacy_100m_unit_api_is_gone(self) -> None:
        """旧口径的 5 个入口已废除/改造，产物里不会再出现「100米计算单元」。"""
        stats = GuangdongStatistics
        for gone in ("marking_units", "marking_unit_rates", "_marking_km", "MARKING_UNIT_RECORDS"):
            self.assertFalse(hasattr(stats, gone), f"{gone} 应已废除")
        for kept in ("marking_points", "marking_side_rates", "marking_overall",
                     "marking_route_units", "marking_per_km", "marking_typical_segments",
                     "marking_long_runs"):
            self.assertTrue(hasattr(stats, kept), f"{kept} 缺失")


class P4f2NativeChartTests(_DocFixture):
    """契约⑤：每张统计图仍是**原生可编辑**图表（回归 P4-Chart-OOXML 的载体结论）。"""

    def test_every_chart_is_native_editable_not_a_picture(self) -> None:
        with zipfile.ZipFile(self.path) as z:
            names = z.namelist()
            charts = [n for n in names if re.fullmatch(r"word/charts/chart\d+\.xml", n)]
            embeds = [n for n in names
                      if n.startswith("word/embeddings/") and n.endswith(".xlsx")]
            xml = etree.fromstring(z.read("word/document.xml"))
        captions = [t for t in xml.iter(f"{W}t")
                    if (t.text or "").strip().startswith("图5-")]
        anchors = xml.findall(f".//{C}chart")
        blips = xml.findall(f".//{A}blip")
        self.assertGreater(len(charts), 0, "没有原生图表部件")
        self.assertEqual(len(charts), len(anchors), "图表部件数与绘图锚点数不一致")
        self.assertEqual(len(charts), len(embeds), "图表数与内嵌工作簿数不一致")
        self.assertEqual(len(charts), len(captions), "图表数与图题数不一致")
        self.assertEqual(blips, [], f"统计图位上还有 {len(blips)} 个位图锚点（应为原生图表）")
        self.assertNotIn("@@GDCHART", etree.tostring(xml, encoding="unicode"),
                         "锚点占位文字残留")
        with zipfile.ZipFile(self.path) as z:                  # 每张图自带缓存 + 数据
            for name in charts:
                root = etree.fromstring(z.read(name))
                self.assertTrue(root.findall(f".//{C}ser"), f"{name} 没有系列")
                self.assertTrue(root.findall(f".//{C}cat"), f"{name} 没有类目缓存")


if __name__ == "__main__":
    unittest.main()
