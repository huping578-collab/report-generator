"""P4-Chart-OOXML：Word 原生图表的**离线**（零 Office）产物级断言。

分层：
* ``OoxmlChartSpecTests``      —— 纯逻辑，不碰 Word。
* ``OoxmlChartDocxTests``      —— 真跑一遍注入，开包核对部件/关系/内嵌工作簿/缓存/表值。
* ``NoOfficeAtGenerationTests``—— 导入钩子 + 源码断言，证明生成路径不拉 Office。
* ``NativeChartWordTests``     —— Word 真机：注入 → 校验 → 改一格 → 重开读回（用户核心诉求）。
"""
import io
import re
import sys
import unittest
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from backend import gd_chart_ooxml as ooxml
from backend import gd_native_chart as native
from backend.gd_native_chart import ChartSpec, build_anchor
from backend.report_engine import GuangdongChapterWriter

WD_INLINE_CHART = 12  # wdInlineShapeChart
BLOCKED = ("win32com", "pythoncom", "pywintypes")

CATS = ["广州市合计", "S41广澳高速东段公司", "G0425广河高速粤水电公司"]
SERIES = [("左侧标线", [55.2, 61.0, 40.5], "#4F81BD"),
          ("右侧标线", [59.9, 48.5, 52.0], "#C0504D"),
          ("缺失率", [None, 3.1, 7.9], "#9BBB59")]


def _word_available() -> bool:
    try:
        _pythoncom, win32 = native._import_com()
    except native.WordChartUnavailable:
        return False
    try:
        app = win32.DispatchEx("Word.Application")
        app.Quit()
        return True
    except Exception:                                    # noqa: BLE001
        return False


WORD_OK = _word_available()
_REASON = "本机无 Microsoft Word 自动化环境，Word 真机用例跳过"


def _build_docx(path: Path, specs, width_cm=15.0, height_cm=7.5):
    doc = Document()
    for index, spec in enumerate(specs, 1):
        spec.anchor = build_anchor(index)
        GuangdongChapterWriter._chart_anchor(doc, spec.anchor)
        doc.add_paragraph(f"图5-{index} 抽检高速标线合格率对比图").style = "Caption"
    ooxml.save_with_charts(doc, path, specs, width_cm=width_cm,
                           height_cm=height_cm, log=lambda *_: None)
    return specs


class OoxmlChartSpecTests(unittest.TestCase):
    """不碰 Word 的纯逻辑部分。"""

    def test_seed_asset_is_packaged(self) -> None:
        """种子必须落在仓库稳定路径，不能指向 Temp（桌面版无 Word 也要能生成）。"""
        seed = ooxml.seed_path()
        self.assertTrue(seed.is_file(), seed)
        self.assertEqual(seed.name, "gd_native_chart_seed.docx")
        self.assertNotIn("Temp", str(seed))
        self.assertEqual(seed.parent.name, "assets")

    def test_number_format_maps_percent_specs(self) -> None:
        self.assertEqual(ChartSpec(["a"], [("s", [1], "#4F81BD")],
                                   value_fmt="%.1f").number_format(), "0.0")
        self.assertEqual(ChartSpec(["a"], [("s", [1], "#4F81BD")],
                                   value_fmt="%.2f").number_format(), "0.00")

    def test_long_or_many_labels_rotate(self) -> None:
        # P4c 口径：①图短标签横排，②图「路线号+管养单位」长标签旋转。
        short = ChartSpec(["全省抽检高速", "广州市抽检高速"], [("s", [1, 2], "#4F81BD")])
        self.assertFalse(short.need_rotate())
        self.assertTrue(ChartSpec(["G0425广河高速粤水电公司"],
                                  [("s", [1], "#4F81BD")]).need_rotate())
        self.assertTrue(ChartSpec([f"类目{i}" for i in range(13)],
                                  [("s", [1] * 13, "#4F81BD")]).need_rotate())

    def test_chart_type_and_anchors(self) -> None:
        bar = ChartSpec(["a"], [("s", [1], "#4F81BD")])
        self.assertEqual(bar.chart_type(), native.XL_COLUMN_CLUSTERED)
        self.assertEqual(ChartSpec(["a"], [("s", [1], "#4F81BD")], kind="line").chart_type(),
                         native.XL_LINE_MARKERS)
        self.assertEqual(ChartSpec(["a"], [("s", [1], "#4F81BD")], horizontal=True).chart_type(),
                         native.XL_BAR_CLUSTERED)
        self.assertEqual(build_anchor(7), "@@GDCHART0007@@")

    def test_bgr_conversion(self) -> None:
        self.assertEqual(native._bgr("#4F81BD"), 0xBD814F)   # Word/Excel 用 BGR
        self.assertEqual(native._bgr("C0504D"), 0x4D50C0)

    def test_anchor_paragraph_matches_picture_layout(self) -> None:
        """锚点段必须与 `_picture` 同版式：居中、段前 12 磅、段后 3 磅。"""
        with TemporaryDirectory() as tmp:
            doc = Document()
            paragraph = GuangdongChapterWriter._chart_anchor(doc, build_anchor(1))
            self.assertEqual(paragraph.text, "@@GDCHART0001@@")
            self.assertEqual(paragraph.alignment, WD_ALIGN_PARAGRAPH.CENTER)
            self.assertEqual(paragraph.paragraph_format.space_before, Pt(12))
            self.assertEqual(paragraph.paragraph_format.space_after, Pt(3))

    def test_chart_xml_carries_no_word_default_residue(self) -> None:
        """回归：Word 建图默认 3 系列 4 类目，残值 1.8/4.5 曾被当数据画出来。
        这里逐点核对缓存里没有 spec 之外的任何 `<c:pt>`。"""
        xml = ooxml.build_chart_xml(ChartSpec(list(CATS), list(SERIES)))
        text = xml.decode("utf-8")
        self.assertIn("Sheet1!$A$2:$A$4", text)          # 公式指向实际区域
        self.assertIn("Sheet1!$D$2:$D$4", text)
        for residue in ("1.8", "4.5", "类别 3", "类别 4", "系列 3"):
            self.assertNotIn(f"<c:v>{residue}</c:v>", text, f"残留 Word 默认值 {residue}")
        # 每系列 1 个系列名 ptCount(1) + 1 个类别 ptCount(3) + 1 个数值 ptCount(3) = 9
        self.assertEqual(text.count('<c:ptCount val="1"/>'), 3, "3 个系列名")
        self.assertEqual(text.count('<c:ptCount val="3"/>'), 6, "3 系列的类别缓存 + 数值缓存")

    def test_series_length_must_match_category_count(self) -> None:
        """点数 ≠ 类目数必须报出可定位的错，而不是下标越界。"""
        bad = ChartSpec(["a", "b", "c"], [("左", [1.0, 2.0], "#4F81BD")])
        bad.anchor = build_anchor(1)
        with self.assertRaises(ooxml.OoxmlChartError) as ctx:
            ooxml.build_chart_xml(bad)
        self.assertIn("2 个点", str(ctx.exception))
        self.assertIn("3 个", str(ctx.exception))

    def test_line_and_horizontal_variants(self) -> None:
        line = ooxml.build_chart_xml(ChartSpec(list(CATS), list(SERIES[:2]), kind="line"))
        self.assertIn(b"<c:lineChart>", line)
        self.assertIn(b'<c:symbol val="circle"/>', line)
        bar = ooxml.build_chart_xml(ChartSpec(list(CATS), list(SERIES[:2]), horizontal=True))
        self.assertIn(b'<c:barDir val="bar"/>', bar)


class OoxmlChartDocxTests(unittest.TestCase):
    """真跑一遍离线注入 + 开包强断言（不需要 Word）。"""

    def test_native_parts_rels_workbook_and_cache_all_agree(self) -> None:
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "第五部分.docx"
            # 注意：不能用 `[spec] * n`——那是同一个对象，anchor 会被后一个覆盖。
            specs = _build_docx(out, [ChartSpec(list(CATS), list(SERIES), ylabel="合格率（%）",
                                                ylim=(0, 100))
                                      for _ in range(2)])
            with zipfile.ZipFile(out) as archive:
                names = archive.namelist()
            charts = [n for n in names if re.fullmatch(r"word/charts/chart\d+\.xml", n)]
            embeds = [n for n in names if n.startswith("word/embeddings/")]
            self.assertEqual(len(charts), 2, charts)
            self.assertEqual(len(embeds), 2, embeds)
            self.assertIn("word/charts/style1.xml", names)
            self.assertIn("word/charts/colors1.xml", names)
            self.assertFalse([n for n in names if n.startswith("word/media/")],
                             "统计图不再以 PNG 入库")
            # 校验器本身就是最强断言（部件/关系/锚点/缓存/工作簿/表值逐图逐点）
            ooxml.verify_docx_charts(out, specs)
            text = "\n".join(p.text for p in Document(out).paragraphs)
            self.assertNotIn("@@GDCHART", text, "锚点文字必须被图表替换掉")

    def test_chart_parts_and_ids_are_unique_per_chart(self) -> None:
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "第五部分.docx"
            # 三张图类目数各不相同（2/3/4），逼出 ptCount 与公式必须跟着变
            specs = _build_docx(out, [
                ChartSpec([f"类目{i}" for i in range(n)],
                          [(name, list(range(n)), color) for name, _v, color in SERIES[:2]])
                for n in (2, 3, 4)])
            with zipfile.ZipFile(out) as z:
                document = z.read("word/document.xml").decode("utf-8")
                rels = z.read("word/_rels/document.xml.rels").decode("utf-8")
            self.assertEqual(document.count("<c:chart "), 3)
            self.assertEqual(len(set(re.findall(r'r:id="(rId\d+)"', document))), 3)
            self.assertEqual(len(set(re.findall(r'<wp:docPr id="(\d+)"', document))), 3)
            all_rids = re.findall(r'Id="(rId\d+)"', rels)
            self.assertEqual(len(set(all_rids)), len(all_rids), "rId 不得重复")
            ooxml.verify_docx_charts(out, specs)

    def test_gap_values_are_blank_in_both_workbook_and_cache(self) -> None:
        """None = 断点：内嵌工作簿留空、``c:numCache`` 不写该 pt（idx 跳号）。"""
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "第五部分.docx"
            _build_docx(out, [ChartSpec(list(CATS), list(SERIES))])
            with zipfile.ZipFile(out) as z:
                chart = z.read("word/charts/chart1.xml").decode("utf-8")
                book = z.read("word/embeddings/gd_chart1.xlsx")
            third = re.findall(r"<c:numCache>.*?</c:numCache>", chart, re.S)[2]
            self.assertIn('<c:pt idx="1"><c:v>3.1</c:v></c:pt>', third)
            self.assertNotIn('<c:pt idx="0"', third)          # 第 0 点是断点，不落 pt
            from openpyxl import load_workbook
            grid = [[c.value for c in row] for row in
                    load_workbook(io.BytesIO(book)).worksheets[0].iter_rows()]
            self.assertIsNone(grid[1][3])                      # 缺失率 × 广州市合计 留空

    def test_failed_injection_keeps_previous_file_intact(self) -> None:
        """原子交付：校验不过就抛，且旧文件一个字节都不动、不留临时文件。"""
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "第五部分.docx"
            good = [ChartSpec(list(CATS), list(SERIES))]
            _build_docx(out, good)
            before = out.read_bytes()
            # 第二个 spec 的锚点在文档里不存在 → 注入必失败
            bad = [good[0], ChartSpec(list(CATS), list(SERIES))]
            bad[1].anchor = build_anchor(99)
            with self.assertRaises(ooxml.OoxmlChartError):
                ooxml.inject_charts(out, bad, log=lambda *_: None)
            self.assertEqual(out.read_bytes(), before, "失败后原文件被改动了")
            self.assertFalse(list(out.parent.glob("*.gdchart*.docx")), "残留临时文件")

    def test_chart_styles_are_per_chart_not_shared(self) -> None:
        """回归（2026-09-28 真机踩到）：多张图表的 rels 若指向**同一个** style/colors 部件，
        Word 打开文档直接报「在试图打开文件时遇到错误」——N=1 能开、N=2 就挂。
        这里锁死「每图一份」，避免以后为了省 10 KB 又改回共享。"""
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "第五部分.docx"
            specs = _build_docx(out, [
                ChartSpec([f"类目{i}" for i in range(4)],
                          [(name, list(range(4)), color) for name, _v, color in SERIES[:2]])
                for _ in range(3)])
            with zipfile.ZipFile(out) as z:
                names = z.namelist()
                for index in (1, 2, 3):
                    self.assertIn(f"word/charts/style{index}.xml", names)
                    self.assertIn(f"word/charts/colors{index}.xml", names)
                    rels = z.read(f"word/charts/_rels/chart{index}.xml.rels").decode()
                    self.assertIn(f'Target="style{index}.xml"', rels)
                    self.assertIn(f'Target="colors{index}.xml"', rels)
            ooxml.verify_docx_charts(out, specs)

    def test_no_external_workbook_reference(self) -> None:
        """不得链接外部临时 xlsx：关系必须是内部 package 目标且无 TargetMode。"""
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "第五部分.docx"
            _build_docx(out, [ChartSpec(list(CATS), list(SERIES))])
            with zipfile.ZipFile(out) as z:
                rels = z.read("word/charts/_rels/chart1.xml.rels").decode("utf-8")
            self.assertIn("../embeddings/gd_chart1.xlsx", rels)
            self.assertNotIn("TargetMode", rels)
            self.assertNotIn("file:///", rels)


class NoOfficeAtGenerationTests(unittest.TestCase):
    """生成路径不得触碰 Office。"""

    def test_import_hook_blocks_office(self) -> None:
        class _Blocker:
            def find_spec(self, name, path=None, target=None):
                if name.split(".")[0] in BLOCKED:
                    raise ImportError(f"blocked {name}")
                return None

        blocker = _Blocker()
        sys.meta_path.insert(0, blocker)
        try:
            with self.assertRaises(ImportError):
                __import__("win32com.client")
        finally:
            sys.meta_path.remove(blocker)

    def test_ooxml_module_does_not_touch_com(self) -> None:
        """`gd_chart_ooxml` 是纯 ZIP/XML 生成器：绝不 import pywin32、不碰 ChartData。"""
        source = Path(ooxml.__file__).read_text(encoding="utf-8")
        for forbidden in ("import win32com", "import pythoncom", "DispatchEx", "ChartData"):
            self.assertNotIn(forbidden, source, f"生成模块里出现 {forbidden}")


@unittest.skipUnless(WORD_OK, _REASON)
class NativeChartWordTests(unittest.TestCase):
    """Word 真机：注入 → 校验 → 改一格 → 重开读回（用户核心诉求）。"""

    def _build(self, path: Path):
        return _build_docx(path, [ChartSpec(list(CATS), list(SERIES), ylabel="合格率（%）",
                                            ylim=(0, 100))])

    @staticmethod
    def _chart_cache(path: Path):
        """直接读 chart XML 的 ``<c:numCache>``——这是图表真正拿去画的数据。
        绕开 COM：COM 回读走运行期缓存，历史上出过「工作簿对、图不对」的假象。"""
        out = []
        with zipfile.ZipFile(path) as archive:
            for name in sorted(n for n in archive.namelist()
                               if re.fullmatch(r"word/charts/chart\d+\.xml", n)):
                xml = archive.read(name).decode("utf-8")
                out.append([[round(float(v), 2) for v in re.findall(r"<c:v>([^<]*)</c:v>", block)]
                            for block in re.findall(r"<c:numCache>.*?</c:numCache>", xml, re.S)])
        return out

    def test_injected_chart_is_native_and_editable_in_word(self) -> None:
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "第五部分.docx"
            self._build(out)
            charts_read = native.read_charts(out, log=lambda *_: None)
            self.assertEqual(len(charts_read), 1)
            self.assertEqual(charts_read[0]["type"], WD_INLINE_CHART, "HasChart 必须为真")
            self.assertEqual(charts_read[0]["categories"], CATS)
            self.assertEqual([n for n, _ in charts_read[0]["series"]],
                             [name for name, _v, _c in SERIES])
            self.assertEqual(charts_read[0]["workbook"][0],
                             ["类别", "左侧标线", "右侧标线", "缺失率"])

    def test_edit_one_value_and_read_back(self) -> None:
        """用户在 Word 里双击图表改一格 → 图形随之更新。"""
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "第五部分.docx"
            self._build(out)
            target = 12.3
            native.edit_chart_value(out, chart_index=1, row=3, column=2, value=target)
            charts_read = native.read_charts(out, log=lambda *_: None)
            self.assertEqual(len(charts_read), 1)
            self.assertEqual([round(v, 2) for v in charts_read[0]["series"][0][1]],
                             [55.2, target, 40.5])
            self.assertEqual(round(charts_read[0]["workbook"][2][1], 2), target)
            # 落盘缓存也必须同步（否则下一次打开看到的还是旧柱）
            self.assertEqual(self._chart_cache(out),
                             [[[55.2, target, 40.5], [59.9, 48.5, 52.0], [3.1, 7.9]]])
            # 其它系列/其它点不受影响
            self.assertEqual([round(v, 2) for v in charts_read[0]["series"][1][1]],
                             [59.9, 48.5, 52.0])


if __name__ == "__main__":
    unittest.main()
