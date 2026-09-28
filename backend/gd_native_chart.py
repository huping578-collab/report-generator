"""广东第五部分统计图 —— Word 原生可编辑图表注入（brief-v4 §3 最新绘图指令，2026-09-28）。

用户指令：`_g03_chart` 的统计图不再用 PNG 嵌入 Word，改成 **Word 自身绘图**——数据内嵌 docx，
双击图表或「编辑数据」能看到并修改作图的类别与数值，改值后柱/线随之更新。
**禁止** PNG 外观 + 独立数据附件冒充；**禁止**链接外部临时 xlsx。
COM 不可用时抛 :class:`WordChartUnavailable`，绝不静默退回不可编辑的 PNG。

做法（最小改动）：python-docx 侧只在图位写一个文本锚点 + 图题，版式（居中、段前 12pt/段后 3pt、
图题序号、空段）全部沿用现有 `_figure`；本模块在 ``doc.save()`` 之后用 Word COM 打开该 docx，
按锚点顺序逐个换成原生图表、写入内嵌 ChartData 工作簿，再保存。表格与其他照片一概不动。

本机 Word 16.0 / Excel 16.0 实测要点（写在这里是因为都属于「不试不知道」的坑）：

1. 文档未落盘前取 ``ChartData.Workbook`` 必抛「发生意外」——必须先 SaveAs2/Save。
   注入流程打开的是已落盘的 docx，天然满足。
2. ChartData 区域是**一个 ListObject 表**（默认 ``$A$1:$D$5`` / 3 系列 / 4 类别）：
   既不能 ``Cells.Clear()``（表会碎），也不能 ``Chart.SetSourceData()``（抛「发生意外」）。
   正确做法是 ``ListObjects(1).Resize()`` 到目标矩形，再往格子里写字。
   **但 Resize 不会改图表的 SERIES 公式**——必须再逐系列重设公式，见 ``_rebind_series``。
3. 图表数据窗格开着时 Word 会抛 ``RPC_E_CALL_REJECTED``（被呼叫方拒绝接收呼叫），
   所有 COM 调用都要过 :func:`com_retry`。写完数据网格要 ``Workbook.Close()``，
   否则下一张图的 ``ChartData.Activate`` 报「图表数据网格已在…中打开」。
4. ``Application.CentimetersToPoints`` 在 gen_py 缓存下抛「未指定的错误」，直接用常数换算。
5. ``Application.Quit()`` 是异步的：不等进程真退出就删/重开 docx 会报
   ``PermissionError: WinError 32``（文件被 Word 持锁），见 ``_Word._await_release``。
"""
from __future__ import annotations

import time
from dataclasses import dataclass, field

__all__ = [
    "ChartSpec",
    "WordChartError",
    "WordChartUnavailable",
    "build_anchor",
    "edit_chart_value",
    "inject_charts",
    "read_charts",
    "POINTS_PER_CM",
]

# 1 cm = 28.3464567 pt（72/2.54）。不调 Application.CentimetersToPoints，见模块 docstring 第 4 条。
POINTS_PER_CM = 28.3464567

# Excel 图表常量（win32com 的 gen_py 缓存可能过期，写死更省事也更稳）。
XL_COLUMN_CLUSTERED = 51
XL_BAR_CLUSTERED = 57
XL_LINE_MARKERS = 65
XL_CATEGORY, XL_VALUE = 1, 2
XL_LEGEND_TOP_RIGHT = -4152
XL_LABEL_OUTSIDE_END = 2
WD_INLINE_CHART = 12                # wdInlineShapeChart（照片是 3，读图表时要跳过）

# 边框：主题白「深色 15%」= #D9D9D9、1 磅实线（与 `_picture_border` 同一口径）。
GD_CHART_BORDER_RGB = 0xD9D9D9
GD_CHART_BORDER_POINTS = 1.0

# 锚点前缀/后缀。锚点是纯 ASCII，避免中文标点在 Find 里的转义问题。
ANCHOR_PREFIX = "@@GDCHART"
ANCHOR_SUFFIX = "@@"

# Word 忙时抛的两个 HRESULT。
_RPC_E_CALL_REJECTED = -2147418111
_RPC_E_SERVERCALL_RETRYLATER = -2147417846

# 类别标签超过这个长度/数量就旋转 45°，避免压字（对齐 P4c 的 `_need_rotate` 口径：
# ①图标签 6~7 字横排不压，②图「路线号+管养单位」20~27 字必转）。
GD_CHART_LABEL_ROTATE_LEN = 12


class WordChartError(RuntimeError):
    """原生图表注入/读取失败。"""


class WordChartUnavailable(WordChartError):
    """Word COM 不可用（未装 Word、pywin32 缺失、Word 起不来）。

    单独一类，便于上层给出明确的环境依赖提示，而不是产出一张不可编辑的 PNG。
    """


@dataclass
class ChartSpec:
    """一张待注入的原生图表。

    ``anchor`` 留空，由 ``_figure`` 在写锚点那一刻按本次文档内的插入顺序分配（见 `build_anchor`）——
    python-docx 侧不需要知道最终序号，注入侧按 spec 列表顺序取锚点，两边靠同一个列表对齐。
    """

    categories: list
    series: list                       # [(系列名, [值...], "#RRGGBB")]，值可为 None（留空 = 断点）
    anchor: str = ""
    ylabel: str = "合格率（%）"
    kind: str = "bar"                  # "bar" | "line"
    ylim: tuple | None = None
    value_fmt: str = "%.1f"            # "%.1f" / "%.2f"
    horizontal: bool = False
    rotation: int = 0

    def chart_type(self) -> int:
        if self.kind == "line":
            return XL_LINE_MARKERS
        return XL_BAR_CLUSTERED if self.horizontal else XL_COLUMN_CLUSTERED

    def decimals(self) -> int:
        """``value_fmt`` 的小数位数（``"%.1f"`` → 1，``"%.2f"`` → 2，无小数点 → 0）。

        注意 ``%.2f`` 的精度是「点后那个整数 2」，不是「点后有几个字符」——
        按字符数算会得到 1 位小数（第一版就踩了这个坑）。
        """
        if "." not in self.value_fmt:
            return 0
        return int(self.value_fmt.split(".")[-1].rstrip("f"))

    def number_format(self) -> str:
        """matplotlib 的 ``%.1f`` / ``%.2f`` → Excel 数字格式 ``0.0`` / ``0.00``。"""
        digits = self.decimals()
        return "0" if not digits else "0." + "0" * digits

    def need_rotate(self) -> bool:
        if self.rotation:
            return True
        longest = max((len(str(c)) for c in self.categories), default=0)
        return longest > GD_CHART_LABEL_ROTATE_LEN or len(self.categories) > 12


# ============================ COM 基础设施 ============================

def _import_com():
    """延迟导入 pywin32。缺依赖 → WordChartUnavailable（不是 PNG 兜底）。"""
    try:
        import pythoncom
        import win32com.client as win32
    except Exception as exc:                      # pywin32 没装
        raise WordChartUnavailable(
            f"Word 原生图表需要 pywin32，但导入失败：{exc!r}。请 pip install pywin32，"
            f"并安装 Microsoft Word。本工具不会退回 PNG 嵌图（那样图不可编辑）。"
        ) from exc
    return pythoncom, win32


def com_retry(func, attempts=25, delay=0.4, what=""):
    """Word 忙（图表数据窗格开着）时重试；非忙错误直接抛。"""
    last = None
    for _ in range(attempts):
        try:
            return func()
        except Exception as exc:                  # noqa: BLE001 - 需按 hresult 分流
            code = getattr(exc, "hresult", None)
            if code is None and getattr(exc, "args", None):
                code = exc.args[0]
            if code not in (_RPC_E_CALL_REJECTED, _RPC_E_SERVERCALL_RETRYLATER):
                raise
            last = exc
            time.sleep(delay)
    raise WordChartError(f"Word COM 调用持续被拒{('：' + what) if what else ''}：{last!r}")


def _bgr(hex_color: str) -> int:
    """``#RRGGBB`` → Word/Excel 用的 BGR 整数。"""
    value = str(hex_color or "").lstrip("#")
    if len(value) != 6:
        return 0
    red, green, blue = (int(value[i:i + 2], 16) for i in (0, 2, 4))
    return red | (green << 8) | (blue << 16)


def _col_letter(index: int) -> str:
    """1 → A。图表系列数远小于 26，单字母够用。"""
    if index < 1 or index > 26:
        raise WordChartError(f"系列/类目列索引越界：{index}（1..26）")
    return chr(ord("A") + index - 1)


def _process_gone(app) -> bool:
    """Word 进程是否已退出。取进程句柄不可靠（DispatchEx 的 PID 会变），用文档数判。"""
    try:
        return int(app.Documents.Count) == 0 and not int(app.Tasks.Count)
    except Exception:                                 # noqa: BLE001 - 取不到就当还在
        return False


class _Word:
    """Word COM 会话上下文：起进程、保证 Quit、Dispose。

    三个入口（注入/读取/改值）都要「起 Word → 干活 → 一定关掉」，抽这一处是因为
    崩在中间会留下持锁的 WINWORD 进程，下次打开同一 docx 直接报「已在别处打开」。
    """

    def __init__(self, log=print):
        self.pythoncom, self.win32 = _import_com()
        try:
            self.pythoncom.CoInitialize()
        except Exception:                             # 已初始化过，忽略
            pass
        try:
            self.app = self.win32.DispatchEx("Word.Application")
        except Exception as exc:
            raise WordChartUnavailable(
                f"无法启动 Microsoft Word 自动化：{exc!r}。请确认本机已安装 Word 16 且可正常打开。"
            ) from exc
        self.app.Visible = False
        self.log = log
        com_retry(lambda: setattr(self.app, "DisplayAlerts", 0), what="DisplayAlerts")
        self.document = None

    def open(self, docx_path, read_only=False):
        self.document = com_retry(
            lambda: self.app.Documents.Open(str(docx_path), ReadOnly=read_only),
            what="Documents.Open")
        return self.document

    def __enter__(self):
        return self

    def __exit__(self, *_exc):
        # 先关文档再退进程。反序（先 Quit 再 Close）会留下持锁的 WINWORD，
        # 下一个用例/下一次跑批打开同一 docx 直接报「已在别处打开」。
        if self.document is not None:
            try:
                com_retry(lambda: self.document.Close(0), what="Document.Close")
            except Exception:                         # noqa: BLE001 - 关闭失败不掩盖主异常
                pass
            self.document = None
        try:
            self.app.Quit()
        except Exception:                             # noqa: BLE001
            pass
        # Quit 是异步的：进程真正退出会把文件锁放开。不等的话紧接着的
        # TemporaryDirectory 清理会报 PermissionError（WinError 32）。
        self._await_release()
        return False

    def _await_release(self, attempts=20, delay=0.25):
        """等 Word 真正退出（文件锁放开）。超时只警告，不抛——不能因清理问题毁掉已完成的工作。"""
        import time as _time
        for _ in range(attempts):
            try:
                if not self.app.Documents.Count and _process_gone(self.app):
                    return True
            except Exception:                         # noqa: BLE001 - 进程已死也算释放
                return True
            _time.sleep(delay)
        return False


def _workbook(chart, pythoncom):
    """取 ChartData 内嵌工作簿（带 Excel 冷启动重试）。"""
    last = None
    for _ in range(10):
        try:
            com_retry(chart.ChartData.Activate, what="ChartData.Activate")
            return com_retry(lambda: chart.ChartData.Workbook, what="ChartData.Workbook")
        except Exception as exc:                  # noqa: BLE001
            last = exc
            time.sleep(0.4)
    raise WordChartError(f"ChartData.Workbook 不可用：{last!r}")


# ============================ 注入 ============================

def _write_data(chart, spec: ChartSpec, pythoncom):
    """把 spec 写进内嵌工作簿并让图表跟随。"""
    if not spec.categories or not spec.series:
        raise WordChartError(f"图表 {spec.anchor} 没有类别或系列，拒绝写入空图")
    workbook = _workbook(chart, pythoncom)
    sheet = com_retry(lambda: workbook.Worksheets(1), what="Worksheets(1)")
    target = f"A1:{_col_letter(1 + len(spec.series))}{1 + len(spec.categories)}"
    # 必须是 Resize 而不是 Cells.Clear()/SetSourceData，见模块 docstring 第 2 条。
    table = com_retry(lambda: sheet.ListObjects(1), what="ListObjects(1)")
    com_retry(lambda: table.Resize(sheet.Range(target)), what="ListObjects.Resize")

    com_retry(lambda: setattr(sheet.Cells(1, 1), "Value", "类别"))
    for j, (name, _values, _color) in enumerate(spec.series):
        com_retry(lambda j=j, name=name: setattr(sheet.Cells(1, 2 + j), "Value", str(name)))
    for i, category in enumerate(spec.categories):
        com_retry(lambda i=i, c=category: setattr(sheet.Cells(2 + i, 1), "Value", str(category)))
    for j, (_name, values, _color) in enumerate(spec.series):
        for i, value in enumerate(values):
            if value is None:                     # None = 留空单元格（Excel 画成断点）
                continue
            com_retry(lambda i=i, j=j, v=value: setattr(sheet.Cells(2 + i, 2 + j), "Value", float(value)))
    # 必须关掉数据网格：Word 一次只允许一个图表数据网格处于打开状态，不关的话
    # 下一张图的 ChartData.Activate 会抛「图表数据网格已在…中打开」（广州实跑第一张后就撞上）。
    # 顺序要紧：先删多余系列 + 改 SERIES 公式，再关网格。
    _rebind_series(chart, len(spec.series), 1 + len(spec.categories))
    com_retry(lambda: workbook.Close(), what="ChartData Workbook.Close")
    com_retry(lambda: chart.Refresh(), what="chart.Refresh")
    return target


def _rebind_series(chart, series_count: int, last_row: int):
    """把图表的 SERIES 公式重新指到 (类别数 × 系列数) 的实际区域。

    **这是本模块最容易踩的坑**：`ListObjects.Resize` 只改表，不改图表的 SERIES 公式。
    Word 建图时公式写死了默认的 4 行 × 3 系列（如 ``Sheet1!$B$2:$B$5``），
    Resize 之后公式纹丝不动 → 图表仍按 4 行读，写进去的第 3 类目看不见，
    还会把表外的残留默认值（4.5 / 2.8）当数据画出来。
    实测：3 类目 2 系列的图缓存读回 ``[[55.2, 61, 40.5], [59.9, 48.5, 1.8]]``，
    末位 1.8 是 Word 建表时的默认值；改完公式才 ``MATCH``。
    """
    collection = com_retry(lambda: chart.SeriesCollection(), what="SeriesCollection")
    while com_retry(lambda: collection.Count, what="SeriesCollection.Count") > series_count:
        index = com_retry(lambda: collection.Count, what="SeriesCollection.Count")
        com_retry(lambda i=index: collection(i).Delete(), what="Series.Delete")
    for j in range(series_count):
        column = _col_letter(2 + j)
        formula = (f"=SERIES(Sheet1!${column}$1,Sheet1!$A$2:$A${last_row},"
                   f"Sheet1!${column}$2:${column}${last_row},{j + 1})")
        com_retry(lambda j=j, formula=formula: setattr(collection(j + 1), "Formula", formula),
                  what="Series.Formula")


def _style_chart(chart, spec: ChartSpec):
    """图内样式：数值轴标题/范围、图例、柱顶数据标签、类别标签旋转、浅灰细边框。

    刻意**不开** ChartTitle：图题由 Word 的图题段（图5-x …）承担，图内再放标题会重复。
    """
    com_retry(lambda: setattr(chart, "HasTitle", False), what="HasTitle")

    series_collection = com_retry(lambda: chart.SeriesCollection(), what="SeriesCollection")
    for j, (_name, _values, color) in enumerate(spec.series):
        point = com_retry(lambda j=j: series_collection(j + 1), what="SeriesCollection(i)")
        com_retry(lambda p=point: setattr(p.Format.Fill, "Visible", True), what="Fill.Visible")
        com_retry(lambda p=point: setattr(p.Format.Fill.ForeColor, "RGB", _bgr(color)),
                  what="Fill.ForeColor")
        com_retry(lambda p=point: setattr(p, "HasDataLabels", True), what="HasDataLabels")
        labels = com_retry(lambda p=point: p.DataLabels(), what="DataLabels")
        com_retry(lambda lb=labels: setattr(lb, "NumberFormat", spec.number_format()),
                  what="DataLabels.NumberFormat")
        com_retry(lambda lb=labels: setattr(lb, "Position", XL_LABEL_OUTSIDE_END),
                  what="DataLabels.Position")

    com_retry(lambda: setattr(chart, "HasLegend", True), what="HasLegend")
    com_retry(lambda: setattr(chart.Legend, "Position", XL_LEGEND_TOP_RIGHT), what="Legend.Position")

    value_axis = com_retry(lambda: chart.Axes(XL_VALUE), what="Axes(xlValue)")
    com_retry(lambda: setattr(value_axis, "HasTitle", True), what="valueAxis.HasTitle")
    com_retry(lambda: setattr(value_axis.AxisTitle, "Text", spec.ylabel), what="AxisTitle.Text")
    if spec.ylim:
        low, high = spec.ylim
        com_retry(lambda: setattr(value_axis, "MinimumScale", float(low)), what="MinimumScale")
        com_retry(lambda: setattr(value_axis, "MaximumScale", float(high)), what="MaximumScale")

    category_axis = com_retry(lambda: chart.Axes(XL_CATEGORY), what="Axes(xlCategory)")
    if spec.need_rotate():
        # 45° 斜排：大量类目 / 长标签时防止标签互相压住。
        com_retry(lambda: setattr(category_axis.TickLabels, "Orientation", 45),
                  what="TickLabels.Orientation")

    # 浅灰细边框（与 `_picture_border` 同色 #D9D9D9、1 磅实线）。
    line = com_retry(lambda: chart.ChartArea.Format.Line, what="ChartArea.Format.Line")
    com_retry(lambda: setattr(line, "Visible", True), what="Line.Visible")
    com_retry(lambda: setattr(line, "Weight", GD_CHART_BORDER_POINTS), what="Line.Weight")
    com_retry(lambda: setattr(line.ForeColor, "RGB", GD_CHART_BORDER_RGB), what="Line.ForeColor.RGB")


def _find_anchor(doc, anchor: str):
    """按锚点文本定位 Range；找不到直接抛（说明 python-docx 侧与注入侧脱节了）。"""
    rng = doc.Content
    find = rng.Find
    com_retry(lambda: find.ClearFormatting(), what="Find.ClearFormatting")
    com_retry(lambda: setattr(find, "Text", anchor), what="Find.Text")
    com_retry(lambda: setattr(find, "Forward", True), what="Find.Forward")
    com_retry(lambda: setattr(find, "Wrap", 0), what="Find.Wrap")       # wdFindStop
    com_retry(lambda: setattr(find, "MatchWildcards", False), what="Find.MatchWildcards")
    if not com_retry(lambda: find.Execute(), what="Find.Execute"):
        raise WordChartError(f"文档里找不到图表锚点 {anchor!r}，图表顺序已与文档脱节")
    return rng


def inject_charts(docx_path, specs, width_cm: float = 15.0, height_cm: float = 7.5, log=print):
    """把 ``specs`` 里的原生图表注入已落盘的 docx，就地保存。返回注入张数。

    ``specs`` 的顺序 = 文档里锚点的顺序，按序注入，保证图与图题一一对应。
    """
    specs = [spec for spec in specs if spec is not None]
    if not specs:
        return 0
    with _Word(log=log) as word:
        document = word.open(docx_path)
        for spec in specs:
            rng = _find_anchor(document, spec.anchor)
            # 先删锚点文本再插图：Range 塌缩后插入，图表落进锚点所在段，
            # 该段的居中/段前 12pt/段后 3pt 版式因此原样保留。
            start = com_retry(lambda: rng.Start, what="Range.Start")
            com_retry(lambda: setattr(rng, "Text", ""), what="Range.Text=''")
            spot = com_retry(lambda: document.Range(start, start), what="Range")
            shape = com_retry(
                lambda: document.InlineShapes.AddChart2(-1, spec.chart_type(), spot),
                what=f"AddChart2({spec.anchor})")
            com_retry(lambda: setattr(shape, "Width", width_cm * POINTS_PER_CM), what="shape.Width")
            com_retry(lambda: setattr(shape, "Height", height_cm * POINTS_PER_CM), what="shape.Height")
            chart = com_retry(lambda: shape.Chart, what="shape.Chart")
            source = _write_data(chart, spec, word.pythoncom)
            _style_chart(chart, spec)
            log(f"[原生图表] {spec.anchor} → {source} "
                f"（{len(spec.categories)} 类目 / {len(spec.series)} 系列）")
        com_retry(lambda: document.Save(), what="Document.Save")
    return len(specs)


# ============================ 读取（验收/测试用） ============================

def _chart_of(document, chart_index: int):
    """取第 ``chart_index`` 张原生图表（1 起，按 InlineShapes 顺序，跳过照片）。"""
    seen = 0
    count = com_retry(lambda: document.InlineShapes.Count, what="InlineShapes.Count")
    for index in range(1, count + 1):
        shape = com_retry(lambda i=index: document.InlineShapes(i), what="InlineShapes(i)")
        if com_retry(lambda: shape.Type, what="InlineShape.Type") != WD_INLINE_CHART:
            continue
        seen += 1
        if seen == chart_index:
            return com_retry(lambda: shape.Chart, what="shape.Chart")
    raise WordChartError(f"文档里没有第 {chart_index} 张原生图表（共 {seen} 张）")


def read_charts(docx_path, log=print):
    """用 Word COM 打开 docx，逐图回读类别/系列/值。

    返回 ``[{"inline": i, "type": 12, "categories": [...], "series": [(名, [值...]), ...],
    "workbook": [[单元格...], ...]}...]``。``type == 12``（wdInlineShapeChart）即 HasChart 为真。
    """
    results = []
    with _Word(log=log) as word:
        document = word.open(docx_path, read_only=True)
        count = com_retry(lambda: document.InlineShapes.Count, what="InlineShapes.Count")
        for index in range(1, count + 1):
            shape = com_retry(lambda i=index: document.InlineShapes(i), what="InlineShapes(i)")
            shape_type = com_retry(lambda: shape.Type, what="InlineShape.Type")
            if shape_type != WD_INLINE_CHART:    # 照片(3)等跳过
                continue
            chart = com_retry(lambda: shape.Chart, what="shape.Chart")
            collection = com_retry(lambda: chart.SeriesCollection(), what="SeriesCollection")
            series_count = com_retry(lambda: collection.Count, what="SeriesCollection.Count")
            categories, series = [], []
            for i in range(1, series_count + 1):
                point = com_retry(lambda i=i: collection(i), what="SeriesCollection(i)")
                if not categories:
                    categories = [str(v) for v in com_retry(lambda p=point: p.XValues, what="XValues")]
                name = com_retry(lambda p=point: p.Name, what="Series.Name")
                values = [None if v is None else float(v)
                          for v in com_retry(lambda p=point: p.Values, what="Series.Values")]
                series.append((str(name), values))
            grid = []
            workbook = None
            try:
                workbook = _workbook(chart, word.pythoncom)
                sheet = com_retry(lambda: workbook.Worksheets(1), what="Worksheets(1)")
                used = com_retry(lambda: sheet.UsedRange, what="UsedRange")
                rows = com_retry(lambda: used.Rows.Count, what="UsedRange.Rows.Count")
                cols = com_retry(lambda: used.Columns.Count, what="UsedRange.Columns.Count")
                for r in range(1, rows + 1):
                    grid.append([sheet.Cells(r, c).Value for c in range(1, cols + 1)])
            except Exception as exc:              # noqa: BLE001 - 读不到工作簿本身就是要报的事
                raise WordChartError(f"第 {index} 张图的内嵌工作簿读不出来：{exc!r}") from exc
            finally:
                # 同 _write_data：一次只能开一个数据网格，不关会影响下一张。
                if workbook is not None:
                    try:
                        com_retry(lambda: workbook.Close(), what="Workbook.Close")
                    except Exception:             # noqa: BLE001 - 读路径的清理失败不影响已取到的数据
                        pass
            results.append({"inline": index, "type": shape_type, "categories": categories,
                            "series": series, "workbook": grid})
            log(f"[回读] 图{index}: {len(categories)} 类目 / {len(series)} 系列")
    return results


def edit_chart_value(docx_path, chart_index: int, row: int, column: int, value: float):
    """改第 ``chart_index`` 张图内嵌工作簿的 ``Cells(row, column)`` 并保存。

    这正是用户在 Word 里双击图表 →「编辑数据」手动做的事，这里用脚本走同一条路，
    用来证明数据确实可写、且改值后图表随之更新（而不是 PNG 那种改了不动的假可编辑）。

    **顺序是 ``CalculateFullRebuild → chart.Refresh() → workbook.Close()``**，反了不生效
    （2026-09-28 实测，7 种序列两两对照，12 图真实报告）：

    ====================  ================
    序列                   落盘 c:numCache
    ====================  ================
    写 → 算 → 关 → 刷新     旧值 ❌
    写 → 算 → 刷新 → 关 ✅   新值 ✅
    写 → 算 → 激活 → 关→刷新 旧值 ❌
    写 → 算 → 激活 → 刷新→关 新值 ✅
    写 → 算 → 退 Excel → 刷新 旧值 ❌
    写 → 算 → 重设源区域→关   新值 ✅
    ====================  ================

    规律：**工作簿还开着的时候刷新才有效**；一旦 ``Close()``，Word 就把内嵌数据源摘掉，
    后面的 ``Refresh()`` 变成空转，落盘的缓存还是旧值。Word 自己建的图与离线生成的图
    在这里表现完全一致，所以这是 COM 序列问题，不是图表结构问题。
    """
    with _Word() as word:
        document = word.open(docx_path)
        chart = _chart_of(document, chart_index)
        workbook = _workbook(chart, word.pythoncom)
        try:
            sheet = com_retry(lambda: workbook.Worksheets(1), what="Worksheets(1)")
            com_retry(lambda: setattr(sheet.Cells(row, column), "Value", float(value)),
                      what=f"Cells({row},{column}).Value")
            # 让 Excel 真正重算并把新值推给图表；缺这一步图表不动。
            com_retry(lambda: workbook.Application.CalculateFullRebuild(),
                      what="Application.CalculateFullRebuild")
            # 必须在关工作簿**之前**刷新，否则 Word 已经摘掉内嵌数据源，刷新是空转。
            com_retry(lambda: chart.Refresh(), what="chart.Refresh")
        finally:
            com_retry(lambda: workbook.Close(), what="Workbook.Close")
        com_retry(lambda: document.Save(), what="Document.Save")
    return float(value)


def build_anchor(index: int) -> str:
    """锚点文本。纯 ASCII，Word Find 不用转义。"""
    return f"{ANCHOR_PREFIX}{index:04d}{ANCHOR_SUFFIX}"
