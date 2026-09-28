"""P4-Chart-OOXML：Word 原生可编辑图表的**离线**注入（方案 B，生成阶段零 Word/Excel）。

用户最终口径（brief-v4 §3，2026-09-28）：第五部分统计图必须是 Word **原生**图表——
用户在 Word 里双击图表 /「编辑数据」能看到并修改作图用的类别与数值，改值后柱/线随之更新。
生成阶段**不启动 Word/Excel**，只靠 Python 标准库 ZIP/XML + 已装的 openpyxl 直接在 DOCX
的 OPC 包里生成 ``word/charts/chartN.xml``、``word/charts/_rels/chartN.xml.rels``、
``word/embeddings/gd_chartN.xlsx``、绘图锚点，并同步 ``[Content_Types].xml`` 与
``word/_rels/document.xml.rels``。同一进程内没有 Office COM、没有子进程。

做法（为什么必须"生成"而不能只"改"种子）：种子 ``assets/gd_native_chart_seed.docx``
是 Word 真机建的 1 图样本，里面 `c:ser` 只有 2 个系列 × 2 个类目， SERIES 公式写死
``Sheet1!$B$2:$B$3``。直接改字符串既脆又容易留下 Word 建表时的默认残值
（历史实测：Word 建表默认 3 系列 4 类目，Resize 不会改 SERIES 公式，残值 1.8/4.5 会被当数据画）。
所以本模块只从种子取三样**与数据无关**的东西——``style1.xml``/``colors1.xml``
（微软图表样式部件，手写不划算）与 Word 原生绘图锚点的 XML 形状；``c:chart`` 本体
则按 :class:`~backend.gd_native_chart.ChartSpec` 重新生成，缓存/公式/ptCount 全部由同一份
数据推出，因此三者（内嵌工作簿 / ``c:numCache`` / 表值）天然逐点一致。

原子交付：先写同目录临时文件 → :func:`verify_docx_charts` 全量校验 → 校验通过才
:func:`os.replace` 覆盖目标。任一步失败只抛异常，旧的完整文件原封不动，绝不留下
含占位符或图不全的 docx。

Word COM 只用于**独立验收脚本**（真机打开—改一格—保存—重开读回）；生成入口永不触碰，
``backend.gd_native_chart`` 的 COM 实现保留作 QA helper，测试用导入钩子证明生成路径不拉 Office。
"""
from __future__ import annotations

import io
import os
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape, quoteattr

from lxml import etree

from backend.gd_native_chart import ChartSpec

__all__ = [
    "ChartSpec",
    "OoxmlChartError",
    "inject_charts",
    "save_with_charts",
    "seed_path",
    "verify_docx_charts",
]

# ---------------------------------------------------------------------------
# OPC / OOXML 常量
# ---------------------------------------------------------------------------
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

RT_CHART = f"{R}/chart"
RT_CHART_STYLE = "http://schemas.microsoft.com/office/2011/relationships/chartStyle"
RT_CHART_COLORS = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle"
RT_PACKAGE = f"{R}/package"

CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
CT_CHART_STYLE = "application/vnd.ms-office.chartstyle+xml"
CT_CHART_COLORS = "application/vnd.ms-office.chartcolorstyle+xml"
CT_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

EMU_PER_CM = 360000
DEG = 60000                     # a:bodyPr/@rot 的单位（1/60000 度）

# 轴 ID：chart 部件内部唯一即可，固定值与种子一致，避免每次生成都不同（便于 diff）。
CAT_AX_ID = 973072816
VAL_AX_ID = 932770576

# 类别轴标签旋转 45°（Word/Excel 的斜排），与旧 matplotlib 口径一致。
ROTATE_45 = -45 * DEG
# 种子 Word 默认写的是 -1000°（等价于不旋转），沿用以保持非旋转图与 Word 原生一致。
NO_ROTATE = -60000000


class OoxmlChartError(RuntimeError):
    """原生图表 OOXML 注入/校验失败。"""


def seed_path() -> Path:
    """静态种子 docx 的绝对路径。PyInstaller 打包后 ``_MEIPASS`` 优先。

    桌面版必须能在**没有 Word/Excel 的机器**上生成，所以种子是随包静态资源，不是 Temp 文件。
    """
    candidates = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "assets" / "gd_native_chart_seed.docx")
    candidates.append(Path(__file__).resolve().parent.parent / "assets" / "gd_native_chart_seed.docx")
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    # P2-2：打印**全部**候选。原来只报 `candidates[0]`，而 `_MEIPASS` 存在却缺种子时
    # 那条路径可能压根不存在，报错会把人指向一个错误的文件。
    raise OoxmlChartError(
        "找不到静态种子，已尝试：\n  " + "\n  ".join(str(candidate) for candidate in candidates)
        + "\n种子必须在 assets/gd_native_chart_seed.docx，"
        "且 report_generator.spec 的 datas 需把它打进桌面版。")


# ---------------------------------------------------------------------------
# 种子：只取与数据无关的样式部件
# ---------------------------------------------------------------------------
_STYLE_CACHE: dict[str, bytes] = {}


def _seed_styles() -> dict[str, bytes]:
    """从种子 docx 读 ``style1.xml`` / ``colors1.xml``（微软图表样式部件）与绘图锚点样本。"""
    if not _STYLE_CACHE:
        with zipfile.ZipFile(seed_path()) as archive:
            _STYLE_CACHE["style"] = archive.read("word/charts/style1.xml")
            _STYLE_CACHE["colors"] = archive.read("word/charts/colors1.xml")
    return _STYLE_CACHE


# ---------------------------------------------------------------------------
# c:chart XML 生成
# ---------------------------------------------------------------------------
def _num_text(value: float) -> str:
    """数值 → ``<c:v>`` 文本。``repr`` = Python 3 最短往返表示，cache 精度 = 源浮点。

    P2-1：原为 ``%.10g``（10 位有效数字），把 ``59.892086330935257`` 截成 ``59.89208633``，
    让「cache 严格等于源值」这句断言不成立（Word 改存后写回的是全精度）。
    改 repr 后两侧一致；`verify_docx_charts` 的 1e-6 容差断言不受影响（只会更严）。
    """
    return repr(float(value))


def _cat_block(spec: ChartSpec, last_row: int) -> str:
    points = "".join(f'<c:pt idx="{i}"><c:v>{escape(str(cat))}</c:v></c:pt>'
                     for i, cat in enumerate(spec.categories))
    return (f"<c:cat><c:strRef><c:f>Sheet1!$A$2:$A${last_row}</c:f>"
            f'<c:strCache><c:ptCount val="{len(spec.categories)}"/>{points}</c:strCache>'
            "</c:strRef></c:cat>")


def _val_block(name_col: str, values: list, last_row: int) -> str:
    # None = 断点：Excel 留空格子，OOXML 的正确表示就是**不写该 <c:pt>**（idx 跳号），
    # 配合 chartSpace 的 <c:dispBlanksAs val="gap"/>。内嵌工作簿同样留空，两边天然一致。
    points = "".join(f'<c:pt idx="{i}"><c:v>{_num_text(v)}</c:v></c:pt>'
                     for i, v in enumerate(values) if v is not None)
    return (f"<c:val><c:numRef><c:f>Sheet1!${name_col}$2:${name_col}${last_row}</c:f>"
            '<c:numCache><c:formatCode>General</c:formatCode>'
            f'<c:ptCount val="{len(values)}"/>{points}</c:numCache>'
            "</c:numRef></c:val>")


def _check_shape(spec: ChartSpec) -> None:
    """每个系列的点数必须等于类目数——否则缓存/工作簿/图形会三方不一致。

    这里显式报错而不是让下标越界：调用方（``_g03_chart``）拼错长度时要看到
    「哪个系列有几个点」这种能直接定位的信息。
    """
    count = len(spec.categories)
    for index, (name, values, _color) in enumerate(spec.series):
        if len(values) != count:
            raise OoxmlChartError(
                f"图表 {spec.anchor!r} 系列 {index + 1}「{name}」有 {len(values)} 个点，"
                f"类目有 {count} 个；两者必须相等（ptCount 与内嵌工作簿都按类目数写）")


def _d_lbls(spec: ChartSpec) -> str:
    """柱顶/点旁数据标签（与旧 COM ``_style_chart`` 同口径：外侧端 + spec 的小数位）。"""
    return ("<c:dLbls>"
            f'<c:numFmt formatCode={quoteattr(spec.number_format())} sourceLinked="0"/>'
            "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>"
            "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"900\"/></a:pPr>"
            "<a:endParaRPr lang=\"zh-Hans-CN\"/></a:p></c:txPr>"
            '<c:dLblPos val="outEnd"/>'
            '<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>'
            '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>'
            "</c:dLbls>")


def _ser_block(spec: ChartSpec, index: int, last_row: int) -> str:
    name, values, color = spec.series[index]
    column = _col_letter(2 + index)
    hexcolor = str(color or "").lstrip("#")
    if len(hexcolor) != 6:
        hexcolor = "4F81BD"
    head = (f'<c:ser><c:idx val="{index}"/><c:order val="{index}"/>'
            f"<c:tx><c:strRef><c:f>Sheet1!${column}$1</c:f>"
            f'<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{escape(str(name))}</c:v>'
            "</c:pt></c:strCache></c:strRef></c:tx>"
            f'<c:spPr><a:solidFill><a:srgbClr val="{hexcolor}"/></a:solidFill>'
            "<a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>")
    if spec.kind == "line":
        head += ('<c:marker><c:symbol val="circle"/><c:size val="5"/>'
                 f'<c:spPr><a:solidFill><a:srgbClr val="{hexcolor}"/></a:solidFill>'
                 "<a:ln><a:noFill/></a:ln></c:spPr></c:marker>")
    else:
        head += '<c:invertIfNegative val="0"/>'
    return (head + _d_lbls(spec) + _cat_block(spec, last_row)
            + _val_block(column, values, last_row) + "</c:ser>")


def _col_letter(index: int) -> str:
    """1 → A。系列数远小于 26，单字母够用。"""
    if not 1 <= index <= 26:
        raise OoxmlChartError(f"系列列索引越界：{index}（1..26）")
    return chr(ord("A") + index - 1)


def _axis_text_props(rot: int | None) -> str:
    body = f'<a:bodyPr rot="{rot}" ' if rot is not None else "<a:bodyPr "
    return (f"<c:txPr>{body}spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" "
            'vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/>'
            '<a:p><a:pPr><a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" '
            'baseline="0"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/>'
            '<a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/>'
            '<a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr>'
            '<a:endParaRPr lang="zh-Hans-CN"/></a:p></c:txPr>')


def _val_axis(spec: ChartSpec) -> str:
    scaling = '<c:scaling><c:orientation val="minMax"/>'
    if spec.ylim:
        low, high = spec.ylim
        scaling += f'<c:max val="{_num_text(high)}"/><c:min val="{_num_text(low)}"/>'
    scaling += "</c:scaling>"
    # P3-2（用户 2026-09-28）：不要坐标轴标题。这里**不再生成 c:title**，
    # 不是把 val 设 0 —— 轴标题的元素在 catAx/valAx/serAx 三处都要整个不出现。
    # 保留网格线、刻度、数据标签（都在本函数后半段，未动）。
    return (f'<c:valAx><c:axId val="{VAL_AX_ID}"/>{scaling}<c:delete val="0"/>'
            '<c:axPos val="l"/>'
            '<c:majorGridlines><c:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">'
            '<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/>'
            '<a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln>'
            "</c:spPr></c:majorGridlines>"
            '<c:numFmt formatCode="General" sourceLinked="1"/>'
            '<c:majorTickMark val="none"/><c:minorTickMark val="none"/>'
            '<c:tickLblPos val="nextTo"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
            + _axis_text_props(NO_ROTATE) +
            f'<c:crossAx val="{CAT_AX_ID}"/><c:crosses val="autoZero"/>'
            '<c:crossBetween val="between"/></c:valAx>')


def _cat_axis(spec: ChartSpec) -> str:
    rot = ROTATE_45 if spec.need_rotate() else NO_ROTATE
    return (f'<c:catAx><c:axId val="{CAT_AX_ID}"/>'
            '<c:scaling><c:orientation val="minMax"/></c:scaling>'
            '<c:delete val="0"/><c:axPos val="b"/>'
            '<c:numFmt formatCode="General" sourceLinked="1"/>'
            '<c:majorTickMark val="none"/><c:minorTickMark val="none"/>'
            '<c:tickLblPos val="nextTo"/><c:spPr><a:noFill/>'
            '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>'
            '<a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/>'
            "</a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>"
            + _axis_text_props(rot) +
            f'<c:crossAx val="{VAL_AX_ID}"/><c:crosses val="autoZero"/><c:auto val="1"/>'
            '<c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>')


def build_chart_xml(spec: ChartSpec) -> bytes:
    """按 spec 生成完整的 ``c:chartSpace``（含缓存与 SERIES 公式，外部数据指向内嵌 xlsx）。"""
    if not spec.categories or not spec.series:
        raise OoxmlChartError(f"图表 {spec.anchor!r} 没有类别或系列，拒绝写入空图")
    _check_shape(spec)
    last_row = 1 + len(spec.categories)
    sers = "".join(_ser_block(spec, j, last_row) for j in range(len(spec.series)))
    if spec.kind == "line":
        plot = ("<c:lineChart><c:grouping val=\"standard\"/><c:varyColors val=\"0\"/>"
                f'{sers}<c:marker val="1"/><c:smooth val="0"/>'
                f'<c:axId val="{CAT_AX_ID}"/><c:axId val="{VAL_AX_ID}"/></c:lineChart>')
    else:
        bar_dir = "bar" if spec.horizontal else "col"
        plot = ("<c:barChart>"
                f'<c:barDir val="{bar_dir}"/><c:grouping val="clustered"/>'
                f'<c:varyColors val="0"/>{sers}'
                '<c:gapWidth val="219"/><c:overlap val="-27"/>'
                f'<c:axId val="{CAT_AX_ID}"/><c:axId val="{VAL_AX_ID}"/></c:barChart>')
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart">'
        '<c:date1904 val="0"/><c:lang val="zh-CN"/><c:roundedCorners val="0"/>'
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice Requires="c14" '
        'xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">'
        '<c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback>'
        "</mc:AlternateContent>"
        "<c:chart>"
        '<c:autoTitleDeleted val="1"/>'          # 图题由 Word 的「图5-x …」图题段承担，图内不重复
        f"<c:plotArea><c:layout/>{plot}{_cat_axis(spec)}{_val_axis(spec)}"
        "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr></c:plotArea>"
        '<c:legend><c:legendPos val="l"/><c:overlay val="0"/>'   # P3-3：图例在左（原来在底 b）
        "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>"
        + _axis_text_props(None) +
        "</c:legend>"
        '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/>'
        "</c:chart>"
        # 浅灰细边框（主题白「深色 15%」= #D9D9D9、0.75 磅实线，与 _picture_border 同色系）
        "<c:spPr><a:solidFill><a:schemeClr val=\"bg1\"/></a:solidFill>"
        '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>'
        '<a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr>'
        "</a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>"
        '<c:externalData r:id="rId3"><c:autoUpdate val="0"/></c:externalData>'
        "</c:chartSpace>"
    )
    etree.fromstring(xml.encode("utf-8"))     # 结构错误的 XML 立刻炸，不进包
    return xml.encode("utf-8")


def build_workbook(spec: ChartSpec) -> bytes:
    """内嵌数据工作簿：Sheet1 + 一个覆盖实际区域的 ListObject（Word「编辑数据」同款）。"""
    from openpyxl import Workbook
    from openpyxl.worksheet.table import Table, TableStyleInfo

    book = Workbook()
    sheet = book.active
    sheet.title = "Sheet1"
    sheet.cell(row=1, column=1, value="类别")
    for j, (name, _values, _color) in enumerate(spec.series):
        sheet.cell(row=1, column=2 + j, value=str(name))
    for i, category in enumerate(spec.categories):
        sheet.cell(row=2 + i, column=1, value=str(category))
        for j, (_name, values, _color) in enumerate(spec.series):
            value = values[i]
            if value is not None:              # None → 留空（Excel 画成断点）
                sheet.cell(row=2 + i, column=2 + j, value=float(value))
    last_col = _col_letter(1 + len(spec.series))
    last_row = 1 + len(spec.categories)
    table = Table(displayName="Table1", ref=f"A1:{last_col}{last_row}")
    table.tableStyleInfo = TableStyleInfo(name=None, showRowStripes=True)
    sheet.add_table(table)
    buffer = io.BytesIO()
    book.save(buffer)
    return buffer.getvalue()


def _chart_rels(index: int) -> bytes:
    """图表部件自己的 rels：rId1/2 = 本图专属的样式部件，rId3 = 本图内嵌 xlsx。

    样式部件**必须每图一份**（不能多图共用 style1.xml）：Word 打开 2 张以上共享样式
    部件的文档会直接报「在试图打开文件时遇到错误」——2026-09-28 实测，见 inject_charts 注释。
    """
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Relationships xmlns="{PR_NS}">'
        f'<Relationship Id="rId1" Type="{RT_CHART_STYLE}" Target="style{index}.xml"/>'
        f'<Relationship Id="rId2" Type="{RT_CHART_COLORS}" Target="colors{index}.xml"/>'
        f'<Relationship Id="rId3" Type="{RT_PACKAGE}" '
        f'Target="../embeddings/gd_chart{index}.xlsx"/>'
        "</Relationships>"
    ).encode("utf-8")


# ---------------------------------------------------------------------------
# document.xml / rels / [Content_Types].xml
# ---------------------------------------------------------------------------
def _drawing_run(rid: str, docpr_id: int, cx: int, cy: int) -> etree._Element:
    """Word 原生内嵌图表的 ``w:r``（形状与种子完全同构，只换 rId / docPr / 尺寸）。"""
    xml = (
        f'<w:r xmlns:w="{W}" xmlns:wp="{WP}" xmlns:r="{R}">'
        '<w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="10795" b="5715"/>'
        f'<wp:docPr id="{docpr_id}" name="图表 {docpr_id}"/>'
        "<wp:cNvGraphicFramePr/>"
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        f'xmlns:r="{R}" r:id="{rid}"/>'
        "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
    )
    return etree.fromstring(xml.encode("utf-8"))


def _replace_anchor(document_xml: bytes, anchor: str, rid: str, docpr_id: int,
                    cx: int, cy: int) -> bytes:
    """把写锚点的那个段里的锚点文字换成图表绘图。段落的 pPr（居中/段前 12pt/段后 3pt）原样保留。"""
    root = etree.fromstring(document_xml)
    for paragraph in root.iter(f"{{{W}}}p"):
        text = "".join(node.text or "" for node in paragraph.iter(f"{{{W}}}t"))
        if text != anchor:
            continue
        for child in list(paragraph):
            if child.tag in (f"{{{W}}}r", f"{{{W}}}hyperlink", f"{{{W}}}bookmarkStart",
                             f"{{{W}}}bookmarkEnd", f"{{{W}}}proofErr", f"{{{W}}}smartTag"):
                paragraph.remove(child)
        paragraph.append(_drawing_run(rid, docpr_id, cx, cy))
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    raise OoxmlChartError(f"文档里找不到图表锚点 {anchor!r}，图表顺序已与文档脱节")


def _add_relationship(rels_xml: bytes, rid: str, reltype: str, target: str) -> bytes:
    root = etree.fromstring(rels_xml)
    node = etree.SubElement(root, f"{{{PR_NS}}}Relationship")
    node.set("Id", rid)
    node.set("Type", reltype)
    node.set("Target", target)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _add_content_type_overrides(content_types_xml: bytes, parts: list[str]) -> bytes:
    root = etree.fromstring(content_types_xml)
    if not root.xpath('//ct:Default[@Extension="xlsx"]',
                      namespaces={"ct": CT_NS}):
        default = etree.Element(f"{{{CT_NS}}}Default")
        default.set("Extension", "xlsx")
        default.set("ContentType", CT_XLSX)
        root.insert(0, default)
    for name, content_type in parts:
        override = etree.SubElement(root, f"{{{CT_NS}}}Override")
        override.set("PartName", name)
        override.set("ContentType", content_type)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _next_ids(document_xml: bytes, rels_xml: bytes) -> tuple[int, int]:
    """本次文档里下一个可用的 docPr ID 与 rId（图片/旧关系都占号，必须错开）。"""
    root = etree.fromstring(document_xml)
    docpr_ids = [int(node.get("id")) for node in root.iter(f"{{{WP}}}docPr")
                 if (node.get("id") or "").isdigit()]
    rels = etree.fromstring(rels_xml)
    used = []
    for node in rels:
        match = re.fullmatch(r"rId(\d+)", node.get("Id") or "")
        if match:
            used.append(int(match.group(1)))
    return (max(docpr_ids) + 1 if docpr_ids else 1,
            max(used) + 1 if used else 1)


# ---------------------------------------------------------------------------
# 主流程：注入
# ---------------------------------------------------------------------------
def inject_charts(docx_path, specs, width_cm: float = 15.0, height_cm: float = 7.5,
                  log=print) -> int:
    """把 ``specs`` 里的原生图表注入**已落盘**的 docx，就地替换。返回注入张数。

    ``specs`` 顺序 = 文档里锚点顺序。内部先写同目录临时文件并全量校验，通过才覆盖原文件；
    失败抛 :class:`OoxmlChartError`，原文件保持不变。
    """
    specs = [spec for spec in specs if spec is not None]
    if not specs:
        return 0
    for spec in specs:
        if not spec.anchor:
            raise OoxmlChartError("ChartSpec 缺 anchor，_figure 侧没登记成功")
        if not spec.categories or not spec.series:
            raise OoxmlChartError(f"图表 {spec.anchor!r} 没有类别或系列，拒绝写入空图")
    styles = _seed_styles()
    docx_path = Path(docx_path)
    cx, cy = int(round(width_cm * EMU_PER_CM)), int(round(height_cm * EMU_PER_CM))

    with zipfile.ZipFile(docx_path) as archive:
        order = archive.namelist()
        parts = {name: archive.read(name) for name in order}

    document_xml = parts["word/document.xml"]
    rels_xml = parts["word/_rels/document.xml.rels"]
    docpr_id, rid = _next_ids(document_xml, rels_xml)

    chart_names, overrides = [], []
    for index, spec in enumerate(specs, 1):
        chart_part = f"word/charts/chart{index}.xml"
        rid_now = f"rId{rid}"
        document_xml = _replace_anchor(document_xml, spec.anchor, rid_now, docpr_id, cx, cy)
        rels_xml = _add_relationship(rels_xml, rid_now, RT_CHART, f"charts/chart{index}.xml")
        parts[chart_part] = build_chart_xml(spec)
        parts[f"word/charts/_rels/chart{index}.xml.rels"] = _chart_rels(index)
        chart_names.append(chart_part)
        overrides.append((f"/{chart_part}", CT_CHART))
        docpr_id += 1
        rid += 1
        log(f"[原生图表 OOXML] {spec.anchor} → {chart_part} "
            f"（{len(spec.categories)} 类目 / {len(spec.series)} 系列，r:id={rid_now}）")

    # 样式部件**每图一份**（2026-09-28 实测：让多张图表的 rels 指向同一个 style1.xml/
    # colors1.xml，Word 打开 2 张图以上的文档直接报「在试图打开文件时遇到错误」；
    # 每图各给一份立刻能开。OPC 规范允许多个关系指向同一部件，但 Word 的图表加载器不认，
    # 所以这里宁可多存 ~10 KB/图也不能共享。）
    for index in range(1, len(specs) + 1):
        parts[f"word/charts/style{index}.xml"] = styles["style"]
        parts[f"word/charts/colors{index}.xml"] = styles["colors"]
        overrides.append((f"/word/charts/style{index}.xml", CT_CHART_STYLE))
        overrides.append((f"/word/charts/colors{index}.xml", CT_CHART_COLORS))
    for index in range(1, len(specs) + 1):
        parts[f"word/embeddings/gd_chart{index}.xlsx"] = build_workbook(specs[index - 1])

    parts["word/document.xml"] = document_xml
    parts["word/_rels/document.xml.rels"] = rels_xml
    parts["[Content_Types].xml"] = _add_content_type_overrides(
        parts["[Content_Types].xml"], overrides)

    temp_path = docx_path.with_name(docx_path.stem + ".gdchart.tmp.docx")
    try:
        with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as out:
            for name in order:                      # 保持原有条目顺序（含 [Content_Types].xml 在前）
                out.writestr(name, parts[name])
            for name, blob in parts.items():
                if name not in order:
                    out.writestr(name, blob)
        verify_docx_charts(temp_path, specs)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise
    os.replace(temp_path, docx_path)                 # 同卷 → 原子替换
    log(f"[原生图表 OOXML] 已注入 {len(specs)} 张原生图表 → {docx_path}")
    return len(specs)


def save_with_charts(doc, path, specs, width_cm: float = 15.0, height_cm: float = 7.5,
                     log=print):
    """落盘 + 注入 + 校验 + 原子替换，一步到位。``path`` 要么是完整成品，要么原封不动。"""
    specs = [spec for spec in specs if spec is not None]
    path = Path(path)
    if not specs:
        doc.save(path)
        return path
    path.parent.mkdir(parents=True, exist_ok=True)
    staging = path.with_name(path.stem + ".gdchart-stage.docx")
    try:
        doc.save(staging)
        inject_charts(staging, specs, width_cm=width_cm, height_cm=height_cm, log=log)
        os.replace(staging, path)
    except Exception:
        staging.unlink(missing_ok=True)
        raise
    return path


# ---------------------------------------------------------------------------
# 校验（原子交付的闸）
# ---------------------------------------------------------------------------
def _cache_of(chart_xml: bytes) -> dict:
    """解析一张 c:chart 部件：类别、系列名、值、ptCount、公式。"""
    cns = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
    root = etree.fromstring(chart_xml)
    sers = root.xpath("//c:ser", namespaces=cns)
    series = []
    for ser in sers:
        name = ser.xpath("string(c:tx//c:v)", namespaces=cns)
        cat_formula = ser.xpath("string(c:cat//c:f)", namespaces=cns)
        cat_count = int(ser.xpath("string(c:cat//c:ptCount/@val)", namespaces=cns) or 0)
        cats = [(int(pt.get("idx")), pt.xpath("string(c:v)", namespaces=cns))
                for pt in ser.xpath("c:cat//c:pt", namespaces=cns)]
        val_formula = ser.xpath("string(c:val//c:f)", namespaces=cns)
        val_count = int(ser.xpath("string(c:val//c:ptCount/@val)", namespaces=cns) or 0)
        pts = [(int(pt.get("idx")), pt.xpath("string(c:v)", namespaces=cns))
               for pt in ser.xpath("c:val//c:pt", namespaces=cns)]
        series.append({"name": name, "cat_formula": cat_formula, "cat_count": cat_count,
                       "cats": cats, "val_formula": val_formula, "val_count": val_count,
                       "pts": pts})
    return {"series": series,
            "external": root.xpath("string(//c:externalData/@r:id)",
                                   namespaces={**cns,
                                               "r": R})}


def _workbook_grid(blob: bytes) -> list[list]:
    from openpyxl import load_workbook
    book = load_workbook(io.BytesIO(blob))
    sheet = book.worksheets[0]
    grid = [[cell.value for cell in row] for row in sheet.iter_rows()]
    while grid and all(cell is None for cell in grid[-1]):
        grid.pop()
    return grid


def verify_docx_charts(docx_path, specs) -> None:
    """逐图强断言：部件/关系/内嵌工作簿/缓存/表值/无锚点残留，任一不符即抛。

    校验三件事是同一份数据的三种表述：``c:numCache``、内嵌 xlsx、:class:`ChartSpec`
    （而 ChartSpec 本身就来自①/②表的同一批率值，见 ``_g03_chart``）。
    """
    specs = [spec for spec in specs if spec is not None]
    if not specs:
        raise OoxmlChartError("没有待校验的图表")
    cns = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
           "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
           "w": W, "wp": WP, "r": R}
    with zipfile.ZipFile(docx_path) as archive:
        names = set(archive.namelist())
        charts = sorted(n for n in names if re.fullmatch(r"word/charts/chart\d+\.xml", n))
        # 只数**本次注入的**内嵌工作簿：模板本身就带 18 个
        # `Microsoft_Excel_Worksheet*.xlsx`（模板自带的旧图表），那些不是本次产物。
        embeds = sorted(n for n in names
                        if re.fullmatch(r"word/embeddings/gd_chart\d+\.xlsx", n))
        if len(charts) != len(specs):
            raise OoxmlChartError(f"图表部件数不对：{len(charts)} != {len(specs)} 张（{charts}）")
        if len(embeds) != len(specs):
            raise OoxmlChartError(f"内嵌工作簿数不对：{len(embeds)} != {len(specs)}（{embeds}）")
        for index in range(1, len(specs) + 1):
            for required in (f"word/charts/chart{index}.xml",
                             f"word/charts/_rels/chart{index}.xml.rels",
                             f"word/embeddings/gd_chart{index}.xlsx",
                             # 每图一份样式部件：Word 不接受多图共用（见 inject_charts 注释）
                             f"word/charts/style{index}.xml",
                             f"word/charts/colors{index}.xml"):
                if required not in names:
                    raise OoxmlChartError(f"缺部件 {required}")
        # 「统计图不得以 PNG 入库」的正确判据不是「media/ 必须为空」——报告里本来就
        # 有照片（brief-v4：其他照片不动），media/ 非空是正常的。真正的判据是：
        # 每一张 ChartSpec 都必须落成 c:chart 锚点（而不是 a:blip 图片锚点），
        # 且 :func:`inject_charts` 全程不往 word/media/ 写任何东西。
        document = etree.fromstring(archive.read("word/document.xml"))
        rels = etree.fromstring(archive.read("word/_rels/document.xml.rels"))
        content_types = etree.fromstring(archive.read("[Content_Types].xml"))

        rid_target = {node.get("Id"): node.get("Target")
                      for node in rels if node.get("Type") == RT_CHART}
        if len(rid_target) != len(specs):
            raise OoxmlChartError(f"document.xml.rels 里的图表关系数不对：{len(rid_target)}")
        used_rids = [node.get(f"{{{R}}}id")
                     for node in document.xpath("//c:chart", namespaces=cns)]
        if sorted(used_rids) != sorted(rid_target):
            raise OoxmlChartError(f"锚点 rId 与关系对不上：{used_rids} vs {list(rid_target)}")
        # 统计图必须全部是 c:chart 锚点。正文里的 a:blip 是照片（brief-v4 要求保留），
        # 数量不受本次注入影响——注入函数压根不往 word/media/ 写东西。
        chart_anchors = int(document.xpath("count(//c:chart)", namespaces=cns))
        if chart_anchors != len(specs):
            raise OoxmlChartError(f"c:chart 锚点数不对：{chart_anchors} != {len(specs)}")
        docpr_ids = [node.get("id") for node in document.xpath("//wp:docPr", namespaces=cns)]
        if len(set(docpr_ids)) != len(docpr_ids):
            raise OoxmlChartError(f"docPr ID 重复：{docpr_ids}")
        if len(used_rids) != len(specs):
            raise OoxmlChartError(f"正文里的图表锚点数不对：{len(used_rids)} != {len(specs)}")
        leftovers = [t for t in document.xpath("//w:t/text()", namespaces=cns)
                     if "@@GDCHART" in t]
        if leftovers:
            raise OoxmlChartError(f"锚点文字残留：{leftovers}")

        overrides = {node.get("PartName") for node in content_types
                     if node.tag.endswith("Override")}
        for index in range(1, len(specs) + 1):
            chart_part = f"word/charts/chart{index}.xml"
            embed_part = f"word/embeddings/gd_chart{index}.xlsx"
            if f"/{chart_part}" not in overrides:
                raise OoxmlChartError(f"[Content_Types].xml 缺 {chart_part} 的 Override")
            rels_part = f"word/charts/_rels/chart{index}.xml.rels"
            if rels_part not in names:
                raise OoxmlChartError(f"缺图表关系部件 {rels_part}")
            chart_rels = etree.fromstring(archive.read(rels_part))
            book_targets = [n.get("Target") for n in chart_rels if n.get("Type") == RT_PACKAGE]
            if book_targets != [f"../embeddings/gd_chart{index}.xlsx"]:
                raise OoxmlChartError(
                    f"{rels_part} 的内嵌工作簿引用不对：{book_targets}（不得指向外部文件）")
            if any(node.get("TargetMode") == "External" for node in chart_rels):
                raise OoxmlChartError(f"{rels_part} 出现外部引用，图表不得链接外部 xlsx")
            # 样式部件必须指向本图专属的那份。共用一份时 Word 打不开整个文档（实测 N≥2）。
            style_targets = {n.get("Target") for n in chart_rels
                             if n.get("Type") in (RT_CHART_STYLE, RT_CHART_COLORS)}
            if style_targets != {f"style{index}.xml", f"colors{index}.xml"}:
                raise OoxmlChartError(
                    f"{rels_part} 的样式部件引用不对：{sorted(style_targets)}，"
                    f"应为 style{index}.xml / colors{index}.xml（多图共用会导致 Word 打不开）")
            spec = specs[index - 1]
            cache = _cache_of(archive.read(chart_part))
            _assert_one_chart(index, spec, cache)
            grid = _workbook_grid(archive.read(embed_part))
            _assert_workbook(index, spec, grid)


def _assert_one_chart(index: int, spec: ChartSpec, cache: dict) -> None:
    where = f"第 {index} 张图（{spec.anchor}）"
    if cache["external"] != "rId3":
        raise OoxmlChartError(f"{where} c:externalData 指向 {cache['external']!r}，应为 rId3")
    if len(cache["series"]) != len(spec.series):
        raise OoxmlChartError(f"{where} 系列数不符：XML {len(cache['series'])} "
                              f"!= spec {len(spec.series)}")
    last_row = 1 + len(spec.categories)
    expected_cats = [str(cat) for cat in spec.categories]
    for j, (series, (name, values, _color)) in enumerate(zip(cache["series"], spec.series)):
        label = f"{where} 系列 {j + 1}（{name}）"
        if series["name"] != str(name):
            raise OoxmlChartError(f"{label} 系列名不符：XML {series['name']!r}")
        if series["cat_count"] != len(expected_cats):
            raise OoxmlChartError(f"{label} 类别 ptCount={series['cat_count']} "
                                  f"!= 类目数 {len(expected_cats)}")
        if [c for _, c in series["cats"]] != expected_cats:
            raise OoxmlChartError(f"{label} 类别缓存与 spec 不一致："
                                  f"{[c for _, c in series['cats']]} != {expected_cats}")
        if [i for i, _ in series["cats"]] != list(range(len(expected_cats))):
            raise OoxmlChartError(f"{label} 类别 pt idx 不连续：{[i for i, _ in series['cats']]}")
        column = _col_letter(2 + j)
        want_cat_f = f"Sheet1!$A$2:$A${last_row}"
        want_val_f = f"Sheet1!${column}$2:${column}${last_row}"
        if series["cat_formula"] != want_cat_f or series["val_formula"] != want_val_f:
            raise OoxmlChartError(
                f"{label} SERIES 公式未指向实际区域：cat={series['cat_formula']!r} "
                f"val={series['val_formula']!r}，应为 {want_cat_f!r} / {want_val_f!r}")
        if series["val_count"] != len(values):
            raise OoxmlChartError(f"{label} 数值 ptCount={series['val_count']} != 点数 {len(values)}")
        cached = dict(series["pts"])
        for i, value in enumerate(values):
            if value is None:
                if i in cached:
                    raise OoxmlChartError(f"{label} 第 {i} 点应为断点却有残值 {cached[i]!r}")
                continue
            if i not in cached:
                raise OoxmlChartError(f"{label} 第 {i} 点缺缓存值（spec={value}）")
            if abs(float(cached[i]) - float(value)) > 1e-6:
                raise OoxmlChartError(
                    f"{label} 第 {i} 点缓存 {cached[i]!r} != spec {value!r}")


def _assert_workbook(index: int, spec: ChartSpec, grid: list[list]) -> None:
    where = f"第 {index} 张图内嵌工作簿（{spec.anchor}）"
    width = 1 + len(spec.series)
    if len(grid) != 1 + len(spec.categories) or any(len(row) > width for row in grid):
        raise OoxmlChartError(f"{where} 区域尺寸不对：{len(grid)} 行 × "
                              f"{max((len(r) for r in grid), default=0)} 列，"
                              f"应为 {1 + len(spec.categories)} 行 × {width} 列")
    header = [str(cell) for cell in grid[0][:width]]
    want_header = ["类别"] + [str(name) for name, _v, _c in spec.series]
    if header != want_header:
        raise OoxmlChartError(f"{where} 表头不对：{header} != {want_header}")
    for i, category in enumerate(spec.categories):
        row = grid[1 + i]
        if str(row[0]) != str(category):
            raise OoxmlChartError(f"{where} 第 {i + 1} 行类别 {row[0]!r} != {category!r}")
        for j, (_name, values, _color) in enumerate(spec.series):
            cell = row[1 + j] if 1 + j < len(row) else None
            want = values[i]
            if want is None:
                if cell is not None:
                    raise OoxmlChartError(f"{where} 第 {i + 1} 行第 {j + 1} 列应留空，"
                                          f"实为 {cell!r}（会画成假数据）")
                continue
            if cell is None or abs(float(cell) - float(want)) > 1e-6:
                raise OoxmlChartError(
                    f"{where} 第 {i + 1} 行第 {j + 1} 列 {cell!r} != spec {want!r}")
