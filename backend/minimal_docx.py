"""Programmatic DOCX report generation driven by a Markdown skeleton template.

A Markdown template (templates/重庆项目报告模板.md) declares the report
skeleton (headings, static body texts, tables). Dynamic data sections
(overview table, tci/height/bolt statistics, charts, conclusions) are injected
ONLY at explicit anchors <!-- inject:... -->; headings never trigger injection
(Q2: 2.3.x 等静态标题不得被误作锚点吞掉). When an anchor is absent the section
is appended at the end in canonical chapter order. Requires a Markdown skeleton.
"""

from __future__ import annotations

import math
import re
import xml.etree.ElementTree as ET
from contextlib import contextmanager
from contextvars import ContextVar
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from backend import markdown_skeleton
from backend import report_engine as engine

BODY_FONT = "仿宋_GB2312"
HEADING_FONT = "黑体"
LATIN_FONT = "Times New Roman"
BODY_SIZE = Pt(10.5)

_FORMAT_CONFIG = ContextVar("report_format_config", default=None)


def _current_format_config(config=None):
    return config or _FORMAT_CONFIG.get() or markdown_skeleton.DEFAULT_CONFIG


@contextmanager
def format_context(config):
    token = _FORMAT_CONFIG.set(config)
    try:
        yield
    finally:
        _FORMAT_CONFIG.reset(token)


def _alignment(value):
    return {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }[value]


def _set_rfonts(r_pr, east_asia, latin):
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    r_fonts.set(qn("w:eastAsia"), east_asia)
    r_fonts.set(qn("w:ascii"), latin)
    r_fonts.set(qn("w:hAnsi"), latin)


def _set_style_font(style, style_config):
    style.font.name = style_config["latin"]
    style.font.size = Pt(float(style_config["size_pt"]))
    style.font.bold = style_config.get("bold", False)
    _set_rfonts(style._element.get_or_add_rPr(), style_config["east_asia"], style_config["latin"])


def configure_document(doc, config=None):
    config = _current_format_config(config)
    page = config["page"]
    section = doc.sections[0]
    landscape = page["orientation"] == "landscape"
    section.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
    width, height = (29.7, 21) if landscape else (21, 29.7)
    section.page_width = Cm(width)
    section.page_height = Cm(height)
    section.left_margin = Cm(page["left_margin_cm"])
    section.right_margin = Cm(page["right_margin_cm"])
    section.top_margin = Cm(page["top_margin_cm"])
    section.bottom_margin = Cm(page["bottom_margin_cm"])

    normal = doc.styles["Normal"]
    _set_style_font(normal, config["body"])
    normal.paragraph_format.alignment = _alignment(config["body"]["alignment"])
    normal.paragraph_format.line_spacing = config["body"]["line_spacing"]
    for level in range(1, 6):
        style = doc.styles[f"Heading {level}"]
        heading_config = config["heading"][str(level)]
        _set_style_font(style, heading_config)
        style.paragraph_format.alignment = _alignment(heading_config["alignment"])
        style.paragraph_format.space_before = Pt(heading_config["space_before_pt"])
        style.paragraph_format.space_after = Pt(heading_config["space_after_pt"])
        if "line_spacing" in heading_config:
            style.paragraph_format.line_spacing = heading_config["line_spacing"]
        style.paragraph_format.keep_with_next = heading_config["keep_with_next"]
        style.paragraph_format.page_break_before = heading_config["page_break_before"]


def _apply_paragraph(paragraph, style_config, *, clear_indent=False):
    fmt = paragraph.paragraph_format
    fmt.alignment = _alignment(style_config["alignment"])
    if "space_before_pt" in style_config:
        fmt.space_before = Pt(style_config["space_before_pt"])
    if "space_after_pt" in style_config:
        fmt.space_after = Pt(style_config["space_after_pt"])
    if "line_spacing" in style_config:
        fmt.line_spacing = style_config["line_spacing"]
    if "first_line_chars" in style_config:
        # first_line_chars is represented as a fixed point-size indent (not w:firstLineChars).
        fmt.first_line_indent = Pt(float(style_config["size_pt"]) * style_config["first_line_chars"])
    if "keep_with_next" in style_config:
        fmt.keep_with_next = style_config["keep_with_next"]
    if "page_break_before" in style_config:
        fmt.page_break_before = style_config["page_break_before"]
    if clear_indent:
        _clear_indent(paragraph)


def _apply_run(run, style_config, *, bold=None):
    if bold is not None:
        run.bold = bold
    elif "bold" in style_config:
        run.bold = style_config["bold"]
    run.italic = False
    run.font.size = Pt(float(style_config["size_pt"]))
    _set_run_fonts(run, style_config["east_asia"], style_config["latin"])


def _set_run_fonts(run, east_asia=BODY_FONT, latin=LATIN_FONT):
    run.font.name = latin
    run.font.color.rgb = RGBColor(0, 0, 0)
    _set_rfonts(run._element.get_or_add_rPr(), east_asia, latin)


def _clear_indent(paragraph):
    fmt = paragraph.paragraph_format
    fmt.left_indent = Pt(0)
    fmt.right_indent = Pt(0)
    fmt.first_line_indent = Pt(0)
    p_pr = paragraph._p.get_or_add_pPr()
    ind = p_pr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        p_pr.append(ind)
    for key in ("left", "right", "firstLine", "hanging", "leftChars", "rightChars", "firstLineChars", "hangingChars"):
        ind.set(qn(f"w:{key}"), "0")


def _warn(paragraph, level):
    p_pr = paragraph._p.get_or_add_pPr()
    outline = p_pr.find(qn("w:outlineLvl"))
    if outline is None:
        outline = OxmlElement("w:outlineLvl")
        p_pr.append(outline)
    outline.set(qn("w:val"), str(level - 1))


_HEADING_NUMBER_FORMATS = {1: "%1.", 2: "%1.%2", 3: "%1.%2.%3"}
_LEADING_NUMBER_RE = re.compile(r"^\d+(?:\.\d+)*\s+")


def _ensure_heading_numbering(doc):
    """一/二/三级标题多级列表（1./1.1/1.1.1），与 Heading 1-3 样式绑定。"""
    # ponytail: 单个共享 numId；下级随上级递增自动重启（Word 默认行为，不显式写 lvlRestart）
    cached = getattr(doc, "_heading_number_id", None)
    if cached is not None:
        return cached
    numbering = doc.part.numbering_part._element
    used_abstract = [int(node.get(qn("w:abstractNumId"))) for node in numbering.findall(qn("w:abstractNum"))]
    used_num = [int(node.get(qn("w:numId"))) for node in numbering.findall(qn("w:num"))]
    abstract_id = max(used_abstract, default=-1) + 1
    num_id = max(used_num, default=0) + 1
    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "multilevel")
    abstract.append(multi)
    for ilvl, level in enumerate((1, 2, 3)):
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), str(ilvl))
        start = OxmlElement("w:start")
        start.set(qn("w:val"), "1")
        lvl.append(start)
        fmt = OxmlElement("w:numFmt")
        fmt.set(qn("w:val"), "decimal")
        lvl.append(fmt)
        style = OxmlElement("w:pStyle")
        style.set(qn("w:val"), f"Heading{level}")
        lvl.append(style)
        suff = OxmlElement("w:suff")
        suff.set(qn("w:val"), "space")
        lvl.append(suff)
        text = OxmlElement("w:lvlText")
        text.set(qn("w:val"), _HEADING_NUMBER_FORMATS[level])
        lvl.append(text)
        align = OxmlElement("w:lvlJc")
        align.set(qn("w:val"), "left")
        lvl.append(align)
        p_pr = OxmlElement("w:pPr")
        indent = OxmlElement("w:ind")
        indent.set(qn("w:left"), "0")
        indent.set(qn("w:hanging"), "0")
        p_pr.append(indent)
        lvl.append(p_pr)
        # T9：编号符号直接染黑（Word 缺省继承蓝色时 1./1.1/1.1.1 渲染蓝）；标题汉字 run 本身已黑，不动。
        r_pr = OxmlElement("w:rPr")
        color = OxmlElement("w:color")
        color.set(qn("w:val"), "000000")
        r_pr.append(color)
        lvl.append(r_pr)
        abstract.append(lvl)
    # numbering.xml 要求 abstractNum 位于 num 之前
    first_num = next((i for i, child in enumerate(numbering) if child.tag == qn("w:num")), len(numbering))
    numbering.insert(first_num, abstract)
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    ref = OxmlElement("w:abstractNumId")
    ref.set(qn("w:val"), str(abstract_id))
    num.append(ref)
    numbering.append(num)
    doc._heading_number_id = num_id
    return num_id


def _heading(doc, text, level):
    level = max(1, min(int(level), 5))
    config = _current_format_config()
    style_config = config["heading"][str(level)]
    paragraph = doc.add_paragraph(style=f"Heading {level}")
    _apply_paragraph(paragraph, style_config, clear_indent=True)
    _warn(paragraph, style_config["outline_level"] + 1)
    if level <= 3:
        num_id = _ensure_heading_numbering(doc)
        p_pr = paragraph._p.get_or_add_pPr()
        num_pr = OxmlElement("w:numPr")
        ilvl = OxmlElement("w:ilvl")
        ilvl.set(qn("w:val"), str(level - 1))
        num_pr.append(ilvl)
        num_el = OxmlElement("w:numId")
        num_el.set(qn("w:val"), str(num_id))
        num_pr.append(num_el)
        p_pr.append(num_pr)
        text = _LEADING_NUMBER_RE.sub("", str(text), count=1)
    run = paragraph.add_run(text)
    _apply_run(run, style_config)
    # 目录收集：渲染顺序=文档顺序，编号与多级列表一致
    toc_max = int(_current_format_config()["toc"]["max_level"])
    if level <= toc_max:
        counters = getattr(doc, "_toc_counters", None)
        if counters is None:
            counters = [0, 0, 0, 0, 0]
            doc._toc_counters = counters
        counters[level - 1] += 1
        for deeper in range(level, 5):
            counters[deeper] = 0
        number = ".".join(str(c) for c in counters[:level])
        entries = getattr(doc, "_toc_entries", None)
        if entries is None:
            entries = []
            doc._toc_entries = entries
        entries.append((number, str(text)))
    return paragraph


def _body(doc, text):
    style_config = _current_format_config()["body"]
    paragraph = doc.add_paragraph()
    _apply_paragraph(paragraph, style_config)
    run = paragraph.add_run(text)
    _apply_run(run, style_config)
    return paragraph


def _rate_text(value):
    return "" if value is None else f"{value:.2f}"


def _caption(doc, prefix, number, title):
    style_config = _current_format_config()["caption"]
    paragraph = doc.add_paragraph()
    _apply_paragraph(paragraph, style_config, clear_indent=True)
    run = paragraph.add_run(f"{prefix}{number} {title}")
    _apply_run(run, style_config)
    return paragraph


def _caption_text(doc, text):
    """整段表题/图题文本（如模板中的“表2.1-1 人员组织情况”）按 caption 样式渲染。"""
    style_config = _current_format_config()["caption"]
    paragraph = doc.add_paragraph()
    _apply_paragraph(paragraph, style_config, clear_indent=True)
    run = paragraph.add_run(text)
    _apply_run(run, style_config)
    return paragraph


def add_toc(doc, toc_config=None):
    """Insert a configured TOC field, or remove TOC update state when disabled."""
    config = _current_format_config()
    toc_config = toc_config or config["toc"]
    settings = doc.settings._element
    if not toc_config["enabled"]:
        update_fields = settings.find(qn("w:updateFields"))
        if update_fields is not None:
            settings.remove(update_fields)
        return

    title = doc.add_paragraph(style="TOC Heading")
    title_config = dict(config["heading"]["1"])
    title_config.update({"alignment": "center", "bold": True})
    if toc_config.get("east_asia"):
        title_config["east_asia"] = toc_config["east_asia"]
    _apply_paragraph(title, title_config, clear_indent=True)
    title_run = title.add_run(toc_config["title"])
    _apply_run(title_run, title_config)

    paragraph = doc.add_paragraph(style="TOC Heading")
    _clear_indent(paragraph)
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instruction.text = (
        f" TOC \\o \"{toc_config['min_level']}-{toc_config['max_level']}\""
        " \\h \\z \\u "
    )
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend((begin, instruction, separate, end))

    update_fields = settings.find(qn("w:updateFields"))
    if toc_config["update_on_open"]:
        if update_fields is None:
            update_fields = OxmlElement("w:updateFields")
            settings.append(update_fields)
        update_fields.set(qn("w:val"), "true")
    elif update_fields is not None:
        settings.remove(update_fields)


def _add_toc(doc, toc_config=None):
    return add_toc(doc, toc_config)


def _populate_toc_cache(doc):
    """把收集到的标题条目写入 TOC 域的缓存结果（separate 与 end 之间）。

    updateFields 关闭（防 WPS/Word 弹"域可能引用其他文件"）时，Word 不会自动更新
    目录；预填充缓存让目录打开即显示。页码写占位 0，交付前可用 Word 更新目录固化
    （见 work/finalize_toc.py）。
    """
    entries = getattr(doc, "_toc_entries", None)
    if not entries:
        return
    field_run = None
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            instrs = run._r.findall(qn("w:instrText"))
            if any((t.text or "").strip().startswith("TOC") for t in instrs):
                field_run = run._r
                break
        if field_run is not None:
            break
    if field_run is None:
        return
    children = list(field_run)
    sep_idx = end_idx = None
    for index, child in enumerate(children):
        if child.tag == qn("w:fldChar") and child.get(qn("w:fldCharType")) == "separate":
            sep_idx = index
        elif child.tag == qn("w:fldChar") and child.get(qn("w:fldCharType")) == "end":
            end_idx = index
    if sep_idx is None or end_idx is None or end_idx <= sep_idx:
        return
    end_element = children[end_idx]
    body_font = _current_format_config()["body"]

    def _cache_run(text: str) -> OxmlElement:
        r = OxmlElement("w:r")
        r_pr = OxmlElement("w:rPr")
        r_fonts = OxmlElement("w:rFonts")
        r_fonts.set(qn("w:ascii"), "Times New Roman")
        r_fonts.set(qn("w:hAnsi"), "Times New Roman")
        r_fonts.set(qn("w:eastAsia"), body_font["east_asia"])
        r_pr.append(r_fonts)
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), "21")
        r_pr.append(sz)
        r.append(r_pr)
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = text
        r.append(t)
        return r

    # 缓存必须是段落的同级run，不能在w:r内嵌套w:r；每条目录单独成段。
    field_run.remove(end_element)
    previous = field_run.getparent()
    content_width = doc.sections[-1].page_width - doc.sections[-1].left_margin - doc.sections[-1].right_margin
    for number, text in entries:
        paragraph = OxmlElement("w:p")
        properties = OxmlElement("w:pPr")
        outline = OxmlElement("w:outlineLvl")
        outline.set(qn("w:val"), "9")
        properties.append(outline)
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:after"), "0")
        spacing.set(qn("w:line"), "240")
        spacing.set(qn("w:lineRule"), "auto")
        properties.append(spacing)
        tabs = OxmlElement("w:tabs")
        tab_stop = OxmlElement("w:tab")
        tab_stop.set(qn("w:val"), "right")
        tab_stop.set(qn("w:leader"), "dot")
        tab_stop.set(qn("w:pos"), str(round(content_width / 635)))
        tabs.append(tab_stop)
        properties.append(tabs)
        paragraph.append(properties)
        paragraph.append(_cache_run(f"{number} {text}".strip()))
        tab_run = OxmlElement("w:r")
        tab_run.append(OxmlElement("w:tab"))
        paragraph.append(tab_run)
        paragraph.append(_cache_run("0"))
        previous.addnext(paragraph)
        previous = paragraph
    end_run = OxmlElement("w:r")
    end_run.append(end_element)
    previous.append(end_run)


def _report_number_from_cover(spec_text: str) -> str:
    """封面“报告编号：BG-2026-”→页眉右侧变量“BG-2026-”（去“报告编号：”标签前缀）。"""
    text = str(spec_text or "").strip()
    for sep in ("：", ":"):
        label, found, value = text.partition(sep)
        if found and "报告编号" in label:
            return value.strip()
    return text


_MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _m_elem(tag, text=None):
    element = ET.Element("{%s}%s" % (_MATH_NS, tag))
    if text is not None:
        run = ET.SubElement(element, "{%s}r" % _MATH_NS)
        t = ET.SubElement(run, "{%s}t" % _MATH_NS)
        t.text = text
    return element


def _math_run(text):
    """创建直接的 m:r/m:t，避免生成无效的嵌套 m:r。"""
    run = ET.Element("{%s}r" % _MATH_NS)
    t = ET.SubElement(run, "{%s}t" % _MATH_NS)
    t.text = text
    return run


def _tci_formula_omath():
    """附件模板 2.3.1 的可编辑 TCI 公式（不重复输出纯文本公式）。"""
    omath = _m_elem("oMath")
    omath.append(_math_run("TCI="))
    nary = ET.SubElement(omath, "{%s}nary" % _MATH_NS)
    nary_pr = ET.SubElement(nary, "{%s}naryPr" % _MATH_NS)
    chr_el = ET.SubElement(nary_pr, "{%s}chr" % _MATH_NS)
    chr_el.set("{%s}val" % _MATH_NS, "∑")
    lim_loc = ET.SubElement(nary_pr, "{%s}limLoc" % _MATH_NS)
    lim_loc.set("{%s}val" % _MATH_NS, "undOvr")
    nary.append(_m_elem("sub", "i=1"))
    upper = ET.SubElement(nary, "{%s}sup" % _MATH_NS)
    i_zero = ET.SubElement(upper, "{%s}sSub" % _MATH_NS)
    i_zero.append(_m_elem("e", "i"))
    i_zero.append(_m_elem("sub", "0"))
    term = ET.SubElement(nary, "{%s}e" % _MATH_NS)
    w_sub = ET.SubElement(term, "{%s}sSub" % _MATH_NS)
    w_sub.append(_m_elem("e", "w"))
    w_sub.append(_m_elem("sub", "i"))
    omath.append(_math_run("(100−"))
    gd_sub = ET.SubElement(omath, "{%s}sSub" % _MATH_NS)
    gd_sub.append(_m_elem("e", "GD"))
    gd_sub.append(_m_elem("sub", "iTCI"))
    omath.append(_math_run(")/0.7"))
    return omath


def _subscript_omath(base, subscript):
    """生成可编辑下标公式，如 GD_iTCI、w_i、i_0。"""
    omath = _m_elem("oMath")
    scripted = ET.SubElement(omath, "{%s}sSub" % _MATH_NS)
    scripted.append(_m_elem("e", base))
    scripted.append(_m_elem("sub", subscript))
    return omath


def _append_formula_paragraph(doc, omath, *, prefix="", suffix="", centered=False):
    """追加可编辑 OMML；可附加与附件模板完全一致的前后文字。"""
    from lxml import etree as lxml_etree

    ET.register_namespace("m", _MATH_NS)
    paragraph = doc.add_paragraph()
    style_config = _current_format_config()["body"]
    _apply_paragraph(paragraph, style_config, clear_indent=centered)
    if centered:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if prefix:
        run = paragraph.add_run(prefix)
        _apply_run(run, style_config)
    paragraph._p.append(lxml_etree.fromstring(ET.tostring(omath)))
    if suffix:
        run = paragraph.add_run(suffix)
        _apply_run(run, style_config)
    return paragraph


def _cover_top_rule(doc) -> None:
    """封面顶部横线（base.pdf 基准）：空段落 + 段落下框线。"""
    paragraph = doc.add_paragraph()
    _apply_paragraph(
        paragraph,
        {**_current_format_config()["body"], "alignment": "center", "first_line_chars": 0},
        clear_indent=True,
    )
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "9")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def _highlight_yellow(run) -> None:
    """run 加黄色高亮（基准封面：报告编号值/项目名称值/左下黄块）。"""
    r_pr = run._element.get_or_add_rPr()
    highlight = OxmlElement("w:highlight")
    highlight.set(qn("w:val"), "yellow")
    r_pr.append(highlight)


def _cover_paragraphs(doc, spec: str) -> str:
    """按参考 Word 模板渲染封面：| 分隔 8 段，含主标题、城市、编号、单位与日期。

    返回报告编号变量（供页眉右侧使用）。版式对齐 base.pdf 基准：顶部横线、
    编号/项目名黄高亮、项目名下划线填空行、公司/日期 16pt 下沉、左下黄块、固定空段数。
    """
    parts = [p.strip() for p in spec.split("|")]
    if len(parts) != 8:
        raise ValueError(f"封面标记需要 8 段（| 分隔）：{spec}")
    body_font = dict(_current_format_config()["body"])
    bold_center = {**body_font, "alignment": "center", "bold": True, "first_line_chars": 0}

    def _para(text, style_config, highlight=False):
        paragraph = doc.add_paragraph()
        _apply_paragraph(paragraph, style_config, clear_indent=True)
        run = paragraph.add_run(text)
        _apply_run(run, style_config)
        if highlight:
            _highlight_yellow(run)
        return paragraph

    def _blank(style_config):
        paragraph = doc.add_paragraph()
        _apply_paragraph(paragraph, style_config, clear_indent=True)
        return paragraph

    def _fill_line(text, style_config, highlight_value=False):
        """“项目名称：<值>”填空行：标签不加下划线，填空值加单下划线（项目名另加黄高亮）。"""
        label, sep, value = str(text).partition("：")
        if not sep:
            label, sep, value = str(text).partition(":")
        paragraph = doc.add_paragraph()
        _apply_paragraph(paragraph, style_config, clear_indent=True)
        head = paragraph.add_run(label + sep if sep else label)
        _apply_run(head, style_config)
        if value.strip():
            from docx.enum.text import WD_UNDERLINE

            filled = paragraph.add_run(value.strip())
            _apply_run(filled, style_config)
            filled.underline = WD_UNDERLINE.SINGLE
            if highlight_value:
                _highlight_yellow(filled)
        return paragraph

    blank_config = {**body_font, "alignment": "center", "first_line_chars": 0}
    # 顶部横线
    _cover_top_rule(doc)
    # 主标题两行：22pt 黑体（中文）/Times New Roman（英文）加粗居中
    title_font = {**bold_center, "east_asia": "黑体", "latin": "Times New Roman", "size_pt": 22, "line_spacing": 1.5}
    _para(parts[0], title_font)
    _para(parts[1], title_font)
    # 城市名：18pt 黑体加粗居中
    _para(parts[2], {**bold_center, "east_asia": "黑体", "size_pt": 18, "line_spacing": 1.5})
    # 报告编号：14pt 宋体加粗，值黄高亮
    song = "宋体"
    label, sep, number_value = parts[3].partition("：")
    if not sep:
        label, sep, number_value = parts[3].partition(":")
    number_para = doc.add_paragraph()
    _apply_paragraph(number_para, {**bold_center, "east_asia": song, "size_pt": 14, "line_spacing": 1.5}, clear_indent=True)
    label_run = number_para.add_run(label + sep if sep else label)
    _apply_run(label_run, {**bold_center, "east_asia": song, "size_pt": 14})
    if number_value.strip():
        value_run = number_para.add_run(number_value.strip())
        _apply_run(value_run, {**bold_center, "east_asia": song, "size_pt": 14})
        _highlight_yellow(value_run)
    # 编号后 10 空段（基准 P7–P16）
    for _ in range(10):
        _blank(blank_config)
    left_font = {**body_font, "bold": True, "east_asia": song, "size_pt": 14, "line_spacing": 1.5, "first_line_chars": 0}
    _fill_line(parts[4], left_font, highlight_value=True)
    _fill_line(parts[5], left_font)
    # 委托单位后 11 空段（基准 P19–P29）
    for _ in range(11):
        _blank(blank_config)
    # 公司与日期：16pt 宋体加粗居中
    company_font = {**bold_center, "east_asia": song, "size_pt": 16}
    _para(parts[6], company_font)
    _para(parts[7], company_font)
    # 左下黄块（基准 P32：黄高亮空 run sz32）
    yellow_block = doc.add_paragraph()
    _apply_paragraph(yellow_block, blank_config, clear_indent=True)
    block_run = yellow_block.add_run("")
    _apply_run(block_run, {**blank_config, "size_pt": 16})
    _highlight_yellow(block_run)
    # 封面独立成页
    from docx.enum.text import WD_BREAK
    doc.paragraphs[-1].add_run().add_break(WD_BREAK.PAGE)
    return _report_number_from_cover(parts[3])


def _notes_paragraphs(doc, spec: str) -> None:
    """注意事项页（参考模板）：标题宋体16居中，条款宋体14，“联系方式”起宋体15。"""
    parts = [p.strip() for p in spec.split("|") if p.strip()]
    if len(parts) < 2:
        raise ValueError(f"注意事项标记至少需要标题和一条内容：{spec}")
    body_font = dict(_current_format_config()["body"])
    song = "宋体"
    contact_size = 15

    def _para(text, style_config):
        paragraph = doc.add_paragraph()
        _apply_paragraph(paragraph, style_config, clear_indent=True)
        run = paragraph.add_run(text)
        _apply_run(run, style_config)
        return paragraph

    _para(parts[0], {**body_font, "east_asia": song, "size_pt": 16, "alignment": "center", "first_line_chars": 0})
    in_contact = False
    for item in parts[1:]:
        if "联系方式" in item:
            in_contact = True
        _para(item, {**body_font, "east_asia": song, "size_pt": contact_size if in_contact else 14, "first_line_chars": 0, "line_spacing": 1.5})
    from docx.enum.text import WD_BREAK
    doc.paragraphs[-1].add_run().add_break(WD_BREAK.PAGE)


def _table(doc, headers, rows, header_shading=False):
    style_config = _current_format_config()["table"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = {
        "left": WD_TABLE_ALIGNMENT.LEFT,
        "center": WD_TABLE_ALIGNMENT.CENTER,
        "right": WD_TABLE_ALIGNMENT.RIGHT,
        "justify": WD_TABLE_ALIGNMENT.CENTER,
    }[style_config["alignment"]]
    table.style = "Table Grid"

    def fill(cell, value, bold=False):
        cell.text = str(value)
        for paragraph in cell.paragraphs:
            _apply_paragraph(paragraph, style_config, clear_indent=True)
            if header_shading and bold:
                shading = OxmlElement("w:shd")
                shading.set(qn("w:val"), "clear")
                shading.set(qn("w:fill"), style_config["header_shading"])
                cell._tc.get_or_add_tcPr().append(shading)
            run = paragraph.runs[0]
            _apply_run(run, style_config, bold=bold and style_config["header_bold"])
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    for index, header in enumerate(headers):
        fill(table.rows[0].cells[index], header, True)
    for row in rows:
        cells = table.add_row().cells
        for index, value in enumerate(row):
            fill(cells[index], value)
    # 复刻附件模板中的纵向合并：只作用于两张静态方法表。
    def _keep_first_paragraph(cell):
        """合并后删除后续空段落，避免合并单元格内出现多余换行。"""
        tc = cell._tc
        paragraphs = tc.findall(qn("w:p"))
        for extra in paragraphs[1:]:
            tc.remove(extra)

    if list(headers) == ["姓名", "组别", "职务", "职责"] and len(table.rows) >= 5:
        for column in (1, 3):
            _keep_first_paragraph(table.cell(2, column).merge(table.cell(4, column)))
    elif list(headers) == ["类型i", "损坏名称", "损坏程度", "计量单位", "单位扣分", "权重wi", "备  注"] and len(table.rows) >= 3:
        for column in (0, 1, 3, 5):
            _keep_first_paragraph(table.cell(1, column).merge(table.cell(2, column)))
    if not style_config["allow_row_break"]:
        for row in table.rows:
            tr_pr = row._tr.get_or_add_trPr()
            cant_split = OxmlElement("w:cantSplit")
            tr_pr.append(cant_split)
    return table


_DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_REPORT_IMAGE_WIDTH_CM = 15.0
_REPORT_IMAGE_HEIGHT_CM = 8.0


def _unlock_aspect(run) -> None:
    """取消图片锁定纵横比（noChangeAspect=0）。"""
    for locks in run._element.iter(qn("a:graphicFrameLocks")):
        locks.set("noChangeAspect", "0")


def _borderless_table(table):
    """移除表格所有框线（示例图/病害图表格无框线）。"""
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        borders.append(el)
    existing = tbl_pr.find(qn("w:tblBorders"))
    if existing is not None:
        tbl_pr.remove(existing)
    tbl_pr.append(borders)
    return table


def _photo_text_table(doc, image_data, extension, text, temp_dir, name):
    """单列表格：行1=15×8cm图片，行2=说明文字。"""
    # 只有"数据/结果表"后需要显式段落防止 Word 自动合并；
    # 两张示例图片表之间不加空段（用户要求示例图连续排版，无空格分隔）。
    body = doc._body._element
    children = list(body)
    previous = None
    if len(children) >= 2 and children[-1].tag == qn("w:sectPr"):
        previous = children[-2]
    elif len(children) == 1 and children[-1].tag != qn("w:sectPr"):
        previous = children[-1]
    if previous is not None and previous.tag == qn("w:tbl"):
        grid_cols = list(previous.iter(qn("w:gridCol")))
        if len(grid_cols) > 1:
            separator = doc.add_paragraph()
            _apply_paragraph(separator, _current_format_config()["body"], clear_indent=True)
    style_config = _current_format_config()["table"]
    table = doc.add_table(rows=0, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _borderless_table(table)
    image_cell = table.add_row().cells[0]
    paragraph = image_cell.paragraphs[0]
    _apply_paragraph(paragraph, style_config, clear_indent=True)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    image_path = Path(temp_dir) / f"{name}{extension}"
    image_path.write_bytes(image_data)
    run.add_picture(
        str(image_path), width=Cm(_REPORT_IMAGE_WIDTH_CM), height=Cm(_REPORT_IMAGE_HEIGHT_CM)
    )
    _unlock_aspect(run)
    text_cell = table.add_row().cells[0]
    text_paragraph = text_cell.paragraphs[0]
    _apply_paragraph(text_paragraph, style_config, clear_indent=True)
    text_run = text_paragraph.add_run(text)
    _apply_run(text_run, style_config, bold=True)
    return table


def _picture(doc, path, width_cm, height_cm=None):
    """插入居中图片；示例/设备图固定 15×8（_REPORT_IMAGE_*），统计图按调用方传入尺寸。"""
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    if width_cm is None:
        width_cm, height_cm = _REPORT_IMAGE_WIDTH_CM, _REPORT_IMAGE_HEIGHT_CM
    if height_cm is None:
        run.add_picture(str(path), width=Cm(width_cm))
    else:
        run.add_picture(str(path), width=Cm(width_cm), height=Cm(height_cm))
    _unlock_aspect(run)


def _example_table(doc, points, photos=None, temp_dir=None):
    """基准表9结构：每示例点=上行 1-2 张照片 + 下行合并文字行（桩号+数值）。

    photos: {示例点序号: [(字节, 扩展名)]}；无照片时仅输出文字行。
    """
    style_config = _current_format_config()["table"]
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _borderless_table(table)
    for index, record in enumerate(points):
        photo_list = (photos or {}).get(index) or []
        if photo_list and temp_dir is not None:
            cells = table.add_row().cells
            merged = cells[0].merge(cells[1])
            paragraph = merged.paragraphs[0]
            _apply_paragraph(paragraph, style_config, clear_indent=True)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = paragraph.add_run()
            image_data, extension = photo_list[0]
            image_path = Path(temp_dir) / f"example_{index}0{extension}"
            image_path.write_bytes(image_data)
            run.add_picture(
                str(image_path),
                width=Cm(_REPORT_IMAGE_WIDTH_CM),
                height=Cm(_REPORT_IMAGE_HEIGHT_CM),
            )
            _unlock_aspect(run)
        cells = table.add_row().cells
        merged = cells[0].merge(cells[1])

        def text(paragraph, value, bold=True, size=None):
            _apply_paragraph(paragraph, style_config, clear_indent=True)
            run = paragraph.add_run(value)
            _apply_run(run, style_config, bold=bold)
            if size is not None:
                run.font.size = size

        top = merged.paragraphs[0]
        text(top, f"{engine.format_station(record['electronic_station'])}（{record['height']:.2f}mm） {engine.format_station(record['raw_station'])}")
    return table


def _distribution_images(segments, stats, records, temp_dir):
    return engine.report_images(temp_dir, segments, stats, records)


def _tci_distribution_images(segments, tci_stats, temp_dir):
    try:
        return engine.report_tci_images(temp_dir, segments, tci_stats)
    except Exception:
        return {}


def generate_guangdong_stub(city, bundle, output_dir, log=lambda _x: None):
    raise NotImplementedError("广东报告需要用户提供第五章模板；当前程序不会伪造该模板。")


def _section_overview(doc, config, segments):
    has_county = bool(segments and any(s.get("county") for s in segments))
    if has_county:
        counties = sorted({str(s.get("county", "")).strip() for s in segments if str(s.get("county", "")).strip()})
        routes = sorted({str(s.get("route", "")).strip() for s in segments if str(s.get("route", "")).strip()})
        county_text = "、".join(counties) if counties else "重庆市"
        route_text = "、".join(routes) if routes else "G210"
        _body(doc, f"本次对{county_text}{route_text}线波形梁护栏{'和'.join(['护栏横梁中心高度', '螺栓缺失'])}进行自动化检测，按采集路段汇总表划分为{len(segments)}个统计分段。{route_text}上行K2264K2325文件按原始桩号统计，其他文件按电子修正桩号统计。")
        _caption(doc, "表", "1.1", f"{county_text}检测路段情况")
        headers = ["序号", "区县", "路线编号", "路线名", "公路等级", "起点桩号", "止点桩号", "里程(km)", "总里程(km)"]
        rows = [[i, s.get("county", ""), s.get("route", "G210"), s.get("route_name", ""), s.get("grade", ""), engine.format_station(s["start"]), engine.format_station(s["end"]), f"{s['mileage']:.3f}", f"{s.get('total_mileage', s['mileage']):.3f}"] for i, s in enumerate(segments, 1)]
    else:
        routes = sorted({str(s.get("route", "")).strip() for s in segments if str(s.get("route", "")).strip()}) if segments else []
        route_text = "、".join(routes) if routes else "G210"
        _body(doc, f"本次对重庆市{route_text}线波形梁护栏{'和'.join(['护栏横梁中心高度', '螺栓缺失'])}进行自动化检测，按采集路段汇总表划分为{len(segments)}个统计分段。{route_text}上行K2264K2325文件按原始桩号统计，其他文件按电子修正桩号统计。")
        _caption(doc, "表", "1.1", f"{route_text}检测路段情况")
        headers = ["序号", "路线编号", "路线名", "公路等级", "起点桩号", "止点桩号", "里程(km)", "总里程(km)"]
        rows = [[i, s.get("route", "G210"), s.get("route_name", ""), s.get("grade", ""), engine.format_station(s["start"]), engine.format_station(s["end"]), f"{s['mileage']:.3f}", f"{s.get('total_mileage', s['mileage']):.3f}"] for i, s in enumerate(segments, 1)]
    _table(doc, headers, rows, True)


def _section_method(doc, height_stats, bolt_stats):
    if height_stats is not None:
        _body(doc, "二波按550、580、620、650mm分档，580≤h≤620mm为合格；三波按647、677、717、747mm分档，677≤h≤717mm为合格。")
    if bolt_stats is not None:
        _body(doc, "螺栓缺失率按缺失数量÷（拼接螺栓数量+连接螺栓数量+缺失数量）×100%计算。")


def _section_height(doc, segments, height_stats, height_records, images, temp_dir, drawing_id=1000, height_photo_index=None, height_photo_workbook=None):
    # 章标题由模板 #4 提供，此处只写“区县整体→逐段”小节；无数据写明确占位句，不留空章/空标题（Q2）。
    from collections import defaultdict
    if not segments:
        _body(doc, "本章暂无有效高度检测记录。")
        return drawing_id
    has_county = bool(segments and any(s.get("county") for s in segments))
    by_county = defaultdict(list)
    for idx, seg in enumerate(segments):
        stat = height_stats[idx] if height_stats is not None and idx < len(height_stats) else None
        by_county[seg.get("county", "")].append((idx, seg, stat))
    for county_idx, (county, items) in enumerate(by_county.items()):
        _heading(doc, f"{county}整体情况" if county else "整体情况", 2)
        county_has = any(it[2] is not None and any(it[2]["types"][kind]["count"] for kind in ("二波", "三波")) for it in items)
        for kind in ("二波", "三波"):
            total = sum(it[2]["types"][kind]["count"] for it in items if it[2] is not None)
            if not total:
                continue
            good = sum(it[2]["types"][kind]["bins"][2] for it in items if it[2] is not None)
            if county:
                _body(doc, f"{county}{kind}形梁护栏有效检测点共{total}个，横梁中心高度合格率为{good * 100 / total:.2f}%。")
            else:
                _body(doc, f"G210线{kind}形梁护栏有效检测点共{total}个，横梁中心高度合格率为{good * 100 / total:.2f}%。")
        if not county_has:
            _body(doc, f"{county}暂无有效高度检测记录。" if county else "本路段暂无有效高度检测记录。")
        labels = ["h＜550", "550≤h＜580", "580≤h≤620", "620＜h≤650", "h＞650"]
        rows = []
        for idx, seg, stat in items:
            if stat is None:
                continue
            data = stat["types"]["二波"] if stat["types"]["二波"]["count"] else None
            if data:
                county_val = seg.get("county", "")
                route_val = seg.get("route", "G210")
                if has_county:
                    rows.append([county_val, route_val, engine.format_station(stat["segment"]["start"]), engine.format_station(stat["segment"]["end"]), *[f"{value:.2f}%" for value in data["pcts"]]])
                else:
                    rows.append([route_val, engine.format_station(stat["segment"]["start"]), engine.format_station(stat["segment"]["end"]), *[f"{value:.2f}%" for value in data["pcts"]]])
        if rows:
            _caption(doc, "表", f"4.{county_idx+1}.1", f"{county}二波形梁护栏横梁中心高度检测结果" if county else "G210线二波形梁护栏横梁中心高度检测结果")
            if has_county:
                _table(doc, ["区县", "路线编号", "起点桩号", "终点桩号", *labels], rows, True)
            else:
                _table(doc, ["路线编号", "起点桩号", "终点桩号", *labels], rows, True)
        labels = ["h＜647", "647≤h＜677", "677≤h≤717", "717＜h≤747", "h＞747"]
        rows = []
        for idx, seg, stat in items:
            if stat is None:
                continue
            data = stat["types"]["三波"] if stat["types"]["三波"]["count"] else None
            if data:
                county_val = seg.get("county", "")
                route_val = seg.get("route", "G210")
                if has_county:
                    rows.append([county_val, route_val, engine.format_station(stat["segment"]["start"]), engine.format_station(stat["segment"]["end"]), *[f"{value:.2f}%" for value in data["pcts"]]])
                else:
                    rows.append([route_val, engine.format_station(stat["segment"]["start"]), engine.format_station(stat["segment"]["end"]), *[f"{value:.2f}%" for value in data["pcts"]]])
        if rows:
            _caption(doc, "表", f"4.{county_idx+1}.2", f"{county}三波形梁护栏横梁中心高度检测结果" if county else "G210线三波形梁护栏横梁中心高度检测结果")
            if has_county:
                _table(doc, ["区县", "路线编号", "起点桩号", "终点桩号", *labels], rows, True)
            else:
                _table(doc, ["路线编号", "起点桩号", "终点桩号", *labels], rows, True)
        for idx, seg, stat in items:
            route_val = seg.get("route", "G210")
            _heading(doc, f"{route_val}线", 2)
            if stat is None or not any(stat["types"][kind]["count"] for kind in ("二波", "三波")):
                _body(doc, f"本段{route_val}线{engine.format_station(seg['start'])}～{engine.format_station(seg['end'])}暂无有效高度检测记录。")
                continue
            for kind in ("二波", "三波"):
                data = stat["types"][kind]
                if not data["count"]:
                    continue
                figure_number = 3 if kind == "三波" and stat["types"]["二波"]["count"] else 0
                level = "较高" if data["pass"] >= 80 else ("一般" if data["pass"] >= 60 else "偏低")
                detail = engine.percentage_phrases(kind, data["pcts"])
                _body(doc, f"{route_val}线{kind}形梁护栏横梁中心高度有效检测点共{data['count']}个，整体合格率{level}，合格率为{data['pass']:.2f}%。护栏横梁中心高度{detail}。")
                image_set = images.get((idx, kind)) if isinstance(images, dict) else None
                if image_set is None:
                    continue
                drawing_id += 1
                _picture(doc, image_set["line"], 13, 8)
                _caption(doc, "图", f"4.{idx + 2}-{figure_number + 1}", f"{kind}形梁护栏横梁中心高度检测结果")
                _picture(doc, image_set["pie"], 14, 8.5)
                _caption(doc, "图", f"4.{idx + 2}-{figure_number + 2}", f"{kind}形梁护栏横梁中心高度分布情况")
                example_rows = [record for record in (height_records or []) if record["segment"] == idx and record["kind"] == kind]
                example_points = engine.select_height_example_points(
                    example_rows, idx, kind, photo_index=height_photo_index
                )
                if example_points:
                    photos = _match_station_photos(example_points, height_photo_index, height_photo_workbook)
                    _example_table(doc, example_points, photos, temp_dir)
                    _caption(doc, "图", f"4.{idx + 2}-{figure_number + 3}", f"{kind}形梁护栏横梁中心高度自动计算示例")
    return drawing_id


def _match_station_photos(points, photo_index, workbook=None):
    """示例点 -> {序号: [(字节, 扩展名)]}：按 (方向, 桩号) 命中病害照片索引（最多 2 张）。"""
    if not photo_index or workbook is None:
        return {}
    result = {}
    for index, point in enumerate(points):
        if not engine.row_has_height_photo(point, photo_index):
            continue
        direction = engine.normalize_direction(point.get("direction"))
        for station_key in ("raw_station", "electronic_station", "station"):
            station = point.get(station_key)
            if station is None:
                continue
            photos = photo_index.get((direction, round(float(station), 1)))
            if photos:
                result[index] = [
                    engine.read_media(workbook, media_name) for media_name, _ in photos[:2]
                ]
                break
    return result


def _height_point_count(height_stats, index: int) -> int:
    """同区段高度有效检测点总数；缺高度统计时返回 0（视为同样无数据）。"""
    try:
        stat = height_stats[index] if height_stats is not None and index < len(height_stats) else None
    except (TypeError, IndexError):
        return 0
    if not stat or not isinstance(stat.get("types"), dict):
        return 0
    total = 0
    for kind in ("二波", "三波"):
        data = stat["types"].get(kind) or {}
        try:
            total += int(data.get("count") or 0)
        except (TypeError, ValueError):
            continue
    return total


def _section_bolt(doc, segments, bolt_stats, bolt_records, disease_image_index, temp_dir, height_stats=None):
    # 章标题由模板 #5 提供，此处只写“区县整体→逐段”小节。
    # 示例图逻辑不变：0 记录写占位句；有记录时仅关联到病害图才挂图。
    from collections import defaultdict
    if not segments:
        _body(doc, "本章暂无有效螺栓检测记录。")
        return
    has_county = bool(segments and any(s.get("county") for s in segments))
    by_county = defaultdict(list)
    for idx, seg in enumerate(segments):
        stat = bolt_stats[idx] if bolt_stats is not None and idx < len(bolt_stats) else None
        by_county[seg.get("county", "")].append((idx, seg, stat))
    for county_idx, (county, items) in enumerate(by_county.items()):
        _heading(doc, f"{county}整体情况" if county else "整体情况", 2)
        total_splice = sum(it[2]["splice"] for it in items if it[2] is not None)
        total_connection = sum(it[2]["connection"] for it in items if it[2] is not None)
        total_missing = sum(it[2]["missing"] for it in items if it[2] is not None)
        total_rate = engine.bolt_missing_rate(total_splice, total_connection, total_missing)
        if county:
            _body(doc, f"本次采用波形梁护栏螺栓缺失自动化检测方式，对{county}波形梁护栏螺栓缺失情况进行统计，共检出拼接螺栓{total_splice}颗，连接螺栓{total_connection}颗，缺失螺栓{total_missing}颗，整体缺失率为{_rate_text(total_rate)}%。")
        else:
            _body(doc, f"本次采用波形梁护栏螺栓缺失自动化检测方式，对重庆市G210线波形梁护栏螺栓缺失情况进行统计，共检出拼接螺栓{total_splice}颗，连接螺栓{total_connection}颗，缺失螺栓{total_missing}颗，整体缺失率为{_rate_text(total_rate)}%。")
        overall = []
        for idx, seg, stat in items:
            if stat is None:
                continue
            county_val = seg.get("county", "")
            route_val = seg.get("route", "G210")
            if has_county:
                overall.append([county_val, route_val, engine.format_station(stat["segment"]["start"]), engine.format_station(stat["segment"]["end"]), f"{stat['segment']['mileage']:.3f}", stat["splice"], stat["connection"], stat["missing"], f"{_rate_text(stat['rate'])}"])
            else:
                overall.append([route_val, engine.format_station(stat["segment"]["start"]), engine.format_station(stat["segment"]["end"]), f"{stat['segment']['mileage']:.3f}", stat["splice"], stat["connection"], stat["missing"], f"{_rate_text(stat['rate'])}"])
        if has_county:
            overall.append(["合计", "", "", "", "", total_splice, total_connection, total_missing, f"{_rate_text(total_rate)}"])
            _caption(doc, "表", f"5.{county_idx+1}.1", f"{county}波形梁护栏螺栓缺失检测结果" if county else "G210线波形梁护栏螺栓缺失检测结果")
            _table(doc, ["区县", "路线编号", "起点桩号", "止点桩号", "检测里程（km）", "拼接螺栓（颗）", "连接螺栓（颗）", "缺失数量（颗）", "缺失率（%）"], overall, True)
        else:
            overall.append(["合计", "", "", "", total_splice, total_connection, total_missing, f"{_rate_text(total_rate)}"])
            _caption(doc, "表", f"5.{county_idx+1}.1", f"{county}波形梁护栏螺栓缺失检测结果" if county else "G210线波形梁护栏螺栓缺失检测结果")
            _table(doc, ["路线编号", "起点桩号", "止点桩号", "检测里程（km）", "拼接螺栓（颗）", "连接螺栓（颗）", "缺失数量（颗）", "缺失率（%）"], overall, True)
        for idx, seg, stat in items:
            route_val = seg.get("route", "G210")
            section_number = idx + 2
            _heading(doc, f"{route_val}线", 2)
            if stat is None:
                _body(doc, f"本段{route_val}线{engine.format_station(seg['start'])}～{engine.format_station(seg['end'])}暂无有效螺栓检测记录。")
                continue
            if not stat["points"]:
                if _height_point_count(height_stats, idx) > 0:
                    # 同段高度有有效数据：不断言“无波形护栏”，仅说明螺栓明细缺失。
                    _body(doc, "本段螺栓明细无有效记录。")
                else:
                    _body(doc, "本段无波形护栏")
                continue
            _body(doc, f"{route_val}线共检出拼接螺栓{stat['splice']}颗，连接螺栓{stat['connection']}颗，缺失螺栓{stat['missing']}颗，缺失率为{_rate_text(stat['rate'])}%。")
            _caption(doc, "表", f"5.{section_number}-1", f"{route_val}线波形梁护栏螺栓缺失检测结果")
            county_val = seg.get("county", "")
            if has_county:
                _table(doc, ["区县", "路线编号", "起点桩号", "止点桩号", "里程（km）", "拼接螺栓（颗）", "连接螺栓（颗）", "缺失数量（颗）"], [[county_val, route_val, engine.format_station(seg["start"]), engine.format_station(seg["end"]), f"{seg['mileage']:.3f}", stat["splice"], stat["connection"], stat["missing"]]], True)
            else:
                _table(doc, ["路线编号", "起点桩号", "止点桩号", "里程（km）", "拼接螺栓（颗）", "连接螺栓（颗）", "缺失数量（颗）"], [[route_val, engine.format_station(seg["start"]), engine.format_station(seg["end"]), f"{seg['mileage']:.3f}", stat["splice"], stat["connection"], stat["missing"]]], True)
            segment_bolt_rows = [record for record in (bolt_records or []) if record["segment"] == idx]
            bolt_examples = engine.select_bolt_example_points(segment_bolt_rows, disease_image_index)
            bolt_examples = [example for example in bolt_examples if example.get("image") is not None]
            if bolt_examples:
                for example in bolt_examples:
                    image_data, image_extension = engine.read_disease_image(example["image"])
                    _photo_text_table(doc, image_data, image_extension, engine.bolt_example_text(example), temp_dir, f"bolt_{section_number}")
                _caption(doc, "图", f"5.{section_number}-1", f"{route_val}线波形梁护栏螺栓缺失自动识别示例")


def _tci_has_data(st) -> bool:
    """该分段是否有有效 TCI 记录（count>0 才算有数据；0 记录的 TCI=100 仅为缺省值）。"""
    return st is not None and int(st.get("count") or 0) > 0


def _section_tci(doc, segments, tci_stats, tci_images=None, temp_dir=None, tci_photos=None, segment_tci_photos=None):
    """县整体→路线：共用逐公里评分，不将不同单位病害混画为构成图。"""
    routes = engine.tci_route_stats(segments, tci_stats)
    counties = dict.fromkeys(s.get("county", "") for s in segments)
    if tci_stats is None:
        _body(doc, "未提供 TCI 病害清单，沿线设施技术状况未评定。")
        return
    for county in counties:
        _heading(doc, f"{county}整体情况" if county else "整体情况", 2)
        county_segments = [s for s in segments if s.get("county", "") == county]
        route_names = list(dict.fromkeys(s.get("route", "G210") for s in county_segments))
        _body(doc, f"{county}沿线设施共检测{len(route_names)}条路线，总里程{sum(s.get('mileage', 0) for s in county_segments):.3f}km。TCI按原检测范围内整公里桩号划分评定单元，首尾不足一公里单独评定，上、下行分别计分；汇总采用评定单元等权平均。")
        summary = []
        for route in route_names:
            groups = [g for g in routes if g["county"] == county and g["route"] == route]
            source = [s for s in county_segments if s.get("route", "G210") == route]
            units = [u for g in groups for u in g["units"]]
            if not units:
                continue
            score = sum(u["tci"] for u in units) / len(units)
            _body(doc, f"{route}线检测里程{sum(s.get('mileage', 0) for s in source):.3f}km，沿线设施TCI均值为{score:.2f}，评定为{engine.tci_grade(score)}等。")
            summary.append([len(summary)+1, county, route, "/".join(g["direction"] or "未注明" for g in groups),
                            engine.format_station(min(u["start"] for u in units)), engine.format_station(max(u["end"] for u in units)),
                            f"{sum(s.get('mileage', 0) for s in source):.3f}", "/".join(f"{g['tci']:.2f}" for g in groups),
                            "/".join(g["grade"] for g in groups)])
        if summary:
            _caption(doc, "表", "3.1-1", f"{county}沿线设施检查评定汇总表")
            _table(doc, ["序号", "区县", "路线编号", "方向", "起点桩号", "止点桩号", "里程（km）", "TCI", "等级"], summary, True)
        for number, route in enumerate(route_names, 2):
            _heading(doc, f"{route}线", 2)
            source = [(i,s) for i,s in enumerate(segments) if s.get("county", "") == county and s.get("route", "G210") == route]
            ranges = "、".join(dict.fromkeys(f"{engine.format_station(s['start'])}～{engine.format_station(s['end'])}" for _,s in source))
            _body(doc, f"{county}{route}线检测路段为{ranges}。")
            if any(not tci_stats[i].get("units") for i, _ in source):
                _body(doc, "部分检测区间未提供TCI数据，未计入均值；详见附表中的实际评定范围。")
            chart_number = 0
            for i, seg in source:
                for direction in ("上行", "下行", ""):
                    image = (tci_images or {}).get((i, direction))
                    if image is None:
                        continue
                    chart_number += 1
                    direction_text = f"{direction}" if direction else "未注明方向"
                    _body(doc, f"{route}线{direction_text}{engine.format_station(seg['start'])}～{engine.format_station(seg['end'])}段各评定单元TCI指数计算结果如下图所示。")
                    _picture(doc, image, 13, 8)
                    _caption(doc, "图", f"3.{number}-{chart_number}", f"{route}线{direction_text}{engine.format_station(seg['start'])}～{engine.format_station(seg['end'])}段TCI情况")
            used = set()
            for index, _ in source:
                for photo in (segment_tci_photos or {}).get(index, []):
                    description, media_name, extension = photo[:3]
                    workbook = photo[3] if len(photo) > 3 else (tci_photos[0] if tci_photos else None)
                    if workbook is None or description in used:
                        continue
                    used.add(description)
                    data, _ = engine.read_media(workbook, media_name)
                    _photo_text_table(doc, data, extension, description, temp_dir, f"route_tci_{number}_{len(used)}")
            if used:
                _caption(doc, "图", f"3.{number}-2", f"{route}线交安设施典型病害图")


def _section_tci_appendix(doc, segments, tci_stats):
    """附表1：路线整体行及逐公里、逐方向评定明细，跨页重复表头。"""
    groups = engine.tci_route_stats(segments, tci_stats)
    for county in dict.fromkeys(s.get("county", "") for s in segments):
        doc.add_page_break()
        name = county if county.startswith("重庆市") else f"重庆市{county}"
        paragraph = doc.add_paragraph(f"附表1 {name}交安设施技术状况评定明细")
        paragraph.style = doc.styles["Heading 1"]
        heading_format = _current_format_config()["heading"]["1"]
        _apply_paragraph(paragraph, heading_format, clear_indent=True)
        for run in paragraph.runs:
            _apply_run(run, heading_format)
        if not hasattr(doc, "_toc_entries"):
            doc._toc_entries = []
        doc._toc_entries.append(("", paragraph.text))
        # 附表不参与正文第7章自动编号，但仍进入目录。
        num_pr = OxmlElement("w:numPr")
        num_id = OxmlElement("w:numId")
        num_id.set(qn("w:val"), "0")
        num_pr.append(num_id)
        paragraph._p.get_or_add_pPr().append(num_pr)
        rows = []
        for route in dict.fromkeys(s.get("route", "G210") for s in segments if s.get("county", "") == county):
            selected = [g for g in groups if g["county"] == county and g["route"] == route]
            units = [u for g in selected for u in g["units"]]
            if not units:
                continue
            score = sum(u["tci"] for u in units) / len(units)
            rows.append([len(rows)+1, route, county, "/".join(g["direction"] or "未注明" for g in selected), "/", "/", f"{score:.2f}", engine.tci_grade(score)])
            for unit in units:
                rows.append([len(rows)+1, route, county, unit["direction"] or "未注明", engine.format_station(unit["start"]),
                             engine.format_station(unit["end"]), f"{unit['tci']:.2f}", unit["grade"]])
        table = _table(doc, ["序号", "路线编号", "区县", "方向", "起点桩号", "止点桩号", "TCI", "等级"], rows, True)
        table.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))


def _section_conclusion(doc, segments, height_stats, bolt_stats, tci_stats=None):
    # 6.1 结论 + 6.2 建议，结构参考《报告模板-5.docx》附件模板。
    _heading(doc, "结论", 2)
    has_county = bool(segments and any(s.get("county") for s in segments))
    units = [u for stat in tci_stats or [] for u in stat.get("units", [])]
    # —— 沿线设施 TCI 结论 ——
    if units:
        mean = sum(u["tci"] for u in units) / len(units)
        grade = engine.tci_grade(mean)
        weak = [u for u in units if u["tci"] < 80]
        weak_text = ""
        if weak:
            parts = [f"{engine.format_station(u['start'])}～{engine.format_station(u['end'])}段" for u in weak]
            weak_text = f"其中{('、'.join(parts))}评定为{'中等' if all(u['tci'] >= 70 for u in weak) else '次差'}，需要特别注意。"
        if has_county:
            county = next(s.get("county", "") for s in segments if s.get("county"))
            _body(doc, f"{county}沿线设施技术状况TCI整体状况{'良好' if mean >= 80 else '一般'}，{weak_text or '各路段均达到优良等级。'}")
        else:
            _body(doc, f"沿线设施技术状况TCI整体状况{'良好' if mean >= 80 else '一般'}，{weak_text or '各路段均达到优良等级。'}")
        _body(doc, "检测段标志主要病害或缺陷为标志板遮挡、标志板变形、反光膜缺损等；标线主要问题为缺损；波形护栏主要问题为缺损、变形等。")
    # —— 高度结论 ——
    if height_stats is not None:
        for kind in ("二波", "三波"):
            total = sum(item["types"][kind]["count"] for item in height_stats)
            if not total:
                continue
            good = sum(item["types"][kind]["bins"][2] for item in height_stats)
            rate = good * 100 / total
            level = "较低" if rate < 60 else ("一般" if rate < 80 else "良好")
            _body(doc, f"检测路段{kind}形梁护栏横梁中心高度整体合格率{level}，合格率为{rate:.2f}%。")
    # —— 螺栓结论 ——
    if bolt_stats is not None:
        total_splice = sum(item["splice"] for item in bolt_stats)
        total_connection = sum(item["connection"] for item in bolt_stats)
        total_missing = sum(item["missing"] for item in bolt_stats)
        total_rate = engine.bolt_missing_rate(total_splice, total_connection, total_missing)
        _body(doc, f"波形梁护栏螺栓整体缺失率为{_rate_text(total_rate)}%。")

    _heading(doc, "建议", 2)
    _body(doc, "依据《公路养护技术标准》（JTG 5110-2023）对高速公路交通安全设施要求，结合本次抽检路段交通安全设施检测评定结果，为确保路面行驶安全舒适，提出以下养护建议，仅供参考：")
    _body(doc, "1.沿线设施技术状况检查建议")
    _body(doc, "管养单位需针对病害类型，制定针对性的维护措施，如加强设施的日常巡查、及时修复缺损部分，以提升高速公路的整体安全性和美观度。")
    _body(doc, "对交通安全设施技术状况TCI评定为优、良等级的路段，可加强日常养护和巡查，及时排除有损交通安全设施的各种不良因素，对发现的交安设施病害，及时的采取有效措施进行维修。")
    _body(doc, "对交通安全设施技术状况TCI评定为中、次、差等级或病害集中的路段，建议针对性地制定维护方案，优先解决标志遮挡、标线完善和护栏加固等问题，同时加强周期性检查，以提升设施的安全性和耐久性。具体处置建议如下。")
    _body(doc, "标志遮挡，建议：立即清理遮挡交通标志的障碍物，确保交通标志清晰可见。若障碍物为自然生长物（如树木枝叶），应定期修剪；若为人为搭建物，应与相关部门协调拆除。")
    _body(doc, "标线缺损，建议：清理标线残留部分，采用热熔型反光标线涂料重新施划，确保道路使用者能够清晰识别行驶路线。")
    _body(doc, "波形梁变形，建议：拆除变形护栏段，更换同型号波形梁板、立柱及连接件，确保护栏线形顺直，螺栓紧固到位，拼接处平滑过渡，恢复防撞防护功能。")
    _body(doc, "2.波形梁护栏专项检测建议")
    _body(doc, "（1）波形梁护栏横梁中心高度")
    _body(doc, "因波形梁护栏中心高度过低或过高均可能导致护栏防护功能受到严重影响。建议波形梁中心高度合格率低的路段管养单位应制定养护计划对波形梁护栏进行整治。")
    _body(doc, "（2）波形梁护栏螺栓缺失")
    _body(doc, "波形梁护栏的螺栓缺失会导致护栏的防护性能下降，应加强养护巡查。当同一处拼接处有缺少3个（含3个）以上螺栓时，需及时修复，在缺少3个以下螺栓时，本着安全第一的原则，也应及时有条件的进行维修，确保波形梁护栏符合设计要求。")


_STATIC_CONCLUSION_HEADINGS = frozenset({"结论", "建议", "结论与建议"})
_STATIC_CONCLUSION_PARAGRAPHS = ("结论章节由程序根据检测统计自动生成", "建议章节由程序根据检测统计自动生成")


def _strip_section_number(text: str) -> str:
    """去掉“5.1 ”“5 ”等章节编号前缀，便于识别静态结论残留标题。"""
    stripped = (text or "").strip()
    index = 0
    while index < len(stripped) and (stripped[index].isdigit() or stripped[index] in ".．·、）) "):
        index += 1
    return stripped[index:].strip()


_CQ_HEADER_COMPANY = "四川京炜交通工程技术有限公司"


def _paragraph_bottom_border(paragraph, size="6") -> None:
    """段落下框线：页眉/封面横线共用（base.pdf 基准单线）。"""
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    bottom = p_bdr.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        p_bdr.append(bottom)
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")


def _add_page_field(paragraph, instruction: str) -> None:
    """页码域（PAGE / NUMPAGES）：无缓存结果，Word 打开时计算；数字 Times 黑色。"""
    run = paragraph.add_run()
    _set_run_fonts(run, "宋体", LATIN_FONT)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    run._r.append(begin)
    run = paragraph.add_run()
    _set_run_fonts(run, "宋体", LATIN_FONT)
    instr = OxmlElement("w:instrText")
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = f" {instruction} "
    run._r.append(instr)
    run = paragraph.add_run()
    _set_run_fonts(run, "宋体", LATIN_FONT)
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    run._r.append(separate)
    run = paragraph.add_run()
    _set_run_fonts(run, "宋体", LATIN_FONT)
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(end)


def _set_chongqing_header_footer(section, report_no: str) -> None:
    """重庆页眉页脚（base.pdf 基准）：左公司名 + 右报告编号变量 + 页眉下横线；
    页脚“第 X 页 共 Y 页”居中。仅重庆链路调用，广东 writer 不经过此函数。"""
    from docx.enum.text import WD_TAB_ALIGNMENT

    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False
    header = section.header
    paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.text = ""
    content_width = section.page_width - section.left_margin - section.right_margin
    try:
        paragraph.paragraph_format.tab_stops.add_tab_stop(content_width, WD_TAB_ALIGNMENT.RIGHT)
    except Exception:
        pass
    left = paragraph.add_run(_CQ_HEADER_COMPANY)
    _set_run_fonts(left, "宋体", LATIN_FONT)
    left.font.size = Pt(9)
    paragraph.add_run().add_tab()
    right = paragraph.add_run(report_no or "")
    _set_run_fonts(right, "宋体", LATIN_FONT)
    right.font.size = Pt(9)
    _paragraph_bottom_border(paragraph)
    footer = section.footer
    fpara = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    fpara.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fpara.text = ""
    for text in ("第 ",):
        run = fpara.add_run(text)
        _set_run_fonts(run, "宋体", LATIN_FONT)
        run.font.size = Pt(9)
    _add_page_field(fpara, "PAGE")
    mid = fpara.add_run(" 页 共 ")
    _set_run_fonts(mid, "宋体", LATIN_FONT)
    mid.font.size = Pt(9)
    _add_page_field(fpara, "NUMPAGES")
    tail = fpara.add_run(" 页")
    _set_run_fonts(tail, "宋体", LATIN_FONT)
    tail.font.size = Pt(9)


def _cover_section_rule(section) -> None:
    """封面节页眉规则线（base.pdf 基准）：空文本 + 段落下框线。

    封面/注意事项/目录页（第 1 节）共用此页眉；正文节由
    _set_chongqing_header_footer 提供公司名+编号+横线。
    """
    section.header.is_linked_to_previous = False
    header = section.header
    paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    paragraph.text = ""
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "9")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def _start_chongqing_body_section(doc, report_no: str):
    """封面/注意事项/目录节之后另起一节：正文加页眉页脚，页码从 1 重计（base.pdf 第 1 页起于概况）。"""
    from docx.enum.section import WD_SECTION

    first = doc.sections[0]
    body = doc.add_section(WD_SECTION.NEW_PAGE)
    body.page_width = first.page_width
    body.page_height = first.page_height
    body.left_margin = first.left_margin
    body.right_margin = first.right_margin
    body.top_margin = first.top_margin
    body.bottom_margin = first.bottom_margin
    body.orientation = first.orientation
    pg_num = OxmlElement("w:pgNumType")
    pg_num.set(qn("w:start"), "1")
    body._sectPr.append(pg_num)
    _set_chongqing_header_footer(body, report_no)
    # T9：正文节首页也要页眉页脚。titlePg 会随 add_section 从封面节克隆而来，
    # 不清除会导致正文第一页（概况页）无页眉（base.pdf 基准仅封面/目录节无页眉）。
    body.different_first_page_header_footer = False
    return body


def _report_with_skeleton(doc, config, blocks, segments, height_stats, height_records, bolt_stats, bolt_records, tci_stats, tci_records, disease_image_index, images, temp_dir, skeleton_md=None, tci_images=None, height_photo_index=None, height_photo_workbook=None, tci_photos=None, segment_tci_photos=None):
    # 只认显式锚点 <!-- inject:xxx -->，无关键词回退：
    # 静态标题（含 2.3.x）一律按原样渲染，绝不触发注入、不吞掉（Q2）。
    # 章标题由模板 #1~#6 提供，各 _section_* 只写章内小节，不再自写章标题。
    import re as _re
    route_segments, route_heights, route_height_records, route_bolts, route_bolt_records = engine.route_report_data(segments, height_stats, height_records, bolt_stats, bolt_records)
    images = _distribution_images(route_segments, route_heights, route_height_records, temp_dir) if route_heights is not None else {}
    anchor_re = _re.compile(r"<!--\s*inject:\s*(overview|tci|height|bolt|conclusion|tci_appendix)\s*-->")
    _CANONICAL_ORDER = ("overview", "tci", "height", "bolt", "conclusion", "tci_appendix")
    def _writer_for(key: str):
        if key == "overview":
            return lambda: _section_overview(doc, config, segments)
        if key == "tci":
            return lambda: _section_tci(doc, segments, tci_stats, tci_images, temp_dir, tci_photos, segment_tci_photos)
        if key == "height":
            if height_stats is None:
                return lambda: _body(doc, "本章暂无有效高度检测记录。")
            return lambda: _section_height(doc, route_segments, route_heights, route_height_records, images, temp_dir, height_photo_index=height_photo_index, height_photo_workbook=height_photo_workbook)
        if key == "bolt":
            if bolt_stats is None:
                return lambda: _body(doc, "本章暂无有效螺栓检测记录。")
            return lambda: _section_bolt(doc, route_segments, route_bolts, route_bolt_records, disease_image_index, temp_dir, route_heights)
        if key == "tci_appendix":
            return lambda: _section_tci_appendix(doc, segments, tci_stats)
        if key == "conclusion":
            return lambda: _section_conclusion(doc, route_segments, route_heights, route_bolts, tci_stats)
        return lambda: None

    # pending 按显式 key 索引，便于锚点直接命中
    pending_by_key = {
        "overview": _writer_for("overview"),
        "tci": _writer_for("tci"),
        "height": _writer_for("height"),
        "bolt": _writer_for("bolt"),
        "conclusion": _writer_for("conclusion"),
        "tci_appendix": _writer_for("tci_appendix"),
    }
    pending_keys = set(pending_by_key.keys())

    def _inject(key: str):
        if key not in pending_keys:
            return
        writer = pending_by_key[key]
        pending_keys.remove(key)
        writer()

    skeleton_dir = Path(skeleton_md).parent if skeleton_md else None

    # T9：封面（首页）无页眉页脚；目录节后另起正文节（页眉页脚 + 页码从 1 重计）。
    # 无目录块的极简模板退化为单节：首页仍不同（封面干净），其余页共用页眉页脚。
    doc.sections[0].different_first_page_header_footer = True
    # T11：封面节页眉规则线（空文本+下框线），封面/注意事项/目录页顶线可见
    _cover_section_rule(doc.sections[0])
    toc_idx = next((i for i, b in enumerate(blocks) if b.kind == "toc"), None)
    report_no = ""
    body_started = False

    for index, block in enumerate(blocks):
        if toc_idx is not None and index > toc_idx and not body_started:
            body_started = True
            _start_chongqing_body_section(doc, report_no)
        # 1) 显式锚点：任意块（标题/段落）的文本中包含 <!-- inject:xxx -->
        anchor_hit = None
        if block.text:
            m = anchor_re.search(block.text)
            if m:
                anchor_hit = m.group(1)
        if anchor_hit:
            # 显式锚点不渲染原块文本，直接注入
            _inject(anchor_hit)
            continue

        if block.kind == "toc":
            if _current_format_config()["toc"]["enabled"]:
                _add_toc(doc)
        elif block.kind == "cover":
            report_no = _cover_paragraphs(doc, block.text) or report_no
        elif block.kind == "notes":
            _notes_paragraphs(doc, block.text)
        elif block.kind == "heading":
            # 无关键词回退：所有静态标题（含 2.3.x）原样渲染。
            if "G210线K" in block.text and block.level >= 3:
                continue
            if "conclusion" not in pending_keys and _strip_section_number(block.text) in _STATIC_CONCLUSION_HEADINGS:
                # 结论已程序注入：跳过模板残留的静态 5.1结论/5.2建议，避免结论/建议各出现两轮（D2）。
                continue
            _heading(doc, block.text, block.level)
        elif block.kind == "paragraph":
            if block.text:
                if "conclusion" not in pending_keys and any(
                    marker in block.text for marker in _STATIC_CONCLUSION_PARAGRAPHS
                ):
                    continue
                formula_match = _re.match(r"^<!--\s*formula:\s*(tci|gd|w|i0)\s*-->(.*)$", block.text, _re.S)
                if formula_match:
                    formula_key, suffix = formula_match.groups()
                    if formula_key == "tci":
                        _append_formula_paragraph(doc, _tci_formula_omath(), centered=True)
                    elif formula_key == "gd":
                        _append_formula_paragraph(doc, _subscript_omath("GD", "iTCI"), suffix=suffix)
                    elif formula_key == "w":
                        _append_formula_paragraph(doc, _subscript_omath("w", "i"), suffix=suffix)
                    else:
                        _append_formula_paragraph(doc, _subscript_omath("i", "0"), suffix=suffix)
                    continue
                # 模板中的静态表题/图题（“表2.1-1 …”“图 2.3.2-1 …”）按 caption 样式渲染
                if _re.match(r"^[表图]\s*\d", block.text):
                    _caption_text(doc, block.text)
                    continue
                _body(doc, block.text)
        elif block.kind == "table":
            if block.rows:
                _table(doc, block.rows[0], block.rows[1:], False)
        elif block.kind == "picture":
            if skeleton_dir is not None:
                try:
                    media_path = (skeleton_dir / block.caption).resolve()
                    if media_path.is_file():
                        # 按 13cm 宽度插入，高度自适应；失败则跳过
                        try:
                            _picture(doc, media_path, 13, 8)
                            continue
                        except Exception:
                            pass
                except Exception:
                    pass
            # 无有效图片时跳过（报告图片由引擎生成）
            continue
    # 未命中锚点按基准章序追加末尾（不丢数据；模板含全部显式锚点时此分支不触发）
    # T9：目录后无块时在此处另起正文节；无目录时单节页眉页脚落到首页之外的所有页。
    if toc_idx is None:
        _set_chongqing_header_footer(doc.sections[0], report_no)
    elif not body_started:
        _start_chongqing_body_section(doc, report_no)
    for key in _CANONICAL_ORDER:
        if key in pending_keys:
            pending_by_key[key]()


def make_report(config, segments, height_stats, height_records, bolt_stats, bolt_records, tci_stats, tci_records, disease_image_index, temp_dir, log=lambda _x: None, skeleton_md=None):
    """Build the full report document from computed statistics."""
    images = {}  # 路线汇总后统一生成高度图。
    tci_images = _tci_distribution_images(segments, tci_stats, temp_dir) if tci_stats is not None else {}
    if skeleton_md is None:
        raise FileNotFoundError(f"Markdown 模板不存在：{skeleton_md}，仅支持 .md 模板。")
    template = markdown_skeleton.read_template(skeleton_md)
    # T11：病害照片索引（高度示例表挂真实照片）与 TCI 类型示例图
    height_photo_index = {}
    height_photo_workbook = None
    if disease_image_index:
        workbook = next(iter(disease_image_index.values()))[0]["workbook"]
        image_map = engine.build_disease_image_map(workbook)
        if image_map:
            height_photo_index = engine.disease_station_photo_index(workbook, image_map)
            height_photo_workbook = workbook
    tci_photos = None
    segment_tci_photos = {}
    if config.tci_path is not None and tci_stats is not None:
        candidate = Path(config.tci_path)
        files = [candidate] if candidate.is_file() else sorted(candidate.glob("*.xlsx"))
        for workbook in files:
            if workbook.name.startswith("~$"):
                continue
            image_map = engine.build_tci_image_map(workbook)
            for index, photos in engine.build_segment_tci_photos(workbook, image_map, segments).items():
                segment_tci_photos.setdefault(index, []).extend((*photo, workbook) for photo in photos)
    doc = Document()
    with format_context(template.config):
        configure_document(doc, template.config)
        _report_with_skeleton(doc, config, template.blocks, segments, height_stats, height_records, bolt_stats, bolt_records, tci_stats, tci_records, disease_image_index, images, temp_dir, skeleton_md=skeleton_md, tci_images=tci_images, height_photo_index=height_photo_index, height_photo_workbook=height_photo_workbook, tci_photos=tci_photos, segment_tci_photos=segment_tci_photos)
        # 预填充目录缓存（updateFields 关闭时 Word/WPS 打开即显示目录，页码占位待固化）
        _populate_toc_cache(doc)
    try:
        doc.save(config.out_docx)
    except PermissionError as exc:
        raise PermissionError(f"Word文件被占用：{config.out_docx}") from exc
    return config.out_docx



def run(config, segments, height_stats, height_records, bolt_stats, bolt_records, disease_image_index, log=lambda _x: None, skeleton_md=None, tci_stats=None, tci_records=None):
    import tempfile

    with tempfile.TemporaryDirectory(prefix="g210_report_builtin_") as temp_dir:
        config.output_dir.mkdir(parents=True, exist_ok=True)
        log("使用 Markdown 报告模板，以程序化方式生成报告。")
        return make_report(config, segments, height_stats, height_records, bolt_stats, bolt_records, tci_stats, tci_records, disease_image_index, temp_dir, log, skeleton_md=skeleton_md)
