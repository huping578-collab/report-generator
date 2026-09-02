"""Programmatic DOCX report generation driven by a Markdown skeleton template.

A Markdown template (templates/重庆项目报告模板.md) declares the report
skeleton (headings, static body texts, tables). Dynamic data sections
(overview table, height/bolt statistics, charts, conclusions) are injected at
anchor headings; when an anchor is absent the section is appended at the end.
Requires a Markdown skeleton; explicit anchors <!-- inject:... --> are preferred over keyword fallback.
"""

from __future__ import annotations

import math
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


def _set_run_fonts(run, east_asia=BODY_FONT, latin=LATIN_FONT):
    run.font.name = latin
    run.font.color.rgb = RGBColor(0, 0, 0)
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    r_fonts.set(qn("w:eastAsia"), east_asia)
    r_fonts.set(qn("w:ascii"), latin)
    r_fonts.set(qn("w:hAnsi"), latin)


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


def _heading(doc, text, level):
    sizes = {1: 16, 2: 15, 3: 14, 4: 12, 5: BODY_SIZE}
    paragraph = doc.add_paragraph(style=f"Heading {level}")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(10 if level == 1 else 6)
    paragraph.paragraph_format.space_after = Pt(6)
    _clear_indent(paragraph)
    _warn(paragraph, level)
    run = paragraph.add_run(text)
    run.bold = False
    run.italic = False
    run.font.size = sizes.get(level, BODY_SIZE)
    _set_run_fonts(run, HEADING_FONT if level <= 3 else BODY_FONT)
    return paragraph


def _body(doc, text):
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = paragraph.add_run(text)
    run.font.size = BODY_SIZE
    _set_run_fonts(run)
    fmt = paragraph.paragraph_format
    fmt.first_line_indent = Pt(21)
    fmt.line_spacing = 1.5
    return paragraph


def _caption(doc, prefix, number, title):
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _clear_indent(paragraph)
    run = paragraph.add_run(f"{prefix}{number} {title}")
    run.bold = True
    run.font.size = BODY_SIZE
    _set_run_fonts(run, HEADING_FONT)
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(4)
    return paragraph


def _table(doc, headers, rows, header_shading=False):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    def fill(cell, value, bold=False):
        cell.text = str(value)
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _clear_indent(paragraph)
            if header_shading and bold:
                shading = OxmlElement("w:shd")
                shading.set(qn("w:val"), "clear")
                shading.set(qn("w:fill"), "e8edf0")
                cell._tc.get_or_add_tcPr().append(shading)
            run = paragraph.runs[0]
            run.bold = header_shading and bold
            run.font.size = Pt(10)
            _set_run_fonts(run)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    for index, header in enumerate(headers):
        fill(table.rows[0].cells[index], header, True)
    for row in rows:
        cells = table.add_row().cells
        for index, value in enumerate(row):
            fill(cells[index], value)
    return table


def _picture(doc, path, width_cm, height_cm):
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    run.add_picture(str(path), width=Cm(width_cm), height=Cm(height_cm))


def _example_table(doc, points):
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for record in points:
        cells = table.add_row().cells
        merged = cells[0].merge(cells[1])

        def text(paragraph, value, bold=True, size=Pt(10.5)):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = paragraph.add_run(value)
            run.bold = bold
            run.font.size = size
            _set_run_fonts(run)

        top = merged.paragraphs[0]
        text(top, f"{engine.format_station(record['electronic_station'])}（{record['height']:.2f}mm） {engine.format_station(record['raw_station'])}")
    return table


def _distribution_images(segments, stats, records, temp_dir):
    return engine.report_images(temp_dir, segments, stats, records)


def generate_guangdong_stub(city, bundle, output_dir, log=lambda _x: None):
    raise NotImplementedError("广东报告需要用户提供第五章模板；当前程序不会伪造该模板。")


def _section_overview(doc, config, segments):
    _body(doc, f"本次对重庆市G210线波形梁护栏{'和'.join(['护栏横梁中心高度', '螺栓缺失'])}进行自动化检测，按采集路段汇总表划分为{len(segments)}个统计分段。G210上行K2264K2325文件按原始桩号统计，其他文件按电子修正桩号统计。")
    _caption(doc, "表", "1.1", "G210检测路段情况")
    _table(doc, ["序号", "路线编号", "起点桩号", "终点桩号", "里程(km)", "公路等级", "管理单位"], [[i, "G210", engine.format_station(s["start"]), engine.format_station(s["end"]), f"{s['mileage']:.3f}", s["grade"], s["manager"]] for i, s in enumerate(segments, 1)], True)


def _section_method(doc, height_stats, bolt_stats):
    if height_stats is not None:
        _body(doc, "二波按560、580、620、640mm分档，580≤h≤620mm为合格；三波按657、677、717、737mm分档，677≤h≤717mm为合格。")
    if bolt_stats is not None:
        _body(doc, "螺栓缺失率按缺失数量÷（拼接螺栓数量+连接螺栓数量+缺失数量）×100%计算。")


def _section_height(doc, segments, height_stats, height_records, images, temp_dir, drawing_id=1000):
    _heading(doc, "波形梁护栏横梁中心高度检测结果", 1)
    _heading(doc, "G210线整体情况", 2)
    for kind in ("二波", "三波"):
        total = sum(item["types"][kind]["count"] for item in height_stats)
        if not total:
            continue
        good = sum(item["types"][kind]["bins"][2] for item in height_stats)
        _body(doc, f"G210线{kind}形梁护栏有效检测点共{total}个，横梁中心高度合格率为{good * 100 / total:.2f}%。")
    labels = ["h＜560", "560≤h＜580", "580≤h≤620", "620＜h≤640", "h＞640"]
    rows = []
    for item in height_stats:
        data = item["types"]["二波"] if item["types"]["二波"]["count"] else None
        if data:
            rows.append(["二波", "G210", engine.format_station(item["segment"]["start"]), engine.format_station(item["segment"]["end"]), *[f"{value:.2f}%" for value in data["pcts"]]])
    if rows:
        _caption(doc, "表", "3.1.1", "G210线二波形梁护栏横梁中心高度检测结果")
        _table(doc, ["类型", "路线编号", "起点桩号", "终点桩号", *labels], rows, True)
    labels = ["h＜657", "657≤h＜677", "677≤h≤717", "717＜h≤737", "h＞737"]
    rows = []
    for item in height_stats:
        data = item["types"]["三波"] if item["types"]["三波"]["count"] else None
        if data:
            rows.append(["三波", "G210", engine.format_station(item["segment"]["start"]), engine.format_station(item["segment"]["end"]), *[f"{value:.2f}%" for value in data["pcts"]]])
    if rows:
        _caption(doc, "表", "3.1.2", "G210线三波形梁护栏横梁中心高度检测结果")
        _table(doc, ["类型", "路线编号", "起点桩号", "止点桩号", *labels], rows, True)

    for segment_index, item in enumerate(height_stats):
        segment = item["segment"]
        if not any(item["types"][kind]["count"] for kind in ("二波", "三波")):
            continue
        _heading(doc, f"G210线{engine.format_station(segment['start'])}～{engine.format_station(segment['end'])}段", 2)
        for kind in ("二波", "三波"):
            data = item["types"][kind]
            if not data["count"]:
                continue
            level = "较高" if data["pass"] >= 80 else ("一般" if data["pass"] >= 60 else "偏低")
            detail = engine.percentage_phrases(kind, data["pcts"])
            _body(doc, f"G210线{engine.format_station(segment['start'])}～{engine.format_station(segment['end'])}段{kind}形梁护栏横梁中心高度有效检测点共{data['count']}个，整体合格率{level}，合格率为{data['pass']:.2f}%。护栏横梁中心高度{detail}。")
            image_set = images[(segment_index, kind)]
            drawing_id += 1
            _picture(doc, image_set["line"], 13, 8)
            _caption(doc, "图", f"3.{segment_index + 2}-{1}", f"{kind}形梁护栏横梁中心高度检测结果")
            _picture(doc, image_set["pie"], 14, 8.5)
            _caption(doc, "图", f"3.{segment_index + 2}-{2}", f"{kind}形梁护栏横梁中心高度分布情况")
            example_rows = [record for record in (height_records or []) if record["segment"] == segment_index and record["kind"] == kind]
            example_points = engine.select_height_example_points(example_rows, segment_index, kind)
            if example_points:
                _example_table(doc, example_points)
                _caption(doc, "图", f"3.{segment_index + 2}-{3}", f"{kind}形梁护栏横梁中心高度自动计算示例")
    return drawing_id


def _section_bolt(doc, segments, bolt_stats, bolt_records, disease_image_index, temp_dir):
    _heading(doc, "波形梁护栏螺栓缺失", 1)
    _heading(doc, "G210线整体情况", 2)
    total_splice = sum(item["splice"] for item in bolt_stats)
    total_connection = sum(item["connection"] for item in bolt_stats)
    total_missing = sum(item["missing"] for item in bolt_stats)
    total_rate = engine.bolt_missing_rate(total_splice, total_connection, total_missing)
    _body(doc, f"本次采用波形梁护栏螺栓缺失自动化检测方式，对重庆市G210线波形梁护栏螺栓缺失情况进行统计，共检出拼接螺栓{total_splice}颗，连接螺栓{total_connection}颗，缺失螺栓{total_missing}颗，整体缺失率为{total_rate:.2f}%。")
    overall = [["G210", engine.format_station(item["segment"]["start"]), engine.format_station(item["segment"]["end"]), f"{item['segment']['mileage']:.3f}", item["splice"], item["connection"], item["missing"], f"{item['rate']:.2f}"] for item in bolt_stats]
    overall.append(["合计", "", "", "", total_splice, total_connection, total_missing, f"{total_rate:.2f}"])
    _caption(doc, "表", "4.1.1", "G210线波形梁护栏螺栓缺失检测结果")
    _table(doc, ["路线编号", "起点桩号", "止点桩号", "检测里程（km）", "拼接螺栓（颗）", "连接螺栓（颗）", "缺失数量（颗）", "缺失率（%）"], overall, True)
    for section_number, item in enumerate(bolt_stats, 2):
        segment = item["segment"]
        _heading(doc, f"G210线{engine.format_station(segment['start'])}～{engine.format_station(segment['end'])}段", 2)
        if not item["points"]:
            _body(doc, "本段无波形护栏")
            continue
        _body(doc, f"G210线{engine.format_station(segment['start'])}～{engine.format_station(segment['end'])}段共检出拼接螺栓{item['splice']}颗，连接螺栓{item['connection']}颗，缺失螺栓{item['missing']}颗，缺失率为{item['rate']:.2f}%。")
        _caption(doc, "表", f"4.{section_number}-1", f"G210线{engine.format_station(segment['start'])}～{engine.format_station(segment['end'])}段波形梁护栏螺栓缺失检测结果")
        _table(doc, ["路线编号", "起点桩号", "止点桩号", "里程（km）", "拼接螺栓（颗）", "连接螺栓（颗）", "缺失数量（颗）"], [["G210", engine.format_station(segment["start"]), engine.format_station(segment["end"]), f"{segment['mileage']:.3f}", item["splice"], item["connection"], item["missing"]]], True)
        segment_bolt_rows = [record for record in (bolt_records or []) if record["segment"] == section_number - 2]
        bolt_examples = engine.select_bolt_example_points(segment_bolt_rows, disease_image_index)
        bolt_examples = [example for example in bolt_examples if example.get("image") is not None]
        if bolt_examples:
            for example in bolt_examples:
                image_data, image_extension = engine.read_disease_image(example["image"])
                path = Path(temp_dir) / f"bolt_{section_number}.{image_extension.lstrip('.')}"
                path.write_bytes(image_data)
                _picture(doc, path, 16, 8.5)
                _body(doc, engine.bolt_example_text(example))
            _caption(doc, "图", f"4.{section_number}-1", f"G210线{engine.format_station(segment['start'])}～{engine.format_station(segment['end'])}段波形梁护栏螺栓缺失自动识别示例")


def _section_conclusion(doc, height_stats, bolt_stats):
    _heading(doc, "结论与建议", 1)
    _heading(doc, "结论", 2)
    if height_stats is not None:
        for kind in ("二波", "三波"):
            total = sum(item["types"][kind]["count"] for item in height_stats)
            if not total:
                continue
            good = sum(item["types"][kind]["bins"][2] for item in height_stats)
            _body(doc, f"G210检测路段{kind}形梁护栏有效检测点{total}个，合格点{good}个，整体合格率{good * 100 / total:.2f}%。")
    if bolt_stats is not None:
        total_splice = sum(item["splice"] for item in bolt_stats)
        total_connection = sum(item["connection"] for item in bolt_stats)
        total_missing = sum(item["missing"] for item in bolt_stats)
        total_rate = engine.bolt_missing_rate(total_splice, total_connection, total_missing)
        _body(doc, f"G210检测路段共检出拼接螺栓{total_splice}颗、连接螺栓{total_connection}颗，缺失螺栓{total_missing}颗，整体缺失率为{total_rate:.2f}%。")
    _heading(doc, "建议", 2)
    _body(doc, "建议管养单位优先对横梁中心高度合格率较低的分段开展现场复核，结合路缘石、路面加铺及护栏结构实际情况制定整治计划，并在养护后复测；加强护栏连接件养护巡查，对螺栓缺失位置及时补装同规格螺栓并复核紧固状态。")


def _report_with_skeleton(doc, config, blocks, segments, height_stats, height_records, bolt_stats, bolt_records, disease_image_index, images, temp_dir, skeleton_md=None):
    # 显式锚点优先，关键词回退。显式锚点语法（任一满足即注入）：
    #   <!-- inject:overview -->  <!-- inject:height -->  <!-- inject:bolt -->  <!-- inject:conclusion -->
    # 兼容旧模板的关键词匹配：项目概况/整体情况/螺栓/结论
    import re as _re
    anchor_re = _re.compile(r"<!--\s*inject:\s*(overview|height|bolt|conclusion)\s*-->")
    keyword_map = {
        "项目概况": "overview",
        "整体情况": "height",
        "螺栓": "bolt",
        "结论": "conclusion",
    }
    def _writer_for(key: str):
        if key == "overview":
            return lambda: _section_overview(doc, config, segments)
        if key == "height":
            return (lambda: _section_height(doc, segments, height_stats, height_records, images, temp_dir)) if height_stats is not None else (lambda: None)
        if key == "bolt":
            return (lambda: _section_bolt(doc, segments, bolt_stats, bolt_records, disease_image_index, temp_dir)) if bolt_stats is not None else (lambda: None)
        if key == "conclusion":
            return lambda: _section_conclusion(doc, height_stats, bolt_stats)
        return lambda: None

    # pending 按显式 key 索引，便于锚点直接命中
    pending_by_key = {
        "overview": ("项目概况", True, _writer_for("overview")),
        "height": ("整体情况", False, _writer_for("height")),
        "bolt": ("螺栓", False, _writer_for("bolt")),
        "conclusion": ("结论", False, _writer_for("conclusion")),
    }
    pending_keys = set(pending_by_key.keys())

    def _inject(key: str, render_heading: bool, block_text: str = "", level: int = 2):
        if key not in pending_keys:
            return
        keyword, do_render, writer = pending_by_key[key]
        pending_keys.remove(key)
        if do_render and block_text:
            _heading(doc, block_text, min(level, 2))
        writer()

    skeleton_dir = Path(skeleton_md).parent if skeleton_md else None

    for block in blocks:
        # 1) 显式锚点：任意块（标题/段落）的文本中包含 <!-- inject:xxx -->
        anchor_hit = None
        if block.text:
            m = anchor_re.search(block.text)
            if m:
                anchor_hit = m.group(1)
        if anchor_hit:
            # 显式锚点不渲染原块文本，直接注入
            _inject(anchor_hit, False)
            continue

        if block.kind == "heading":
            # 关键词回退
            hit_key = None
            for kw, key in keyword_map.items():
                if kw in block.text and key in pending_keys:
                    hit_key = key
                    break
            if hit_key:
                _keyword, do_render, _ = pending_by_key[hit_key]
                _inject(hit_key, do_render, block.text, block.level)
                continue
            if "G210线K" in block.text and block.level >= 3:
                continue
            _heading(doc, block.text, min(block.level, 2))
        elif block.kind == "paragraph":
            if block.text:
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
    # 未命中锚点追加末尾（不丢数据但位置错，模板编辑时保留关键词或显式锚点可避免）
    for key in list(pending_keys):
        _, _, writer = pending_by_key[key]
        writer()


def make_report(config, segments, height_stats, height_records, bolt_stats, bolt_records, disease_image_index, temp_dir, log=lambda _x: None, skeleton_md=None):
    """Build the full report document from computed statistics."""
    images = _distribution_images(segments, height_stats, height_records, temp_dir) if height_stats is not None else {}
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.orientation = WD_ORIENT.PORTRAIT
    section.left_margin = section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(engine.PROGRAM_NAME)
    run.bold = True
    run.font.size = Pt(22)
    _set_run_fonts(run, HEADING_FONT)
    title.paragraph_format.space_after = Pt(18)

    if skeleton_md is None:
        raise FileNotFoundError(f"Markdown 模板不存在：{skeleton_md}，仅支持 .md 模板。")
    blocks = markdown_skeleton.read_blocks(skeleton_md)
    _report_with_skeleton(doc, config, blocks, segments, height_stats, height_records, bolt_stats, bolt_records, disease_image_index, images, temp_dir, skeleton_md=skeleton_md)
    try:
        doc.save(config.out_docx)
    except PermissionError as exc:
        raise PermissionError(f"Word文件被占用：{config.out_docx}") from exc
    return config.out_docx



def run(config, segments, height_stats, height_records, bolt_stats, bolt_records, disease_image_index, log=lambda _x: None, skeleton_md=None):
    import tempfile

    with tempfile.TemporaryDirectory(prefix="g210_report_builtin_") as temp_dir:
        config.output_dir.mkdir(parents=True, exist_ok=True)
        log("未检测到内置 Word 模板，切换到程序化报告生成模式。")
        return make_report(config, segments, height_stats, height_records, bolt_stats, bolt_records, disease_image_index, temp_dir, log, skeleton_md=skeleton_md)
