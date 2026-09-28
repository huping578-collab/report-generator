from __future__ import annotations

from collections import Counter, defaultdict
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from queue import Empty, Queue
from threading import Thread
from zipfile import ZIP_DEFLATED, ZipFile
from concurrent.futures import ThreadPoolExecutor, as_completed
import csv
import hashlib
import math
import os
import posixpath
import random
import re
import shutil
import sys
import tempfile
import traceback
import xml.etree.ElementTree as ET

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import openpyxl
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabel, DataLabelList

from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import CharacterProperties, Font as DrawingFont, Paragraph, ParagraphProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from backend import markdown_skeleton

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


X = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
ET.register_namespace("w", W)
ET.register_namespace("r", R)
ET.register_namespace("wp", WP)
ET.register_namespace("a", A)
ET.register_namespace("pic", PIC)
q = lambda ns, tag: f"{{{ns}}}{tag}"
wt = lambda tag: q(W, tag)

# 统计图系列配色（brief-v3 需求 #4）：与甲方模板《…模板-1-2.docx》18 张内嵌图表逐值一致。
# 模板 18 个内嵌 xlsx 的 theme accent1/2/3 = 4F81BD/C0504D/9BBB59（Office 2007-2010 调色板），
# Office 2013 色号（4472C4/ED7D31/A5A5A5）在模板 EMF 内 COLORREF 出现 0 次。
# 证据：reviews/findings-chart-colors.md、requirements/v3-template-profile.md §5.2/§5.5。
# TCI 对比图 #2E75B6(柱) / #92D050(线) 为模板显式 srgbClr，与本组不冲突，保持不变。
GD_SERIES_COLORS = ("4F81BD", "C0504D", "9BBB59")
GD_SERIES_HEX = tuple(f"#{value}" for value in GD_SERIES_COLORS)
PIE_COLORS = [*GD_SERIES_COLORS, "FFC000", "4BACC6"]
PROGRAM_NAME = "报告生成工具V0.1"
OUT_XLSX_NAME = "重庆G210护栏统计.xlsx"
OUT_DOCX_NAME = "重庆G210交安设施检测报告.docx"
GRAY_HEADER_FILL = "D9D9D9"


def county_report_stem(county):
    """区县报告主名：重庆市XX区/县交安设施检测报告（已带重庆市前缀的不重复）。"""
    name = str(county or "").strip()
    if name.startswith("重庆市"):
        name = name[len("重庆市"):]
    return f"重庆市{name}交安设施检测报告"


@dataclass
class GuangdongConfig:
    project_dir: Path
    manual_xlsx: Path
    route_xlsx: Path
    output_dir: Path
    marking_threshold: float
    height_threshold: float
    bolt_threshold: float
    marking_dir: Path | None = None
    guardrail_dir: Path | None = None
    route_name_xlsx: Path | None = None
    manual_before_xlsx: Path | None = None
    # 附件《全省高速公路基础信息表》：唯一带桩号、可按 P2 消歧跨主体路线的经营主体来源。
    # 不给时退到同目录 `路线分类表（含经营主体）.xlsx`（只有 (地市,路线) 级归属）。
    owner_xlsx: Path | None = None

    def __post_init__(self):
        values = validate_thresholds(self.marking_threshold, self.height_threshold, self.bolt_threshold)
        self.marking_threshold, self.height_threshold, self.bolt_threshold = values

    @property
    def thresholds(self):
        return {"marking": self.marking_threshold, "height": self.height_threshold, "bolt": self.bolt_threshold}


def validate_thresholds(marking, height, bolt):
    result = []
    for label, value in zip(("标线", "护栏高度", "螺栓缺失数量差值"), (marking, height, bolt)):
        if value is None or str(value).strip() == "":
            raise ValueError(f"{label}一致性阈值不能为空")
        try:
            number = float(value)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"{label}一致性阈值必须为数值") from exc
        if not 0 <= number <= 100:
            raise ValueError(f"{label}一致性阈值必须在0～100之间")
        result.append(number)
    return tuple(result)


def application_root():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def resource_template_path(name):
    return application_root() / "templates" / name


def builtin_template_paths():
    """返回内置报告模板配置；新增模板时只需在此处增加名称和路径。"""
    return {
        "重庆模板": resource_template_path("重庆项目报告模板.md"),
    }


BUILTIN_REPORT_TEMPLATES = builtin_template_paths()


@dataclass
class Config:
    project_dir: Path
    summary_xlsx: Path
    detail_dir: Path
    template_docx: Path
    output_dir: Path
    disease_dir: Path | None = None
    tci_path: Path | None = None
    county: str | None = None

    @property
    def out_xlsx(self):
        if self.county:
            return self.output_dir / f"{county_report_stem(self.county)}.xlsx"
        return self.output_dir / OUT_XLSX_NAME

    @property
    def out_docx(self):
        if self.county:
            return self.output_dir / f"{county_report_stem(self.county)}.docx"
        return self.output_dir / OUT_DOCX_NAME


def station_to_m(value):
    s = str(value or "").strip().upper().replace("K", "")
    match = re.search(r"(-?\d+(?:\.\d+)?)\s*\+\s*(-?\d+(?:\.\d+)?)", s)
    if match:
        return float(match.group(1)) * 1000 + float(match.group(2))
    try:
        return float(s) * 1000
    except ValueError:
        return None


def format_station(meters):
    km = math.floor(meters / 1000)
    remainder = int(round(meters - km * 1000))
    if remainder >= 1000:
        km, remainder = km + 1, remainder - 1000
    return f"K{km}+{remainder:03d}"


def format_station_one_decimal(meters):
    """桩号米数保留1位小数，用于螺栓示例的原始桩号。"""
    km = math.floor(meters / 1000)
    remainder = round(meters - km * 1000, 1)
    if remainder >= 1000:
        km, remainder = km + 1, remainder - 1000
    return f"K{km}+{remainder:05.1f}"


def read_segments(summary_xlsx):
    wb = openpyxl.load_workbook(summary_xlsx, read_only=True, data_only=True)
    if "各区县项目概况" in wb.sheetnames:
        ws = wb["各区县项目概况"]
    elif "Sheet1" in wb.sheetnames:
        ws = wb["Sheet1"]
    else:
        ws = wb.worksheets[0]
    # 读取 header 行（第2行）映射列名到索引，兼容旧版无区县时默认 county="" route="G210"
    header = None
    header_row_idx = 2
    for idx in (2, 1):
        if ws.max_row >= idx:
            values = [str(c.value).strip() if c.value is not None else "" for c in ws[idx]]
            if any(v for v in values):
                if any("区县" in v or "公路等级" in v or "路线编号" in v for v in values):
                    header = values
                    header_row_idx = idx
                    break
    # 若 header 含 "区县" 则按新列读，否则按旧固定索引
    if header is not None and any("区县" in h for h in header):
        def _col(keywords):
            for i, h in enumerate(header):
                for kw in keywords:
                    if kw and kw in h:
                        return i
            return None

        county_idx = _col(["区县"])
        route_idx = _col(["路线编号"])
        route_name_idx = _col(["路线名"])
        grade_idx = _col(["公路等级"])
        manager_idx = _col(["管理单位", "管养单位"])
        start_idx = _col(["起点桩号", "起点"])
        end_idx = _col(["止点桩号", "终点桩号", "止点", "终点"])
        total_idx = _col(["总里程"])
        mileage_idx = None
        for i, h in enumerate(header):
            if "里程" in h and "总里程" not in h:
                mileage_idx = i
                break

        result = []
        last_county = ""
        last_total = None
        for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
            if row is None or all(v is None for v in row):
                continue
            s_val = row[start_idx] if start_idx is not None and start_idx < len(row) else None
            e_val = row[end_idx] if end_idx is not None and end_idx < len(row) else None
            if s_val is None or e_val is None or str(s_val).strip() == "":
                continue
            try:
                start = float(s_val) * 1000
                end = float(e_val) * 1000
            except (TypeError, ValueError):
                continue
            mileage = 0.0
            if mileage_idx is not None and mileage_idx < len(row) and row[mileage_idx] not in (None, ""):
                try:
                    mileage = float(row[mileage_idx])
                except (TypeError, ValueError):
                    mileage = 0.0
            total = mileage
            if total_idx is not None and total_idx < len(row) and row[total_idx] not in (None, ""):
                try:
                    total = float(row[total_idx])
                except (TypeError, ValueError):
                    total = mileage
            raw_county = str(row[county_idx] or "").strip() if county_idx is not None and county_idx < len(row) else ""
            if raw_county:
                county = raw_county
                last_county = county
            else:
                county = last_county
            route = str(row[route_idx] or "").strip() if route_idx is not None and route_idx < len(row) else "G210"
            if not route:
                route = "G210"
            route_name = str(row[route_name_idx] or "").strip() if route_name_idx is not None and route_name_idx < len(row) else ""
            grade = str(row[grade_idx] or "").strip() if grade_idx is not None and grade_idx < len(row) else ""
            manager = str(row[manager_idx] or "").strip() if manager_idx is not None and manager_idx < len(row) else ""
            result.append({
                "county": county,
                "route": route,
                "route_name": route_name,
                "grade": grade,
                "manager": manager,
                "start": start,
                "end": end,
                "mileage": mileage,
                "total_mileage": total,
            })
        wb.close()
        if not result:
            raise ValueError("分段汇总表中未读取到起点桩号和终点桩号。")
        return result
    result = []
    for row in ws.iter_rows(min_row=header_row_idx + 1 if header is not None else 3, values_only=True):
        if len(row) < 9 or row[6] is None or row[7] is None:
            continue
        try:
            start = float(row[6]) * 1000
            end = float(row[7]) * 1000
        except (TypeError, ValueError):
            continue
        result.append({
            "county": "",
            "route": "G210",
            "route_name": "",
            "grade": str(row[3] or ""),
            "manager": str(row[4] or ""),
            "start": start,
            "end": end,
            "mileage": float(row[8] or 0),
            "total_mileage": float(row[8] or 0),
        })
    wb.close()
    if not result:
        raise ValueError("分段汇总表中未读取到起点桩号和终点桩号。")
    return result


def county_short_names(county):
    """区县全称+简称候选（兼容“渝北=渝北区”口径）。"""
    # ponytail: 只剥一层区/县/自治县后缀；民族自治县前缀缩写需别名表时再加
    full = str(county or "").strip()
    names = {full} if full else set()
    for suffix in ("自治县", "区", "县"):
        if full.endswith(suffix) and len(full) > len(suffix):
            names.add(full[: -len(suffix)])
            break
    return names


def extract_counties_from_filenames(filenames, known_counties):
    """按已知区县在文件名中命中（全称优先、简称其次），保已知顺序去重。"""
    names = [str(item or "") for item in filenames]
    matched = []
    for county in known_counties:
        full = str(county)
        if any(full and full in item for item in names):
            matched.append(county)
            continue
        shorts = county_short_names(county) - {full}
        if any(short and short in item for short in shorts for item in names):
            matched.append(county)
    return matched


def resolve_report_counties(segments, filenames=(), override=None):
    """文件名→区县→划分表分段的路由口径（Q1~Q4 用户已书面确认：无区县手动选、多区县逐个生成、缺行报错、兼容简称）。"""
    known = []
    for segment in segments or []:
        county = str(segment.get("county") or "").strip()
        if county and county not in known:
            known.append(county)
    if not known:
        return []
    if override is not None:
        want = str(override).strip()
        for county in known:
            if want and want in county_short_names(county):
                return [county]
        raise ValueError(f"区县“{want}”在区县划分情况表中无对应路线分段（候选：{'、'.join(known)}），已中断。")
    files = [str(item or "") for item in (filenames or []) if str(item or "").strip()]
    if not files:
        return list(known)
    matched = extract_counties_from_filenames(files, known)
    if not matched:
        raise ValueError(f"文件名中未识别到区县名称，请在“区县”下拉框中手动选择（候选：{'、'.join(known)}），已中断。")
    return matched


def county_route_summary(segments):
    """区县→{路线编号}映射：识别各区县在汇总表中对应的道路编号及分段。

    数据收集后各区县报告按“区县整体→每个区段详情”展开（总结文字、统计表格、
    matplotlib 统计图、病害/螺栓示例图片匹配），分段定位统一走“路线+桩号”
    口径（见 _segment_index）。
    """
    summary = {}
    for segment in segments or []:
        county = str(segment.get("county") or "").strip()
        route = str(segment.get("route") or "").strip() or "G210"
        summary.setdefault(county, {}).setdefault(route, 0)
        summary[county][route] += 1
    return summary


def cell_value(cell, shared_strings=None):
    if cell.get("t") == "inlineStr":
        return "".join(t.text or "" for t in cell.findall(f".//{q(X, 't')}"))
    value = cell.find(q(X, "v"))
    if value is None or value.text is None:
        return None
    if cell.get("t") == "s" and shared_strings is not None:
        try:
            return shared_strings[int(value.text)]
        except (ValueError, IndexError):
            return None
    try:
        return float(value.text)
    except ValueError:
        return value.text


def _xlsx_sheet_files(archive):
    return sorted(
        name for name in archive.namelist()
        if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
    )


def _xlsx_sheet_labels(archive):
    """Map worksheet XML parts to their human-readable workbook sheet names."""
    workbook_name = "xl/workbook.xml"
    rels_name = "xl/_rels/workbook.xml.rels"
    if workbook_name not in archive.namelist() or rels_name not in archive.namelist():
        return {}
    workbook = ET.fromstring(archive.read(workbook_name))
    relationships = ET.fromstring(archive.read(rels_name))
    targets = {relation.get("Id"): relation.get("Target") for relation in relationships}
    labels = {}
    for sheet in workbook.findall(f".//{q(X, 'sheet')}"):
        target = targets.get(sheet.get(q(R, "id")))
        if not target:
            continue
        part = posixpath.normpath(posixpath.join("xl", target)).lstrip("/")
        labels[part] = sheet.get("name", part)
    return labels


def _xml_row_values(element, shared_strings):
    values = {}
    for cell in element.findall(q(X, "c")):
        match = re.match(r"[A-Z]+", cell.get("r", ""))
        if match:
            values[match.group(0)] = cell_value(cell, shared_strings)
    return values


def _iter_sheet_rows_with_header(archive, sheet_name, shared_strings, header_predicate):
    """Yield (Excel row number, row) after the first matching header row.

    Unlike the old first-row reader, this keeps blank header columns in every row,
    which makes an explicitly present remark column distinguishable from an old
    workbook that has no remark column at all.
    """
    headers = None
    with archive.open(sheet_name) as stream:
        for _, element in ET.iterparse(stream, events=("end",)):
            if element.tag != q(X, "row"):
                continue
            row_number = int(element.get("r", "0") or 0)
            values = _xml_row_values(element, shared_strings)
            if headers is None:
                labels = {str(value).strip() for value in values.values() if value is not None}
                if header_predicate(labels):
                    headers = {column: str(value).strip() for column, value in values.items() if value is not None and str(value).strip()}
            elif headers:
                row = {header: None for header in headers.values()}
                for column, value in values.items():
                    row[headers.get(column, column)] = value
                yield row_number, row
            element.clear()


def _iter_source_rows(path, header_predicate):
    """Read the first worksheet whose actual headers identify the requested source."""
    with ZipFile(path) as archive:
        shared_strings = _xlsx_shared_strings(archive)
        labels = _xlsx_sheet_labels(archive)
        for sheet_name in _xlsx_sheet_files(archive):
            rows = _iter_sheet_rows_with_header(archive, sheet_name, shared_strings, header_predicate)
            found = False
            for row_number, row in rows:
                found = True
                yield labels.get(sheet_name, sheet_name), row_number, row
            if found:
                return


def _height_header(labels):
    return "护栏类型" in labels and any("中心高度" in label for label in labels)


def _bolt_header(labels):
    return any("拼接螺栓数量" in label for label in labels) and any("连接螺栓" in label for label in labels)


def _iter_height_rows_with_source(path):
    return _iter_source_rows(path, _height_header)


def iter_height_rows(path):
    for _, _, row in _iter_height_rows_with_source(path):
        yield row


def _iter_bolt_rows_with_source(path):
    return _iter_source_rows(path, _bolt_header)


def iter_bolt_rows(path):
    """Locate the worksheet by its bolt headers and stream its rows."""
    for _, _, row in _iter_bolt_rows_with_source(path):
        yield row


def guardrail_type(value):
    text = str(value or "").strip()
    if "三" in text:
        return "三波"
    if any(word in text for word in ("双", "两", "二")):
        return "二波"
    return None


def remark_marker_allowed(row):
    """Return whether a row is eligible under the 重庆螺栓备注口径.

    New files use ``备注标记``; older files used ``备注`` or had no remark
    column.  A missing column remains compatible, while present values are
    normalized before comparison.
    """
    if "备注标记" in row:
        value = row.get("备注标记")
    elif "备注" in row:
        value = row.get("备注")
    else:
        return True
    return str(value or "").strip() in {"", "无备注"}


HEIGHT_DESIGN = {"二波": 600.0, "三波": 697.0}
HEIGHT_LIMITS = {"二波": (580.0, 620.0), "三波": (677.0, 717.0)}


def height_deviation_over_10cm(kind, height):
    design = HEIGHT_DESIGN.get(kind)
    return design is not None and abs(float(height) - design) > 100


def bolt_missing_ratio(splice, connection, missing):
    denominator = splice + connection + missing
    return missing / denominator if denominator else None


def format_progress(stage, current, total, item=""):
    stage = str(stage or "").strip()
    current, total = int(current), int(total)
    if not stage or total <= 0 or current < 0 or current > total:
        raise ValueError("进度消息参数无效")
    suffix = f" {str(item).strip()}" if str(item or "").strip() else ""
    return f"[progress] {stage} {current}/{total}{suffix}"


_PROGRESS_RE = re.compile(
    r"^\[progress\]\s+(?P<stage>.+?)\s+(?P<current>\d+)/(?P<total>\d+)"
    r"(?:\s+(?P<item>.*))?$"
)


def parse_progress(message):
    match = _PROGRESS_RE.match(str(message or "").strip())
    if not match:
        return None
    current, total = int(match.group("current")), int(match.group("total"))
    if total <= 0 or current > total:
        return None
    return {
        "stage": match.group("stage"),
        "current": current,
        "total": total,
        "item": match.group("item") or "",
    }


HEIGHT_DESIGN = {"二波": 600, "三波": 697}
HEIGHT_PASS_TOLERANCE = 20


def height_limits(kind, tolerance=50):
    """分档边界＝设计值∓容差（外边界）与设计值∓20mm（内边界）；默认±50与现口径一致。"""
    design = HEIGHT_DESIGN.get(kind, 600)
    return (design - tolerance, design - HEIGHT_PASS_TOLERANCE,
            design + HEIGHT_PASS_TOLERANCE, design + tolerance)


def height_bin_labels(kind, tolerance=50):
    low, low_in, high_in, high = height_limits(kind, tolerance)
    return [
        f"h＜{low}", f"{low}≤h＜{low_in}", f"{low_in}≤h≤{high_in}",
        f"{high_in}＜h≤{high}", f"h＞{high}",
    ]


def bin_index(kind, height, limits=None):
    limits = limits or height_limits(kind)
    if height < limits[0]:
        return 0
    if height < limits[1]:
        return 1
    if height <= limits[2]:
        return 2
    if height <= limits[3]:
        return 3
    return 4


def collect_records(segments, detail_dir, log=lambda _: None):
    records, excluded = [], Counter()
    conflicts = Counter()
    files = sorted(Path(detail_dir).glob("*.xlsx"))
    files = [p for p in files if not p.name.startswith("~$")]
    if not files:
        raise FileNotFoundError("明细文件夹中未找到交安设施现场检测Excel。")
    for index, path in enumerate(files, 1):
        if path.name.startswith("G210-G210-"):
            continue
        log(format_progress("解析明细", index, len(files), path.name))
        use_raw = path.name.startswith("G210上行K2264K2325-")
        file_dir = file_direction(path)
        for source_sheet, source_row, row in _iter_height_rows_with_source(path):
            kind = guardrail_type(row.get("护栏类型"))
            height = row.get("梁板中心高度(mm)")
            if kind is None or not isinstance(height, (int, float)) or height <= 0:
                continue
            remark = str(row.get("异常标记") or row.get("备注") or "").strip()
            if remark not in {"", "无备注"}:
                excluded[remark] += 1
                continue
            _count_direction_conflict(conflicts, path, row.get("方向"), file_dir)
            basis = "原始桩号" if use_raw else "电子修正桩号"
            raw_station = station_to_m(row.get("原始桩号"))
            electronic_station = station_to_m(row.get("电子修正桩号"))
            station = raw_station if use_raw else electronic_station
            if station is None:
                continue
            route = row.get("路线编号") or row.get("路线") or ""
            segment = _segment_index(segments, station, route)
            if segment is None:
                continue
            segment_data = segments[segment]
            records.append({
                "file": path.name,
                "source_sheet": source_sheet,
                "source_row": source_row,
                "direction": file_dir,
                "route": _route(route or segment_data.get("route", "")),
                "county": str(segment_data.get("county") or "").strip(),
                "station": station,
                "raw_station": raw_station,
                "electronic_station": electronic_station,
                "basis": basis,
                "kind": kind,
                "height": float(height),
                "segment": segment,
            })
    unique = {}
    for record in records:
        key = (
            record["direction"], round(record["station"], 3),
            record["kind"], round(record["height"], 3),
        )
        unique.setdefault(key, record)
    _log_direction_conflicts(conflicts, log)
    return list(unique.values()), len(records) - len(unique), excluded


def collect_bolt_records(segments, detail_dir, log=lambda _: None):
    records = []
    conflicts = Counter()
    files = sorted(Path(detail_dir).glob("*.xlsx"))
    files = [p for p in files if not p.name.startswith("~$")]
    if not files:
        raise FileNotFoundError("明细文件夹中未找到交安设施现场检测Excel。")

    def number(value):
        try:
            return float(value or 0)
        except (TypeError, ValueError):
            return 0.0

    for index, path in enumerate(files, 1):
        log(format_progress("解析螺栓明细", index, len(files), path.name))
        use_raw = path.name.startswith("G210上行K2264K2325-")
        file_dir = file_direction(path)
        for source_sheet, source_row, row in _iter_bolt_rows_with_source(path):
            if not remark_marker_allowed(row):
                continue
            _count_direction_conflict(conflicts, path, row.get("方向"), file_dir)
            basis = "原始桩号" if use_raw else "电子修正桩号"
            raw_station = station_to_m(row.get("原始桩号"))
            electronic_station = station_to_m(row.get("电子修正桩号"))
            station = raw_station if use_raw else electronic_station
            if station is None:
                continue
            route = row.get("路线编号") or row.get("路线") or ""
            segment = _segment_index(segments, station, route)
            if segment is None:
                continue
            segment_data = segments[segment]
            splice = number(row.get("拼接螺栓数量（颗）"))
            splice_missing = number(row.get("拼接螺栓缺失数量（颗）"))
            connection = number(row.get("连接螺栓数量（颗）"))
            connection_missing = number(row.get("连接螺栓缺失数量（颗）"))
            if splice == connection == splice_missing == connection_missing == 0:
                continue
            records.append({
                "file": path.name,
                "source_sheet": source_sheet,
                "source_row": source_row,
                "direction": file_dir,
                "route": _route(route or segment_data.get("route", "")),
                "county": str(segment_data.get("county") or "").strip(),
                "station": station,
                "raw_station": raw_station,
                "electronic_station": electronic_station,
                "basis": basis,
                "segment": segment,
                "splice": splice,
                "splice_missing": splice_missing,
                "connection": connection,
                "connection_missing": connection_missing,
            })
    unique = {}
    for record in records:
        key = (
            record["direction"], round(record["station"], 3),
            record["splice"], record["splice_missing"],
            record["connection"], record["connection_missing"],
        )
        unique.setdefault(key, record)
    _log_direction_conflicts(conflicts, log)
    return list(unique.values()), len(records) - len(unique)


def normalize_direction(value):
    text = str(value or "")
    if "上" in text:
        return "上行"
    if "下" in text:
        return "下行"
    return text.strip()


def file_direction(path):
    """重庆口径：源文件名含“下行”判为下行，否则上行；不采用表内方向字段。"""
    return "下行" if "下行" in Path(path).name else "上行"


def _count_direction_conflict(conflicts, path, row_direction, file_dir):
    """表内方向字段与文件名判定不一致时计数，统计仍按文件名口径。"""
    row_dir = normalize_direction(row_direction)
    if row_dir in ("上行", "下行") and row_dir != file_dir:
        conflicts[(path.name, file_dir, row_dir)] += 1


def _log_direction_conflicts(conflicts, log):
    for (name, file_dir, row_dir), count in sorted(conflicts.items()):
        log(f"方向冲突提示：{name} 按文件名判为{file_dir}，表内{count}条记为{row_dir}，统计以文件名口径为准。")


def _segment_index(segments, station, route="", county="", direction=""):
    """唯一落段；相邻边界归后段，身份不明且重叠时不猜测。"""
    route = _route(route) if route else ""
    county = str(county or "").strip()
    if county == "重庆市":
        county = ""
    direction = normalize_direction(direction)
    candidates = []
    for index, segment in enumerate(segments):
        start, end = sorted((segment["start"], segment["end"]))
        if not start <= station <= end:
            continue
        if route and _route(segment.get("route", "")) != route:
            continue
        if county and segment.get("county") and county not in county_short_names(segment["county"]):
            continue
        if direction and segment.get("direction") and direction != normalize_direction(segment["direction"]):
            continue
        candidates.append(index)
    interior = [i for i in candidates if station < max(segments[i]["start"], segments[i]["end"])]
    candidates = interior or candidates
    return candidates[0] if len(candidates) == 1 else None


def _xlsx_shared_strings(archive):
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return ["".join(node.text or "" for node in item.findall(f".//{q(X, 't')}")) for item in root]


def _xlsx_sheet_rows(archive, sheet_name, shared_strings):
    rows = {}
    headers = None
    with archive.open(sheet_name) as stream:
        for _, element in ET.iterparse(stream, events=("end",)):
            if element.tag != q(X, "row"):
                continue
            row_number = int(element.get("r", "0") or 0)
            values = {}
            for cell in element.findall(q(X, "c")):
                match = re.match(r"[A-Z]+", cell.get("r", ""))
                if match:
                    values[match.group(0)] = cell_value(cell, shared_strings)
            labels = {str(value).strip() for value in values.values() if value is not None}
            if headers is None and (
                "原始桩号" in labels
                or ("路线编号" in labels and any("缺损" in label for label in labels))
            ):
                headers = {column: str(value).strip() for column, value in values.items() if value is not None and str(value).strip()}
            elif headers is not None:
                row = {header: None for header in headers.values()}
                for column, value in values.items():
                    row[headers.get(column, column)] = value
                rows[row_number] = row
            element.clear()
    return rows


def collect_disease_image_index(disease_dir, log=lambda _: None):
    """Index only bolt-disease photos, retaining workbook/sheet/row identity."""
    disease_dir = Path(disease_dir)
    files = sorted(path for path in disease_dir.glob("*.xlsx") if not path.name.startswith("~$"))
    if not files:
        raise FileNotFoundError("病害清单文件夹中未找到病害清单Excel。")
    image_index = {}
    indexed = 0
    for file_number, path in enumerate(files, 1):
        log(format_progress("索引病害图片", file_number, len(files), path.name))
        with ZipFile(path) as archive:
            shared_strings = _xlsx_shared_strings(archive)
            # 每个 sheet 只解析一次：旧写法在照片行循环里重复解析整张表，4314 行照片 = 4314 次全表解析。
            rows_by_sheet = {}
            for (sheet_name, excel_row), photos in build_disease_image_map(path).items():
                if sheet_name not in rows_by_sheet:
                    rows_by_sheet[sheet_name] = _xlsx_sheet_rows(archive, sheet_name, shared_strings)
                row = rows_by_sheet[sheet_name].get(excel_row, {})
                if "螺栓缺失" not in str(row.get("病害类型") or ""):
                    continue
                raw_station = station_to_m(row.get("原始桩号"))
                if raw_station is None:
                    continue
                try:
                    quantity = int(round(float(row.get("工程量") or 0)))
                except (TypeError, ValueError):
                    quantity = 0
                for media_name, extension in photos:
                    descriptor = {
                        "workbook": path,
                        "media": media_name,
                        "extension": extension,
                        "source_sheet": sheet_name,
                        "source_row": excel_row,
                        "direction": file_direction(path),
                        "route": _route(row.get("路线编号") or row.get("路线") or ""),
                        "county": str(row.get("区县") or row.get("所属区县") or row.get("区域") or "").strip(),
                        "raw_station": raw_station,
                        "quantity": quantity,
                        "role": "bolt",
                    }
                    key = (descriptor["direction"], round(raw_station, 1))
                    image_index.setdefault(key, []).append(descriptor)
                    indexed += 1
    for descriptors in image_index.values():
        descriptors.sort(key=lambda item: (str(item["workbook"]), item["source_sheet"], item["source_row"], item["media"]))
    log(f"病害图片索引完成：{len(files)}个文件，{indexed}张螺栓缺失图片。")
    return image_index


def match_disease_image(record, image_index):
    if not image_index or record.get("raw_station") is None:
        return None
    key = (normalize_direction(record.get("direction")), round(record["raw_station"], 1))
    candidates = [item for item in image_index.get(key, []) if item.get("role") in (None, "bolt")]
    route = _route(record.get("route") or "")
    county = str(record.get("county") or "").strip()
    if route:
        candidates = [item for item in candidates if not item.get("route") or _route(item.get("route")) == route]
    if county:
        candidates = [item for item in candidates if not item.get("county") or item.get("county") in county_short_names(county)]
    if not candidates:
        return None
    total_missing = int(round(record["splice_missing"] + record["connection_missing"]))
    return next((item for item in candidates if item["quantity"] == total_missing), candidates[0])


def read_disease_image(descriptor):
    with ZipFile(descriptor["workbook"]) as archive:
        data = archive.read(descriptor["media"])
    if data.startswith(b"\x89PNG\x0d\x0a\x1a\x0a"):
        return data, ".png"
    if data.startswith(b"\xff\xd8\xff"):
        return data, ".jpeg"
    extension = Path(descriptor["media"]).suffix.lower() or ".png"
    return data, extension


def _xlsx_header_columns(archive, sheet_name, shared_strings, names):
    wanted = tuple(names or ())
    with archive.open(sheet_name) as stream:
        for _, element in ET.iterparse(stream, events=("end",)):
            if element.tag != q(X, "row"):
                continue
            values = _xml_row_values(element, shared_strings)
            if any(any(name in str(value or "") for name in wanted) for value in values.values()):
                element.clear()
                return {
                    column: str(value).strip()
                    for column, value in values.items()
                    if value is not None and str(value).strip()
                }
            element.clear()
    return {}


def _anchored_image_map(path, photo_column, workbook_sheet_names=None, photo_headers=()):
    """解析工作簿各 sheet 的浮动图片锚点，建立 (sheet名, excel行号) -> [(media部件名, 扩展名)] 映射。

    - 优先按实际图片表头定位图片列；没有可用图片表头时兼容旧的固定列。
    - 跳过表头行（from.row==0，含装饰图/logo）。
    - 值为媒体部件名（懒读，不加载图片字节），调用方按需从同一工作簿读取。
    - workbook_sheet_names: 可选 {sheet部件名: 显示sheet名}，缺省用 sheetN 部件名。
    """
    path = Path(path)
    result = {}
    with ZipFile(path) as archive:
        sheet_names = _xlsx_sheet_files(archive)
        for sheet_name in sheet_names:
            sheet_rels_name = posixpath.join(
                posixpath.dirname(sheet_name), "_rels", posixpath.basename(sheet_name) + ".rels"
            )
            if sheet_rels_name not in archive.namelist():
                continue
            sheet_relationships = ET.fromstring(archive.read(sheet_rels_name))
            drawing_relation = next(
                (relation for relation in sheet_relationships if relation.get("Type", "").endswith("/drawing")),
                None,
            )
            if drawing_relation is None:
                continue
            drawing_name = posixpath.normpath(posixpath.join(posixpath.dirname(sheet_name), drawing_relation.get("Target"))).lstrip("/")
            drawing_rels_name = posixpath.join(
                posixpath.dirname(drawing_name), "_rels", posixpath.basename(drawing_name) + ".rels"
            ).lstrip("/")
            if drawing_name not in archive.namelist() or drawing_rels_name not in archive.namelist():
                continue
            drawing_relationships = ET.fromstring(archive.read(drawing_rels_name))
            media_by_rid = {
                relation.get("Id"): posixpath.normpath(
                    posixpath.join(posixpath.dirname(drawing_name), relation.get("Target"))
                )
                for relation in drawing_relationships
                if relation.get("Type", "").endswith("/image")
            }
            header_columns = _xlsx_header_columns(archive, sheet_name, _xlsx_shared_strings(archive), photo_headers)
            dynamic_column = next(
                (
                    index
                    for index, label in (
                        (ord(column) - ord("A"), value)
                        for column, value in header_columns.items()
                        if len(column) == 1 and column.isalpha()
                    )
                    if any(header in label for header in photo_headers)
                ),
                None,
            )
            # Column letters can be multi-letter; calculate the zero-based index.
            if dynamic_column is None:
                for column, label in header_columns.items():
                    if any(header in label for header in photo_headers):
                        dynamic_column = 0
                        for char in column:
                            dynamic_column = dynamic_column * 26 + ord(char.upper()) - ord("A") + 1
                        dynamic_column -= 1
                        break
            drawing = ET.fromstring(archive.read(drawing_name))
            anchors = []
            for anchor in list(drawing):
                anchor_from = anchor.find(q(XDR, "from"))
                if anchor_from is None:
                    continue
                row_node = anchor_from.find(q(XDR, "row"))
                col_node = anchor_from.find(q(XDR, "col"))
                blip = anchor.find(f".//{q(A, 'blip')}")
                if row_node is None or col_node is None or blip is None:
                    continue
                row_index = int(row_node.text or 0)
                col_index = int(col_node.text or 0)
                if row_index <= 0:
                    continue
                media_name = media_by_rid.get(blip.get(q(R, "embed")))
                if not media_name:
                    continue
                # openpyxl 写绝对 Target（/xl/media/...），Excel 写相对路径；统一去前导斜杠
                media_name = media_name.lstrip("/")
                if media_name not in archive.namelist():
                    continue
                extension = Path(media_name).suffix.lower() or ".png"
                label = sheet_name
                if workbook_sheet_names:
                    label = workbook_sheet_names.get(sheet_name, sheet_name)
                excel_row = row_index + 1
                anchors.append((row_index, col_index, media_name, extension))
            columns = {col for _, col, _, _ in anchors if dynamic_column is not None and col == dynamic_column}
            if not columns and photo_column is not None:
                columns = {photo_column}
            for row_index, col_index, media_name, extension in anchors:
                if col_index not in columns:
                    continue
                result.setdefault((label, row_index + 1), []).append((media_name, extension))
    return result


def build_disease_image_map(disease_path):
    """建立病害明细工作簿的 (sheet名, excel行号) -> [(媒体部件, 扩展名)] 图片映射（病害照片列）。"""
    return _anchored_image_map(disease_path, photo_column=12, photo_headers=("病害照片", "照片", "图片"))


def build_tci_image_map(tci_path):
    """建立 TCI 工作簿的 (sheet名, excel行号) -> [(媒体部件, 扩展名)] 图片映射（图片列）。"""
    return _anchored_image_map(tci_path, photo_column=15, photo_headers=("图片", "照片", "病害照片"))


def read_media(path, media_name):
    """从工作簿读取媒体部件字节并识别扩展名。"""
    with ZipFile(path) as archive:
        data = archive.read(media_name)
    if data.startswith(b"\x89PNG\x0d\x0a\x1a\x0a"):
        return data, ".png"
    if data.startswith(b"\xff\xd8\xff"):
        return data, ".jpeg"
    return data, Path(media_name).suffix.lower() or ".png"


def disease_station_photo_index(disease_path, image_map, role=None):
    """Index disease photos by station, with source and record-role identity."""
    disease_path = Path(disease_path)
    result = {}
    with ZipFile(disease_path) as archive:
        shared_strings = _xlsx_shared_strings(archive)
        sheet_names = _xlsx_sheet_files(archive)
        for sheet_name in sheet_names:
            if not any(key[0] == sheet_name for key in image_map):
                continue
            rows = _xlsx_sheet_rows(archive, sheet_name, shared_strings)
            for excel_row, row in rows.items():
                photos = image_map.get((sheet_name, excel_row))
                if not photos:
                    continue
                disease_type = str(row.get("病害类型") or "")
                if role == "bolt" and "螺栓" not in disease_type:
                    continue
                if role == "height" and ("螺栓" in disease_type or "高度" not in disease_type):
                    continue
                direction = file_direction(disease_path)
                route = _route(row.get("路线编号") or row.get("路线") or "")
                county = str(row.get("区县") or row.get("所属区县") or row.get("区域") or "").strip()
                for column in ("原始桩号", "电子修正桩号"):
                    station = station_to_m(row.get(column))
                    if station is not None:
                        for media_name, extension in photos:
                            result.setdefault((direction, round(station, 1)), []).append({
                                "workbook": disease_path,
                                "media": media_name,
                                "extension": extension,
                                "source_sheet": sheet_name,
                                "source_row": excel_row,
                                "direction": direction,
                                "route": route,
                                "county": county,
                                "station": station,
                                "role": role or "disease",
                            })
    return result


def tci_type_photo_index(tci_path, image_map, routes=None):
    """TCI 工作簿：{病害类型: [(媒体部件, 扩展名)]}，每类取第一条有图记录。"""
    tci_path = Path(tci_path)
    with ZipFile(tci_path) as archive:
        shared_strings = _xlsx_shared_strings(archive)
        rows_by_sheet = {sheet_name: _xlsx_sheet_rows(archive, sheet_name, shared_strings) for sheet_name in _xlsx_sheet_files(archive)}
    type_columns = [
        ("防护设施缺损（处）", "防护设施缺损"),
        ("标志缺损（处）", "交通标志缺损"),
        ("标线缺损（m）", "交通标线缺损"),
    ]
    result = {}
    route_names = {_route(route) for route in routes or ()}
    for (sheet, excel_row), photos in image_map.items():
        row = rows_by_sheet.get(sheet, {}).get(excel_row)
        if not row:
            continue
        if routes:
            row_route = _route(row.get("路线编号"))
            if row_route not in route_names:
                continue
        for column, label in type_columns:
            value = row.get(column)
            if value is not None:
                try:
                    positive = float(value) > 0
                except (TypeError, ValueError):
                    positive = False
                if positive:
                    result.setdefault(label, photos)
    return result


def build_segment_tci_photos(tci_path, image_map, segments):
    """TCI 工作簿：段序号 -> [(病害描述, 媒体部件, 扩展名)]（仅含图片的行）。

    按 路线+方向+桩号（原始/电子修正）落在 segments 区间内匹配；病害描述取
    「防护设施缺损/交通标志缺损/交通标线缺损」三列中第一个非空值（具体病害名，
    如“反光膜污染”“标志板遮挡”）。
    """
    tci_path = Path(tci_path)
    if not segments or not image_map:
        return {}
    with ZipFile(tci_path) as archive:
        shared_strings = _xlsx_shared_strings(archive)
        rows_by_sheet = {sheet_name: _xlsx_sheet_rows(archive, sheet_name, shared_strings) for sheet_name in _xlsx_sheet_files(archive)}
    desc_columns = ("防护设施缺损", "交通标志缺损", "交通标线缺损")
    result = {}
    for (sheet, excel_row), photos in image_map.items():
        row = rows_by_sheet.get(sheet, {}).get(excel_row)
        if not row:
            continue
        direction = file_direction(tci_path)
        station = None
        for column in ("电子修正桩号", "原始桩号"):
            station = station_to_m(row.get(column))
            if station is not None:
                break
        if station is None:
            continue
        description = ""
        for column in desc_columns:
            value = str(row.get(column) or "").strip()
            if value:
                description = value
                break
        route = _route(row.get("路线编号"))
        county = row.get("区县") or row.get("所属区县") or row.get("区域") or ""
        index = _segment_index(segments, station, route, county, direction)
        if index is not None:
            for media_name, extension in photos:
                result.setdefault((index, direction), []).append((description or "病害", media_name, extension))
    return result


def bolt_missing_rate(splice, connection, missing):
    """螺栓缺失率：缺失数÷（现有拼接数+现有连接数+缺失数）×100%。"""
    ratio = bolt_missing_ratio(splice, connection, missing)
    return ratio * 100 if ratio is not None else None


DIRECTIONS = ("上行", "下行")


def _group_by_segment(records):
    """按区段一次性归组，避免每个区段重复扫描全部记录。"""
    grouped = {}
    for record in records:
        grouped.setdefault(record["segment"], []).append(record)
    return grouped


def _bolt_totals(rows):
    splice = sum(record["splice"] for record in rows)
    connection = sum(record["connection"] for record in rows)
    missing = sum(record["splice_missing"] + record["connection_missing"] for record in rows)
    return {
        "splice": int(round(splice)),
        "connection": int(round(connection)),
        "missing": int(round(missing)),
        "rate": bolt_missing_rate(splice, connection, missing),
        "points": len(rows),
    }


def merge_bolt_totals(items):
    """合并若干螺栓统计（按数量求和后重算缺失率）。"""
    splice = sum(item["splice"] for item in items)
    connection = sum(item["connection"] for item in items)
    missing = sum(item["missing"] for item in items)
    return {
        "splice": splice, "connection": connection, "missing": missing,
        "rate": bolt_missing_rate(splice, connection, missing),
        "points": sum(item.get("points", 0) for item in items),
    }


def make_bolt_stats(segments, records):
    grouped = _group_by_segment(records)
    stats = []
    for index, segment in enumerate(segments):
        rows = grouped.get(index, [])
        by_direction = {
            direction: _bolt_totals([record for record in rows if record["direction"] == direction])
            for direction in DIRECTIONS
        }
        stats.append({"segment": segment, "by_direction": by_direction, **_bolt_totals(rows)})
    return stats


def _height_type_bins(rows, kind, tolerance=50, pass_bins=(2,)):
    """分档统计；pass_bins 决定合格率取哪几档（默认取中间档＝设计值±20mm）。"""
    heights = [record["height"] for record in rows if record["kind"] == kind]
    limits = height_limits(kind, tolerance)
    bins = [0] * 5
    for height in heights:
        bins[bin_index(kind, height, limits)] += 1
    percentages = [round(value * 100 / len(heights), 2) if heights else 0 for value in bins]
    passed = sum(bins[index] for index in pass_bins)
    return {
        "count": len(heights), "bins": bins,
        "pcts": percentages,
        "pass": round(passed * 100 / len(heights), 2) if heights else 0,
    }


def merge_height_types(items):
    """合并若干 {kind: 分档} 统计（按点数求和后重算占比/合格率）。"""
    merged = {}
    for kind in ("二波", "三波"):
        count = sum(item[kind]["count"] for item in items)
        bins = [sum(item[kind]["bins"][index] for item in items) for index in range(5)]
        percentages = [round(bins[index] * 100 / count, 2) if count else 0 for index in range(5)]
        merged[kind] = {
            "count": count, "bins": bins,
            "pcts": percentages, "pass": percentages[2] if count else 0,
        }
    return merged


def make_stats(segments, records):
    grouped = _group_by_segment(records)
    stats = []
    for segment_index, segment in enumerate(segments):
        rows = grouped.get(segment_index, [])
        by_direction = {
            direction: {
                kind: _height_type_bins([record for record in rows if record["direction"] == direction], kind)
                for kind in ("二波", "三波")
            }
            for direction in DIRECTIONS
        }
        stats.append({
            "segment": segment,
            "types": merge_height_types(list(by_direction.values())),
            "by_direction": by_direction,
        })
    return stats


# --- TCI 沿线设施技术状况 (JTG 5210-2018 3 项: 防护轻10/重30 w0.25, 标志20 w0.25, 标线0.1/m 每10m1分不足10m计10m w0.20, /0.7) ---
TCI_WEIGHTS = (0.25, 0.25, 0.20)
TCI_DIVISOR = 0.7
TCI_DEDUCT_LIGHT = 10
TCI_DEDUCT_HEAVY = 30
TCI_DEDUCT_SIGN = 20
TCI_DEDUCT_MARKING_PER_M = 0.1

def tci_gd(light: int, heavy: int, sign: int, marking_m: float):
    gd1 = min(100, light * TCI_DEDUCT_LIGHT + heavy * TCI_DEDUCT_HEAVY)
    gd2 = min(100, sign * TCI_DEDUCT_SIGN)
    gd3 = 0
    if marking_m:
        gd3 = min(100, __import__('math').ceil(float(marking_m) / 10.0) * 1)
    return gd1, gd2, gd3

def compute_tci(light: int, heavy: int, sign: int, marking_m: float) -> float:
    gd1, gd2, gd3 = tci_gd(light, heavy, sign, marking_m)
    return round((TCI_WEIGHTS[0] * (100 - gd1) + TCI_WEIGHTS[1] * (100 - gd2) + TCI_WEIGHTS[2] * (100 - gd3)) / TCI_DIVISOR, 4)

def tci_grade(tci: float) -> str:
    if tci >= 90:
        return "优"
    if tci >= 80:
        return "良"
    if tci >= 70:
        return "中"
    if tci >= 60:
        return "次"
    return "差"

def _tci_station_value(row: dict):
    for key in ("电子修正桩号", "标注修正桩号", "原始桩号"):
        v = row.get(key)
        if v not in (None, ""):
            m = station_to_m(v)
            if m is not None:
                return m
    for key in ("电子修正桩号", "标注修正桩号"):
        v = row.get(key)
        try:
            if v is not None and str(v).strip():
                return float(str(v).strip()) * 1000
        except:
            pass
    return None

def collect_tci_records(segments, tci_path, log=lambda _: None):
    records = []
    conflicts = Counter()
    if tci_path is None:
        return records
    pth = Path(tci_path)
    files = []
    if pth.is_file():
        files = [pth]
    elif pth.is_dir():
        files = sorted(pth.glob("*.xlsx"))
        files = [f for f in files if not f.name.startswith("~$")]
    else:
        log(f"TCI 路径不存在：{pth}")
        return records
    for file_number, fpath in enumerate(files, 1):
        log(format_progress("解析TCI", file_number, len(files), fpath.name))
        file_dir = file_direction(fpath)
        try:
            wb = openpyxl.load_workbook(fpath, data_only=True, read_only=True)
        except Exception as e:
            log(f"TCI 读取失败 {fpath.name}: {e}")
            continue
        data_sheet_names = []
        for candidate_name in wb.sheetnames:
            candidate_ws = wb[candidate_name]
            candidate_head = []
            for candidate_row in candidate_ws.iter_rows(min_row=1, max_row=6, values_only=True):
                candidate_head.extend(str(value).strip() for value in candidate_row if value is not None)
            if "路线编号" in candidate_head and any("缺损" in value for value in candidate_head):
                data_sheet_names.append(candidate_name)
        sheet_name = (
            "病害明细表" if "病害明细表" in data_sheet_names
            else (data_sheet_names[0] if data_sheet_names else wb.sheetnames[0])
        )
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        if not rows:
            continue
        header_row_idx = None
        col_map = {}
        route_idx = None
        for i, row in enumerate(rows[:6], 1):
            vals = [str(v).strip() if v is not None else "" for v in row]
            if "区域" in vals and "路线编号" in vals:
                header_row_idx = i
                route_idx = next(
                    (idx for idx, value in enumerate(vals) if "路线编号" in value or value == "路线"),
                    None,
                )
                if i < len(rows):
                    sub = [str(v).strip() if v is not None else "" for v in rows[i]]
                    for idx, (h, s) in enumerate(zip(vals, sub)):
                        if s == "轻":
                            col_map[idx] = "light"
                        elif s == "重":
                            col_map[idx] = "heavy"
                    for idx, h in enumerate(vals):
                        if "标志缺损" in h:
                            col_map[idx] = "sign"
                        elif "标线缺损" in h:
                            col_map[idx] = "marking"
                        elif "电子修正桩号" in h:
                            col_map[idx] = "electron"
                        elif "标注修正桩号" in h:
                            col_map[idx] = "annotation"
                        elif "原始桩号" in h:
                            col_map[idx] = "original"
                break
        if header_row_idx is None:
            for i, row in enumerate(rows[:5], 1):
                vals = [str(v).strip() if v is not None else "" for v in row]
                if any("防护设施缺损" in v for v in vals):
                    header_row_idx = i
                    route_idx = next(
                        (idx for idx, value in enumerate(vals) if "路线编号" in value or value == "路线"),
                        None,
                    )
                    for idx, h in enumerate(vals):
                        if "防护" in h and "轻" in h:
                            col_map[idx] = "light"
                        elif "防护" in h and "重" in h:
                            col_map[idx] = "heavy"
                        elif "标志缺损" in h:
                            col_map[idx] = "sign"
                        elif "标线缺损" in h:
                            col_map[idx] = "marking"
                        elif "电子修正桩号" in h:
                            col_map[idx] = "electron"
                        elif "标注修正桩号" in h:
                            col_map[idx] = "annotation"
                        elif "原始桩号" in h:
                            col_map[idx] = "original"
                    break
        if header_row_idx is None:
            header_row_idx = 1
        route_value = ""
        if route_idx is None:
            route_match = re.search(r"([GSHXY]\d+)", fpath.name, re.IGNORECASE)
            route_value = route_match.group(1) if route_match else ""
        start = header_row_idx + 1
        if header_row_idx < len(rows) and "轻" in "".join(str(x) for x in rows[header_row_idx] if x is not None):
            start = header_row_idx + 2
        for excel_row, row in enumerate(rows[start-1:], start):
            vals = list(row)
            if not any(v not in (None, "") for v in vals):
                continue
            station = None
            for k in ("electron", "annotation", "original"):
                idxs = [i for i,c in col_map.items() if c==k]
                for idx in idxs:
                    if idx < len(vals) and vals[idx] not in (None, ""):
                        station = station_to_m(vals[idx])
                        if station is None:
                            try:
                                station = float(str(vals[idx]).strip())*1000
                            except:
                                station=None
                        if station is not None:
                            break
                if station is not None:
                    break
            if station is None:
                continue
            def num_at(typ):
                idxs = [i for i,c in col_map.items() if c==typ]
                for idx in idxs:
                    if idx < len(vals) and vals[idx] not in (None, ""):
                        try:
                            return float(str(vals[idx]).strip())
                        except:
                            return 0
                return 0
            if not any(c in col_map.values() for c in ("light","heavy","sign","marking")):
                try:
                    light = float(str(vals[8]).strip()) if len(vals)>8 and vals[8] not in (None,"") else 0
                except:
                    light=0
                try:
                    heavy = float(str(vals[9]).strip()) if len(vals)>9 and vals[9] not in (None,"") else 0
                except:
                    heavy=0
                try:
                    sign = float(str(vals[10]).strip()) if len(vals)>10 and vals[10] not in (None,"") else 0
                except:
                    sign=0
                try:
                    marking = float(str(vals[11]).strip()) if len(vals)>11 and vals[11] not in (None,"") else 0
                except:
                    marking=0
            else:
                light = num_at("light")
                heavy = num_at("heavy")
                sign = num_at("sign")
                marking = num_at("marking")
            row_route = vals[route_idx] if route_idx is not None and route_idx < len(vals) else route_value
            headers = [str(v or "").strip() for v in rows[header_row_idx - 1]]
            def identity_value(names):
                index = next((i for i, h in enumerate(headers) if h in names), None)
                return str(vals[index] or "").strip() if index is not None and index < len(vals) else ""
            county = identity_value(("区县", "所属区县", "区域"))
            _count_direction_conflict(conflicts, fpath, identity_value(("方向", "行驶方向")), file_dir)
            direction = file_dir
            seg_idx = _segment_index(segments, station, row_route, county, direction)
            if seg_idx is None:
                continue
            if any(not math.isfinite(v) or v < 0 for v in (light, heavy, sign, marking)):
                raise ValueError(f"{fpath.name}：桩号{format_station(station)}的 TCI 病害数量必须为非负有限数值")
            records.append({"segment": seg_idx, "station": station, "county": segments[seg_idx].get("county", ""),
                            "route": row_route, "direction": direction, "file": str(fpath),
                            "source_sheet": sheet_name, "source_row": excel_row,
                            "light": int(light), "heavy": int(heavy), "sign": int(sign), "marking": float(marking)})
        log(f"TCI 病害 {fpath.name}: 落段 {len([r for r in records if True])} 条")
    _log_direction_conflicts(conflicts, log)
    return records

def make_tci_stats(segments, tci_records):
    """在原检测区段内按整公里、方向评定；汇总直接平均单元而非段均值。"""
    stats = []
    for idx, seg in enumerate(segments):
        rows = [r for r in tci_records if r["segment"] == idx]
        totals = {key: sum(r[key] for r in rows) for key in ("light", "heavy", "sign", "marking")}
        directions = {normalize_direction(r.get("direction")) for r in rows}

        units = []
        start, end = sorted((seg["start"], seg["end"]))
        bounds = [start, *range((math.floor(start / 1000) + 1) * 1000, math.ceil(end / 1000) * 1000, 1000), end]
        for direction in ("上行", "下行", ""):
            if direction not in directions or start == end:
                continue
            for left, right in zip(bounds, bounds[1:]):
                selected = [r for r in rows if normalize_direction(r.get("direction")) == direction
                            and left <= r["station"] and (r["station"] < right or r["station"] == right == end)]
                values = {key: sum(r[key] for r in selected) for key in totals}
                score = compute_tci(**dict(light=values["light"], heavy=values["heavy"], sign=values["sign"], marking_m=values["marking"]))
                units.append(dict(county=seg.get("county", ""), route=seg.get("route", "G210"), direction=direction,
                                  start=left, end=right, mileage=(right-left)/1000, source_segment=idx,
                                  **values, tci=score, grade=tci_grade(score), count=len(selected)))
        score = sum(u["tci"] for u in units) / len(units) if units else None
        stats.append(dict(segment=seg, **totals, units=units, tci=score,
                          grade=tci_grade(score) if score is not None else "", count=len(rows)))
    return stats


def tci_route_stats(segments, tci_stats):
    """县+路线+方向汇总，保留全部原区段单元以供正文、附表和 Excel 共用。"""
    groups = {}
    for stat in tci_stats or []:
        for unit in stat.get("units", []):
            key = (unit["county"], unit["route"], unit["direction"])
            groups.setdefault(key, []).append(unit)
    result = []
    for (county, route, direction), units in sorted(groups.items(), key=lambda item: (item[0][0], item[0][1], 0 if item[0][2] == "上行" else 1)):
        units = sorted(units, key=lambda u: (u["start"], u["end"]))
        score = sum(u["tci"] for u in units) / len(units)
        result.append(dict(county=county, route=route, direction=direction, units=units,
                           mileage=sum(u["mileage"] for u in units), tci=score, grade=tci_grade(score)))
    return result


def route_report_data(segments, height_stats, height_records, bolt_stats, bolt_records):
    """重庆正文按县、道路编号汇总；原分段仅作为来源，间断里程不补齐。"""
    groups = {}
    for index, segment in enumerate(segments):
        groups.setdefault((segment.get("county", ""), segment.get("route", "G210")), []).append(index)
    routes, heights, bolts, mapping = [], [], [], {}
    for route_index, indices in enumerate(groups.values()):
        source = [segments[i] for i in indices]
        route = dict(source[0], start=min(min(s["start"], s["end"]) for s in source),
                     end=max(max(s["start"], s["end"]) for s in source),
                     mileage=sum(s.get("mileage", 0) for s in source), source_segments=source)
        routes.append(route)
        mapping.update({i: route_index for i in indices})
        if height_stats is not None:
            types = {}
            for kind in ("二波", "三波"):
                data = [height_stats[i]["types"][kind] for i in indices if height_stats[i] is not None]
                count = sum(d["count"] for d in data)
                bins = [sum(d["bins"][j] for d in data) for j in range(5)]
                pcts = [round(n * 100 / count, 2) if count else 0 for n in bins]
                types[kind] = dict(count=count, bins=bins, pcts=pcts, **{"pass": pcts[2]})
            heights.append(dict(segment=route, types=types))
        if bolt_stats is not None:
            data = [bolt_stats[i] for i in indices if bolt_stats[i] is not None]
            values = {key: sum(d[key] for d in data) for key in ("splice", "connection", "missing", "points")}
            bolts.append(dict(segment=route, **values, rate=bolt_missing_rate(values["splice"], values["connection"], values["missing"])))
    def remap(records):
        return [dict(r, segment=mapping[r["segment"]]) for r in records or []]
    return routes, heights if height_stats is not None else None, remap(height_records), bolts if bolt_stats is not None else None, remap(bolt_records)


def style_sheet(ws, widths):
    thin = Side(style="thin", color="808080")
    for row in ws.iter_rows():
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="000000")
        cell.fill = PatternFill("solid", fgColor=GRAY_HEADER_FILL)
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def add_interval_sheets(out_path, summary_path, log=lambda _: None, segments=None):
    """按给定分段（默认重读汇总表全量）展开区间明细表。

    区县子报告必须传入该区县的 segments，使区间表收敛到本区县/本区段；
    无区县维度（旧汇总表）或“仅更新区间明细”模式仍走全量汇总表。
    """
    if segments is None:
        segments = read_segments(summary_path)
    wb = openpyxl.load_workbook(out_path)
    for name in list(wb.sheetnames):
        if name.startswith("区间"):
            del wb[name]
    detail = wb["检测明细"]
    headers = {cell.value: cell.column for cell in detail[1]}
    grouped = {(format_station(s["start"]), format_station(s["end"])): [] for s in segments}
    for row in detail.iter_rows(min_row=2, values_only=True):
        key = (row[headers["所属分段起点"] - 1], row[headers["所属分段终点"] - 1])
        if key in grouped:
            normalized_kind = guardrail_type(row[headers["护栏类型"] - 1]) or row[headers["护栏类型"] - 1]
            grouped[key].append((
                row[headers["统计桩号"] - 1], normalized_kind,
                row[headers["梁板中心高度(mm)"] - 1],
            ))
    columns = [
        "序号", "桩号", "护栏类型", "梁板中心高度(mm)",
        "标准值（580mm）", "标准值（620mm）", "标准值（677mm）", "标准值（717mm）",
    ]
    thin = Side(style="thin", color="808080")
    for index, (start, end) in enumerate(grouped, 1):
        ws = wb.create_sheet(f"区间{index:02d}_{start[1:]}-{end[1:]}")
        ws.append(columns)
        rows = sorted(grouped[(start, end)], key=lambda item: (item[0], item[1]))
        for sequence, (station, kind, height) in enumerate(rows, 1):
            ws.append([sequence, station, kind, height, 580, 620, 677, 717])
        for row in ws.iter_rows():
            for cell in row:
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="000000")
            cell.fill = PatternFill("solid", fgColor=GRAY_HEADER_FILL)
        for col, width in enumerate([9, 18, 14, 23, 19, 19, 19, 19], 1):
            ws.column_dimensions[get_column_letter(col)].width = width
        ws.freeze_panes, ws.sheet_view.showGridLines = "A2", False
        ws["J1"], ws["K1"], ws["J2"], ws["K2"] = "统计区间", f"{start}～{end}", "有效记录数", len(rows)
        ws["J1"].font = ws["J2"].font = Font(bold=True)
        if rows:
            table = Table(displayName=f"IntervalDetail{index:02d}", ref=f"A1:H{len(rows)+1}")
            table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
            ws.add_table(table)
    wb.save(out_path)
    log(f"区间明细表已更新：{len(segments)}个区间，共{sum(map(len, grouped.values()))}条记录。")


def make_excel(
    config, segments,
    height_stats=None, height_records=None, height_duplicates=0, excluded=None,
    bolt_stats=None, bolt_records=None, bolt_duplicates=0,
    tci_stats=None, tci_records=None,
    log=lambda _: None,
):
    height_records = height_records or []
    bolt_records = bolt_records or []
    tci_records = tci_records or []
    excluded = excluded or Counter()
    has_county = bool(segments and any(s.get("county") for s in segments))
    # 各表末尾的“合计”行按区县分组；无区县维度时退化为一行全部路线。
    county_groups = list(dict.fromkeys(str(s.get("county") or "") for s in segments)) if has_county else [""]

    def _county_indices(county: str):
        return [index for index, segment in enumerate(segments)
                if (str(segment.get("county") or "") if has_county else "") == county]

    def _county_of(segment_index):
        """记录所属区县；索引越界返回 None（不参与任何区县合计）。"""
        if not isinstance(segment_index, int) or not 0 <= segment_index < len(segments):
            return None
        return str(segments[segment_index].get("county") or "") if has_county else ""

    wb = Workbook()
    wb.remove(wb.active)
    if height_stats is not None:
        for title, kind in (("二波统计", "二波"), ("三波统计", "三波")):
            labels = height_bin_labels(kind)
            ws = wb.create_sheet(title)
            if has_county:
                ws.append(["序号", "区县", "路线编号", "公路等级", "管理单位", "起点桩号", "终点桩号", "统计口径", "检测点数", *labels, "合格率"])
            else:
                ws.append(["序号", "路线编号", "公路等级", "管理单位", "起点桩号", "终点桩号", "统计口径", "检测点数", *labels, "合格率"])
            sequence = 0
            for segment_index, item in enumerate(height_stats):
                data = item["types"][kind]
                if not data["count"]:
                    continue
                sequence += 1
                segment = item["segment"]
                bases = {r["basis"] for r in height_records if r["segment"] == segment_index and r["kind"] == kind}
                basis = next(iter(bases)) if len(bases) == 1 else "原始桩号+电子修正桩号（按来源文件分别采用）"
                if has_county:
                    ws.append([
                        sequence, segment.get("county", ""), segment.get("route", "G210"), segment["grade"], segment["manager"],
                        format_station(segment["start"]), format_station(segment["end"]),
                        basis, data["count"], *[value / 100 for value in data["pcts"]], data["pass"] / 100,
                    ])
                else:
                    ws.append([
                        sequence, "G210", segment["grade"], segment["manager"],
                        format_station(segment["start"]), format_station(segment["end"]),
                        basis, data["count"], *[value / 100 for value in data["pcts"]], data["pass"] / 100,
                    ])
            if has_county:
                for row in ws.iter_rows(min_row=2, min_col=10, max_col=15):
                    for cell in row:
                        cell.number_format = "0.00%"
                style_sheet(ws, [7, 12, 10, 14, 24, 14, 14, 42, 11, 13, 14, 14, 14, 13, 12])
            else:
                for row in ws.iter_rows(min_row=2, min_col=9, max_col=14):
                    for cell in row:
                        cell.number_format = "0.00%"
                style_sheet(ws, [7, 10, 14, 24, 14, 14, 42, 11, 13, 14, 14, 14, 13, 12])
        ws = wb.create_sheet("检测明细")
        if has_county:
            ws.append(["序号", "来源文件", "方向", "统计桩号", "桩号口径", "护栏类型", "梁板中心高度(mm)", "区县", "路线编号", "所属分段起点", "所属分段终点"])
        else:
            ws.append(["序号", "来源文件", "方向", "统计桩号", "桩号口径", "护栏类型", "梁板中心高度(mm)", "所属分段起点", "所属分段终点"])
        ordered = sorted(height_records, key=lambda item: (item["segment"], item["direction"], item["station"]))
        for sequence, record in enumerate(ordered, 1):
            segment = segments[record["segment"]]
            if has_county:
                ws.append([
                    sequence, record["file"], record["direction"], format_station(record["station"]),
                    record["basis"], record["kind"], record["height"],
                    segment.get("county", ""), segment.get("route", "G210"),
                    format_station(segment["start"]), format_station(segment["end"]),
                ])
            else:
                ws.append([
                    sequence, record["file"], record["direction"], format_station(record["station"]),
                    record["basis"], record["kind"], record["height"],
                    format_station(segment["start"]), format_station(segment["end"]),
                ])
        if has_county:
            style_sheet(ws, [8, 62, 10, 16, 14, 12, 20, 12, 10, 16, 16])
        else:
            style_sheet(ws, [8, 62, 10, 16, 14, 12, 20, 16, 16])
        # 高度分档对照：分档外边界＝设计值±50 / ±40，合格率＝落在该容差区间内的比例。
        # 二波/三波分档边界数值不同，按波形各出一张对照表（±50 与 ±40 各一组）。
        for tolerance in (50, 40):
            for kind in ("二波", "三波"):
                labels = height_bin_labels(kind, tolerance)
                ws = wb.create_sheet(f"{kind}对照（±{tolerance}mm）")
                if has_county:
                    header = ["序号", "区县", "路线编号", "公路等级", "管理单位", "起点桩号", "终点桩号", "统计口径", "检测点数"]
                    widths = [7, 12, 10, 14, 24, 14, 14, 42, 11, 13, 14, 14, 14, 13, 12]
                else:
                    header = ["序号", "路线编号", "公路等级", "管理单位", "起点桩号", "终点桩号", "统计口径", "检测点数"]
                    widths = [7, 10, 14, 24, 14, 14, 42, 11, 13, 14, 14, 14, 13, 12]
                ws.append([*header, *labels, f"合格率（±{tolerance}mm）"])
                sequence = 0
                for segment_index, item in enumerate(height_stats):
                    segment = item["segment"]
                    rows = [record for record in height_records
                            if record["segment"] == segment_index and record["kind"] == kind]
                    if not rows:
                        continue
                    data = _height_type_bins(rows, kind, tolerance, pass_bins=(1, 2, 3))
                    sequence += 1
                    bases = {record["basis"] for record in rows}
                    basis = next(iter(bases)) if len(bases) == 1 else "原始桩号+电子修正桩号（按来源文件分别采用）"
                    ws.append([
                        sequence,
                        *([segment.get("county", "")] if has_county else []),
                        segment.get("route", "G210"), segment["grade"], segment["manager"],
                        format_station(segment["start"]), format_station(segment["end"]),
                        basis, data["count"],
                        *[value / 100 for value in data["pcts"]], data["pass"] / 100,
                    ])
                # 合计行：整个区县各路线合并后按同一容差重新分档，合格率＝落在该容差区间内的点占比。
                for county in county_groups:
                    rows = [record for record in height_records
                            if record["kind"] == kind and _county_of(record["segment"]) == county]
                    if not rows:
                        continue
                    data = _height_type_bins(rows, kind, tolerance, pass_bins=(1, 2, 3))
                    indices = sorted({record["segment"] for record in rows})
                    ws.append([
                        "合计",
                        *([county] if has_county else []),
                        "全部路线", "", "",
                        format_station(min(segments[i]["start"] for i in indices)),
                        format_station(max(segments[i]["end"] for i in indices)),
                        "", data["count"],
                        *[value / 100 for value in data["pcts"]], data["pass"] / 100,
                    ])
                for row in ws.iter_rows(min_row=2, min_col=len(header) + 1, max_col=len(header) + 6):
                    for cell in row:
                        cell.number_format = "0.00%"
                style_sheet(ws, widths)

    if bolt_stats is not None:
        ws = wb.create_sheet("螺栓缺失统计")
        if has_county:
            ws.append(["序号", "区县", "路线编号", "起点桩号", "止点桩号", "检测里程（km）", "拼接螺栓（颗）", "连接螺栓（颗）", "缺失数量（颗）", "缺失率（%）"])
        else:
            ws.append(["序号", "路线编号", "起点桩号", "止点桩号", "检测里程（km）", "拼接螺栓（颗）", "连接螺栓（颗）", "缺失数量（颗）", "缺失率（%）"])
        sequence = 0
        for item in bolt_stats:
            sequence += 1
            segment = item["segment"]
            if has_county:
                ws.append([
                    sequence, segment.get("county", ""), segment.get("route", "G210"), format_station(segment["start"]), format_station(segment["end"]),
                    segment["mileage"], item["splice"], item["connection"], item["missing"], item["rate"] / 100 if item["rate"] is not None else None,
                ])
            else:
                ws.append([
                    sequence, "G210", format_station(segment["start"]), format_station(segment["end"]),
                    segment["mileage"], item["splice"], item["connection"], item["missing"], item["rate"] / 100 if item["rate"] is not None else None,
                ])
        # 合计行：整个区县各路线螺栓缺失汇总（按数量求和后重算缺失率）。
        for county in county_groups:
            indices = _county_indices(county)
            items = [bolt_stats[i] for i in indices if bolt_stats[i] is not None]
            if not items:
                continue
            total = merge_bolt_totals(items)
            ws.append([
                "合计",
                *([county] if has_county else []),
                "全部路线",
                format_station(min(segments[i]["start"] for i in indices)),
                format_station(max(segments[i]["end"] for i in indices)),
                sum(float(segments[i].get("mileage") or 0) for i in indices),
                total["splice"], total["connection"], total["missing"],
                total["rate"] / 100 if total["rate"] is not None else None,
            ])
        if has_county:
            for cell in ws["J"][1:]:
                cell.number_format = "0.00%"
            style_sheet(ws, [8, 12, 11, 16, 16, 16, 18, 18, 18, 14])
        else:
            for cell in ws["I"][1:]:
                cell.number_format = "0.00%"
            style_sheet(ws, [8, 11, 16, 16, 16, 18, 18, 18, 14])

        ws = wb.create_sheet("螺栓缺失明细")
        if has_county:
            ws.append([
                "序号", "来源文件", "方向", "统计桩号", "桩号口径",
                "拼接螺栓（颗）", "拼接螺栓缺失（颗）", "连接螺栓（颗）", "连接螺栓缺失（颗）",
                "区县", "路线编号", "所属分段起点", "所属分段终点",
            ])
        else:
            ws.append([
                "序号", "来源文件", "方向", "统计桩号", "桩号口径",
                "拼接螺栓（颗）", "拼接螺栓缺失（颗）", "连接螺栓（颗）", "连接螺栓缺失（颗）",
                "所属分段起点", "所属分段终点",
            ])
        ordered = sorted(bolt_records, key=lambda item: (item["segment"], item["direction"], item["station"]))
        for sequence, record in enumerate(ordered, 1):
            segment = segments[record["segment"]]
            if has_county:
                ws.append([
                    sequence, record["file"], record["direction"], format_station(record["station"]), record["basis"],
                    record["splice"], record["splice_missing"], record["connection"], record["connection_missing"],
                    segment.get("county", ""), segment.get("route", "G210"),
                    format_station(segment["start"]), format_station(segment["end"]),
                ])
            else:
                ws.append([
                    sequence, record["file"], record["direction"], format_station(record["station"]), record["basis"],
                    record["splice"], record["splice_missing"], record["connection"], record["connection_missing"],
                    format_station(segment["start"]), format_station(segment["end"]),
                ])
        if has_county:
            style_sheet(ws, [8, 62, 10, 16, 14, 18, 22, 18, 22, 12, 10, 16, 16])
        else:
            style_sheet(ws, [8, 62, 10, 16, 14, 18, 22, 18, 22, 16, 16])

    if tci_stats is not None:
        ws = wb.create_sheet("沿线设施统计")
        if has_county:
            ws.append(["序号", "区县", "路线编号", "起点桩号", "止点桩号", "检测里程（km）", "防护-轻（处）", "防护-重（处）", "标志缺损（处）", "标线缺损（m）", "TCI", "等级"])
        else:
            ws.append(["序号", "路线编号", "起点桩号", "止点桩号", "检测里程（km）", "防护-轻（处）", "防护-重（处）", "标志缺损（处）", "标线缺损（m）", "TCI", "等级"])
        for idx, item in enumerate(tci_stats, 1):
            seg = item["segment"]
            if has_county:
                ws.append([idx, seg.get("county",""), seg.get("route","G210"), format_station(seg["start"]), format_station(seg["end"]), f"{seg['mileage']:.3f}", item["light"], item["heavy"], item["sign"], item["marking"], item["tci"], item["grade"]])
            else:
                ws.append([idx, seg.get("route","G210"), format_station(seg["start"]), format_station(seg["end"]), f"{seg['mileage']:.3f}", item["light"], item["heavy"], item["sign"], item["marking"], item["tci"], item["grade"]])
        # 合计行：整个区县各路线沿线设施汇总；TCI 按本区县全部评定单元等权平均。
        for county in county_groups:
            items = [item for item in tci_stats
                     if (str(item["segment"].get("county") or "") if has_county else "") == county]
            units = [unit for item in items for unit in item.get("units", [])]
            if not items or not units:
                continue
            score = sum(unit["tci"] for unit in units) / len(units)
            ws.append([
                "合计",
                *([county] if has_county else []),
                "全部路线",
                format_station(min(item["segment"]["start"] for item in items)),
                format_station(max(item["segment"]["end"] for item in items)),
                f"{sum(float(item['segment'].get('mileage') or 0) for item in items):.3f}",
                sum(item["light"] for item in items), sum(item["heavy"] for item in items),
                sum(item["sign"] for item in items), sum(item["marking"] for item in items),
                score, tci_grade(score),
            ])
        if has_county:
            style_sheet(ws, [8, 12, 11, 16, 16, 16, 14, 14, 14, 14, 10, 8])
        else:
            style_sheet(ws, [8, 11, 16, 16, 16, 14, 14, 14, 14, 10, 8])

        units_sheet = wb.create_sheet("TCI公里评定明细")
        units_sheet.append(["序号", "路线编号", "区县", "方向", "起点桩号", "止点桩号", "TCI", "等级"])
        for group in tci_route_stats(segments, tci_stats):
            for unit in group["units"]:
                units_sheet.append([units_sheet.max_row, unit["route"], unit["county"], unit["direction"],
                                    format_station(unit["start"]), format_station(unit["end"]), unit["tci"], unit["grade"]])
        for cell in units_sheet["G"][1:]:
            cell.number_format = "0.00"
        style_sheet(units_sheet, [8, 12, 14, 12, 18, 18, 12, 10])
        ws = wb.create_sheet("沿线设施明细")
        if has_county:
            ws.append(["序号", "区县", "路线编号", "桩号", "轻", "重", "标志", "标线(m)", "所属分段起点", "所属分段终点"])
        else:
            ws.append(["序号", "路线编号", "桩号", "轻", "重", "标志", "标线(m)", "所属分段起点", "所属分段终点"])
        ordered = sorted(tci_records, key=lambda x: (x["segment"], x["station"]))
        for seq, rec in enumerate(ordered, 1):
            seg = segments[rec["segment"]]
            if has_county:
                ws.append([seq, seg.get("county",""), seg.get("route","G210"), format_station(rec["station"]), rec["light"], rec["heavy"], rec["sign"], rec["marking"], format_station(seg["start"]), format_station(seg["end"])])
            else:
                ws.append([seq, seg.get("route","G210"), format_station(rec["station"]), rec["light"], rec["heavy"], rec["sign"], rec["marking"], format_station(seg["start"]), format_station(seg["end"])])
        if has_county:
            style_sheet(ws, [8, 12, 11, 16, 8, 8, 8, 10, 16, 16])
        else:
            style_sheet(ws, [8, 11, 16, 8, 8, 8, 10, 16, 16])

    ws = wb.create_sheet("统计说明")
    notes = [
        ("项目", "说明"),
        ("统计分段", "按G210采集路段信息汇总表的起终点桩号分段。"),
        ("桩号口径", "G210上行K2264K2325文件使用原始桩号；其他文件使用电子修正桩号。"),
    ]
    if height_stats is not None:
        notes.extend([
            ("中心高度统计数值", "使用护栏高度表中的梁板中心高度(mm)列。"),
            ("中心高度备注过滤", f"仅保留异常标记为空或无备注的数据；排除{sum(excluded.values())}条：{dict(excluded)}。"),
            ("中心高度去重", f"完全一致的重叠记录去重，剔除{height_duplicates}条。"),
            ("区间明细标准值", "标准值列为580mm、620mm、677mm和717mm（设计值±20mm）。"),
            ("高度分档对照", "不再在工作簿内绘图；新增二波/三波各两张对照表，分档外边界取设计值±50mm（二波550/580/620/650、三波647/677/717/747）和±40mm（二波560/580/620/640、三波657/677/717/737），合格率为落在该容差区间内的点占比。"),
        ])
    if tci_stats is not None:
        notes.extend([
            ("TCI 统计", "原区段内按整公里桩号、方向划分，首尾不足公里单独评定；路线和区县按所有评定单元等权平均。TCI=Σwᵢ(100-GDᵢ)/0.7，w=[0.25,0.25,0.20]，防护轻10/重30、标志20/处、标线每10m1分不足10m计10m，GD封顶100，等级 优≥90 良≥80 中≥70 次≥60 差<60。"),
            ("TCI 数据来源", f"TCI 病害明细共{len(tci_records)}条，分段{len(tci_stats)}。"),
        ])
    if bolt_stats is not None:
        notes.extend([
            ("螺栓统计数值", "使用拼接螺栓数量、拼接螺栓缺失数量、连接螺栓数量和连接螺栓缺失数量列。"),
            ("螺栓缺失总数", "缺失数量为拼接螺栓缺失数量与连接螺栓缺失数量之和。"),
            ("螺栓缺失率", "缺失数量÷（拼接螺栓数量+连接螺栓数量+缺失数量）×100%。"),
            ("螺栓去重", f"完全一致的重叠记录去重，剔除{bolt_duplicates}条。"),
        ])
    if height_stats is not None or bolt_stats is not None or tci_stats is not None:
        notes.append(("合计行", "分档对照表、螺栓缺失统计和沿线设施统计的末行为“合计”（写在序号列，路线编号列为“全部路线”），表示该区县全部路线的汇总：分档对照按全部检测点重新计算分档占比与合格率，螺栓按数量求和后重算缺失率，沿线设施汇总的TCI取本区县全部评定单元等权平均。"))
    for row in notes:
        ws.append(row)
    style_sheet(ws, [18, 110])
    config.output_dir.mkdir(parents=True, exist_ok=True)
    wb.save(config.out_xlsx)
    log(f"统计工作簿已生成：{config.out_xlsx}")
    if height_stats is not None:
        add_interval_sheets(config.out_xlsx, config.summary_xlsx, log, segments)


def run_properties(bold=False, size=21):
    properties = ET.Element(wt("rPr"))
    fonts = ET.SubElement(properties, wt("rFonts"))
    fonts.set(wt("ascii"), "Times New Roman")
    fonts.set(wt("hAnsi"), "Times New Roman")
    fonts.set(wt("eastAsia"), "宋体")
    fonts.set(wt("cs"), "Times New Roman")
    if bold:
        ET.SubElement(properties, wt("b")); ET.SubElement(properties, wt("bCs"))
    if size is not None:
        ET.SubElement(properties, wt("sz")).set(wt("val"), str(size))
        ET.SubElement(properties, wt("szCs")).set(wt("val"), str(size))
    return properties


def paragraph_properties(style=None, center=False, num_id=None, level=None, keep_next=None):
    properties = ET.Element(wt("pPr"))
    if style:
        ET.SubElement(properties, wt("pStyle")).set(wt("val"), str(style))
    if keep_next is not None:
        ET.SubElement(properties, wt("keepNext")).set(wt("val"), "1" if keep_next else "0")
    if num_id is not None:
        number = ET.SubElement(properties, wt("numPr"))
        ET.SubElement(number, wt("ilvl")).set(wt("val"), str(level))
        ET.SubElement(number, wt("numId")).set(wt("val"), str(num_id))
    indent = ET.SubElement(properties, wt("ind"))
    indent.set(wt("left"), "0"); indent.set(wt("right"), "0"); indent.set(wt("firstLine"), "0")
    ET.SubElement(properties, wt("jc")).set(wt("val"), "center" if center else "left")
    return properties


def append_run(p, text, bold=False, size=21):
    run = ET.SubElement(p, wt("r")); run.append(run_properties(bold, size))
    node = ET.SubElement(run, wt("t")); node.text = str(text)
    if str(text).startswith(" ") or str(text).endswith(" "):
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return run


def body_paragraph(text="", center=False, bold=False, style=None, keep_next=None):
    p = ET.Element(wt("p")); p.append(paragraph_properties(style, center, keep_next=keep_next))
    append_run(p, text, bold)
    return p


def heading_paragraph(text, level, num_id):
    style = {1: "2", 2: "3", 3: "4"}[level]
    p = ET.Element(wt("p")); p.append(paragraph_properties(style, False, num_id, level - 1))
    append_run(p, text, True, None)
    return p


def append_field(p, instruction, result, bold=False):
    begin = ET.SubElement(p, wt("r")); begin.append(run_properties(bold)); ET.SubElement(begin, wt("fldChar")).set(wt("fldCharType"), "begin")
    code = ET.SubElement(p, wt("r")); code.append(run_properties(bold)); text = ET.SubElement(code, wt("instrText")); text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve"); text.text = f" {instruction} "
    separate = ET.SubElement(p, wt("r")); separate.append(run_properties(bold)); ET.SubElement(separate, wt("fldChar")).set(wt("fldCharType"), "separate")
    append_run(p, result, bold)
    end = ET.SubElement(p, wt("r")); end.append(run_properties(bold)); ET.SubElement(end, wt("fldChar")).set(wt("fldCharType"), "end")


def caption_paragraph(
    title, label, section_number, sequence,
    bookmark_id=None, bookmark_name=None,
):
    """创建包含标题2章节号的图/表题注，编号格式为“章节号-序号”。"""
    p = ET.Element(wt("p")); p.append(paragraph_properties("9", True, keep_next=False))
    if bookmark_id is not None and bookmark_name:
        start = ET.SubElement(p, wt("bookmarkStart"))
        start.set(wt("id"), str(bookmark_id)); start.set(wt("name"), bookmark_name)
    append_run(p, f"{label} ", True)
    # Word“题注编号-包含章节号”的字段结构：章节起始样式为标题2，
    # STYLEREF取得当前标题2编号，SEQ按标题2重新开始计数。
    append_field(p, 'STYLEREF 2 \\s', str(section_number), True)
    append_run(p, "-", True)
    append_field(p, f'SEQ {label} \\* ARABIC \\s 2', str(sequence), True)
    if bookmark_id is not None and bookmark_name:
        end = ET.SubElement(p, wt("bookmarkEnd")); end.set(wt("id"), str(bookmark_id))
    append_run(p, "  " + title, True)
    return p


def append_reference(p, bookmark_name, visible_text):
    append_field(p, f"REF {bookmark_name} \\h", visible_text, False)


def word_table(headers, rows, header_shading=False, keep_together=False):
    table = ET.Element(wt("tbl")); properties = ET.SubElement(table, wt("tblPr"))
    borders = ET.SubElement(properties, wt("tblBorders"))
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = ET.SubElement(borders, wt(side)); node.set(wt("val"), "single"); node.set(wt("sz"), "4")
    for row_index, row in enumerate([headers] + rows):
        tr = ET.SubElement(table, wt("tr"))
        if keep_together:
            tr_properties = ET.SubElement(tr, wt("trPr")); ET.SubElement(tr_properties, wt("cantSplit"))
        for value in row:
            tc = ET.SubElement(tr, wt("tc")); tc_properties = ET.SubElement(tc, wt("tcPr"))
            if row_index == 0 and header_shading:
                shading = ET.SubElement(tc_properties, wt("shd"))
                shading.set(wt("val"), "pct15")
                shading.set(wt("color"), "auto")
                shading.set(wt("fill"), "auto")
            p = ET.SubElement(tc, wt("p"))
            # 表格内容允许自然分页，不设置“与下段同页”(w:keepNext)。
            p.append(paragraph_properties(None, True, keep_next=False))
            append_run(p, value, row_index == 0)
    return table


def height_example_table(records):
    """创建2列表格：奇数行留空放图，偶数行合并后显示点位信息。"""
    table = ET.Element(wt("tbl")); properties = ET.SubElement(table, wt("tblPr"))
    width = ET.SubElement(properties, wt("tblW")); width.set(wt("w"), "9000"); width.set(wt("type"), "dxa")
    borders = ET.SubElement(properties, wt("tblBorders"))
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = ET.SubElement(borders, wt(side)); node.set(wt("val"), "single"); node.set(wt("sz"), "4"); node.set(wt("color"), "808080")
    grid = ET.SubElement(table, wt("tblGrid"))
    for _ in range(2):
        ET.SubElement(grid, wt("gridCol")).set(wt("w"), "4500")

    for record in records:
        image_row = ET.SubElement(table, wt("tr")); row_properties = ET.SubElement(image_row, wt("trPr"))
        height = ET.SubElement(row_properties, wt("trHeight")); height.set(wt("val"), "2500"); height.set(wt("hRule"), "atLeast")
        for _ in range(2):
            cell = ET.SubElement(image_row, wt("tc")); cell_properties = ET.SubElement(cell, wt("tcPr"))
            cell_width = ET.SubElement(cell_properties, wt("tcW")); cell_width.set(wt("w"), "4500"); cell_width.set(wt("type"), "dxa")
            ET.SubElement(cell_properties, wt("vAlign")).set(wt("val"), "center")
            paragraph = ET.SubElement(cell, wt("p")); paragraph.append(paragraph_properties(None, True))

        info_row = ET.SubElement(table, wt("tr")); cell = ET.SubElement(info_row, wt("tc")); cell_properties = ET.SubElement(cell, wt("tcPr"))
        cell_width = ET.SubElement(cell_properties, wt("tcW")); cell_width.set(wt("w"), "9000"); cell_width.set(wt("type"), "dxa")
        ET.SubElement(cell_properties, wt("gridSpan")).set(wt("val"), "2")
        ET.SubElement(cell_properties, wt("vAlign")).set(wt("val"), "center")
        paragraph = ET.SubElement(cell, wt("p")); paragraph.append(paragraph_properties(None, True))
        electronic = record.get("electronic_station")
        raw = record.get("raw_station")
        electronic_text = format_station(electronic) if electronic is not None else "—"
        raw_text = format_station(raw) if raw is not None else "—"
        append_run(paragraph, f"{electronic_text}（{record['height']:.2f}mm） {raw_text}")
    return table


def replace_text(root, old, new):
    for node in root.findall(f".//{wt('t')}"):
        if node.text:
            node.text = node.text.replace(old, new)


def add_multilevel_numbering(numbering_xml):
    root = ET.fromstring(numbering_xml)
    abstracts = [int(node.get(wt("abstractNumId"))) for node in root.findall(wt("abstractNum"))]
    nums = [int(node.get(wt("numId"))) for node in root.findall(wt("num"))]
    abstract_id = max(abstracts, default=0) + 1
    abstract = ET.Element(wt("abstractNum")); abstract.set(wt("abstractNumId"), str(abstract_id))
    ET.SubElement(abstract, wt("multiLevelType")).set(wt("val"), "multilevel")
    for level, (style, text) in enumerate((("2", "%1"), ("3", "%1.%2"), ("4", "%1.%2.%3"))):
        item = ET.SubElement(abstract, wt("lvl")); item.set(wt("ilvl"), str(level))
        ET.SubElement(item, wt("start")).set(wt("val"), "1")
        ET.SubElement(item, wt("numFmt")).set(wt("val"), "decimal")
        # 不写lvlRestart：Word默认在出现上一级标题时重启下一层编号。
        # 显式写入与当前层级冲突的值会导致Word修复并删除列表定义。
        ET.SubElement(item, wt("pStyle")).set(wt("val"), style)
        # 编号后使用一个空格，不使用Word默认制表符；标题保持无列表缩进。
        ET.SubElement(item, wt("suff")).set(wt("val"), "space")
        ET.SubElement(item, wt("lvlText")).set(wt("val"), text)
        ET.SubElement(item, wt("lvlJc")).set(wt("val"), "left")
        ppr = ET.SubElement(item, wt("pPr")); ind = ET.SubElement(ppr, wt("ind")); ind.set(wt("left"), "0"); ind.set(wt("hanging"), "0")
    # numbering.xml要求所有abstractNum位于num之前。
    first_num_index = next((i for i, child in enumerate(list(root)) if child.tag == wt("num")), len(root))
    root.insert(first_num_index, abstract)
    chapter_nums = {}
    next_id = max(nums, default=0) + 1
    for chapter in (1, 3, 4, 5):
        num = ET.SubElement(root, wt("num")); num.set(wt("numId"), str(next_id))
        ET.SubElement(num, wt("abstractNumId")).set(wt("val"), str(abstract_id))
        override = ET.SubElement(num, wt("lvlOverride")); override.set(wt("ilvl"), "0")
        ET.SubElement(override, wt("startOverride")).set(wt("val"), str(chapter))
        # 每章使用独立编号实例，并显式将小节、子小节从1开始。
        for level in (1, 2):
            override = ET.SubElement(num, wt("lvlOverride")); override.set(wt("ilvl"), str(level))
            ET.SubElement(override, wt("startOverride")).set(wt("val"), "1")
        chapter_nums[chapter] = next_id; next_id += 1
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), chapter_nums


def percentage_phrases(kind, percentages):
    limits = (550, 580, 620, 650) if kind == "二波" else (647, 677, 717, 747)
    labels = [
        f"小于{limits[0]}mm的约占{{:.2f}}%",
        f"介于{limits[0]}～{limits[1]}mm的约占{{:.2f}}%",
        f"介于{limits[1]}～{limits[2]}mm的约占{{:.2f}}%",
        f"介于{limits[2]}～{limits[3]}mm的约占{{:.2f}}%",
        f"大于{limits[3]}mm的约占{{:.2f}}%",
    ]
    return "，".join(labels[index].format(value) for index, value in enumerate(percentages) if value > 0)


def order_example_records(records):
    """示例图片顺序：上行桩号递增，下行桩号递减；混合时上行组在前。"""
    def key(record):
        direction = str(record.get("direction") or "")
        station = record.get("station")
        station = float(station) if station is not None else 0.0
        if "上" in direction:
            return 0, station
        if "下" in direction:
            return 1, -station
        return 2, station
    return sorted(records, key=key)


def matching_height_photos(row, photo_index):
    """Return only identity-matching photos at the first matching station basis."""
    if not photo_index:
        return []
    direction = normalize_direction(row.get("direction"))
    for station_key in ("raw_station", "electronic_station", "station"):
        station = row.get(station_key)
        if station is None:
            continue
        candidates = photo_index.get((direction, round(float(station), 1)), [])
        matches = []
        for candidate in candidates:
            if isinstance(candidate, dict):
                if candidate.get("role") not in (None, "height", "disease"):
                    continue
                row_route = _route(row.get("route") or "")
                candidate_route = _route(candidate.get("route") or "")
                if row_route and candidate_route and row_route != candidate_route:
                    continue
                row_county = str(row.get("county") or "").strip()
                candidate_county = str(candidate.get("county") or "").strip()
                if row_county and candidate_county and candidate_county not in county_short_names(row_county):
                    continue
            if candidate not in matches:
                matches.append(candidate)
        if matches:
            return matches
    return []


def row_has_height_photo(row, photo_index):
    """Use the same identity filter for example selection and photo rendering."""
    return bool(matching_height_photos(row, photo_index))


def select_height_example_points(rows, segment_index=0, kind="", photo_index=None):
    """选自动计算示例点：先筛出能挂上病害照片的检测点，再按数量/高度分位选点。

    提供照片索引时只用可挂图点（无照片点不参与选点；本段本方向本波形一个可挂图点
    都没有就返回空，由调用方不输出无图示例表）；未提供照片索引时退回全部检测点。
    """
    if photo_index:
        rows = [row for row in rows if row_has_height_photo(row, photo_index)]
        if not rows:
            return []
    ordered = sorted(rows, key=lambda item: (item["height"], item["station"]))
    count = len(ordered)
    if not count:
        return []
    if count > 1000:
        selected = []
        for index in range(4):
            group = ordered[index * count // 4:(index + 1) * count // 4]
            selected.append(group[len(group) // 2])
        return order_example_records(selected)
    if count > 100:
        quarter_count = max(1, math.ceil(count * 0.25))
        randomizer = random.Random(f"{segment_index}|{kind}|{count}")
        return order_example_records([
            randomizer.choice(ordered[:quarter_count]),
            randomizer.choice(ordered[-quarter_count:]),
        ])
    # 偶数个点时采用靠前的中位点，确保选中的是实际存在的数据点。
    return order_example_records([ordered[len(ordered) // 2]])


def select_bolt_example_points(rows, disease_image_index=None):
    """按缺失阈值和病害图片可匹配性选择拼接、连接螺栓示例。"""
    def attach_image(row):
        descriptor = match_disease_image(row, disease_image_index) if disease_image_index is not None else None
        return row, descriptor

    candidates = [attach_image(row) for row in rows]
    if disease_image_index is not None:
        candidates = [(row, image) for row, image in candidates if image is not None]
    splice_only_rows = [
        (row, image) for row, image in candidates
        if 0 < row["splice_missing"] <= 12 and row["connection_missing"] <= 0
    ]
    connection_rows = [
        (row, image) for row, image in candidates
        if 0 < row["connection_missing"] <= 2 and row["splice_missing"] <= 12
    ]
    splice_only_rows.sort(key=lambda item: (-item[0]["splice_missing"], item[0]["station"]))
    connection_rows.sort(key=lambda item: (
        -item[0]["connection_missing"],
        -(item[0]["splice_missing"] + item[0]["connection_missing"]),
        item[0]["station"],
    ))

    selected_rows = []
    if splice_only_rows:
        selected_rows.append(splice_only_rows[0])
    if connection_rows:
        selected_rows.append(connection_rows[0])
    elif selected_rows:
        # 没有任何连接螺栓缺失点时，只保留一个仅拼接螺栓缺失点。
        selected_rows = selected_rows[:1]

    selected = [
        {
            "record": row,
            "missing": int(round(row["splice_missing"] + row["connection_missing"])),
            "image": image,
        }
        for row, image in selected_rows
    ]
    ordered_records = order_example_records([example["record"] for example in selected])
    by_identity = {id(example["record"]): example for example in selected}
    return [by_identity[id(record)] for record in ordered_records]


def bolt_example_text(example):
    """螺栓示例标题：电子（修正）桩号 + 缺失颗数 + 原始桩号，两个桩号均保留一位小数。"""
    record = example["record"]
    raw = record.get("raw_station")
    electronic = record.get("electronic_station")
    raw_text = format_station_one_decimal(raw) if raw is not None else "—"
    electronic_text = format_station_one_decimal(electronic) if electronic is not None else "—"
    return f"{electronic_text} 螺栓缺失{example['missing']}颗 {raw_text}"


def report_images(temp_dir, segments, stats, records):
    """按区段×方向分别绘制高度折线图与分档分布图（无图内标题）。"""
    images = {}
    plt.rcParams["font.sans-serif"] = ["SimSun", "Microsoft YaHei", "Arial Unicode MS"]
    plt.rcParams["axes.unicode_minus"] = False
    grouped = {}
    for record in records or []:
        grouped.setdefault(
            (record["segment"], normalize_direction(record.get("direction")), record["kind"]), []
        ).append(record)
    for segment_index, item in enumerate(stats):
        for direction in DIRECTIONS:
            for kind in ("二波", "三波"):
                rows = sorted(
                    grouped.get((segment_index, direction, kind), []),
                    key=lambda record: record["station"],
                )
                if not rows:
                    continue
                key = (segment_index, direction, kind)
                line_path = Path(temp_dir) / f"line_{segment_index}_{direction}_{kind}.png"
                pie_path = Path(temp_dir) / f"pie_{segment_index}_{direction}_{kind}.png"
                x = list(range(len(rows))); heights = [row["height"] for row in rows]
                figure, axis = plt.subplots(figsize=(13 / 2.54, 8 / 2.54), dpi=180)
                axis.plot(x, heights, color=GD03_COLORS[0], linewidth=1, label="梁板中心高度(mm)")
                standards = (580, 620) if kind == "二波" else (677, 717)
                for standard, color in zip(standards, GD03_COLORS[1:3]):
                    axis.plot(
                        x, [standard] * len(x), color=color, linewidth=2,
                        label=f"标准值（{standard}mm）",
                    )
                axis.set_ylim(300, 850); axis.grid(True, alpha=0.25)
                ticks = sorted(set(int(i * (len(x) - 1) / min(9, max(1, len(x) - 1))) for i in range(min(10, len(x)))))
                axis.set_xticks(ticks); axis.set_xticklabels([format_station(rows[i]["station"]) for i in ticks], rotation=30, ha="right", fontsize=7)
                axis.legend(loc="upper center", bbox_to_anchor=(0.5, -0.22), ncol=3, frameon=False)
                figure.subplots_adjust(left=0.10, right=0.98, top=0.96, bottom=0.30)
                figure.savefig(line_path, transparent=False); plt.close(figure)

                data = (item.get("by_direction") or {}).get(direction, {}).get(kind) or _height_type_bins(rows, kind)
                values = data["pcts"]
                labels = (["h＜550", "550≤h＜580", "580≤h≤620", "620＜h≤650", "h＞650"] if kind == "二波" else
                          ["h＜647", "647≤h＜677", "677≤h≤717", "717＜h≤747", "h＞747"])
                nonzero = [(label, value, f"#{PIE_COLORS[i]}") for i, (label, value) in enumerate(zip(labels, values)) if value > 0]
                figure, axis = plt.subplots(figsize=(14 / 2.54, 8.5 / 2.54), dpi=180)
                wedges, _, _ = axis.pie(
                    [value for _, value, _ in nonzero], colors=[color for _, _, color in nonzero],
                    autopct=lambda pct: f"{pct:.2f}%" if pct > 0 else "", pctdistance=1.15,
                    textprops={"fontsize": 8},
                )
                axis.legend(wedges, [label for label, _, _ in nonzero], loc="center left", bbox_to_anchor=(1.0, 0.5), frameon=False)
                figure.subplots_adjust(left=0.02, right=0.76, top=0.88, bottom=0.05)
                figure.savefig(pie_path, transparent=False); plt.close(figure)
                images[key] = {"line": line_path, "pie": pie_path}
    return images


def report_tci_images(temp_dir, segments, tci_stats):
    """按道路编号的原检测区段、方向分别绘制TCI，样式与高度折线图统一。"""
    images = {}
    plt.rcParams["font.sans-serif"] = ["SimSun", "Microsoft YaHei", "Arial Unicode MS"]
    plt.rcParams["axes.unicode_minus"] = False
    for seg_idx, stat in enumerate(tci_stats or []):
        for direction in ("上行", "下行", ""):
            units = [u for u in stat.get("units", []) if normalize_direction(u.get("direction")) == direction]
            if not units:
                continue
            figure, axis = plt.subplots(figsize=(13 / 2.54, 8 / 2.54), dpi=180)
            x = list(range(len(units)))
            axis.plot(x, [u["tci"] for u in units], color=GD_SERIES_HEX[0], linewidth=1,
                      marker="o", markersize=2, label="TCI")
            axis.set_ylim(0, 105)
            axis.grid(True, alpha=0.25)
            axis.set_ylabel("TCI")
            # 点位过多时抽样横坐标刻度（与高度折线图一致），避免桩号标签重叠。
            ticks = sorted(set(int(i * (len(x) - 1) / min(9, max(1, len(x) - 1))) for i in range(min(10, len(x)))))
            axis.set_xticks(ticks)
            axis.set_xticklabels([format_station(units[i]["start"]) for i in ticks], rotation=30, ha="right", fontsize=7)
            axis.legend(loc="upper center", bbox_to_anchor=(0.5, -0.22), ncol=1, frameon=False)
            figure.subplots_adjust(left=0.10, right=0.98, top=0.96, bottom=0.30)
            path = Path(temp_dir) / f"tci_{len(images)}.png"
            figure.savefig(path, transparent=False)
            plt.close(figure)
            images[(seg_idx, direction)] = path
    return images

def picture_paragraph(rel_id, drawing_id, width_cm=13, height_cm=8):
    cx, cy = int(width_cm * 360000), int(height_cm * 360000)
    p = ET.Element(wt("p")); p.append(paragraph_properties(None, True))
    run = ET.SubElement(p, wt("r")); run.append(run_properties())
    drawing = ET.SubElement(run, wt("drawing")); inline = ET.SubElement(drawing, q(WP, "inline"))
    extent = ET.SubElement(inline, q(WP, "extent")); extent.set("cx", str(cx)); extent.set("cy", str(cy))
    effect = ET.SubElement(inline, q(WP, "effectExtent"));
    for side in ("l", "t", "r", "b"): effect.set(side, "0")
    docpr = ET.SubElement(inline, q(WP, "docPr")); docpr.set("id", str(drawing_id)); docpr.set("name", f"G210Chart{drawing_id}")
    frame = ET.SubElement(inline, q(WP, "cNvGraphicFramePr")); ET.SubElement(frame, q(A, "graphicFrameLocks")).set("noChangeAspect", "1")
    graphic = ET.SubElement(inline, q(A, "graphic")); data = ET.SubElement(graphic, q(A, "graphicData")); data.set("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture")
    pic = ET.SubElement(data, q(PIC, "pic")); nv = ET.SubElement(pic, q(PIC, "nvPicPr"));
    cpr = ET.SubElement(nv, q(PIC, "cNvPr")); cpr.set("id", "0"); cpr.set("name", f"G210Image{drawing_id}")
    ET.SubElement(nv, q(PIC, "cNvPicPr")); fill = ET.SubElement(pic, q(PIC, "blipFill"))
    blip = ET.SubElement(fill, q(A, "blip")); blip.set(q(R, "embed"), rel_id)
    stretch = ET.SubElement(fill, q(A, "stretch")); ET.SubElement(stretch, q(A, "fillRect"))
    shape = ET.SubElement(pic, q(PIC, "spPr")); transform = ET.SubElement(shape, q(A, "xfrm"))
    offset = ET.SubElement(transform, q(A, "off")); offset.set("x", "0"); offset.set("y", "0")
    size = ET.SubElement(transform, q(A, "ext")); size.set("cx", str(cx)); size.set("cy", str(cy))
    geometry = ET.SubElement(shape, q(A, "prstGeom")); geometry.set("prst", "rect"); ET.SubElement(geometry, q(A, "avLst"))
    return p


def make_docx(
    config, segments,
    height_stats=None, height_records=None,
    bolt_stats=None, bolt_records=None,
    tci_stats=None, tci_records=None,
    disease_image_index=None,
    log=lambda _: None,
    require_template=True,
):
    if config.template_docx.suffix.lower() != ".md":
        raise ValueError(f"仅支持 Markdown 模板：{config.template_docx}，请使用 .md 模板。")
    if not config.template_docx.is_file():
        raise FileNotFoundError(f"Markdown 模板不存在：{config.template_docx}")
    from backend import minimal_docx
    log("使用 Markdown 报告模板。")
    return minimal_docx.run(
        config, segments,
        height_stats=height_stats, height_records=height_records,
        bolt_stats=bolt_stats, bolt_records=bolt_records,
        tci_stats=tci_stats, tci_records=tci_records,
        disease_image_index=disease_image_index, log=log,
        skeleton_md=config.template_docx,
    )

def _iter_input_filenames(config):
    """明细/病害/TCI 输入文件名（去临时文件），供区县路由识别。"""
    names = []
    for directory in (config.detail_dir, config.disease_dir):
        try:
            if directory is not None and Path(directory).is_dir():
                names.extend(path.name for path in Path(directory).glob("*.xlsx") if not path.name.startswith("~$"))
        except OSError:
            continue
    tci = getattr(config, "tci_path", None) or getattr(config, "tci_xlsx", None)
    try:
        if tci is not None:
            tci = Path(tci)
            if tci.is_file():
                names.append(tci.name)
            elif tci.is_dir():
                names.extend(path.name for path in tci.glob("*.xlsx") if not path.name.startswith("~$"))
    except OSError:
        pass
    return names


def _apply_county_routing(config, segments, county_override, log):
    """先文件名识区县、再按区县划分情况表过滤对应路线分段；单区县时顶层文件直接用区县命名。"""
    matched = resolve_report_counties(segments, _iter_input_filenames(config), override=county_override)
    if not matched:
        return segments
    known = []
    for segment in segments:
        county = str(segment.get("county") or "").strip()
        if county and county not in known:
            known.append(county)
    if set(matched) < set(known):
        segments = [s for s in segments if str(s.get("county") or "").strip() in set(matched)]
        log(f"按文件名识别到区县：{'、'.join(matched)}，仅生成对应区县分段报告。")
    if not segments:
        raise ValueError("区县划分情况表中无对应路线分段，已中断。")
    final = {str(s.get("county") or "").strip() for s in segments} - {""}
    if len(final) == 1 and not getattr(config, "county", None):
        config.county = next(iter(final))
    return segments


def generate_statistics_and_report(
    config, log=lambda _: None,
    process_height=True, process_bolts=False, process_alongline=False, process_tci=False,
    require_template=True, county_override=None,
):
    segments = read_segments(config.summary_xlsx)
    segments = _apply_county_routing(config, segments, county_override, log)
    height_records = []; height_stats = None; height_duplicates = 0; excluded = Counter()
    bolt_records = []; bolt_stats = None; bolt_duplicates = 0
    disease_image_index = None
    tci_records = []; tci_stats = None
    if process_alongline or process_tci:
        # TCI 沿线设施：读取 tci_path 指定的病害明细（支持单文件或目录）
        tci_src = getattr(config, "tci_path", None) or getattr(config, "tci_xlsx", None)
        if tci_src is None or not Path(tci_src).exists():
            log(f"TCI 未提供有效路径（{tci_src}），跳过沿线设施统计。")
        else:
            tci_records = collect_tci_records(segments, tci_src, log)
            tci_stats = make_tci_stats(segments, tci_records)
            log(f"TCI 有效记录{len(tci_records)}条；分段{len(tci_stats)}段。")
    if process_height:
        height_records, height_duplicates, excluded = collect_records(segments, config.detail_dir, log)
        height_stats = make_stats(segments, height_records)
    if process_bolts:
        bolt_records, bolt_duplicates = collect_bolt_records(segments, config.detail_dir, log)
        bolt_stats = make_bolt_stats(segments, bolt_records)
        if config.disease_dir is None or not Path(config.disease_dir).is_dir():
            raise FileNotFoundError("处理螺栓缺失时，请选择有效的病害清单文件夹。")
        disease_image_index = collect_disease_image_index(config.disease_dir, log)
    for county, routes in sorted(county_route_summary(segments).items()):
        if not county:
            continue
        detail = "、".join(f"{route}{count}段" for route, count in sorted(routes.items()))
        log(f"区县{county}：识别到道路编号{detail}。")

    def _write_outputs(cfg, segs, h_stats, h_recs, b_stats, b_recs, t_stats, t_recs, step=None):
        # step=(第几个区县, 区县总数, 区县名)：给出可量化的阶段进度，避免进度条长时间不动。
        if step:
            log(format_progress("生成区县统计", step[0], step[1], step[2]))
        make_excel(
            cfg, segs,
            height_stats=h_stats, height_records=h_recs,
            height_duplicates=height_duplicates, excluded=excluded,
            bolt_stats=b_stats, bolt_records=b_recs, bolt_duplicates=bolt_duplicates,
            tci_stats=t_stats, tci_records=t_recs,
            log=log,
        )
        if step:
            log(format_progress("生成区县报告", step[0], step[1], step[2]))
        make_docx(
            cfg, segs,
            height_stats=h_stats, height_records=h_recs,
            bolt_stats=b_stats, bolt_records=b_recs,
            tci_stats=t_stats, tci_records=t_recs,
            disease_image_index=disease_image_index, log=log,
            require_template=require_template,
        )

    # R2：不再生成整体报告，只逐区县生成。单区县时顶层直接用区县命名即为该区县报告；
    # 多区县时只生成各区县子目录报告，不生成顶层整体文件。无区县维度（旧汇总表）保留 legacy 行为。
    counties = sorted({str(s.get("county") or "").strip() for s in segments if str(s.get("county") or "").strip()})
    if counties and segments:
        if len(counties) == 1:
            if not getattr(config, "county", None):
                config.county = counties[0]
            _write_outputs(config, segments, height_stats, height_records, bolt_stats, bolt_records, tci_stats, tci_records,
                           step=(1, 1, counties[0]))
        else:
            from collections import defaultdict
            by_county = defaultdict(list)
            for idx, seg in enumerate(segments):
                by_county[seg.get("county", "")].append(idx)
            county_items = [(name, ids) for name, ids in by_county.items() if name]
            for position, (county, idxs) in enumerate(county_items, 1):
                sub_segments = [segments[i] for i in idxs]
                # 重映射记录的 segment 索引到子集 0..n-1
                def _remap(records):
                    out = []
                    mp = {orig: new for new, orig in enumerate(idxs)}
                    for r in records:
                        if r["segment"] in mp:
                            nr = dict(r)
                            nr["segment"] = mp[r["segment"]]
                            out.append(nr)
                    return out
                # 浅拷贝 stats 后再调整 segment 引用，避免污染其他区县
                sub_height_stats = [dict(height_stats[i]) for i in idxs] if height_stats else None
                if sub_height_stats:
                    for ns, st in zip(sub_segments, sub_height_stats):
                        st["segment"] = ns
                sub_bolt_stats = [dict(bolt_stats[i]) for i in idxs] if bolt_stats else None
                if sub_bolt_stats:
                    for ns, st in zip(sub_segments, sub_bolt_stats):
                        st["segment"] = ns
                sub_tci_stats = [dict(tci_stats[i]) for i in idxs] if tci_stats else None
                if sub_tci_stats:
                    for ns, st in zip(sub_segments, sub_tci_stats):
                        st["segment"] = ns
                sub_height_records = _remap(height_records)
                sub_bolt_records = _remap(bolt_records)
                sub_tci_records = _remap(tci_records)
                sub_out = config.output_dir / county
                sub_out.mkdir(parents=True, exist_ok=True)
                sub_cfg = Config(config.project_dir, config.summary_xlsx, config.detail_dir, config.template_docx, sub_out, config.disease_dir, getattr(config, "tci_path", None), county=county)
                # 复用 make_excel/make_docx 生成子报告
                try:
                    _write_outputs(sub_cfg, sub_segments, sub_height_stats, sub_height_records, sub_bolt_stats, sub_bolt_records, sub_tci_stats, sub_tci_records,
                                   step=(position, len(county_items), county))
                    log(f"区县分报告已生成：{county} -> {sub_out}")
                except Exception as e:
                    log(f"区县 {county} 分报告生成失败：{e}")
    else:
        _write_outputs(config, segments, height_stats, height_records, bolt_stats, bolt_records, tci_stats, tci_records)

    if process_height:
        log(f"中心高度有效记录{len(height_records)}条；备注排除{sum(excluded.values())}条；重复排除{height_duplicates}条。")
    if process_bolts:
        log(f"螺栓有效记录{len(bolt_records)}条；重复排除{bolt_duplicates}条。")
    return {"height": height_stats, "bolts": bolt_stats, "tci": tci_stats}


def _workbook_head_text(path, max_rows=3, max_sheets=3):
    """读取工作簿前几张表前几行的文本，供 discover 回退按表头确认文件类型。"""
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return ""
    try:
        texts = []
        for sheet_name in wb.sheetnames[:max_sheets]:
            try:
                ws = wb[sheet_name]
            except Exception:
                continue
            try:
                for row in ws.iter_rows(min_row=1, max_row=max_rows, values_only=True):
                    for value in row or ():
                        if value is not None and str(value).strip():
                            texts.append(str(value).strip())
            except Exception:
                continue
        return "\n".join(texts)
    except Exception:
        return ""
    finally:
        try:
            wb.close()
        except Exception:
            pass


def _fallback_summary_file(folder):
    """汇总表回退：*汇总*/*整理*.xlsx，按表头含路线/桩号列确认；旧死文件名优先（见 discover_paths）。"""
    candidates = []
    for path in sorted(folder.rglob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        if "汇总" not in path.name and "整理" not in path.name:
            continue
        head = _workbook_head_text(path)
        if "路线" in head and "桩号" in head:
            candidates.append(path)
    candidates.sort(key=lambda p: (0 if "汇总" in p.name else 1, str(p)))
    return candidates[0] if candidates else None


def _fallback_detail_dir(folder):
    """明细回退：*明细*.xlsx（排除病害清单），按表头含护栏/螺栓列确认；返回含匹配最多的目录。"""
    grouped: dict = {}
    for path in sorted(folder.rglob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        if "明细" not in path.name or "病害" in path.name:
            continue
        head = _workbook_head_text(path)
        if "护栏类型" in head or "梁板中心高度" in head or "拼接螺栓数量" in head:
            grouped.setdefault(path.parent, []).append(path)
    if not grouped:
        return None
    return max(grouped.items(), key=lambda item: (len(item[1]), str(item[0])))[0]


def _fallback_disease_dir(folder):
    """病害清单回退：*病害*.xlsx；返回含匹配最多的目录。"""
    grouped: dict = {}
    for path in sorted(folder.rglob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        if "病害" not in path.name:
            continue
        grouped.setdefault(path.parent, []).append(path)
    if not grouped:
        return None
    return max(grouped.items(), key=lambda item: (len(item[1]), str(item[0])))[0]


def discover_paths(folder):
    folder = Path(folder)
    summary = next(iter(folder.rglob("G210采集路段信息汇总.xlsx")), None)
    detail_candidates = []
    disease_candidates = []
    tci_candidates = []
    for directory in [folder, *[p for p in folder.rglob("*") if p.is_dir()]]:
        detail_count = sum(1 for path in directory.glob("*.xlsx") if "交安设施现场检测" in path.name and "明细" in path.name)
        disease_count = sum(1 for path in directory.glob("*.xlsx") if "交安设施现场检测" in path.name and "病害清单" in path.name)
        if detail_count:
            detail_candidates.append((detail_count, directory))
        if disease_count:
            disease_candidates.append((disease_count, directory))
        if directory.name and "TCI" in directory.name:
            tci_candidates.append((sum(1 for path in directory.glob("*.xlsx")), directory))
    detail = max(detail_candidates, default=(0, None))[1]
    disease = max(disease_candidates, default=(0, None))[1]
    tci = max(tci_candidates, default=(0, None))[1]
    # 旧文件名优先；真实文件（整理/明细/病害命名）走回退识别。
    if summary is None:
        summary = _fallback_summary_file(folder)
    if detail is None:
        detail = _fallback_detail_dir(folder)
    if disease is None:
        disease = _fallback_disease_dir(folder)
    return summary, detail, disease, tci


class GuardrailApp(tk.Tk):
    MODES = (
        "完整生成（统计、报告、区间明细）",
        "仅更新区间明细",
    )

    def __init__(self):
        super().__init__()
        self.title("G210交安设施统计与报告工具")
        self.geometry("980x760")
        self.minsize(860, 620)
        self.queue = Queue()
        self.running = False
        self.vars = {name: tk.StringVar() for name in ("project", "summary", "detail", "disease", "tci", "output")}
        self.template_name = tk.StringVar(value=next(iter(BUILTIN_REPORT_TEMPLATES)))
        self.item_vars = {
            "沿线设施": tk.BooleanVar(value=False),
            "中心高度": tk.BooleanVar(value=True),
            "螺栓缺失": tk.BooleanVar(value=True),
        }
        self.mode = tk.StringVar(value=self.MODES[0])
        self._build_ui()
        self.after(100, self._poll_queue)

    def _build_ui(self):
        style = ttk.Style(self)
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 16, "bold"))
        container = ttk.Frame(self, padding=18); container.pack(fill="both", expand=True)
        ttk.Label(container, text="G210交安设施统计与报告工具", style="Title.TLabel").pack(anchor="w", pady=(0, 14))
        form = ttk.LabelFrame(container, text="文件与文件夹", padding=12); form.pack(fill="x")
        rows = [
            ("项目文件夹", "project", self.choose_project, "选择文件夹"),
            ("分段汇总表", "summary", lambda: self.choose_file("summary", [("Excel", "*.xlsx")]), "选择文件"),
            ("检测明细文件夹", "detail", lambda: self.choose_folder("detail"), "选择文件夹"),
            ("病害清单文件夹", "disease", lambda: self.choose_folder("disease"), "选择文件夹"),
            ("输出文件夹", "output", lambda: self.choose_folder("output"), "选择文件夹"),
        ]
        for index, (label, key, command, button_text) in enumerate(rows[:4]):
            ttk.Label(form, text=label, width=15).grid(row=index, column=0, sticky="w", pady=5)
            ttk.Entry(form, textvariable=self.vars[key]).grid(row=index, column=1, sticky="ew", padx=8, pady=5)
            ttk.Button(form, text=button_text, command=command, width=12).grid(row=index, column=2, pady=5)
        template_row = 4
        ttk.Label(form, text="Word报告模板", width=15).grid(row=template_row, column=0, sticky="w", pady=5)
        ttk.Combobox(
            form,
            textvariable=self.template_name,
            values=tuple(BUILTIN_REPORT_TEMPLATES),
            state="readonly",
        ).grid(row=template_row, column=1, sticky="ew", padx=8, pady=5)
        ttk.Label(form, text="内置配置", foreground="#777777", width=12).grid(row=template_row, column=2, pady=5)
        output_label, output_key, output_command, output_button = rows[-1]
        output_row = 5
        ttk.Label(form, text=output_label, width=15).grid(row=output_row, column=0, sticky="w", pady=5)
        ttk.Entry(form, textvariable=self.vars[output_key]).grid(row=output_row, column=1, sticky="ew", padx=8, pady=5)
        ttk.Button(form, text=output_button, command=output_command, width=12).grid(row=output_row, column=2, pady=5)
        form.columnconfigure(1, weight=1)

        items = ttk.LabelFrame(container, text="处理细分项", padding=10); items.pack(fill="x", pady=(14, 0))
        for name in ("沿线设施", "中心高度", "螺栓缺失"):
            ttk.Checkbutton(items, text=name, variable=self.item_vars[name]).pack(side="left", padx=(0, 22))
        ttk.Label(items, text="沿线设施暂保留选择项，尚未开发", foreground="#777777").pack(side="left")

        action = ttk.LabelFrame(container, text="运行功能", padding=12); action.pack(fill="x", pady=14)
        ttk.Combobox(action, textvariable=self.mode, values=self.MODES, state="readonly", width=45).pack(side="left", fill="x", expand=True)
        self.run_button = ttk.Button(action, text="开始运行", command=self.start_run, width=14); self.run_button.pack(side="left", padx=(12, 0))
        self.open_button = ttk.Button(action, text="打开输出文件夹", command=self.open_output, width=16); self.open_button.pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(container, mode="indeterminate"); self.progress.pack(fill="x", pady=(0, 10))
        log_frame = ttk.LabelFrame(container, text="运行日志", padding=8); log_frame.pack(fill="both", expand=True)
        self.log_box = scrolledtext.ScrolledText(log_frame, wrap="word", font=("Consolas", 10), state="disabled")
        self.log_box.pack(fill="both", expand=True)
        ttk.Label(container, text="提示：若提示文件被占用，请先关闭已打开的Excel或Word文件。", foreground="#666666").pack(anchor="w", pady=(8, 0))

    def log(self, text):
        self.queue.put(("log", text))

    def choose_project(self):
        folder = filedialog.askdirectory(title="选择包含统计资料的项目文件夹")
        if not folder:
            return
        self.vars["project"].set(folder)
        summary, detail, disease, tci = discover_paths(folder)
        if summary: self.vars["summary"].set(str(summary))
        if detail: self.vars["detail"].set(str(detail))
        if disease: self.vars["disease"].set(str(disease))
        if tci: self.vars["tci"].set(str(tci))
        self.vars["output"].set(str(summary.parent if summary else Path(folder)))
        self._append_log("已导入项目文件夹并自动识别相关文件。")

    def choose_file(self, key, types):
        path = filedialog.askopenfilename(filetypes=types)
        if path: self.vars[key].set(path)

    def choose_folder(self, key):
        path = filedialog.askdirectory()
        if path: self.vars[key].set(path)

    def build_config(self):
        project = Path(self.vars["project"].get() or ".").resolve()
        output = Path(self.vars["output"].get() or project).resolve()
        template_name = self.template_name.get()
        if template_name not in BUILTIN_REPORT_TEMPLATES:
            raise ValueError(f"未知的内置Word报告模板：{template_name}")
        return Config(
            project_dir=project,
            summary_xlsx=Path(self.vars["summary"].get()).resolve(),
            detail_dir=Path(self.vars["detail"].get()).resolve(),
            template_docx=BUILTIN_REPORT_TEMPLATES[template_name].resolve(),
            output_dir=output,
            disease_dir=Path(self.vars["disease"].get()).resolve() if self.vars["disease"].get() else None,
        )

    def validate(self, config):
        mode = self.mode.get()
        selected = {name for name, variable in self.item_vars.items() if variable.get()}
        if not selected:
            raise ValueError("请至少勾选一个处理细分项。")
        if not selected.intersection({"中心高度", "螺栓缺失"}):
            raise ValueError("沿线设施功能暂未开发，请同时勾选中心高度或螺栓缺失。")
        if mode == self.MODES[0]:
            if not config.summary_xlsx.is_file(): raise FileNotFoundError("请选择有效的分段汇总表。")
            if not config.detail_dir.is_dir(): raise FileNotFoundError("请选择有效的检测明细文件夹。")
            if "螺栓缺失" in selected and (config.disease_dir is None or not config.disease_dir.is_dir()):
                raise FileNotFoundError("处理螺栓缺失时，请选择有效的病害清单文件夹。")
            if not config.template_docx.is_file():
                raise FileNotFoundError(f"内置报告模板不存在：{config.template_docx}")
        elif mode == self.MODES[2]:
            if "中心高度" not in selected: raise ValueError("更新区间明细仅适用于中心高度，请勾选中心高度。")
            if not config.out_xlsx.is_file(): raise FileNotFoundError(f"未找到统计工作簿：{config.out_xlsx}")
            if not config.summary_xlsx.is_file(): raise FileNotFoundError("请选择有效的分段汇总表。")
        elif mode == self.MODES[3]:
            if "中心高度" not in selected: raise ValueError("生成图表仅适用于中心高度，请勾选中心高度。")
            if not config.out_xlsx.is_file(): raise FileNotFoundError(f"未找到统计工作簿：{config.out_xlsx}")

    def start_run(self):
        if self.running:
            return
        try:
            config = self.build_config(); self.validate(config)
        except Exception as error:
            messagebox.showerror("参数错误", str(error)); return
        self.running = True; self.run_button.configure(state="disabled"); self.progress.start(10)
        self._append_log("=" * 60); self._append_log(f"开始：{self.mode.get()}")
        selected = {name for name, variable in self.item_vars.items() if variable.get()}
        Thread(target=self._worker, args=(config, self.mode.get(), selected), daemon=True).start()

    def _worker(self, config, mode, selected):
        try:
            options = {
                "process_height": "中心高度" in selected,
                "process_bolts": "螺栓缺失" in selected,
                "process_alongline": "沿线设施" in selected,
            }
            if mode == self.MODES[0]:
                generate_statistics_and_report(config, self.log, **options)
            else:
                add_interval_sheets(config.out_xlsx, config.summary_xlsx, self.log)
            self.queue.put(("done", "处理完成。"))
        except PermissionError:
            self.queue.put(("error", "文件被占用，无法保存。请关闭Excel或Word后重试。"))
        except Exception as error:
            self.queue.put(("error", f"{error}\n\n{traceback.format_exc()}"))

    def _poll_queue(self):
        try:
            while True:
                kind, text = self.queue.get_nowait()
                if kind == "log": self._append_log(text)
                elif kind == "done":
                    self._append_log(text); self._finish(); messagebox.showinfo("完成", text)
                elif kind == "error":
                    self._append_log("错误：" + text); self._finish(); messagebox.showerror("运行失败", text)
        except Empty:
            pass
        self.after(100, self._poll_queue)

    def _finish(self):
        self.running = False; self.progress.stop(); self.run_button.configure(state="normal")

    def _append_log(self, text):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", str(text) + "\n"); self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def open_output(self):
        folder = Path(self.vars["output"].get() or self.vars["project"].get() or ".")
        if folder.is_dir(): os.startfile(folder)
        else: messagebox.showwarning("提示", "输出文件夹不存在。")


# ==================== 广东项目模板层 ====================

def _norm_header(value):
    return re.sub(r"[\s（）()\[\]【】_:：\-]+", "", str(value or "")).lower()


def _filename_k_range(path):
    """源文件名中的 K 区间（G234上行K3140K3144.xlsx → (3140000, 3144000)）。

    只认 4 位 K（既有丢弃口径，见 _convert）；不因放宽而改变已验收市的行数与段长。
    """
    match = re.search(r"K(\d{4})K(\d{4})", Path(str(path)).name)
    if not match:
        return None
    lo, hi = (int(value) * 1000 for value in match.groups())
    return (min(lo, hi), max(lo, hi))


def _filename_k_span(path):
    """文件名声明的 K 区间（1~5 位：G234上行K3140K3144 与 S16下行K18K0 都识别），单位米。

    只用于「一行有多个桩号列时挑哪一列」，不用于丢弃行 —— 丢弃仍走 _filename_k_range 的 4 位口径。
    """
    match = re.search(r"K(\d{1,5})K(\d{1,5})", Path(str(path)).name, re.I)
    if not match:
        return None
    lo, hi = (int(value) * 1000 for value in match.groups())
    return (min(lo, hi), max(lo, hi))


# 护栏明细各指标的桩号列取用优先级（模板「五、」：高度按原始桩号的 4 m 输出栅格计数，
# 故高度优先原始桩号；螺栓按电子修正桩号）。
GD_STATION_COLUMNS = {
    "height": ("原始桩号", "电子修正桩号", "标注修正桩号", "桩号"),
    "bolt": ("电子修正桩号", "标注修正桩号", "原始桩号", "桩号"),
}

# ponytail: 改列的「列间偏差」下限 30 km —— 实测两个制度分布分离得很开：
# ① 栅格错位（原始桩号 4 m 输出栅格 vs 电子修正桩号并点）健康市最大 28.0 km（东莞 S3 螺栓）；
# ② 坏值（设备计数回绕 9999+988、串到别段 K2159/K2482）最小 48.7 km（广州 G0421 螺栓）。
# 两制度之间 28~48 km 是空的。上下界各留 7%+ / 37% 余量；新增制度时需重跑
# python/p1b_spread.py 复核间隔是否仍为空，否则调这个常量。
GD_STATION_SWITCH_MIN = 30_000


def _pick_station(row, kind, source=""):
    """取该行的桩号原值：按 kind 的优先级取第一个非空列；当优先列不可信时改用文件名 K 区间内的列。

    采集明细里同一行有 原始/标注修正/电子修正 三个桩号列，优先列在两种情形下不可信：
    ① 设备计数回绕，原始桩号出现 9999+988 这类溢出值（广州 S16/S18/S29/S39/S2(广河)、
       深圳 G0422/G94/S28/S29/S31/S33/S86、惠州 S6/S30/S9925、中山 S5311、江门 S49、东莞 S22/G220 实测）；
    ② 串到别的路段，广州 S2 的电子修正桩号 K2159、S29 的 K2482（文件名声明 K0-K70 / K0-K45）。
    两者都落在文件名声明的 K 区间之外，且与同行其它列的偏差 ≫ 栅格错位，故改用区间内的列。

    门槛：只在偏差 ≥ GD_STATION_SWITCH_MIN 时改列。只落在区间外、但偏差只有几米~几十公里的，
    是文件名少报路段（如东莞 S3 文件名 K30K20 而数据 K2-K30）或栅格错位，保持原优先级不动 ——
    否则会把已验收市的段长改掉。全部列都不在区间内时同样保持原优先级，不丢行。
    """
    span = _filename_k_span(source)
    columns = GD_STATION_COLUMNS[kind]
    first = None
    primary_meters = None
    for name in columns:
        raw = _first(row, name)
        if raw in (None, ""):
            continue
        meters = _station_first(raw)
        if first is None:
            first, primary_meters = raw, meters
            if not span or meters is None or span[0] <= meters <= span[1]:
                return raw
            continue
        if not span or meters is None or not (span[0] <= meters <= span[1]):
            continue
        if primary_meters is not None and abs(meters - primary_meters) < GD_STATION_SWITCH_MIN:
            continue  # 偏差只是栅格错位，不是坏值
        return raw
    return first


def _first(row, *aliases, default=None):
    normalized = {_norm_header(k): v for k, v in row.items()}
    for alias in aliases:
        key = _norm_header(alias)
        if key in normalized and normalized[key] not in (None, ""):
            return normalized[key]
    return default


def _float(value):
    try:
        number = float(value)
        return number if math.isfinite(number) and not isinstance(value, bool) else None
    except (TypeError, ValueError, OverflowError):
        return None


def _route(value):
    return re.sub(r"\s+", "", str(value or "")).upper()


def station_span(rows):
    """按 (路线,方向) 分组后从明细行聚合桩号范围（模板②表「起止桩号」）。"""
    values = [float(row["station_m"]) for row in rows if row.get("station_m") is not None]
    return (min(values), max(values)) if values else (None, None)


def parse_station_range(text):
    """解析“K966+300~K966+200”“K22~K23”这类桩号范围，返回排序后的 (米, 米)；解析不出返回 None。"""
    import re
    values = []
    for part in re.split(r"~|～|--|—|－|至", str(text or "")):
        match = re.search(r"(\d+)(?:\s*[+＋]\s*(\d+))?", part)
        if match:
            values.append(int(match.group(1)) * 1000 + int(match.group(2) or 0))
    return tuple(sorted(values)) if len(values) == 2 else None


def merge_spans(items):
    """合并相邻/重叠桩号区间。"""
    merged = []
    for start, end in sorted(items):
        if merged and start <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], end)
        else:
            merged.append([start, end])
    return [tuple(pair) for pair in merged]


def load_manual_review_spans(path):
    """读取一期人工复核对比文件的桩号范围，按 (路线, 方向) 合并为米制区间。"""
    spans = {}
    if not path:
        return spans
    try:
        import openpyxl
        book = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception:
        return spans
    for sheet in book.worksheets:
        indices = None
        for row in sheet.iter_rows(values_only=True):
            cells = ["" if cell is None else str(cell).strip() for cell in row]
            if not "".join(cells):
                continue
            if indices is None:
                indices = {token: next((index for index, cell in enumerate(cells) if token in cell), None)
                           for token in ("路线", "方向", "桩号")}
                continue
            route_index, direction_index, range_index = (indices[token] for token in ("路线", "方向", "桩号"))
            if None in (route_index, direction_index, range_index):
                continue
            span = parse_station_range(cells[range_index])
            if not span or not cells[route_index]:
                continue
            spans.setdefault((cells[route_index], cells[direction_index]), []).append(span)
    return {key: merge_spans(value) for key, value in spans.items()}


def review_covered_km(phase_spans, key, start_km, end_km):
    """该抽检路段内被人工复核覆盖的里程（公里）；phase_spans 为一期复核的 {(路线, 方向): [区间]}。"""
    spans = (phase_spans or {}).get(key)
    if not spans or start_km is None or end_km is None:
        return 0.0
    low, high = sorted((float(start_km), float(end_km)))
    total = sum(max(0.0, min(end, high * 1000) - max(start, low * 1000)) for start, end in spans)
    return total / 1000


def detect_manual_before_workbook(project_dir):
    """在项目资料里查找“进场前人工自动化对比”文件；找不到返回 None（备注不写进场前段）。"""
    root = Path(project_dir)
    if not root.is_dir():
        return None
    for path in sorted(root.rglob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        if "进场前" in str(path) and ("对比" in path.name or "自动化" in path.name):
            return path
    return None


def load_route_segments(path):
    """读取路线表中的抽检路段行（附件清单表口径：起点/终点/里程为公里数，管养单位取路线表）。"""
    rows = []
    if not path:
        return rows
    try:
        import openpyxl
        book = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception:
        return rows
    for sheet in book.worksheets:
        header = {}
        city = block_type = ""
        for row in sheet.iter_rows(values_only=True):
            cells = ["" if cell is None else str(cell).strip() for cell in row]
            if not "".join(cells):
                continue
            if not header:
                if "地市" in cells and any(cell in ("起点桩号", "路段起点") for cell in cells):
                    header = {cell: index for index, cell in enumerate(cells) if cell}
                continue

            def _cell(*names):
                return next((cells[header[name]] for name in names
                             if name in header and header[name] < len(cells)), "")

            if cells[0]:                      # 首列块标记：高速 / 国省道
                block_type = cells[0]
            if _cell("地市"):
                city = _cell("地市").rstrip("市")
            route, direction = _cell("路线", "线路号"), _cell("方向")
            start = _float(_cell("起点桩号", "路段起点"))
            if not route or start is None:
                continue
            kind = _cell("道路等级", "类型") or block_type
            rows.append({
                "city": city, "route": route, "direction": direction,
                "category": "高速公路" if "高速" in kind else ("普通国省道" if kind else ""),
                "start": start, "end": _float(_cell("终点桩号", "路段终点")),
                "length": _float(_cell("检测里程（km）", "里程", "检测里程")),
                "manager": _cell("管养单位"),
            })
    return rows


OWNER_ALIASES = ("经营主体", "主体", "经营主体名称")

# 路线号录入笔误勘误表：{(归一地市, 路线号原值): 正确路线号}。
# 依据：用户 2026-09-27 澄清「茂名/G291 实为 S291」——茂名该段抽检的是省道 S291，
# 源表「进场前人工复核原始数据（缺标线数据）.xlsx」螺栓缺失 sheet 的 5 行把 S291 误写成 G291。
# P1n-2：同一份「进场前人工自动化对比_统计结果.xlsx」螺栓 sheet 另有 1 行连前缀都没写（`291`），
# 桩号 K165~K166 落在同文件高度/标线 sheet 记为 S291 的 K164+900~K167+200 区间内 → 同一处笔误。
# 该行若不归一，category_for_records 返回 None（分类表没有裸 `291` 行），整行被丢出报告。
# 分类表本就有 茂名/S291/普通国省道 行，故不新增表行，只在记录用路线号归一入口改写。
ROUTE_CODE_ERRATA = {("茂名", "G291"): "S291", ("茂名", "291"): "S291"}

# 实测抽检段落入两个经营主体时写入分类表第 5 列的取值（非主体名，仅作「本行跨主体」标记）。
CROSS_OWNER = "跨主体"


# 抽检明细文件名：<路线码><方向/补测等描述>K<起>K<止>-<路线码>-…-明细-<时间戳>.xlsx
# K 值是桩号公里数（K63.4 允许小数）。经营主体消歧按该实测区间落在附件哪一段判定。
MEASURED_SEGMENT_FILE = re.compile(r"^(?P<code>[GSXY]\d{1,5})[^\\/]*?K(?P<a>\d+(?:\.\d+)?)K(?P<b>\d+(?:\.\d+)?)", re.I)


def load_measured_segments(root, route_index):
    """明细文件名里的实测抽检桩号区间（km）→ {(归一地市, 路线号): [(起, 止), …]}。

    用途：路线分类表没有桩号列时，用「本市实测抽检段落入附件哪一段」判定经营主体
    （P1d-b 裁决：不加桩号列，改用实测段消歧）。找不到目录/市名时返回空 dict，绝不抛异常。
    """
    if not root or not route_index or not Path(root).is_dir():
        return {}
    names = sorted({norm for norm, _ in route_index.mapping if norm})
    base = Path(root)
    # root 本身是某个市目录时直接用它，否则逐个子目录找市名（与扫描器 _city_from_folder 同一口径）
    folders = [base] if any(name in base.name for name in names) else [p for p in base.iterdir() if p.is_dir()]
    segments = {}
    for folder in sorted(folders):
        city = next((name for name in names if name in folder.name), None)
        if not city:
            continue
        for path in folder.rglob("*.xlsx"):
            if path.name.startswith("~$"):
                continue
            match = MEASURED_SEGMENT_FILE.match(path.name)
            if not match:
                continue
            lo, hi = sorted((float(match.group("a")), float(match.group("b"))))
            segments.setdefault((city, _route(match.group("code"))), []).append((lo, hi))
    return segments


def merge_owner_table(route_index, route_xlsx, owner_xlsx=None, log=None):
    """补齐经营主体来源：附件分段索引 + 同目录带该列的路线表。

    优先级：`owner_xlsx`（附件《全省高速公路基础信息表》，唯一带桩号、可按 P2 消歧的源）→
    同目录的 `路线分类表（含经营主体）.xlsx`（P1d 导出，只有 (地市,路线) 级归属）→ 都没有就原样返回，
    后续如实显示「—」/「缺失」，不猜。

    P1m：`owner_xlsx` 为 None 时**自动发现**附件（与本函数的「同目录发现」同思路），
    否则调用方忘传 `GuangdongConfig.owner_xlsx` 就会静默退化成路线级口径 ——
    P1m 实测 `run_gd.py`（传）与 `run_gd_all.py`（没传）跑出的广州 ①表 655=540+115 vs 655=541+114。
    任何入口跑同一市必须逐值相同，所以发现放在这里这个唯一出口，而不是每个 run_*.py 各传一次。
    """
    if not getattr(route_index, "owner_index", None):
        found = Path(owner_xlsx) if owner_xlsx else discover_owner_xlsx(route_xlsx)
        if found and Path(found).is_file():
            route_index.owner_index = load_owner_index(found)
            if log:
                log(f"经营主体附件（权威来源，带桩号）：{found}"
                    f"（{'显式传入' if owner_xlsx else '自动发现'}），分段索引 {len(route_index.owner_index)} 组")
        elif log:
            log("经营主体附件：未找到（显式路径不存在且未自动发现），"
                "退化为路线分类表「经营主体」列（路线级、无桩号）——①表两行可能不闭合")
    if getattr(route_index, "owners", None) or not route_xlsx:
        return route_index
    try:
        folder = Path(route_xlsx).parent
    except Exception:
        return route_index
    for path in sorted(folder.glob("*.xlsx")) if folder.is_dir() else ():
        if path == Path(route_xlsx) or path.name.startswith("~$"):
            continue
        try:
            other = RouteCategoryIndex.from_file(path)
        except Exception:
            continue
        if getattr(other, "owners", None):
            route_index.owners = other.owners
            return route_index
    return route_index


def discover_owner_xlsx(route_xlsx=None):
    """自动发现附件《2026年全省高速公路基础信息表.xlsx》：路线表同目录 → 引擎仓 references → 常见项目目录。

    找不到返回 None（调用方如实退化为路线级口径并打日志，不静默猜）。
    ponytail: 只认这一个文件名；将来附件改名时改这里一处，升级路径 = 读配置里的候选名列表。
    """
    name = "2026年全省高速公路基础信息表.xlsx"
    folders = []
    if route_xlsx:
        try:
            folders.append(Path(route_xlsx).parent)
        except Exception:
            pass
    folders.append(application_root() / "references")
    folders.append(application_root() / "requirements")
    for folder in folders:
        try:
            hit = Path(folder) / name
            if hit.is_file():
                return hit
        except Exception:
            continue
    return None


def load_owner_index(path):
    """附件《全省高速公路基础信息表》→ {(归一地市, 路线号): [分段…]}，经营主体按原值读入。

    惰性加载：文件不存在/读不出返回空 dict（调用方视为「无附件」，不报错）。
    分段元素 {start, end, owner, keeper}；起点>止点时归一（附件实测已升序，仍做保护）。
    """
    if not path or not Path(path).is_file():
        return {}
    index = {}
    try:
        import openpyxl
        book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return {}
    try:
        for sheet in book.worksheets:
            header = None
            for values in sheet.iter_rows(values_only=True):
                cells = [_norm_header(v) for v in values]
                if header is None:
                    if "路线编码" in cells and "地市" in cells:
                        header = {name: i for i, name in enumerate(cells) if name}
                    continue

                def _cell(name):
                    index_ = header.get(_norm_header(name)) if header else None
                    return values[index_] if index_ is not None and index_ < len(values) else None

                code = _route(_cell("路线编码"))
                city = RouteCategoryIndex._norm_city(_cell("地市"))
                if not code or not city:
                    continue
                start, end = _float(_cell("起点桩号")), _float(_cell("止点桩号"))
                lo, hi = sorted(v for v in (start, end) if v is not None) or (None, None)
                owner = _first(dict(zip(header, values)), *OWNER_ALIASES)
                index.setdefault((city, code), []).append({
                    "start": lo,
                    "end": hi,
                    "owner": re.sub(r"\s+", "", str(owner)) if owner else "",
                    "keeper": str(_cell("管养单位") or "").strip(),
                })
    finally:
        book.close()
    return index


def match_owner(owner_index, city, route, start=None, end=None, keeper=None):
    """按 A1 §6.1 的 P0→P4 顺序定经营主体；命中多个主体且无法消歧 → None。

    P0 路线编码等值 + P1 地市等值 → P2 桩号区间重叠（有起止桩号时）
    → P3 管养单位等值 → P4 全城该路线唯一主体。
    禁止：按路线号跨市唯一判定、禁止「广东/省/集团」等关键字回退。

    P1h：P3 只在管养单位**覆盖整段查询桩号**时才用来收窄。清单表行的 `manager` 是
    明细行里取到的第一个管养单位（`_gd_inspection_segments_from_detail` 的 `next(...)`），
    跨主体路线上它只管一段；拿它收窄会把整行判给那一个主体（G0425 广州 40km 整行判成
    非集团、G4 39km 同），正是 A1 §6.2 禁止的「取多数/取第一段」。管养单位盖不满查询区间
    时放弃 P3，保留 P2 的多主体结果交调用方按桩号拆分。
    """
    segs = (owner_index or {}).get((RouteCategoryIndex._norm_city(city), _route(route)))
    if not segs:
        return None
    lo, hi = _float(start), _float(end)
    if lo is not None and hi is not None:
        lo, hi = min(lo, hi), max(lo, hi)
    else:
        lo = hi = None
    hits = [s for s in segs if lo is None or not s["start"] is None
            and min(s["end"], hi) - max(s["start"], lo) > 0] or segs
    if keeper:
        narrowed = [s for s in hits if s["keeper"] == keeper and s["start"] is not None]
        # 收窄后仍须盖住整段查询桩号，否则说明管养单位只覆盖其中一段，不能据此定整行主体。
        if narrowed and (lo is None or (min(s["start"] for s in narrowed) <= lo + 1e-6
                                        and max(s["end"] for s in narrowed) >= hi - 1e-6)):
            hits = narrowed
    values = {s["owner"] for s in hits if s["owner"]}
    return values.pop() if len(values) == 1 else None


def gd_number_text(value):
    """公里数显示：整数不带小数点，小数按原值（附件清单表口径）。"""
    if value is None:
        return "—"
    return f"{float(value):g}"


# 管养单位名 → 分支标识（P1i）。同一分支在附件与明细明细行里公司名写法不同：
# 附件「广东省公路建设有限公司博深分公司」vs 明细「广东博大高速公路有限公司博深分公司」，
# P3 的等值比较因此落空，附件未覆盖的桩号空洞判不出主体（深圳 G0422 差 7.131km）。
# 分支名（分公司/管理处/养护所/管理所 之前的部分）才是稳定标识；无分支后缀的单位名
# 原样返回（这类名称在附件内本身唯一）。实测附件 129 个管养单位归一后
# 仅 1 组碰撞且同为「非集团」，无跨经营主体歧义。
GD_BRANCH_SUFFIXES = ("分公司", "管理处", "养护所", "管理所")
GD_CORP_ENDINGS = ("有限公司", "股份公司", "公司", "集团", "中心", "管理局")


def gd_branch_key(name):
    """管养单位名 → 分支标识（去空白 + 取分支后缀前的分支名）。"""
    text = re.sub(r"\s+", "", str(name or ""))
    for suffix in GD_BRANCH_SUFFIXES:
        index = text.rfind(suffix)
        if index > 0:
            head, cut = text[:index], 0
            for ending in GD_CORP_ENDINGS:
                at = head.rfind(ending)
                if at >= 0:
                    cut = max(cut, at + len(ending))
            return head[cut:] + suffix
    return text


# 省交通集团归属判定（数据源无该字段，按管养单位名称启发式判定）：
# 依据：广东高速公路省交通集团下属单位名称普遍含「广东/省/集团/交通集团」（如「广东省高速公路有限公司」）。
# 局限：名称不含关键词的省属单位会被判为「非省交通集团」，需人工核对；判定不出时按「非省交通集团」计。
PROVINCIAL_MANAGER_KEYWORDS = ("广东", "省", "集团", "交通集团")


def gd_manager_lookup(route_segments, city=""):
    """C4：路线分类表（清单表同源）的管养单位映射。

    返回 {(路线, 方向): 管养单位} 以及该路线在路线表中只对应一个管养单位时的 {(路线, ""): 管养单位}
    兜底键（明细行缺方向时用）。明细行 manager 写法不统一（同一单位多种简称/全称）会让同一管养单位
    在 ① 表被拆成多行，与引言句「分属 N 家管养单位」自相矛盾，故 ①/②/③ 表一律以路线表口径为准。
    数据源即清单表行本身（与引言句、清单表完全同源）；标注了他市的路线表行剔除，无 city 字段的明细分
    组行保留（明细聚合口径无 city 字段）。
    """
    wanted = str(city or "").rstrip("市")
    lookup = {}
    per_route = {}
    for row in route_segments or []:
        row_city = str(row.get("city") or "").strip()
        if wanted and row_city and row_city != wanted:
            continue
        route = str(row.get("route") or "")
        name = str(row.get("manager") or "").strip()
        if not route or not name:
            continue
        lookup.setdefault((route, str(row.get("direction") or "")), name)
        per_route.setdefault(route, set()).add(name)
    for route, names in per_route.items():
        if len(names) == 1:
            lookup.setdefault((route, ""), next(iter(names)))
    return lookup


def gd_manager_name(lookup, row):
    """按 (路线, 方向) → (路线, "") 取路线表管养单位；路线表缺该路段时如实写「—」，不回退明细行口径。"""
    route = str(row.get("route") or "")
    return (lookup.get((route, str(row.get("direction") or ""))) or lookup.get((route, "")) or "—")


def gd_route_managed_rows(rows, lookup):
    """C4：把明细行的 manager 覆写为路线表口径（路线表整体缺失时保持原值，生产链路必有清单表）。"""
    if not lookup:
        return rows
    return [dict(row, manager=gd_manager_name(lookup, row)) for row in rows]


# 模板-1「（三）工作建议」3.养护提升建议 原文（该节模板无数据占位，逐字照抄；C1）：
# 结构 = 引导段（仿宋_GB2312 12，不加粗）+ 4 个 `（n）` 引导段（同字体）+ 5 个 `①②③` 二级标题（黑体 12 bold）+ 正文段。
GD_MAINTENANCE_LEAD = "针对本次暴露出的高速与普通国省道差距大、老旧两波护栏病害集中、标线衰减快、螺栓易松脱、标志遮挡等共性问题，完善常态化养护机制，实现交安设施可持续良好运行。"
GD_MAINTENANCE_ADVICE = (
    ("（1）建立差异化、精准化的养护策略", (
        ("①高速公路，加强属地养护监管", (
            "督促运营管理单位压实养护主体责任，针对高速公路交安设施建立常态化巡查养护机制，明确日常巡查、定期检测的频次与责任到人，对发现的病害问题建立整改台账，按照轻重缓急安排养护资金及时处治，确保设施功能完好。",
        )),
        ("②普通国省道，统筹养护", (
            "由属地交通运输主管部门统筹做好普通国省道交安设施的养护规划与资金调配，针对老旧路段存量设施基础薄弱、病害多发的特点，按年度分批安排专项养护工程开展集中整治，优先解决老旧两波护栏防护能力不足、标线大面积衰减等影响通行安全的突出问题。督促养护部门按规范定期开展沿线交安设施巡查，对发现的标志歪倾、护栏变形、螺栓缺失等问题第一时间处置，实现小问题及时消险、大问题纳入计划整治的管控模式，逐步缩小与高速公路交安设施运行水平的差距。",
        )),
    )),
    ("（2）分类型完善设施养护策略", (
        ("①防护护栏", (
            "对老旧两波波形梁护栏建立专项档案，两波护栏高度合格率普遍偏低，在路面罩面、大修工程中同步复核调整护栏高度，路面加铺后同步抬升护栏，避免只铺路面不调护栏。",
            "高填方、桥梁过渡段、重载交通路段，将护栏螺栓紧固、缺失补装纳入季度常态化巡检，重点防范拼接螺栓松脱脱落。",
            "对开口护栏防护不足、应设未设防护等安全隐患，结合安全提升工程逐步完善。",
        )),
        ("②交通标线", (
            "高速公路继续保持较好养护水平，重点关注同一路线不同管养单位标线表现不均衡现象。",
            "普通国省道标线逆反射问题突出，建议推广高耐磨、长寿命标线材料，适当延长标线使用寿命；结合交通流量、磨损情况合理确定复划周期，扭转大量标线夜间失效的现状。",
            "路面改造同步做好旧标线清除，减少旧标线残留病害。",
        )),
        ("③交通标志与防眩设施", (
            "将标志树木修剪纳入常态化养护，定期清理路侧乔木对标志板遮挡问题；及时修复缺损、变形标志。",
            "高速公路加强防眩设施养护，重点处置防眩网锈蚀、植物防眩失效，降低防眩设施对TCI指标的扣分项；普通国省道结合照明条件按需设置防眩设施。",
        )),
    )),
    ("（3）资金保障与技术能力提升", (
        ("", (
            "加大资金投入：适当加大普通国省道交通安全设施专项养护资金投入。普通国省道交安设施短板突出，与高速公路差距巨大，在年度养护资金分配中，向标线翻新、护栏维修、螺栓加固倾斜，避免重路面轻交安。",
            "加强技术培训：加强一线养护人员技术培训，普及《公路技术状况评定标准》《公路交通安全设施设计规范》，解决部分基层施工人员及管理人员对护栏安装高度、设施设置规范不熟悉的问题，从源头减少新建、改造项目的交安设施质量缺陷。",
        )),
    )),
    ("（4）病害源头治理", (
        ("", (
            "在公路改扩建、路面罩面、护栏改造项目阶段，严格把控交安设施施工质量，严格复核护栏中心高度、螺栓扭矩、标线材料性能，从源头减少后期养护整改压力，实现建管养一体化提升。",
        )),
    )),
)
# 模板-1「（三）工作建议」1.重点路段处治建议（如有）原文（该节模板仅 1 段，无子标题、无表；C3）
GD_KEY_ROUTE_ADVICE = "针对典型状况不佳路段，组织排查补装，建立每月专项巡检台账，紧盯整改后运行状态。总结经验，提升全市普通国省道交安设施运行水平。问题分散、合格率低的路段，由属地管养单位制定半年整改计划，落实处治措施，按时报送进展，由上级交通运输主管部门抽查核验，确保整改落实。"


def manager_is_provincial(manager):
    text = str(manager or "")
    return any(keyword in text for keyword in PROVINCIAL_MANAGER_KEYWORDS)


def _gd_route_name(row, route_names, city=""):
    return f"{row['route']}{route_names.get(row['route']) or ''}（{manager_display(row.get('manager'), city)}）"


def gd_route_rate_sentence(category, indicator, rows, route_names, value, city=""):
    """②各抽检路段情况前置句（D10，合格率类，句型逐字照模板-1）。

    「抽检的各{类别}路段{指标}明细如下表所示。合格率超90%的路段分别有…，合格率较低的路段有…。」
    空名单如实降级为「本次无…」。
    """
    names = [(value(row), _gd_route_name(row, route_names, city)) for row in rows if value(row) is not None]
    high = [name for rate, name in names if rate >= 0.9]
    low = [name for rate, name in sorted(names)[:3]]
    return ("抽检的各%s路段%s明细如下表所示。" % (category, indicator)
            + ("合格率超90%的路段分别有" + "、".join(high) if high else "本次无合格率超90%的路段")
            + "，" + ("合格率较低的路段有" + "、".join(low) if low else "本次无合格率较低的路段") + "。")


def gd_route_missing_sentence(category, indicator, rows, route_names, value, city=""):
    """②各抽检路段情况前置句（D10，缺失率类，句型逐字照模板-1）。

    「抽检的各{类别}路段{指标}明细如下表所示。缺失率低于1%的路段分别有…，缺失率较高的路段有…。」
    """
    names = [(value(row), _gd_route_name(row, route_names, city)) for row in rows if value(row) is not None]
    low = [name for rate, name in names if rate < 0.01]
    high = [name for rate, name in sorted(names, reverse=True)[:3]]
    return ("抽检的各%s路段%s明细如下表所示。" % (category, indicator)
            + ("缺失率低于1%的路段分别有" + "、".join(low) if low else "本次无缺失率低于1%的路段")
            + "，" + ("缺失率较高的路段有" + "、".join(high) if high else "本次无缺失率较高的路段") + "。")


def gd_managers_in_order(rows):
    """按明细出现顺序取出管养单位（去重），① 管理单位级表格的行顺序。"""
    managers = []
    for row in rows:
        manager = str(row.get("manager") or "")
        if manager and manager not in managers:
            managers.append(manager)
    return managers


def gd_class_groups(rows, category):
    """① 分组：普通国省道只取 普通国道/普通省道（合计行由调用处补），跳过与分支同名的汇总项。"""
    return [(label, subset) for label, subset in gd_road_class_subsets(rows, category) if label != category]


def gd_height_segment_kinds(rows, item):
    """典型路段的两波/三波合格率：按该路段桩号范围的护栏明细重算（与 ①② 同口径）。"""
    subset = [row for row in rows
              if row.get("route") == item["route"] and row.get("direction") == item.get("direction")
              and row.get("station_m") is not None
              and item["start_m"] <= row["station_m"] < item["end_m"]]
    stats = GuangdongStatistics.height_overall(subset)
    return stats["二波"]["rate"], stats["三波"]["rate"]


def _contiguous_runs(values):
    """同值连续段起点与长度，用于清单表 `类型` 列的纵向合并（D4）。"""
    runs = []
    start = 0
    for index in range(1, len(values) + 1):
        if index == len(values) or values[index] != values[start]:
            runs.append((start, index - start))
            start = index
    return runs


def gd_inspection_segments(bundle):
    """抽检路段清单：优先取路线表中该市的抽检路段行（附件口径），路线表缺该市时回退明细行聚合。"""
    city = str(bundle.get("city") or "").strip()
    rows = [row for row in (bundle.get("route_segments") or [])
            if row.get("city") == city.rstrip("市")]
    if not rows:
        rows = _gd_inspection_segments_from_detail(bundle)
    phases = bundle.get("review_phases") or {}
    for row in rows:
        row["start_text"] = gd_number_text(row.get("start"))
        row["end_text"] = gd_number_text(row.get("end"))
        row["length_text"] = gd_number_text(row.get("length"))
        if row.get("length") is None:
            row["length"] = row.get("length_km") or 0.0
        row["length_km"] = float(row.get("length") or 0.0)
        # 备注口径（附件）：人工复核：进场前X公里，抽检后Y公里；两期都没覆盖则不写
        covered = [f"{phase}{review_covered_km(phases.get(phase), (row['route'], row['direction']), row.get('start'), row.get('end')):g}公里"
                   for phase in ("进场前", "抽检后")
                   if review_covered_km(phases.get(phase), (row["route"], row["direction"]), row.get("start"), row.get("end"))]
        row["remark"] = f"人工复核：{'，'.join(covered)}" if covered else ""
    return rows


def gd_mileage(inspection, category=None, manager=None, route=None, direction=None, class_prefix=None, owner=None):
    """抽检里程(km)：只从清单表那批抽检路段行求和；任一维度为 None = 该维不筛选。

    分组键维度：道路类别 / 管养单位 / 路线 / 方向 / 道路类别细分(普通国道G、普通省道S) / 经营主体。
    率仍由检测记录算，本函数只出里程。无匹配行或里程全缺 → None（显示「—」），
    不回退检测点数/100m 单元数/全路网里程（brief-v3 §6 + v3-mileage-perkm-contract §2.1）。
    """
    if owner is not None:
        # P1h：跨主体行按桩号分段拆出属于该主体的里程，而不是整行都算进去或都丢掉。
        # 拆分行的 `owner` 为 None（判不出单一归属），只要有 owner_spans 就要计入。
        # 走原来的整行循环前，先把这些行的分段里程加进来（下方循环会跳过它们）。
        hit = total = 0.0
        for row in inspection or ():
            spans = row.get("owner_spans")
            if not spans or row.get("category") != category:
                continue
            total += sum(b - a for a, b, value in spans if value == owner)
            hit += 1.0
    else:
        hit = total = 0.0
    for row in inspection or ():
        if category is not None and row.get("category") != category:
            continue
        if manager is not None and str(row.get("manager") or "") != str(manager):
            continue
        if route is not None and str(row.get("route") or "") != str(route):
            continue
        if direction is not None and str(row.get("direction") or "") != str(direction):
            continue
        if class_prefix is not None and not str(row.get("route") or "").upper().startswith(class_prefix):
            continue
        if owner is not None:
            if row.get("owner_spans"):
                continue          # 已按分段计过，不重复整行计入
            if str(row.get("owner") or "") != str(owner):
                continue
        km = row.get("length_km")
        if km is None:
            continue
        hit, total = 1.0, total + float(km)
    return total if hit else None


def gd_class_subset(rows, class_prefix=None):
    """按道路类别细分（G=普通国道 / S=普通省道）取检测记录子集。

    P1n-1：① 表「普通国道」「普通省道」两行此前只把 class_prefix 传给 gd_mileage（里程列对），
    速率列仍用**全市** rows 算，于是 45 张普通①表的「全市/国道/省道」三行速率完全相同。
    判据与 gd_mileage 里的 class_prefix 逐字一致，两处必须同步，否则里程与速率口径分裂。
    """
    if class_prefix is None:
        return list(rows or [])
    return [row for row in (rows or [])
            if str(row.get("route") or "").upper().startswith(class_prefix)]


def gd_km_text(km):
    """里程单元格：无值写「—」，否则整数显示（`{:,}` 对 float 不取整，会输出 226.818）。"""
    return "—" if km is None else f"{km:.0f}"


def gd_km_split_text(total, *parts):
    """① 表「全市 = 分项之和」三行的里程文本：全市普通四舍五入，分项按最大余数法补齐恒等。

    分项各自四舍五入后可能与全市差 1（深圳 163+75=238≠237），甲方一眼就会加出来。
    先各自向下取整，再把 round(全市)-Σfloor 的差值按小数余数从大到小逐个 +1，
    保证 sum(分项整数) == 整(全市)（深圳 163+75=238≠237 → 163+74=237）。
    无值（None）时该格照旧写「—」且不参与分配；全市为 None 时全走原逻辑。

    P1m：补齐只允许吃掉**取整残差**，绝不摊派真实差额。
    原实现对任意 target-Σfloor 都逐个 +1，于是把「附件未覆盖、判不出主体」的真实缺口也摊进
    两侧：广州 500.647+75.021=575.668（差 79.000）被摊成 540+115，①表整数看着闭合，
    却与说明句/引言的精确值直接矛盾（甲方把说明句的数字一加就露馅）。
    判据用**真实缺口** `total - Σparts` 而非缺口格数：取整残差最多 len(parts) 格，
    缺口 ≥0.5 km 必然来自未核定里程 → 各分项按自身值取整，如实显示不闭合，
    由 `gd_owner_gap_note` 点名路段解释。
    """
    total_text = gd_km_text(total)
    if total is None or any(part is None for part in parts):
        return (total_text,) + tuple(gd_km_text(part) for part in parts)
    target = int(f"{total:.0f}")
    base = [math.floor(part) for part in parts]
    order = sorted(range(len(parts)), key=lambda i: -(parts[i] - base[i]))
    steps = 0 if abs(total - sum(parts)) >= 0.5 else max(0, min(target - sum(base), len(parts)))
    for step in range(steps):
        base[order[step % len(order)]] += 1
    return (total_text,) + tuple(f"{value}" for value in base)


def gd_km_row_texts(groups):
    """① 表各行里程单元格文本；groups = [(行标签, stats 或 None)]，与表行同序。

    高速分支的「全市 / 非省交通集团 / 省交通集团」三行走 gd_km_split_text（最大余数法），
    其余行（普通国省道的全市/国道/省道、②表、清单表）一律走 gd_km_text 原逻辑。
    """
    texts = [gd_km_text(stats["km"]) if stats else "" for _label, stats in groups]
    labels = [label for label, _ in groups]
    if "非省交通集团" in labels and "省交通集团" in labels:
        i, j = labels.index("非省交通集团"), labels.index("省交通集团")
        total = next((groups[k][1]["km"] for k in range(i - 1, -1, -1) if groups[k][1]), None)
        if total is not None and groups[i][1] and groups[j][1]:
            # P1m：本市有高速抽检行时，某主体 0 段是**核定出来的 0 km**（云浮非省交通集团），
            # 写 0 不写「—」——「—」在本项目里表示「无数据」（全省行同理留空），两者不能混。
            # gd_mileage 对 0 段返回 None（hit=0），这里按 0 参与恒等，与引言/说明句同一口径。
            texts[i], texts[j] = gd_km_split_text(
                total, groups[i][1]["km"] or 0.0, groups[j][1]["km"] or 0.0)[1:]
    return texts


# 附件《全省高速公路基础信息表》`经营主体` 的两个原值；分组只用原值，禁止关键字回退（A1 §6.1 P6）。
GD_OWNER_PROVINCIAL = "集团"
GD_OWNER_NON_PROVINCIAL = "非集团"


def gd_owner_gap_note(inspection, category="高速公路"):
    """①表下方的经营主体核定说明句（P1i 硬要求：差额必须在产物正文可见，不能只打日志）。

    差额归零 → 事实句「已全部核定」；有差额 → 如实列出未定主体的路线与里程及原因
    （附件未覆盖该桩号 / 管养单位在附件内对应多个主体）。无数据时返回 None，不写占位。

    P1m：三个数与 ① 表**同源同取整** —— 走 `gd_km_split_text`（① 表 `gd_km_row_texts` 调的同一函数），
    句子按整数写，`合计 = 非集团 + 集团 (+ 差额)` 精确成立；原先这里按 `:.3f` 写精确值、
    ① 表按最大余数法写整数，两套取整，广州 500.647/75.021（说明句）对 540/115（①表）直接矛盾。
    差额分支列出**每条未定主体路段及其里程**（原先只列 `owner is None` 的路线，
    跨主体拆分失败的行 owner 非 None，整列出来是「—」，等于没交代差额来自哪里）。
    """
    if not inspection:
        return None
    total = gd_mileage(inspection, category=category)
    if total is None:
        return None
    parts = {owner: gd_mileage(inspection, category=category, owner=owner) or 0.0
             for owner in (GD_OWNER_NON_PROVINCIAL, GD_OWNER_PROVINCIAL)}
    gap = total - sum(parts.values())
    cells = gd_km_split_text(total, parts[GD_OWNER_NON_PROVINCIAL], parts[GD_OWNER_PROVINCIAL])
    if abs(gap) < 5e-4:
        return (f"本表「抽检{category}」经营主体已全部核定：非省交通集团 {cells[1]} km 与"
                f"省交通集团 {cells[2]} km 合计 {cells[0]} km，与该行里程一致，"
                f"「非省交通集团」与「省交通集团」两行之和等于抽检里程。")
    # 差额按**已显示的整数**倒推，保证句内 合计 = 非集团 + 集团 + 差额 逐位成立
    # （用 float 的 gap 四舍五入会与句子里另外三个整数对不上，甲方一加就露馅）。
    left = int(cells[0]) - int(cells[1]) - int(cells[2])
    return (f"本表「抽检{category}」经营主体核定说明：抽检里程合计 {cells[0]} km，"
            f"其中非省交通集团 {cells[1]} km、省交通集团 {cells[2]} km，"
            f"另有 {left} km 尚未核定经营主体，三者之和等于抽检里程合计；"
            f"该差额对应路段：{gd_owner_gap_routes(inspection, category)}。"
            f"原因为《2026年全省高速公路基础信息表》未覆盖上述路段的桩号区间，"
            f"且其管养单位在附件中未唯一对应一个经营主体，故不作推定、不摊派，"
            f"该部分里程不计入上述两行。")


def gd_owner_gap_routes(inspection, category="高速公路"):
    """差额里程对应的路段清单（`G2518 3.153 km`）；无差额返回「无」。

    差额的两种来源都要点名（P1m，江门 G2518 实测）：
      ① 整行判不出主体：`owner` 既不是集团也不是非集团（None / `跨主体` / 其他值）且无桩号分段；
      ② 拆分行**内部空洞**：`owner_spans` 首尾盖住整行（`gd_cross_owner_spans` 的守卫），
         但附件在该行桩号区间内没覆盖到，中间留了洞 —— 这段里程两侧都不落，
         却因为「这行有 spans」而被漏掉（P1m 首版就栽在这，江门说明句写了「无」）。
    两者都按**该行实际未落侧的里程**点名，不写「—」。
    """
    out = []
    for row in inspection or ():
        if row.get("category") != category:
            continue
        km = _float(row.get("length_km")) or 0.0
        spans = row.get("owner_spans")
        if spans:
            hole = km - sum(b - a for a, b, value in spans if value in (GD_OWNER_PROVINCIAL, GD_OWNER_NON_PROVINCIAL))
        elif str(row.get("owner") or "") in (GD_OWNER_PROVINCIAL, GD_OWNER_NON_PROVINCIAL):
            continue
        else:
            hole = km
        if hole > 5e-4:
            out.append(f"{row.get('route') or '—'} {hole:.3f} km")
    return "、".join(out) or "无"


# 跨主体行只在「≥2 个主体且桩号分段首尾完整覆盖该行起止」时才拆；其余情况返回 None，
# 让该行走单一 owner 的原路径 —— 与 P1b 行为逐值一致，不给非跨主体的行带来任何回归。
def gd_cross_owner_spans(computed, start, end):
    computed = [s for s in (computed or []) if s]
    lo, hi = _float(start), _float(end)
    if lo is None or hi is None or len(computed) < 2 or len({owner for _a, _b, owner in computed}) < 2:
        return None
    lo, hi = min(lo, hi), max(lo, hi)
    if abs(computed[0][0] - lo) > 1e-6 or abs(computed[-1][1] - hi) > 1e-6:
        return None
    return computed


def gd_owner_map(inspection):
    """{(路线, 方向): 经营主体原值 | [(起, 止, 经营主体)] 桩号分段}；判不出主体的行不出现。

    P1h：跨主体路线（G0425/G4 广州）的一行抽检里程横跨两个主体，单值装不下，①表按
    (路线,方向) 取值的率分母会把整行算进一侧。这里改挂 gd_cross_owner_spans 桩号分段，
    由 gd_mileage/gd_owner_rows 按桩号拆分，两侧都不重不漏（差额才可能归零）。
    注意：拆分行的 `owner` 是 None（判不出单一归属），但它有 owner_spans，必须进 map，
    否则整行里程在①表两侧都消失、差额反而变大。
    """
    out = {}
    for row in inspection or ():
        spans = row.get("owner_spans")
        if not (row.get("owner") or spans):
            continue
        key = (str(row.get("route") or ""), str(row.get("direction") or ""))
        out[key] = spans if spans else row["owner"]
    return out


def gd_owner_rows(rows, owner_map, owner):
    """明细行里归属于该经营主体的子集（①表经营主体两行的率分母/分子）。

    owner_map 的值是桩号分段时，明细行按自身 station_m 落段判定归属；落不进任何一段
    （跨主体切缝上的罕见测点）不计入，两侧都不吞。
    """
    out = []
    for row in rows or ():
        key = (str(row.get("route") or ""), str(row.get("direction") or ""))
        found = owner_map.get(key)
        if isinstance(found, list):
            km = _float(row.get("station_m"))
            if km is None:
                continue
            km /= 1000.0                      # 明细行 station_m 是米，桩号分段是 km
            last = found[-1][1]
            hit = next((value for a, b, value in found
                        if a <= km < b or (km == b and b == last)), None)
            if hit == owner:
                out.append(row)
        elif found == owner:
            out.append(row)
    return out


def _gd_inspection_segments_from_detail(bundle):
    """后备口径：三类指标明细行按 (道路类别, 路线, 方向) 合并出抽检路段范围。"""
    grouped = {}
    for key in ("marking", "height", "bolt"):
        for row in bundle.get(key) or []:
            # 类别判不出（None/空）的行与 build_bundles 的「无法判定类别」同一出口：只进 issues/优先处治，
            # 不进清单表、不进①②分母。按类别判定，不按地市二次判定（广河高速 S2 等有类别的行照常保留）。
            if row.get("station_m") is None or not row.get("route") or not str(row.get("category") or "").strip():
                continue
            identity = (str(row["category"]), str(row["route"]), str(row.get("direction") or ""))
            grouped.setdefault(identity, []).append(row)
    result = []
    for (category, route, direction), rows in grouped.items():
        start_m, end_m = station_span(rows)
        length_km = abs(float(end_m) - float(start_m)) / 1000 if start_m is not None else 0.0
        result.append({
            "category": category, "route": route, "direction": direction,
            "start": start_m / 1000 if start_m is not None else None,
            "end": end_m / 1000 if end_m is not None else None,
            "length_km": length_km, "length": length_km,
            "manager": next((row.get("manager") for row in rows if row.get("manager")), None),
            "remark": next((str(row.get("remark")) for row in rows if row.get("remark")), ""),
        })
    return sorted(result, key=lambda item: (0 if item["category"] == "高速公路" else 1,
                                            item["route"], item["direction"]))


def gd_km_detail(item, threshold=0.5, limit=4):
    """附件③表述：逐公里看，全线 N 个整公里段中有 M 段总体合格率低于 x%，其中 K… 段为 a%…。"""
    rated = []
    for row in (item.get("km_items") or []):
        value = row.get("rate")
        if value is None:
            values = [float(row[key]) for key in ("left_rate", "right_rate") if row.get(key) is not None]
            value = sum(values) / len(values) if values else None
        if value is not None:
            rated.append((row, float(value)))
    if not rated:
        return ""
    low = [(row, value) for row, value in rated if value < threshold]
    worst = sorted(rated, key=lambda pair: pair[1])[:limit]
    detail = "、".join(f"K{row.get('km')}段为{value:.1%}" for row, value in worst)
    if not low:
        return f"逐公里看，全线{len(rated)}个整公里段总体合格率均不低于{threshold:.0%}，最低为{detail}。"
    return (f"逐公里看，全线{len(rated)}个整公里段中有{len(low)}段总体合格率低于{threshold:.0%}，"
            f"其中{detail}，为该路段合格率最低的集中区间。")


def gd_gtype_note(gtype):
    """附件③表述：该段全部为三波形梁护栏，标准中心高度697mm。"""
    text = str(gtype or "").strip()
    if "三波" in text:
        return "该段全部为三波形梁护栏，标准中心高度697 mm。"
    if "两波" in text or "双波" in text:
        return "该段全部为两波形梁护栏，标准中心高度600 mm。"
    return ""


def _safe_city_component(value):
    text = str(value or "").strip()
    reserved = {"CON","PRN","AUX","NUL",*[f"COM{i}" for i in range(1,10)],*[f"LPT{i}" for i in range(1,10)]}
    if (not text or text in {".",".."} or Path(text).is_absolute() or
            re.search(r'[<>:"/\\|?*\x00-\x1f]', text) or text.rstrip(" .") != text or
            text.upper().split(".")[0] in reserved):
        raise ValueError(f"地市名称包含非法路径字符：{value}")
    return text


def _safe_excel_value(value):
    if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def _guardrail_type(value):
    text = str(value or "")
    return guardrail_type(value) or text.strip()


def _station_first(value):
    text = str(value or "")
    match = re.search(r"K?\s*(\d+(?:\.\d+)?)\s*\+\s*(\d+(?:\.\d+)?)", text, re.I)
    if match:
        return float(match.group(1)) * 1000 + float(match.group(2))
    match = re.search(r"K\s*(\d+(?:\.\d+)?)", text, re.I)
    if match:
        return float(match.group(1)) * 1000
    return station_to_m(text.split("-")[0].split("~")[0])


class RouteCategoryIndex:
    VALID = {"高速公路", "普通国省道"}
    OWNER_ALIASES = ("经营主体", "主体", "经营主体名称")

    def __init__(self, mapping, owners=None, names=None):
        self.mapping = {}
        self.display = {}
        for (city, route), value in mapping.items():
            norm = self._norm_city(city)
            self.mapping[(norm, _route(route))] = str(value).strip()
            self._remember_display(norm, str(city).strip())
        # 经营主体列（P1d）：{(归一地市, 路线号): 主体原值}；表里没有该列时为空 dict。
        self.owners = {}
        for (city, route), value in (owners or {}).items():
            text = re.sub(r"\s+", "", str(value or ""))
            if text:
                self.owners[(self._norm_city(city), _route(route))] = text
        # 路线名称反向索引（P1d-b）：{(归一地市, 路线名称): 该市唯一对应路线号}。
        # 只收本市唯一的名称，供记录里「路线编号列填了名称」时兜底（S2 的「广河高速」）。
        self.route_of_name = {}
        ambiguous = set()
        for (city, route), name in (names or {}).items():
            text = _route(name)
            key = (self._norm_city(city), text)
            if not text or key in ambiguous:
                continue
            if key in self.route_of_name and self.route_of_name[key] != _route(route):
                self.route_of_name.pop(key)
                ambiguous.add(key)          # 本市同名多行 → 不做名称兜底
            else:
                self.route_of_name[key] = _route(route)
        self.owner_index = {}    # 附件分段索引；由 load_owner_index 惰性挂载
        self.measured = {}       # 明细文件名实测段索引；由 load_measured_segments 惰性挂载

    def _remember_display(self, norm, raw):
        """记录地市原始写法；同地市出现“韶关/韶关市”两种写法时优先保留带“市”的完整形式。"""
        current = self.display.get(norm)
        if current is None:
            self.display[norm] = raw
        elif raw.endswith("市") and not current.endswith("市") and raw[:-1] == current:
            self.display[norm] = raw

    @staticmethod
    def _norm_city(value):
        """地市匹配前归一化：去空白；末尾“市”仅在三字及以上时去掉（韶关市↔韶关）。"""
        text = re.sub(r"\s+", "", str(value or ""))
        return text[:-1] if text.endswith("市") and len(text) > 2 else text

    @staticmethod
    def city_display(value):
        """显示层市名统一带“市”（模板取证：正文「共抽检XX市高速公路…」、表行标签「XX市抽检高速」「XX市合计」）。

        只用于渲染（正文/表题/表行标签/产物文件名与子目录名）；分类表「地市」列与数据里的管养单位名原样保留。
        """
        text = re.sub(r"\s+", "", str(value or ""))
        return text if not text or text.endswith("市") else f"{text}市"

    @classmethod
    def _route_key(cls, city, value):
        """记录用路线号归一入口：去空白大写 + 录入笔误勘误（ROUTE_CODE_ERRATA）。"""
        norm = cls._norm_city(city)
        code = _route(value)
        return norm, ROUTE_CODE_ERRATA.get((norm, code), code)

    @staticmethod
    def _row_category(value):
        """道路类别/道路等级单元格 → 高速公路/普通国省道。
        “高速”开头为高速公路；“国省道”等同普通国省道；一级、二级等普通公路等级亦为普通国省道。"""
        text = re.sub(r"\s+", "", str(value or ""))
        if not text:
            return None
        if text in RouteCategoryIndex.VALID:
            return text
        if "国省道" in text or "普通公路" in text:
            return "普通国省道"
        if text.startswith("高速"):
            return "高速公路"
        if re.match(r"^(一级|二级|三级|四级|等外|快速)", text):
            return "普通国省道"
        return None

    @classmethod
    def _sheet_category(cls, title):
        """按工作表名推断道路类别（附件3_4 格式无“道路类别”列）。"""
        normalized = _norm_header(title)
        if "国省道" in normalized:
            return "普通国省道"
        if "高速" in normalized:
            return "高速公路"
        return None

    @staticmethod
    def _is_total_row(values):
        """汇总行（“总计/合计/小计”）不参与路线分类。"""
        return any(str(v).strip() in {"总计", "合计", "小计"} for v in values if v not in (None, ""))

    @classmethod
    def from_file(cls, path):
        """支持四种路线表格式，均按表头字段识别，不依赖文件名：
        1) 旧格式：首行表头含 地市/地区 + 路线/路线编号/路线编码 + 道路类别/公路类别；
        2) 附件3_4 格式：工作表名区分类别，首行大标题、第2行表头；
        3) 省检线路统计格式：表头含 地市 + 路线/线路号，类别取自左侧首列填充值
           （“高速/国省道”，向下填充）或“道路等级”列；
        4) 无表头格式：整表按 地市、路线号、路线名、方向、起点、终点、里程、道路等级、管养单位
           的固定列序识别（首列地市向下填充）。
        “路线名称/线路名”只是路线名的别名，不单独充当路线号。
        无地市+路线表头的工作表（如“里程汇总”）自动跳过。
        """
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        mapping = {}
        owners = {}
        names = {}
        seen = {}
        try:
            for ws in wb.worksheets:
                sheet_category = cls._sheet_category(ws.title)
                rows = list(ws.iter_rows(values_only=True))
                header_index = None
                for index, values in enumerate(rows):
                    cells = {_norm_header(v) for v in values if v not in (None, "")}
                    if ("地市" in cells or "地区" in cells) and any(
                            name in cells for name in ("路线", "路线编号", "路线编码", "线路号")):
                        header_index = index
                        break
                if header_index is not None:
                    data_rows = cls._header_rows(rows, header_index, sheet_category)
                else:
                    data_rows = cls._positional_rows(rows, sheet_category)
                for raw_city, route, category, owner, name in data_rows:
                    if not raw_city or not route or not category:
                        continue
                    if category not in cls.VALID:
                        continue
                    key = (cls._norm_city(raw_city), _route(route))
                    if key in seen and seen[key] != category:
                        raise ValueError(f"路线分类冲突：{raw_city}/{key[1]}")
                    seen[key] = category
                    mapping[(raw_city, _route(route))] = category
                    if owner:
                        owners[(raw_city, _route(route))] = owner
                    if name:
                        names[(raw_city, _route(route))] = name
        finally:
            wb.close()
        if not mapping:
            raise ValueError("路线分类表未读取到有效记录")
        return cls(mapping, owners, names)

    @classmethod
    def _header_rows(cls, rows, header_index, sheet_category):
        """有表头表：逐行取地市/路线/类别，类别回退顺序为
        “道路类别/公路类别/道路等级”列 → 左侧首列填充值 → 工作表名。
        最左列常是无表头的块标记列（表头单元格为空，值为“高速/国省道”并向下填充）。"""
        headers = list(rows[header_index])
        head_labels = [_norm_header(h) for h in headers]
        city_col = next((i for i, name in enumerate(head_labels) if name in ("地市", "地区")), None)
        # 最左列在“地市”列左侧且自身无表头 → 块标记列（“高速/国省道”），取其填充值作为类别来源。
        lead_col = 0 if city_col is not None and city_col > 0 and not head_labels[0] else None
        carried = None
        carried_city = None
        for values in rows[header_index + 1:]:
            if cls._is_total_row(values):
                continue
            row = dict(zip(headers, values))
            if lead_col is not None:
                lead = values[lead_col] if lead_col < len(values) else None
                lead = None if _norm_header(lead) in ("", "总计", "合计") else lead
                if lead:
                    carried = lead
            # 地市为合并单元格或仅区块首行填写时，向下填充。
            cell_city = values[city_col] if (city_col is not None and city_col < len(values)) else None
            if cell_city not in (None, ""):
                carried_city = str(cell_city).strip()
            raw_city = carried_city
            route = _first(row, "路线", "路线编号", "路线编码", "线路号")
            category = _first(row, "道路类别", "公路类别", "道路等级")
            category = cls._row_category(category) or (
                cls._row_category(carried) if lead_col is not None else None) or sheet_category
            owner = _first(row, *cls.OWNER_ALIASES)
            if not raw_city or not route:
                continue
            # 「路线名称」只作别名，不单独充当路线号（见 from_file 文档字符串）
            name = _first(row, "路线名称", "线路名", default="")
            yield raw_city, route, category, str(owner).strip() if owner else "", str(name or "").strip()

    @classmethod
    def _positional_rows(cls, rows, sheet_category):
        """无表头表：固定列序 地市/路线号/路线名/方向/起点/终点/里程/道路等级/管养单位，
        首列地市向下填充；列数不足或前两列非“地市+路线号”形状时返回空。"""
        route_re = re.compile(r"^[GS]\d", re.I)
        carried = None
        for values in rows:
            if cls._is_total_row(values):
                continue
            cells = list(values) + [None] * max(0, 9 - len(values))
            city = str(cells[0]).strip() if cells[0] not in (None, "") else ""
            route = cells[1]
            if city:
                carried = city
            route_text = str(route).strip() if route not in (None, "") else ""
            if not route_re.match(route_text):
                continue
            level = cls._row_category(cells[7])
            category = level or sheet_category
            if not carried or not category:
                continue
            # 无表头格式第 3 列即路线名（见本函数文档字符串的固定列序）
            yield carried, route, category, "", str(cells[2] or "").strip() if cells[2] else ""

    def category(self, city, route):
        key = (self._norm_city(city), _route(route))
        if key not in self.mapping:
            raise KeyError(f"路线分类缺失：{str(city).strip()}/{key[1]}")
        return self.mapping[key]

    def resolve_city(self, route):
        cities = sorted({self.display.get(norm, norm) for norm, number in self.mapping if number == _route(route)})
        if not cities:
            raise KeyError(f"路线分类缺失：{_route(route)}")
        if len(cities) > 1:
            raise ValueError(f"路线{_route(route)}跨地市，无法唯一补全城市：{','.join(cities)}")
        return cities[0]

    def rows(self, city=None):
        target = self._norm_city(city) if city is not None else None
        return [{"city": self.display.get(norm, norm), "route": number, "category": value}
                for (norm, number), value in sorted(self.mapping.items()) if target is None or norm == target]

    def category_or_none(self, city, route):
        """查路线类别，查不到返回 None（不抛异常）。用于对人工对比记录等补充类别信息。"""
        key = self._route_key(city, route)
        return self.mapping.get(key)

    def owner_route_key(self, city, route):
        """清单表/明细里的路线值 → 路线索引键。源表「路线编号」列填的是路线名称时
        （广州 S2 的「广河高速」），按本市唯一的路线名称反查编码（与 category_for_records
        第 2 级同口径：route_of_name 只收本市唯一的名称，多义名不反查，不做模糊匹配）。
        """
        key = self._route_key(city, route)
        if key in self.mapping or key in self.owners or key in (self.owner_index or {}):
            return key
        by_name = self.route_of_name.get(key)
        return (key[0], by_name) if by_name else key

    def _owner_of_slice(self, segs, a, b, keepers):
        """桩号片段 [a,b]（km）的经营主体：A1 §6.1 的 P2 区间重叠，无覆盖时 P3 管养单位。

        P2 命中多个主体 → None（判不出，禁止取多数/取第一段）；P2 无覆盖且片段内管养单位
        在附件里唯一对应一个主体 → 该主体；否则 None，由调用方如实留空并计入未命中。
        """
        owners = {s["owner"] for s in segs
                  if s["owner"] and s["start"] is not None
                  and min(s["end"], b) - max(s["start"], a) > 0}
        if len(owners) == 1:
            return owners.pop()
        if owners:
            return None
        names = {k for lo, hi, k in keepers if k and lo < b and hi > a}
        if len(names) != 1:
            return None
        keeper = next(iter(names))
        # P1i：管养单位按分支标识比较，不按整串等值 —— 同一分支的公司名在附件与明细里
        # 写法不同（深圳 G0422 博深分公司），整串等值会让附件未覆盖的空洞判不出主体。
        found = {s["owner"] for s in segs
                 if s["owner"] and gd_branch_key(s["keeper"]) == gd_branch_key(keeper)}
        return found.pop() if len(found) == 1 else None

    def owner_spans(self, city, route, start, end, keepers=()):
        """[起,止]（km）按附件分段切成 [(起, 止, 经营主体)]，各段互不重叠、拼回原区间。

        跨主体路线（G0425 广州、G4 广州）的一行抽检里程本来会整段落空，导致①表
        「非集团+集团 < 全市」；这里按 P2 桩号区间把该行拆到两侧，差额才可能归零。
        `keepers` = [(桩号km, 桩号km, 管养单位)]，只用于附件覆盖不到的尾段（P3）。
        判不出的片段不返回（调用方如实报未命中，不摊派）。
        """
        segs = (self.owner_index or {}).get(self.owner_route_key(city, route)) or []
        lo, hi = _float(start), _float(end)
        if not segs or lo is None or hi is None:
            return []
        lo, hi = min(lo, hi), max(lo, hi)
        cuts = sorted({lo, hi} | {c for s in segs if s["start"] is not None
                                  for c in (s["start"], s["end"]) if lo < c < hi})
        out = []
        for a, b in zip(cuts, cuts[1:]):
            owner = self._owner_of_slice(segs, a, b, keepers)
            if owner:
                out.append((a, b, owner))
        return out

    def owner_for(self, city, route, start=None, end=None, keeper=None):
        """经营主体：表内列优先；表里没有该列时回退附件匹配（match_owner）。

        两者都没有 → None（调用方写空单元格，禁止写「—」等占位，避免被当主体值参与分组）。
        """
        key = self.owner_route_key(city, route)
        found = self.owners.get(key)
        # P1b：表里的 `跨主体` 只是「这条路线本市跨两个主体」的标记，不是可用的归属值 ——
        # 放行到下面的附件 + 桩号消歧（P2→P3→P4）去定；定不出才回落到跨主体。
        if found and found != CROSS_OWNER:
            return found
        if not self.owner_index:
            return found or None      # 没有附件可消歧：`跨主体` 标记原样回传，不降级成「缺失」
        owner = match_owner(self.owner_index, city, route, start, end, keeper)
        if owner is not None or start is not None or end is not None:
            return owner
        # 分类表没有桩号列：附件内跨主体时，用本市实测抽检段（明细文件名 K 区间）消歧。
        # 实测段全部落入同一主体 → 该主体；跨两个主体 → 「跨主体」；判不出 → None。
        spans = (self.measured or {}).get(key) or []
        values = {value for lo, hi in spans
                  if (value := match_owner(self.owner_index, city, route, lo, hi, keeper))}
        if len(values) == 1:
            return values.pop()
        return CROSS_OWNER if len(values) > 1 else None

    def owner_conflicts(self):
        """附件里 (地市,路线) 跨多个经营主体、且本次无桩号可消歧的清单（供导出时如实上报）。"""
        if not self.owner_index:
            return []
        out = []
        for (city, code) in sorted(self.owner_index):
            if self.owners.get((city, code)):
                continue
            if len(self.owner_index[(city, code)]) > 1:
                out.append((self.display.get(city, city), code))
        return out

    def category_for_records(self, city, route):
        """R2 记录用道路类别兜底（用户口径：记录按数据所在文件夹定市）。

        文件夹定市后，文件夹内可能混有“地市”列写其他市的记录，其路线可能不在本市路线表里；
        这里绝不允许抛异常中断整市报告，兜底顺序为：
        1) (本市, 路线号) 命中 → 直接用；
        2) 路线号查不到但命中本市某行的「路线名称」→ 用该行类别（源表把编号列填成名称，
           如广州 S2 的「广河高速」；本市同名多行时不兜底）；
        3) 全表按路线号唯一命中（贯通路线）→ **仅当该路线号就是本市的行时**沿用该类别。
           P1d-b：第 2 级原先只看路线号不看地市，江门明细里 12 个他市省道码（S272/S278/S282/
           S283/S291/S292/S296/S355/S356/S368/S373/S380）借此拿到「普通国省道」类别进入江门
           ①②分母，抽检里程虚增 34km(+58.6%)、高度合格率被拉低 0.13pp（R3 §4）。现在归属他市的
           路线号一律拒绝借用，与无法判定的行走同一出口（记 issues、类别留空、不进分母）。
        4) 仍无法判定 → 返回 None，由调用方把该行记入 issues（记录仍归属该文件夹的市）。
        """
        key = self._route_key(city, route)
        if key in self.mapping:
            return self.mapping[key]
        by_name = self.route_of_name.get(key)
        if by_name and (key[0], by_name) in self.mapping:
            return self.mapping[(key[0], by_name)]
        hits = {(owner_city, value) for (owner_city, number), value in self.mapping.items()
                if number == key[1]}
        if len(hits) == 1 and next(iter(hits))[0] == key[0]:
            return next(iter(hits))[1]
        return None

    def row_category(self, row):
        """按记录中的 地市/地区 + 路线 字段查类别，查不到返回 None。"""
        city = _first(row, "地市", "地区")
        route = _first(row, "路线", "路线编号", "路线编码")
        if city is None or route is None:
            return None
        return self.category_or_none(city, route)


class GuangdongInputScanner:
    """递归扫描项目资料，先按表头判型，再补全已判型标线 CSV 的文件名元数据。"""
    EXCLUDE_DIRS = ("原始数据",)  # 原始采集数据体量巨大且非统计表，默认跳过，避免误选根目录时卡死
    MAX_FILE_BYTES = 50 * 1024 * 1024  # 单文件上限，超大的直接跳过并记入问题清单，作为二次防护

    def __init__(self, root, route_index=None, default_city=None, log=lambda _x: None):
        self.root = Path(root)
        self.route_index = route_index
        self.log = log
        # R2（用户口径）：扫描根（数据所在文件夹）即权威市，文件夹内记录的“地市”列不再决定归属。
        self.default_city = default_city or self._city_from_folder()
        self.issues = []
        if self.default_city:
            log(f"  扫描根归属市：{self.default_city}（{self.root.name}）")
        else:
            # 文件夹名解析不出市（如把省的汇总目录当根）：保持旧口径，按地市列/文件内多数市判定。
            log(f"  未识别到归属市，按地市列口径：{self.root.name}")

    def _city_from_folder(self):
        """地市为空的行默认归属其所在文件夹对应的市（如“东莞标线数据明细”“东莞-最终提交9.14”）；
        用于跨地市贯通路线（G15/G105等）无法单凭路线号唯一补全城市时的回退。
        传入单个文件（标线统计表）时向上查其所在文件夹，否则表名不含市名会丢失全部空地市行。"""
        if not self.route_index:
            return None
        names = [self.root.name]
        if self.root.is_file():
            names += [parent.name for parent in list(self.root.parents)[:3]]
        for name in names:
            for (norm, _route_num) in self.route_index.mapping:
                if norm and norm in name:
                    return self.route_index.display.get(norm, norm)
        return None

    @staticmethod
    def _kind(headers):
        h = {_norm_header(x) for x in headers if x is not None}
        has = lambda prefix: any(value.startswith(_norm_header(prefix)) for value in h)
        if any("逆反亮度系数" in value or "逆反射亮度系数" in value for value in h) and (has("计算区间") or has("桩号")):
            return "marking"
        if has("拼接螺栓数量") and has("连接螺栓缺失数量"):
            return "bolt"
        if (has("梁板中心高度") or has("护栏中心高度")) and has("护栏类型"):
            return "height"
        return None

    def _xlsx_tables(self, path):
        dimension_a1 = False
        try:
            with ZipFile(path) as archive:
                xml = archive.read("xl/worksheets/sheet1.xml")
            dimension_a1 = b'<dimension ref="A1"' in xml
        except Exception:
            xml = b""
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for index, ws in enumerate(wb.worksheets):
            if index == 0 and dimension_a1:
                continue
            values = list(ws.iter_rows(values_only=True))
            if values:
                yield ws.title, values[0], values[1:]
        wb.close()
        # 真实护栏文件 dimension 可能错误标为 A1；复用基线 XML 流读取回退。
        if dimension_a1:
            try:
                rows = list(iter_height_rows(path))
                if rows:
                    headers = list(rows[0])
                    yield "sheet1.xml回退", headers, [[row.get(h) for h in headers] for row in rows]
            except Exception as exc:
                self.issues.append(f"{path.name}: XML读取回退失败：{exc}")

    @staticmethod
    def _csv_table(path):
        last = None
        for encoding in ("utf-8-sig", "gbk", "utf-8"):
            try:
                with path.open("r", encoding=encoding, newline="") as stream:
                    rows = list(__import__("csv").reader(stream))
                return rows[0], rows[1:]
            except UnicodeDecodeError as exc:
                last = exc
        raise last

    def _metadata_from_marking_name(self, path):
        match = re.search(r"([GSHXY]\d+)\s*(上行|下行)", path.stem, re.I)
        if not match:
            return {}
        route, direction = match.groups()
        position_match = re.search(r"标线\d+", path.stem)
        position = position_match.group(0) if position_match else ""
        city = None
        km = re.search(r"K(\d+(?:\.\d+)?)K(\d+(?:\.\d+)?)", path.stem, re.I)
        segment = f"{_route(route)}{direction}K{km.group(1)}-K{km.group(2)}" if km else ""
        return {"city": city, "route": _route(route), "direction": direction,
                "marking_position": position, "file_segment": segment}

    def _convert(self, kind, row, source, sheet, source_row=None, issues=None):
        city = str(_first(row, "地市", "地区", default="") or "").strip()
        route = _route(_first(row, "路线编号", "路线"))
        direction = str(_first(row, "方向", "检测方向", default="") or "").strip()
        metadata = self._metadata_from_marking_name(source)
        if not city: city = str(metadata.get("city") or "")
        if not route: route = metadata.get("route", "")
        if not direction: direction = metadata.get("direction", "")
        # R2（用户口径）：数据所在文件夹即权威市——文件夹内记录即使“地市”列写的是别的市，
        # 也一律计入该文件夹对应的市；“地市”列不再决定归属，也不再因此丢弃。
        # 已知副作用：同一贯通路线在多个市文件夹中各有一份时，两市报告各自计入该路段，不再跨市去重
        # （原去重理由见 filter_own_city_marking 注释）；用户明确要求以文件夹定市并接受该口径。
        declared_city = ""
        if self.default_city:
            declared_city = city
            city = self.default_city
        if not city and route and self.route_index:
            try:
                city = self.route_index.resolve_city(route)
            except ValueError:
                # 跨地市贯通路线（G15/G105等）：回退到项目文件夹隐含的地市
                if self.default_city:
                    city = self.default_city
                else:
                    raise
        if kind in ("height", "bolt"):
            # P1b P0：同一行有 3 个桩号列，优先列可能回绕(9999+988)或串到别段(K2159)，
            # 由 _pick_station 在「文件名声明的 K 区间内」挑列；区间外全部列都不在时保持原优先级。
            # 模板「五、」规定护栏横梁中心以 4 m 为输出间距：高度以 4 m 输出栅格（原始桩号）为计数单位。
            # 电子修正桩号会把不同的 4 m 输出点并到同一桩号（源文件实测 209/10028 行同桩号），
            # 使输出点数与两波合格率系统性偏低（实测 14210 点/45.90% vs 17029 点/37.45%）。
            station_raw = _pick_station(row, kind, source)
        else:
            station_raw = _first(row, "标注修正桩号", "电子修正桩号", "原始桩号", "桩号", "桩号范围", "计算区间")
        explicit_segment = _first(row, "检测区段", "统计区段", "区段", "桩号范围")
        base = {"city": city, "route": route, "direction": direction, "station_m": _station_first(station_raw),
                "segment": str(explicit_segment or metadata.get("file_segment") or station_raw or ""), "manager": str(_first(row, "管养单位", default="") or ""), "interval": str(_first(row, "计算区间", default="") or ""), "source": str(source), "sheet": sheet, "source_row": source_row, "source_range": str(station_raw or "")}
        if declared_city and RouteCategoryIndex._norm_city(declared_city) != RouteCategoryIndex._norm_city(base["city"]):
            base["declared_city"] = declared_city   # 仅用于“保留其他市标记记录 N 条”日志，不参与归属
        location = f"{Path(source).name}/{sheet}/第{source_row}行"
        if not base["city"] or not base["route"]:
            raise ValueError(f"{location}：缺少地市或路线元数据")
        if _float(base["station_m"]) is None:
            raise ValueError(f"{location}：桩号无法解析：{station_raw}")
        if kind in ("height", "bolt"):
            bounds = _filename_k_range(source)
            if bounds and not (bounds[0] <= float(base["station_m"]) <= bounds[1]):
                return None
        if kind == "height":
            value = _float(_first(row, "梁板中心高度(mm)", "护栏中心高度(mm)", "梁板中心高度", "护栏中心高度", "梁板中心高度毫米", "护栏中心高度mm"))
            remark = str(_first(row, "异常标记", "备注标记", "备注", default="") or "").strip()
            if value is None or value <= 0:
                # 桥梁/隧道路段标记行：不参与高度统计，但记录供报告生成区段说明
                if remark and ("桥梁" in remark or "隧道" in remark):
                    base.update(guardrail_note=remark)
                    return base
                return None
            base.update(guardrail_type=_guardrail_type(_first(row, "护栏类型")), height=value)
            if remark and ("桥梁" in remark or "隧道" in remark):
                base["guardrail_note"] = remark
        elif kind == "bolt":
            remark = str(_first(row, "备注标记", "异常标记", "备注", default="") or "").strip()
            vals = [_float(_first(row, name, name+"颗")) or 0.0 for name in ("拼接螺栓数量","拼接螺栓缺失数量","连接螺栓数量","连接螺栓缺失数量")]
            outline = _float(_first(row, "轮廓标数量（个）", "轮廓标数量(个)", "轮廓标数量")) or 0.0
            base.update(splice=vals[0], splice_missing=vals[1], connection=vals[2], connection_missing=vals[3], outline=outline)
            if remark and ("桥梁" in remark or "隧道" in remark):
                base["guardrail_note"] = remark
        else:
            side_columns = (("主车道左侧标线逆反亮度系数", "左侧标线", "左侧标线逆反射目标值"), ("主车道右侧标线逆反亮度系数", "右侧标线", "右侧标线逆反射目标值"))
            wide = any(_norm_header(column) in {_norm_header(k) for k in row} for column, _, _ in side_columns)
            if not direction or (wide and not base["manager"].strip()):
                raise ValueError(f"{location}：缺少检测方向或管养单位元数据")
            if wide:
                endpoints = re.fullmatch(r"\s*(K?\s*\d+(?:\.\d+)?\s*\+\s*\d+(?:\.\d+)?)\s*[-~～—–－]\s*(K?\s*\d+(?:\.\d+)?\s*\+\s*\d+(?:\.\d+)?)\s*", str(station_raw or ""), re.I)
                if not endpoints:
                    raise ValueError(f"{location}：桩号区间无法解析：{station_raw}")
                base["station_m"], base["end_m"] = map(_station_first, endpoints.groups())
                if any(_float(base[k]) is None for k in ("station_m", "end_m")):
                    raise ValueError(f"{location}：桩号区间必须为有限数值")
                sides = [(_first(row, column), position, _first(row, target)) for column, position, target in side_columns]
            else:
                sides = [(_first(row, "逆反亮度系数", "逆反射亮度系数"),
                          str(_first(row, "标线位置", "标线名称", default=metadata.get("marking_position", ""))),
                          _first(row, "逆反亮度系数目标值", "逆反射亮度系数目标值", default=80))]
            records = []
            for raw_value, position, raw_target in sides:
                if raw_value is None or str(raw_value).strip() == "":
                    continue  # 空侧不补值，也不丢弃另一有效侧。
                value, target = _float(raw_value), _float(raw_target)
                if value is None or value < 0 or target is None or target < 0:
                    message = f"{location}/{position}：标线值及目标值必须为非负有限数值（值={raw_value}，目标={raw_target}）"
                    if issues is None: raise ValueError(message)
                    issues.append(message)
                    continue
                records.append(dict(base, marking_position=position, value=value, target=target))
            return records if wide else (records[0] if records else None)
        return base

    def _process_file(self, path):
        """单个文件处理（并行任务单元）"""
        records = {"height": [], "bolt": [], "marking": [], "notes": [], "issues": []}
        try:
            tables = self._xlsx_tables(path) if path.suffix.lower() == ".xlsx" else [(str(path.stem), *self._csv_table(path))]
            for sheet, headers, rows in tables:
                kind = self._kind(headers)
                if not kind: continue
                carried_manager = {}   # (路线, 方向) → 上一行的管养单位
                for source_row, values in enumerate(rows, 2):
                    if not any(value not in (None, "") for value in values): continue
                    row = dict(zip(headers, values))
                    # 源表按块留空：同一路线方向内 管养单位 常只写首行，后续行沿用，否则整段被判为缺元数据而丢失
                    manager_key = next((k for k in row if k and _norm_header(str(k)) == _norm_header("管养单位")), None)
                    if manager_key:
                        scope = (str(_first(row, "路线编号", "路线", default="") or ""), str(_first(row, "方向", "检测方向", default="") or ""))
                        current = str(row.get(manager_key) or "").strip()
                        if current:
                            carried_manager[scope] = current
                        elif carried_manager.get(scope):
                            row[manager_key] = carried_manager[scope]
                    try:
                        record = self._convert(kind, row, path, sheet, source_row, records["issues"])
                        if isinstance(record, list):
                            for item in record:
                                if item: records[kind].append(item)
                        elif record:
                            if record.get("guardrail_note"):
                                records["notes"].append(record)
                            if (record.get("height") is not None or record.get("splice") is not None
                                    or not record.get("guardrail_note")):
                                records[kind].append(record)
                    except Exception as exc:
                        records["issues"].append(str(exc))
        except Exception as exc:
            records["issues"].append(f"{path.name}：读取失败：{exc}")
        return records

    def scan(self):
        """并行扫描项目资料，优化大数据量性能。root 可为文件夹，也可为单个文件（标线直接指向“标线统计”表）。"""
        base = self.root.parent if self.root.is_file() else self.root
        if not base.is_dir(): raise FileNotFoundError(f"项目资料文件夹不存在：{base}")
        if self.root.is_file():
            files = [self.root]
        else:
            files = [p for p in sorted(base.rglob("*")) if p.is_file() and not p.name.startswith("~$") and p.suffix.lower() in (".xlsx", ".csv")]
        # 跳过显式排除的文件夹（如“原始数据”），避免误选根目录时读取海量原始采集文件而卡死
        kept, skipped_dir = [], []
        for p in files:
            parts = p.relative_to(base).parts
            (skipped_dir if any(part in self.EXCLUDE_DIRS for part in parts) else kept).append(p)
        files = kept
        # 单文件体积上限，超大的直接跳过并记入问题清单，作为二次防护
        huge = [(p, p.stat().st_size) for p in files if p.stat().st_size > self.MAX_FILE_BYTES]
        files = [p for p in files if p.stat().st_size <= self.MAX_FILE_BYTES]
        for p, size in huge:
            self.issues.append(f"{p.name}：文件过大（约{size/1048576:.0f}MB）已跳过，疑似非统计表原始数据")
        if skipped_dir or huge:
            self.log(f"待识别文件 {len(files)} 个（已跳过 {len(skipped_dir)} 个排除文件夹内文件、{len(huge)} 个超大文件）")
        else:
            self.log(f"待识别文件 {len(files)} 个")
        result = {"height": [], "bolt": [], "marking": [], "notes": [], "issues": self.issues}
        # ponytail: 并行处理，max_workers=4 避免内存溢出，必要时可调高
        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = {executor.submit(self._process_file, f): f for f in files}
            done = 0
            for future in as_completed(futures):
                partial = future.result()
                done += 1
                for kind in ("height", "bolt", "marking", "notes"):
                    result[kind].extend(partial[kind])
                result["issues"].extend(partial["issues"])
                if done % 10 == 0 or done == len(files):
                    self.log(format_progress("扫描资料", done, len(files), futures[future].name))
        return result



def detect_guangdong_data_folders(project_dir):
    """扫描项目文件夹，自动识别标线数据和护栏数据所在的子文件夹。
    返回 (marking_dir, guardrail_dir)，可能为 None。
    """
    project_dir = Path(project_dir)
    marking_candidates = {}
    guardrail_candidates = {}

    dirs_to_check = [project_dir] + sorted(
        d for d in project_dir.rglob("*")
        if d.is_dir() and not d.name.startswith(".")
        and d.name not in GuangdongInputScanner.EXCLUDE_DIRS
        and len(d.relative_to(project_dir).parts) <= 3
    )

    for subdir in dirs_to_check:
        marking_count = 0
        guardrail_count = 0
        files = sorted(
            f for f in subdir.iterdir()
            if f.is_file() and not f.name.startswith("~$") and f.suffix.lower() in (".xlsx", ".csv")
        )
        for f in files[:20]:  # 每个目录最多检查20个文件
            try:
                if f.stat().st_size > 50 * 1024 * 1024:
                    continue
            except OSError:
                continue
            try:
                if f.suffix.lower() == ".xlsx":
                    wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
                    for ws in wb.worksheets[:2]:
                        headers = [str(v) for v in next(ws.iter_rows(min_row=1, max_row=1, values_only=True), []) if v is not None]
                        kind = GuangdongInputScanner._kind(headers)
                        if kind == "marking":
                            marking_count += 1
                        elif kind in ("height", "bolt"):
                            guardrail_count += 1
                    wb.close()
                elif f.suffix.lower() == ".csv":
                    hdrs = None
                    for encoding in ("utf-8-sig", "gbk", "utf-8"):
                        try:
                            with open(f, "r", encoding=encoding) as fh:
                                hdrs = next(csv.reader(fh), [])
                            break
                        except UnicodeDecodeError:
                            continue
                    if hdrs is None:
                        continue
                    kind = GuangdongInputScanner._kind(hdrs)
                    if kind == "marking":
                        marking_count += 1
                    elif kind in ("height", "bolt"):
                        guardrail_count += 1
            except Exception:
                continue
        if marking_count:
            marking_candidates[str(subdir)] = marking_count
        if guardrail_count:
            guardrail_candidates[str(subdir)] = guardrail_count

    def _best(candidates):
        if not candidates:
            return None
        return max(candidates, key=lambda path: (
            candidates[path], len(Path(path).relative_to(project_dir).parts), path))

    marking_dir = _best(marking_candidates)
    guardrail_dir = _best(guardrail_candidates)
    return marking_dir, guardrail_dir


# 导入源命名约定：标线以“标线统计”表为唯一数据源（“标线区间统计”与“单路线明细”均为其派生物，
# 同目录读取会重复计数），护栏高度与螺栓取“护栏数据明细”文件夹内的全部文件。
GUANGDONG_MARKING_TABLE = "标线统计"
GUANGDONG_GUARDRAIL_FOLDER = "护栏数据明细"


def _is_guardrail_dir(name):
    """护栏明细目录实测有「护栏数据明细/护栏明细数据/护栏明细/2. 护栏高度、螺栓数据」等形态（字序不定、
    可带序号前缀、可嵌在「云浮明细」等父目录下——层级由调用方 rglob 逐层递归），统一按「护栏」+（「明细」或「数据」）判定。"""
    return "护栏" in name and ("明细" in name or "数据" in name)


def detect_guangdong_sources(project_dir):
    """识别导入源，返回 (标线源列表, 护栏源列表)。

    按市存放时每个市各命中一份，因此返回多个源；不要求市名与文件夹名一致，
    记录的市别取自表内数据，故“广州/广东护栏数据明细”“中山统计标线数据”这类
    命名差异无需特判。命名全部未命中时回退到按表头识别目录（兼容旧格式）。
    """
    root = Path(project_dir)
    if not root.is_dir():
        raise FileNotFoundError(f"项目资料文件夹不存在：{root}")

    def _is_marking_table(name):
        return (not name.startswith("~$") and name.endswith((".xlsx", ".csv"))
                and GUANGDONG_MARKING_TABLE in name and "区间" not in name)

    marking = sorted(
        f for f in root.rglob("*")
        if f.is_file() and _is_marking_table(f.name)
        and f.stat().st_size <= GuangdongInputScanner.MAX_FILE_BYTES
    )
    guardrail = sorted(
        d for d in root.rglob("*")
        if d.is_dir() and _is_guardrail_dir(d.name)
    )
    # 命中按类独立：标线命中不得短路护栏的表头回退（旧版 `if marking or guardrail` 会在
    # “有标线、护栏目录名未命中”时整块跳过回退 → 护栏高度/螺栓静默为 0）。
    if not marking or not guardrail:
        # ponytail: 回退结果是单目录（旧版行为）；命名约定覆盖当前全部输入，仅补漏不改变已命中项
        fallback_marking, fallback_guardrail = detect_guangdong_data_folders(root)
        marking = marking or ([Path(fallback_marking)] if fallback_marking else [])
        guardrail = guardrail or ([Path(fallback_guardrail)] if fallback_guardrail else [])
    return marking, guardrail


def folder_city_for(path, route_index):
    """R2：按数据所在文件夹（含文件名）判定归属市，复用扫描器的文件夹定市口径；识别不到返回 None。

    人工复核对比表与其所在目录同属该市（如「…/东莞-最终提交9.14/进场后人工复核数据对比/
    人工自动化对比东莞.xlsx」），因此两期对比表都用同一口径定市，不再按表内“地市”列丢行。
    """
    if not path or not route_index:
        return None
    return GuangdongInputScanner(path, route_index).default_city


def filter_own_city_marking(rows, folder_city=None):
    """标线归属过滤。

    folder_city（扫描根文件夹对应的市，R2 用户口径）已知时，文件夹内全部记录一律计入该市，
    “地市”列不再决定归属、也不丢弃任何记录。
    归属市未知（folder_city=None，如把省的汇总目录当根）时保持旧口径：按文件内多数地市判定
    该文件归属，返回 (归属市, 本市记录)。

    旧口径的去重理由（仅在归属市未知时仍生效）：跨市路段会同时列在多个市的文件中且实测值不同
    （如博深高速 G0422 的同一桩号在东莞与深圳文件中分别为 551.55 / 440.51），只取本市那份，
    避免重复计数与数值冲突。
    """
    if not rows:
        return folder_city, rows
    if folder_city:
        # R2（用户口径）：数据所在文件夹即权威市。同一贯通路线在多市文件夹中各有一份时，
        # 两市报告各自计入该路段，不再跨市去重——这是用户明确要求的口径。
        return folder_city, rows
    own = Counter(RouteCategoryIndex._norm_city(r.get("city")) for r in rows).most_common(1)[0][0]
    return own, [r for r in rows if RouteCategoryIndex._norm_city(r.get("city")) == own]


def own_city_marking(rows, log=lambda _m: None, folder_city=None):
    """标线入账前统一过滤：文件夹市已知时全部保留（R2 用户口径，仅记录其他市标记条数）；
    归属市未知时按文件内多数地市取数并记录忽略条数。"""
    own, kept = filter_own_city_marking(rows, folder_city)
    if folder_city:
        other = sum(1 for row in rows if row.get("declared_city"))
        log(f"  按文件夹归属{own}取数，保留其他市标记记录 {other} 条")
    elif own and len(kept) != len(rows):
        log(f"  按归属{own}取数：忽略其他市的重复路段 {len(rows) - len(kept)} 条")
    return kept


MARKING_SEGMENT_FIELDS = ("city", "category", "route", "direction", "manager", "segment")


class GuangdongStatistics:
    HEIGHT_LIMITS = HEIGHT_LIMITS

    @classmethod
    def group_marking_ranges(cls, rows):
        """只归组宽表的相接/重叠输入区间；源方向、行号、两端和旧格式区段不变。"""
        groups = cls._group([r for r in rows if "end_m" in r], ("city", "route", "direction", "manager"))
        for selected in groups.values():
            connected = []
            start = end = None
            def flush():
                if connected:
                    segment = f"{format_station_one_decimal(start)}-{format_station_one_decimal(end)}"
                    for row in connected: row["segment"] = segment
            for row in sorted(selected, key=lambda r: min(r["station_m"], r["end_m"])):
                low, high = sorted((row["station_m"], row["end_m"]))
                if end is None or low > end:
                    flush()
                    connected = []
                    start = low
                    end = high
                else:
                    end = max(end, high)
                connected.append(row)
            flush()

    @classmethod
    def height_summary(cls, rows):
        result = {}
        for kind in ("二波", "三波"):
            selected = [r for r in rows if _guardrail_type(r.get("guardrail_type")) == kind and _float(r.get("height")) is not None]
            low, high = cls.HEIGHT_LIMITS[kind]
            values = [float(r["height"]) for r in selected]
            qualified = sum(low <= v <= high for v in values)
            over = sum(height_deviation_over_10cm(kind, v) for v in values)
            result[kind] = {"valid_count": len(values), "average": sum(values)/len(values) if values else None,
                            "qualified_count": qualified, "qualified_rate": qualified/len(values) if values else None,
                            "over_10cm_count": over}
        return result

    @staticmethod
    def bolt_summary(rows):
        totals = {k: sum(float(r.get(k, 0) or 0) for r in rows) for k in ("splice","splice_missing","connection","connection_missing")}
        missing = totals["splice_missing"] + totals["connection_missing"]
        totals.update(missing_total=missing,
                      missing_rate=bolt_missing_ratio(totals["splice"], totals["connection"], missing))
        return totals

    @staticmethod
    def marking_summary(rows):
        values=[float(r["value"]) for r in rows if _float(r.get("value")) is not None]
        qualified=sum(float(r["value"]) >= float(r.get("target",80)) for r in rows if _float(r.get("value")) is not None)
        return {"valid_count":len(values),"average":sum(values)/len(values) if values else None,
                "qualified_count":qualified,"qualified_rate":qualified/len(values) if values else None}

    @staticmethod
    def _group(rows, fields):
        grouped = {}
        for row in rows:
            key = tuple(row.get(field, "") for field in fields)
            grouped.setdefault(key, []).append(row)
        return grouped

    @classmethod
    def marking_segment_summary(cls, rows):
        fields = (*MARKING_SEGMENT_FIELDS, "marking_position")
        result = []
        for key, selected in cls._group(rows, fields).items():
            summary = cls.marking_summary(selected)
            result.append(dict(zip(fields, key), **summary))
        return sorted(result, key=lambda row: tuple(str(row.get(field, "")) for field in fields))

    @classmethod
    def marking_segment_pair_summary(cls, rows):
        fields = MARKING_SEGMENT_FIELDS
        result = []
        for key, selected in cls._group(rows, fields).items():
            item = dict(zip(fields, key))
            side_names = marking_side_names(r.get("marking_position") for r in selected)
            for pos in side_names:
                summary = cls.marking_summary([r for r in selected if str(r.get("marking_position", "")) == pos])
                item[f"{pos}_valid_count"] = summary["valid_count"]
                item[f"{pos}_average"] = summary["average"]
                item[f"{pos}_qualified_count"] = summary["qualified_count"]
                item[f"{pos}_qualified_rate"] = summary["qualified_rate"]
            item["_side_names"] = side_names
            result.append(item)
        return sorted(result, key=lambda row: tuple(str(row.get(field, "")) for field in fields))

    @classmethod
    def height_segment_summary(cls, rows):
        fields = ("category", "route", "direction", "segment", "guardrail_type")
        result = []
        for key, selected in cls._group(rows, fields).items():
            kind = _guardrail_type(key[-1])
            if kind not in cls.HEIGHT_LIMITS:
                continue
            summary = cls.height_summary(selected).get(kind, {})
            result.append(dict(zip(fields, key), **summary))
        return sorted(result, key=lambda row: tuple(str(row.get(field, "")) for field in fields))

    @classmethod
    def bolt_segment_summary(cls, rows):
        fields = ("category", "route", "direction", "segment")
        result = []
        for key, selected in cls._group(rows, fields).items():
            result.append(dict(zip(fields, key), **cls.bolt_summary(selected)))
        return sorted(result, key=lambda row: tuple(str(row.get(field, "")) for field in fields))

    @staticmethod
    def continuous_marking_weak(rows, minimum_length_m=3000):
        groups = {}
        for row in rows:
            key=(row.get("route"),row.get("direction"),row.get("marking_position"))
            groups.setdefault(key,[]).append(row)
        weak=[]
        for key, selected in groups.items():
            selected=sorted(selected,key=lambda r:r.get("station_m") if r.get("station_m") is not None else float("inf"))
            low=high=previous=None
            def flush():
                # 止点取最后一条不合格单元的 end_m 高端（旧实现取起点，系统性少一个单元长度，且恰好达阈值的区间被漏报）。
                if low is not None and high is not None and high-low >= minimum_length_m:
                    weak.append({"route":key[0],"direction":key[1],"marking_position":key[2],"start_m":low,"end_m":high})
            for row in selected:
                station=row.get("station_m"); value=_float(row.get("value")); target=_float(row.get("target")) or 80
                bad=station is not None and value is not None and value < target
                if not bad:
                    flush(); low=high=previous=None; continue
                end=_float(row.get("end_m"))
                cell_low, cell_high=sorted((station, end if end is not None else station))
                if previous is None:
                    low, high, previous = cell_low, cell_high, station
                elif abs(station-previous) <= 50:
                    low, high, previous = min(low, cell_low), max(high, cell_high), station
                else:
                    flush(); low, high, previous = cell_low, cell_high, station
            flush()
        return weak


    # ==================== GD03 四段式聚合（brief-gd03 §3） ====================

    MARKING_UNIT_RECORDS = 5

    @staticmethod
    def _route_key(route):
        text = str(route or "")
        prefix = text.rstrip("0123456789")
        digits = text[len(prefix):]
        return (prefix, int(digits) if digits.isdigit() else 0, text)

    @staticmethod
    def _span_km(units):
        """同一路线跨区段取并集后合计里程（km）。"""
        spans = {}
        for unit in units:
            start, end = unit.get("start_m"), unit.get("end_m")
            if start is None or end is None:
                continue
            key = str(unit.get("route") or "")
            low, high = spans.get(key, (start, end))
            spans[key] = (min(low, start), max(high, end))
        return sum(high - low for low, high in spans.values()) / 1000

    @staticmethod
    def _side_of(position):
        return marking_side_names([position]).get(str(position), str(position))

    @classmethod
    def marking_units(cls, rows):
        """100 m 平均值基准：每侧每 5 条 20 m 记录（按桩号排序）归为一个计算单元。

        分组口径与附件一致：路线+方向+管养单位+标线位置（不按检测区段拆分，
        否则跨区段处会把 5 条记录截断，导致合格率偏差）。
        """
        fields = ("route", "direction", "manager", "marking_position")
        units = []
        for key, selected in cls._group(rows, fields).items():
            ordered = sorted((r for r in selected if _float(r.get("value")) is not None and r.get("station_m") is not None),
                             key=lambda r: r["station_m"])
            for index in range(0, len(ordered), cls.MARKING_UNIT_RECORDS):
                chunk = ordered[index:index + cls.MARKING_UNIT_RECORDS]
                values = [float(r["value"]) for r in chunk]
                targets = [_float(r.get("target")) for r in chunk]
                target = max((t for t in targets if t is not None), default=80.0)
                average = sum(values) / len(values)
                ends = [_float(r.get("end_m")) for r in chunk]
                ends = [e for e in ends if e is not None]
                units.append(dict(zip(fields, key), count=len(values), average=average, target=target,
                                  qualified=average >= target, start_m=min(r["station_m"] for r in chunk),
                                  end_m=max(ends) if ends else max(r["station_m"] for r in chunk)))
        return units

    @classmethod
    def marking_unit_rates(cls, units):
        """按标线侧汇总 100 m 单元的合格率。"""
        result = {}
        for unit in units:
            side = cls._side_of(unit.get("marking_position"))
            item = result.setdefault(side, {"unit_count": 0, "qualified_count": 0})
            item["unit_count"] += 1
            item["qualified_count"] += 1 if unit.get("qualified") else 0
        for item in result.values():
            item["qualified_rate"] = item["qualified_count"] / item["unit_count"] if item["unit_count"] else None
        return result

    @classmethod
    @staticmethod
    def _marking_km(units):
        """检测里程（km）：每侧 100 m 计算单元数 × 0.1（oracle 口径，跨区段按单元累计）。"""
        sides = {}
        for unit in units:
            side = GuangdongStatistics._side_of(unit.get("marking_position"))
            sides[side] = sides.get(side, 0) + 1
        return round(0.1 * max(sides.values()), 1) if sides else 0.0

    @classmethod
    def marking_overall(cls, units):
        rates = cls.marking_unit_rates(units)
        left = rates.get("左侧标线", {}).get("qualified_rate")
        right = rates.get("右侧标线", {}).get("qualified_rate")
        return {"left_rate": left, "right_rate": right, "overall_rate": _gd03_overall(left, right),
                "unit_count": len(units),
                "qualified_count": sum(1 for unit in units if unit.get("qualified")),
                "left_units": rates.get("左侧标线", {}).get("unit_count", 0),
                "right_units": rates.get("右侧标线", {}).get("unit_count", 0)}

    @classmethod
    def marking_route_units(cls, rows):
        """②“路线—管养单位”合并（跨检测区段）的 100 m 单元合格率。"""
        units = cls.marking_units(rows)
        result = []
        for key, selected in cls._group(units, ("route", "manager")).items():
            rates = cls.marking_unit_rates(selected)
            left = rates.get("左侧标线", {}).get("qualified_rate")
            right = rates.get("右侧标线", {}).get("qualified_rate")
            starts = [u["start_m"] for u in selected]
            ends = [u["end_m"] for u in selected]
            result.append(dict(zip(("route", "manager"), key), left_rate=left, right_rate=right,
                                overall_rate=_gd03_overall(left, right), km=cls._marking_km(selected),
                                unit_count=len(selected), start_m=min(starts), end_m=max(ends)))
        return sorted(result, key=lambda row: cls._route_key(row["route"]))

    @classmethod
    def marking_per_km(cls, rows):
        """逐公里 100 m 单元合格率（左侧/右侧/总体）。"""
        buckets = {}
        for unit in cls.marking_units(rows):
            km = int(unit["start_m"] // 1000)
            entry = buckets.setdefault((unit.get("route"), unit.get("direction"), km), {})
            side = cls._side_of(unit.get("marking_position"))
            cell = entry.setdefault(side, [0, 0])
            cell[0] += 1
            cell[1] += 1 if unit.get("qualified") else 0
        result = []
        for (route, direction, km), sides in buckets.items():
            rates = {name: (cell[1] / cell[0] if cell[0] else None) for name, cell in sides.items()}
            result.append({"route": route, "direction": direction, "km": km,
                           "left_rate": rates.get("左侧标线"), "right_rate": rates.get("右侧标线"),
                           "overall_rate": _gd03_overall(rates.get("左侧标线"), rates.get("右侧标线")),
                           "unit_count": sum(cell[0] for cell in sides.values())})
        return sorted(result, key=lambda row: (cls._route_key(row["route"]), str(row["direction"] or ""), row["km"]))

    @classmethod
    def _km_runs(cls, items, keep):
        """按「km 相邻且达标」切连续段：km 不相邻（中间整公里无检测数据）即断开，
        不把有缺口的公里拼成假连续段（brief-v3 §5 用户修正：不能跨缺口假装连续）。"""
        runs, current, previous = [], [], None
        for item in items:
            if keep(item) and (previous is None or item["km"] == previous + 1):
                current.append(item)
            else:
                if current:
                    runs.append(current)
                current = [item] if keep(item) else []
            previous = item["km"]
        if current:
            runs.append(current)
        return runs

    @classmethod
    def _marking_runs(cls, items, threshold=0.8):
        return cls._km_runs(items, lambda item: item["overall_rate"] is not None
                            and item["overall_rate"] < threshold)

    @classmethod
    def marking_typical_segments(cls, rows, limit=3, threshold=0.8):
        """④典型状况不佳路段：按路线/方向取低于阈值的连续公里段。"""
        grouped = {}
        for item in cls.marking_per_km(rows):
            grouped.setdefault((item["route"], item["direction"]), []).append(item)
        candidates = []
        for (route, direction), items in grouped.items():
            items.sort(key=lambda row: row["km"])
            runs = cls._marking_runs(items, threshold)
            if not runs:
                continue
            run = min(runs, key=lambda block: sum(i["overall_rate"] for i in block) / len(block))
            rates = [i["overall_rate"] for i in run if i["overall_rate"] is not None]
            lefts = [i["left_rate"] for i in run if i["left_rate"] is not None]
            rights = [i["right_rate"] for i in run if i["right_rate"] is not None]
            candidates.append({"route": route, "direction": direction, "km_items": run,
                               "start_m": run[0]["km"] * 1000, "end_m": (run[-1]["km"] + 1) * 1000,
                               "length_km": run[-1]["km"] - run[0]["km"] + 1,
                               "average": sum(rates) / len(rates), "min_rate": min(rates), "max_rate": max(rates),
                               "left_rate": sum(lefts) / len(lefts) if lefts else None,
                               "right_rate": sum(rights) / len(rights) if rights else None})
        return sorted(candidates, key=lambda row: (row["average"], -row["length_km"]))[:limit]

    @classmethod
    def marking_long_runs(cls, rows, window=30, threshold=0.8, min_windows=3):
        """⑤⑥长连续路段（oracle 口径）：按“路线—管养单位—标线侧”取 100 m 单元序列，
        30 个单元（3 km）滑动窗口不合格率 > 80% 的最长连续窗口段，长度 ≥ 3 km 时输出。"""
        grouped = {}
        for unit in cls.marking_units(rows):
            side = cls._side_of(unit.get("marking_position"))
            grouped.setdefault((unit.get("route"), unit.get("manager"), side), []).append(unit)
        output = []
        for (route, manager, side), items in grouped.items():
            items.sort(key=lambda u: u["start_m"])
            flags = [bool(u.get("qualified")) for u in items]
            n = len(flags)
            if n < window:
                continue
            fails = [0] * (n + 1)
            for index, ok in enumerate(flags):
                fails[index + 1] = fails[index] + (0 if ok else 1)
            bad = [fails[i + window] - fails[i] > threshold * window for i in range(n - window + 1)]
            best = cur = 0
            best_end = 0
            for index, value in enumerate(bad):
                cur = cur + 1 if value else 0
                if cur > best:
                    best, best_end = cur, index
            if best < min_windows:
                continue
            start_index = best_end - best + 1
            fail_rates = [(fails[i + window] - fails[i]) / window for i in range(start_index, best_end + 1)]
            output.append({"route": route, "manager": manager, "position": items[0].get("marking_position"),
                           "position_name": side,
                           "start_m": items[start_index]["start_m"],
                           "end_m": items[best_end]["start_m"] + window * 100,
                           "length_km": round(3 + (best - 1) * 0.1, 1),
                           "fail_rate": sum(fail_rates) / len(fail_rates)})
        return sorted(output, key=lambda row: (-row["length_km"], row["fail_rate"]))

    @staticmethod
    def height_kind_stats(rows, kind):
        subset = [r for r in rows if _guardrail_type(r.get("guardrail_type")) == kind and _float(r.get("height")) is not None]
        low, high = HEIGHT_LIMITS[kind]
        qualified = sum(1 for r in subset if low <= float(r["height"]) <= high)
        return {"count": len(subset), "qualified": qualified,
                "rate": qualified / len(subset) if subset else None}

    @classmethod
    def height_overall(cls, rows):
        rows = [r for r in rows if _float(r.get("height")) is not None]
        total = len(rows)
        qualified = 0
        for row in rows:
            kind = _guardrail_type(row.get("guardrail_type"))
            if kind in HEIGHT_LIMITS:
                low, high = HEIGHT_LIMITS[kind]
                qualified += 1 if low <= float(row["height"]) <= high else 0
        km = {(str(r.get("route") or ""), str(r.get("direction") or ""), int(r["station_m"] // 1000))
              for r in rows if r.get("station_m") is not None}
        return {"valid_count": total, "qualified_count": qualified,
                "rate": qualified / total if total else None, "km_segments": len(km),
                "二波": cls.height_kind_stats(rows, "二波"), "三波": cls.height_kind_stats(rows, "三波")}

    @classmethod
    def height_route_units(cls, rows):
        """②按“路线—管养单位”合并的护栏中心高度合格率。"""
        result = []
        for key, selected in cls._group([r for r in rows if _float(r.get("height")) is not None], ("route", "direction", "manager")).items():
            stats = cls.height_overall(selected)
            start_m, end_m = station_span(selected)
            result.append(dict(zip(("route", "direction", "manager"), key),
                               start_m=start_m, end_m=end_m, **stats))
        return sorted(result, key=lambda row: cls._route_key(row["route"]))

    @classmethod
    def height_manager_units(cls, rows):
        result = []
        for key, selected in cls._group([r for r in rows if _float(r.get("height")) is not None], ("manager",)).items():
            stats = cls.height_overall(selected)
            result.append({"manager": key[0],
                           "routes": sorted({str(r.get("route") or "") for r in selected}, key=cls._route_key), **stats})
        return sorted(result, key=lambda row: (row["rate"] is None, row["rate"]))

    @classmethod
    def height_per_km(cls, rows):
        buckets = {}
        for row in rows:
            kind = _guardrail_type(row.get("guardrail_type"))
            if kind not in HEIGHT_LIMITS or _float(row.get("height")) is None or row.get("station_m") is None:
                continue
            key = (row.get("route"), row.get("direction"), int(row["station_m"] // 1000))
            entry = buckets.setdefault(key, [0, 0, 0])
            low, high = HEIGHT_LIMITS[kind]
            entry[0] += 1
            entry[1] += 1 if low <= float(row["height"]) <= high else 0
            entry[2] += 1 if height_deviation_over_10cm(kind, float(row["height"])) else 0
        return [{"route": key[0], "direction": key[1], "km": key[2], "count": value[0],
                 "rate": value[1] / value[0] if value[0] else None,
                 "over_ratio": value[2] / value[0] if value[0] else None, "over_count": value[2]}
                for key, value in sorted(buckets.items(), key=lambda item: (cls._route_key(item[0][0]), str(item[0][1] or ""), item[0][2]))]

    @classmethod
    def height_typical_segments(cls, rows, limit=3, threshold=0.8):
        grouped = {}
        for item in cls.height_per_km(rows):
            grouped.setdefault((item["route"], item["direction"]), []).append(item)
        candidates = []
        for (route, direction), items in grouped.items():
            items.sort(key=lambda row: row["km"])
            runs = cls._km_runs(items, lambda item: item["rate"] is not None
                                and item["rate"] < threshold)
            if not runs:
                continue
            run = min(runs, key=lambda block: sum(i["rate"] for i in block) / len(block))
            rates = [i["rate"] for i in run if i["rate"] is not None]
            candidates.append({"route": route, "direction": direction, "km_items": run,
                               "start_m": run[0]["km"] * 1000, "end_m": (run[-1]["km"] + 1) * 1000,
                               "length_km": run[-1]["km"] - run[0]["km"] + 1,
                               "average": sum(rates) / len(rates), "min_rate": min(rates), "max_rate": max(rates),
                               "count": sum(i["count"] for i in run)})
        return sorted(candidates, key=lambda row: (row["average"], -row["length_km"]))[:limit]

    @classmethod
    def height_over10_runs(cls, rows, threshold=0.05):
        """⑤⑥整公里内（按波形梁类型）|偏差|>100 mm 占比超过阈值的连续公里段。"""
        buckets = {}
        for row in rows:
            kind = _guardrail_type(row.get("guardrail_type"))
            if kind not in HEIGHT_LIMITS or _float(row.get("height")) is None or row.get("station_m") is None:
                continue
            key = (row.get("route"), row.get("direction"), int(row["station_m"] // 1000), kind)
            entry = buckets.setdefault(key, [0, 0, 0])
            low, high = HEIGHT_LIMITS[kind]
            entry[0] += 1
            entry[1] += 1 if low <= float(row["height"]) <= high else 0
            entry[2] += 1 if height_deviation_over_10cm(kind, float(row["height"])) else 0
        grouped = {}
        for (route, direction, km, kind), value in buckets.items():
            grouped.setdefault((route, direction, kind), []).append(
                {"route": route, "direction": direction, "km": km, "kind": kind, "count": value[0],
                 "rate": value[1] / value[0] if value[0] else None,
                 "over_ratio": value[2] / value[0] if value[0] else None, "over_count": value[2]})
        runs = []
        for (route, direction, kind), group in grouped.items():
            group.sort(key=lambda row: row["km"])
            for current in cls._km_runs(group, lambda item: item["over_ratio"] is not None
                                        and item["over_ratio"] > threshold):
                runs.append((route, direction, kind, current))
        output = []
        for route, direction, kind, block in runs:
            output.append({"route": route, "direction": direction, "kind": kind, "km_items": block,
                           "start_m": block[0]["km"] * 1000, "end_m": (block[-1]["km"] + 1) * 1000,
                           "length_km": block[-1]["km"] - block[0]["km"] + 1,
                           "count": sum(i["count"] for i in block),
                           "over_count": sum(i["over_count"] for i in block),
                           "over_ratio": sum(i["over_count"] for i in block) / sum(i["count"] for i in block),
                           "rate": sum(i["rate"] * i["count"] for i in block if i["rate"] is not None) / sum(i["count"] for i in block)})
        return sorted(output, key=lambda row: (-row["length_km"], row["rate"]))

    @classmethod
    def bolt_overall(cls, rows):
        splice = sum(float(r.get("splice", 0) or 0) for r in rows)
        splice_missing = sum(float(r.get("splice_missing", 0) or 0) for r in rows)
        connection = sum(float(r.get("connection", 0) or 0) for r in rows)
        connection_missing = sum(float(r.get("connection_missing", 0) or 0) for r in rows)
        outline = sum(float(r.get("outline", 0) or 0) for r in rows)
        splice_should = splice + splice_missing
        conn_should = connection + connection_missing + outline
        missing = splice_missing + connection_missing
        total = splice_should + conn_should
        km = {(str(r.get("route") or ""), str(r.get("direction") or ""), int(r["station_m"] // 1000))
              for r in rows if r.get("station_m") is not None}
        severe_splice = sum(1 for r in rows if float(r.get("splice_missing", 0) or 0)
                            and (float(r.get("splice", 0) or 0) + float(r.get("splice_missing", 0) or 0))
                            and float(r["splice_missing"]) / (float(r["splice"]) + float(r["splice_missing"])) > 0.5)
        severe_conn = sum(1 for r in rows if float(r.get("connection_missing", 0) or 0)
                          and (float(r.get("connection", 0) or 0) + float(r.get("connection_missing", 0) or 0) + float(r.get("outline", 0) or 0))
                          and float(r["connection_missing"]) / (float(r["connection"]) + float(r["connection_missing"]) + float(r.get("outline", 0) or 0)) > 0.5)
        return {"splice": splice, "splice_missing": splice_missing, "splice_should": splice_should,
                "connection": connection, "connection_missing": connection_missing, "conn_should": conn_should,
                "outline": outline, "missing": missing, "total": total,
                "splice_rate": splice_missing / splice_should if splice_should else None,
                "conn_rate": connection_missing / conn_should if conn_should else None,
                "rate": missing / total if total else None, "km_segments": len(km),
                "severe_splice_points": severe_splice, "severe_conn_points": severe_conn}

    @classmethod
    def bolt_route_units(cls, rows):
        grouped = {}
        for key, selected in cls._group(rows, ("route", "direction", "manager")).items():
            start_m, end_m = station_span(selected)
            unit = dict(zip(("route", "direction", "manager"), key),
                        start_m=start_m, end_m=end_m, **cls.bolt_overall(selected))
            sources = frozenset(str(r.get("source") or "") for r in selected)
            grouped.setdefault(sources, []).append(unit)
        result = []
        for group in grouped.values():
            if len(group) > 1:
                # 同一检测文件按桩号拆分到多个管养单位时，单处严重缺失按文件级统计（与 oracle 一致）
                splice_total = sum(int(r["severe_splice_points"]) for r in group)
                conn_total = sum(int(r["severe_conn_points"]) for r in group)
                for row in group:
                    row["severe_splice_points"] = splice_total
                    row["severe_conn_points"] = conn_total
            result.extend(group)
        return sorted(result, key=lambda row: (row["rate"] is None, row["rate"] or 0, cls._route_key(row["route"])))

    @classmethod
    def bolt_manager_units(cls, rows):
        result = []
        for key, selected in cls._group(rows, ("manager",)).items():
            result.append({"manager": key[0],
                           "routes": sorted({str(r.get("route") or "") for r in selected}, key=cls._route_key),
                           **cls.bolt_overall(selected)})
        return sorted(result, key=lambda row: (row["rate"] is None, -(row["rate"] or 0)))

    @classmethod
    def bolt_per_km(cls, rows):
        buckets = {}
        for row in rows:
            if row.get("station_m") is None:
                continue
            key = (row.get("route"), row.get("direction"), int(row["station_m"] // 1000))
            entry = buckets.setdefault(key, {"splice": 0.0, "splice_missing": 0.0, "connection": 0.0, "connection_missing": 0.0, "outline": 0.0, "missing": 0.0, "severe": 0, "count": 0})
            splice = float(row.get("splice", 0) or 0)
            splice_missing = float(row.get("splice_missing", 0) or 0)
            connection = float(row.get("connection", 0) or 0)
            connection_missing = float(row.get("connection_missing", 0) or 0)
            outline = float(row.get("outline", 0) or 0)
            missing = splice_missing + connection_missing
            entry["splice"] += splice
            entry["splice_missing"] += splice_missing
            entry["connection"] += connection
            entry["connection_missing"] += connection_missing
            entry["outline"] += outline
            entry["missing"] += missing
            entry["count"] += 1
            denominator = splice + splice_missing + connection + connection_missing + outline
            if missing and denominator and missing / denominator > 0.5:
                entry["severe"] += 1
        result = []
        for key, entry in buckets.items():
            denominator = entry["splice"] + entry["splice_missing"] + entry["connection"] + entry["connection_missing"] + entry["outline"]
            splice_should = entry["splice"] + entry["splice_missing"]
            conn_should = entry["connection"] + entry["connection_missing"] + entry["outline"]
            total_should = splice_should + conn_should
            result.append({"route": key[0], "direction": key[1], "km": key[2], "splice": entry["splice"],
                           "connection": entry["connection"], "outline": entry["outline"],
                           "missing": entry["missing"], "severe": entry["severe"],
                           "splice_should": splice_should, "splice_missing": entry["splice_missing"],
                           "splice_rate": entry["splice_missing"] / splice_should if splice_should else 0.0,
                           "conn_should": conn_should, "conn_missing": entry["connection_missing"],
                           "conn_rate": entry["connection_missing"] / conn_should if conn_should else 0.0,
                           "total_should": total_should, "total_missing": entry["missing"],
                           "total_rate": entry["missing"] / total_should if total_should else 0.0,
                           "rate": entry["missing"] / denominator if denominator else None})
        return sorted(result, key=lambda row: (cls._route_key(row["route"]), str(row["direction"] or ""), row["km"]))

    @classmethod
    def bolt_typical_segments(cls, rows, limit=3, threshold=0.05):
        grouped = {}
        for item in cls.bolt_per_km(rows):
            grouped.setdefault((item["route"], item["direction"]), []).append(item)
        candidates = []
        for (route, direction), items in grouped.items():
            items.sort(key=lambda row: row["km"])
            runs = cls._km_runs(items, lambda item: item["rate"] is not None
                                and item["rate"] > threshold)
            if not runs:
                continue
            run = max(runs, key=lambda block: sum(i["missing"] for i in block))
            rates = [i["rate"] for i in run if i["rate"] is not None]
            candidates.append({"route": route, "direction": direction, "km_items": run,
                               "start_m": run[0]["km"] * 1000, "end_m": (run[-1]["km"] + 1) * 1000,
                               "length_km": run[-1]["km"] - run[0]["km"] + 1,
                               "average": sum(rates) / len(rates), "max_rate": max(rates),
                               "missing": sum(i["missing"] for i in run),
                               "severe": sum(i["severe"] for i in run)})
        return sorted(candidates, key=lambda row: (-row["average"], -row["missing"]))[:limit]

    @classmethod
    def bolt_over5_runs(cls, rows, threshold=0.03):
        grouped = {}
        for item in cls.bolt_per_km(rows):
            grouped.setdefault((item["route"], item["direction"]), []).append(item)
        output = []
        for (route, direction), items in grouped.items():
            items.sort(key=lambda row: row["km"])
            for current in cls._km_runs(items, lambda item: item["rate"] is not None
                                        and item["rate"] > threshold):
                output.append((route, direction, current))
        rows_out = []
        for route, direction, block in output:
            splice_should = sum(i["splice_should"] for i in block)
            conn_should = sum(i["conn_should"] for i in block)
            total_should = sum(i["total_should"] for i in block)
            rows_out.append({"route": route, "direction": direction, "km_items": block,
                             "start_m": block[0]["km"] * 1000, "end_m": (block[-1]["km"] + 1) * 1000,
                             "length_km": block[-1]["km"] - block[0]["km"] + 1,
                             "missing": sum(i["missing"] for i in block), "severe": sum(i["severe"] for i in block),
                             # ③ 整公里清单新增列（D13）：拼接/连接缺失率按应安装数量计
                             "splice_rate": (sum(i["splice_missing"] for i in block) / splice_should) if splice_should else None,
                             "conn_rate": (sum(i["conn_missing"] for i in block) / conn_should) if conn_should else None,
                             "rate": sum(i["missing"] for i in block) / (total_should or 1)})
        return rows_out

    @classmethod
    def bolt_severe_clusters(cls, rows, min_points=3, gap_m=20):
        """单处严重缺失（单立柱缺失率>50%）连续出现 ≥3 处的公里段。"""
        buckets = {}
        for row in rows:
            if row.get("station_m") is None:
                continue
            splice = float(row.get("splice", 0) or 0)
            connection = float(row.get("connection", 0) or 0)
            missing = float(row.get("splice_missing", 0) or 0) + float(row.get("connection_missing", 0) or 0)
            denominator = splice + connection + missing
            if not missing or not denominator or missing / denominator <= 0.5:
                continue
            key = (row.get("route"), row.get("direction"), int(row["station_m"] // 1000))
            buckets.setdefault(key, []).append((float(row["station_m"]), missing))
        output = []

        def _cluster(route, direction, km, chunk):
            return {"route": route, "direction": direction, "km": km, "points": len(chunk),
                    "start_m": chunk[0][0], "end_m": chunk[-1][0], "span_m": chunk[-1][0] - chunk[0][0],
                    "max_missing": max(value for _, value in chunk)}

        for (route, direction, km), points in buckets.items():
            points.sort()
            current = [points[0]]
            for point in points[1:]:
                if point[0] - current[-1][0] <= gap_m:
                    current.append(point)
                else:
                    if len(current) >= min_points:
                        output.append(_cluster(route, direction, km, current))
                    current = [point]
            if len(current) >= min_points:
                output.append(_cluster(route, direction, km, current))
        return sorted(output, key=lambda row: (cls._route_key(row["route"]), str(row["direction"] or ""), row["km"]))

    @classmethod
    def fill_managers(cls, marking_rows, *collections):
        """用标线数据按（路线、方向、公里）为护栏/螺栓记录补齐管养单位。"""
        lookup = {}
        for row in marking_rows:
            manager = row.get("manager")
            if not manager or row.get("station_m") is None:
                continue
            key = (str(row.get("route") or ""), str(row.get("direction") or ""), int(row["station_m"] // 1000))
            lookup.setdefault(key, manager)
        # 标线未覆盖的公里（如文件末行 K3144+000）取同路线同方向最近的公里映射，避免边界行丢失管养单位
        km_index = {}
        for route, direction, km in lookup:
            km_index.setdefault((route, direction), []).append(km)
        for kms in km_index.values():
            kms.sort()
        for rows in collections:
            for row in rows:
                if row.get("manager") or row.get("station_m") is None:
                    continue
                route = str(row.get("route") or "")
                direction = str(row.get("direction") or "")
                km = int(row["station_m"] // 1000)
                manager = lookup.get((route, direction, km))
                if not manager:
                    candidates = km_index.get((route, direction))
                    if candidates:
                        manager = lookup[(route, direction, min(candidates, key=lambda item: (abs(item - km), item)))]
                row["manager"] = manager or ""
        return lookup


# 人工复核「结果一致性」口径（brief-v3 §7 用户裁决）：只比较两侧合格判定是否相同，
# 不再用测值偏差/允许偏差阈值。合格判定按设施标准算，比较器行级 consistent、三张明细表的
# 合格判定列与结果一致性列共用下面这三个函数，禁止各自再写一套标准。
GD_HEIGHT_STANDARDS=(("三波",697.0),("两波",600.0),("双波",600.0))
GD_HEIGHT_TOLERANCE_MM=20.0
GD_MARKING_TARGET_WHITE=80.0   # 标线源表无颜色/目标值列：白色目标 80（黄色 50，见 gd_marking_color 同源推导）
GD_BOLT_DIFF_THRESHOLD=5.0     # 螺栓「结果一致性」= 缺失数量差值阈值（颗/柱），界面 boltThreshold 可改，缺省 5


def gd_height_standard(gtype):
    """护栏中心高度标准中心值(mm)：三波 697、两波/双波 600；型式无法识别返回 None（不伪判）。"""
    text=str(gtype or "")
    for key,value in GD_HEIGHT_STANDARDS:
        if key in text: return value
    return None


def gd_pass(indicator, value, gtype=""):
    """单侧合格判定：合格/不合格；数值缺失或标准无法确定时返回 None。
    螺栓没有合格判定（M2：改按缺失数量差值定一致性），一律返回 None。"""
    if value is None: return None
    if indicator=="marking":
        return "合格" if value>=GD_MARKING_TARGET_WHITE else "不合格"
    if indicator=="height":
        standard=gd_height_standard(gtype)
        if standard is None: return None
        return "合格" if abs(value-standard)<=GD_HEIGHT_TOLERANCE_MM else "不合格"
    return None


def gd_consistent(indicator, manual, automatic, gtype="", bolt_threshold=None):
    """标线/高度：两侧合格判定相同→True；一合格一不合格→False；任一侧无法判定→None（禁止 None==None 视作一致）。
    螺栓（brief-v3 §7 第二轮裁决）：对比表没有合格判定列，改按缺失数量差值判定——
    |人工缺失合计−自动缺失合计| ≥ 阈值 → False（不一致），< 阈值 → True；任一侧无数值 → None（显示 —）。"""
    if indicator=="bolt":
        if manual is None or automatic is None: return None
        limit=GD_BOLT_DIFF_THRESHOLD if bolt_threshold is None else bolt_threshold
        return abs(manual-automatic) < limit
    left=gd_pass(indicator,manual,gtype); right=gd_pass(indicator,automatic,gtype)
    if left is None or right is None: return None
    return left==right


class ManualAutoComparator:
    def __init__(self, thresholds): self.thresholds = thresholds

    @staticmethod
    def read_file(path):
        wb=openpyxl.load_workbook(path,read_only=True,data_only=True)
        records=[]; issues=[]
        name=Path(path).name
        source="正式检测后" if "进场后" in name else ("进场前" if "进场前" in name else "")
        for ws in wb.worksheets:
            title=ws.title
            rows=list(ws.iter_rows(values_only=True))
            if not rows: continue
            first="|".join(_norm_header(value) for value in rows[0] if value not in (None,""))
            second="|".join(_norm_header(value) for value in rows[1] if value not in (None,"")) if len(rows)>1 else ""
            combined=first+"|"+second
            if ("人工复核护栏中心平均高度" in combined or "人工护栏中心高度" in combined) and "自动化护栏中心高度" in combined:
                indicator="height"; start=1
            elif "人工逆反射亮度系数平均值" in combined and "自动化逆反射亮度系数平均值" in combined:
                indicator="marking"; start=1
            elif ("拼接螺栓缺失数量" in combined and "连接螺栓缺失数量" in combined and
                  ("人工复核螺栓缺失数量" in combined or "自动化螺栓缺失数量" in combined)):
                indicator="bolt"; start=2
            else:
                issues.append(f"{Path(path).name}/{title}：未按业务表头识别到人工复核指标")
                continue
            headers=list(rows[0])
            for row_index, values in enumerate(rows[start:], start=start+1):
                if not any(v not in (None,"") for v in values): continue
                row=dict(zip(headers,values)); city=_first(row,"地市","地区"); route=_first(row,"路线","路线编号"); direction=_first(row,"方向"); segment=_first(row,"桩号范围","桩号","计算区间")
                # 原样保留展示字段，供报告生成三张分项对比表
                display={"gtype":str(_first(row,"护栏类型",default="") or "").strip(),
                         "position":str(_first(row,"护栏位置","标线位置",default="") or "").strip(),
                         "remark":str(_first(row,"备注",default="") or "").strip()}
                if indicator=="bolt":
                    # 缺一侧数据不可算一致（M1）：两个单元格皆空 → 该侧缺失（None），不静默补 0 当合格
                    def _bolt_side(left,right):
                        pair=[_float(left),_float(right)]
                        return None if all(v is None for v in pair) else sum(v or 0 for v in pair)
                    manual=_bolt_side(values[6],values[7]) if len(values)>9 else None
                    automatic=_bolt_side(values[8],values[9]) if len(values)>9 else None
                    display.update(msplice=_float(values[6]),mconn=_float(values[7]),asplice=_float(values[8]),aconn=_float(values[9]))
                elif indicator=="height":
                    manual=_float(_first(row,"人工复核护栏中心平均高度mm","人工护栏中心高度","人工值"))
                    automatic=_float(_first(row,"自动化护栏中心高度mm标注修正桩号匹配","自动化护栏中心高度标注修正桩号匹配","自动化护栏中心高度mm电子修正桩号匹配","自动化护栏中心高度mm","自动化护栏中心高度","自动化值"))
                else:
                    manual=_float(_first(row,"人工逆反射亮度系数平均值","人工值")); automatic=_float(_first(row,"自动化逆反射亮度系数平均值","自动化值"))
                if not city or not route or manual is None or automatic is None:
                    issues.append(f"{Path(path).name}/{title}/第{row_index}行：人工对比记录字段缺失或非数值"); continue
                records.append({"indicator":indicator,"city":str(city).strip(),"route":_route(route),"direction":str(direction or ""),"segment":str(segment or ""),"source":source,"manual":manual,"automatic":automatic,**display})
        wb.close(); return records,issues

    @staticmethod
    def summarize(detail):
        """按指标汇总人工/自动化对比明细（可传入按道路类别过滤后的明细）。
        标线、螺栓缺失使用相对偏差(%)；护栏中心高度使用绝对偏差(mm)——偏差/MAE/落控计数仅作描述性统计。
        一致性占比按用户最新裁决（brief-v3 §7）用行级 consistent（标线/高度＝两侧合格判定是否相同；
        螺栓＝缺失数量差值 < 阈值）：判定为 None（无法判定）的行不计入分母，另计 unjudged_count。
        within_count 是落阈行数：标线按相对偏差(%)、高度按绝对偏差(mm)、螺栓按缺失数量差值(颗/柱)。"""
        summary={}
        for indicator in ("marking","height","bolt"):
            selected=[x for x in detail if x.get("indicator")==indicator]
            if indicator=="height":
                computable=[x for x in selected if x.get("absolute_difference") is not None]
                metric="absolute_difference"; signed_metric="signed_difference"
            else:
                computable=[x for x in selected if x.get("relative_deviation") is not None]
                metric="relative_deviation"; signed_metric="signed_relative_deviation"
            deviations=[x[metric] for x in computable]; within=sum(bool(x.get("within_threshold")) for x in computable)
            judged=[x for x in selected if x.get("consistent") is not None]
            consistent_count=sum(bool(x.get("consistent")) for x in judged)
            signed=[x[signed_metric] for x in computable if x.get(signed_metric) is not None]
            diffs=[x["signed_difference"] for x in computable if x.get("signed_difference") is not None]
            rels=[x["signed_relative_deviation"] for x in computable if x.get("signed_relative_deviation") is not None]
            abs_signed=[abs(value) for value in signed]
            summary[indicator]={"paired_count":len(selected),"computable_count":len(computable),"min":min(deviations) if deviations else None,
                "max":max(deviations) if deviations else None,"average":sum(deviations)/len(deviations) if deviations else None,
                "within_count":within,"consistent_count":consistent_count,"unjudged_count":len(selected)-len(judged),
                "consistency_rate":consistent_count/len(judged) if judged else None,
                "zero_manual_count":sum(x.get("relative_deviation") is None for x in selected),
                "diff_min":min(diffs) if diffs else None,"diff_max":max(diffs) if diffs else None,
                "diff_average":sum(diffs)/len(diffs) if diffs else None,
                "rel_min":min(rels) if rels else None,"rel_max":max(rels) if rels else None,
                "rel_average":sum(rels)/len(rels) if rels else None,
                "mae":sum(abs_signed)/len(abs_signed) if abs_signed else (sum(deviations)/len(deviations) if deviations else None),
                "mae_relative":sum(abs(value) for value in rels)/len(rels) if rels else None,
                "within5_count":sum(1 for value in abs_signed if value<=5),
                "within10_count":sum(1 for value in abs_signed if value<=10),
                "high_count":sum(1 for value in signed if value>0),
                "low_count":sum(1 for value in signed if value<0),
                "zero_count":sum(1 for value in signed if value==0)}
        return summary

    def compare(self, records):
        detail=[]
        for record in records:
            item=dict(record); manual=_float(item.get("manual")); automatic=_float(item.get("automatic"))
            if manual is None or automatic is None: continue
            item["absolute_difference"]=abs(automatic-manual)
            item["relative_deviation"]=None if manual == 0 else abs(automatic-manual)/abs(manual)*100
            item["signed_difference"]=automatic-manual
            item["signed_relative_deviation"]=None if manual == 0 else (automatic-manual)/abs(manual)*100
            indicator=item["indicator"]; threshold=self.thresholds[indicator]
            if indicator=="height":
                # 护栏中心高度：使用绝对偏差(mm)判断，不使用百分比
                item["within_threshold"]=item["absolute_difference"] <= threshold
            elif indicator=="bolt":
                # M2（第二轮裁决）：螺栓改按缺失数量差值（颗/柱）判断，差值 < 阈值即落阈
                item["within_threshold"]=item["absolute_difference"] < threshold
            else:
                # 标线：使用相对偏差(%)判断
                item["within_threshold"]=None if manual == 0 else item["relative_deviation"] <= threshold
            # 结果一致性（brief-v3 §7）：标线/高度只比两侧合格判定是否相同，与测值偏差/阈值无关；
            # 螺栓无合格判定列，按两侧缺失数量差值与阈值比较。任一侧无法判定 → None（显示 —）。
            item["consistent"]=gd_consistent(indicator, manual, automatic, item.get("gtype") or "", threshold)
            detail.append(item)
        return detail, self.summarize(detail)





# 附件模板-4 各表列宽（cm）：按表头签名套用。
GD_TABLE_WIDTHS = {
    # ① 汇总表列宽（模板-1 实测，structure-diff §2.5）：高速 = 管理单位级，普通 = 类别级
    ("管理单位", "抽检里程(km)", "总体合格率(%)", "主车道左侧合格率(%)", "主车道右侧合格率(%)"): [2.81, 2.19, 2.31, 2.16, 2.53],
    ("管理单位", "抽检里程(km)", "总体合格率(%)", "两波合格率(%)", "三波合格率(%)"): [2.81, 2.19, 2.31, 2.16, 2.53],
    ("管理单位", "抽检里程(km)", "总体缺失率(%)", "拼接螺栓缺失率(%)", "连接螺栓缺失率(%)"): [2.81, 2.19, 2.31, 2.16, 2.53],
    ("道路类别", "抽检里程(km)", "总体合格率(%)", "主车道左侧合格率(%)", "主车道右侧合格率(%)"): [3.53, 1.98, 2.31, 2.16, 2.53],
    ("类别", "抽检里程(km)", "总体合格率(%)", "两波合格率(%)", "三波合格率(%)"): [3.34, 2.19, 2.31, 2.16, 2.53],
    ("类别", "抽检里程(km)", "总体缺失率(%)", "拼接螺栓缺失率(%)", "连接螺栓缺失率(%)"): [3.59, 1.96, 2.31, 2.16, 2.53],
    ("路线编号", "路线名称", "管养单位", "起止桩号", "里程(km)", "总体合格率%", "左侧合格率%", "右侧合格率%"): [1.16, 1.63, 4.10, 2.17, 1.23, 1.75, 1.75, 1.75],
    ("路线编号", "路线名称", "管养单位", "检测范围", "里程(km)", "总体合格率(%)", "左侧合格率(%)", "右侧合格率(%)"): [1.17, 1.53, 3.46, 2.36, 1.15, 1.87, 1.87, 1.87],
    ("路线编号", "路线名称", "管养单位", "起止桩号", "里程(km)", "总体合格率(%)", "两波合格率(%)", "三波合格率(%)"): [1.01, 1.60, 3.64, 2.21, 1.18, 1.78, 1.78, 1.78],
    ("路线编号", "路线名称", "管养单位", "起止桩号", "里程（Km）", "总体缺失率%", "拼接缺失率%", "连接缺失率%"): [1.11, 1.77, 3.72, 2.49, 1.36, 1.60, 1.37, 1.38],
    ("路线编号", "路线名称", "管养单位", "起止桩号", "里程（Km)", "总体缺失率（%）", "拼接缺失率（%）", "连接缺失率（%）"): [1.06, 1.81, 3.66, 2.59, 1.15, 1.56, 1.53, 1.52],
    ("路线编号", "管养单位", "起止桩号", "总体合格率(%)", "左侧合格率(%)", "右侧合格率(%)"): [1.19, 5.88, 2.16, 1.84, 1.84, 1.78],
    # ③ 其它分支表（tsv 实测）
    ("路线编号", "管养单位", "起止桩号", "总体合格率(%)", "两波合格率(%)", "三波合格率(%)"): [1.09, 5.99, 2.14, 1.86, 1.82, 1.78],
    ("路线编号", "管养单位", "起止桩号", "总体合格率%", "两波合格率%", "三波合格率%"): [1.09, 5.99, 2.14, 1.86, 1.82, 1.78],
    ("路线编号", "路线名称", "管养单位", "起止桩号", "里程(km)", "总体合格率%", "两波合格率%", "三波合格率%"): [1.01, 1.60, 3.96, 2.14, 1.20, 1.69, 1.71, 1.67],
    ("路线", "管养单位", "起止桩号", "缺失处数", "单处最大缺失(颗/柱)"): [1.23, 4.48, 3.09, 1.21, 2.03],
    ("路线", "管养单位", "起止桩号", "总体缺失率%", "拼接缺失率%", "连接缺失率%"): [1.27, 4.48, 2.27, 1.46, 1.46, 1.43],
    ("路线", "路线名称", "管养单位", "起止桩号", "跨度(m)", "连续处数", "单处最大缺失(颗/柱)"): [1.23, 1.88, 3.29, 3.25, 1.61, 1.27, 2.10],
    ("路线", "路线名称", "管养单位", "起止桩号", "总体缺失率%", "拼接缺失率%", "连接缺失率%"): [1.27, 2.26, 4.48, 2.27, 1.46, 1.46, 1.43],
    # 人工复核（标线两分支同宽；高度/螺栓两分支列宽不同 → 见 GD_TABLE_WIDTHS_BY_CAPTION）
    ("路线编号", "桩号区段", "标线颜色", "人工检测复核结果", "人工检测复核结果", "自动化检测结果", "自动化检测结果", "结果一致性"): [1.34, 3.55, 1.63, 1.63, 1.79, 1.46, 1.76, 1.58],
    ("路线编号", "桩号区段", "护栏类型", "合格值(mm)", "人工检测复核结果", "人工检测复核结果", "自动化检测结果", "自动化检测结果", "结果一致性"): [1.23, 3.26, 1.0, 2.17, 1.25, 1.59, 1.23, 1.62, 1.44],
    ("路线", "桩号区段", "护栏类型", "人工检测复核结果", "人工检测复核结果", "人工检测复核结果", "自动化检测结果", "自动化检测结果", "自动化检测结果", "结果一致性"): [1.52, 2.56, 1.26, 1.31, 1.31, 1.32, 1.32, 1.32, 1.32, 1.37],
    ("类型", "路线编号", "路线名称", "检测方向", "起点桩号", "终点桩号", "段长（Km)", "管养单位", "备注"): [1.36, 1.04, 1.78, 1.09, 1.09, 1.12, 1.31, 4.25, 2.47],
}

# 表头签名相同但两分支列宽不同的表 → 按表题取（tsv 实测）；表题已去掉 `表5-N ` 前缀。
GD_TABLE_WIDTHS_BY_CAPTION = {
    "高速公路标线逆反射亮度系数不佳路段汇总表": [1.19, 5.88, 2.16, 1.84, 1.84, 1.78],
    "普通国省道标线逆反射亮度系数不佳路段汇总表": [1.13, 5.91, 2.21, 1.85, 1.79, 1.8],
    "高速公路波形梁护栏中心高度人工复核对比明细表": [1.23, 3.26, 1.0, 2.17, 1.25, 1.59, 1.23, 1.62, 1.44],
    "普通国省道波形梁护栏中心高度人工复核对比明细表": [1.23, 3.29, 1.0, 2.16, 1.25, 1.59, 1.23, 1.62, 1.43],
    "高速公路路侧波形梁护栏螺栓缺失人工复核对比明细表": [1.52, 2.56, 1.26, 1.31, 1.31, 1.32, 1.32, 1.32, 1.32, 1.37],
    "普通国省道路侧波形梁护栏螺栓缺失人工复核对比明细表": [1.44, 3.29, 1.18, 1.23, 1.23, 1.24, 1.24, 1.24, 1.24, 1.29],
}


def gd_table_captions(document):
    """文档顺序下每张表对应的表题（去掉 `表5-N ` 前缀；无题注的表为空串）。"""
    from docx.oxml.ns import qn

    captions, pending = [], ""
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            text = "".join(node.text or "" for node in child.iter(qn("w:t"))).strip()
            if text.startswith("表5-"):
                pending = re.sub(r"^表5-\d+\s*", "", text)
        elif child.tag == qn("w:tbl"):
            captions.append(pending)
            pending = ""
    return captions


def _norm_header_text(value):
    return "".join(str(value or "").split())


def apply_gd_table_widths(document):
    """按表题（分支列宽不同时）或表头签名套用附件列宽；未命中的表保持默认宽度。"""
    from docx.shared import Cm

    captions = gd_table_captions(document)
    applied = 0
    for index, table in enumerate(document.tables):
        if not table.rows:
            continue
        headers = tuple(_norm_header_text(cell.text) for cell in table.rows[0].cells)
        widths = GD_TABLE_WIDTHS_BY_CAPTION.get(captions[index] if index < len(captions) else "")
        widths = widths or GD_TABLE_WIDTHS.get(headers)
        if not widths or len(widths) != len(table.columns):
            continue
        for column, width in zip(table.columns, widths):
            column.width = Cm(width)
        for row in table.rows:
            for cell, width in zip(row.cells, widths):
                cell.width = Cm(width)
        applied += 1
    return applied


def apply_gd_table_heights(document):
    """T2h/A2：按表题套用模板-1 行高（表头/数据行两档，缺值 397）。
    用 atLeast 而非 exact：多行单元格（类型列换行）不会被截断。"""
    from docx.enum.table import WD_ROW_HEIGHT_RULE
    from docx.shared import Emu

    captions = gd_table_captions(document)
    applied = 0
    for index, table in enumerate(document.tables):
        if not table.rows:
            continue
        key = captions[index] if index < len(captions) else ""
        if key.endswith("交通安全设施抽检路段清单"):
            key = "{市}交通安全设施抽检路段清单"  # T2i：题注含市名，tsv 键用占位
        header, data = GD_TABLE_HEIGHTS_BY_CAPTION.get(key, (397, 397))
        for position, row in enumerate(table.rows):
            row.height = Emu(int(header if position == 0 else data) * 635)
            row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        applied += 1
    return applied


_STATION_SEPARATORS = ("\uff5e", "\u301c", "\u007e")  # ～ 〜 ~


def normalize_station_text(value):
    """T2k：桩号区间连接符统一为 ASCII `-`（docx 与 xlsx 共用的单一文本出口）。"""
    text = str(value)
    for old in _STATION_SEPARATORS:
        text = text.replace(old, "-")
    return text


def normalize_station_separators(document):
    """T2j：docx 全文（正文+表格，含源自输入数据的 `～`/`〜`/`~`）归一化；`—` 不动。"""
    from docx.oxml.ns import qn

    changed = 0
    for node in document.element.iter(qn("w:t")):
        original = node.text or ""
        fixed = normalize_station_text(original)
        if fixed != original:
            node.text = fixed
            changed += 1
    return changed


def normalize_workbook_separators(workbook):
    """T2k：xlsx 全部工作表单元格归一化（源数据自带的 `～`/`~` 桩号区间同样覆盖）。

    ponytail: 工作簿内不存在「量纲区间」单元格（扫描验证：命中项 100% 为桩号），故整体归一化；
    若将来 xlsx 里出现 `580～620mm` 类标注，改此处为只匹配 `K\d` 桩号的正则。"""
    changed = 0
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    fixed = normalize_station_text(cell.value)
                    if fixed != cell.value:
                        cell.value = fixed
                        changed += 1
    return changed


def gd_stacked_category(value):
    """T2h/A4：清单表「类型」列照模板手动换行（高速\\n公路 / 普通\\n国省道）。"""
    text = str(value or "")
    for head in ("高速公路", "普通国省道"):
        if text.startswith(head):
            return head[:2] + "\n" + head[2:]
    return text

GD_TABLE_HEIGHTS_BY_CAPTION = {
    # T2i：键 = 产物表题（`{市}` 为占位）；值 = (表头行高, 数据行高)，twips，缺值 fallback=397。
    # 权威源 requirements/template-rows.tsv（含模板题注别名与逐行原始值）。
    "{市}交通安全设施抽检路段清单": (397, 397),
    "抽检普通国省道护栏横梁中心高度合格率汇总表": (397, 397),
    "抽检普通国省道标线逆反射亮度系数合格率汇总表": (340, 340),
    "普通国省道各抽检路段标线逆反射亮度系数合格率汇总表": (397, 397),
    "普通国省道各抽检路段波形梁护栏中心高度合格率汇总表": (510, 510),
    "普通国省道各路段公司螺栓缺失率汇总表": (397, 432),
    "普通国省道整公里螺栓缺失率大于3%严重路段清单": (397, 397),
    "普通国省道标线逆反射亮度系数不佳路段汇总表": (454, 454),
    "普通国省道标线逆反射亮度系数人工复核对比明细表": (397, 397),
    "普通国省道波形梁护栏中心高度不佳路段汇总表": (601, 283),
    "普通国省道波形梁护栏中心高度人工复核对比明细表": (397, 283),
    "普通国省道波形梁护栏螺栓缺失率汇总表": (397, 397),
    "普通国省道螺栓缺失不佳长连续段清单": (397, 397),
    "普通国省道路侧波形梁护栏螺栓缺失人工复核对比明细表": (283, 283),
    "高速公路各管理单位护栏横梁中心高度合格率汇总表": (397, 397),
    "高速公路各管理单位标线逆反射亮度系数合格率汇总表": (397, 397),
    "高速公路各管理单位波形梁护栏螺栓缺失率汇总表": (397, 397),
    "高速公路各路段公司标线逆反射亮度系数合格率汇总表": (510, 510),
    "高速公路各路段公司波形梁护栏中心高度合格率汇总表": (510, 510),
    "高速公路各路段公司螺栓缺失率汇总表": (510, 510),
    "高速公路整公里螺栓缺失率大于3%路段清单": (397, 397),
    "高速公路标线逆反射亮度系数不佳路段汇总表": (600, 283),
    "高速公路标线逆反射亮度系数人工复核对比明细表": (397, 397),
    "高速公路波形梁护栏中心高度不佳路段汇总表": (601, 283),
    "高速公路波形梁护栏中心高度人工复核对比明细表": (397, 283),
    "高速公路螺栓缺失不佳长连续段清单": (397, 397),
    "高速公路路侧波形梁护栏螺栓缺失人工复核对比明细表": (283, 283),
}


class GuangdongChapterWriter:
    @classmethod
    def _format_config(cls):
        from backend import minimal_docx
        return minimal_docx._current_format_config()

    @staticmethod
    def _fmt(value, digits=2):
        if value is None or value == "": return "—"
        if isinstance(value, (int, float)): return f"{value:,.{digits}f}"
        return str(value)

    @staticmethod
    def _pct(value):
        return "—" if value is None else f"{value:.2%}"

    @staticmethod
    def _segment_text(value):
        text=str(value or "—")
        return text if text.endswith("段") or text == "—" else text + "段"

    @staticmethod
    def _picture(doc, path, width):
        """插图统一居中、无缩进。直接 add_picture 会继承正文首行缩进把图片右推。"""
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from backend import minimal_docx
        paragraph = doc.add_paragraph()
        paragraph.add_run().add_picture(str(path), width=width)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        minimal_docx._clear_indent(paragraph)
        return paragraph

    @classmethod
    def _add_table(cls, doc, headers, rows, merge=None, vmerge=None, tail_merge=None):
        """表格：所有单元格文字无缩进、水平居中、垂直居中；表头加粗仿宋_GB2312。
        merge={起始列: 跨列数}：两级表头——第一行主表头（横向合并），其下列子表头；单列表头纵向合并两行。
        tail_merge={起始列: 跨列数}：尾行横向合并（② 表合计行照模板 `G4G4G4G4....` 形态）。"""
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from backend import minimal_docx
        format_config = cls._format_config()["table"]

        def _style(cell, bold=False):
            for p in cell.paragraphs:
                minimal_docx._apply_paragraph(p, format_config, clear_indent=True)
                pPr=p._p.get_or_add_pPr()
                ind=pPr.find(qn('w:ind'))
                if ind is None:
                    ind=OxmlElement('w:ind'); pPr.append(ind)
                ind.set(qn('w:firstLineChars'),'0'); ind.set(qn('w:firstLine'),'0')
                # T2h/A3：模板单元格行距 = 12pt 固定
                _sp = pPr.find(qn('w:spacing'))
                if _sp is None:
                    _sp = OxmlElement('w:spacing'); pPr.append(_sp)
                _sp.set(qn('w:line'), '240'); _sp.set(qn('w:lineRule'), 'exact')
                tcPr=cell._tc.get_or_add_tcPr()
                vAlign=tcPr.find(qn('w:vAlign'))
                if vAlign is None:
                    vAlign=OxmlElement('w:vAlign'); tcPr.append(vAlign)
                vAlign.set(qn('w:val'),'center')  # 垂直居中
                for run in p.runs:
                    minimal_docx._apply_run(run, format_config, bold=bold and format_config["header_bold"])

        if merge:
            # 两级表头：merge={列号: (主表头名, 跨列数)}，原生 XML 构建（gridSpan/vMerge），物理列 = 逻辑列 - Σ(跨列数-1)
            from docx.table import _Cell
            merge = {c: (v[0], int(v[1])) for c, v in merge.items()}
            phys_of={}; p=0; i=0
            while i < len(headers):
                if i in merge:
                    n=merge[i][1]
                    for j in range(i, i+n): phys_of[j]=p
                    p+=n-1; i+=n
                else:
                    phys_of[i]=p; p+=1; i+=1
            physical=p
            table=doc.add_table(rows=2,cols=physical); table.style="Table Grid"
            # 清空默认两行，手工构建 tr
            tbl=table._tbl
            for tr in list(tbl.findall(qn('w:tr'))): tbl.remove(tr)
            def _tc(text, span=None, vmerge=None, bold=True):
                tc=OxmlElement('w:tc'); tcPr=OxmlElement('w:tcPr')
                if span: sp=OxmlElement('w:gridSpan'); sp.set(qn('w:val'),str(span)); tcPr.append(sp)
                if vmerge is not None:
                    vm=OxmlElement('w:vMerge'); vm.set(qn('w:val'),vmerge) if vmerge!="continue" else None; tcPr.append(vm)
                v_align=OxmlElement('w:vAlign'); v_align.set(qn('w:val'),'center'); tcPr.append(v_align)
                tc.append(tcPr); par=OxmlElement('w:p'); r=OxmlElement('w:r')
                rPr=OxmlElement('w:rPr')
                if bold and format_config["header_bold"]: b=OxmlElement('w:b'); rPr.append(b)
                fonts=OxmlElement('w:rFonts'); fonts.set(qn('w:eastAsia'),format_config["east_asia"]); fonts.set(qn('w:ascii'),format_config["latin"]); fonts.set(qn('w:hAnsi'),format_config["latin"]); rPr.append(fonts)
                pPr=OxmlElement('w:pPr')
                _sp=OxmlElement('w:spacing'); _sp.set(qn('w:line'),'240'); _sp.set(qn('w:lineRule'),'exact'); pPr.append(_sp)
                jc=OxmlElement('w:jc'); jc.set(qn('w:val'),format_config["alignment"] if format_config["alignment"] != "justify" else "center"); pPr.append(jc)
                r.append(rPr)
                t=OxmlElement('w:t'); t.text=str(text); r.append(t); par.append(pPr); par.append(r); tc.append(par)
                return tc
            top=OxmlElement('w:tr'); sub=OxmlElement('w:tr')
            i=0
            while i < len(headers):
                if i in merge and merge[i][0]:
                    parent,n=merge[i]
                    top.append(_tc(parent, span=n-1))
                    for j in range(i+1, i+n):
                        sub.append(_tc(headers[j]))
                    i+=n
                elif i in merge:
                    top.append(_tc(headers[i], vmerge="restart"))
                    sub.append(_tc("", vmerge="continue"))
                    i+=1
                else:
                    top.append(_tc(headers[i], vmerge="restart"))
                    sub.append(_tc("", vmerge="continue"))
                    i+=1
            tbl.append(top); tbl.append(sub)
        else:
            table=doc.add_table(rows=1,cols=len(headers)); table.style="Table Grid"
            for index, header in enumerate(headers):
                cell=table.rows[0].cells[index]; cell.text=str(header); _style(cell, True)
        if not rows:
            blank_cols = physical if merge else len(headers)
            rows = [["—"] + [""] * (blank_cols - 1)]
        for row in rows:
            cells=table.add_row().cells
            for index,value in enumerate(row):
                cells[index].text=str(value)
                _style(cells[index])
        if vmerge:
            # 数据行纵向合并：vmerge={列号: [(起始数据行, 行数), ...]}（表头占第 0 行）
            for col, groups in vmerge.items():
                for first, count in groups:
                    if count < 2:
                        continue
                    cell = table.rows[first + 1].cells[col].merge(table.rows[first + count].cells[col])
                    for extra in cell.paragraphs[1:]:
                        extra._p.getparent().remove(extra._p)
        if tail_merge:
            # D2：尾行（合计行）横向合并，文本取起始列原值；多余并列段落清掉
            first_data = 2 if merge else 1
            for col, span in tail_merge.items():
                if span < 2 or len(table.rows) <= first_data:
                    continue
                last = table.rows[-1]
                merged = last.cells[col].merge(last.cells[col + span - 1])
                for extra in merged.paragraphs[1:]:
                    extra._p.getparent().remove(extra._p)
        # T2h/A1：模板每表 tblLayout=fixed（缺它会按内容 auto-fit，窄列被压扁、列宽漂移）
        _layout = OxmlElement("w:tblLayout"); _layout.set(qn("w:type"), "fixed")
        table._tbl.tblPr.append(_layout)
        if not format_config["allow_row_break"]:
            for row in table.rows:
                tr_pr = row._tr.get_or_add_trPr()
                tr_pr.append(OxmlElement("w:cantSplit"))
        # 参考件 70 表中 21 表设置表头跨页重复；两级表头时两行都重复。
        from docx.enum.table import WD_TABLE_ALIGNMENT
        table.alignment = WD_TABLE_ALIGNMENT.CENTER  # 表格整体居中于页面（单元格文字居中已在 _style 处理）
        header_rows = 2 if merge else 1
        for index, row in enumerate(table.rows):
            if index >= header_rows:
                break
            row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
        # 表头跨页重复会让“只有表头的孤行表”落在页脚：给表头行加 keepNext，
        # 强制表头与首行数据同页（Word 逐行排版时才会把整表推到下一页）。
        for index, row in enumerate(table.rows):
            if index >= header_rows:
                break
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.keep_with_next = True
        return table

    @staticmethod
    def _strip_template_notes(doc):
        """删除模板自带的 HTML 注释说明段（<!-- ... -->），它们不是报告内容。"""
        for paragraph in list(doc.paragraphs):
            text = paragraph.text.strip()
            if text.startswith("<!--") and text.endswith("-->"):
                paragraph._p.getparent().remove(paragraph._p)

    @classmethod
    def _body(cls, doc, text, keep_next=False):
        from backend import minimal_docx
        format_config = cls._format_config()["body"]
        paragraph = doc.add_paragraph()
        minimal_docx._apply_paragraph(paragraph, format_config)
        run = paragraph.add_run(text)
        minimal_docx._apply_run(run, format_config)
        if keep_next:
            paragraph.paragraph_format.keep_with_next = True
        return paragraph

    @classmethod
    def _indent_existing_body(cls, doc):
        """给已有正文段落补模板配置的首行缩进；跳过标题、题注与空段。"""
        from backend import minimal_docx
        format_config = cls._format_config()["body"]
        for paragraph in doc.paragraphs:
            if not paragraph.text.strip():
                continue
            style=paragraph.style.name if paragraph.style is not None else ""
            if style.startswith("Heading") or style == "Caption":
                continue
            minimal_docx._apply_paragraph(paragraph, format_config)

    @classmethod
    def _remove_template_placeholder(cls, doc):
        """删除模板自带的占位章节结构：从首个 '（一）' 起，到首个 '（三）'前，
        移除所有直系 body 段落/表格元素（含 Heading2/3/4 占位标题及其下属正文），
        但保留末尾 sectPr（节属性）。若模板不含 '（三）'，则删到 sectPr 之前为止。"""
        from docx.oxml.ns import qn
        body=doc.element.body
        children=list(body)
        start_idx=None; end_idx=len(children)
        for i,el in enumerate(children):
            if el.tag.endswith('}sectPr'):
                end_idx=i  # 永远不删除节属性
                break
            if el.tag.endswith('}p'):
                texts="".join(node.text or "" for node in el.iter(qn('w:t')))
                if start_idx is None and texts.startswith("（一）"):
                    start_idx=i
                elif start_idx is not None and texts.startswith("（三）"):
                    end_idx=i; break
        if start_idx is None:
            return
        for el in children[start_idx:end_idx]:
            body.remove(el)

    @staticmethod
    def _outline_level(paragraph):
        """返回 Word 实际大纲等级（1-based）；优先读取 w:outlineLvl。"""
        from docx.oxml.ns import qn
        import re
        pPr = paragraph._p.pPr
        if pPr is not None:
            outline = pPr.find(qn("w:outlineLvl"))
            if outline is not None:
                try:
                    return int(outline.get(qn("w:val"))) + 1
                except (TypeError, ValueError):
                    pass
        style = paragraph.style.name if paragraph.style is not None else ""
        match = re.search(r"Heading\s*(\d+)", style, re.I)
        return int(match.group(1)) if match else None

    @staticmethod
    def _clear_paragraph_indent(paragraph):
        """标题/题注不保留模板首行、左右或悬挂缩进。"""
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Pt
        fmt = paragraph.paragraph_format
        fmt.left_indent = Pt(0)
        fmt.right_indent = Pt(0)
        fmt.first_line_indent = Pt(0)
        pPr = paragraph._p.get_or_add_pPr()
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            pPr.append(ind)
        for key in ("left", "right", "firstLine", "hanging", "leftChars", "rightChars", "firstLineChars", "hangingChars"):
            ind.set(qn(f"w:{key}"), "0")

    @classmethod
    def _apply_heading_format(cls, paragraph, level, main_title=False):
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Pt, RGBColor
        level = max(1, int(level or 1))
        pPr = paragraph._p.get_or_add_pPr()
        outline = pPr.find(qn("w:outlineLvl"))
        if outline is None:
            outline = OxmlElement("w:outlineLvl")
            pPr.append(outline)
        outline.set(qn("w:val"), str(level - 1))
        paragraph.alignment = 0
        cls._clear_paragraph_indent(paragraph)
        east_asia = "黑体" if level in (1, 2, 3) else "仿宋_GB2312"
        size = 16 if main_title else {1: 16, 2: 15, 3: 14, 4: 12, 5: 10.5}.get(level, 10.5)
        for run in paragraph.runs:
            run.bold = False
            run.italic = False
            run.font.size = Pt(size)
            run.font.color.rgb = RGBColor(0, 0, 0)
            cls._set_run_fonts(run, east_asia)

    @staticmethod
    def _outline_level(paragraph):
        from docx.oxml.ns import qn
        import re
        pPr = paragraph._p.pPr
        if pPr is not None:
            node = pPr.find(qn("w:outlineLvl"))
            if node is not None:
                try:
                    return int(node.get(qn("w:val"))) + 1
                except (TypeError, ValueError):
                    pass
        style = paragraph.style.name if paragraph.style is not None else ""
        match = re.search(r"Heading\s*(\d+)", style, re.I)
        return int(match.group(1)) if match else None

    @staticmethod
    def _clear_paragraph_indent(paragraph):
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Pt
        fmt = paragraph.paragraph_format
        fmt.left_indent = Pt(0)
        fmt.right_indent = Pt(0)
        fmt.first_line_indent = Pt(0)
        pPr = paragraph._p.get_or_add_pPr()
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            pPr.append(ind)
        for key in ("left", "right", "firstLine", "hanging", "leftChars", "rightChars", "firstLineChars", "hangingChars"):
            ind.set(qn(f"w:{key}"), "0")

    @classmethod
    def _apply_heading_format(cls, paragraph, level, main_title=False):
        from backend import minimal_docx
        level = max(1, min(int(level or 1), 5))
        heading_config = cls._format_config()["heading"][str(level)]
        minimal_docx._apply_paragraph(paragraph, heading_config, clear_indent=True)
        minimal_docx._warn(paragraph, heading_config["outline_level"] + 1)
        for run in paragraph.runs:
            minimal_docx._apply_run(run, heading_config)

    @classmethod
    def _heading(cls, doc, text, level):
        level = max(1, min(int(level or 1), 5))
        paragraph = doc.add_paragraph()
        if f"Heading {level}" in doc.styles:
            paragraph.style = doc.styles[f"Heading {level}"]
        paragraph.add_run(text)
        cls._apply_heading_format(paragraph, level)
        return paragraph

    # ==================== GD03 四段式章节渲染（brief-gd03 §1–§2） ====================

    @staticmethod
    def _g03_rate(value, digits=1):
        return "—" if value is None else f"{value * 100:.{digits}f}%"

    @staticmethod
    def _g03_num(value, digits=1):
        return "—" if value is None else f"{value * 100:.{digits}f}"

    @staticmethod
    def _g03_int(value):
        return f"{int(round(float(value or 0))):,}"

    @staticmethod
    def _g03_cell(value):
        """T2i：③ 表空值写 `/`；有数据写两位小数（不带 %，列头已写单位）。"""
        return "/" if value is None else GuangdongChapterWriter._g03_num(value, 2)

    @staticmethod
    def _gd03_range(start_m, end_m):
        if start_m is None or end_m is None:
            return "—"
        return f"{format_station(start_m)}-{format_station(end_m)}"

    @staticmethod
    def _g03_km_range(km):
        """③表逐公里起止桩号：K407-K408（模板表 12/15/33/36 逐公里写法，不带米数）。"""
        return f"K{km}-K{km + 1}"

    @classmethod
    def _g03_km_manager(cls, rows, item, km, city=""):
        """③表逐公里行的管养单位：按该公里回查明细行（管养单位可能段内变化）。"""
        return cls._manager_for_km(rows, dict(item, start_m=int(km) * 1000), city)

    @classmethod
    def _g03_km_kinds(cls, rows, item, km):
        """③表逐公里行的两波/三波合格率：按该公里桩号范围从 raw height 明细分别重算
        （brief-v3 §5：不得把段级两波/三波率复制给每公里）。"""
        subset = dict(item, start_m=int(km) * 1000, end_m=(int(km) + 1) * 1000)
        return gd_height_segment_kinds(rows, subset)

    @staticmethod
    def _marking_route_ranges(rows):
        """按路线聚合标线桩号范围（供普通国省道③④“检测范围”列）。"""
        spans = {}
        for row in rows:
            station = _float(row.get("station_m"))
            if station is None:
                continue
            key = str(row.get("route") or "")
            end = _float(row.get("end_m"))
            high = end if end is not None else station
            low, top = spans.get(key, (station, station))
            spans[key] = (min(low, station), max(top, high))
        return {key: f"{key}：{format_station(low)}-{format_station(high)}" for key, (low, high) in spans.items()}

    @staticmethod
    def _manager_for_km(rows, run, city=""):
        """⑤/⑥清单表“管养单位”列：按清单首公里从明细行回查。"""
        km = int(run["start_m"] // 1000)
        for row in rows:
            if (row.get("route") == run["route"] and row.get("direction") == run["direction"]
                    and row.get("station_m") is not None and int(row["station_m"] // 1000) == km
                    and row.get("manager")):
                return manager_display(row.get("manager"), city)
        return "—"

    @classmethod
    def _g03_chart(cls, base, prefix, key, categories, series, ylabel="合格率（%）", kind="bar", rotation=0, value_fmt="%.1f", ylim=None, horizontal=False, scale=100.0):
        path = Path(base) / f"{prefix}gd03_{key}.png"
        path.parent.mkdir(parents=True, exist_ok=True)
        builder = gd03_bar_chart if kind == "bar" else gd03_line_chart
        kwargs = dict(ylabel=ylabel, rotation=rotation, value_fmt=value_fmt)
        if ylim is not None:
            kwargs["ylim"] = ylim
        if horizontal:
            kwargs["horizontal"] = True
        # 口径统一：合格率/缺失率传入的是 0–1 比率，默认换算为百分数（与附件图表单位一致）；计数类传 scale=1。
        percent_series = [(name, [None if value is None else value * scale + 0.0 for value in values], color)
                          for name, values, color in series]
        try:
            builder(str(path), categories, percent_series, **kwargs)
        except Exception:
            return None
        return str(path)

    @classmethod
    def _marking_gd03_section(cls, doc, category, rows, table, figure, base, chart_prefix, city="", route_names=None,
                              inspection=None):
        heading = lambda text: cls._heading(doc, text, 5)
        route_names = route_names or {}
        units = GuangdongStatistics.marking_units(rows)
        if not units:
            heading("①总体情况")
            cls._body(doc, f"本次未读取到{category}标线逆反射亮度系数有效检测数据，该指标暂不评价。")
            return
        overall = GuangdongStatistics.marking_overall(units)
        route_units = GuangdongStatistics.marking_route_units(rows)

        heading("①总体情况")
        cls._body(doc, f"{city}抽检{'高速公路' if category == '高速公路' else '普通国省道'}主车道标线总体合格率为"
                       f"{cls._g03_num(overall['overall_rate'])}%，主车道左侧标线合格率为{cls._g03_num(overall['left_rate'])}%，"
                       f"主车道右侧标线合格率为{cls._g03_num(overall['right_rate'])}%。")
        # G1（用户第 3 轮）：① 总体情况 下只保留合格率/缺失率句——模板该节无统计口径段，删除原第 2 段。
        # P1b：km 改由 gd_mileage 从清单表行求和（mileage 键 = 分组键），不再按检测单元数算。
        def _marking_group(subset, **km_key):
            # P1n-1：class_prefix 同时筛「检测记录子集」与里程口径，速率不再照抄全市值
            subset_units = GuangdongStatistics.marking_units(gd_class_subset(subset, km_key.get("class_prefix")))
            overall_sub = GuangdongStatistics.marking_overall(subset_units)
            return {"km": gd_mileage(inspection, category=category, **km_key),
                    "overall": overall_sub["overall_rate"], "left": overall_sub["left_rate"],
                    "right": overall_sub["right_rate"]}

        if category == "高速公路":
            # brief-v3 §1：① 表 = 表头 + 4 行（全省抽检高速 / XX市抽检高速 / 非省交通集团 / 省交通集团）。
            # 第 1 行是全省行，4 个数据格全部留空（本地无全省抽检记录，不填占位也不填「—」）。
            # 模板表2 写「其中非省交通集团」、表11/14/17 写「非省交通集团」；统一取后者（3/5 张如此）。
            owner_map = gd_owner_map(inspection)
            groups = [("全省抽检高速", None),
                      (f"{RouteCategoryIndex.city_display(city)}抽检高速", _marking_group(rows)),
                      ("非省交通集团", _marking_group(
                          gd_owner_rows(rows, owner_map, GD_OWNER_NON_PROVINCIAL), owner=GD_OWNER_NON_PROVINCIAL)),
                      ("省交通集团", _marking_group(
                          gd_owner_rows(rows, owner_map, GD_OWNER_PROVINCIAL), owner=GD_OWNER_PROVINCIAL))]
            first_col = "管理单位"
            caption = "高速公路各管理单位标线逆反射亮度系数合格率汇总表"
            figure_title = "抽检高速标线逆反射亮度系数合格率对比图"
        else:
            groups = [("全省抽检普通国省道", None), ("全省抽检普通国道", None), ("全省抽检普通省道", None),
                      (f"{RouteCategoryIndex.city_display(city)}抽检普通国省道", _marking_group(rows)),
                      (f"{RouteCategoryIndex.city_display(city)}抽检普通国道", _marking_group(rows, class_prefix="G")),
                      (f"{RouteCategoryIndex.city_display(city)}抽检普通省道", _marking_group(rows, class_prefix="S"))]
            first_col = "道路类别"
            caption = "抽检普通国省道标线逆反射亮度系数合格率汇总表"
            figure_title = "抽检普通国省道标线逆反射亮度系数合格率对比图"
        km_texts = gd_km_row_texts(groups)   # P1j：高速三行整数恒等（最大余数法）
        table([first_col, "抽检里程(km)", "总体合格率(%)", "主车道左侧合格率(%)", "主车道右侧合格率(%)"],
              [[label, "", "", "", ""] if stats is None else
               [label, km_texts[idx], cls._g03_num(stats["overall"]), cls._g03_num(stats["left"]),
                cls._g03_num(stats["right"])]
               for idx, (label, stats) in enumerate(groups)],
              caption)
        # P1i：①表「全市 = 非集团+集团」的核定说明必须进产物正文（表下段落），不能只打日志。
        if category == "高速公路" and (note := gd_owner_gap_note(inspection, category)):
            cls._body(doc, note)
        chart = cls._g03_chart(base, chart_prefix, f"marking_units_{category}", [label for label, _ in groups],
                               [("主车道左侧标线", [stats["left"] if stats else None for _, stats in groups], GD03_COLORS[0]),
                                ("主车道右侧标线", [stats["right"] if stats else None for _, stats in groups], GD03_COLORS[1]),
                                ("总体", [stats["overall"] if stats else None for _, stats in groups], GD03_COLORS[2])])
        if chart:
            figure(chart, figure_title)

        heading("②各抽检路段情况")
        cls._body(doc, gd_route_rate_sentence(category, "标线逆反射亮度系数总体合格率", route_units, route_names,
                                             lambda row: row["overall_rate"], city))
        headers = ["路线编号", "路线名称", "管养单位", "起止桩号", "里程(km)", "总体合格率%", "左侧合格率%", "右侧合格率%"]
        if category != "高速公路":
            # 模板-1 分支差异逐字照抄：普通分支第 4 列 `检测范围`，且第 6–8 列列头带括号
            headers[3] = "检测范围"
            headers[5:] = ["总体合格率(%)", "左侧合格率(%)", "右侧合格率(%)"]
        route_rows = [[row["route"], route_names.get(row["route"], "—"), manager_display(row.get("manager"), city),
                       cls._gd03_range(row.get("start_m"), row.get("end_m")),
                       gd_km_text(gd_mileage(inspection, category=category, route=row["route"],
                                             direction=row.get("direction"), manager=row.get("manager"))),
                       cls._g03_num(row["overall_rate"]), cls._g03_num(row["left_rate"]), cls._g03_num(row["right_rate"])]
                      for row in route_units]
        # D2：模板该表尾行为合计行（前 4 列并成一格，文本 `XX市合计`）
        _route_totals = _marking_group(rows)
        route_rows.append([f"{city}合计", "", "", "", gd_km_text(_route_totals["km"]), cls._g03_num(_route_totals["overall"]),
                           cls._g03_num(_route_totals["left"]), cls._g03_num(_route_totals["right"])])
        table(headers, route_rows,
              "高速公路各路段公司标线逆反射亮度系数合格率汇总表" if category == "高速公路"
              else "普通国省道各抽检路段标线逆反射亮度系数合格率汇总表",
              tail_merge={0: 4})
        chart = cls._g03_chart(base, chart_prefix, f"marking_route_{category}",
                               [f"{row['route']}{manager_display(row.get('manager'), city)}" for row in route_units],
                               [("左侧标线", [row["left_rate"] for row in route_units], GD03_COLORS[0]),
                                ("右侧标线", [row["right_rate"] for row in route_units], GD03_COLORS[1]),
                                ("总体", [row["overall_rate"] for row in route_units], GD03_COLORS[2])])
        if chart:
            figure(chart, "高速公路各路段公司标线逆反射亮度系数合格率对比图" if category == "高速公路"
                   else "普通国省道各抽检路段标线逆反射亮度系数合格率对比图")

        if category == "高速公路":
            heading("③典型状况不佳路段及原因分析")
            cls._body(doc, "根据抽检路段每公里合格率评定明细结果，筛查左侧或右侧标线合格率小于20%的路段，并对其中长连续的标线逆反射亮度系数典型状况不佳路段进行统计分析。")
            typical = GuangdongStatistics.marking_typical_segments(rows, limit=3)
            if not typical:
                cls._body(doc, "未识别到标线合格率明显偏低的典型路段。")
            if typical:
                # P1a（brief-v3 §5）：③表逐公里列明，一个典型段每个有检测数据的公里一行，
                # 各行取该公里原始 100 m 单元复算值，不用段平均填充。
                table(["路线编号", "管养单位", "起止桩号", "总体合格率(%)", "左侧合格率(%)", "右侧合格率(%)"],
                      [[item["route"], cls._g03_km_manager(rows, item, km["km"], city),
                        cls._g03_km_range(km["km"]), cls._g03_num(km["overall_rate"]),
                        cls._g03_num(km["left_rate"]), cls._g03_num(km["right_rate"])]
                       for item in typical for km in item["km_items"]],
                      f"{category}标线逆反射亮度系数不佳路段汇总表")
            for index, item in enumerate(typical, 1):
                # P1b P3：模板逐字为 `1）S51下行K162+000-K142+000段`，`下行` 后无空格。
                seg = f"{item['route']}{item['direction']}{cls._gd03_range(item['start_m'], item['end_m'])}"
                cls._typical_segment_heading(doc, f"{index}）{seg}段")
                cls._body(doc, f"该路段由{manager_display(cls._manager_for_km(rows, item, city), city)}管养，"
                               f"标线100米计算单元总体合格率仅{cls._g03_rate(item['average'])}，"
                               f"左侧{cls._g03_rate(item['left_rate'])}、右侧{cls._g03_rate(item['right_rate'])}。"
                               + gd_km_detail(item))
                cls._body(doc, "原因分析：该路段标线逆反射亮度系数偏低，主要受路面标线自然磨耗、车辆轮迹带污染、"
                               "重载交通渠化作用及标线施划年限较长等因素影响，需现场复核标线磨损与污染状况。")
                # D16/Q1：模板无「典型状况不佳路段标线合格率分布图」，逐公里明细入工作簿「标线典型路段逐公里」表
            # D16/Q1：模板无「长连续不合格路段清单」表与「长连续路段分布图」，正文只留结论，
            # 明细与分档数据写进交安设施统计图表工作簿「标线长连续不合格」表。
            runs = GuangdongStatistics.marking_long_runs(rows)
            if runs:
                cls._body(doc, f"共识别长连续不合格路段{len(runs)}处，最长{max(row['length_km'] for row in runs):.2f}公里，"
                               "建议纳入现场复核与处治计划，明细见交安设施统计图表工作簿。")
            else:
                cls._body(doc, "未识别到100米基准下3公里窗口不合格率超过80%的长连续路段。")
        else:
            heading("③典型状况不佳路段及原因分析")
            cls._body(doc, "根据抽检路段每公里合格率评定明细结果，筛查左侧或右侧标线合格率小于20%的路段，并对其中长连续的标线逆反射亮度系数典型状况不佳路段进行统计分析。")
            typical = GuangdongStatistics.marking_typical_segments(rows, limit=3)
            if not typical:
                cls._body(doc, "未识别到标线合格率明显偏低的典型路段。")
            table(["路线编号", "管养单位", "起止桩号", "总体合格率(%)", "左侧合格率(%)", "右侧合格率(%)"],
                  [[item["route"], cls._g03_km_manager(rows, item, km["km"], city),
                    cls._g03_km_range(km["km"]), cls._g03_num(km["overall_rate"]),
                    cls._g03_num(km["left_rate"]), cls._g03_num(km["right_rate"])]
                   for item in typical for km in item["km_items"]],
                  f"{category}标线逆反射亮度系数不佳路段汇总表")
            for index, item in enumerate(typical, 1):
                # P1b P3：模板逐字为 `1）S51下行K162+000-K142+000段`，`下行` 后无空格。
                seg = f"{item['route']}{item['direction']}{cls._gd03_range(item['start_m'], item['end_m'])}"
                cls._typical_segment_heading(doc, f"{index}）{seg}段")
                cls._body(doc, f"该路段由{manager_display(cls._manager_for_km(rows, item, city), city)}管养，"
                               f"标线100米计算单元总体合格率仅{cls._g03_rate(item['average'])}，"
                               f"左侧{cls._g03_rate(item['left_rate'])}、右侧{cls._g03_rate(item['right_rate'])}。"
                               + gd_km_detail(item))
                cls._body(doc, "原因分析：标线逆反射亮度系数偏低主要与标线磨耗、污染及施划年限有关，建议现场复核。")
            # D16/Q1：同高速分支，删表删图，数据入工作簿「标线长连续不合格」表
            runs = GuangdongStatistics.marking_long_runs(rows)
            if runs:
                cls._body(doc, f"共识别长连续不合格路段{len(runs)}处，建议纳入现场复核与处治计划，明细见交安设施统计图表工作簿。")
            else:
                cls._body(doc, "未识别到100米基准下3公里窗口不合格率超过80%的长连续路段。")

    @classmethod
    def _height_gd03_section(cls, doc, category, rows, table, figure, base, chart_prefix, city="", route_names=None,
                             inspection=None):
        heading = lambda text: cls._heading(doc, text, 5)
        route_names = route_names or {}
        if not [row for row in rows if _float(row.get("height")) is not None]:
            heading("①总体情况")
            cls._body(doc, f"本次未读取到{category}波形梁护栏中心高度有效检测数据，该指标暂不评价。")
            return
        stats = GuangdongStatistics.height_overall(rows)
        route_units = GuangdongStatistics.height_route_units(rows)
        heading("①总体情况")
        cls._body(doc, f"{city}抽检{'高速公路' if category == '高速公路' else '普通国省道'}波形梁护栏中心高度总体合格率为"
                       f"{cls._g03_num(stats['rate'])}%，其中两波护栏合格率{cls._g03_num(stats['二波']['rate'])}%，"
                       f"三波护栏合格率{cls._g03_num(stats['三波']['rate'])}%。")
        # G1（用户第 3 轮）：① 总体情况 下只保留合格率句（删除原有效检测点/判定标准段）。
        # P1b：km 改由 gd_mileage 从清单表行求和（分组键 = 里程分组键），不再按覆盖公里桶数算。
        def _height_group(subset, **km_key):
            # P1n-1：同 _marking_group，速率按道路类别子集算
            sub_stats = GuangdongStatistics.height_overall(gd_class_subset(subset, km_key.get("class_prefix")))
            return {"km": gd_mileage(inspection, category=category, **km_key), "overall": sub_stats["rate"],
                    "two": sub_stats["二波"]["rate"], "three": sub_stats["三波"]["rate"]}

        if category == "高速公路":
            owner_map = gd_owner_map(inspection)
            groups = [("全省抽检高速", None),
                      (f"{RouteCategoryIndex.city_display(city)}抽检高速", _height_group(rows)),
                      ("非省交通集团", _height_group(
                          gd_owner_rows(rows, owner_map, GD_OWNER_NON_PROVINCIAL), owner=GD_OWNER_NON_PROVINCIAL)),
                      ("省交通集团", _height_group(
                          gd_owner_rows(rows, owner_map, GD_OWNER_PROVINCIAL), owner=GD_OWNER_PROVINCIAL))]
            first_col = "管理单位"
            caption = "高速公路各管理单位护栏横梁中心高度合格率汇总表"
            figure_title = "抽检高速波形梁护栏中心高度合格率对比图"
        else:
            groups = [("全省抽检普通国省道", None), ("全省抽检普通国道", None), ("全省抽检普通省道", None),
                      (f"{RouteCategoryIndex.city_display(city)}抽检普通国省道", _height_group(rows)),
                      (f"{RouteCategoryIndex.city_display(city)}抽检普通国道", _height_group(rows, class_prefix="G")),
                      (f"{RouteCategoryIndex.city_display(city)}抽检普通省道", _height_group(rows, class_prefix="S"))]
            first_col = "类别"
            caption = "抽检普通国省道护栏横梁中心高度合格率汇总表"
            figure_title = "抽检普通国省道波形梁护栏横梁中心高度合格率对比图"
        km_texts = gd_km_row_texts(groups)   # P1j：高速三行整数恒等（最大余数法）
        table([first_col, "抽检里程(km)", "总体合格率(%)", "两波合格率(%)", "三波合格率(%)"],
              [[label, "", "", "", ""] if stats is None else
               [label, km_texts[idx], cls._g03_num(stats["overall"]), cls._g03_num(stats["two"]),
                cls._g03_num(stats["three"])]
               for idx, (label, stats) in enumerate(groups)],
              caption)
        if category == "高速公路" and (note := gd_owner_gap_note(inspection, category)):
            cls._body(doc, note)
        chart = cls._g03_chart(base, chart_prefix, f"height_units_{category}", [label for label, _ in groups],
                               [("两波护栏", [stats["two"] if stats else None for _, stats in groups], GD03_COLORS[0]),
                                ("三波护栏", [stats["three"] if stats else None for _, stats in groups], GD03_COLORS[1]),
                                ("总体", [stats["overall"] if stats else None for _, stats in groups], GD03_COLORS[2])])
        if chart:
            figure(chart, figure_title)

        heading("②各抽检路段情况")
        cls._body(doc, gd_route_rate_sentence(category, "波形梁护栏中心高度总体合格率", route_units, route_names,
                                             lambda row: row["rate"], city))
        height_headers = ["路线编号", "路线名称", "管养单位", "起止桩号", "里程(km)", "总体合格率(%)", "两波合格率(%)", "三波合格率(%)"]
        if category != "高速公路":
            # 模板-1 分支差异逐字照抄：普通分支该三列列头不带括号
            height_headers[5:] = ["总体合格率%", "两波合格率%", "三波合格率%"]
        height_rows = [[row["route"], route_names.get(row["route"], "—"), manager_display(row.get("manager"), city),
                        cls._gd03_range(row.get("start_m"), row.get("end_m")),
                        gd_km_text(gd_mileage(inspection, category=category, route=row["route"],
                                              direction=row.get("direction"), manager=row.get("manager"))),
                        cls._g03_num(row["rate"]), cls._g03_num(row["二波"]["rate"]), cls._g03_num(row["三波"]["rate"])]
                       for row in route_units]
        # D2：同标线②，尾行为合计行（前 4 列合并）
        _height_totals = _height_group(rows)
        height_rows.append([f"{city}合计", "", "", "", gd_km_text(_height_totals["km"]), cls._g03_num(_height_totals["overall"]),
                            cls._g03_num(_height_totals["two"]), cls._g03_num(_height_totals["three"])])
        table(height_headers, height_rows,
              "高速公路各路段公司波形梁护栏中心高度合格率汇总表" if category == "高速公路"
              else "普通国省道各抽检路段波形梁护栏中心高度合格率汇总表",
              tail_merge={0: 4})
        chart = cls._g03_chart(base, chart_prefix, f"height_route_{category}",
                               [f"{row['route']}{manager_display(row.get('manager'), city)}" for row in route_units],
                               [("二波护栏", [row["二波"]["rate"] for row in route_units], GD03_COLORS[0]),
                                ("三波护栏", [row["三波"]["rate"] for row in route_units], GD03_COLORS[1]),
                                ("总体", [row["rate"] for row in route_units], GD03_COLORS[2])])
        if chart:
            figure(chart, "高速公路各路段公司波形梁护栏中心高度合格率对比" if category == "高速公路"
                   else "普通国省道波形梁护栏中心高度合格率对比图")

        if category == "高速公路":
            heading("③典型状况不佳路段及原因分析")
            cls._body(doc, "根据抽检路段每公里合格率评定明细结果，筛查整体波形梁护栏中心高度合格率小于50%的路段，并对其中长连续的波形梁护栏中心高度典型状况不佳路段进行统计分析。")
            typical = GuangdongStatistics.height_typical_segments(rows, limit=3)
            if not typical:
                cls._body(doc, "未识别到波形梁护栏中心高度合格率明显偏低的典型路段。")
            if typical:
                # P1a（brief-v3 §5）：③表逐公里列明，两波/三波按该公里 raw 明细分别重算。
                table(["路线编号", "管养单位", "起止桩号", "总体合格率(%)", "两波合格率(%)", "三波合格率(%)"],
                      [[item["route"], cls._g03_km_manager(rows, item, km["km"], city),
                        cls._g03_km_range(km["km"]), cls._g03_cell(km["rate"]),
                        *[cls._g03_cell(rate) for rate in cls._g03_km_kinds(rows, item, km["km"])]]
                       for item in typical for km in item["km_items"]],
                      f"{category}波形梁护栏中心高度不佳路段汇总表")
            for index, item in enumerate(typical, 1):
                # P1b P3：模板逐字为 `1）S51下行K162+000-K142+000段`，`下行` 后无空格。
                seg = f"{item['route']}{item['direction']}{cls._gd03_range(item['start_m'], item['end_m'])}"
                cls._typical_segment_heading(doc, f"{index}）{seg}段")
                cls._body(doc, f"{cls._gd03_range(item.get('start_m'), item.get('end_m'))}区间合格率普遍低于"
                               f"{cls._g03_rate(item.get('max_rate'))}，"
                               + (gd_km_detail(item, 0.35, 3) or
                                  f"逐公里看，最低单公里合格率{cls._g03_rate(item['min_rate'])}。")
                               + gd_gtype_note(item.get("gtype")))
                cls._body(doc, "原因分析：护栏中心高度偏差主要受路面加铺、路缘石抬高、路基沉陷及立柱埋深不足等因素影响，"
                               "建议结合路面结构与立柱埋深现场核查后实施抬升、调整或更换。")
            # D16/Q1：模板无「超10cm长连续路段清单」表与「最低公里段分布图」，
            # 正文只留结论，明细与逐公里分档入工作簿「高度超10cm长连续」「高度最低公里段」表
            runs = GuangdongStatistics.height_over10_runs(rows)
            if runs:
                cls._body(doc, f"共识别整公里偏差超过10 cm占比大于5%的长连续路段{len(runs)}处，"
                               "建议结合路面结构与立柱埋深核查后处治，明细见交安设施统计图表工作簿。")
            else:
                cls._body(doc, "未识别到整公里偏差超过10 cm占比大于5%的长连续路段。")
        else:
            heading("③典型状况不佳路段及原因分析")
            cls._body(doc, "根据抽检路段每公里合格率评定明细结果，筛查整体波形梁护栏中心高度合格率小于50%的路段，并对其中长连续的波形梁护栏中心高度典型状况不佳路段进行统计分析。")
            typical = GuangdongStatistics.height_typical_segments(rows, limit=3)
            if not typical:
                cls._body(doc, "未识别到波形梁护栏中心高度合格率明显偏低的典型路段。")
            table(["路线编号", "管养单位", "起止桩号", "总体合格率%", "两波合格率%", "三波合格率%"],
                  [[item["route"], cls._g03_km_manager(rows, item, km["km"], city),
                    cls._g03_km_range(km["km"]), cls._g03_cell(km["rate"]),
                    *[cls._g03_cell(rate) for rate in cls._g03_km_kinds(rows, item, km["km"])]]
                   for item in typical for km in item["km_items"]],
                  f"{category}波形梁护栏中心高度不佳路段汇总表")
            for index, item in enumerate(typical, 1):
                # P1b P3：模板逐字为 `1）S51下行K162+000-K142+000段`，`下行` 后无空格。
                seg = f"{item['route']}{item['direction']}{cls._gd03_range(item['start_m'], item['end_m'])}"
                cls._typical_segment_heading(doc, f"{index}）{seg}段")
                cls._body(doc, f"{cls._gd03_range(item.get('start_m'), item.get('end_m'))}区间合格率普遍低于"
                               f"{cls._g03_rate(item.get('max_rate'))}，"
                               + (gd_km_detail(item, 0.35, 3) or
                                  f"逐公里看，最低单公里合格率{cls._g03_rate(item['min_rate'])}。")
                               + gd_gtype_note(item.get("gtype")))
                cls._body(doc, "原因分析：护栏中心高度偏差与路面加铺、路缘石及立柱埋深等因素有关，建议现场核查。")
            # D16/Q1：同高速分支
            runs = GuangdongStatistics.height_over10_runs(rows)
            if runs:
                cls._body(doc, f"共识别整公里偏差超过10 cm占比大于5%的长连续路段{len(runs)}处，"
                               "建议现场核查后处治，明细见交安设施统计图表工作簿。")
            else:
                cls._body(doc, "未识别到整公里偏差超过10 cm占比大于5%的长连续路段。")

    @classmethod
    def _bolt_gd03_section(cls, doc, category, rows, table, figure, base, chart_prefix, city="", route_names=None,
                           inspection=None):
        heading = lambda text: cls._heading(doc, text, 5)
        route_names = route_names or {}
        if not rows:
            heading("①总体情况")
            cls._body(doc, f"本次未读取到{category}波形梁护栏螺栓有效检测数据，该指标暂不评价。")
            return
        stats = GuangdongStatistics.bolt_overall(rows)
        route_units = GuangdongStatistics.bolt_route_units(rows)
        heading("①总体情况")
        cls._body(doc, f"{city}抽检{'高速公路' if category == '高速公路' else '普通国省道'}路侧波形梁护栏螺栓总体缺失率为"
                       f"{cls._g03_num(stats['rate'], 2)}%，其中拼接螺栓缺失率{cls._g03_num(stats['splice_rate'], 2)}%，"
                       f"连接螺栓缺失率{cls._g03_num(stats['conn_rate'], 2)}%。")
        # G1（用户第 3 轮）：① 总体情况 下只保留缺失率句（删除原应安装/缺失数量口径段）。
        # P1b：km 改由 gd_mileage 从清单表行求和（分组键 = 里程分组键），不再按覆盖公里桶数算。
        def _bolt_group(subset, **km_key):
            # P1n-1：同 _marking_group，速率按道路类别子集算
            sub_stats = GuangdongStatistics.bolt_overall(gd_class_subset(subset, km_key.get("class_prefix")))
            return {"km": gd_mileage(inspection, category=category, **km_key), "overall": sub_stats["rate"],
                    "splice": sub_stats["splice_rate"], "conn": sub_stats["conn_rate"]}

        if category == "高速公路":
            owner_map = gd_owner_map(inspection)
            groups = [("全省抽检高速", None),
                      (f"{RouteCategoryIndex.city_display(city)}抽检高速", _bolt_group(rows)),
                      ("非省交通集团", _bolt_group(
                          gd_owner_rows(rows, owner_map, GD_OWNER_NON_PROVINCIAL), owner=GD_OWNER_NON_PROVINCIAL)),
                      ("省交通集团", _bolt_group(
                          gd_owner_rows(rows, owner_map, GD_OWNER_PROVINCIAL), owner=GD_OWNER_PROVINCIAL))]
            first_col = "管理单位"
            caption = "高速公路各管理单位波形梁护栏螺栓缺失率汇总表"
            figure_title = "高速公路各管理单位波形梁护栏螺栓缺失率对比图"
        else:
            groups = [("全省抽检普通国省道", None), ("全省抽检普通国道", None), ("全省抽检普通省道", None),
                      (f"{RouteCategoryIndex.city_display(city)}抽检普通国省道", _bolt_group(rows)),
                      (f"{RouteCategoryIndex.city_display(city)}抽检普通国道", _bolt_group(rows, class_prefix="G")),
                      (f"{RouteCategoryIndex.city_display(city)}抽检普通省道", _bolt_group(rows, class_prefix="S"))]
            first_col = "类别"
            caption = "普通国省道波形梁护栏螺栓缺失率汇总表"
            figure_title = "抽检普通国省道波形梁护栏螺栓缺失率对比图"
        km_texts = gd_km_row_texts(groups)   # P1j：高速三行整数恒等（最大余数法）
        table([first_col, "抽检里程(km)", "总体缺失率(%)", "拼接螺栓缺失率(%)", "连接螺栓缺失率(%)"],
              [[label, "", "", "", ""] if stats is None else
               [label, km_texts[idx], cls._g03_num(stats["overall"], 2), cls._g03_num(stats["splice"], 2),
                cls._g03_num(stats["conn"], 2)]
               for idx, (label, stats) in enumerate(groups)],
              caption)
        if category == "高速公路" and (note := gd_owner_gap_note(inspection, category)):
            cls._body(doc, note)
        chart = cls._g03_chart(base, chart_prefix, f"bolt_units_{category}", [label for label, _ in groups],
                               [("拼接", [stats["splice"] if stats else None for _, stats in groups], GD03_COLORS[0]),
                                ("连接", [stats["conn"] if stats else None for _, stats in groups], GD03_COLORS[1]),
                                ("总体", [stats["overall"] if stats else None for _, stats in groups], GD03_COLORS[2])],
                               ylabel="螺栓缺失率（%）", value_fmt="%.2f",
                               ylim=(0, max(10.0, max((stats["overall"] or 0) * 100 * 1.1
                                                       for _, stats in groups if stats))))
        if chart:
            figure(chart, figure_title)

        # 模板-1 普通分支该处写作 `②各路段抽检情况`（逐字照抄，不做统一）
        heading("②各路段抽检情况" if category == "普通国省道" else "②各抽检路段情况")
        cls._body(doc, gd_route_missing_sentence(category, "波形梁护栏螺栓缺失率", route_units, route_names,
                                                 lambda row: row["rate"], city))
        bolt_headers = ["路线编号", "路线名称", "管养单位", "起止桩号", "里程（Km）", "总体缺失率%", "拼接缺失率%", "连接缺失率%"]
        if category != "高速公路":
            # 模板-1 分支差异逐字照抄：普通分支为全角括号版（`里程（Km)` 右括号半角，照抄模板）
            bolt_headers[4:] = ["里程（Km)", "总体缺失率（%）", "拼接缺失率（%）", "连接缺失率（%）"]
        bolt_rows = [[row["route"], route_names.get(row["route"], "—"), manager_display(row.get("manager"), city),
                      cls._gd03_range(row.get("start_m"), row.get("end_m")),
                      gd_km_text(gd_mileage(inspection, category=category, route=row["route"],
                                            direction=row.get("direction"), manager=row.get("manager"))),
                      cls._g03_num(row["rate"], 2), cls._g03_num(row["splice_rate"], 2), cls._g03_num(row["conn_rate"], 2)]
                     for row in route_units]
        # D2：同标线②，尾行为合计行（前 4 列合并）
        _bolt_totals = _bolt_group(rows)
        bolt_rows.append([f"{city}合计", "", "", "", gd_km_text(_bolt_totals["km"]), cls._g03_num(_bolt_totals["overall"], 2),
                          cls._g03_num(_bolt_totals["splice"], 2), cls._g03_num(_bolt_totals["conn"], 2)])
        table(bolt_headers, bolt_rows,
              "高速公路各路段公司螺栓缺失率汇总表" if category == "高速公路"
              else "普通国省道各路段公司螺栓缺失率汇总表",
              tail_merge={0: 4})
        chart = cls._g03_chart(base, chart_prefix, f"bolt_route_{category}",
                               [f"{row['route']}{manager_display(row.get('manager'), city)}" for row in route_units],
                               [("拼接", [row["splice_rate"] for row in route_units], GD03_COLORS[0]),
                                ("连接", [row["conn_rate"] for row in route_units], GD03_COLORS[1]),
                                ("总体", [row["rate"] for row in route_units], GD03_COLORS[2])],
                               ylabel="螺栓缺失率（%）", ylim=(0, max(10.0, max((row["rate"] or 0) * 100 * 1.1 for row in route_units))), value_fmt="%.2f")
        if chart:
            figure(chart, "高速公路各路段公司波形梁护栏螺栓缺失率对比图" if category == "高速公路"
                   else "普通国省道各路段波形梁护栏螺栓缺失率对比图")

        if category == "高速公路":
            heading("③典型状况不佳路段及原因分析")
            cls._body(doc, "根据抽检路段螺栓缺失率评定明细结果，筛查“将同公里路段内相邻严重缺失桩号间距不超过20m的归为同一区段，且该区段内包含至少3处”或“整公里螺栓缺失率大于3%”的路段，并对其中螺栓缺失典型状况不佳路段进行统计分析。")
            typical = GuangdongStatistics.bolt_typical_segments(rows, limit=3)
            if not typical:
                cls._body(doc, "未识别到螺栓缺失率明显偏高的典型路段。")
            for index, item in enumerate(typical, 1):
                # P1b P3：模板逐字为 `1）S51下行K162+000-K142+000段`，`下行` 后无空格。
                seg = f"{item['route']}{item['direction']}{cls._gd03_range(item['start_m'], item['end_m'])}"
                cls._typical_segment_heading(doc, f"{index}）{seg}段")
                cls._body(doc, f"该路段螺栓缺失率均值{cls._g03_rate(item['average'], 2)}，"
                               f"最高单公里缺失率{cls._g03_rate(item['max_rate'], 2)}，"
                               f"缺失螺栓{cls._g03_int(item['missing'])}颗，单处严重缺失{cls._g03_int(item['severe'])}处。")
                cls._body(doc, "原因分析：螺栓缺失多与交通荷载振动、施工遗留、养护更换不及时及连接件锈蚀有关，"
                               "缺失会削弱护栏整体连接强度，建议优先补齐并检查梁板搭接与连接件状态。")
            clusters = GuangdongStatistics.bolt_severe_clusters(rows)
            runs = GuangdongStatistics.bolt_over5_runs(rows)
            if clusters:
                table(["路线", "管养单位", "起止桩号", "缺失处数", "单处最大缺失(颗/柱)"],
                      [[row["route"], cls._manager_for_km(rows, row, city), cls._gd03_range(row["start_m"], row["end_m"]),
                        row["points"], cls._g03_int(row["max_missing"])] for row in clusters],
                      "高速公路螺栓缺失不佳长连续段清单")
            if runs:
                # P1a（brief-v3 §5）：整公里清单逐公里列明；上方严重缺失簇为物理区段，不强拆。
                table(["路线", "管养单位", "起止桩号", "总体缺失率%", "拼接缺失率%", "连接缺失率%"],
                      [[row["route"], cls._g03_km_manager(rows, row, km["km"], city),
                        cls._g03_km_range(km["km"]), cls._g03_num(km["total_rate"], 2),
                        cls._g03_num(km["splice_rate"], 2), cls._g03_num(km["conn_rate"], 2)]
                       for row in runs for km in row["km_items"]],
                      "高速公路整公里螺栓缺失率大于3%路段清单")
            if not clusters and not runs:
                cls._body(doc, "未识别到连续严重缺失或整公里缺失率大于3%的长连续路段。")
            # D16/Q1：模板无「单处严重缺失公里段分布图」，逐公里严重缺失数据入工作簿「螺栓严重缺失公里段」表
        else:
            heading("③典型状况不佳路段及原因分析")
            cls._body(doc, f'按“同公里段内相邻严重缺失桩号距离≤20m 视为同一段、段内连续≥3处”以及“整公里螺栓缺失率大于3%”两个并列口径，抽检的{city}普通国省道存在如下典型状况不佳路段：')
            typical = GuangdongStatistics.bolt_typical_segments(rows, limit=3)
            if not typical:
                cls._body(doc, "未识别到螺栓缺失率明显偏高的典型路段。")
            for index, item in enumerate(typical, 1):
                # P1b P3：模板逐字为 `1）S51下行K162+000-K142+000段`，`下行` 后无空格。
                seg = f"{item['route']}{item['direction']}{cls._gd03_range(item['start_m'], item['end_m'])}"
                cls._typical_segment_heading(doc, f"{index}）{seg}段")
                cls._body(doc, f"该路段螺栓缺失率均值{cls._g03_rate(item['average'], 2)}，"
                               f"缺失螺栓{cls._g03_int(item['missing'])}颗。")
                cls._body(doc, "原因分析：螺栓缺失与交通荷载、施工遗留及养护更换不及时有关，建议优先补齐。")
            clusters = GuangdongStatistics.bolt_severe_clusters(rows)
            runs = GuangdongStatistics.bolt_over5_runs(rows)
            if clusters:
                table(["路线", "路线名称", "管养单位", "起止桩号", "跨度(m)", "连续处数", "单处最大缺失(颗/柱)"],
                      [[row["route"], route_names.get(row["route"], "—"), cls._manager_for_km(rows, row, city),
                        cls._gd03_range(row["start_m"], row["end_m"]), cls._g03_int(row["span_m"]), row["points"],
                        cls._g03_int(row["max_missing"])] for row in clusters],
                      "普通国省道螺栓缺失不佳长连续段清单")
            if runs:
                # P1a（brief-v3 §5）：整公里清单逐公里列明；上方严重缺失簇为物理区段，不强拆。
                table(["路线", "路线名称", "管养单位", "起止桩号", "总体缺失率%", "拼接缺失率%", "连接缺失率%"],
                      [[row["route"], route_names.get(row["route"], "—"),
                        cls._g03_km_manager(rows, row, km["km"], city), cls._g03_km_range(km["km"]),
                        cls._g03_num(km["total_rate"], 2), cls._g03_num(km["splice_rate"], 2),
                        cls._g03_num(km["conn_rate"], 2)]
                       for row in runs for km in row["km_items"]],
                      "普通国省道整公里螺栓缺失率大于3%严重路段清单")
            if not clusters and not runs:
                cls._body(doc, "未识别到连续严重缺失或整公里缺失率大于3%的长连续路段。")
            # D16/Q1：同高速分支

    @classmethod
    def _comparison_gd03_section(cls, doc, category, detail, thresholds, table, marking_rows=None):
        """人工复核对比情况（对标附件）：说明段 + 三张人工复核对比明细表，不带小节编号。"""
        groups = {key: [row for row in detail if row.get("indicator") == key] for key in ("marking", "height", "bolt")}
        # P1-1：源数据无标线颜色字段，按该行匹配到的自动化单元「目标值」推导（80→白色、50→黄色）
        marking_rows = marking_rows or []

        def _num(value, digits=2, sign=False):
            if value is None: return "—"
            return f"{value:+.{digits}f}" if sign else f"{value:.{digits}f}"

        def _int(value):
            return "—" if value is None else f"{int(value)}"

        def _gtype_short(value):
            text = str(value or "").strip()
            return (text[:-2] if text.endswith("护栏") else text) or "—"

        def _height_grade(gtype):
            """P3-1：合格值按模板带公差（600±20mm / 697±20mm）；标准值与判定共用 gd_height_standard，禁止第二套。"""
            standard = gd_height_standard(gtype)
            return "—" if standard is None else f"{int(standard)}±{int(GD_HEIGHT_TOLERANCE_MM)}mm"

        def _judge(indicator, value, gtype=""):
            return gd_pass(indicator, value, gtype) or "—"

        def _consistent(row):
            """结果一致性（brief-v3 §7 用户裁决）：标线/高度只比较两侧合格判定是否相同——
            同为合格或同为不合格记一致，一合格一不合格记不一致；任一侧无法判定记 —（不以空值相等视作一致）。
            螺栓对比表没有合格判定列，按两侧缺失数量的差值判定（差值≥阈值记不一致，阈值见界面设置）。"""
            flag = row.get("consistent")
            return "—" if flag is None else ("一致" if flag else "不一致")

        if not detail:
            cls._body(doc, "本次未提供人工复核对比记录。")
            return

        city = next((str(row["city"]).strip() for row in detail if row.get("city")), "")
        counts = {key: len(groups[key]) for key in groups}
        cls._body(doc, f"对比分析{city}{category}抽检路段标线逆反射亮度系数（{counts['marking']}个路段）、"
                       f"波形梁护栏中心高度（{counts['height']}个路段）和螺栓缺失情况（{counts['bolt']}个路段）的"
                       "人工检测复核结果与自动化检测结果，排除极个别路段在人工复核前进行局部养护处治导致结果发生变化的特殊情况，"
                       "总体而言，人工复核数据与自动化检测结果具有良好的一致性。各指标明细对比情况详见下表。")
        # 口径说明：标线/高度一致性 = 两侧合格判定是否相同（不再用测值偏差/允许偏差阈值）；螺栓按缺失数量差值
        bolt_limit = thresholds.get("bolt", GD_BOLT_DIFF_THRESHOLD)
        cls._body(doc, "说明：本表“结果一致性”对标线逆反射亮度系数、波形梁护栏中心高度只比较人工复核与自动化检测两侧的合格判定结果，"
                       "同为合格或同为不合格记为一致，一合格一不合格记为不一致；任一侧合格判定无法确定时记为“—”，不参与一致性占比统计。"
                       "合格判定按设施标准执行：标线逆反射亮度系数以该颜色目标值为限（白色80、黄色50），"
                       "波形梁护栏中心高度按护栏型式取合格值（三波697±20mm、两波/双波600±20mm）。"
                       f"波形梁护栏螺栓缺失对比表没有合格判定列，改按人工与自动化两侧缺失数量的差值判定："
                       f"差值小于{bolt_limit:g}记为一致，不小于该值记为不一致。")

        marking = groups["marking"]
        if marking:
            table(["路线编号", "桩号区段", "标线颜色", "人工检测复核结果", "测值", "合格判定",
                   "自动化检测结果", "测值", "合格判定", "结果一致性"],
                  [[row.get("route") or "—", row.get("segment") or "—", gd_marking_color(row, marking_rows),
                    _num(row.get("manual")), _judge("marking", row.get("manual")),
                    _num(row.get("automatic")), _judge("marking", row.get("automatic")), _consistent(row)]
                   for row in marking],
                  f"{category}标线逆反射亮度系数人工复核对比明细表",
                  merge={3: ("人工检测复核结果", 3), 6: ("自动化检测结果", 3)})
            # brief-v3 §7：源表 备注 为「现场震荡标线」的行，其差异受标线型式影响，写在表下说明段（不加表列）
            spans = [f"{row.get('route') or '—'}{row.get('direction') or ''}{normalize_station_text(row.get('segment'))}"
                     for row in marking if "震荡标线" in str(row.get("remark") or "")]
            if spans:
                cls._body(doc, f"注：以上 {len(spans)} 个路段现场为震荡标线，其人工复核与自动化检测结果的差异"
                               f"受标线型式影响。涉及桩号区段：{'、'.join(spans)}。")
        else:
            cls._body(doc, "本次未提供标线逆反射亮度系数人工复核对比记录。")

        height = groups["height"]
        if height:
            table(["路线编号", "桩号区段", "护栏类型", "合格值(mm)", "人工检测复核结果", "测值", "合格判定",
                   "自动化检测结果", "测值", "合格判定", "结果一致性"],
                  [[row.get("route") or "—", row.get("segment") or "—", _gtype_short(row.get("gtype")),
                    _height_grade(row.get("gtype")), _num(row.get("manual")),
                    _judge("height", row.get("manual"), row.get("gtype")), _num(row.get("automatic")),
                    _judge("height", row.get("automatic"), row.get("gtype")), _consistent(row)]
                   for row in height],
                  f"{category}波形梁护栏中心高度人工复核对比明细表",
                  merge={4: ("人工检测复核结果", 3), 7: ("自动化检测结果", 3)})
        else:
            cls._body(doc, "本次未提供波形梁护栏中心高度人工复核对比记录。")

        bolt = groups["bolt"]
        if bolt:
            table(["路线", "桩号区段", "护栏类型", "人工检测复核结果", "拼接螺栓缺失", "连接螺栓缺失", "总体缺失",
                   "自动化检测结果", "拼接螺栓缺失", "连接螺栓缺失", "总体缺失", "结果一致性"],
                  [[row.get("route") or "—", row.get("segment") or "—", _gtype_short(row.get("gtype")),
                    _int(row.get("msplice")), _int(row.get("mconn")), _int(row.get("manual")),
                    _int(row.get("asplice")), _int(row.get("aconn")), _int(row.get("automatic")), _consistent(row)]
                   for row in bolt],
                  f"{category}路侧波形梁护栏螺栓缺失人工复核对比明细表",
                  merge={3: ("人工检测复核结果", 4), 7: ("自动化检测结果", 4)})
            # 自动化数量多于人工时，按附件口径补充表后原因说明（人工复核前已完成病害整改）
            for row in bolt:
                if (row.get("automatic") or 0) > (row.get("manual") or 0):
                    cls._body(doc, f"{row.get('route')}路线{row.get('segment')}区间，自动化检测结果的病害数量多于人工复核结果的原因，"
                                   "是人工复核前管养单位已针对该区间开展了病害整改作业，可处置的缺陷均已完成补装螺栓处理。")
        else:
            cls._body(doc, "本次未提供波形梁护栏螺栓缺失人工复核对比记录。")

    @classmethod
    def _advice_gd03_section(cls, doc, bundle, table):
        """工作建议（对标附件）：1.重点路段处治建议（高速/普通优先处治路段表+闭环督办）→ 2.迎国评 → 3.养护提升。"""
        weak = bundle.get("weak_segments", [])
        city = ""
        for key in ("marking", "height", "bolt", "comparison_detail"):
            for row in bundle.get(key, []):
                if row.get("city"):
                    city = str(row["city"]).strip()
                    break
            if city:
                break
        cls._body(doc, "结合本次检测评估的结果，针对不同类型问题提出分层分类的改进意见与措施，有效推动整体效能提升与可持续发展。")
        cls._heading(doc, "1.重点路段处治建议（如有）", 3)
        # C3：模板该节仅 1 个正文段（无子标题、无表）；优先处治清单与督办要求并入工作簿「优先处治路段」表（Q1）
        cls._body(doc, GD_KEY_ROUTE_ADVICE)
        if weak:
            cls._body(doc, f"优先处治路段明细见《{city or '本项目'}交安设施统计图表.xlsx》「优先处治路段」工作表。")
        else:
            cls._body(doc, "本次未识别到满足连续3 km标线不合格、护栏中心高度偏差超过10 cm或螺栓缺失率超过5%的典型薄弱路段，建议保持既有养护节奏并加强日常巡查。")
        cls._heading(doc, "2.迎国评工作建议", 3)
        cls._body(doc, "结合2026年国评交安专项要求，提出简易可行、行之有效的迎检攻坚建议。")
        cls._body(doc, "围绕国评交通安全设施TCI指标、护栏防护能力两大核心评分项，抓重点、补短板，高效开展迎检攻坚。")
        # C2：按模板原文补 `（1）` 前缀（模板无句号）
        cls._body(doc, "（1）分类开展显性问题突击整治", keep_next=True)
        # F1：模板-1「（1）分类开展显性问题突击整治」下 3 段原文逐字照抄
        cls._body(doc, "清除标志遮挡：标志板遮挡是高速公路和普通国省道共有的突出问题。组织一次全线排查，集中修剪遮挡交通标志的树木和植被，成本低、见效快。")
        cls._body(doc, "高速公路方面：更换变形护栏板并同步排查补齐护栏缺失螺栓，修复锈蚀、缺损防眩网；对路段旧标线未清除问题开展专项清理。")
        cls._body(doc, "普通国省道方面：优先处置波形梁护栏缺损并补齐护栏缺失螺栓、防护设施防护能力不足隐患，快速翻新缺损严重的标线。")
        cls._body(doc, "（2）做好迎检路段现场排查", keep_next=True)
        cls._body(doc, "国评前组织管养单位开展徒步+车载快速巡检，消除护栏缺口、标志缺失、标线模糊等一眼可见问题。")
        cls._body(doc, "（3）统筹力量，差异化投入", keep_next=True)
        cls._body(doc, "高速公路以管养分公司落实主体责任；普通国省道以区县公路事务中心为责任主体，市级统筹协调养护资金、技术力量，对薄弱区县予以帮扶，避免普通国省道拖整体后腿。")
        cls._heading(doc, "3.养护提升建议", 3)
        # C1：按模板-1 逐字重建（引导段 + 4 个 `（n）` + 5 个 `①②③` 二级标题 + 正文段），替换本工程原有 3 段
        cls._body(doc, GD_MAINTENANCE_LEAD)
        for lead, groups in GD_MAINTENANCE_ADVICE:
            cls._body(doc, lead, keep_next=True)
            for title, paragraphs in groups:
                if title:
                    cls._heading(doc, title, 5)
                for text in paragraphs:
                    cls._body(doc, text)

    @classmethod
    def _comparison_table(cls, doc, detail, thresholds, table_adder=None):
        """人工复核对比：先给判断标准说明，再按 护栏中心高度/螺栓缺失/标线逆反射 生成三张分项对比表。"""
        # M1（brief-v3 §7 用户裁决）：标线/高度一致性只比两侧合格判定；M2：螺栓比缺失数量差值。测值偏差仅作描述性统计
        bolt_limit = thresholds.get("bolt", GD_BOLT_DIFF_THRESHOLD)
        cls._body(doc, "人工复核的“结果一致性”对标线逆反射亮度系数、波形梁护栏中心高度按人工复核与自动化检测两侧的合格判定是否相同判定："
                       "同为合格或同为不合格记为一致，一合格一不合格记为不一致；任一侧合格判定无法确定时记为“—”，不计入一致性占比。"
                       "合格判定标准：标线逆反射亮度系数以该颜色目标值为限（白色80、黄色50），"
                       "波形梁护栏中心高度按护栏型式取合格值（三波697±20mm、两波/双波600±20mm）。"
                       "波形梁护栏螺栓缺失对比表没有合格判定列，按人工与自动化两侧缺失数量的差值判定："
                       f"差值小于{bolt_limit:g}记为一致，不小于该值记为不一致。"
                       "下述平均偏差、偏差极值等数值仅为描述性统计，不参与一致性判定。")
        # 精简版偏差分析：只保留平均偏差+一致性占比，不输出 min～max 偏差范围
        summary = ManualAutoComparator.summarize(detail)
        names = {"marking": "标线逆反射亮度系数", "height": "波形梁护栏中心高度", "bolt": "波形梁护栏螺栓缺失"}
        sentences = []
        for key, name in names.items():
            s = summary.get(key, {})
            if s.get("computable_count"):
                unit = "%" if key in ("marking", "bolt") else " mm"
                sentences.append(f"{name}:自动化检测与人工复核平均偏差{cls._fmt(s.get('average'))}{unit}，一致性占比{cls._pct(s.get('consistency_rate'))}")
        if sentences:
            cls._body(doc, "偏差范围与一致性占比分析：" + "；".join(sentences) + "。")
        fmt = lambda v: "" if v is None else (f"{v:g}" if isinstance(v, float) else str(v))
        def emit(headers, rows, title, merge=None):
            if table_adder is None:
                cls._add_table(doc, headers, rows, merge=merge)
            else:
                table_adder(headers, rows, title, merge)

        height=[r for r in detail if r.get("indicator")=="height"]
        bolt=[r for r in detail if r.get("indicator")=="bolt"]
        marking=[r for r in detail if r.get("indicator")=="marking"]

        emit(["路线","护栏类型","护栏位置","方向","桩号范围","人工复核护栏中心平均高度(mm)","自动化护栏中心高度(mm)","备注"],
             [[r.get("route"),r.get("gtype"),r.get("position"),r.get("direction"),r.get("segment"),fmt(r.get("manual")),fmt(r.get("automatic")),r.get("remark")] for r in height],
             "人工复核护栏中心高度对比表")

        # 螺栓两级表头：主表头横跨2个子列（拼接/连接），无空单元格
        if bolt:
            emit(["路线","护栏类型","护栏位置","方向","桩号范围","人工复核螺栓缺失数量","拼接螺栓缺失数量","连接螺栓缺失数量","自动化螺栓缺失数量","拼接螺栓缺失数量","连接螺栓缺失数量","备注"],
                 [[r.get("route"),r.get("gtype"),r.get("position"),r.get("direction"),r.get("segment"),fmt(r.get("msplice")),fmt(r.get("mconn")),fmt(r.get("asplice")),fmt(r.get("aconn")),r.get("remark")] for r in bolt],
                 "人工复核螺栓缺失对比表", merge={5:("人工复核螺栓缺失数量",3),8:("自动化螺栓缺失数量",3)})

        emit(["路线","标线位置","方向","桩号范围","人工逆反射亮度系数平均值","自动化逆反射亮度系数平均值"],
             [[r.get("route"),r.get("position"),r.get("direction"),r.get("segment"),fmt(r.get("manual")),fmt(r.get("automatic"))] for r in marking],
             "人工复核标线逆反射对比表")

    @staticmethod
    def _set_run_fonts(run, east_asia="仿宋_GB2312", latin="Times New Roman"):
        from backend import minimal_docx
        minimal_docx._set_run_fonts(run, east_asia, latin)

    @classmethod
    def _configure_heading_styles(cls, doc):
        from backend import minimal_docx
        minimal_docx.configure_document(doc, cls._format_config())

    @classmethod
    def _format_existing_headings(cls, doc):
        main_prefix = "五、交通安全设施技术状况检测评价"
        for paragraph in doc.paragraphs:
            if paragraph.text.strip().startswith(main_prefix) and "结果" in paragraph.text:
                for run in paragraph.runs:
                    run.text = run.text.replace("结果", "情况")
            style = paragraph.style.name if paragraph.style is not None else ""
            text = paragraph.text.strip()
            level = cls._outline_level(paragraph)
            is_main = text.startswith(main_prefix)
            is_heading = is_main or level is not None or style.startswith("Heading")
            if not is_heading:
                continue
            level = level or (1 if is_main else 1)
            if level <= 5 and f"Heading {level}" in doc.styles:
                paragraph.style = doc.styles[f"Heading {level}"]
            cls._apply_heading_format(paragraph, level, main_title=is_main)

    @classmethod
    def _format_chart_caption(cls, paragraph):
        from backend import minimal_docx
        format_config = cls._format_config()["caption"]
        minimal_docx._apply_paragraph(paragraph, format_config, clear_indent=True)
        # P2-2：题注 keep_with_next，避免表题与表体被拆到两页
        paragraph.paragraph_format.keep_with_next = True
        for run in paragraph.runs:
            minimal_docx._apply_run(run, format_config)

    @classmethod
    def _format_all_run_fonts(cls, doc):
        from docx.oxml.ns import qn
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from backend import minimal_docx
        format_config = cls._format_config()
        main_prefix = "五、交通安全设施技术状况检测评价"
        for paragraph in doc.paragraphs:
            style = paragraph.style.name if paragraph.style is not None else ""
            text = paragraph.text.strip()
            level = cls._outline_level(paragraph)
            is_main = text.startswith(main_prefix)
            is_heading = is_main or level is not None or style.startswith("Heading")
            if style == "Caption":
                cls._format_chart_caption(paragraph)
            elif style == "TOC Heading":
                minimal_docx._apply_paragraph(paragraph, format_config["heading"]["1"], clear_indent=True)
                for run in paragraph.runs:
                    minimal_docx._apply_run(run, format_config["heading"]["1"])
            elif is_heading:
                cls._apply_heading_format(paragraph, level or 1, main_title=is_main)
            elif paragraph._p.findall(".//" + qn("w:drawing")) or paragraph._p.findall(".//" + qn("w:pict")):
                # 图片段落不是正文：套用正文格式会加回首行缩进并改成两端对齐，把插图推右
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                minimal_docx._clear_indent(paragraph)
            else:
                minimal_docx._apply_paragraph(paragraph, format_config["body"])
                for run in paragraph.runs:
                    minimal_docx._apply_run(run, format_config["body"])
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        minimal_docx._apply_paragraph(paragraph, format_config["table"], clear_indent=True)
                        for run in paragraph.runs:
                            minimal_docx._apply_run(run, format_config["table"])

    @classmethod
    def _typical_segment_heading(cls, doc, text):
        """G3（用户第 3 轮）：③ 典型段标题 `1）…段` 按模板为**正文级**——
        Normal 段落、pPr 不含 outlineLvl（不进导航大纲），直接格式 楷体_GB2312 12 bold。
        不能用 Heading 5：模板该标题 `style=Normal` 且无 outlineLvl。字体仍由
        `_apply_typical_heading_font` 在最后一遍兜底（否则会被 `_format_all_run_fonts` 重置）。"""
        from docx.shared import Pt
        paragraph = doc.add_paragraph()
        run = paragraph.add_run(text)
        run.bold = True
        run.font.size = Pt(12)
        cls._set_run_fonts(run, "楷体_GB2312")
        paragraph.paragraph_format.keep_with_next = True  # 与 T2f P2-3 同类：标题不与正文被拆页
        return paragraph

    @classmethod
    def _apply_typical_heading_font(cls, doc):
        """D15/G3：① 小节下的典型段标题 `1）2）3）…` 按模板用 楷体_GB2312 12 bold（正文级）。

        G3 起 `1）` 段为 Normal（无 outlineLvl），故按文本而非大纲层级识别；
        只在最后一遍字体格式化（_format_all_run_fonts）之后覆盖字体，否则会被重置回仿宋。
        """
        import re
        from docx.shared import Pt
        for paragraph in doc.paragraphs:
            if not re.match(r"^\d+）", paragraph.text.strip()):
                continue
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(12)
                cls._set_run_fonts(run, "楷体_GB2312")

    @staticmethod
    def _alpha_label(index):
        value = int(index)
        result = ""
        while value:
            value, remainder = divmod(value - 1, 26)
            result = chr(ord("a") + remainder) + result
        return result + "."

    @staticmethod
    def _chapter_number(doc):
        import re
        chinese = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}
        for paragraph in doc.paragraphs:
            match = re.match(r"^([一二三四五六七八九十]+|\d+)、", paragraph.text.strip())
            if not match:
                continue
            token = match.group(1)
            if token.isdigit():
                return int(token)
            if token == "十":
                return 10
            if len(token) == 2 and token[0] == "十":
                return 10 + chinese[token[1]]
            if len(token) == 2 and token[1] == "十":
                return chinese[token[0]] * 10
            if len(token) == 3 and token[1] == "十":
                return chinese[token[0]] * 10 + chinese[token[2]]
            return chinese.get(token, 1)
        return 1

    @classmethod
    def _marking_segment_sentence(cls, row):
        names = row.get("_side_names") or {}
        ordered = list(names)
        if not ordered:
            ordered = ["标线2", "标线3"]; names = marking_side_names(ordered)
        parts = []
        for pos in ordered:
            name = names[pos]
            count, avg, rate = row.get(f"{pos}_valid_count", 0) or 0, cls._fmt(row.get(f"{pos}_average")), cls._pct(row.get(f"{pos}_qualified_rate"))
            parts.append(f"{name}共{count:,}个有效计算单元，平均逆反射亮度系数{avg}，合格率{rate}")
        return "该区段" + "；".join(parts) + "。"

    @staticmethod
    def _height_segment_sentence(rows):
        parts = []
        total = 0
        for row in rows:
            count = int(row.get("valid_count") or 0)
            total += count
            parts.append(f"{_guardrail_type(row.get('guardrail_type'))}护栏{count:,}个有效点，平均高度{GuangdongChapterWriter._fmt(row.get('average'))} mm，合格率{GuangdongChapterWriter._pct(row.get('qualified_rate'))}")
        return f"该区段护栏中心高度共检测{total:,}个有效点，其中" + "；".join(parts) + "。"

    @staticmethod
    def _bolt_segment_sentence(row):
        return (f"该区段识别现有拼接螺栓{int(row.get('splice', 0) or 0):,}颗、连接螺栓{int(row.get('connection', 0) or 0):,}颗，检出缺失螺栓{int(row.get('missing_total', 0) or 0):,}颗，螺栓缺失率为{GuangdongChapterWriter._pct(row.get('missing_rate'))}。")

    @staticmethod
    def _guardrail_note_counts(notes):
        """按区段聚合桥梁/隧道标记：{(路线,方向,区段): {"桥梁":n,"隧道":n}}。"""
        result = {}
        for row in notes or []:
            remark = str(row.get("guardrail_note") or "")
            key = (row.get("route"), row.get("direction"), row.get("segment"))
            counts = result.setdefault(key, {"桥梁": 0, "隧道": 0})
            if "桥梁" in remark: counts["桥梁"] += 1
            if "隧道" in remark: counts["隧道"] += 1
        return result

    @staticmethod
    def _guardrail_note_sentence(counts):
        if not counts or not (counts["桥梁"] or counts["隧道"]): return None
        if counts["桥梁"] and counts["隧道"]: return "该区段为桥梁和隧道路段，区段内无护栏"
        if counts["桥梁"]: return "该区段为桥梁路段，区段内无护栏"
        return "该区段为隧道路段，区段内无护栏"

    @staticmethod
    def _guardrail_no_valid_point_sentence(counts):
        if not counts or not (counts["桥梁"] or counts["隧道"]): return None
        if counts["桥梁"] and counts["隧道"]: return "当前区段为桥梁和隧道路段，无有效检测点位。"
        if counts["桥梁"]: return "当前区段为桥梁路段，无有效检测点位。"
        return "当前区段为隧道路段，无有效检测点位。"

    @classmethod
    def write(cls,city,bundle,output_dir,template,thresholds):
        from backend import markdown_skeleton, minimal_docx
        city=_safe_city_component(city)
        template=Path(template)
        if not template.is_file(): raise FileNotFoundError(f"Word模板不存在：{template}")
        if template.suffix.lower() != ".md":
            raise ValueError(f"仅支持 Markdown 模板：{template}，请使用 .md 模板。")
        template_data = markdown_skeleton.read_template(template)
        with minimal_docx.format_context(template_data.config):
            return cls._write_document(city,bundle,output_dir,template,thresholds,template_data)

    @classmethod
    def _write_document(cls,city,bundle,output_dir,template,thresholds,template_data):
        from docx import Document
        from docx.shared import Pt
        from backend import minimal_docx
        doc = Document()
        cls._configure_heading_styles(doc)
        blocks = template_data.blocks
        category_titles = (("高速公路", "（一）高速公路交安设施技术状况"), ("普通国省道", "（二）普通国省道交安设施技术状况"))
        # 道路类别骨架随该类别动态检测内容一起渲染，避免先输出两套同名章节。
        category_blocks = {title: [] for _, title in category_titles}
        intro_blocks = []
        target = intro_blocks
        for block in blocks:
            if block.kind == "heading" and block.level <= 2:
                target = category_blocks.get(block.text, intro_blocks)
            target.append(block)
        def _skeleton_picture(caption: str):
            try:
                media_path = (template.parent / caption).resolve()
                if media_path.is_file():
                    try:
                        cls._picture(doc, media_path, Pt(380))
                    except Exception:
                        pass
            except Exception:
                pass

        def _render(blocks):
            markdown_skeleton.render_skeleton(
                doc, blocks,
                heading=lambda text, level: cls._heading(doc, text, level),
                body=lambda text: cls._body(doc, text),
                table=lambda headers, rows, shading: cls._add_table(doc, headers, rows),
                picture=_skeleton_picture,
                toc=lambda: minimal_docx.add_toc(doc, template_data.config["toc"]),
                replace=intro_replace,
            )
        GuangdongStatistics.fill_managers(bundle.get("marking") or [], bundle.get("height") or [], bundle.get("bolt") or [])

        # 清单表 `类型` 列按高速块/普通国省道块纵向合并（D4）：先按分支稳定排序，保证同值连续
        inspection = sorted(gd_inspection_segments(bundle),
                            key=lambda row: 0 if row.get("category") == "高速公路" else 1)

        km_totals = {}

        for row in inspection:

            km_totals[row["category"]] = km_totals.get(row["category"], 0.0) + row["length_km"]

        # P1b：经营主体按附件原值定（A1 §6.1 P0→P4，用本行起止桩号消歧），判不出写 None 走「—」，不猜。
        route_index = bundle.get("route_index")
        owner_for = getattr(route_index, "owner_for", None)
        for row in inspection:
            row["owner"] = (owner_for(city, row.get("route"), row.get("start"), row.get("end"),
                                      row.get("manager")) if owner_for else None) or None
        # P1h：分类表标「跨主体」的行，一行抽检里程横跨两个主体，单值装不下 → ①表两侧都漏、
        # 出现「非集团+集团 < 全市」的差额。按 A1 §6.1 的 P2 桩号区间把这行拆到两侧
        # （附件覆盖不到的尾段用 P3 管养单位），拆不干净就不拆，仍如实报差额。
        # 只对 CROSS_OWNER 行动手：其余行走 P1b 的单值路径，逐值不变。
        owner_spans_of = getattr(route_index, "owner_spans", None)
        if owner_spans_of:
            _keepers = defaultdict(list)
            for _kind in ("marking", "height", "bolt"):
                for _r in bundle.get(_kind) or ():
                    if _r.get("station_m") is None or not _r.get("manager"):
                        continue
                    _km = float(_r["station_m"]) / 1000
                    _keepers[(str(_r.get("route") or ""), str(_r.get("direction") or ""))].append(
                        (_km, _km, str(_r["manager"])))
            _split, _unresolved = 0, []
            for row in inspection:
                # 已有单一具体主体的行走 P1b 单值路径，逐值不变；只处理判不出/跨主体的行。
                if row.get("owner") not in (None, CROSS_OWNER):
                    continue
                _sp = gd_cross_owner_spans(
                    owner_spans_of(city, row.get("route"), row.get("start"), row.get("end"),
                                   _keepers.get((str(row.get("route") or ""),
                                                 str(row.get("direction") or "")), ())),
                    row.get("start"), row.get("end"))
                if _sp:
                    row["owner_spans"] = _sp
                    _split += 1
                else:
                    _unresolved.append(str(row.get("route")))
            if _split or _unresolved:
                print(f"[经营主体] 跨主体行按桩号拆分 {_split} 行"
                      + (f"；未能拆分（差额如实保留）：{'、'.join(_unresolved)}" if _unresolved else ""))
        if owner_for:
            _miss = [str(row.get("route")) for row in inspection if not row.get("owner")]
            print(f"[经营主体] 清单表 {len(inspection) - len(_miss)}/{len(inspection)} 行可定经营主体"
                  + (f"；未命中：{'、'.join(_miss)}" if _miss else ""))

        def _intro_stats(category):
            """（一）/（二）小节引言的里程/路线/管养单位/省交通集团里程（D6，全部来自本市抽检路段行）。"""
            subset = [row for row in inspection if row.get("category") == category]
            routes, managers = [], []
            for row in subset:
                if row.get("route") and row["route"] not in routes:
                    routes.append(row["route"])
                name = str(row.get("manager") or "").strip()
                if name and name not in managers:
                    managers.append(name)
            km_total = sum(float(row.get("length_km") or 0.0) for row in subset)
            # P1b：集团/非集团里程改用与 ① 表同源的经营主体分组（附件原值），
            # 弃用 manager_is_provincial 关键字启发式（A1 §6.1 P6 明令禁止）。
            provincial = gd_mileage(inspection, category=category, owner=GD_OWNER_PROVINCIAL) or 0.0
            other = gd_mileage(inspection, category=category, owner=GD_OWNER_NON_PROVINCIAL) or 0.0
            # 跨主体/未命中的行两边都不落，差额如实记录，不硬摊到某一侧（A1 §6.2/§6.3）。
            unassigned = round(km_total - provincial - other, 3)
            # 只对高速分支报不闭合：普通国省道的 ① 表按道路类别分组，不用经营主体（A1 §7：附件无国省道行）。
            if category == "高速公路" and abs(unassigned) > 0.5:
                cross = sorted({f"{row.get('route')}={row.get('owner') or '缺失'}"
                                for row in subset if row.get("owner") not in (GD_OWNER_PROVINCIAL, GD_OWNER_NON_PROVINCIAL)})
                print(f"[经营主体] {city}{category}：集团 {provincial:.3f} + 非集团 {other:.3f} "
                      f"< 全市 {km_total:.3f}，差额 {unassigned:.3f} km（{len(cross)} 行：{'、'.join(cross)}）")
            return {"routes": routes, "managers": managers, "km": km_total,
                    "provincial": provincial, "other": other,
                    # P1k：引言里程取整必须与 ① 表同源同法（最大余数法），否则甲方把引言两个
                    # 分项加起来会与 ① 表对不上（深圳曾 75+163=238 ≠ ① 表 237=163+74）。
                    # 高速分支走 gd_km_split_text（① 表 gd_km_row_texts 调用的同一函数）；
                    # 普通分支没有「全市=分项」三行关系，① 表走 gd_km_text，引言同法。
                    "km_texts": (gd_km_split_text(km_total, other, provincial)
                                 if category == "高速公路"
                                 else (gd_km_text(km_total),))}

        def _route_label(route, names):
            """P1c：引言路线列表「路线号 + 路线名称」（brief-v3 需求 #3）。
            名称取路线分类表权威值；查不到就只写路线号（不猜、不用别处措辞顶替），
            由叙述名称扫描如实报为「未命中」。"""
            return f"{route}{names.get(route) or ''}"

        _names = bundle.get("route_names") or {}   # 清单表/①②③表与引言路线列表共用同一份权威名称
        _high = _intro_stats("高速公路")
        _ord = _intro_stats("普通国省道")
        _ord_g = sum(float(row.get("length_km") or 0.0) for row in inspection
                     if row.get("category") == "普通国省道" and str(row.get("route") or "").upper().startswith("G"))
        _ord_s = _ord["km"] - _ord_g
        intro_replace = {"{{地市}}": city,
                         "{{高速抽检里程}}": _high["km_texts"][0],
                         "{{普通国省道抽检里程}}": _ord["km_texts"][0],
                         "{{高速路线列表}}": "、".join(_route_label(route, _names) for route in _high["routes"]) or "—",
                         "{{高速路线数}}": f"{len(_high['routes'])}",
                         "{{高速管养单位数}}": f"{len(_high['managers'])}",
                         "{{高速省集团里程}}": _high["km_texts"][2],
                         "{{高速非省集团里程}}": _high["km_texts"][1],
                         "{{普通路线列表}}": "、".join(_route_label(route, _names) for route in _ord["routes"]) or "—",
                         "{{普通路线数}}": f"{len(_ord['routes'])}",
                         "{{普通管养单位数}}": f"{len(_ord['managers'])}",
                         "{{普通国道里程}}": f"{_ord_g:.0f}",
                         "{{普通省道里程}}": f"{_ord_s:.0f}"}
        # D6 口径：里程/路线数/管养单位数/省交通集团里程均由本市抽检路段（路线表）行汇总，不编造
        _render(intro_blocks)
        # P1-2：引言段 mcd·m-²·lx-¹ 的 -²/-¹ 按模板设为上标
        for _paragraph in doc.paragraphs:
            apply_marking_superscripts(_paragraph)
        chapter_no = 5
        caption_counts={"figure":0,"table":0}

        def _caption(kind,title):
            caption_counts[kind]+=1
            prefix="图" if kind=="figure" else "表"
            cap=doc.add_paragraph(f"{prefix}{chapter_no}-{caption_counts[kind]} {title}")
            cap.style="Caption"
            cls._format_chart_caption(cap)

        def _table(headers,rows,title,merge=None,vmerge=None,tail_merge=None):
            _caption("table",title)
            return cls._add_table(doc,headers,rows,merge=merge,vmerge=vmerge,tail_merge=tail_merge)

        def _figure(path,title):
            cls._picture(doc, path, Pt(440))
            _caption("figure",title)

        route_names=_names

        if inspection:
            def _gdk(value):
                """D19：清单表 起点/终点桩号、段长 一律输出整数公里（四舍五入）。"""
                return "—" if value is None else f"{int(float(value) + 0.5)}"

            _table(["类型", "路线编号", "路线名称", "检测方向", "起点桩号", "终点桩号", "段长（Km)", "管养单位", "备注"],
                   [[gd_stacked_category(row["category"]), row["route"], route_names.get(row["route"], "—"),
                     row["direction"] or "—", _gdk(row.get("start")), _gdk(row.get("end")), _gdk(row.get("length")),
                     row.get("manager") or "—", row.get("remark") or "—"]
                    for row in inspection],
                   f"{city}交通安全设施抽检路段清单",
                   vmerge={0: _contiguous_runs([row["category"] for row in inspection])})
        for category,chapter_title in category_titles:
            all_mark=[r for r in bundle.get("marking",[]) if r.get("category")==category]
            all_height=[r for r in bundle.get("height",[]) if r.get("category")==category]
            all_bolt=[r for r in bundle.get("bolt",[]) if r.get("category")==category]
            all_detail=[r for r in bundle.get("comparison_detail",[]) if r.get("category")==category]
            chart_prefix=f"{city}_{category}_"
            base=Path(output_dir)/city/"charts"
            def _images(suffix):
                return sorted(str(p) for p in base.glob(f"{chart_prefix}*{suffix}*.png")) if base.exists() else []

            if category_blocks[chapter_title]:
                _render(category_blocks[chapter_title])
            else:
                cls._heading(doc,chapter_title,2)
                cls._heading(doc,"1.沿线设施技术状况TCI",3)

            cls._heading(doc,"2.标线、护栏自动化检测",3)
            cls._heading(doc,"（1）标线逆反射亮度系数情况",4)
            segment_sort_key=lambda row: tuple(str(row.get(field) or "") for field in ("route","direction","segment"))
            GuangdongStatistics.fill_managers(all_mark,all_height,all_bolt)
            # C4：①/②/③ 表管养单位统一到路线分类表口径（与清单表、引言句同源；缺则「—」）
            manager_lookup=gd_manager_lookup(inspection,city)
            all_mark=gd_route_managed_rows(all_mark,manager_lookup)
            all_height=gd_route_managed_rows(all_height,manager_lookup)
            all_bolt=gd_route_managed_rows(all_bolt,manager_lookup)
            cls._marking_gd03_section(doc,category,all_mark,_table,_figure,base,chart_prefix,city,route_names,
                                      inspection=inspection)

            cls._heading(doc,"（2）波形梁护栏中心高度情况",4)
            cls._height_gd03_section(doc,category,all_height,_table,_figure,base,chart_prefix,city,route_names,
                                     inspection=inspection)

            cls._heading(doc,"（3）波形梁护栏螺栓缺失情况",4)
            cls._bolt_gd03_section(doc,category,all_bolt,_table,_figure,base,chart_prefix,city,route_names,
                                   inspection=inspection)
            cls._heading(doc,"（4）人工复核对比情况",4)
            cls._comparison_gd03_section(doc,category,all_detail,thresholds,_table,all_mark)

        cls._heading(doc,"（三）工作建议",2)
        cls._advice_gd03_section(doc,bundle,_table)
        cls._indent_existing_body(doc)
        apply_gd_table_widths(doc)
        apply_gd_table_heights(doc)
        normalize_station_separators(doc)
        cls._strip_template_notes(doc)
        cls._format_all_run_fonts(doc)
        cls._apply_typical_heading_font(doc)
        folder=Path(output_dir)/city; folder.mkdir(parents=True,exist_ok=True); path=folder/f"{city}在役公路技术状况检测评价报告第五部分.docx"
        try: doc.save(path)
        except PermissionError as exc: raise PermissionError(f"Word文件被占用：{path}") from exc
        return path



# ==================== 广东项目图表生成 ====================

def _gd_height_charts(rows, route, direction, chart_dir, prefix=""):
    """护栏中心高度：每个区段、每种护栏类型分别绘制折线图与结果分布饼图。"""
    images = {}
    safe_route = route.replace("/", "_").replace("\\", "_")
    safe_dir = direction.replace("/", "_")
    for segment in sorted({str(r.get("segment", "")) for r in rows}):
        seg_rows = [r for r in rows if str(r.get("segment", "")) == segment]
        safe_seg = segment.replace("/", "_").replace("\\", "_").replace(" ", "_") or "segment"
        for kind in ("二波", "三波"):
            kind_rows = sorted(
                [r for r in seg_rows if _guardrail_type(r.get("guardrail_type")) == kind
                 and _float(r.get("height")) is not None and r.get("station_m") is not None],
                key=lambda r: r["station_m"],
            )
            if not kind_rows:
                continue
            heights = [float(r["height"]) for r in kind_rows]
            x = list(range(len(kind_rows)))
            seg_text = str(segment or "") + ("段" if segment and not str(segment).endswith("段") else "")

            line_path = chart_dir / f"{prefix}height_{safe_route}_{safe_dir}_{safe_seg}_{kind}_line.png"
            fig, ax = plt.subplots(figsize=(13 / 2.54, 8 / 2.54), dpi=180)
            ax.plot(x, heights, color=GD03_COLORS[0], linewidth=1, label="梁板中心高度(mm)")
            standards = (580, 620) if kind == "二波" else (677, 717)
            for std, color in zip(standards, GD03_COLORS[1:3]):
                ax.plot(x, [std] * len(x), color=color, linewidth=2, label=f"标准值（{std}mm）")
            ax.set_ylim(300, 850)
            ax.grid(True, alpha=0.25)
            tick_count = min(10, len(x))
            ticks = sorted(set(int(i * (len(x) - 1) / (tick_count - 1)) for i in range(tick_count))) if tick_count > 1 else [0]
            ax.set_xticks(ticks)
            ax.set_xticklabels([format_station(kind_rows[i]["station_m"]) for i in ticks], rotation=30, ha="right", fontsize=7)
            ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.22), ncol=3, frameon=False)
            fig.subplots_adjust(left=0.10, right=0.98, top=0.92, bottom=0.30)
            fig.savefig(line_path, transparent=False, bbox_inches="tight", pad_inches=0.15)
            plt.close(fig)

            pie_path = chart_dir / f"{prefix}height_{safe_route}_{safe_dir}_{safe_seg}_{kind}_pie.png"
            limits = (550, 580, 620, 650) if kind == "二波" else (647, 677, 717, 747)
            labels = [
                f"h＜{limits[0]}", f"{limits[0]}≤h＜{limits[1]}",
                f"{limits[1]}≤h≤{limits[2]}", f"{limits[2]}＜h≤{limits[3]}", f"h＞{limits[3]}",
            ]
            bins = [0] * 5
            for h in heights:
                if h < limits[0]: bins[0] += 1
                elif h < limits[1]: bins[1] += 1
                elif h <= limits[2]: bins[2] += 1
                elif h <= limits[3]: bins[3] += 1
                else: bins[4] += 1
            total = len(heights)
            pcts = [b * 100 / total if total else 0 for b in bins]
            nonzero = [(lbl, v, f"#{PIE_COLORS[i]}") for i, (lbl, v) in enumerate(zip(labels, pcts)) if v > 0]
            if nonzero:
                fig, ax = plt.subplots(figsize=(14 / 2.54, 8.5 / 2.54), dpi=180)
                wedges, _, _ = ax.pie(
                    [v for _, v, _ in nonzero], colors=[c for _, _, c in nonzero],
                    autopct=lambda pct: f"{pct:.2f}%" if pct > 0 else "", pctdistance=1.15,
                    textprops={"fontsize": 8},
                )
                ax.legend(wedges, [lbl for lbl, _, _ in nonzero], loc="center left", bbox_to_anchor=(1.0, 0.5), frameon=False)
                fig.subplots_adjust(left=0.02, right=0.76, top=0.88, bottom=0.05)
                fig.savefig(pie_path, transparent=False, bbox_inches="tight", pad_inches=0.15)
                plt.close(fig)
                images[f"height_{safe_route}_{safe_dir}_{safe_seg}_{kind}"] = {"line": line_path, "pie": pie_path}
            else:
                images[f"height_{safe_route}_{safe_dir}_{safe_seg}_{kind}"] = {"line": line_path}
    return images

def _gd_bolt_charts(rows, route, direction, chart_dir, prefix=""):
    """护栏螺栓缺失：每个区段绘制一组柱状图（拼接/连接/缺失数量）。"""
    images = {}
    safe_route = route.replace("/", "_").replace("\\", "_")
    safe_dir = direction.replace("/", "_")
    segments = {}
    for r in rows:
        segments.setdefault(str(r.get("segment", "")), []).append(r)
    for seg in sorted(segments.keys()):
        seg_rows = segments[seg]
        s = sum(float(r.get("splice", 0) or 0) for r in seg_rows)
        c = sum(float(r.get("connection", 0) or 0) for r in seg_rows)
        m = sum(float(r.get("splice_missing", 0) or 0) + float(r.get("connection_missing", 0) or 0) for r in seg_rows)
        safe_seg = seg.replace("/", "_").replace("\\", "_").replace(" ", "_") or "segment"
        seg_text = str(seg or "") + ("段" if seg and not str(seg).endswith("段") else "")
        bar_path = chart_dir / f"{prefix}bolt_{safe_route}_{safe_dir}_{safe_seg}_bar.png"
        fig, ax = plt.subplots(figsize=(13 / 2.54, 8 / 2.54), dpi=180)
        category_vals = [int(round(s)), int(round(c)), int(round(m))]
        bar_labels = ["拼接螺栓", "连接螺栓", "缺失数量"]
        bar_colors = [GD03_COLORS[0], GD03_COLORS[1], "#FF0000"]
        ax.bar(range(3), category_vals, 0.5, color=bar_colors)
        ax.set_xticks(range(3))
        ax.set_xticklabels(bar_labels, fontsize=9)
        ax.set_ylabel("螺栓数量（颗）")
        for i, v in enumerate(category_vals):
            ax.text(i, v, str(v), ha="center", va="bottom", fontsize=9)
        ax.grid(True, axis="y", alpha=0.25)
        fig.subplots_adjust(left=0.10, right=0.98, top=0.92, bottom=0.12)
        fig.savefig(bar_path, transparent=False, bbox_inches="tight", pad_inches=0.15)
        plt.close(fig)
        images[f"bolt_{safe_route}_{safe_dir}_{safe_seg}"] = {"bar": bar_path}
    return images

def _gd_marking_charts(rows, route, direction, chart_dir, prefix=""):
    """逐报告区段、逐侧绘制实测值与各行实际目标值，不跨身份混图。"""
    images = {}
    groups = GuangdongStatistics._group(rows, (*MARKING_SEGMENT_FIELDS, "marking_position"))
    for pos_rows in groups.values():
        pos_rows = sorted(
            [r for r in pos_rows if _float(r.get("value")) is not None and r.get("station_m") is not None],
            key=lambda r: r["station_m"],
        )
        if not pos_rows:
            continue
        values = [float(r["value"]) for r in pos_rows]
        x = list(range(len(pos_rows)))

        line_path = _marking_chart_path(chart_dir, pos_rows[0])
        fig, ax = plt.subplots(figsize=(13 / 2.54, 8 / 2.54), dpi=180)
        ax.plot(x, values, color=GD03_COLORS[0], linewidth=1, marker="." if len(x) == 1 else None, label="逆反射亮度系数")
        ax.plot(x, [r.get("target", 80) for r in pos_rows], color=GD03_COLORS[1], linewidth=2, linestyle="--", label="目标值")
        ax.grid(True, alpha=0.25)
        tick_count = min(10, len(x))
        ticks = sorted(set(int(i * (len(x) - 1) / (tick_count - 1)) for i in range(tick_count))) if tick_count > 1 else [0]
        ax.set_xticks(ticks)
        ax.set_xticklabels([format_station(pos_rows[i]["station_m"]) for i in ticks], rotation=30, ha="right", fontsize=7)
        ax.set_ylabel("逆反射亮度系数")
        ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.22), ncol=3, frameon=False)
        fig.subplots_adjust(left=0.10, right=0.98, top=0.92, bottom=0.30)
        fig.savefig(line_path, transparent=False, bbox_inches="tight", pad_inches=0.15)
        plt.close(fig)

        images[line_path.stem] = {"line": line_path}
    return images

def guangdong_report_images(bundle, output_dir, log=lambda _: None):
    """为广东项目城市 bundle 生成护栏高度、螺栓缺失、标线逆反射图表。"""
    city = _safe_city_component(bundle["city"])
    chart_dir = Path(output_dir) / city / "charts"
    chart_dir.mkdir(parents=True, exist_ok=True)

    plt.rcParams["font.sans-serif"] = ["SimSun", "Microsoft YaHei", "Arial Unicode MS"]
    plt.rcParams["axes.unicode_minus"] = False

    all_images = {}
    for kind in ("height", "bolt", "marking"):
        rows = bundle.get(kind, [])
        if not rows:
            continue
        groups = {}
        for row in rows:
            key = (str(row.get("route", "")), str(row.get("direction", "")))
            groups.setdefault(key, []).append(row)
        for (route, direction), group_rows in groups.items():
            category = group_rows[0].get("category") or ""
            prefix = f"{city}_{category}_" if category else ""
            try:
                if kind == "height":
                    all_images.update(_gd_height_charts(group_rows, route, direction, chart_dir, prefix))
                elif kind == "bolt":
                    all_images.update(_gd_bolt_charts(group_rows, route, direction, chart_dir, prefix))
                elif kind == "marking":
                    all_images.update(_gd_marking_charts(group_rows, route, direction, chart_dir, prefix))
            except Exception as exc:
                log(f"图表生成警告（{kind}/{route}/{direction}）：{exc}")

    if all_images:
        log(f"{city}图表已生成：{len(all_images)}组，保存至{chart_dir}")
    return all_images


def marking_side_names(positions):
    """显式左右不重命名；旧编号对沿用小号左、大号右，其他名称原样保留。"""
    positions = {str(p) for p in positions if p not in (None, "")}
    numbered = sorted((p for p in positions if re.fullmatch(r"标线\d+", p)), key=lambda p: int(p[2:]))
    names = {p: p for p in positions}
    if len(numbered) == 1:
        names[numbered[0]] = "右侧标线" if numbered[0] == "标线3" else "左侧标线"
    elif len(numbered) == 2:
        names.update(zip(numbered, ("左侧标线", "右侧标线")))
    return dict(sorted(names.items(), key=lambda item: ({"左侧标线": 0, "右侧标线": 1}.get(item[1], 2), item[0])))


GD03_COLORS = GD_SERIES_HEX

COMPANY_PARENTS = (
    "广东省公路建设有限公司",
    "广东省路桥建设发展有限公司",
    "广东省高速公路有限公司",
    "广东交通实业投资有限公司",
    "广东省南粤交通投资建设有限公司",
    "广东省交通集团有限公司",
)


def gd_marking_color(row, marking_rows):
    """P1-1 标线颜色：由该行匹配到的自动化单元目标值推导（80→白色、50→黄色），判定不了写 —。
    源数据（人工对比表）无颜色字段；先按路线+方向+自动化测值定位自动化单元，
    命不中则退化为该路线方向全部单元；目标值不唯一（无法判定）时写 —，不回填道路等级。"""
    automatic = row.get("automatic")
    same_route = [r for r in marking_rows or []
                  if str(r.get("route") or "").strip().upper() == str(row.get("route") or "").strip().upper()
                  and str(r.get("direction") or "").strip() == str(row.get("direction") or "").strip()]
    exact = [r for r in same_route if automatic is not None and r.get("value") is not None
             and abs(float(r["value"]) - float(automatic)) <= 0.05]
    targets = {round(float(r["target"])) for r in (exact or same_route) if r.get("target") is not None}
    if len(targets) != 1:
        return "—"
    return {80: "白色", 50: "黄色"}.get(targets.pop(), "—")


_MARKING_SUPERSCRIPT_RE = re.compile(r"(?<=mcd·m)(-²)|(?<=·lx)(-¹)")


def apply_marking_superscripts(paragraph):
    """把引言中「mcd·m-²·lx-¹」的 -²/-¹ 设为上标（模板-1 run 级上标，共 4 处）；返回上标 run 数。"""
    if len(paragraph.runs) != 1 or "mcd·m" not in paragraph.text:
        return 0
    text = paragraph.text
    parts, last = [], 0
    for match in _MARKING_SUPERSCRIPT_RE.finditer(text):
        parts.append((text[last:match.start()], False))
        parts.append((match.group(), True))
        last = match.end()
    parts.append((text[last:], False))
    if not any(superscript for _, superscript in parts):
        return 0
    run = paragraph.runs[0]
    rpr = run._element.get_or_add_rPr()
    run.text = parts[0][0]
    count = 0
    for piece, superscript in parts[1:]:
        if not piece:
            continue
        new_run = paragraph.add_run(piece)
        new_run._element.insert(0, deepcopy(rpr))
        if superscript:
            new_run.font.superscript = True
            count += 1
    return count


def split_manager_company(manager):
    """管养单位 →（母公司，二级/路段公司）；非公司制单位返回（None，原值）。"""
    text = str(manager or "").strip()
    for parent in COMPANY_PARENTS:
        if text.startswith(parent):
            return parent, text[len(parent):].strip() or text
    return None, text


def manager_display(manager, city=""):
    """表格/正文显示名：统一到路线分类表口径（与清单表一致），不做二次简写；缺值写 —。
    P2-5：T2c 已把①/②/③表分组键统一到路线表口径，T2f 把显示名也统一
    （此前取二级公司/去地市前缀，导致同一实体在不同表出现两种写法）。"""
    return str(manager or "").strip() or "—"


def _gd03_overall(left, right):
    """总体合格率 =（左侧合格率 + 右侧合格率）/2，两侧均先按 1 位小数取整（oracle 口径）。"""
    if left is not None and right is not None:
        return (round(left * 100, 1) + round(right * 100, 1)) / 200.0
    return left if left is not None else right


def _gd03_axes():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    matplotlib.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "DejaVu Sans"]
    matplotlib.rcParams["axes.unicode_minus"] = False
    return plt


def gd03_bar_chart(path, categories, series, ylabel="合格率（%）", ylim=(0, 100), rotation=0, value_fmt="%.1f", horizontal=False):
    """GD03 分组柱状图（附件样式）：柱顶数据标签、水平网格线、图例右上带框、纵轴 0–100。"""
    plt = _gd03_axes()
    if horizontal:
        figure, axes = plt.subplots(figsize=(max(6.4, 3.2), max(3.0, 0.42 * len(categories) + 1.4)), dpi=160)
        height = 0.8 / max(len(series), 1)
        for index, (name, values, color) in enumerate(series):
            offsets = [position + (index - (len(series) - 1) / 2) * height for position in range(len(categories))]
            bars = axes.barh(offsets, [float("nan") if value is None else float(value) for value in values],
                             height=height * 0.9, label=name, color=color)
            axes.bar_label(bars, fmt=value_fmt, fontsize=7, padding=1)
        axes.set_yticks(range(len(categories)))
        axes.set_yticklabels(categories, fontsize=8)
        axes.invert_yaxis()
        axes.set_xlim(*ylim)
        axes.set_xlabel(ylabel, fontsize=9)
        axes.grid(axis="x", color="#D9D9D9", linewidth=0.8)
        axes.set_axisbelow(True)
        axes.legend(loc="lower right", frameon=True, fontsize=8)
        for side in ("top", "right"):
            axes.spines[side].set_visible(False)
        figure.tight_layout()
        figure.savefig(path)
        plt.close(figure)
        return str(path)
    figure, axes = plt.subplots(figsize=(max(6.4, 1.15 * len(categories) + 2.2), 3.8), dpi=160)
    width = 0.8 / max(len(series), 1)
    for index, (name, values, color) in enumerate(series):
        offsets = [position + (index - (len(series) - 1) / 2) * width for position in range(len(categories))]
        bars = axes.bar(offsets, [float("nan") if value is None else float(value) for value in values],
                        width=width * 0.9, label=name, color=color)
        axes.bar_label(bars, fmt=value_fmt, fontsize=7, padding=1)
    axes.set_xticks(range(len(categories)))
    # P1c：② 图类别标签是「路线号+管养单位」（实测 8 类时长 20~27 字，rotation=0 时相邻标签
    # 压字到读不出）。Excel 图表会自动旋转长标签，matplotlib 不会 —— 这里按标签长度补旋转，
    # 内容一字不改。阈值 12：① 图标签 6~7 字（实测压字 0，保持横排）；② 图 20~27 字（必转）。
    _need_rotate = rotation or max((len(str(c)) for c in categories), default=0) > 12
    axes.set_xticklabels(categories, rotation=45 if _need_rotate else 0, fontsize=8,
                         ha="right" if _need_rotate else "center")
    axes.set_ylim(*ylim)
    axes.set_ylabel(ylabel, fontsize=9)
    axes.grid(axis="y", color="#D9D9D9", linewidth=0.8)
    axes.set_axisbelow(True)
    axes.legend(loc="upper right", frameon=True, fontsize=8)
    for side in ("top", "right"):
        axes.spines[side].set_visible(False)
    figure.tight_layout()
    # P1c：标签转 45° 后首尾两条会伸出画布被切（人眼实测 image2 左/右端「司」字缺字）。
    # 与本文件其它图表函数一致用 bbox_inches="tight" 让 matplotlib 自算边界，图形内容不变。
    figure.savefig(path, bbox_inches="tight", pad_inches=0.1)
    plt.close(figure)
    return str(path)


def gd03_line_chart(path, categories, series, ylabel="合格率（%）", ylim=(0, 100), rotation=45, value_fmt="%.1f"):
    """GD03 折线图（附件样式）：圆点标记、水平网格线、图例右上带框、x 轴标签旋转 45°。"""
    plt = _gd03_axes()
    figure, axes = plt.subplots(figsize=(max(6.4, 0.6 * len(categories) + 2.2), 3.8), dpi=160)
    for name, values, color in series:
        axes.plot(range(len(categories)), [float("nan") if value is None else float(value) for value in values],
                  marker="o", markersize=4, linewidth=1.6, label=name, color=color)
    axes.set_xticks(range(len(categories)))
    axes.set_xticklabels(categories, rotation=rotation, fontsize=7, ha="right" if rotation else "center")
    axes.set_ylim(*ylim)
    axes.set_ylabel(ylabel, fontsize=9)
    axes.grid(axis="y", color="#D9D9D9", linewidth=0.8)
    axes.set_axisbelow(True)
    axes.legend(loc="upper right", frameon=True, fontsize=8)
    for side in ("top", "right"):
        axes.spines[side].set_visible(False)
    figure.tight_layout()
    figure.savefig(path)
    plt.close(figure)
    return str(path)


# 分类表缺失的抽检路段（用户 2026-09-28 裁决）：阳江的进场后人工复核对比表里出现 G324（螺栓缺失 12 行 +
# 人工自动化对比_统计结果 1 行），但分类表只有 云浮/惠州/肇庆 的 G324 行，阳江本市的没有。
# 只补这一行（普通国省道，经营主体留空），不因「导入回归 miss」增删其它行。
EXTRA_ROUTE_ROWS = (("阳江市", "G324", "福州－昆明", "普通国省道"),)


def write_guangdong_route_workbook_all(route_index, route_names, output_dir, log=lambda _: None,
                                       filename="路线分类表（含路线名称）.xlsx"):
    """导出一份覆盖全部地市、补充路线名称的路线分类表，可直接作为后续「路线分类表」导入使用。

    P1d：第 5 列为「经营主体」（表内列优先，其次附件匹配 + 实测段消歧）；取不到写空单元格，
    不写「—」「缺失」等占位，避免被当主体值参与分组。前 4 列顺序与语义保持不变。
    P1d-b：普通国省道行经营主体一律留空（附件无对应分段，无权威来源）；补齐 EXTRA_ROUTE_ROWS。
    """
    getter = getattr(route_index, "rows", None)
    rows = list(getter()) if callable(getter) else []
    present = {(RouteCategoryIndex._norm_city(row.get("city")), _route(row.get("route"))) for row in rows}
    # 只补「本市已有行但缺这一条」的行：非广东/单市表不凭空多行；本工具产物回灌时不重复补
    missing = [item for item in EXTRA_ROUTE_ROWS
               if (RouteCategoryIndex._norm_city(item[0]), _route(item[1])) not in present
               and any(RouteCategoryIndex._norm_city(row.get("city")) == RouteCategoryIndex._norm_city(item[0])
                       for row in rows)]
    if missing:
        rows.extend({"city": city, "route": route, "category": category, "name": name}
                    for city, route, name, category in missing)
    if not rows:
        return None
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill

    folder = Path(output_dir)
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / filename
    names = route_names or {}
    owner_for = getattr(route_index, "owner_for", None)
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "路线分类表"
    sheet.append(["地市", "路线编号", "路线名称", "道路类别", "经营主体"])
    for cell in sheet[1]:
        cell.font = Font(name="宋体", size=11, bold=True)
        cell.fill = PatternFill("solid", fgColor="D9D9D9")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    filled = misses = 0
    miss_list = []
    for row in sorted(rows, key=lambda item: (str(item.get("category") or ""), str(item.get("route") or ""))):
        name = names.get(row.get("route"), "") or row.get("name") or ""
        # 普通国省道不写经营主体（附件只有高速分段；裁决要求一律空单元格）
        owner = (owner_for(row.get("city"), row.get("route"))
                 if owner_for and row.get("category") == "高速公路" else None)
        if owner:
            filled += 1
        else:
            misses += 1
            miss_list.append(f"{row.get('city')}/{row.get('route')}")
        sheet.append([row.get("city"), row.get("route"), name, row.get("category"), owner or None])
    for column, width in zip("ABCDE", (14, 16, 34, 16, 14)):
        sheet.column_dimensions[column].width = width
    book.save(path)
    log(f"路线分类表（全量含路线名称+经营主体）已导出：{path}")
    log(f"经营主体覆盖 {filled}/{filled + misses} 行；未命中 {misses} 行：{'、'.join(miss_list) if miss_list else '无'}")
    return path


def write_guangdong_route_workbook(bundle, output_dir, log=lambda _: None):
    """导出本市的路线分类表（补充路线名称），文件可直接作为后续「路线分类表」导入使用。"""
    rows = list(bundle.get("route_rows") or [])
    if not rows:
        return None
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill

    city = _safe_city_component(bundle["city"])
    folder = Path(output_dir) / city
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / f"{city}路线分类表.xlsx"
    names = bundle.get("route_names") or {}
    # P1b：第 5 列「经营主体」（任务书第 3 项）。值取 A1 §6.1 的 P0→P4 匹配结果，匹配不到写「缺失」；
    # docx 内表格列不动（模板列固定，加列会破坏逐字对齐）。
    route_index = bundle.get("route_index")
    owner_for = getattr(route_index, "owner_for", None)
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "路线分类表"
    sheet.append(["地市", "路线编号", "路线名称", "道路类别", "经营主体"])
    for cell in sheet[1]:
        cell.font = Font(name="宋体", size=11, bold=True)
        cell.fill = PatternFill("solid", fgColor="D9D9D9")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    owner_missing = []
    for row in sorted(rows, key=lambda item: (str(item.get("category") or ""), str(item.get("route") or ""))):
        # P1d-b 裁决：普通国省道在附件里没有分段，无权威来源，一律「缺失」，不写「非集团」。
        owner = ((owner_for(row.get("city") or bundle.get("city"), row.get("route")) if owner_for else None)
                 or "缺失")
        if owner == "缺失":
            owner_missing.append(f"{row.get('city') or bundle.get('city')}/{row.get('route')}")
        sheet.append([row.get("city") or bundle.get("city"), row.get("route"),
                      names.get(row.get("route"), ""), row.get("category"), owner])
    for column, width in zip("ABCDE", (14, 16, 34, 16, 12)):
        sheet.column_dimensions[column].width = width
    book.save(path)
    filled = len(rows) - len(owner_missing)
    log(f"路线分类表已导出：{path}")
    log(f"经营主体有值 {filled}/{len(rows)} 行；「缺失」{len(owner_missing)} 行"
        + (f"：{'、'.join(owner_missing)}" if owner_missing else ""))
    return path


def _marking_chart_path(chart_dir, row):
    # 身份包含路线/方向/管养/区段/侧；摘要防止路径字符清洗碰撞，Word按同一键精确取图。
    identity = tuple(str(row.get(field, "")) for field in (*MARKING_SEGMENT_FIELDS, "marking_position"))
    key = hashlib.sha256(repr(identity).encode("utf-8")).hexdigest()
    return Path(chart_dir) / f"marking_{key}_line.png"


def _side_display(pos):
    """单个标线位置编号的显示名（图表标题用）。"""
    return marking_side_names([pos]).get(str(pos), str(pos))


def write_gd_removed_data_sheets(bundle, workbook):
    """D16：正文已删除的表/图数据落到统计工作簿，信息不丢（用户裁决 Q1）。

    对应关系：`标线长连续不合格`（原表5-5/5-20）、`高度超10cm长连续`（原表5-9/5-24）、
    `优先处治路段`（原表5-31/5-32）、`标线典型路段逐公里`/`高度最低公里段`/`螺栓严重缺失公里段`（原 4 张分布图）。
    """
    from openpyxl.styles import Alignment, Border, Font, Side
    thin = Side(style="thin", color="B7B7B7")
    city = str(bundle.get("city") or "")

    def new_sheet(name, headers):
        sheet = workbook.create_sheet(name)
        for offset, text in enumerate(headers, 1):
            cell = sheet.cell(1, offset, text)
            cell.font = Font(name="宋体", size=10, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        return sheet

    def fill(sheet, rows):
        for index, row in enumerate(rows, 2):
            for offset, value in enumerate(row, 1):
                sheet.cell(index, offset, _safe_excel_value(value))
        return sheet.title, len(rows)

    def span(start_m, end_m):
        if start_m is None or end_m is None:
            return "—"
        return f"{format_station(start_m)}-{format_station(end_m)}"

    groups = []
    for category in ("高速公路", "普通国省道"):
        groups.append((category, {
            "marking": [r for r in bundle.get("marking") or [] if r.get("category") == category],
            "height": [r for r in bundle.get("height") or [] if r.get("category") == category],
            "bolt": [r for r in bundle.get("bolt") or [] if r.get("category") == category],
        }))

    added = []
    added.append(fill(new_sheet("标线长连续不合格", ["道路类别", "路线", "管养单位", "标线侧", "起止桩号", "长度（km）", "3 km窗口不合格率（%）"]),
                      [[category, r["route"], manager_display(r.get("manager"), city), r["position_name"], span(r["start_m"], r["end_m"]),
                        round(r["length_km"], 2), (r["fail_rate"] or 0) * 100]
                       for category, group in groups for r in GuangdongStatistics.marking_long_runs(group["marking"])]))

    added.append(fill(new_sheet("高度超10cm长连续", ["道路类别", "路线", "方向", "管养单位", "公里区间", "长度（km）", "波形梁类型", "检测点数（个）", "超10cm占比（%）"]),
                      [[category, r["route"], r["direction"], GuangdongChapterWriter._manager_for_km(group["height"], r, city),
                        span(r["start_m"], r["end_m"]), r["length_km"], r["kind"], r["count"], (r["over_ratio"] or 0) * 100]
                       for category, group in groups for r in GuangdongStatistics.height_over10_runs(group["height"])]))

    added.append(fill(new_sheet("优先处治路段", ["道路类别", "路线", "方向", "区段", "问题类型", "起止桩号", "管养单位", "原因"]),
                      [[r.get("category") or "—", r.get("route"), r.get("direction"), r.get("segment"),
                        r.get("type"), span(r.get("start_m"), r.get("end_m")), r.get("manager"), r.get("reason")]
                       for r in bundle.get("weak_segments") or []]))

    added.append(fill(new_sheet("标线典型路段逐公里", ["道路类别", "路线", "方向", "公里", "总体合格率(%)", "左侧合格率(%)", "右侧合格率(%)", "有效计算单元数"]),
                      [[category, item["route"], item["direction"], r["km"], (r["overall_rate"] or 0) * 100,
                        (r["left_rate"] or 0) * 100, (r["right_rate"] or 0) * 100, r["unit_count"]]
                       for category, group in groups
                       for item in GuangdongStatistics.marking_typical_segments(group["marking"], limit=3)
                       for r in item["km_items"]]))

    added.append(fill(new_sheet("高度最低公里段", ["道路类别", "路线", "方向", "公里", "总体合格率(%)", "检测点数"]),
                      [[category, r["route"], r["direction"], r["km"], (r["rate"] or 0) * 100, r["count"]]
                       for category, group in groups
                       for r in sorted([x for x in GuangdongStatistics.height_per_km(group["height"]) if x["rate"] is not None],
                                       key=lambda x: x["rate"])]))

    added.append(fill(new_sheet("螺栓严重缺失公里段", ["道路类别", "路线", "方向", "公里", "单处严重缺失（处）", "缺失率(%)"]),
                      [[category, r["route"], r["direction"], r["km"], r["severe"], (r["rate"] or 0) * 100]
                       for category, group in groups
                       for r in GuangdongStatistics.bolt_per_km(group["bolt"]) if r["severe"]]))

    return added


def write_guangdong_chart_workbook(bundle, output_dir, log=lambda _: None):
    """生成专门的图表统计工作簿：护栏中心高度、护栏螺栓缺失、标线逆反射各占一个工作表。

    图表为可编辑的原生Excel图表；配色、线条粗细、图例位置与标题文字均与Word文档插图一致。
    """
    city = _safe_city_component(bundle["city"])
    folder = Path(output_dir) / city; folder.mkdir(parents=True, exist_ok=True)
    write_guangdong_route_workbook(bundle, output_dir, log)
    path = folder / f"{city}交安设施统计图表.xlsx"
    thin = Side(style="thin", color="B7B7B7")

    def font_axes(chart):
        for axis in (chart.x_axis, chart.y_axis):
            try:
                axis.txPr = RichText(p=[Paragraph(
                    pPr=ParagraphProperties(defRPr=CharacterProperties(
                        latin=DrawingFont(typeface="Times New Roman"), ea=DrawingFont(typeface="宋体"), sz=700)),
                    endParaRPr=CharacterProperties(sz=700))])
            except Exception:
                pass

    def line_style(series, color, width_emu, dash=None):
        series.graphicalProperties.line.solidFill = color
        series.graphicalProperties.line.width = width_emu
        if dash:
            series.graphicalProperties.line.dashStyle = dash
        series.smooth = False

    def header_cells(ws, row, headers, start_col=1):
        for offset, text in enumerate(headers):
            cell = ws.cell(row, start_col + offset, text)
            cell.font = Font(name="宋体", size=10, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def seg_text(segment):
        return str(segment) + ("" if str(segment).endswith("段") or not str(segment) else "段")

    wb = Workbook(); wb.remove(wb.active)
    counts = {"高度折线": 0, "高度分布": 0, "螺栓柱状": 0, "标线折线": 0, "四段式柱状": 0, "逐公里折线": 0}

    # ---- 工作表1：护栏中心高度（逐区段、逐类型折线图+分布饼图，与Word插图同款） ----
    ws = wb.create_sheet("护栏中心高度")
    groups = {}
    for r in bundle.get("height", []):
        kind = _guardrail_type(r.get("guardrail_type"))
        value = _float(r.get("height"))
        if kind not in ("二波", "三波") or value is None or r.get("station_m") is None:
            continue
        groups.setdefault((str(r.get("segment") or ""), kind), []).append((r["station_m"], value))
    line_slot = pie_slot = 2
    for (segment, kind), points in sorted(groups.items()):
        points.sort(key=lambda item: item[0])
        if not points:
            continue
        standards = (580, 620) if kind == "二波" else (677, 717)
        limits = (550, 580, 620, 650) if kind == "二波" else (647, 677, 717, 747)
        header_row = line_slot
        header_cells(ws, header_row, [
            "桩号", "梁板中心高度(mm)", f"标准值（{standards[0]}mm）", f"标准值（{standards[1]}mm）",
            "高度区间", "点数",
        ])
        bin_labels = [f"h＜{limits[0]}", f"{limits[0]}≤h＜{limits[1]}", f"{limits[1]}≤h≤{limits[2]}",
                      f"{limits[2]}＜h≤{limits[3]}", f"h＞{limits[3]}"]
        bins = [0] * 5
        for index, (station, value) in enumerate(points, 1):
            row = header_row + index
            ws.cell(row, 1, format_station(station))
            ws.cell(row, 2, value)
            ws.cell(row, 3, standards[0])
            ws.cell(row, 4, standards[1])
            if value < limits[0]: bins[0] += 1
            elif value < limits[1]: bins[1] += 1
            elif value <= limits[2]: bins[2] += 1
            elif value <= limits[3]: bins[3] += 1
            else: bins[4] += 1
        for index, (label, count) in enumerate(zip(bin_labels, bins), 1):
            ws.cell(header_row + index, 5, label)
            ws.cell(header_row + index, 6, count)
        last = header_row + len(points)

        chart = LineChart(); chart.visible_cells_only = False
        chart.add_data(Reference(ws, min_col=2, max_col=4, min_row=header_row, max_row=last), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=header_row + 1, max_row=last))
        chart.height, chart.width = 8, 13
        chart.y_axis.scaling.min, chart.y_axis.scaling.max, chart.y_axis.majorUnit = 300, 850, 50
        chart.x_axis.delete = chart.y_axis.delete = False
        chart.legend.position = "b"
        line_style(chart.series[0], GD_SERIES_COLORS[0], 12700)
        line_style(chart.series[1], GD_SERIES_COLORS[1], 25400)
        line_style(chart.series[2], GD_SERIES_COLORS[2], 25400)
        ws.add_chart(chart, f"H{line_slot}")
        counts["高度折线"] += 1; line_slot += 17

        pie = PieChart()
        pie.add_data(Reference(ws, min_col=6, min_row=header_row, max_row=header_row + 5), titles_from_data=True)
        pie.set_categories(Reference(ws, min_col=5, min_row=header_row + 1, max_row=header_row + 5))
        pie.height, pie.width = 8.5, 14
        pie.legend.position = "r"
        pie.dataLabels = DataLabelList()
        pie.dataLabels.showVal, pie.dataLabels.showPercent = False, True
        pie.dataLabels.numFmt = "0.00%"
        pie.dataLabels.showCatName = pie.dataLabels.showSerName = pie.dataLabels.showLegendKey = False
        pie.series[0].data_points = [DataPoint(idx=i, spPr=GraphicalProperties(solidFill=color)) for i, color in enumerate(PIE_COLORS)]
        zero_labels = []
        for index, count in enumerate(bins):
            if not count:
                label = DataLabel(idx=index); label.delete = True; zero_labels.append(label)
        if zero_labels:
            pie.dataLabels.dLbl = zero_labels
        ws.add_chart(pie, f"S{pie_slot}")
        counts["高度分布"] += 1; pie_slot = max(pie_slot + 17, header_row + len(points) + 3)

    # ---- 工作表2：护栏螺栓缺失（区段汇总表+逐区段柱状图，柱色与Word一致） ----
    ws = wb.create_sheet("护栏螺栓缺失")
    headers = ["路线", "方向", "检测区段", "拼接螺栓（颗）", "连接螺栓（颗）", "缺失数量（颗）"]
    header_cells(ws, 1, headers)
    summary_rows = {}
    for r in bundle.get("bolt", []):
        key = (str(r.get("route") or ""), str(r.get("direction") or ""), str(r.get("segment") or ""))
        item = summary_rows.setdefault(key, {"splice": 0, "connection": 0, "missing": 0})
        item["splice"] += float(r.get("splice", 0) or 0) + 0
        item["connection"] += float(r.get("connection", 0) or 0)
        item["missing"] += float(r.get("splice_missing", 0) or 0) + float(r.get("connection_missing", 0) or 0)
    for index, (key, item) in enumerate(sorted(summary_rows.items()), 2):
        ws.cell(index, 1, key[0]); ws.cell(index, 2, key[1]); ws.cell(index, 3, key[2])
        ws.cell(index, 4, int(round(item["splice"]))); ws.cell(index, 5, int(round(item["connection"]))); ws.cell(index, 6, int(round(item["missing"])))
    for column, width in zip("ABCDEF", (10, 10, 26, 15, 15, 15)):
        ws.column_dimensions[column].width = width
    bar_slot = 2
    for index in range(2, len(summary_rows) + 2):
        chart = BarChart(); chart.type = "col"
        chart.add_data(Reference(ws, min_col=4, max_col=6, min_row=index, max_row=index), from_rows=True)
        chart.set_categories(Reference(ws, min_col=4, max_col=6, min_row=1))
        chart.height, chart.width = 8, 13
        chart.legend = None
        chart.x_axis.delete = chart.y_axis.delete = False
        chart.y_axis.title = "螺栓数量（颗）"
        chart.gapWidth = 80
        chart.series[0].data_points = [DataPoint(idx=i, spPr=GraphicalProperties(solidFill=color))
                                       for i, color in enumerate((GD_SERIES_COLORS[0], GD_SERIES_COLORS[1], "FF0000"))]
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        chart.dataLabels.dLblPos = "outEnd"
        ws.add_chart(chart, f"H{bar_slot}")
        counts["螺栓柱状"] += 1; bar_slot += 17

    # ---- 工作表3：标线逆反射；源桩号/行号及实际目标随每个点导出 ----
    ws = wb.create_sheet("标线逆反射")
    marking_groups = GuangdongStatistics._group(bundle.get("marking", []), (*MARKING_SEGMENT_FIELDS, "marking_position"))
    slot = 2
    for key, rows in sorted(marking_groups.items()):
        points = sorted((r for r in rows if _float(r.get("value")) is not None and r.get("station_m") is not None), key=lambda r: r["station_m"])
        if not points:
            continue
        header_row = slot
        header_cells(ws, header_row, ["桩号", "逆反射亮度系数", "目标值", "地市", "道路类别", "路线", "方向", "管养单位", "检测区段", "标线位置", "源文件", "源工作表", "源行号", "源桩号范围", "起桩米数", "止桩米数"])
        for index, point in enumerate(points, 1):
            values = [format_station_one_decimal(point["station_m"]), point["value"], point.get("target", 80),
                      *[point.get(field) for field in (*MARKING_SEGMENT_FIELDS, "marking_position", "source", "sheet", "source_row", "source_range", "station_m", "end_m")]]
            for column, value in enumerate(values, 1):
                ws.cell(header_row + index, column, _safe_excel_value(value))
        last = header_row + len(points)
        chart = LineChart(); chart.visible_cells_only = False
        chart.add_data(Reference(ws, min_col=2, max_col=3, min_row=header_row, max_row=last), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=header_row + 1, max_row=last))
        chart.title = " ".join(str(x) for x in key[2:])
        chart.height, chart.width = 8, 13
        chart.x_axis.delete = chart.y_axis.delete = False
        chart.y_axis.title = "逆反射亮度系数"
        chart.legend.position = "b"
        line_style(chart.series[0], GD_SERIES_COLORS[0], 12700)
        if len(points) == 1: chart.series[0].marker.symbol = "circle"
        line_style(chart.series[1], GD_SERIES_COLORS[1], 25400, dash="dash")
        ws.add_chart(chart, f"R{slot}")
        counts["标线折线"] += 1; slot = max(slot + 17, last + 2)

    # ---- 工作表4：GD03 四段式分组柱状图与逐公里折线图（与Word插图同款数据/配色） ----
    ws = wb.create_sheet("四段式图表")

    def add_grouped(ws, slot, headers, rows, colors, ylabel="合格率（%）", ymax=100):
        if not rows:
            return slot
        header_cells(ws, slot, headers)
        for index, row in enumerate(rows, 1):
            for column, value in enumerate(row, 1):
                ws.cell(slot + index, column, value)
        last = slot + len(rows)
        chart = BarChart(); chart.type = "col"; chart.grouping = "clustered"
        chart.add_data(Reference(ws, min_col=2, max_col=len(headers), min_row=slot, max_row=last), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=slot + 1, max_row=last))
        chart.height, chart.width = 8, 13
        chart.y_axis.scaling.min, chart.y_axis.scaling.max = 0, ymax
        chart.x_axis.delete = chart.y_axis.delete = False
        chart.legend.position = "tr"
        chart.y_axis.title = ylabel
        chart.dataLabels = DataLabelList(); chart.dataLabels.showVal = True
        chart.dataLabels.numFmt = "0.0"; chart.dataLabels.dLblPos = "outEnd"
        for series, color in zip(chart.series, colors):
            series.graphicalProperties.solidFill = color
            series.graphicalProperties.line.solidFill = color
        ws.add_chart(chart, f"H{slot}")
        counts["四段式柱状"] += 1
        return max(slot + 17, last + 2)

    def add_line(ws, slot, headers, rows, colors, ylabel="合格率（%）", ymax=100):
        if not rows:
            return slot
        header_cells(ws, slot, headers)
        for index, row in enumerate(rows, 1):
            for column, value in enumerate(row, 1):
                ws.cell(slot + index, column, value)
        last = slot + len(rows)
        chart = LineChart(); chart.visible_cells_only = False
        chart.add_data(Reference(ws, min_col=2, max_col=len(headers), min_row=slot, max_row=last), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=slot + 1, max_row=last))
        chart.height, chart.width = 8, 13
        chart.y_axis.scaling.min, chart.y_axis.scaling.max = 0, ymax
        chart.x_axis.delete = chart.y_axis.delete = False
        chart.legend.position = "tr"
        chart.y_axis.title = ylabel
        for series, color in zip(chart.series, colors):
            line_style(series, color, 12700)
            series.marker.symbol = "circle"
        ws.add_chart(chart, f"H{slot}")
        counts["逐公里折线"] += 1
        return max(slot + 17, last + 2)

    slot = 2
    for category in ("高速公路", "普通国省道"):
        mark = [r for r in bundle.get("marking", []) if r.get("category") == category]
        slot = add_grouped(ws, slot, ["路线（管养单位）", "左侧标线", "右侧标线", "总体"],
                           [[f"{r['route']}{manager_display(r.get('manager'), city)}", (r["left_rate"] or 0) * 100,
                             (r["right_rate"] or 0) * 100, (r["overall_rate"] or 0) * 100]
                            for r in GuangdongStatistics.marking_route_units(mark)],
                           GD_SERIES_COLORS)
        for item in GuangdongStatistics.marking_typical_segments(mark, limit=3):
            slot = add_line(ws, slot, ["公里", "左侧标线", "右侧标线", "总体"],
                            [[str(r["km"]), (r["left_rate"] or 0) * 100, (r["right_rate"] or 0) * 100,
                              (r["overall_rate"] or 0) * 100] for r in item["km_items"]],
                            GD_SERIES_COLORS)
        height = [r for r in bundle.get("height", []) if r.get("category") == category]
        slot = add_grouped(ws, slot, ["路线（管养单位）", "二波护栏", "三波护栏", "总体"],
                           [[f"{r['route']}{manager_display(r.get('manager'), city)}", (r["二波"]["rate"] or 0) * 100,
                             (r["三波"]["rate"] or 0) * 100, (r["rate"] or 0) * 100]
                            for r in GuangdongStatistics.height_route_units(height)],
                           GD_SERIES_COLORS)
        bolt = [r for r in bundle.get("bolt", []) if r.get("category") == category]
        bolt_rows = GuangdongStatistics.bolt_route_units(bolt)
        bolt_max = max([(r["rate"] or 0) * 100 for r in bolt_rows] or [5]) * 1.2
        slot = add_grouped(ws, slot, ["路线（管养单位）", "缺失率"],
                           [[f"{r['route']}{manager_display(r.get('manager'), city)}", (r["rate"] or 0) * 100] for r in bolt_rows],
                           (GD_SERIES_COLORS[0],), ylabel="螺栓缺失率（%）", ymax=max(5.0, bolt_max))
    for column, width in zip("ABCD", (24, 12, 12, 12)):
        ws.column_dimensions[column].width = width

    summary = wb.create_sheet("标线区段汇总")
    header_cells(summary, 1, ["地市", "道路类别", "路线", "方向", "管养单位", "检测区段", "标线位置", "有效单元数", "均值", "合格单元数", "合格率"])
    for item in GuangdongStatistics.marking_segment_summary(bundle.get("marking", [])):
        summary.append([_safe_excel_value(item.get(field)) for field in (*MARKING_SEGMENT_FIELDS, "marking_position", "valid_count", "average", "qualified_count", "qualified_rate")])

    for column, width in (("A", 14), ("B", 18), ("C", 18), ("D", 18), ("E", 16), ("F", 10)):
        ws.column_dimensions[column].width = width
        wb["护栏中心高度"].column_dimensions[column].width = width
    # D16/Q1：正文删除的清单/分档/长连续数据全部落进本工作簿（信息不丢）
    for name, count in write_gd_removed_data_sheets(bundle, wb):
        log(f"统计工作簿已补工作表：{name}（{count} 行）")
    normalize_workbook_separators(wb)
    try:
        wb.save(path)
    except PermissionError as exc:
        raise PermissionError(f"Excel文件被占用：{path}") from exc
    log(f"图表统计工作簿已生成：{path}（" + "，".join(f"{name}{count}个" for name, count in counts.items()) + "）")
    return path


class GuangdongBatchRunner:
    @staticmethod
    def add_weak_segments(bundle):
        weak=[]
        for row in GuangdongStatistics.continuous_marking_weak(bundle.get("marking",[])):
            weak.append({"type":"标线连续3km不合格",**row,"segment":f"{format_station(row['start_m'])}-{format_station(row['end_m'])}","reason":"同一路线、方向和标线位置连续不合格长度达到3 km"})
        for row in GuangdongStatistics.height_segment_summary(bundle.get("height",[])):
            if (row.get("over_10cm_count") or 0) > 0:
                weak.append({"type":"护栏高度偏差超10cm","route":row.get("route"),"direction":row.get("direction"),"segment":row.get("segment"),"reason":f"{_guardrail_type(row.get('guardrail_type'))}护栏偏差超10 cm点数{row.get('over_10cm_count')}个"})
        for row in GuangdongStatistics.bolt_segment_summary(bundle.get("bolt",[])):
            if (row.get("missing_rate") or 0) > 0.05:
                weak.append({"type":"螺栓缺失率超5%","route":row.get("route"),"direction":row.get("direction"),"segment":row.get("segment"),"reason":f"螺栓缺失率{row.get('missing_rate'):.2%}"})
        bundle["weak_segments"]=weak
        return bundle

    @staticmethod
    def run_bundles(bundles,output_dir,template,thresholds,log=lambda _x:None):
        result={"success":[],"failed":{},"warnings":[]}
        for city,bundle in bundles.items():
            try:
                if bundle.get("_error"):
                    raise ValueError(bundle["_error"])
                log(f"生成{city}数据图表")
                guangdong_report_images(bundle, output_dir, log)
                log(f"生成{city}第五章Word")
                GuangdongChapterWriter.write(city,bundle,output_dir,template,thresholds)
                log(f"生成{city}图表统计工作簿")
                write_guangdong_chart_workbook(bundle,output_dir,log)
                charts_dir=Path(output_dir)/_safe_city_component(city)/"charts"
                if charts_dir.is_dir():
                    shutil.rmtree(charts_dir,ignore_errors=True)
                    log(f"已清理临时图片文件夹：{charts_dir}")
                result["success"].append(city)
            except Exception as exc:
                result["failed"][city]=str(exc); log(f"{city}失败：{exc}")
        return result

    @staticmethod
    def build_bundles(scanned,route_index,manual_records=None,thresholds=None):
        norm=RouteCategoryIndex._norm_city
        labels={}
        for kind in ("height","bolt","marking","notes"):
            for row in scanned.get(kind,[]):
                labels.setdefault(norm(row["city"]),[]).append(str(row["city"]).strip())
        # P1d-b：显示层市名统一带「市」（模板取证：正文「共抽检XX市高速公路…」、表行标签「XX市合计」、
        # 产物文件名/子目录名同为带「市」形式）。分类表「地市」列与管养单位名不受影响。
        cities={key:RouteCategoryIndex.city_display(next((x for x in values if x.endswith("市")),values[0]))
                for key,values in labels.items()}
        bundles={}
        comparator=ManualAutoComparator(thresholds or {"marking":0,"height":0,"bolt":0})
        for key in sorted(cities):
            city=cities[key]
            bundle={"city":city,"route_rows":route_index.rows(city),"issues":list(scanned.get("issues",[])),"weak_segments":[],"unknown_categories":[],"route_index":route_index}
            unknown={}
            try:
                for kind in ("height","bolt","marking","notes"):
                    bundle[kind]=[dict(r) for r in scanned.get(kind,[]) if norm(r["city"])==key]
                    for row in bundle[kind]:
                        row["city"]=city
                        # R2：记录已按文件夹归属计入，路线可能不在本市路线表里。兜底绝不抛异常，
                        # 无法判定类别的行仍归属该市并记入 issues，不中断整市报告。
                        row["category"]=route_index.category_for_records(city,row["route"])
                        if row["category"] is None:
                            unknown[(kind,row["route"])]=unknown.get((kind,row["route"]),0)+1
                for (kind,route),count in sorted(unknown.items()):
                    bundle["unknown_categories"].append({"kind":kind,"route":route,"count":count})
                    bundle["issues"].append(f"{city}：路线{route}不在本市路线表且无法按路线号唯一判定道路类别，{count} 条{kind}记录已按文件夹归属{city}计入明细（类别留空，报告不中断）")
                GuangdongStatistics.group_marking_ranges(bundle["marking"])
                detail,summary=comparator.compare([r for r in (manual_records or []) if norm(r.get("city"))==key]); bundle["comparison_detail"]=detail; bundle["comparison_summary"]=summary
                for row in detail: row["category"]=route_index.category_for_records(row.get("city"),row.get("route"))
                GuangdongBatchRunner.add_weak_segments(bundle)
            except Exception as exc:
                bundle["_error"]=str(exc)
                for kind in ("height","bolt","marking","notes"):
                    bundle.setdefault(kind,[])
                bundle.setdefault("comparison_detail",[]); bundle.setdefault("comparison_summary",{}); bundle.setdefault("unknown_categories",[])
            bundles[city]=bundle
        return bundles


PROJECT_TEMPLATES = ("重庆项目模板", "广东项目模板")



def load_route_names(path):
    """交工路表「公路路线基本情况明细表」→ {路线编号: 路线名称}。

    ponytail: 只取“路线编号/路线名称”两列，表头按表内文字定位（多层表头无需写死行号）；
    需要按县级区划细分名称或补充管养单位时再扩列。
    """
    if not path:
        return {}
    names = {}
    workbook = openpyxl.load_workbook(Path(path), read_only=True, data_only=True)
    for sheet in workbook.worksheets:
        code_col = name_col = None
        for row in sheet.iter_rows(values_only=True):
            cells = ["" if value is None else str(value).replace("\n", "").strip() for value in row]
            if code_col is None:
                if "路线编号" in cells and "路线名称" in cells:
                    code_col, name_col = cells.index("路线编号"), cells.index("路线名称")
                continue
            code = cells[code_col] if code_col < len(cells) else ""
            name = cells[name_col] if name_col < len(cells) else ""
            if code and name:
                names.setdefault(code, name)
        if names:
            break
    workbook.close()
    return names


def detect_route_name_workbook(project_dir):
    """在项目资料里查找交工路表；命名未命中返回 None（报告该列显示 —）。"""
    root = Path(project_dir)
    if not root.is_dir():
        return None
    for path in sorted(root.rglob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        if "交公路" in path.name or "路线基本情况" in path.name:
            return path
    return None


def gd_road_class_subsets(rows, category):
    """①总体情况按附件分道路类别：高速一类；普通国省道分国道(G)/省道(S)。"""
    if category != "普通国省道":
        return [("高速公路", rows)]
    upper = lambda row: str(row.get("route") or "").upper()
    return [("普通国省道", rows),
            ("普通国道", [row for row in rows if upper(row).startswith("G")]),
            ("普通省道", [row for row in rows if upper(row).startswith("S")])]


def guardrail_missing_warning(project_dir, scanned):
    """禁止静默 0：高度/螺栓都解析为 0、而项目内确实存在含护栏表头的明细时返回警告（含所在目录），否则 None。

    ponytail: 只在「解析结果全 0」这一条路径上做一次表头嗅探，正常有数据的项目零开销。
    """
    if scanned.get("height") or scanned.get("bolt"):
        return None
    _marking_dir, guardrail_dir = detect_guangdong_data_folders(project_dir)
    if not guardrail_dir:
        return None
    return f"警告：护栏数据源解析为 0 条，但 {guardrail_dir} 内存在护栏表头文件，请核对该目录的数据源选择"


def run_guangdong_project(config, log=lambda _x: None):
    for label,path,kind in (("项目资料",config.project_dir,"dir"),("人工自动化对比表",config.manual_xlsx,"file"),("路线分类表",config.route_xlsx,"file")):
        exists=path.is_dir() if kind=="dir" else path.is_file()
        if not exists: raise FileNotFoundError(f"{label}不存在：{path}")
    if config.marking_dir and not Path(config.marking_dir).exists():
        raise FileNotFoundError(f"标线统计表不存在：{config.marking_dir}")
    if config.guardrail_dir and not config.guardrail_dir.is_dir():
        raise FileNotFoundError(f"护栏数据文件夹不存在：{config.guardrail_dir}")
    route_index=merge_owner_table(RouteCategoryIndex.from_file(config.route_xlsx),config.route_xlsx,
                                 getattr(config,"owner_xlsx",None),log); log("路线分类表读取完成")

    scanned = {"height": [], "bolt": [], "marking": [], "notes": [], "issues": []}
    if config.marking_dir or config.guardrail_dir:
        marking_dir = config.marking_dir
        guardrail_dir = config.guardrail_dir
        if marking_dir and guardrail_dir and Path(marking_dir) == Path(guardrail_dir):
            log(f"扫描数据文件夹：{marking_dir}")
            scanner = GuangdongInputScanner(marking_dir, route_index, log=log)
            partial = scanner.scan()
            scanned["marking"].extend(own_city_marking(partial["marking"], log, scanner.default_city))
            scanned["height"].extend(partial["height"])
            scanned["bolt"].extend(partial["bolt"])
            scanned["notes"].extend(partial["notes"])
            scanned["issues"].extend(partial["issues"])
        else:
            if marking_dir:
                log(f"扫描标线数据源：{marking_dir}")
                scanner = GuangdongInputScanner(marking_dir, route_index, log=log)
                partial = scanner.scan()
                scanned["marking"].extend(own_city_marking(partial["marking"], log, scanner.default_city))
                scanned["issues"].extend(partial["issues"])
            if guardrail_dir:
                log(f"扫描护栏数据文件夹：{guardrail_dir}")
                partial = GuangdongInputScanner(guardrail_dir, route_index, log=log).scan()
                scanned["height"].extend(partial["height"])
                scanned["bolt"].extend(partial["bolt"])
                scanned["notes"].extend(partial["notes"])
                scanned["issues"].extend(partial["issues"])
    else:
        marking_sources, guardrail_sources = detect_guangdong_sources(config.project_dir)
        if marking_sources or guardrail_sources:
            for kind, sources in (("marking", marking_sources), ("guardrail", guardrail_sources)):
                for source in sources:
                    log(f"扫描{'标线' if kind=='marking' else '护栏'}数据源：{source}")
                    scanner = GuangdongInputScanner(source, route_index, log=log)
                    partial = scanner.scan()
                    if kind == "marking":
                        scanned["marking"].extend(own_city_marking(partial["marking"], log, scanner.default_city))
                    else:
                        scanned["height"].extend(partial["height"])
                        scanned["bolt"].extend(partial["bolt"])
                        scanned["notes"].extend(partial["notes"])
                    scanned["issues"].extend(partial["issues"])
        else:
            scanned = GuangdongInputScanner(config.project_dir, route_index, log=log).scan()

    log(f"数据识别完成：标线{len(scanned['marking'])}、高度{len(scanned['height'])}、螺栓{len(scanned['bolt'])}")
    warning = guardrail_missing_warning(config.project_dir, scanned)
    if warning:
        log(warning)
        scanned["issues"].append(warning)
    if not any(scanned[k] for k in ("marking","height","bolt")): raise ValueError("未识别到任何有效数据")
    manual,manual_issues=ManualAutoComparator.read_file(config.manual_xlsx); scanned["issues"].extend(manual_issues)
    # R2：人工复核对比明细行同样以所在市文件夹定市（“地市”列不再决定归属），否则 build_bundles
    # 的 norm(city)==key 过滤会把整批行丢掉。
    manual_city=folder_city_for(config.manual_xlsx, route_index)
    if manual_city:
        for row in manual: row["city"]=manual_city
    # P1n-2：进场前对比表此前**只**用于「人工复核覆盖里程」（load_manual_review_spans），
    # 其明细行从未进 manual → 三张人工复核对比明细表整支缺进场前数据。
    # 实测 15 市进场前覆盖率全为 0%（不是 3 市）：进场后工作簿只是「碰巧」含同名路线
    # （中山 G105/S268），桩号段与进场前那份完全不同，按路线名比对会误判为「两支都有」。
    # 两支按并集读入：进场前/进场后是同一路线的两次独立测量（惠州 G324、云浮 S265 各有
    # 1 段桩号相同但人工/自动测值与护栏型式都不同，是两次测量而非重复行），故不按段去重。
    before_xlsx=config.manual_before_xlsx or detect_manual_before_workbook(config.project_dir)
    if before_xlsx and Path(before_xlsx)!=Path(config.manual_xlsx):
        before,before_issues=ManualAutoComparator.read_file(before_xlsx); scanned["issues"].extend(before_issues)
        before_city=folder_city_for(before_xlsx, route_index)
        if before_city:
            for row in before: row["city"]=before_city
        seen={(r.get("indicator"),r.get("route"),r.get("direction"),r.get("segment"),
               r.get("manual"),r.get("automatic")) for r in manual}
        manual.extend(row for row in before
                      if (row.get("indicator"),row.get("route"),row.get("direction"),row.get("segment"),
                          row.get("manual"),row.get("automatic")) not in seen)
        log(f"进场前人工复核对比明细并入 {len(before)} 行（{before_xlsx}）")
    route_name_path=config.route_name_xlsx or detect_route_name_workbook(config.project_dir)
    route_names=load_route_names(route_name_path) if route_name_path else {}
    if route_names: log(f"路线名称表读取完成：{route_name_path}（{len(route_names)}条路线）")
    bundles=GuangdongBatchRunner.build_bundles(scanned,route_index,manual,config.thresholds)
    # R2：文件夹定市后可能混入不在本市路线表、又无法按路线号唯一判别的记录；口径变化如实暴露到日志。
    unknown_rows=sum(item["count"] for bundle in bundles.values() for item in bundle.get("unknown_categories",[]))
    if unknown_rows:
        detail="；".join(f"{bundle['city']}/{item['kind']}/{item['route']} {item['count']} 行"
                        for bundle in bundles.values() for item in bundle.get("unknown_categories",[]))
        log(f"无法判定道路类别的记录 {unknown_rows} 行（记入对应市 issues，报告未中断）：{detail}")
    else:
        log("无法判定道路类别的记录 0 行")
    route_segments=load_route_segments(config.route_xlsx)
    if route_segments: log(f"路线表起止桩号读取完成：{config.route_xlsx}（{len(route_segments)}个抽检路段）")
    review_before=load_manual_review_spans(before_xlsx)
    review_after=load_manual_review_spans(config.manual_xlsx)
    if review_before: log(f"进场前人工复核覆盖读取完成（{len(review_before)}个路段）")
    for bundle in bundles.values():
        bundle["route_names"]=route_names
        bundle["route_segments"]=route_segments
        bundle["review_phases"]={"进场前": review_before, "抽检后": review_after}
    # P1d-b：本次项目目录的实测抽检段（明细文件名 K 区间），供分类表导出按桩号消歧经营主体
    route_index.measured=load_measured_segments(config.project_dir, route_index)
    write_guangdong_route_workbook_all(route_index,route_names,config.output_dir,log)
    template=resource_template_path("广东项目第五章模板.md")
    return GuangdongBatchRunner.run_bundles(bundles,config.output_dir,template,config.thresholds,log)


class ReportGeneratorApp(tk.Tk):
    """双模板统一 GUI；重庆逻辑调用原基线，广东逻辑调用批量运行器。"""
    def __init__(self):
        super().__init__(); self.title(PROGRAM_NAME); self.geometry("980x860"); self.minsize(880,720)
        self.project_template=tk.StringVar(value=PROJECT_TEMPLATES[0]); self.queue=Queue(); self.running=False
        keys=("cq_project","cq_summary","cq_detail","cq_disease","cq_tci","cq_output","gd_project","gd_marking_dir","gd_guardrail_dir","gd_manual","gd_route","gd_output","gd_marking","gd_height","gd_bolt")
        self.vars={k:tk.StringVar() for k in keys}
        self.vars["gd_marking"].set("7"); self.vars["gd_height"].set("5"); self.vars["gd_bolt"].set("5")
        self._build(); self.after(100,self._poll)

    def _pick(self,key,file=False):
        path=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")]) if file else filedialog.askdirectory()
        if path:
            self.vars[key].set(path)
            if key == "gd_project":
                self._auto_detect_gd_folders(path)

    def _auto_detect_gd_folders(self, project_dir):
        """自动识别导入源。多市输入时不预填，交由引擎按市逐个识别，避免只处理一个市。"""
        self._append("正在自动识别标线和护栏数据源...")
        try:
            marking, guardrail = detect_guangdong_sources(project_dir)
            if len(marking) <= 1 and len(guardrail) <= 1:
                if marking: self.vars["gd_marking_dir"].set(str(marking[0]))
                if guardrail: self.vars["gd_guardrail_dir"].set(str(guardrail[0]))
            suffix = "" if len(marking) <= 1 and len(guardrail) <= 1 else "（多市输入，运行时按市识别）"
            self._append(f"识别完成：标线源{len(marking)}个，护栏源{len(guardrail)}个{suffix}")
        except Exception as exc:
            self._append(f"自动识别失败：{exc}")

    def _row(self,parent,row,label,key,file=False):
        ttk.Label(parent,text=label,width=23).grid(row=row,column=0,sticky="w",pady=4)
        is_threshold = key in ("gd_marking","gd_height","gd_bolt")
        is_editable_path = key in ("gd_marking_dir","gd_guardrail_dir")
        state = "normal" if (is_threshold or is_editable_path) else "readonly"
        ttk.Entry(parent,textvariable=self.vars[key],state=state).grid(row=row,column=1,sticky="ew",padx=8)
        if not is_threshold: ttk.Button(parent,text="选择",command=lambda:self._pick(key,file),width=10).grid(row=row,column=2)

    def _build(self):
        root=ttk.Frame(self,padding=16); root.pack(fill="both",expand=True)
        ttk.Label(root,text=PROGRAM_NAME,font=("Microsoft YaHei UI",16,"bold")).pack(anchor="w")
        select=ttk.LabelFrame(root,text="项目模板",padding=8); select.pack(fill="x",pady=10)
        for name in PROJECT_TEMPLATES: ttk.Radiobutton(select,text=name,value=name,variable=self.project_template,command=self._switch).pack(side="left",padx=12)
        self.forms=ttk.Frame(root); self.forms.pack(fill="x")
        self.cq=ttk.LabelFrame(self.forms,text="重庆项目输入",padding=10); self.gd=ttk.LabelFrame(self.forms,text="广东项目输入",padding=10)
        for i,(label,key,file) in enumerate((("项目资料文件夹","cq_project",False),("分段汇总表","cq_summary",True),("检测明细文件夹","cq_detail",False),("病害清单文件夹","cq_disease",False),("TCI病害清单","cq_tci",True),("输出文件夹","cq_output",False))): self._row(self.cq,i,label,key,file)
        for i,(label,key,file) in enumerate((("项目资料文件夹","gd_project",False),("标线统计表","gd_marking_dir",True),("护栏数据文件夹","gd_guardrail_dir",False),("人工自动化对比表","gd_manual",True),("路线分类表","gd_route",True),("输出文件夹","gd_output",False),("标线一致性阈值（%）","gd_marking",False),("护栏高度一致性阈值（mm）","gd_height",False),("螺栓缺失数量差值阈值（颗/柱）","gd_bolt",False))): self._row(self.gd,i,label,key,file)
        self.cq.columnconfigure(1,weight=1); self.gd.columnconfigure(1,weight=1); self._switch()
        actions=ttk.Frame(root); actions.pack(fill="x",pady=10); self.run_button=ttk.Button(actions,text="开始运行",command=self.start); self.run_button.pack(side="left"); ttk.Button(actions,text="打开输出文件夹",command=self.open_output).pack(side="left",padx=8)
        self.progress=ttk.Progressbar(root,mode="indeterminate"); self.progress.pack(fill="x")
        self.log=scrolledtext.ScrolledText(root,state="disabled",height=16); self.log.pack(fill="both",expand=True,pady=8)

    def _switch(self):
        self.cq.pack_forget(); self.gd.pack_forget(); (self.gd if self.project_template.get()==PROJECT_TEMPLATES[1] else self.cq).pack(fill="x")

    def start(self):
        if self.running:return
        try:
            if self.project_template.get()==PROJECT_TEMPLATES[1]:
                cfg=GuangdongConfig(Path(self.vars['gd_project'].get()),Path(self.vars['gd_manual'].get()),Path(self.vars['gd_route'].get()),Path(self.vars['gd_output'].get()),self.vars['gd_marking'].get(),self.vars['gd_height'].get(),self.vars['gd_bolt'].get(),marking_dir=Path(self.vars['gd_marking_dir'].get()) if self.vars['gd_marking_dir'].get() else None,guardrail_dir=Path(self.vars['gd_guardrail_dir'].get()) if self.vars['gd_guardrail_dir'].get() else None)
                worker=lambda:run_guangdong_project(cfg,lambda x:self.queue.put(("log",x)))
            else:
                base=Path(self.vars['cq_project'].get() or "."); cfg=Config(base,Path(self.vars['cq_summary'].get()),Path(self.vars['cq_detail'].get()),BUILTIN_REPORT_TEMPLATES[next(iter(BUILTIN_REPORT_TEMPLATES))],Path(self.vars['cq_output'].get() or base),Path(self.vars['cq_disease'].get()) if self.vars['cq_disease'].get() else None, Path(self.vars['cq_tci'].get()) if self.vars.get('cq_tci') and self.vars['cq_tci'].get() else None)
                # TCI 触发：若填写了 TCI 路径则处理沿线设施
                tci_on = bool(self.vars.get('cq_tci') and self.vars['cq_tci'].get() and Path(self.vars['cq_tci'].get()).exists())
                worker=lambda tci_on=tci_on:generate_statistics_and_report(cfg,lambda x:self.queue.put(("log",x)),process_height=True,process_bolts=True, process_alongline=tci_on, process_tci=tci_on)
        except Exception as exc:messagebox.showerror("参数错误",str(exc));return
        self.running=True;self.run_button.configure(state="disabled");self.progress.start(10)
        def run():
            try:self.queue.put(("done",worker()))
            except Exception as exc:self.queue.put(("error",f"{exc}\n{traceback.format_exc()}"))
        Thread(target=run,daemon=True).start()

    def _poll(self):
        try:
            while True:
                kind,value=self.queue.get_nowait()
                if kind=="log":self._append(value)
                else:
                    self.running=False;self.progress.stop();self.run_button.configure(state="normal")
                    if kind=="done":self._append(f"完成：{value}");messagebox.showinfo("完成",str(value))
                    else:self._append(f"错误：{value}");messagebox.showerror("运行失败",str(value))
        except Empty:pass
        self.after(100,self._poll)

    def _append(self,text):
        self.log.configure(state="normal");self.log.insert("end",str(text)+"\n");self.log.see("end");self.log.configure(state="disabled")
    def open_output(self):
        key="gd_output" if self.project_template.get()==PROJECT_TEMPLATES[1] else "cq_output"; path=Path(self.vars[key].get() or ".")
        if path.is_dir():os.startfile(path)
        else:messagebox.showwarning("提示","输出文件夹不存在")


def launch_app():
    ReportGeneratorApp().mainloop()


if __name__ == "__main__":
    launch_app()
