from __future__ import annotations

import json
import os
import threading
import traceback
from pathlib import Path
from typing import Any

import webview

from backend import report_engine as engine


class DesktopBridge:
    """Native operations exposed to the HTML frontend through pywebview."""

    def __init__(self) -> None:
        self._window: webview.Window | None = None
        self._running = False
        self._lock = threading.Lock()
        self._progress = 5
        self._stage = 0

    def attach_window(self, window: webview.Window) -> None:
        self._window = window

    def get_environment(self) -> dict[str, Any]:
        templates = self._template_paths()
        return {
            "desktop": True,
            "program": engine.PROGRAM_NAME,
            "templates": {
                name: {"path": str(path), "exists": path.is_file()}
                for name, path in templates.items()
            },
        }

    def choose_path(self, kind: str, key: str, template: str) -> dict[str, Any]:
        if self._window is None:
            return {"ok": False, "error": "桌面窗口尚未初始化。"}

        try:
            if kind == "file":
                selected = self._window.create_file_dialog(
                    webview.FileDialog.OPEN,
                    allow_multiple=False,
                    file_types=("Excel 工作簿 (*.xlsx)",),
                )
            else:
                selected = self._window.create_file_dialog(
                    webview.FileDialog.FOLDER,
                    allow_multiple=False,
                )
            if not selected:
                return {"ok": True, "cancelled": True}

            path = str(Path(selected[0]).resolve())
            detected: dict[str, str] = {}
            if key == "projectPath":
                detected = self._detect_project_paths(Path(path), template)
            elif key == "summaryPath" and template != "gd":
                detected = {"countyOptions": self._summary_counties(Path(path))}
            return {"ok": True, "cancelled": False, "path": path, "detected": detected}
        except Exception as exc:
            return {"ok": False, "error": f"无法打开路径选择器：{exc}"}

    def start_run(self, payload: dict[str, Any]) -> dict[str, Any]:
        with self._lock:
            if self._running:
                return {"ok": False, "error": "已有任务正在运行，请等待当前任务完成。"}
            self._running = True

        try:
            self._validate_payload(payload)
        except Exception as exc:
            with self._lock:
                self._running = False
            return {"ok": False, "error": str(exc)}

        worker = threading.Thread(target=self._run_worker, args=(payload,), daemon=True)
        worker.start()
        return {"ok": True, "started": True}

    def open_output(self, path: str) -> dict[str, Any]:
        try:
            folder = Path(path).expanduser().resolve()
            if not folder.is_dir():
                return {"ok": False, "error": "输出文件夹不存在，请先完成生成或重新选择输出位置。"}
            os.startfile(folder)
            return {"ok": True}
        except Exception as exc:
            return {"ok": False, "error": f"无法打开输出文件夹：{exc}"}

    def _run_worker(self, payload: dict[str, Any]) -> None:
        try:
            self._progress, self._stage = 5, 0
            self._emit("run", status="running", progress=5, stage=0, message="开始校验项目资料")
            template = payload["template"]
            values = payload["values"]
            if template == "gd":
                self._run_guangdong(values)
            else:
                self._run_chongqing(values)
            self._emit("run", status="complete", progress=100, stage=3, message="全部输出文件已保存")
        except PermissionError:
            message = "文件被占用，无法保存。请关闭已打开的 Excel 或 Word 文件后重试。"
            self._emit("run", status="error", progress=0, stage=0, message=message)
        except Exception as exc:
            self._emit(
                "run",
                status="error",
                progress=0,
                stage=0,
                message=str(exc),
                detail=traceback.format_exc(),
            )
        finally:
            with self._lock:
                self._running = False

    def _run_chongqing(self, values: dict[str, Any]) -> None:
        template_path = self._template_paths()["重庆项目报告模板"]
        config = engine.Config(
            project_dir=Path(values["projectPath"]),
            summary_xlsx=Path(values["summaryPath"]),
            detail_dir=Path(values["detailPath"]),
            template_docx=template_path,
            output_dir=Path(values["outputPath"]),
            disease_dir=Path(values["diseasePath"]) if values.get("diseasePath") else None,
            tci_path=Path(values["tciPath"]) if values.get("tciPath") else None,
        )
        engine.generate_statistics_and_report(
            config,
            log=self._on_engine_log,
            generate_charts_first=True,
            process_height=True,
            process_bolts=True,
            process_tci=bool(values.get("tciPath")),
            require_template=False,
            county_override=(str(values.get("countySelect") or "").strip() or None),
        )

    def _run_guangdong(self, values: dict[str, Any]) -> None:
        config = engine.GuangdongConfig(
            project_dir=Path(values["projectPath"]),
            manual_xlsx=Path(values["manualPath"]),
            route_xlsx=Path(values["routePath"]),
            output_dir=Path(values["outputPath"]),
            marking_threshold=values["markingThreshold"],
            height_threshold=values["heightThreshold"],
            bolt_threshold=values["boltThreshold"],
            marking_dir=Path(values["markingPath"]) if values.get("markingPath") else None,
            guardrail_dir=Path(values["guardrailPath"]) if values.get("guardrailPath") else None,
        )
        engine.run_guangdong_project(config, log=self._on_engine_log)

    def _validate_payload(self, payload: dict[str, Any]) -> None:
        if payload.get("template") not in {"cq", "gd"}:
            raise ValueError("项目模板无效。")
        values = payload.get("values")
        if not isinstance(values, dict):
            raise ValueError("项目配置无效。")

        required = {
            "cq": ("projectPath", "summaryPath", "detailPath", "diseasePath", "tciPath", "outputPath"),
            "gd": ("projectPath", "manualPath", "routePath", "outputPath"),
        }[payload["template"]]
        labels = {
            "projectPath": "项目资料文件夹",
            "summaryPath": "分段汇总表",
            "detailPath": "检测明细文件夹",
            "diseasePath": "病害清单文件夹",
            "tciPath": "TCI数据文件夹",
            "markingPath": "标线数据文件夹",
            "guardrailPath": "护栏数据文件夹",
            "manualPath": "人工自动化对比表",
            "routePath": "路线分类表",
            "outputPath": "输出文件夹",
        }
        for key in required:
            if not str(values.get(key, "")).strip():
                raise ValueError(f"请选择{labels[key]}。")

        file_keys = {"summaryPath", "manualPath", "routePath"}
        for key in required:
            path = Path(values[key])
            if key in file_keys and not path.is_file():
                raise FileNotFoundError(f"{labels[key]}不存在：{path}")
            if key not in file_keys and key != "outputPath" and not path.is_dir():
                raise FileNotFoundError(f"{labels[key]}不存在：{path}")

        output = Path(values["outputPath"])
        output.mkdir(parents=True, exist_ok=True)

        if payload["template"] == "gd":
            template_path = self._template_paths()["广东项目第五章模板"]
            if not template_path.is_file():
                raise FileNotFoundError(
                    f"缺少内置 Word 模板：{template_path.name}。请将模板放入程序 templates 文件夹。"
                )
            engine.validate_thresholds(
                values.get("markingThreshold"),
                values.get("heightThreshold"),
                values.get("boltThreshold"),
            )

    @staticmethod
    def _summary_counties(summary: Path) -> list[str]:
        try:
            counties = []
            for segment in engine.read_segments(Path(summary)):
                county = str(segment.get("county") or "").strip()
                if county and county not in counties:
                    counties.append(county)
            return counties
        except Exception:
            return []

    def _detect_project_paths(self, folder: Path, template: str) -> dict[str, str]:
        if template == "gd":
            marking, guardrail = engine.detect_guangdong_data_folders(folder)
            result = {"outputPath": str(folder)}
            if marking:
                result["markingPath"] = str(marking)
            if guardrail:
                result["guardrailPath"] = str(guardrail)
            return result

        summary, detail, disease, tci = engine.discover_paths(folder)
        result = {"outputPath": str(summary.parent if summary else folder)}
        if summary:
            result["summaryPath"] = str(summary)
            result["countyOptions"] = self._summary_counties(summary)
        if detail:
            result["detailPath"] = str(detail)
        if disease:
            result["diseasePath"] = str(disease)
        if tci:
            result["tciPath"] = str(tci)
        return result

    def _template_paths(self) -> dict[str, Path]:
        return {
            "重庆项目报告模板": engine.resource_template_path("重庆项目报告模板.md"),
            "广东项目第五章模板": engine.resource_template_path("广东项目第五章模板.md"),
        }

    def _on_engine_log(self, message: Any) -> None:
        text = str(message)
        progress, stage = self._progress_from_log(text)
        # 总进度与阶段只增不减：长耗时步骤之间的杂项日志不会让进度条回退。
        self._progress = max(self._progress, progress)
        self._stage = max(self._stage, stage)
        self._emit("log", status="running", progress=self._progress, stage=self._stage, message=text)

    # 各阶段在总进度中的区间，按真实运行实测耗时占比划分（两江新区+万盛+重庆市资料，总 177.5s：
    # 解析明细≈7s、解析螺栓≈19s、索引病害图片≈4s、逐区县统计与报告≈145s）。
    _PHASE_RANGES = {
        "解析TCI": (1, 5, 6),
        "解析明细": (1, 6, 15),
        "解析螺栓明细": (1, 15, 26),
        "索引病害图片": (1, 26, 30),
        "扫描资料": (1, 5, 30),  # 广东扫描阶段
    }
    # 逐区县循环：同一个 1/n 计数器，区县内「统计工作簿」占前一半、「报告」占后一半。
    _COUNTY_RANGE = (30, 99)
    _COUNTY_PHASES = {"生成区县统计": (2, 0.0), "生成区县报告": (3, 0.5)}
    _STAGE_START = (5, 5, 30, 62)
    # 只认强信号词：完成/生成/报告 这类词在过程中频繁出现，会把进度提前顶到阶段起点。
    _STAGE_KEYWORDS = (
        (3, ("保存", "导出", "写入")),
        (2, ("工作簿", "统计表", "汇总", "图表")),
        (1, ("扫描", "识别", "索引", "解析", "读取", "校验", "导入")),
    )

    def _progress_from_log(self, message: str) -> tuple[int, int]:
        structured = engine.parse_progress(message)
        if structured:
            name = structured["stage"]
            current, total = structured["current"], structured["total"]
            if name in self._COUNTY_PHASES:
                stage, offset = self._COUNTY_PHASES[name]
                start, end = self._COUNTY_RANGE
                position = (current - 1 + offset) / total
                return round(start + (end - start) * position), stage
            span = self._PHASE_RANGES.get(name)
            if span:
                stage, start, end = span
                return round(start + (end - start) * (current / total)), stage
        for stage, keywords in self._STAGE_KEYWORDS:
            if any(word in message for word in keywords):
                return self._STAGE_START[stage], stage
        return self._STAGE_START[0], 0

    def _emit(self, event: str, **payload: Any) -> None:
        if self._window is None:
            return
        data = json.dumps({"event": event, **payload}, ensure_ascii=False)
        try:
            self._window.evaluate_js(f"window.desktopEvents && window.desktopEvents({data})")
        except Exception:
            pass
