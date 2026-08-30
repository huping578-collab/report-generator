from __future__ import annotations

import sys
from pathlib import Path

import webview

from bridge import DesktopBridge


def resource_path(relative: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative


def main() -> None:
    frontend = resource_path("frontend/index.html")
    if not frontend.is_file():
        raise FileNotFoundError(f"前端文件不存在：{frontend}")

    api = DesktopBridge()
    window = webview.create_window(
        "报告生成工具",
        url=str(frontend),
        js_api=api,
        width=1380,
        height=880,
        min_size=(920, 680),
        background_color="#F3F6F8",
        text_select=True,
        confirm_close=True,
    )
    api.attach_window(window)
    webview.start(
        gui="edgechromium",
        debug=False,
        private_mode=False,
    )


if __name__ == "__main__":
    from backend import report_engine as engine

    log_dir = engine.application_root() / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_stream = open(log_dir / "startup.log", "w", encoding="utf-8")
    sys.stderr = log_stream
    try:
        main()
    except Exception:
        import traceback as _tb

        log_stream.write(_tb.format_exc())
        log_stream.flush()
        raise
