#!/usr/bin/env python3
"""Stage build assets into dist\报告生成工具.

Runs from build.bat. Uses pathlib (UTF-8) so Chinese paths are codepage-safe
when cmd.exe hands them over; cmd's own mkdir/copy of Chinese paths would
create mojibake directories under a non-UTF-8 codepage.
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
APP_DIR = ROOT / "dist" / "报告生成工具"


def main() -> int:
    app_templates = APP_DIR / "templates"
    app_templates.mkdir(parents=True, exist_ok=True)
    src_templates = ROOT / "templates"
    # Markdown 模板为当前模板形态；docx 模板不再随构建分发。
    for src in list(src_templates.glob("*.md")) + [src_templates / "模板放置说明.txt"]:
        if src.is_file():
            shutil.copy2(src, app_templates / src.name)
    for name in ("使用说明.txt", "version.txt"):
        src = ROOT / name
        if src.is_file():
            shutil.copy2(src, APP_DIR / name)
    print("Staged assets into", APP_DIR, flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
