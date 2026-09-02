#!/usr/bin/env python3
"""Standalone post-build verifier for report-generator release artifacts.

Runs at the end of build-release.bat to avoid codepage mojibake when the
batch file hands Chinese paths to a one-liner Python command via cmd.exe.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path


def main() -> int:
    root = Path(os.environ.get("RELEASE_ROOT") or os.getcwd())
    expected = [
        root / "dist" / "报告生成工具" / "报告生成工具.exe",
        root / "dist" / "updater" / "updater.exe",
        root / "dist" / "报告生成工具-Setup.exe",
    ]
    missing = [
        str(path)
        for path in expected
        if not path.is_file() or path.read_bytes()[:2] != b"MZ"
    ]
    if missing:
        print("[ERROR] Invalid build artifacts: " + ", ".join(missing), flush=True)
        return 1
    print("Release build verified:", flush=True)
    for path in expected:
        print("  " + str(path) + f" ({path.stat().st_size} bytes)", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
