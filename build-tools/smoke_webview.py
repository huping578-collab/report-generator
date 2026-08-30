from __future__ import annotations

import sys
import time
from pathlib import Path

import webview

frontend = Path(__file__).resolve().parent / "frontend" / "index.html"


def on_done(): 
    print("GUI_OK", flush=True)


def main():
    window = webview.create_window(
        "Smoke Test",
        url=str(frontend),
        width=900,
        height=700,
        background_color="#F3F6F8",
    )
    window.events.loaded += on_done
    webview.start(gui="edgechromium", debug=False, private_mode=False)
    print("WEBVIEW_EXITED", flush=True)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"SMOKE_FAIL {exc}", flush=True)
        sys.exit(1)
