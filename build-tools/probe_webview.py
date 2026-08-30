import sys
from pathlib import Path


def probe():
    root = Path(__file__).resolve().parent
    frontend = root / "frontend" / "index.html"
    print("frontend exists:", frontend.is_file(), file=sys.stderr)

    import webview

    print("webview module:", webview.__file__, file=sys.stderr)
    print("edge platform avail:", hasattr(webview, "platforms") and hasattr(webview.platforms, "edgechromium"), file=sys.stderr)

    import webview.platforms.edgechromium as edge

    print("edgechromium import ok", file=sys.stderr)

    import clr_loader

    print("clr_loader ok", file=sys.stderr)

    import pythonnet

    print("pythonnet ok", file=sys.stderr)

    window = webview.create_window("probe", url=str(frontend), width=800, height=600)

    def loaded():
        print("LOADED_OK", file=sys.stderr)

    window.events.loaded += loaded
    webview.start(gui="edgechromium", debug=False)
    print("START_RETURNED", file=sys.stderr)


if __name__ == "__main__":
    try:
        probe()
    except Exception as exc:
        import traceback

        traceback.print_exc()
        sys.exit(1)
