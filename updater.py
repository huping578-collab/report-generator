from __future__ import annotations

import json
import re
import subprocess
import sys
import tempfile
import threading
from pathlib import Path
from typing import Any, Callable
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen

APP_NAME = "报告生成工具"
ASSET_NAME = "report-generator-Setup.exe"
RELEASE_API = "https://api.github.com/repos/huping578-collab/report-generator/releases/latest"
DOWNLOAD_HOSTS = frozenset({"github.com", "objects.githubusercontent.com"})


def resource_path(relative: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative


def parse_version(value: str) -> tuple[int, int, int]:
    match = re.fullmatch(r"v?(\d+)\.(\d+)\.(\d+)", value.strip())
    if not match:
        raise ValueError(f"版本号格式无效：{value!r}，应为 X.Y.Z。")
    return tuple(int(part) for part in match.groups())


def is_newer(candidate: tuple[int, int, int], current: tuple[int, int, int]) -> bool:
    return candidate > current


def _validate_asset_url(asset: dict[str, Any]) -> None:
    url = asset.get("browser_download_url")
    parsed = urlparse(url if isinstance(url, str) else "")
    if parsed.scheme != "https" or parsed.hostname not in DOWNLOAD_HOSTS:
        raise ValueError("更新资产下载地址不是受信任的 HTTPS GitHub 地址。")


def select_installer_asset(release: dict[str, Any]) -> dict[str, Any]:
    assets = release.get("assets") if isinstance(release, dict) else None
    if not isinstance(assets, list):
        raise ValueError("GitHub Release 响应缺少有效资产列表。")
    for asset in assets:
        if isinstance(asset, dict) and asset.get("name") == ASSET_NAME:
            _validate_asset_url(asset)
            return asset
    raise ValueError(f"最新 GitHub Release 缺少资产：{ASSET_NAME}")


def validate_download(data: bytes, expected_size: int | None = None) -> None:
    if not data.startswith(b"MZ"):
        raise ValueError("下载文件不是有效的 Windows 安装程序。")
    if expected_size is not None and len(data) != expected_size:
        raise ValueError(f"下载文件大小不符：实际 {len(data)} 字节，预期 {expected_size} 字节。")


def fetch_latest_release() -> dict[str, Any]:
    request = Request(
        RELEASE_API,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "report-generator-updater",
        },
    )
    try:
        with urlopen(request, timeout=20) as response:
            if getattr(response, "status", 200) >= 400:
                raise RuntimeError(f"GitHub 返回 HTTP {response.status}。")
            return json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        raise RuntimeError(f"GitHub 请求失败：HTTP {exc.code}。") from exc
    except (URLError, TimeoutError, OSError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"无法读取 GitHub Release：{exc}") from exc


def check_for_update(
    current_version: str | None = None,
    fetcher: Callable[[], dict[str, Any]] = fetch_latest_release,
) -> dict[str, Any]:
    current_text = current_version or resource_path("version.txt").read_text(encoding="utf-8").strip()
    current = parse_version(current_text)
    release = fetcher()
    tag = release.get("tag_name") if isinstance(release, dict) else None
    latest = parse_version(tag if isinstance(tag, str) else "")
    asset = select_installer_asset(release)
    return {
        "current_version": current_text.lstrip("v"),
        "latest_version": ".".join(str(part) for part in latest),
        "release": release,
        "asset": asset,
        "update_available": is_newer(latest, current),
    }


def download_installer(asset: dict[str, Any]) -> Path:
    _validate_asset_url(asset)
    expected_size = asset.get("size")
    if not isinstance(expected_size, int) or expected_size <= 0:
        expected_size = None

    temp_file = tempfile.NamedTemporaryFile(
        prefix="report-generator-update-",
        suffix=".exe",
        delete=False,
    )
    temp_path = Path(temp_file.name)
    temp_file.close()
    try:
        request = Request(
            asset["browser_download_url"],
            headers={"User-Agent": "report-generator-updater"},
        )
        total = 0
        with urlopen(request, timeout=120) as response:
            with temp_path.open("wb") as output:
                while chunk := response.read(1024 * 1024):
                    output.write(chunk)
                    total += len(chunk)
        if expected_size is not None and total != expected_size:
            raise ValueError(f"下载文件大小不符：实际 {total} 字节，预期 {expected_size} 字节。")
        with temp_path.open("rb") as downloaded:
            if downloaded.read(2) != b"MZ":
                raise ValueError("下载文件不是有效的 Windows 安装程序。")
        return temp_path
    except (HTTPError, URLError, TimeoutError, OSError) as exc:
        temp_path.unlink(missing_ok=True)
        raise RuntimeError(f"下载更新程序失败：{exc}") from exc
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def launch_installer(path: Path) -> None:
    subprocess.Popen([str(path)], close_fds=True)


def main() -> None:
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.title(f"{APP_NAME} - 检查更新")
    root.resizable(False, False)
    root.geometry("460x220")

    status = tk.StringVar(value="正在检查 GitHub 最新版本…")
    detail = tk.StringVar(value="")
    check_button = tk.Button(root, text="重新检查", state=tk.DISABLED)
    update_button = tk.Button(root, text="下载并安装更新", state=tk.DISABLED)
    result: dict[str, Any] = {}

    tk.Label(root, text=APP_NAME, font=("Microsoft YaHei", 15, "bold")).pack(pady=(22, 8))
    tk.Label(root, textvariable=status, wraplength=410).pack(pady=4)
    tk.Label(root, textvariable=detail, wraplength=410, justify="left").pack(pady=4)
    buttons = tk.Frame(root)
    buttons.pack(pady=12)
    check_button.pack(in_=buttons, side=tk.LEFT, padx=5)
    update_button.pack(in_=buttons, side=tk.LEFT, padx=5)

    def show_error(exc: Exception) -> None:
        status.set("检查更新失败")
        detail.set(str(exc))
        check_button.config(state=tk.NORMAL)
        messagebox.showerror("检查更新失败", str(exc), parent=root)

    def apply_result() -> None:
        if result["update_available"]:
            status.set(f"发现新版本：{result['latest_version']}")
            body = result["release"].get("body") or ""
            detail.set(f"当前版本：{result['current_version']}\n{body[:500]}")
            update_button.config(state=tk.NORMAL)
        else:
            status.set(f"当前已是最新版本：{result['current_version']}")
            detail.set("GitHub Releases 中没有更新版本。")
        check_button.config(state=tk.NORMAL)

    def check_worker() -> None:
        try:
            result.update(check_for_update())
            root.after(0, apply_result)
        except Exception as exc:
            root.after(0, lambda error=exc: show_error(error))

    def check() -> None:
        check_button.config(state=tk.DISABLED)
        update_button.config(state=tk.DISABLED)
        status.set("正在检查 GitHub 最新版本…")
        threading.Thread(target=check_worker, daemon=True).start()

    def install_worker() -> None:
        try:
            installer = download_installer(result["asset"])
            launch_installer(installer)
            root.after(0, root.destroy)
        except Exception as exc:
            root.after(0, lambda error=exc: show_error(error))

    def install() -> None:
        update_button.config(state=tk.DISABLED)
        status.set("正在下载更新安装包…")
        threading.Thread(target=install_worker, daemon=True).start()

    check_button.config(command=check)
    update_button.config(command=install)
    root.after(100, check)
    root.mainloop()


if __name__ == "__main__":
    main()
