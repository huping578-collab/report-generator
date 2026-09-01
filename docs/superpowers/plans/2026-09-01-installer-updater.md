# 安装器与更新器实现计划

> **面向 AI 代理的工作者：** 按本计划逐任务实现；每个代码任务先写失败测试，再实现最小代码并回归。

**目标：** 为报告生成工具构建 Windows 安装器、独立 GitHub Releases 更新器和自动发布工作流。

**架构：** 主程序仍由现有 PyInstaller spec 构建。更新器是一个使用标准库的独立 tkinter EXE，读取内置版本、查询固定 GitHub Releases API、校验并下载精确命名的安装器，再交给 Inno Setup 覆盖安装。Inno Setup 负责安装、卸载、快捷方式和升级，不让更新器直接替换正在运行的程序文件。

**技术栈：** Python 3.11 标准库、PyInstaller、Inno Setup 6、GitHub Actions。

---

## 文件职责

- 创建：`version.txt` — 唯一三段式版本号。
- 创建：`updater.py` — Release 查询、版本比较、资产校验、下载和 tkinter 界面。
- 创建：`installer.iss` — Inno Setup 安装/卸载定义。
- 创建：`build-release.bat` — 构建主程序、更新器和安装器。
- 创建：`.github/workflows/release.yml` — tag 触发的 Windows Release 构建与上传。
- 修改：`build.bat` — 把版本文件和模板资料复制到主程序发布目录。
- 修改：`tests/test_report_generator.py` — 继续使用仓库唯一测试程序，加入更新器核心逻辑测试。
- 修改：`README.md` — 写明安装器构建、Release 资产命名和更新方式。

## 任务 1：版本读取和更新器核心逻辑

**目标：** 在不启动 GUI 的情况下完成可靠的版本比较、Release 资产筛选和安装包校验。

**文件：**
- 创建：`version.txt`
- 创建：`updater.py`
- 修改：`tests/test_report_generator.py`

- [ ] **步骤 1：写失败测试**

在现有 `tests/test_report_generator.py` 中导入 `parse_version`、`is_newer`、`select_installer_asset` 和 `validate_installer_bytes`，覆盖：

```python
self.assertEqual(parse_version("v1.2.3"), (1, 2, 3))
self.assertTrue(is_newer((1, 2, 4), (1, 2, 3)))
self.assertFalse(is_newer((1, 2, 3), (1, 2, 3)))
asset = select_installer_asset({"assets": [{"name": "报告生成工具-Setup.exe", "browser_download_url": "https://github.com/a/b/releases/download/v1/报告生成工具-Setup.exe", "size": 2}]})
self.assertEqual(asset["name"], "报告生成工具-Setup.exe")
with self.assertRaises(ValueError):
    validate_installer_bytes(b"not-an-exe", 11)
```

- [ ] **步骤 2：运行测试确认失败**

运行：

```bash
.venv/Scripts/python.exe -m unittest tests.test_report_generator.DesktopBridgeTests.test_updater_core_validation -v
```

预期：FAIL，因为 `updater.py` 和核心函数尚未存在。

- [ ] **步骤 3：实现最少核心逻辑**

`updater.py` 采用以下接口和规则：

```python
APP_VERSION = "0.1.0"  # 实际值从打包进 EXE 的 version.txt 读取
ASSET_NAME = "报告生成工具-Setup.exe"
RELEASE_API = "https://api.github.com/repos/huping578-collab/report-generator/releases/latest"

def parse_version(value: str) -> tuple[int, int, int]: ...
def is_newer(candidate: tuple[int, int, int], current: tuple[int, int, int]) -> bool: ...
def select_installer_asset(release: dict) -> dict: ...
def validate_download(data: bytes, expected_size: int | None) -> None: ...
```

核心实现必须：只接受三段数字版本；只筛选精确资产名；只接受 HTTPS 且 host 为 `github.com` 或 `objects.githubusercontent.com`；下载内容必须以 `MZ` 开头，且在 Release 提供 size 时字节数必须一致。所有失败都抛出 `ValueError` 或 `RuntimeError`，不启动安装器。

- [ ] **步骤 4：运行测试确认通过**

运行同一个测试目标，预期 PASS；再运行完整测试程序，预期所有原有测试和新增测试通过。

- [ ] **步骤 5：提交**

```bash
git add version.txt updater.py tests/test_report_generator.py
git commit -m "feat(更新器): 添加 GitHub Release 更新核心逻辑"
```

## 任务 2：更新器 GUI、下载和安装启动

**目标：** 将任务 1 的可测试核心接入最小 Windows GUI。

**文件：**
- 修改：`updater.py`
- 修改：`tests/test_report_generator.py`

- [ ] **步骤 1：写失败测试**

加入下载 URL 校验测试：错误协议、错误主机、错误资产名均拒绝；使用 `unittest.mock.patch` 验证 `subprocess.Popen` 只在校验成功后调用。

- [ ] **步骤 2：运行测试确认失败**

运行新增的更新器测试目标，预期 FAIL。

- [ ] **步骤 3：实现最少 GUI**

增加：

- `fetch_latest_release()`：发送带 `Accept: application/vnd.github+json` 和 `User-Agent` 的请求，解析 JSON。
- `download_installer(asset)`：流式写入 `%TEMP%` 临时 EXE，校验 size 和 `MZ`，失败时删除临时文件。
- `launch_installer(path)`：使用 `subprocess.Popen([str(path)])` 启动，不使用 shell。
- `main()`：tkinter 窗口显示当前版本、检查结果、Release 版本和错误；有新版本时提供“下载并安装”按钮。
- 路径、仓库和资产名使用常量；不从用户输入拼接命令。

- [ ] **步骤 4：运行测试确认通过**

运行更新器测试和完整测试；预期无失败。用 `python updater.py` 做一次 GUI 启动冒烟，不点击下载。

- [ ] **步骤 5：提交**

```bash
git add updater.py tests/test_report_generator.py
git commit -m "feat(更新器): 添加 Release 检查与安装启动界面"
```

## 任务 3：Inno Setup 安装器

**目标：** 将现有主程序目录和更新器安装为用户可卸载的 Windows 应用。

**文件：**
- 创建：`installer.iss`

- [ ] **步骤 1：实现安装定义**

要求：

- `AppId` 固定，`AppVersion` 由 `build-release.bat` 通过 `/DAppVersion` 传入。
- 默认目录 `{localappdata}\\Programs\\报告生成工具`，`PrivilegesRequired=lowest`。
- 主程序来自 `dist\\报告生成工具\\*`；更新器来自 `dist\\更新程序\\更新程序.exe`。
- 模板目录单独复制并标记 `uninsneveruninstall`，避免卸载删除用户模板。
- 创建主程序和更新器的开始菜单/桌面快捷方式。
- 使用 `报告生成工具-Setup.exe` 作为输出名。
- 安装后可选启动主程序；不设置强制静默升级和自动回滚。

- [ ] **步骤 2：用最小构建检查编译**

由任务 4 的构建脚本调用 Inno Setup；预期生成 `dist\\报告生成工具-Setup.exe`，文件首字节为 `MZ`。

- [ ] **步骤 3：提交**

```bash
git add installer.iss
git commit -m "feat(安装器): 添加 Inno Setup 安装与卸载配置"
```

## 任务 4：本地发布构建脚本

**目标：** 一条命令构建主程序、更新器和安装器。

**文件：**
- 创建：`build-release.bat`
- 修改：`build.bat`

- [ ] **步骤 1：实现脚本**

`build-release.bat` 必须：校验 `version.txt`、校验 `.venv\\Scripts\\python.exe` 和 Inno Setup `ISCC.exe`；调用现有 `build.bat`；用 PyInstaller `--onefile --windowed` 构建 `更新程序.exe` 并把 `version.txt` 作为数据文件打包；调用 `ISCC.exe /DAppVersion=<version> installer.iss`；最后检查两个 EXE 存在且首字节为 `MZ`。

`build.bat` 增加复制 `version.txt` 到 `dist\\报告生成工具`，保持既有主程序构建行为不变。

- [ ] **步骤 2：本地构建验证**

运行：

```bat
build-release.bat
```

预期存在：

```text
dist\\报告生成工具\\报告生成工具.exe
dist\\更新程序\\更新程序.exe
dist\\报告生成工具-Setup.exe
```

- [ ] **步骤 3：提交**

```bash
git add build-release.bat build.bat
 git commit -m "chore(构建): 添加安装器与更新器发布脚本"
```

## 任务 5：GitHub Release 自动发布

**目标：** 推送 `vX.Y.Z` tag 时自动构建并发布可供更新器下载的安装包。

**文件：**
- 创建：`.github/workflows/release.yml`

- [ ] **步骤 1：实现工作流**

工作流在 `windows-latest` 上执行：checkout、安装 Python 3.11、安装 `requirements.txt`、安装 Inno Setup、检查 tag 去掉 `v` 后等于 `version.txt`、运行 `build-release.bat`，并将 `dist/报告生成工具-Setup.exe` 上传到同名 GitHub Release。

- [ ] **步骤 2：静态检查**

检查 YAML 可解析、资产路径和更新器固定资产名完全一致；不把 token、用户数据或模板上传到 Release。

- [ ] **步骤 3：提交**

```bash
git add .github/workflows/release.yml
 git commit -m "ci(发布): 添加 Windows Release 自动构建"
```

## 任务 6：文档、全量验证和交付

**目标：** 让用户知道如何构建、发布和使用更新器，并验证最终交付物。

**文件：**
- 修改：`README.md`

- [ ] **步骤 1：补充文档**

说明：安装器构建命令、安装默认目录、开始菜单更新入口、Release tag 规则 `vX.Y.Z`、安装器资产名 `报告生成工具-Setup.exe`、更新器只消费 GitHub Releases，以及 Windows 10 需要 WebView2 Runtime。

- [ ] **步骤 2：全量验证**

运行：

```bash
.venv/Scripts/python.exe -m py_compile app.py bridge.py backend/report_engine.py backend/minimal_docx.py updater.py tests/test_report_generator.py
.venv/Scripts/python.exe -m unittest tests.test_report_generator -v
build-release.bat
```

核对：测试无失败；三个 EXE 存在且为 PE 文件；安装器可启动；安装后主程序和更新器快捷方式存在；更新器在当前版本无新 Release 或 API 不可用时给出可读提示，不崩溃。

- [ ] **步骤 3：提交并检查工作树**

```bash
git add README.md
git commit -m "docs(发布): 补充安装器与更新器使用说明"
git status --short --branch
git diff --check
```
