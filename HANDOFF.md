# HANDOFF — 报告生成工具桌面版

> 交接日期：2026-08-30
> 上一位开发：Hermes Agent（frontender profile）
> 接续端：另一台 Windows 设备，通过 GitHub 仓库同步

## 1. 项目速览

原 `报告生成工具V0.1-5.py` 是一个 Tkinter 单窗口工具，生成重庆 G210 和广东多地市的交安设施检测统计 Excel 与 Word 报告。桌面版把它改造成 **HTML/CSS/JS 前端（已确认的"工程资料路由台"三栏界面）+ Python 后端（原报表引擎原样保留）** 的 Windows 桌面应用，用 pywebview 做 WebView 壳，PyInstaller 打包 EXE。

- 界面：左侧模板导航 + 中间资料配置 + 右侧运行监控（进度、阶段、日志、开始生成/打开输出）
- 数据只在本机处理；无网络依赖（图表用 matplotlib，文档用 python-docx）
- 重庆流程**不需要 Word 模板**（程序化生成）；广东流程**需要**模板

## 2. 仓库

- GitHub（私有）：https://github.com/huping578-collab/report-generator
- 默认分支 `main`，远端已同步（当前 commit `7c04598`）

## 3. 目录结构

```text
app.py                     # 桌面壳入口：pywebview 窗口 + 启动日志
bridge.py                  # 前后端桥接：get_environment / choose_path / start_run / open_output；校验、后台线程、进度映射
backend/
  report_engine.py         # 原 Python 报表引擎（双模板完整逻辑，未删改业务）
  minimal_docx.py          # 重庆程序化报告生成（无模板模式）
frontend/
  index.html               # 已确认的界面（CSS 内联）
  app.js                   # 交互 + pywebview 桥接调用 + 状态机
templates/                 # 模板放置（.docx 用户提供，gitignore 排除；说明文件入库）
tests/
  test_report_generator.py # 唯一长期测试程序
  artifacts/               # 测试产物目录；Git 仅保留 .gitkeep
build-tools/
  create_icon.py           # 生成 assets/app.ico（PIL）
report_generator.spec      # PyInstaller spec（collect_all webview + pythonnet）
build.bat                  # 一键构建（调用 spec + 复制模板说明/使用说明）
README.md / 使用说明.txt   # 面向开发 / 最终用户
DESIGN.md / PRODUCT.md     # 设计系统与产品文档
requirements.txt           # pywebview==6.2.1 pyinstaller==6.22.2 openpyxl matplotlib python-docx
```

## 4. 从头接续（新设备）

```bat
git clone https://github.com/huping578-collab/report-generator.git
cd report-generator
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

注意：仓库不含 `.venv` / `build` / `dist` / `logs` / `artifacts`（gitignore），首次需重建。Windows 10 需安装 Microsoft Edge WebView2 Runtime（Win11 自带）。

## 5. 验证矩阵（全绿基准）

| 项目 | 命令 | 结果 |
|---|---|---|
| 报告生成回归 | `.venv\Scripts\python.exe -m unittest tests.test_report_generator -v` | 13 passed |
| 源码模式 GUI | `.venv\Scripts\python.exe app.py` | 窗口出现，pywebview loaded |
| 打包 EXE | `build.bat` | 生成 `dist\报告生成工具\报告生成工具.exe` |

## 6. 关键决策与坑位（必须知道的）

1. **模板路径**：`report_engine.builtin_template_paths()` 原按 `__file__` 定位，打包后会指向 PyInstaller 内部目录。已改为 `application_root()`：源码模式=项目根，冻结模式=EXE 同级目录。文件：`backend/report_engine.py:104-118`。
2. **程序化生成模式**：重庆 `require_template=False`，模板缺失时 `make_docx` 转到 `minimal_docx.py`（python-docx 从头顶起：标题、章节编号、表格、matplotlib 图片、题注、横向/竖向页面）。广东强制 `require_template=True`（未提供则前端禁用开始生成）。
3. **`server_args={"log": True}` 不可用**：pywebview 6.2.1 的 BottleServer 不接受该参数，`webview.start()` 直接抛 TypeError 秒退。曾因旧 EXE 用此参数构建导致"启动后无窗口"，务必用当前 `app.py`（无此参数）重建。
4. **构建前必须关掉运行中的 EXE**：PyInstaller COLLECT 要清理旧 `dist/报告生成工具`，运行中的 EXE 会锁 `base_library.zip` 导致 PermissionError；构建前先关闭，避免锁定 `dist` 产物。
5. **EXE 启动故障排查**：`logs/startup.log` 捕获 stderr/异常，位于 EXE 同级目录（源码模式在项目根）。空文件但窗口没出现 → 查 WebView2 运行时。
6. **git-bash + PowerShell 脚本**：内联 `$_` 会被 MSYS 转义破坏（`Get-Process | Where-Object { $_... }` 会报 CommandNotFoundException）。一律写成 `.ps1` 文件再 `powershell -File` 执行。
7. **中文字符路径 OK**：`C:\文件\工作工具台\报告生成工具` 全程可用；pywebview/PyInstaller 均正常。
8. **venv 的 python 才能跑**：系统 `python` 没有 webview/PyInstaller，必须用 `.venv\Scripts\python.exe`。

## 7. 已知问题 / 后续优先项

- [ ] **广东模板入仓**：`templates/广东项目第五章模板.docx` 由用户提供后放入（gitignore 已排除 .docx，如要入库需改 .gitignore）；放入后重启应用即可用。
- [ ] **真实数据端到端验证**：目前 EXE 用真实界面 + 桥接验证；统计/报告流水线是以合成数据跑通 `minimal_docx` 与引擎单元测试。用户手头有真实 G210 数据后建议跑一次完整"开始生成"，核对输出 Excel/Word 与旧版 Tkinter 程序一致。
- [ ] **P3 遗留**（原型阶段 reviewer 记录，均已在本版解决或迁移）：
  - Toast 已从右下角移到右上角（避开"打开输出"），见 `index.html .toast`。
  - 移动端已加粘性 `run-actions`（`sticky bottom`），但桌面 EXE 固定 1380×880，移动布局仅在浏览器预览生效。
- [ ] **仅浏览器打开时**：本地文件功能禁用（runButton/chooseButton disabled），显示"浏览器预览模式"，这是有意的降级提示，不是 bug。
- [ ] **未做**：EXE 数字签名（SmartScreen 会提示"更多信息→仍要运行"）；自动更新；广东程序化生成（广东仍需模板）。

## 8. 交接约定

- 日常改动先本地 commit（`git add -A && git commit`），推送需用户明确说"上传/合并"。
- 后续测试只修改 `tests/test_report_generator.py`，不要新增其他测试程序。
- 测试产物仅写入 `tests/artifacts/`；Git 只保留该目录的 `.gitkeep`。
- 改前端、后端或桥接后均运行 `.venv\Scripts\python.exe -m unittest tests.test_report_generator -v`。
- 动了 `app.py` 或依赖后重新运行 `build.bat`。
- `frontend/index.html` 与原型项目 `C:\FakeD\HermesTeam\projects\报告生成工具桌面化\frontend\报告生成工具-界面原型.html` 同源（原型已另存，如需对照可查）。
- 原程序副本：`C:\文件\工作工具台\报告生成工具V0.1-5.py`（未改动，仅拷贝进 backend 改名 report_engine.py）。
