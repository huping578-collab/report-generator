# 报告生成工具

将原有 Tkinter 报告程序封装为基于 HTML/CSS/JavaScript + Python 的 Windows 桌面应用。界面负责项目配置、文件选择、进度与日志；原 Python 模块继续负责 Excel/Word 统计、图表与报告生成。

## 功能

- 重庆项目：分段汇总、护栏高度、螺栓缺失、图表与 Word 报告
- 广东项目：标线/护栏数据扫描、多地市批量报告、人工复核对比与图表工作簿
- 本机文件选择、自动路径识别、实时运行日志、输出目录快捷打开
- 数据仅在本机处理，不上传到网络

## 模板（Markdown 骨架）

报告模板已切换为 **Markdown 骨架**（`templates/*.md`），程序化章节在锚点注入，docx 仅保留为兼容回退（已移除）：

- **重庆项目**：`templates/重庆项目报告模板.md`（约 7 KB，含封面/注意事项/目录/章节/表格，图片为 `![name](media/name)` 占位）。显式锚点 `<!-- inject:overview -->` / `<!-- inject:height -->` / `<!-- inject:bolt -->` / `<!-- inject:conclusion -->` 优先于关键词（项目概况/整体情况/螺栓/结论）匹配；未命中则追加末尾。图片按模板同级 media 路径渲染。
- **广东项目**：`templates/广东项目第五章模板.md`（封面 + `{{地市}}` 占位，`##` 起章节由程序按区段生成；当前仅封面受模板控制，全章模板驱动需重构 Writer）。

模板为可选：缺失时重庆流程会报错提示仅支持 .md（需补入模板），广东流程同样校验模板存在。

## 本地运行

```bat
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe app.py
```

Windows 11 自带 WebView2；Windows 10 需安装 Microsoft Edge WebView2 Runtime。

## 测试

```bat
.venv\Scripts\python.exe -m unittest tests.test_report_generator -v
```

长期测试程序仅保留 `tests/test_report_generator.py`。测试输出统一写入 `tests/artifacts/`，该目录在 Git 中只保留 `.gitkeep`。

## 构建 Windows 桌面版

```bat
build.bat
```

构建输出：

```text
dist/报告生成工具/报告生成工具.exe
```

分发时请保留整个 `dist/报告生成工具/` 文件夹结构。若模板暂未提供，可在构建后将模板补入该目录下的 `templates/` 文件夹。

## 安装器与更新器

本地构建完整发布包：

```bat
build-release.bat
```

构建前需要：

- `.venv\\Scripts\\python.exe` 已安装 `requirements.txt`。
- Windows 已安装 Inno Setup 6（`ISCC.exe`）。
- Windows 10 已安装 Microsoft Edge WebView2 Runtime。

构建输出：

```text
dist\\报告生成工具-Setup.exe
```

安装器默认安装到 `%LocalAppData%\\Programs\\报告生成工具`，同时创建主程序和更新程序的开始菜单入口。更新程序只检查 GitHub Releases，不从源码分支下载或在用户电脑编译。

发布新版本时：

1. 修改 `version.txt`，例如 `0.1.1`。
2. 提交并推送代码。
3. 创建同版本 tag，例如 `v0.1.1` 并推送：`git push origin v0.1.1`。
4. GitHub Actions 自动构建并发布资产 `报告生成工具-Setup.exe`。

更新程序依赖该精确资产名；如果 Release 没有该资产或网络不可用，会显示错误且不会修改现有安装。

启动后程序会在 `logs/startup.log` 中记录启动过程中的错误；运行时若界面无响应，可检查该日志。

## 目录

```text
app.py                     桌面窗口入口
bridge.py                  HTML 与 Python 的本地桥接
backend/report_engine.py   原报告业务逻辑
frontend/index.html        已确认的桌面界面
frontend/app.js            交互与桌面桥接调用
build-tools/create_icon.py 图标生成脚本
tests/                     唯一 Python 测试程序与测试产物目录
templates/                 Markdown 报告骨架（.md）+ 说明文件
```

## 隐私与分发

- 项目资料、输出报告、Excel、Word 与日志均不提交到 GitHub。
- 构建产物默认不入仓库。
- 未签名 EXE 首次运行可能触发 Windows SmartScreen；选择“更多信息 → 仍要运行”。
