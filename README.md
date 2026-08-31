# 报告生成工具

将原有 Tkinter 报告程序封装为基于 HTML/CSS/JavaScript + Python 的 Windows 桌面应用。界面负责项目配置、文件选择、进度与日志；原 Python 模块继续负责 Excel/Word 统计、图表与报告生成。

## 功能

- 重庆项目：分段汇总、护栏高度、螺栓缺失、图表与 Word 报告
- 广东项目：标线/护栏数据扫描、多地市批量报告、人工复核对比与图表工作簿
- 本机文件选择、自动路径识别、实时运行日志、输出目录快捷打开
- 数据仅在本机处理，不上传到网络

## Word 模板

原始程序依赖两个 Word 模板。桌面版已内置"程序化生成"模式：

- **重庆项目**：`templates/重庆项目报告模板.docx` 可选。缺失时程序自动切换为程序化生成，产出同结构的检测报告文档（标题、章节、表格、图片、编号均程序生成）。存在该模板时沿用模板内容生成。
- **广东项目**：`templates/广东项目第五章模板.docx` 必需。广东报告复用项目第五章文体与封面结构，未提供该模板时应用会提示并将"开始生成"禁用。

因此可以**完全不提供任何模板直接使用重庆流程**；广东流程需补入原模板。

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
templates/                 用户提供的 Word 模板
```

## 隐私与分发

- 项目资料、输出报告、Excel、Word 与日志均不提交到 GitHub。
- 构建产物默认不入仓库。
- 未签名 EXE 首次运行可能触发 Windows SmartScreen；选择“更多信息 → 仍要运行”。
