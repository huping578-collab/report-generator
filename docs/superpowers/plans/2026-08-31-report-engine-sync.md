# 桌面版报告引擎同步实现计划

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推荐）或 superpowers:executing-plans 逐任务实现此计划。步骤使用复选框（`- [ ]`）语法来跟踪进度。

**目标：** 将最新报告业务引擎完整同步到桌面版，保留桌面资源定位和重庆无模板能力，并把所有自动化测试合并成一个程序。

**架构：** `backend/report_engine.py` 继续作为唯一业务引擎，`bridge.py` 接口不变。同步采用“最新源文件为基线 + 重放两处桌面适配”，而不是逐函数猜测补丁；单一 `tests/test_report_generator.py` 覆盖桌面桥接、无模板重庆和广东关键回归。

**技术栈：** Python 3.11、unittest、openpyxl、python-docx、matplotlib、pywebview、Git。

---

## 文件结构

- 创建：`tests/test_report_generator.py` — 唯一测试程序，含现有测试与最新业务回归。
- 创建：`tests/artifacts/.gitkeep` — 测试过程产物统一目录占位。
- 修改：`backend/report_engine.py` — 最新业务引擎 + 桌面适配。
- 修改：`.gitignore` — 忽略 `tests/artifacts/*`，保留 `.gitkeep`。
- 修改：`README.md`、`HANDOFF.md` — 测试命令与单文件约定。
- 删除：`tests/test_bridge.py`、`tests/test_minimal_docx.py`、`tests/cdp_verify.mjs`。
- 删除：`build-tools/capture-window.ps1`、`launch-and-probe.ps1`、`probe-*.ps1`、`probe_webview.py`、`smoke_webview.py`、`verify-launch.ps1` — 仅用于旧测试/探针流程。
- 保留：`build-tools/create_icon.py`、`build.bat`、`report_generator.spec` — 正式构建链路。

### 任务 1：建立唯一测试程序并得到 RED

**文件：**
- 创建：`tests/test_report_generator.py`
- 读取并迁移：`tests/test_bridge.py`、`tests/test_minimal_docx.py`
- 测试：`tests/test_report_generator.py`

- [ ] **步骤 1：合并现有 8 项测试**

将 `DesktopBridgeTests`、`MinimalDocxTests` 及其构造函数复制进同一个测试模块，导入保持：

```python
from bridge import DesktopBridge
from backend import minimal_docx, report_engine as engine
```

- [ ] **步骤 2：增加桌面适配保护测试**

```python
def test_resource_template_path_uses_application_root(self):
    self.assertEqual(
        engine.resource_template_path("x.docx"),
        engine.application_root() / "templates" / "x.docx",
    )
```

保留现有 `test_engine_routes_to_minimal_without_template`，用于防止同步时丢失 `require_template=False`。

- [ ] **步骤 3：增加最新广东业务 RED 测试**

使用最小 `Document()` 和临时输出目录调用 `GuangdongChapterWriter.write()`，构造三类区段：仅二波、仅三波、二者均无但含桥梁 note。断言：

```python
self.assertIn("其中二波护栏", text)
self.assertIn("其中三波护栏", text)
self.assertEqual(text.count("当前区段为桥梁路段，无有效检测点位。"), 2)
self.assertNotIn("区段内无护栏", text)
self.assertNotIn("共检测0个有效点，其中。", text)
```

另构造 scanner 记录并断言 notes 行仍有：

```python
self.assertEqual(note_row["guardrail_note"], "桥梁地段")
```

对 `_add_table` 生成的 DOCX 原始 XML 断言物理格/逻辑格/网格列 `8/10/10`，`gridSpan=[1,1,1,1,1,2,2,1]`；断言偏差句有“平均偏差”“一致性占比”且无 `min～max` 范围。

- [ ] **步骤 4：运行唯一测试程序验证 RED**

运行：

```bash
python -m unittest tests.test_report_generator -v
```

预期：现有桌面/重庆测试通过；新增广东测试因旧引擎缺少桥隧区段级逻辑或旧表头结构而 FAIL，而不是导入错误。

- [ ] **步骤 5：提交 RED 测试**

```bash
git add tests/test_report_generator.py
git commit -m "test: 汇总桌面版回归并覆盖最新广东规则"
```

### 任务 2：同步最新业务引擎并保留桌面适配

**文件：**
- 修改：`backend/report_engine.py`
- 测试：`tests/test_report_generator.py`

- [ ] **步骤 1：以最新源文件替换业务引擎**

复制：

```text
D:\京炜交通\米奇妙妙屋工具箱\报告生成\报告生成工具V0.1-5.py
→ backend/report_engine.py
```

- [ ] **步骤 2：恢复桌面资源定位**

加入 `import sys`，并在 `builtin_template_paths()` 前恢复：

```python
def application_root():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def resource_template_path(name):
    return application_root() / "templates" / name
```

将重庆模板改为：

```python
"重庆模板": resource_template_path("重庆项目报告模板.docx")
```

- [ ] **步骤 3：恢复重庆无模板入口**

给 `make_docx(..., require_template=True)` 恢复参数；模板不存在时：

```python
if not config.template_docx.exists():
    if not require_template:
        from backend import minimal_docx
        log("未检测到内置 Word 模板，切换到程序化报告生成模式。")
        return minimal_docx.run(
            config, segments,
            height_stats=height_stats, height_records=height_records,
            bolt_stats=bolt_stats, bolt_records=bolt_records,
            disease_image_index=disease_image_index, log=log,
        )
    raise FileNotFoundError(f"Word模板不存在：{config.template_docx}")
```

给 `generate_statistics_and_report(..., require_template=True)` 恢复参数，并向 `make_docx` 传递该值。

- [ ] **步骤 4：运行唯一测试程序验证 GREEN**

```bash
python -m unittest tests.test_report_generator -v
```

预期：全部 PASS。

- [ ] **步骤 5：检查桌面调用签名**

```bash
python -c "from bridge import DesktopBridge; from backend import report_engine as e; print(e.PROGRAM_NAME, e.resource_template_path('x.docx'))"
```

预期：导入成功，路径位于仓库 `templates`。

- [ ] **步骤 6：提交业务同步**

```bash
git add backend/report_engine.py
git commit -m "fix: 同步最新报告业务引擎到桌面版"
```

### 任务 3：整理测试目录并删除旧探针

**文件：**
- 修改：`.gitignore`、`README.md`、`HANDOFF.md`
- 创建：`tests/artifacts/.gitkeep`
- 删除：旧测试与 probe/smoke 文件清单

- [ ] **步骤 1：建立统一产物目录**

`.gitignore` 增加：

```gitignore
tests/artifacts/*
!tests/artifacts/.gitkeep
```

- [ ] **步骤 2：删除多余测试程序和测试探针**

删除文件结构章节列出的旧测试与 probe/smoke 文件，只保留 `tests/test_report_generator.py` 这一份测试程序。

- [ ] **步骤 3：更新文档命令**

README 测试命令改为：

```bat
.venv\Scripts\python.exe -m unittest tests.test_report_generator -v
```

HANDOFF 写明后续测试只修改该文件，产物仅写 `tests/artifacts/`。

- [ ] **步骤 4：验证目录约束**

运行 Python 脚本枚举测试程序，断言仅有：

```text
tests/test_report_generator.py
```

并确认 `build-tools` 只剩正式构建文件。

- [ ] **步骤 5：提交清理**

```bash
git add -A
git commit -m "chore: 统一测试程序和测试产物目录"
```

### 任务 4：完整回归与最终审查

**文件：**
- 验证：`backend/report_engine.py`
- 验证：`tests/test_report_generator.py`

- [ ] **步骤 1：编译与统一测试**

```bash
python -m py_compile app.py bridge.py backend/report_engine.py backend/minimal_docx.py tests/test_report_generator.py
python -m unittest tests.test_report_generator -v
```

预期：退出码 0、0 failures。

- [ ] **步骤 2：运行真实佛山/珠海回归**

测试程序通过环境变量或参数读取仓库外真实数据，所有生成物写入 `tests/artifacts/real-data/`。检查：仅一类波形有数据时不出现无有效点句；两类均无时出现一次；旧句和 0 点残句均为 0；三节共同区段相对顺序一致。

- [ ] **步骤 3：检查仓库卫生**

```bash
git status --short
git diff --check HEAD~3..HEAD
git ls-files tests build-tools
```

确认没有模板、真实数据、输出 DOCX/XLSX、日志、截图或凭据进入 Git。

- [ ] **步骤 4：最终本地提交**

若验证产生文档修正：

```bash
git add -A
git commit -m "docs: 更新桌面版测试与交接说明"
```

不执行 `git push`。
