# R13 复核报告（cq-report-iter3）

- 复核人：@coordinator（独立复核；@reviewer 子代理于 2026-09-06 23:25 因上游 HTTP 429 中断、无产出，本次由 coordinator 重新独立运行全部验证命令并目检渲染）
- 复核时间：2026-09-06 23:40（CST）
- 需求：`requirements/brief.md` 的 R13-1 ~ R13-5
- 基准附件：`requirements/报告模板-3.docx`（SHA256 `e6369aeb…ef11`，与用户附件及旧基准副本一致）
- 结论：**通过**（无 P0/P1/P2；P3 仅 1 条记录不阻断）

---

## 1. 验收证据（均为本次复核真实重跑输出）

### 1.1 单元测试
```
./.venv/Scripts/python.exe -m unittest tests.test_report_generator
Ran 103 tests in 5.393s
OK (skipped=2)
```
（103 项全绿，含 R13 新增/改造断言：公式只出现一次、OMML 可编辑、表间存在独立段落、15×8cm、updateFields 缺省关闭。）

### 1.2 真实数据重跑 + 程序化验收
```
regen_cq_real_sample.py   → 城口县 docx 3,194,256 B / 万州区 docx 5,716,585 B（另 xlsx、_generation.log）
work/verify_r13.py        → PASS 万州区: formulas=4 images=9 all=15x8 updateFields=absent externalFields=0
                            PASS 城口县: formulas=4 images=14 all=15x8 updateFields=absent externalFields=0
                            R13 VERIFIED
```
verify_r13 覆盖：
- 21 条附件静态文字逐字存在（含 3 个静态表题、2 个设备图题注）；
- 公式仅 4 个 OMML（TCI 主公式 + GD/w/i₀ 三个下标变量），无纯文本公式残留；
- 全部 `<wp:extent>` = (5400000, 2880000) = 宽15cm × 高8cm；
- settings.xml 无 `w:updateFields`；instrText 无 `INCLUDETEXT/INCLUDEPICTURE/LINK/RD`；关系包无 External TargetMode；
- document.xml 顶层相邻节点不存在 `<w:tbl><w:tbl>` 连续（螺栓结果表与示例表未合并）。

### 1.3 Word 打开/导出（弹窗问题验证）
```
python export_mini.py  → word started / doc opened / exported / DONE 1278454
```
用 Word COM（`Documents.Open` 只读 + `ExportAsFixedFormat`）实际打开生成文档成功导出 PDF（visual/sample-chengkou.pdf，16 页）。打开过程无任何域更新弹窗；程序化检查确认 settings 无 updateFields、无外部文件域——即触发"该文档包含的域可能引用了其他文件"的根因（`w:updateFields` + 外部链接域）已被移除。

### 1.4 人眼目检（Word 渲染 PDF → PNG）
| 页 | 内容 | 目检结果 |
|---|---|---|
| p05 | 2.2/2.3.1、人员表（纵向合并）、扣分表（合并后"1/防护设施缺损/处/0.25"仅显示一次） | 通过：无重复、无乱码、无重叠 |
| p06 | TCI 公式与式中 4 变量（OMML 上下标） | 通过：公式唯一、上下标可辨 |
| p07 | 2.3.2/2.3.3 设备图（15×8）与文字 | 通过：图片正常、无溢出 |
| p14/p15 | 螺栓检测结果表与自动识别示例表 | 通过：两表之间有独立段落，未粘连 |

---

## 2. 需求逐条核验

| 需求 | 实现 | 核验 |
|---|---|---|
| R13-1 已有公式后删除纯文本公式 | 模板改为 `<!-- formula:tci -->` 语义块；渲染器仅输出 OMML；变量改用 `formula:gd/w/i0` 内联公式 | ✅ 文档中 4 个 oMath，无 `TCI = Σ_{i=1}^{i₀}…` 纯文本 |
| R13-2 螺栓结果表与示例表分开 | `_photo_text_table` 前探测 body 末节点（含 sectPr 移位），若上一节点为 tbl 则插独立空段 | ✅ 顶层无相邻 tbl（程序断言）+ p14/p15 目检 |
| R13-3 前两章静态文字与附件逐字一致 | 模板 2.1/2.2/2.3 全部替换为附件原文（含"第一册建工程》"缺书名号、"i  —— "空距等原样瑕疵）；3 个静态表题、2 个设备图题注补入；人员/扣分表纵向合并复刻 | ✅ 21 条逐字存在；动态项目概况（区县/路线/里程）仍按真实数据 |
| R13-4 打开不弹外部域提示 | 重庆模板 `update_on_open=false`；DEFAULT_CONFIG 默认 false；渲染不注入外部文件域；图片均内嵌 | ✅ settings 无 updateFields、无外部域、无 External rels；Word 实开导出无弹窗 |
| R13-5 所有图片 15cm×8cm | `_photo_text_table`/`_example_table`/`_picture` 统一常量 `_REPORT_IMAGE_WIDTH_CM=15.0 / HEIGHT=8.0`，统计图、示例图、设备图全部走同一常量 | ✅ 全 extent 15×8 |

---

## 3. 未发现问题分级

- P0/P1：无。
- P2：无。
- P3（记录，不阻断）：基准附件 2.3.3 首句"依据公路养护工程质量检验评定标准 第一册建工程》(JTG…"原文缺书名号（`《公路养护工程质量检验评定标准 第一册 土建工程》` 被误写成 `第一册建工程》`），已按 R13-3"逐字一致"要求原样保留。若用户希望修正此错别字，属数据/原文修正，需用户确认。

## 4. 已知限制
- 设备图为附件第 1/2 张内嵌图（image1.png/image2.png）导出，另存 `templates/assets/`。
- 目录"错误!未定义书签"为基准模板目录域固有现象（已记录，非本轮引入）。
- 广东链路未改动（git diff 无广东文件）。

## 6. 修正轮（2026-09-07 用户 3 点修正 + 错别字确认）

| 修正 | 实现 | 核验 |
|---|---|---|
| a. 图片尺寸只针对示例图片 | `_picture` 恢复调用方参数；统计图 13×8（折线）/14×8.5（饼图），示例/设备图 `_REPORT_IMAGE_*`=15×8 | ✅ 城口 extent 分布：13×8 ×6 + 14×8.5 ×4 + 15×8 ×4（verify_r13 两卷 PASS） |
| b. 目录未生成 | `_heading` 收集标题 → `_populate_toc_cache` 预填充 TOC 域缓存（updateFields 保持 false，不弹窗）→ `work/finalize_toc.py` Word COM 更新目录固化真实页码 | ✅ PDF p3 目录完整（1概况→1 … 6结论→12，与正文逐条一致）；updateFields 仍缺失（无弹窗） |
| c. 示例图片单元格间不加空格 | `_photo_text_table` 仅当上一节点为多列表格（结果/数据表）时插空段；单列示例表之间不插，Word 渲染自动合并连续排版 | ✅ 城口 5.3 段两示例图合并为一张表（图/文/图/文 连续，p16 文本证实）；多列表格后无相邻 tbl |
| 错别字 | 2.3.3 首句修正为《公路养护工程质量检验评定标准 第一册 土建工程》（JTG 5220-2020） | ✅ 模板/样例/PDF p7 文本三层确认 |

实验结论：`w:updateFields=true` + 纯内部域（TOC/PAGE/NUMPAGES）在 Word COM 打开**不弹窗**（exp_updatefields.py，OPENED OK in 0.4s）；WPS 对 updateFields=true 存在弹窗行为，故正式生成保持 updateFields=false + 目录缓存预填充方案，Word/WPS 均兼容。

## 7. 最终证据

```
./.venv/Scripts/python.exe -m unittest tests.test_report_generator   # 103 OK (skipped=2)
work/regen_cq_real_sample.py        # 城口 3,194,669B / 万州 5,717,210B
work/verify_r13.py                  # 两卷 PASS（公式4/示例图15x8/无外部域/目录缓存）
work/finalize_toc.py                # Word COM 更新目录并保存（页码固化）
export_mini.py + render_png.py      # sample-chengkou.pdf 17 页（目录 1 页+正文 13 页+封面等）
```

## 5. 复核人独立执行的命令清单
```
./.venv/Scripts/python.exe -m unittest tests.test_report_generator
./.venv/Scripts/python.exe work/regen_cq_real_sample.py
./.venv/Scripts/python.exe work/verify_r13.py
python export_mini.py && python render_png.py
```
全部真实执行并记录输出（见 1.1–1.4）。
