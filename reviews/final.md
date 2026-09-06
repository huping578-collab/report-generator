# 终审报告 T7（R1–R6 独立复核）

- 审核人：@reviewer（只读审查＋写报告，未改源码；发现问题只记录不修复）
- 日期：2026-09-05；仓库：`C:/文件/工作工具台/报告生成工具`（分支 main，有未提交改动，未 commit/push）
- 需求：`C:/FakeD/HermesTeam/projects/cq-report-iter3/requirements/brief.md`（R1–R6）；基准画像：`.../requirements/baseline-template-profile.md`（源自 `C:\文件\工作工具台\重庆报告数据\报告模板.docx`）
- 分级口径：P0 阻塞发布 / P1 严重（错数据/错报告）/ P2 一般缺陷（应修）/ P3 轻微或信息项
- R5 视觉逐项对比详见 companion 报告 `reviews/visual-compare.md`，本报告 §R5 为摘要＋分级。

## R1 字体抽查 ✅ 通过

复跑 `C:/FakeD/HermesTeam/projects/cq-report-iter3/font_spotcheck.py`（工作目录为仓库），真实输出：

```
万州区卷 | runs: 107 | latin: ['Times New Roman'] | colors: ['000000']
城口县卷 | runs: 113 | latin: ['Times New Roman'] | colors: ['000000']
```

- 两份样例 docx 非空 run 的 latin 唯一 `Times New Roman`、颜色唯一黑（000000），含封面标题。
- 源码核对：`backend/minimal_docx.py:342` 封面主标题已为 `"latin": "Times New Roman"`（brief 点名的黑体问题已改）；`_set_run_fonts` 统一黑色（`:140`）。
- 测试锁定：`ChongqingFontTimesTests`（封面 latin 非黑体＋全文档 runs 黑/Times 双断言）本次全绿中。
- 分级：无问题。R5 中的 Arial→Times 差异为 R1 要求的**有意偏离**（visual-compare.md D1）。

## R2 去整体报告 ✅ 通过

- 样例目录 `tests/artifacts/cq-real-sample/` 下**无顶层 docx/xlsx**：仅 `城口县/`、`万州区/` 两个子目录（各含 1 docx＋1 xlsx）＋`_generation.log`。
- 测试断言存在且通过：
  - `test_multi_county_files_generate_each_county_report`（约 L1816）：多区县时断言顶层 `OUT_DOCX_NAME`/`OUT_XLSX_NAME` 不存在（L1824–1826 注释 `# R2：多区县时不再生成顶层整体报告`）。
  - `test_single_county_file_filters_segments_and_names_top_files`（约 L1828）：单区县时顶层即区县报告、不建子目录副本（L1835–1836）。
- 分级：无问题。

## R3 数据逻辑 ✅ 通过（附 1 个 P2、2 个 P3）

- 城口县卷含**区县整体＋逐区段小节**：`城口县整体情况`（TCI/螺栓/高度各一）＋逐段 `G211线K1157+874～K1201+315段 / G211线K1202+566～K1377+179段 / G347线K1857+419～K1877+099段`；高度两段 G211 有总结文字＋统计表（T7/T8）＋统计图题注（图x.2-1/2-2 分布）＋示例图题注（图x.2-3）＋示例图文表（T9–T12）；media 8 张图；xlsx 含 `区间01–03` 三个区段 sheet，分段与 docx 一致。
- 与 `0. 各区县路段整理.xlsx` 收敛（省检 sheet）：城口＝row16（G211 1157.874–1201.315）/row17（G211 1202.566–1377.179）/row125（G347 1857.419–1877.099），桩号与报告逐段标题**逐段一致**；万州 9 段＝row20–23（G211×4）/80–81（G318×2）/131–132（G348×2）/209（G542×1），与万州卷 9 个逐段小节及 xlsx `区间01–09` 一致。`_generation.log` 口径一致（城口“G211 2段、G347 1段”；万州“G211 4段、G318 2段、G348 2段、G542 1段”）。
- **P2**：螺栓章节逐段写“本段无波形护栏”，但同区段高度章节有 8999 个有效点（城口 G211 两段；螺栓有效记录 0 条 vs 高度 9032 条）——文案与数据矛盾，建议无螺栓数据时改中性表述或核查数据源。
- **P3**：城口 G347 段在高度章节无小节（该段无高度有效点时无交代，而螺栓章节有“本段无…”占位），建议统一无数据占位策略。
- **P3**：万州卷高度章节仅整体、无逐段无图（media 0），为无数据 fallback，正确但与基准形态不同，特此记录。

## R4 清 Word 模板代码 ✅ 通过

- `grep -rn 'Document(' backend/ bridge.py`：仅两处空白创建——`backend/minimal_docx.py:839`（重庆主链路程序化建文档）与 `backend/report_engine.py:3324`（`GuangdongChapterWriter._write_document`，广东链路）。重庆链路无模板加载。
- 重庆入口 `report_engine.py:1679` 拒收一切非 `.md` 模板（`仅支持 Markdown 模板`），`FileNotFoundError` 守卫完整；残留含“模板”字样的行均为 Markdown 模板语义或广东链路历史命名（`Word模板不存在` 位于广东 writer L3312，brief 硬约束要求 `REPORT_E2E_TEMPLATE_DOCX` 等历史命名不动，属合规残留）。
- 广东测试全绿：本次独立复跑 75 tests 中广东三类（`GuangdongTemplateConfigTests`/`GuangdongBusinessRegressionTests`/桥接广东用例）全部 ok。
- 分级：无问题。

## R5 视觉对比 摘要（详 visual-compare.md）

一致项：封面 8 段＋注意事项逐字一致；目录 TOC 域等效；概况/组织实施骨架与检测依据 10 条一致；核心表头口径一致；题注 `表x.x/图x.x-x` 同体系；表头灰底 D9D9D9；全文全黑。
差异项与分级：D1 R1 有意偏离 Arial→Times（✅）；D2 **P2 结论/建议各重复两轮**（两卷皆然，疑静态模板与注入叠加）；D3 P3 明细载体 docx 附表→xlsx 区间 sheet；D4 P3 TCI 表增病害计数列、缺方向列；D5 P3 万州高度无图（数据驱动）；D6 P2 螺栓文案矛盾（同 R3 P2）；D7 P3 字号 10.5pt vs 基准 10pt、样式名差异（约束内）；D8 P3 螺栓/高度章节顺序与基准相反。

## R6 桌面端去说明文字 ✅ 通过

- `grep <small`：frontend 零命中（仅 CSS 变量 `--radius-small`，非文案）；`grep 说明/副标题`：`index.html`＋`app.js` 零命中。
- 界面抽查（`frontend/index.html` 全文）：无头部副标题、无段落式说明；仅字段 label＋功能性 placeholder（`尚未选择…`/`默认使用项目资料文件夹…`/`未选择时扫描项目文件夹…`/`自动识别（推荐）`），符合 brief 默认口径；`git diff` 显示本次工作区删除了全部 `<small>` 提示与头部副标题行。
- 分级：无问题。

## 独立复跑（7）✅ 真实输出

```
./.venv/Scripts/python.exe -m unittest tests.test_report_generator
................................................................ss...........
Ran 75 tests in 3.468s
OK (skipped=2)
```

- 2 skipped 均为 `RealDataE2ETests`（`set REPORT_E2E_REAL=1 to run` 环境门控，历史预置，非本次引入）。
- 广东用例（`-v` 抽查 15 项，含 writer/toc/表头/桥接/业务回归）全部 ok。

## 结论：**有条件通过**

条件（任一未关闭前不建议发布）：
1. （P2）确认并修复结论/建议双轮重复（D2）：若为模板静态 `5.1/5.2` 与 `inject:conclusion` 叠加，去重其一后重跑测试并重审该节。
2. （P2）处理螺栓“本段无波形护栏”与高度有效数据并存的矛盾（D6/R3-P2）：改中性无数据表述或修正数据源后重审该节。
3. 本次终审基于**未提交工作区**（`backend/minimal_docx.py、backend/report_engine.py、frontend/app.js、frontend/index.html、tests/test_report_generator.py` 5 文件 modified）；发布前固定版本并重跑 ` ./.venv/Scripts/python.exe -m unittest tests.test_report_generator` 全绿。

## 未验证范围

1. 未人工打开 Word 目视（排版分页、页眉页脚、图片渲染、目录域更新后效果）；对比为程序化抽取。
2. 未验证打印/签章/骑缝章等纸质流程要素；封面编号 `BG-2026-` 截断属数据驱动，未深究。
3. `REPORT_E2E_REAL=1` 真实数据广东 e2e（2 skipped）未跑；广东链路仅以回归测试覆盖，未生成广东样例。
4. 未跑性能/大文件压力测试；未验证 24MB 基准模板本身的正确性（基准仅作画像输入）。
5. P3 项（D3/D4/D5/D7/D8：明细载体、TCI 方向列、字号样式名、章节顺序）属可接受差异或待产品确认，未阻塞但建议一并答复。
