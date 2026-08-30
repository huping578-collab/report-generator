# Design

## Surface
**Operate**：用户配置本地资料并执行一次长任务；操作速度、校验反馈和状态可追踪性优先于营销表达。

## Identity
“工程资料路由台”：以冷白测绘底格、墨蓝工作台和青绿状态色，把文件输入、识别、统计、输出组织成一条可追踪流程。

## Tokens
- Canvas `#F3F6F8`
- Panel `#FFFFFF`
- Ink `#142B3A`
- Secondary ink `#385160`
- Muted `#5B707C`
- Line `#D6E0E5`
- Accent `#0F766E`
- Accent dark `#0B5F59`
- Danger `#B42318`
- Radius `12px` / compact `8px`
- UI font：Microsoft YaHei UI；数据字体：Bahnschrift / DIN Alternate

## Layout
- 左栏：项目模板与真实 4 阶段流程。
- 中栏：资料配置、广东阈值、预计输出。
- 右栏：进度、阶段、实时日志与运行操作。
- 1120px 以下监控区下移；760px 以下改为单列，运行按钮区域保持可达。

## Signature
左侧纵向“资料输入 → 数据识别 → 统计计算 → 报告输出”节点与右侧运行阶段同步，使产品的批处理逻辑成为界面结构，而不是装饰。

## Motion
仅使用 140–180ms 可中断状态过渡；按钮按压缩放至 0.96；遵循 `prefers-reduced-motion`。

## Accessibility
- 语义按钮、标签与标题层级。
- 全局 `:focus-visible`，阈值组使用 `:focus-within`。
- 桌面控件最小 40px，窄屏最小 44px。
- 状态、日志和 Toast 使用 `aria-live`，颜色不是唯一状态信号。
