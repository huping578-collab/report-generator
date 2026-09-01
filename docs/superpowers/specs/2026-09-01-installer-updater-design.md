# 报告生成工具安装器与更新器设计规格

**目标**：为 Windows 桌面版提供可安装、可卸载的安装包，以及从 GitHub Releases 获取最新版安装包并执行更新的独立更新器。

## 范围

### 安装器

- 使用 Inno Setup 生成 `报告生成工具-Setup.exe`。
- 安装 PyInstaller 产出的整个 `报告生成工具` 目录，而不是只复制主 EXE，保留 WebView2、Python 运行库、前端资源、模板目录和运行日志目录。
- 安装到用户可写的默认目录 `%LocalAppData%\\Programs\\报告生成工具`，避免普通用户写入 `Program Files` 时产生额外权限问题。
- 创建桌面快捷方式、开始菜单快捷方式和卸载入口。
- 安装“报告生成工具”和“检查更新”两个入口。
- 安装器升级时复用同一 AppId，并保留用户提供的模板与生成数据；应用文件由安装器覆盖更新。

### 更新器

- 使用 Python 标准库 `urllib.request`、`json`、`tkinter`、`tempfile`、`hashlib` 和 `subprocess`，不新增运行时依赖。
- 固定访问 `https://api.github.com/repos/huping578-collab/report-generator/releases/latest`。
- 读取本地 `version.txt`，使用语义版本比较判断是否需要更新。
- 只接受 GitHub Releases 中名称精确为 `报告生成工具-Setup.exe` 的资产；下载地址必须为 HTTPS，且主机必须为 `github.com` 或 `objects.githubusercontent.com`。
- 下载到临时目录，校验 HTTP 状态、内容长度和 `MZ` PE 文件头后再启动安装器；下载失败不改动现有安装。
- 无新版本时显示当前版本已是最新；有新版本时显示版本号、发布日期和发布说明摘要。
- 启动安装器后退出更新器，由 Inno Setup 负责覆盖安装和卸载旧版本。
- API 不可用、Release 没有目标资产、版本格式错误或下载文件无效时给出可读错误，不执行安装。

### 版本与发布

- `version.txt` 保存三段式版本号，例如 `0.1.0`，是源码、PyInstaller、Inno Setup 和更新器共同读取的唯一版本来源。
- `build-release.bat` 顺序执行：校验版本 → PyInstaller 构建主程序 → PyInstaller 构建更新器 → Inno Setup 生成安装器。
- GitHub Actions 在推送 `v*.*.*` tag 时构建 Windows 安装器，并创建/更新对应 GitHub Release，上传 `报告生成工具-Setup.exe`。
- 更新器只消费 Release 资产，不从 `main` 分支下载源码或在用户电脑重新编译。

## 非目标

- 本次不实现差分更新、后台静默更新、自动回滚、签名证书和增量补丁。
- 本次不把项目资料、报告、模板、日志或 `.venv` 打进发布资产。
- 本次不自动创建 GitHub Release；由 tag 工作流发布。构建本身不推送代码或发布版本。

## 验收标准

1. 在干净 Windows 用户目录中运行安装器，主程序和更新器均可从开始菜单启动，安装目录包含主 EXE、更新器、`frontend`、`templates` 和运行时文件。
2. 卸载入口可正常移除程序文件，但不删除用户单独放入的模板和输出目录。
3. 更新器能解析成功的 GitHub Release 响应，并筛选目标安装包；无目标资产或非法 URL 时拒绝更新。
4. 版本比较正确处理相同版本、较新版本、较旧版本和 `v` 前缀。
5. 下载中断、HTTP 错误、内容长度不符、非 PE 文件均不会启动安装器。
6. 本地单元测试、PyInstaller 构建、Inno Setup 编译均通过；生成的安装器文件存在且为有效 PE 文件。
7. GitHub Actions 配置语法正确，tag 发布时使用 `version.txt` 构建并上传精确名称的安装器资产。
