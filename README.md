# OfficeBridge

**在把 Mac 制作的文件交给 Windows 之前，检查可见的兼容风险，并按选择生成保留可编辑内容的字体替换副本。**

v0.1.0 支持 PPTX、XLSX，运行时在本机处理，不上传文档。它是规则检查与保守修改工具，不是 Office 排版引擎，也不能保证文件在另一台电脑上完全一致。源码采用 MIT 许可。

## 直接使用

打开发布目录中的 [`release/OfficeBridge.html`](release/OfficeBridge.html)。这是包含所需代码和样例的单个 HTML 文件，无需安装 Node.js 或 Office，可断网使用。请使用可信来源的发布文件。

如果拿到的是源码 ZIP，解压后先按下方“开发与构建”生成这个文件；源码包不重复包含发布产物。便携 ZIP 则可直接解压并打开其中的 `OfficeBridge.html`。

1. 先点“演示文稿”或“电子表格”，用合成样例熟悉报告。
2. 选择或拖入自己的 `.pptx` / `.xlsx` 文件。
3. 阅读风险说明、证据和建议。警告表示需要核对，不等于已经发生错误。
4. 需要改字体时，再主动勾选具体建议；**默认全部不选**。也可在高级选项中指定原字体与目标字体，再重新检查。目标电脑仍须有相应字体。
5. 下载检查报告；采用替换后，另行下载 `.windows.pptx` / `.windows.xlsx` 副本及变更报告。保留原件。
6. 在接收方实际使用的 Windows 软件中打开，核对版面、文字、公式和播放效果。见 [Windows 实际验证说明](docs/WINDOWS-VERIFICATION.md)。

字体替换可能改变字宽、换行及外观；工具不会安装或嵌入字体，不证明字体授权，也不会猜测已经丢失的文字。

## 支持范围

| 内容 | 检查 | 自动修改范围 |
| --- | --- | --- |
| PPTX 字体 | 文字、主题、母版、版式、图表及项目符号的字体声明 | 仅主动选择的字体名称替换 |
| PPTX 其他内容 | 外部资源、媒体、嵌入对象、动画、部分图形效果及字符风险 | 不改写；不把幻灯片或对象转成图片 |
| PPTX 文本边界 | 有限坐标和显式字号的启发式提示 | 不移动文本框，不自动调整字号 |
| XLSX 字体 | 单元格样式、富文本及 DrawingML 字体 | 仅主动选择的字体名称替换 |
| XLSX 数据与功能 | 公式版本及外部依赖、日期系统、长数值与前导零等风险 | 不改公式、数值、日期、链接或打印设置 |
| 已签名或含宏内容 | 安全可读取时提供有限检查 | 禁止修复 |
| Strict OOXML | 有限只读检查 | 禁止修复 |
| Keynote、Numbers、Pages、DOCX、PDF、CSV、旧版 Office 文件 | 当前不支持 | 不处理 |
| 加密、损坏或超过处理上限的文件 | 拒绝处理或报告结构问题 | 不尝试恢复 |

未知的包内内容会被保留。修复会记录修改部件、原字体、目标字体和次数，并逐字节核对未修改部件。重新打包会改变 ZIP 容器；修复副本不是与原件字节完全相同的文件。

### 怎样理解验证状态

- 已完成 macOS 上 Chrome 的浏览器流程检查，以及自动化规则、保真和异常输入测试。
- 面向现代 Chrome / Edge 提供使用入口；**Edge、其他浏览器和 Windows 浏览器环境尚未实际验证**。
- **尚未在真实 Windows 桌面 Office 中完成打开、编辑、版面或播放验证。** 报告中的 `windowsOfficeTested` 固定为 `false`。
- 可选的 Windows PowerShell 验证脚本需要目标电脑上已安装的桌面 Office；它仍属于实验功能，打开成功也不代表版面一致或公式正确。

## 处理上限与隐私

浏览器每批最多 20 个文件、合计 250 MiB。单文件上限 100 MiB；包内最多 5,000 个条目，解压总量 200 MiB，单个条目 100 MiB，单个 XML 12 MiB、XML 总量 40 MiB。大于 1 MiB 的条目超过 300 倍压缩比时拒绝处理；流式解压发现实际内容超过声明，或连续超过 128 KiB 输入没有输出时，也会停止。

浏览器每次检查或修复设有 60 秒超时：超时后停止本次 Worker 并重新初始化，后续文件可继续检查。该计时依赖浏览器调度，不是精确 CPU 限额；CLI 没有自动超时。复杂或恶意文件仍可能消耗较多资源，OfficeBridge **不是杀毒器或不可信文档沙盒**。

应用不上传文稿。导出的报告可能含文件名、内部路径、外部链接、公式片段和其他检查证据；不要未经检查就公开报告或真实文档。更多说明见 [SECURITY.md](SECURITY.md)。

## 开发与构建

需要 Node.js **22.12 或更新版本**。在源码目录执行：

```sh
npm ci
npm test
npm run build
npm run build:portable
```

`npm ci` 安装开发依赖时可能访问 npm；生成的应用处理文档时不联网。`npm run dev` 启动本地开发页面。修改源码后先执行 `npm run build`，再用 `npm run build:portable` 将该次构建生成的文件打包为 `release/OfficeBridge.html`。

```sh
npm run release:local
npm run release:verify
```

本地发布命令在 `release/` 整理单 HTML、便携 ZIP、源码 ZIP 与校验清单；验证命令核对包内源码、当前工作区和 SHA-256。详见 [GitHub 发布步骤](docs/PUBLISHING.md)。当前尚未假定已存在的公开仓库或下载地址；请以实际拿到的本地发布文件为准。

## 命令行批量使用

安装源码依赖后，可直接运行：

```sh
node scripts/cli.mjs check "演示文稿.pptx" "表格.xlsx" --out reports
node scripts/cli.mjs repair "演示文稿.pptx" --accept-font-substitutions --out repaired
node scripts/cli.mjs --help
```

省略命令时默认为 `check`，默认输出目录为 `officebridge-reports/`。检查只写 JSON；`repair` 的确认参数表示接受本次报告中**所有可用的字体替换建议**，建议先检查报告。签名、结构异常等文件不生成修复副本；没有字体建议时仅保存报告。

所有 CLI 输出使用排他创建：已存在的文件、大小写或 Unicode 规范化冲突均不会被覆盖。单项失败后继续处理后续文件；退出码 `0` 为完成，`1` 为存在失败或受阻文件，`2` 为用法错误。报告保存 SHA-256 与验证状态，修复报告另附原件指纹和变更记录。

## 参与与许可

[架构](docs/ARCHITECTURE.md) · [引擎接口](docs/ENGINE-CONTRACT.md) · [测试记录](docs/TESTING.md) · [贡献指南](CONTRIBUTING.md) · [更新记录](CHANGELOG.md)

本项目代码采用 [MIT License](LICENSE)。第三方依赖遵循各自许可，详见 [第三方许可](THIRD_PARTY_NOTICES.md)。本项目不分发专有字体。
