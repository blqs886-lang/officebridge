# 在 Windows Office 中做一次实际打开检查

这个可选脚本会调用 Windows 电脑上**已安装的桌面版 Excel / PowerPoint**，只读打开选中的文件，记录 Office 版本、页数或工作表数量，以及打开前后的文件指纹。它不会修改原文件。

它验证的是“这个文件能否被这台电脑上的 Office 打开”。**打开成功不等于排版一致、公式正确、动画正常或已证明能够继续编辑。** 这些仍需在目标电脑人工核对。浏览器中的 OfficeBridge 报告与此脚本报告是两个独立结果；脚本不会回写浏览器报告中的验证状态。

## 怎么使用

1. 将 OfficeBridge 项目解压到 Windows 的普通本地文件夹。安装的 Office 需要已完成首次启动、激活和登录。
2. 保存手头工作，然后完全关闭 **Excel 和 PowerPoint**。检查期间也不要打开它们。脚本发现它们已经运行时会停止，绝不会替你结束进程。
3. 在项目文件夹空白处按住 Shift，点鼠标右键，选择“在此处打开 PowerShell 窗口”。也可以在资源管理器地址栏输入 `powershell` 后回车。
4. 输入下面这行，再在弹出的文件选择框里选择原稿、修复副本，或两者：

   ```powershell
   powershell.exe -NoProfile -STA -File .\scripts\verify-windows.ps1
   ```

5. 结果文件会保存在项目的 `work` 文件夹，名字以 `OfficeBridge-Windows-` 开头，格式为 JSON。窗口中也会显示成功与失败结果。检查完成后，再用 Office 正常打开文件核对版面。

如果系统策略不允许运行脚本，请先阅读脚本，按所在单位的规定交给管理员运行。无需降低 Office 宏设置，也不要为了运行本脚本关闭受保护视图。如果 Office 出现需要处理的提示，可取消提示并查看报告；脚本不会绕过安全提示。

也可以指定文件和需要核对的字体，例如：

```powershell
.\scripts\verify-windows.ps1 -Path 'C:\Documents\演示文稿.windows.pptx' -TargetFonts 'Arial','Microsoft YaHei'
```

多文件建议使用前面的选择框。自定义输出位置可以加 `-OutputPath 'C:\Documents\Windows-check.json'`；输出目录需要已经存在。脚本会在打开 Office 之前拒绝与输入相同的路径、已有文件、网络输出路径和 Windows 保留文件名，并在最终写入时再次使用“只创建新文件”方式，避免覆盖已有成果。

## 怎么看结果

| 字段 | 含义 |
| --- | --- |
| `fileName` / `sha256` | 被检查的文件名与内容指纹，可与浏览器报告对应 |
| `windowsVersion` / `officeVersion` / `officeBuild` | 本次实际使用的环境；结论不自动适用于另一台电脑或另一个 Office 版本 |
| `actualOfficeOpenPerformed` | 是否实际调用了 Office 的打开方法；前置检查拒绝时为 false |
| `openedSuccessfully` | Office 是否返回了文档对象并读到了页数或工作表数量 |
| `slideCount` / `sheetCount` | 幻灯片数 / 普通工作表数，不含宏工作表；不是内容完整性比较 |
| `sourceUnchanged` | 前后 SHA-256 是否一致；脚本也在检查期间保持原文件只读锁 |
| `visualConsistencyVerified` 等字段 | 固定为 false：本脚本没有做版面比较、计算正确性验证、编辑测试或动画播放 |
| `targetFonts` | 可选字体家族名称是否能被 Windows 字体清单找到，同时核对系统默认、英文及简繁中文名称；不验证中文字形是否齐全、云字体、实际替换或排版效果 |
| `error` / `cleanupNote` | 失败原因，以及是否需要手动关闭此次检查留下的 Office 窗口 |

首次运行失败时，先看 `error`。未安装对应桌面 Office、已有 Office 进程、外部链接、受保护视图或 Office 初始化提示，都可能使自动检查无法完成；这不等同于文件损坏。

## 为什么部分文件会拒绝自动打开

此脚本只对依赖较少的 PPTX / XLSX 进行自动检查。发现宏、XLM 宏工作表、外部关系（包括普通超链接）、数据连接、部分联网函数、链接数据类型、数据模型、网页加载项、OLE 嵌入对象或 ActiveX 控件时，会在 Office 打开之前停止该文件。这是可选验证脚本的保守范围，并不说明这些功能不能在 Windows 使用。

Excel 的 `UpdateLinks=0` 只约束工作簿链接更新，不能被当成对所有外部连接和自动刷新行为的保证。因此脚本在打开前排除外部依赖，并用手动计算模式检查文件；它不调用重算、刷新、保存或修复。微软的参数定义见 [Workbooks.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open) 与 [Application.Calculation](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculation)。

Excel 和 PowerPoint 均设置 `AutomationSecurity=3`，关闭通过自动化打开文件时的宏。微软特别说明，这一设置**不会单独禁用 Excel 4.0 宏**，所以脚本还预先拒绝相关部件和名称。[Excel 宏自动化设置](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity) · [PowerPoint 宏自动化设置](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.automationsecurity)

PowerPoint 使用只读、隐藏窗口方式打开，随后关闭且不保存；方法含义见 [Presentations.Open](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentations.open)。此脚本不会导出图片、PDF 或视频，也不会对公式、宏、加载项和媒体执行测试。

## 会不会影响正在做的工作

运行前必须关闭 Excel 与 PowerPoint。每个文件检查都新建一个应用会话，核对它的窗口所属进程与启动时间；检查结束只关闭自己的文档，只有该新会话已没有其他文档时才请求退出。无法确认归属、发现其他文档，或关闭失败时，会保留进程并写入说明。脚本不使用强制终止进程命令。

如果在检查期间另外打开 Office，仍存在竞争情况，所以不要同时操作这些应用。Office 还可能加载你原本配置的启动加载项；这个脚本不是安全沙盒，适合检查你自己制作和信任的文档。它不会更改系统或 Office 的信任中心策略。

## 当前验证状态与开发者检查

此脚本在 macOS 上编写。**尚未在装有桌面 Office 的 Windows 电脑上完成实际打开测试，因此当前作为可选实验功能提供。** 语法解析或在未安装 Office 的 Windows 上得到预期拒绝，都不能冒充 Office 打开验证。

2026-10-25 已在 Windows 虚拟机的 **Windows PowerShell 5.1.26100.9457** 中通过内存语法解析和前置检查测试：合成普通工作簿可通过包检查，包含连接的包会被拒绝，空输入路径、同名输出路径与 Windows 保留文件名会被拒绝；未安装 Excel / PowerPoint 时，安装检查会按预期拒绝。字体清单测试也确认，中文 Windows 需要核对本地化家族名称，脚本已覆盖这一情况。测试没有启动 Office，也没有更改虚拟机设置。真实 Office 打开行为仍待具有桌面 Office 的环境验证。

开发者可在 Windows PowerShell 中执行下面的只读语法检查；这不会运行脚本或启动 Office：

```powershell
$tokens = $null
$parseErrors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile(
    (Resolve-Path '.\scripts\verify-windows.ps1').Path,
    [ref]$tokens,
    [ref]$parseErrors
)
if ($parseErrors.Count) { $parseErrors | Format-List; exit 1 }
```

仓库的 `.github/workflows/ci.yml` 已包含同样的 Windows PowerShell 5.1 解析检查，也会检查中文脚本所需的 UTF-8 BOM。CI 使用 Node.js 24 在 Windows、macOS 和 Linux 上运行测试与构建，不发布或上传本地 `work` 内容。通用 CI 机器通常没有桌面 Office；要获得真实打开证据，需要用户自己的 Windows Office 环境，或配置了 Office 的专用测试电脑。
