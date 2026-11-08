# 用 Windows Excel 对照原稿和副本的打印分页

便携 ZIP 里附有一项可选的桌面 Office 检查。一次选中原稿和 OfficeBridge 生成的表格副本后，它会用这台电脑的 Excel 打开两份文件，列出各工作表的打印页数，以及哪一行、哪一列开始换页。日常使用离线网页无需运行脚本。

## 怎么使用

1. 将便携 ZIP 完整解压到 Windows 本地文件夹。电脑需要已激活的桌面版 Excel，打印预览可正常使用。检查演示文稿时还需安装 PowerPoint，并先保存、关闭已打开的 PowerPoint 窗口。
2. 打开刚解压的文件夹，在资源管理器地址栏输入 `powershell`，按回车。
3. 输入以下命令。在打开的选择框中，按住 Ctrl 同时选中原稿和副本，例如 `预算.xlsx` 和 `预算.windows.xlsx`。选择框一次只能选同一文件夹里的文件。

   ```powershell
   powershell.exe -NoProfile -STA -File .\scripts\verify-windows.ps1
   ```

4. 查看窗口最下方的中文结果。若显示“由 3 页变为 4 页”，下一行会列出原稿和副本在哪些行或列前换页。打开两份文件的打印预览，查看这些位置。报告保存在解压目录的 `work` 文件夹中。

要对照 Mac 上的样子，请使用 Mac 原稿的打印预览或 PDF。上述脚本测量的是同一台 Windows 电脑上两份 Excel 文件的差异。纸张、打印机、缩放、字体和 Excel 版本都会影响结果。两份文件的页数和分页点即使相同，仍要留意文字是否在单元格内换行或被截断。

如果系统策略阻止运行脚本，可以直接打开两份文件，在 Excel 的“文件 → 打印”中核对。若需在单位电脑上运行脚本，请遵循管理员的策略。

## 结果里有哪些信息

| 结果 | 含义 |
| --- | --- |
| 工作表的打印页数和分页点相同 | Excel 为这台电脑的当前打印机计算的结果一致 |
| 打印分页或工作表有变化 | 查找到页数、行列分页点、工作表顺序或可见性差异。窗口会列出具体位置 |
| 未能完整核对打印分页 | 例如打印机不可用、部分分页点读取失败或工作表过多。报告保留已取得的数据 |
| 只看到文件的打开结果 | 请核对同一次选择中有原件和唯一的 `.windows.xlsx` 副本。重名和改名的文件需要手动配对 |
| 单独选择 PPTX | 记录桌面 PowerPoint 能否打开、幻灯片数量和文件指纹。动画和排版请在 PowerPoint 中查看 |

JSON 报告的 `files[].printLayout.sheets` 记录可见工作表的 `pages`、`beforeRows`、`beforeColumns`。隐藏工作表不参与打印分页对照。每个文件最多读取 50 张工作表，每张表最多读取横向及纵向各 128 个分页点。超出时会提示需要人工核对。多个打印区域、图表工作表和特定打印选区仍应通过打印预览复核。`printComparisons` 按原件与副本的文件名配对，`different`、`same`、`unknown` 对应上述三种对照结论。

每份文件还有 `sha256`、`sourceUnchanged` 和对应的 Office 版本。浏览器检查报告独立于桌面报告。脚本全程锁定原件只读，用单独新开的 Excel 实例工作，结束时关闭自己打开的文件，只在没有其他文档时退出新实例。检查期间请等本次操作结束再操作检查窗口。

也可以像下面这样指定路径、要核对的字体和新报告位置。

```powershell
.\scripts\verify-windows.ps1 -Path 'C:\Documents\预算.xlsx','C:\Downloads\预算.windows.xlsx' -TargetFonts 'Arial','Microsoft YaHei' -OutputPath 'C:\Documents\分页对照.json'
```

字体检查记录 Windows 已安装字体的家族名。对照范围覆盖中文本地名和英文名，实际字形、云字体和 Office 的自动替换需要在源文件中核对。报告需要保存到已有本地目录中的新文件名。

## 为什么部分文件会拒绝自动打开

此脚本只对依赖较少的 PPTX / XLSX 进行自动检查。发现宏、XLM 宏工作表、外部关系（包括普通超链接）、数据连接、部分联网函数、链接数据类型、数据模型、网页加载项、OLE 嵌入对象或 ActiveX 控件时，会在 Office 打开之前停止该文件。这是可选验证脚本的保守范围，并不说明这些功能不能在 Windows 使用。

Excel 的 `UpdateLinks=0` 只约束工作簿链接更新，不能被当成对所有外部连接和自动刷新行为的保证。因此脚本在打开前排除外部依赖，并用手动计算模式检查文件；它不调用重算、刷新、保存或修复。微软的参数定义见 [Workbooks.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open) 与 [Application.Calculation](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculation)。

Excel 和 PowerPoint 均设置 `AutomationSecurity=3`，关闭通过自动化打开文件时的宏。微软特别说明，这一设置**不会单独禁用 Excel 4.0 宏**，所以脚本还预先拒绝相关部件和名称。[Excel 宏自动化设置](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity) · [PowerPoint 宏自动化设置](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.automationsecurity)

PowerPoint 使用只读、隐藏窗口方式打开，随后关闭文稿。图片、PDF、视频导出和动画播放需要在桌面应用中另行操作。[PowerPoint 的打开方法](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentations.open)


## 验证记录

正式版开发期间，在 Windows 11 的桌面 Excel、PowerPoint 中测试了同名工作簿配对、文件只读打开和原件指纹。Excel 检查涵盖了由 3 页变为 4 页的长表、保持单页的 Numbers 导出工作簿，以及总页数相同但换页位置从第 18 行移动到第 21 行的样本。另核对隐藏工作表跳过、格式损坏返回“未能完整核对”和只选较少影响排版的位置后分页稳定。测试保留系统的脚本执行策略，通过 Windows PowerShell 5.1 在内存中运行同一份检查器正文。若策略禁用从磁盘运行脚本，可用 Excel 的打印预览完成手工对照。

GitHub 的 Windows 测试另外使用 Windows PowerShell 解析脚本并核对 UTF-8 BOM。普通 CI 机器没有桌面 Office，其结果不能代替真实的排版验证。[项目测试记录](TESTING.md)
