#requires -Version 5.1
<#
.SYNOPSIS
Read-only opening check in installed Windows desktop Excel / PowerPoint.
.DESCRIPTION
Does not save documents, recalculate formulas, refresh connections, install fonts,
change Trust Center policy, export previews, or establish visual equivalence.
Close ALL Excel and PowerPoint windows before starting; keep them closed during
the check. The script never kills a process. See docs/WINDOWS-VERIFICATION.md.
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)] [ValidateNotNullOrEmpty()] [string[]] $Path,
    [ValidateNotNullOrEmpty()] [string] $OutputPath,
    [ValidateNotNullOrEmpty()] [string[]] $TargetFonts = @()
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw '此可选检查需要 Windows，以及已安装的桌面版 Microsoft Office。'
}
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    throw '请用 -STA 参数启动 Windows PowerShell，详见 docs/WINDOWS-VERIFICATION.md。'
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.Windows.Forms
if (-not ('OfficeBridge.NativeWindow' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace OfficeBridge {
    public static class NativeWindow {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    }
}
'@
}

function Get-OfficeProcesses {
    return ,@(Get-Process -Name EXCEL, POWERPNT -ErrorAction SilentlyContinue)
}

function Assert-OfficeClosed {
    if ((Get-OfficeProcesses).Count -gt 0) {
        throw 'Excel 或 PowerPoint 正在运行。请先保存工作并关闭这两个应用，再重试。脚本没有关闭任何进程。'
    }
}

function Assert-OfficeInstalled([string] $ProgId) {
    if ($null -eq [type]::GetTypeFromProgID($ProgId)) { throw "未安装 ${ProgId}；此检查需要桌面版 Microsoft Office。" }
}

function Get-InstalledFontNames {
    Add-Type -AssemblyName System.Drawing
    $collection = [Drawing.Text.InstalledFontCollection]::new()
    $names = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    try {
        foreach ($family in $collection.Families) {
            [void]$names.Add($family.Name)
            # GDI+ may expose a localized family name on Chinese Windows. Check
            # English and both Chinese locales before reporting a missing name.
            foreach ($locale in @(1033, 2052, 1028)) {
                try { [void]$names.Add($family.GetName($locale)) } catch { }
            }
        }
    }
    finally { $collection.Dispose() }
    return ,$names
}

function Release-ComReference([object] $Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) }
        catch { Write-Verbose 'COM 引用已失效；继续释放其他资源。' }
    }
}

function Resolve-ReportPath([string] $RequestedPath, [string[]] $InputPaths) {
    if (-not $RequestedPath) {
        $workFolder = Join-Path (Split-Path $PSScriptRoot -Parent) 'work'
        [void][IO.Directory]::CreateDirectory($workFolder)
        $RequestedPath = Join-Path $workFolder ('OfficeBridge-Windows-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 6) + '.json')
    }
    if ([string]::IsNullOrWhiteSpace($RequestedPath)) { throw '输出路径不能为空。' }
    $resolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RequestedPath)
    if ($resolved -notmatch '^[A-Za-z]:\\' -or $resolved.StartsWith('\\')) { throw '报告必须保存到本地文件系统路径。' }
    if ([IO.Path]::GetExtension($resolved) -ine '.json') { throw '输出文件路径必须以 .json 结尾。' }
    $leaf = [IO.Path]::GetFileName($resolved)
    if ($leaf.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0) { throw '输出文件名包含 Windows 不允许的字符。' }
    if ($leaf -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)' -or $leaf -match '[. ]$') { throw '输出文件名是 Windows 保留名称或带有无效结尾。' }
    if (-not [IO.Directory]::Exists([IO.Path]::GetDirectoryName($resolved))) { throw '输出目录不存在，请先选择一个已存在的本地文件夹。' }
    foreach ($inputPath in $InputPaths) {
        if ([string]::IsNullOrWhiteSpace($inputPath)) { continue }
        try { $inputResolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($inputPath) }
        catch { continue } # The per-file check will report invalid inputs.
        if ([StringComparer]::OrdinalIgnoreCase.Equals($resolved, $inputResolved)) { throw '报告路径不能与输入文件相同。' }
    }
    if (Test-Path -LiteralPath $resolved) { throw '报告文件已经存在。请选择新文件名；不会覆盖已有成果。' }
    return $resolved
}

function Read-PackageXml([IO.Compression.ZipArchiveEntry] $Entry) {
    if ($Entry.Length -gt 12MB) { throw '某个 XML 部件超过 12 MB 的检查上限。' }
    $settings = New-Object Xml.XmlReaderSettings
    $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $settings.MaxCharactersInDocument = 12MB
    $reader = $null
    $stream = $Entry.Open()
    try {
        $reader = [Xml.XmlReader]::Create($stream, $settings)
        $document = New-Object Xml.XmlDocument
        $document.XmlResolver = $null
        $document.Load($reader)
        return ,$document
    }
    finally {
        if ($null -ne $reader) { $reader.Dispose() }
        $stream.Dispose()
    }
}

function Assert-LocalPackage([IO.Stream] $Stream, [string] $Extension) {
    # This conservative preflight is separate from the browser's package check.
    # UpdateLinks=0 does not cover every connection, OLE object, or online formula.
    # Reject such packages before Office opens them rather than claiming they
    # cannot refresh. Read-only does not, by itself, block external access.
    $archive = $null
    try {
        $Stream.Position = 0
        $archive = [IO.Compression.ZipArchive]::new($Stream, [IO.Compression.ZipArchiveMode]::Read, $true)
        if ($archive.Entries.Count -gt 5000) { throw '文档部件数量超过检查上限。' }
        $names = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        [long]$expanded = 0
        [long]$xmlBytes = 0
        $mainVerified = $false
        foreach ($entry in $archive.Entries) {
            $name = $entry.FullName
            if (-not $names.Add($name)) { throw '文档包含重复或大小写冲突的部件，不能自动检查。' }
            if ($name -match '(^/|\\|(^|/)\.\.(/|$)|[\x00-\x1f])') { throw '文档内部路径不安全，已停止。' }
            $expanded += $entry.Length
            if ($entry.Length -gt 100MB -or $expanded -gt 200MB) { throw '文档解压大小超过检查上限。' }
            if ($entry.Length -gt 1MB -and $entry.Length / [Math]::Max(1, $entry.CompressedLength) -gt 300) { throw '文档压缩比例异常，已停止检查。' }
            if ($name -match '(?i)(vba(project|data|signature)|/(intl)?macrosheets/|/activeX/|/embeddings/|/externalLinks/|/richData/|/webextensions/|/model/|/(connections|queryTable\d*)\.xml$)') {
                throw '文件含宏、外部依赖、控件或嵌入对象，请改用 Office 人工检查。'
            }
            if ($name -notmatch '(?i)\.(xml|rels)$') { continue }
            $xmlBytes += $entry.Length
            if ($xmlBytes -gt 40MB) { throw 'XML 内容超过 40 MB 的检查上限。' }
            $doc = Read-PackageXml $entry
            if ($name -eq '[Content_Types].xml') {
                $expectedPart = if ($Extension -eq '.xlsx') { '/xl/workbook.xml' } else { '/ppt/presentation.xml' }
                $expectedType = if ($Extension -eq '.xlsx') { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' } else { 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml' }
                foreach ($node in $doc.SelectNodes('//*[local-name()="Override" or local-name()="Default"]')) {
                    $contentType = $node.GetAttribute('ContentType')
                    if ($contentType -match '(?i)(macroEnabled|macrosheet|vbaProject|activeX|oleObject|webextension|richvalue|richdata|model\+data)') { throw '文件含宏或活动内容类型，请人工检查。' }
                    if ($node.GetAttribute('PartName') -eq $expectedPart -and $contentType -eq $expectedType) { $mainVerified = $true }
                }
            }
            foreach ($node in $doc.SelectNodes('//*[local-name()="Relationship"]')) {
                if ($node.GetAttribute('TargetMode') -eq 'External') { throw '文件含外部关系（包括超链接），不在本次自动打开检查范围内。' }
                if ($node.GetAttribute('Type') -match '(?i)(vbaProject|xl(intl)?Macrosheet|oleObject|activeX|connections|externalLink|webextension|richdata|richvalue)') { throw '文件含活动内容或外部关系，请人工检查。' }
            }
            if ($doc.SelectNodes('//*[local-name()="connection" or local-name()="queryTable" or local-name()="oleObject" or local-name()="control" or local-name()="externalBook" or local-name()="ddeLink" or local-name()="oleLink"]').Count -gt 0) {
                throw '文件含数据连接、外部链接或嵌入控件，请人工检查。'
            }
            foreach ($node in $doc.SelectNodes('//*[local-name()="definedName"]')) {
                foreach ($flag in @('xlm', 'function', 'vbProcedure')) {
                    if ($node.GetAttribute($flag) -match '^(1|true)$') { throw '名称带有宏或自定义函数标记，请人工检查。' }
                }
            }
            foreach ($node in $doc.SelectNodes('//*[local-name()="f" or local-name()="formula" or local-name()="formula1" or local-name()="formula2" or local-name()="definedName" or local-name()="calculatedColumnFormula" or local-name()="totalsRowFormula"]')) {
                $formula = $node.InnerText -replace '"(?:[^"]|"")*"', '""'
                if ($formula -match '(?i)(?:^|[^A-Z0-9_.])(?:_xlfn\.)?(WEBSERVICE|FILTERXML|IMAGE|STOCKHISTORY|RTD|PY|CUBEVALUE|CUBEMEMBER|CUBESET|CUBESETCOUNT|CUBERANKEDMEMBER|CUBEMEMBERPROPERTY|CUBEKPIMEMBER|CALL|REGISTER|EXEC|EVALUATE)\s*\(' -or $formula -match '(?i)_xll\.') {
                    throw '公式可能调用在线服务、加载项或活动功能，请人工检查。'
                }
                if ($formula -match "'(?:[^']|'')*\[[^\]]+\](?:[^']|'')*'!|(?:^|[^A-Z0-9_])\[[^\]]+\][^'!+\-*/^&=<>(),;\s]+!") {
                    throw '公式引用外部工作簿，请人工检查。'
                }
            }
        }
        if (-not $mainVerified) { throw '此文件不是受支持的无宏 PPTX / XLSX 文档包。' }
    }
    finally { if ($null -ne $archive) { $archive.Dispose() } }
}

function Get-StreamSha256([IO.Stream] $Stream) {
    $hash = [Security.Cryptography.SHA256]::Create()
    try {
        $Stream.Position = 0
        return ([BitConverter]::ToString($hash.ComputeHash($Stream))).Replace('-', '').ToLowerInvariant()
    }
    finally { $hash.Dispose() }
}

function Test-OneDocument([string] $InputPath) {
    $result = [ordered]@{
        fileName = $InputPath
        sha256 = $null; sha256After = $null; sourceUnchanged = $null
        openedSuccessfully = $false; actualOfficeOpenPerformed = $false
        officeApplication = $null; officeVersion = $null; officeBuild = $null
        processOwnershipVerified = $false; sheetCount = $null; slideCount = $null
        visualConsistencyVerified = $false; formulaCalculationVerified = $false
        editabilityVerified = $false; animationAndMediaVerified = $false
        error = $null; cleanupNote = $null
    }
    $source = $null; $app = $null; $documents = $null; $document = $null
    $blank = $null; $items = $null; $owned = $false; $startTime = $null
    try {
        if ([string]::IsNullOrWhiteSpace($InputPath)) { throw '输入文件路径不能为空。' }
        $result.fileName = [IO.Path]::GetFileName($InputPath)
        $file = Get-Item -LiteralPath $InputPath
        if ($file.PSProvider.Name -ne 'FileSystem') { throw '请选择本地文件系统中的文档。' }
        if ($file.PSIsContainer -or $file.Extension -notin @('.pptx', '.xlsx')) { throw '请选择 .pptx 或 .xlsx 文件。' }
        if ($file.FullName.StartsWith('\\')) { throw '请选择本地文件；此检查不直接打开网络共享路径。' }
        if ($file.Length -gt 100MB) { throw '文件超过 100 MB 的检查上限。' }
        $result.fileName = $file.Name
        # Keep a read lock for the entire check; another process cannot overwrite
        # or delete this source while Office is inspecting it.
        $source = [IO.File]::Open($file.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        $result.sha256 = Get-StreamSha256 $source
        Assert-LocalPackage $source ($file.Extension.ToLowerInvariant())
        Assert-OfficeClosed
        $progId = if ($file.Extension -ieq '.xlsx') { 'Excel.Application' } else { 'PowerPoint.Application' }
        $processName = if ($file.Extension -ieq '.xlsx') { 'EXCEL' } else { 'POWERPNT' }
        Assert-OfficeInstalled $progId
        $startTime = [DateTime]::UtcNow
        $app = New-Object -ComObject $progId
        [uint32]$officeProcessId = 0
        [void][OfficeBridge.NativeWindow]::GetWindowThreadProcessId([IntPtr]$app.HWND, [ref]$officeProcessId)
        $process = Get-Process -Id $officeProcessId
        if ($process.ProcessName -ine $processName -or $process.StartTime.ToUniversalTime() -lt $startTime -or (Get-OfficeProcesses).Count -ne 1) {
            throw '无法确认 Office 进程是本次新建的独立会话，已停止。不会终止任何进程。'
        }
        $owned = $true
        $app.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
        $result.processOwnershipVerified = $true
        $result.officeApplication = $progId
        $result.officeVersion = [string]$app.Version
        $result.officeBuild = [string]$process.MainModule.FileVersionInfo.ProductVersion
        if ($file.Extension -ieq '.xlsx') {
            $app.Visible = $false
            $app.EnableEvents = $false
            $app.DisplayAlerts = $false
            $app.AskToUpdateLinks = $false
            $documents = $app.Workbooks
            if ($documents.Count -ne 0) { throw 'Excel 启动时打开了其他工作簿，已停止检查以保护它。' }
            # A first blank workbook establishes manual calculation before the
            # real document is opened, regardless of its saved calcMode.
            $blank = $documents.Add(-4167) # xlWBATWorksheet
            $app.Calculation = -4135 # xlCalculationManual
            $app.CalculateBeforeSave = $false
            $missing = [Type]::Missing
            $app.AutomationSecurity = 3
            $result.actualOfficeOpenPerformed = $true
            $document = $documents.Open($file.FullName, 0, $true, $missing, '', '', $true, $missing, $missing, $false, $false, $missing, $false, $false, 0)
            $items = $document.Worksheets
            $result.sheetCount = [int]$items.Count
        }
        else {
            $documents = $app.Presentations
            if ($documents.Count -ne 0) { throw 'PowerPoint 中有其他演示文稿，已停止检查以保护它。' }
            $app.AutomationSecurity = 3
            $result.actualOfficeOpenPerformed = $true
            $document = $documents.Open($file.FullName, -1, 0, 0) # ReadOnly, named, hidden
            $items = $document.Slides
            $result.slideCount = [int]$items.Count
        }
        $result.openedSuccessfully = $true
    }
    catch { $result.error = $_.Exception.Message }
    finally {
        Release-ComReference $items
        if ($null -ne $document) {
            try {
                if ($result.officeApplication -eq 'Excel.Application') { $document.Close($false) }
                else { $document.Saved = -1; $document.Close() } # Discard in-memory changes; never Save.
            }
            catch { $result.cleanupNote = '无法关闭被检查文档，请手动关闭且不要保存。' }
            Release-ComReference $document
        }
        if ($null -ne $blank) {
            try { $blank.Close($false) } catch { $result.cleanupNote = '无法关闭本次创建的临时空白工作簿。' }
            Release-ComReference $blank
        }
        if ($null -ne $app) {
            try {
                # A user may have opened another document during the check.
                # Quit only our confirmed instance and only when it is empty.
                if ($owned -and $null -ne $documents -and $documents.Count -eq 0) { $app.Quit() }
                else { $result.cleanupNote = '无法确认进程归属或已无其他文档，已保留 Office 窗口；没有强制终止进程。' }
            }
            catch { $result.cleanupNote = 'Office 无法自动退出；没有强制终止进程。' }
        }
        Release-ComReference $documents
        Release-ComReference $app
        if ($null -ne $source) {
            try {
                $result.sha256After = Get-StreamSha256 $source
                $result.sourceUnchanged = $result.sha256After -eq $result.sha256
            }
            catch {
                $result.sourceUnchanged = $null
                $result.cleanupNote = '无法读取检查后的文件指纹，不能确认文件是否保持不变。'
            }
            finally { $source.Dispose() }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    return [pscustomobject]$result
}

Assert-OfficeClosed
if (-not $Path -or $Path.Count -eq 0) {
    $dialog = New-Object Windows.Forms.OpenFileDialog
    try {
        $dialog.Title = 'OfficeBridge — 选择要在 Windows Office 中只读检查的文件'
        $dialog.Filter = 'Office 文档 (*.pptx;*.xlsx)|*.pptx;*.xlsx'
        $dialog.Multiselect = $true
        if ($dialog.ShowDialog() -ne [Windows.Forms.DialogResult]::OK) { Write-Host '已取消，没有打开文件。'; return }
        $Path = $dialog.FileNames
    }
    finally { $dialog.Dispose() }
}

$OutputPath = Resolve-ReportPath $OutputPath $Path

$fontResults = @()
if ($TargetFonts.Count -gt 0) {
    $families = Get-InstalledFontNames
    $fontResults = @($TargetFonts | ForEach-Object { [pscustomobject]@{ name = $_; installedFamilyNameFound = $families.Contains($_) } })
}

$results = @($Path | ForEach-Object { Test-OneDocument $_ })
$report = [ordered]@{
    schemaVersion = 1
    checkerVersion = '0.1.0'
    checkedAtUtc = [DateTime]::UtcNow.ToString('o')
    windowsVersion = [Environment]::OSVersion.VersionString
    powerShellVersion = $PSVersionTable.PSVersion.ToString()
    scope = '仅验证在已安装的 Windows 桌面 Office 中只读打开并读取页数；未验证版面一致性、公式结果、动画或可编辑性。'
    targetFonts = $fontResults
    fontCheckLimitation = '字体家族名称同时核对系统默认、英文及简繁中文名称；不验证字符覆盖、Office 云字体、实际替换或布局。'
    files = $results
}
$json = $report | ConvertTo-Json -Depth 8
$output = [IO.File]::Open($OutputPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
try {
    $encoded = ([Text.UTF8Encoding]::new($false)).GetBytes($json)
    $output.Write($encoded, 0, $encoded.Length)
}
finally { $output.Dispose() }
Write-Host ('报告已保存：' + $OutputPath)
Write-Host '成功打开不代表已经验证版面一致、公式正确或可编辑性。'
$results | Select-Object fileName, openedSuccessfully, sheetCount, slideCount, sourceUnchanged, error | Format-Table -Wrap
