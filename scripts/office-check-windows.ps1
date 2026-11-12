#requires -Version 5.1
[CmdletBinding()]
param([switch] $Worker, [string] $RequestPath)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
. (Join-Path $PSScriptRoot 'verify-windows.ps1') -LibraryOnly

function Write-NewText([string] $Path, [string] $Text) {
    $stream = [IO.File]::Open($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    try {
        $bytes = [Text.UTF8Encoding]::new($false).GetBytes($Text)
        $stream.Write($bytes, 0, $bytes.Length)
    }
    finally { $stream.Dispose() }
}
function Read-PdfPayload([string] $Path, [string] $Name) {
    $file = Get-Item -LiteralPath $Path
    if ($file.Length -gt 25MB -or $file.Length -lt 5) { throw 'PDF 预览超过 25 MiB 或内容为空。请在 Office 里分段导出后核对。' }
    $bytes = [IO.File]::ReadAllBytes($file.FullName)
    if ([Text.Encoding]::ASCII.GetString($bytes, 0, 5) -ne '%PDF-') { throw '这份预览不是有效的 PDF。' }
    return @{ fileName = $Name; base64 = [Convert]::ToBase64String($bytes); sha256 = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
}
function Write-ProgressFile([string] $Directory, [string] $Message) {
    [IO.File]::WriteAllText((Join-Path $Directory 'progress.txt'), $Message, [Text.UTF8Encoding]::new($false))
}
function Invoke-OfficeCheck($Request) {
    $folder = $Request.outputDirectory
    $report = [ordered]@{
        schema = 'officebridge-native-check'; schemaVersion = 1; toolVersion = '1.1.0'
        createdAt = [DateTime]::UtcNow.ToString('o')
        environment = @{ platform = 'windows'; os = [Environment]::OSVersion.VersionString; powerShellVersion = $PSVersionTable.PSVersion.ToString() }
        documents = @(); referencePdf = $null; comparisons = @()
    }
    if ($Request.referencePath) {
        Write-ProgressFile $folder '正在读取原稿参照 PDF…'
        $report.referencePdf = Read-PdfPayload $Request.referencePath ([IO.Path]::GetFileName($Request.referencePath))
    }
    $choices = @(@{ role = 'original'; path = $Request.originalPath })
    if ($Request.copyPath) { $choices += @{ role = 'copy'; path = $Request.copyPath } }
    foreach ($choice in $choices) {
        Write-ProgressFile $folder ('正在用本机 Office 核对「' + [IO.Path]::GetFileName($choice.path) + '」…')
        $pdfName = $choice.role + '.pdf'
        $pdfPath = Join-Path $folder $pdfName
        $item = Test-OneDocument $choice.path $pdfPath (Join-Path $folder 'office-owner.json')
        $item | Add-Member -NotePropertyName role -NotePropertyValue $choice.role
        $item | Add-Member -NotePropertyName format -NotePropertyValue ([IO.Path]::GetExtension($choice.path).TrimStart('.').ToLowerInvariant())
        $item | Add-Member -NotePropertyName pdf -NotePropertyValue $null
        if ($item.pdfExported) {
            try { $item.pdf = Read-PdfPayload $pdfPath ([IO.Path]::GetFileNameWithoutExtension($item.fileName) + '.pdf') }
            catch { $item.pdfError = $_.Exception.Message }
        }
        $report.documents += $item
    }
    if ($report.documents.Count -eq 2 -and $report.documents[0].format -eq 'xlsx' -and $report.documents[1].format -eq 'xlsx') {
        $report.comparisons += Compare-PrintLayout $report.documents[0] $report.documents[1]
    }
    $json = $report | ConvertTo-Json -Depth 14
    if ([Text.Encoding]::UTF8.GetByteCount($json) -gt 110MB) { throw '本次核对记录过大，请分别核对较小的文件。' }
    $jsonPath = Join-Path $folder 'OfficeBridge.officebridge-check.json'
    Write-NewText $jsonPath $json
    $appPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'OfficeBridge.html'
    $resultPath = Join-Path $folder 'OfficeBridge-result.html'
    if ([IO.File]::Exists($appPath)) {
        $html = [IO.File]::ReadAllText($appPath)
        $safe = $json.Replace('<', '\u003c').Replace([string][char]0x2028, '\u2028').Replace([string][char]0x2029, '\u2029')
        $marker = '<script type="application/json" id="officebridge-native-data">' + $safe + '</script>'
        $position = $html.LastIndexOf('</body>', [StringComparison]::OrdinalIgnoreCase)
        if ($position -lt 0) { throw '随包网页内容不完整，请重新解压 OfficeBridge 完整包。' }
        Write-NewText $resultPath $html.Insert($position, $marker)
    }
    else {
        $rows = @($report.documents | ForEach-Object {
            $name = [Net.WebUtility]::HtmlEncode($_.fileName)
            $state = if ($_.openedSuccessfully) { '已在本机 Office 打开' } else { [Net.WebUtility]::HtmlEncode($_.error) }
            '<li><strong>' + $name + '</strong><p>' + $state + '</p></li>'
        }) -join ''
        Write-NewText $resultPath ('<!doctype html><html lang="zh-CN"><meta charset="utf-8"><title>OfficeBridge 核对结果</title><body style="max-width:800px;margin:40px auto;font:18px/1.7 system-ui"><h1>本机 Office 核对结果</h1><ul>' + $rows + '</ul><p>同文件夹中包含 PDF 预览和核对记录。用 OfficeBridge 网页导入 OfficeBridge.officebridge-check.json，可查看完整排版对照。</p></body></html>')
    }
    Write-ProgressFile $folder '核对完成，可以打开结果查看。'
    $incomplete = @($report.documents | Where-Object { !$_.openedSuccessfully -or !$_.sourceUnchanged -or !$_.pdf }).Count
    Write-NewText (Join-Path $folder 'complete.json') (@{ html = $resultPath; report = $jsonPath; incomplete = $incomplete } | ConvertTo-Json)
}

if ($Worker) {
    try {
        if (!$RequestPath) { throw '缺少本次核对参数。' }
        $request = [IO.File]::ReadAllText($RequestPath) | ConvertFrom-Json
        Invoke-OfficeCheck $request
        exit 0
    }
    catch {
        if ($RequestPath -and [IO.File]::Exists($RequestPath)) {
            $directory = Split-Path $RequestPath -Parent
            [IO.File]::WriteAllText((Join-Path $directory 'failed.txt'), $_.Exception.Message, [Text.UTF8Encoding]::new($false))
        }
        exit 1
    }
}

[Windows.Forms.Application]::EnableVisualStyles()
$form = New-Object Windows.Forms.Form
$form.Text = 'OfficeBridge · 本机 Office 核对'
$form.Size = New-Object Drawing.Size(790, 520)
$form.MinimumSize = New-Object Drawing.Size(720, 490)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object Drawing.Font('Microsoft YaHei UI', 10)
$layout = New-Object Windows.Forms.TableLayoutPanel
$layout.Dock = 'Fill'; $layout.Padding = New-Object Windows.Forms.Padding(24)
$layout.ColumnCount = 3; $layout.RowCount = 9
[void]$layout.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Absolute', 135)))
[void]$layout.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent', 100)))
[void]$layout.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Absolute', 84)))
$form.Controls.Add($layout)

$intro = New-Object Windows.Forms.Label
$intro.Text = '选好文件后点“开始核对”。本机的 Excel 或 PowerPoint 会生成预览，结果会在浏览器里打开。'
$intro.Dock = 'Fill'; $intro.AutoSize = $true
$layout.Controls.Add($intro, 0, 0); $layout.SetColumnSpan($intro, 3)
$fields = @{}
$buttons = @()
$labels = @(
    @('originalPath', '要核对的原件', 'Office 文档 (*.pptx;*.xlsx)|*.pptx;*.xlsx'),
    @('copyPath', '修复副本（可选）', 'Office 文档 (*.pptx;*.xlsx)|*.pptx;*.xlsx'),
    @('referencePath', '原稿 PDF（可选）', 'PDF 文件 (*.pdf)|*.pdf'),
    @('outputDirectory', '结果保存位置', '')
)
for ($i = 0; $i -lt $labels.Count; $i++) {
    $record = $labels[$i]
    $label = New-Object Windows.Forms.Label
    $label.Text = $record[1]; $label.Dock = 'Fill'; $label.TextAlign = 'MiddleLeft'
    $box = New-Object Windows.Forms.TextBox
    $box.Name = $record[0]; $box.AccessibleName = $record[1]
    $box.Dock = 'Fill'; $box.Margin = New-Object Windows.Forms.Padding(3, 8, 8, 8)
    $button = New-Object Windows.Forms.Button
    $button.Text = '选择…'; $button.Dock = 'Fill'; $button.Margin = New-Object Windows.Forms.Padding(3, 6, 0, 6)
    $button.Tag = $record[0]; $button.AccessibleName = '选择' + $record[1]
    $fields[$record[0]] = $box; $buttons += $button
    $layout.Controls.Add($label, 0, $i + 1); $layout.Controls.Add($box, 1, $i + 1); $layout.Controls.Add($button, 2, $i + 1)
    $button.Add_Click({
        $key = $this.Tag
        if ($key -eq 'outputDirectory') {
            $picker = New-Object Windows.Forms.FolderBrowserDialog
            $picker.Description = '选择核对结果的保存位置'
            $picker.SelectedPath = $fields[$key].Text
            try { if ($picker.ShowDialog($form) -eq 'OK') { $fields[$key].Text = $picker.SelectedPath } }
            finally { $picker.Dispose() }
        }
        else {
            $picker = New-Object Windows.Forms.OpenFileDialog
            $picker.Filter = if ($key -eq 'referencePath') { 'PDF 文件 (*.pdf)|*.pdf' } else { 'Office 文档 (*.pptx;*.xlsx)|*.pptx;*.xlsx' }
            $picker.Title = if ($key -eq 'originalPath') { '选择要在本机 Office 核对的原件' } elseif ($key -eq 'copyPath') { '选择 OfficeBridge 生成的修复副本' } else { '选择另一台电脑上正确显示时导出的 PDF' }
            try { if ($picker.ShowDialog($form) -eq 'OK') { $fields[$key].Text = $picker.FileName; $fields[$key].SelectionStart = $fields[$key].Text.Length } }
            finally { $picker.Dispose() }
        }
    })
}
$fields.outputDirectory.Text = Join-Path $env:LOCALAPPDATA 'OfficeBridge\Results'
$help = New-Object Windows.Forms.Label
$help.Text = '想核对 Mac → Windows 的效果，可再选一份 Mac 导出的原稿 PDF。只选原件也能生成本机 Office 预览。'
$help.AutoSize = $true; $help.Dock = 'Fill'; $help.ForeColor = [Drawing.Color]::FromArgb(65, 83, 110)
$layout.Controls.Add($help, 0, 5); $layout.SetColumnSpan($help, 3)
$status = New-Object Windows.Forms.Label
$status.AccessibleName = '核对状态'
$status.Text = '文件在本机核对。需要已激活的桌面 Excel 或 PowerPoint。'
$status.Dock = 'Fill'; $status.AutoSize = $true; $status.Padding = New-Object Windows.Forms.Padding(0, 12, 0, 12)
$layout.Controls.Add($status, 0, 6); $layout.SetColumnSpan($status, 3)
$actions = New-Object Windows.Forms.FlowLayoutPanel
$actions.Dock = 'Fill'; $actions.AutoSize = $true
$start = New-Object Windows.Forms.Button
$start.Text = '开始核对'; $start.AutoSize = $true; $start.Padding = New-Object Windows.Forms.Padding(15, 7, 15, 7)
$cancel = New-Object Windows.Forms.Button
$cancel.Text = '停止本次核对'; $cancel.AutoSize = $true; $cancel.Padding = $start.Padding; $cancel.Enabled = $false
$open = New-Object Windows.Forms.Button
$open.Text = '打开结果文件夹'; $open.AutoSize = $true; $open.Padding = $start.Padding; $open.Enabled = $false
$actions.Controls.AddRange(@($start, $cancel, $open))
$layout.Controls.Add($actions, 0, 7); $layout.SetColumnSpan($actions, 3)
$form.AcceptButton = $start

$task = @{ process = $null; folder = $null; started = $null }
$timer = New-Object Windows.Forms.Timer
$timer.Interval = 400
function Set-Working([bool] $Working) {
    $start.Enabled = !$Working; $cancel.Enabled = $Working
    foreach ($field in $fields.Values) { $field.Enabled = !$Working }
    foreach ($button in $buttons) { $button.Enabled = !$Working }
}
function Stop-Check([string] $Message) {
    $timer.Stop()
    if ($task.process -and !$task.process.HasExited) { $task.process.Kill() }
    # A blocking Office dialog may outlive its automation client. Reveal only
    # the exact newly owned window so the user can resolve it normally.
    $ownerPath = if ($task.folder) { Join-Path $task.folder 'office-owner.json' } else { $null }
    if ($ownerPath -and [IO.File]::Exists($ownerPath)) {
        try {
            $owner = [IO.File]::ReadAllText($ownerPath) | ConvertFrom-Json
            $process = Get-Process -Id $owner.pid -ErrorAction Stop
            if ($process.StartTime.ToUniversalTime().ToString('o') -eq $owner.startedUtc) {
                [void][OfficeBridge.NativeWindow]::ShowWindow([IntPtr]$owner.hwnd, 5)
            }
        }
        catch { }
    }
    Set-Working $false
    $open.Enabled = [bool]$task.folder
    $status.Text = $Message
}
$start.Add_Click({
    try {
        if (![IO.File]::Exists($fields.originalPath.Text)) { throw '请先选择要核对的原件。' }
        if ([IO.Path]::GetExtension($fields.originalPath.Text).ToLowerInvariant() -notin @('.pptx', '.xlsx')) { throw '请选择 .pptx 演示文稿或 .xlsx 表格。' }
        if ($fields.copyPath.Text -and ![IO.File]::Exists($fields.copyPath.Text)) { throw '找不到修复副本，请重新选择。' }
        if ($fields.copyPath.Text -and [IO.Path]::GetExtension($fields.originalPath.Text) -ine [IO.Path]::GetExtension($fields.copyPath.Text)) { throw '原件和修复副本应为同一种格式。' }
        if ($fields.referencePath.Text -and ![IO.File]::Exists($fields.referencePath.Text)) { throw '找不到原稿 PDF，请重新选择。' }
        $base = [IO.Path]::GetFullPath($fields.outputDirectory.Text)
        if ($base.StartsWith('\\')) { throw '请选择本机文件夹保存结果。' }
        [void][IO.Directory]::CreateDirectory($base)
        $task.folder = Join-Path $base ('OfficeBridge-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 6))
        [void][IO.Directory]::CreateDirectory($task.folder)
        $request = @{ originalPath = $fields.originalPath.Text; copyPath = $fields.copyPath.Text; referencePath = $fields.referencePath.Text; outputDirectory = $task.folder }
        $requestFile = Join-Path $task.folder 'request.json'
        Write-NewText $requestFile ($request | ConvertTo-Json)
        $info = New-Object Diagnostics.ProcessStartInfo
        $info.FileName = Join-Path $PSHOME 'powershell.exe'
        $info.Arguments = '-NoProfile -STA -ExecutionPolicy Bypass -File "' + $PSCommandPath + '" -Worker -RequestPath "' + $requestFile + '"'
        $info.UseShellExecute = $false; $info.CreateNoWindow = $true
        $task.process = [Diagnostics.Process]::Start($info)
        $task.started = [DateTime]::UtcNow
        Set-Working $true; $open.Enabled = $false
        $status.Text = '正在启动本机 Office，请稍候…'
        $timer.Start()
    }
    catch { $status.Text = $_.Exception.Message }
})
$cancel.Add_Click({ Stop-Check '本次核对已停止。若 Office 仍显示窗口，请先保存您自行打开的文档，再关闭该 Office 窗口后重试。' })
$open.Add_Click({ if ($task.folder) { Start-Process explorer.exe -ArgumentList ('"' + $task.folder + '"') } })
$timer.Add_Tick({
    try {
        if (!$task.process) { return }
        $progress = Join-Path $task.folder 'progress.txt'
        if ([IO.File]::Exists($progress)) { $status.Text = [IO.File]::ReadAllText($progress) }
        if ($task.process.HasExited) {
            $timer.Stop(); Set-Working $false; $open.Enabled = $true
            $complete = Join-Path $task.folder 'complete.json'
            $failed = Join-Path $task.folder 'failed.txt'
            if ([IO.File]::Exists($complete)) {
                $result = [IO.File]::ReadAllText($complete) | ConvertFrom-Json
                $status.Text = if ($result.incomplete -gt 0) { '本次有文件未能完整核对，请在结果页查看原因后重试。' } else { '核对完成，已打开结果。可在结果页查看预览和需要复核的项目。' }
                Start-Process -FilePath $result.html
                foreach ($name in @('request.json', 'progress.txt', 'office-owner.json', 'complete.json')) {
                    Remove-Item -LiteralPath (Join-Path $task.folder $name) -ErrorAction SilentlyContinue
                }
            }
            elseif ([IO.File]::Exists($failed)) { $status.Text = [IO.File]::ReadAllText($failed) }
            else { $status.Text = '本次核对未能完成。请查看 Office 窗口，确认已登录并能正常打开文件后重试。' }
        }
        elseif (([DateTime]::UtcNow - $task.started).TotalSeconds -gt 240) {
            Stop-Check 'Office 响应较慢，已停止核对。请查看 Office 提示；若需关闭该窗口，请先保存您自行打开的文档。'
        }
    }
    catch { Stop-Check ('本次核对中断。' + $_.Exception.Message) }
})
$form.Add_FormClosing({ if ($task.process -and !$task.process.HasExited) { Stop-Check '已关闭本次核对。' } })
try { [void]$form.ShowDialog() }
finally { $timer.Dispose(); $form.Dispose() }
