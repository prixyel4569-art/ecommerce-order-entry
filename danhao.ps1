<#
    单号录入 - 双击即用的单号登记工具
    流程：单号录入\单号\yyyy-MM-dd\单号记录.md
    表格：| 序号 | 单号 | 备注 |
    参数：
      -Add   命令行追加（调试/批处理用）
      -List  命令行列出今日记录
      -SelfTest  构建界面并检查今日/明日目录（自检）
#>
[CmdletBinding()]
param(
    [string]$Root,
    [string[]]$Add,
    [switch]$List,
    [switch]$SelfTest
)

$ErrorActionPreference = 'Stop'

# ---------- 路径与文件 ----------
function Resolve-RootPath {
    param([string]$RootPath)
    if ($RootPath) {
        if (Test-Path -LiteralPath $RootPath) { return (Resolve-Path -LiteralPath $RootPath).Path }
        return [System.IO.Path]::GetFullPath($RootPath)
    }
    if ($PSScriptRoot) { return $PSScriptRoot }
    return (Get-Location).Path
}

function Get-DayInfo {
    param([string]$RootPath, [datetime]$Date)
    $d = $Date.ToString('yyyy-MM-dd')
    $baseDir = Join-Path $RootPath '单号'
    $dir = Join-Path $baseDir $d
    [pscustomobject]@{
        Date    = $d
        BaseDir = $baseDir
        Dir     = $dir
        File    = (Join-Path $dir '单号记录.md')
    }
}

function Get-AvailableDates {
    param([string]$RootPath, [datetime]$Today)
    $dates = @{}
    $dates[$Today.ToString('yyyy-MM-dd')] = $Today.Date
    $baseDir = Join-Path $RootPath '单号'
    if (Test-Path -LiteralPath $baseDir) {
        foreach ($item in (Get-ChildItem -LiteralPath $baseDir -Directory -ErrorAction SilentlyContinue)) {
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParseExact(
                    $item.Name,
                    'yyyy-MM-dd',
                    [System.Globalization.CultureInfo]::InvariantCulture,
                    [System.Globalization.DateTimeStyles]::None,
                    [ref]$parsed)) {
                if ($parsed.Date -le $Today.Date) {
                    $dates[$parsed.ToString('yyyy-MM-dd')] = $parsed.Date
                }
            }
        }
    }
    return @($dates.Values | Sort-Object -Descending)
}

$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Ensure-DayFile {
    param($Info)
    if (-not (Test-Path -LiteralPath $Info.Dir)) {
        New-Item -ItemType Directory -Path $Info.Dir -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $Info.File)) {
        $text = "# 单号记录 - $($Info.Date)`r`n`r`n| 序号 | 单号 | 备注 |`r`n| ---- | ---- | ---- |`r`n"
        [System.IO.File]::WriteAllText($Info.File, $text, $script:Utf8NoBom)
    }
}

function Get-NewLine {
    param($Info)
    $t = [System.IO.File]::ReadAllText($Info.File, [System.Text.Encoding]::UTF8)
    if ($t.Contains("`r`n")) { return "`r`n" }
    return "`n"
}

function Read-Records {
    param($Info)
    Ensure-DayFile $Info
    $text = [System.IO.File]::ReadAllText($Info.File, [System.Text.Encoding]::UTF8)
    $rows = New-Object System.Collections.ArrayList
    foreach ($line in ($text -split "`r?`n")) {
        if ($line -match '^\|\s*(\d+)\s*\|\s*([^|]*?)\s*\|\s*([^|]*?)\s*\|\s*$') {
            $id = $matches[2]
            if ($id -match '^-+$' -or $id -eq '单号') { continue }
            [void]$rows.Add([pscustomobject]@{
                    No   = [int]$matches[1]
                    Id   = $id
                    Note = $matches[3]
                })
        }
    }
    return , $rows
}

function Parse-Ids {
    param([string]$Text)
    $items = New-Object System.Collections.ArrayList
    foreach ($line in ($Text -split "`r?`n")) {
        $l = $line.Trim()
        if (-not $l) { continue }
        foreach ($p in [regex]::Split($l, '[\s,;，；、|]+')) {
            $p = $p.Trim()
            if ($p) { [void]$items.Add($p) }
        }
    }
    return , $items
}

function Add-Records {
    param($Info, [string[]]$Ids, [string]$Note = '')
    $existing = Read-Records $Info
    $seen = @{}
    $maxNo = 0
    foreach ($r in $existing) {
        $seen[$r.Id] = $true
        if ($r.No -gt $maxNo) { $maxNo = $r.No }
    }
    $next = $maxNo + 1
    $nl = Get-NewLine $Info
    $sb = New-Object System.Text.StringBuilder
    $added = New-Object System.Collections.ArrayList
    $dups = New-Object System.Collections.ArrayList
    foreach ($raw in $Ids) {
        $id = ("$raw").Trim()
        if (-not $id) { continue }
        if ($seen.ContainsKey($id)) { [void]$dups.Add($id); continue }
        $seen[$id] = $true
        $safeNote = ("$Note").Replace('|', '/').Trim()
        [void]$sb.Append("| $next | $id | $safeNote |$nl")
        [void]$added.Add($id)
        $next++
    }
    if ($sb.Length -gt 0) {
        $cur = [System.IO.File]::ReadAllText($Info.File, [System.Text.Encoding]::UTF8)
        if (-not $cur.EndsWith("`n") -and -not $cur.EndsWith("`r")) {
            [void]$sb.Insert(0, $nl)
        }
        [System.IO.File]::AppendAllText($Info.File, $sb.ToString(), $script:Utf8NoBom)
    }
    [pscustomobject]@{
        Added = $added
        Dups  = $dups
        Total = (Read-Records $Info).Count
    }
}

# ---------- 命令行模式 ----------
$script:RootPath = Resolve-RootPath $Root
$script:TodayDate = (Get-Date).Date
$script:TodayInfo = Get-DayInfo -RootPath $script:RootPath -Date $script:TodayDate
$script:ViewDate = $script:TodayDate
$script:Info = $script:TodayInfo

if ($Add) {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $ids = Parse-Ids (($Add -join "`n"))
    $res = Add-Records -Info $script:Info -Ids $ids -Note ''
    Write-Host ("文件: " + $script:Info.File)
    Write-Host ("新增: " + $res.Added.Count + " 条 -> " + (($res.Added) -join ', '))
    Write-Host ("重复跳过: " + $res.Dups.Count)
    Write-Host ("今日累计: " + $res.Total)
    exit 0
}

if ($List) {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $rows = Read-Records $script:Info
    Write-Host ("文件: " + $script:Info.File)
    Write-Host ("条数: " + $rows.Count)
    foreach ($r in $rows) { Write-Host ("{0,3}  {1}  {2}" -f $r.No, $r.Id, $r.Note) }
    exit 0
}

# ---------- 图形界面 ----------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$cAccent = [System.Drawing.Color]::FromArgb(47, 111, 237)
$cAccentDark = [System.Drawing.Color]::FromArgb(35, 88, 194)
$cBg = [System.Drawing.Color]::FromArgb(247, 248, 250)
$cText = [System.Drawing.Color]::FromArgb(33, 37, 41)
$cGray = [System.Drawing.Color]::FromArgb(120, 128, 140)
$fontBase = New-Object System.Drawing.Font('微软雅黑', 9.5)
$fontBold = New-Object System.Drawing.Font('微软雅黑', 10, [System.Drawing.FontStyle]::Bold)
$fontTitle = New-Object System.Drawing.Font('微软雅黑', 15, [System.Drawing.FontStyle]::Bold)
$fontMono = New-Object System.Drawing.Font('Consolas', 10.5)

function New-AccentButton {
    param([string]$Text, $X, $Y, $W, $H, [bool]$Primary)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.Location = New-Object System.Drawing.Point($X, $Y)
    $b.Size = New-Object System.Drawing.Size($W, $H)
    $b.FlatStyle = 'Flat'
    $b.FlatAppearance.BorderSize = 1
    $b.Cursor = [System.Windows.Forms.Cursors]::Hand
    $b.Font = $fontBase
    if ($Primary) {
        $b.BackColor = $cAccent
        $b.ForeColor = [System.Drawing.Color]::White
        $b.FlatAppearance.BorderColor = $cAccent
        $b.Font = $fontBold
    }
    else {
        $b.BackColor = [System.Drawing.Color]::White
        $b.ForeColor = $cText
        $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(210, 214, 220)
    }
    return $b
}

$form = New-Object System.Windows.Forms.Form
$form.Text = '单号录入'
$form.Size = New-Object System.Drawing.Size(680, 640)
$form.MinimumSize = New-Object System.Drawing.Size(620, 520)
$form.StartPosition = 'CenterScreen'
$form.BackColor = $cBg
$form.Font = $fontBase
$form.KeyPreview = $true

$header = New-Object System.Windows.Forms.Panel
$header.Dock = 'Top'
$header.Height = 74
$header.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($header)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = '单号录入'
$lblTitle.Font = $fontTitle
$lblTitle.ForeColor = $cText
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(16, 12)
$header.Controls.Add($lblTitle)

$lblViewDate = New-Object System.Windows.Forms.Label
$lblViewDate.Text = '查看日期：'
$lblViewDate.ForeColor = $cGray
$lblViewDate.AutoSize = $true
$lblViewDate.Location = New-Object System.Drawing.Point(406, 16)
$lblViewDate.Anchor = 'Top, Right'
$header.Controls.Add($lblViewDate)

$cmbViewDate = New-Object System.Windows.Forms.ComboBox
$cmbViewDate.DropDownStyle = 'DropDownList'
$cmbViewDate.Font = $fontBase
$cmbViewDate.Location = New-Object System.Drawing.Point(478, 10)
$cmbViewDate.Size = New-Object System.Drawing.Size(168, 28)
$cmbViewDate.Anchor = 'Top, Right'
$header.Controls.Add($cmbViewDate)

$lblPath = New-Object System.Windows.Forms.Label
$lblPath.Text = ('今日文件：' + $script:Info.File)
$lblPath.Font = New-Object System.Drawing.Font('微软雅黑', 8.5)
$lblPath.ForeColor = $cGray
$lblPath.AutoSize = $false
$lblPath.AutoEllipsis = $true
$lblPath.Location = New-Object System.Drawing.Point(18, 46)
$lblPath.Size = New-Object System.Drawing.Size(620, 20)
$lblPath.Anchor = 'Top, Left, Right'
$header.Controls.Add($lblPath)

$lblTip = New-Object System.Windows.Forms.Label
$lblTip.Text = '① 在下方粘贴单号，每行一个（也可用逗号 / 空格 / 顿号分隔），重复的会自动跳过'
$lblTip.ForeColor = $cGray
$lblTip.AutoSize = $true
$lblTip.Location = New-Object System.Drawing.Point(16, 84)
$form.Controls.Add($lblTip)

$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Multiline = $true
$txtInput.ScrollBars = 'Vertical'
$txtInput.AcceptsReturn = $true
$txtInput.Font = $fontMono
$txtInput.BorderStyle = 'FixedSingle'
$txtInput.Location = New-Object System.Drawing.Point(16, 108)
$txtInput.Size = New-Object System.Drawing.Size(632, 116)
$txtInput.Anchor = 'Top, Left, Right'
$form.Controls.Add($txtInput)

$lblNote = New-Object System.Windows.Forms.Label
$lblNote.Text = '② 备注（可选）'
$lblNote.ForeColor = $cGray
$lblNote.AutoSize = $true
$lblNote.Location = New-Object System.Drawing.Point(16, 240)
$form.Controls.Add($lblNote)

$txtNote = New-Object System.Windows.Forms.TextBox
$txtNote.BorderStyle = 'FixedSingle'
$txtNote.Location = New-Object System.Drawing.Point(126, 236)
$txtNote.Size = New-Object System.Drawing.Size(190, 25)
$txtNote.Anchor = 'Top, Left'
$form.Controls.Add($txtNote)

$btnSave = New-AccentButton -Text '保存与追加 (Ctrl+Enter)' -X 330 -Y 234 -W 176 -H 30 -Primary $true
$btnSave.Anchor = 'Top, Right'
$form.Controls.Add($btnSave)

$btnCheck = New-AccentButton -Text '核对差异…' -X 516 -Y 234 -W 132 -H 30 -Primary $false
$btnCheck.Anchor = 'Top, Right'
$form.Controls.Add($btnCheck)

$btnClear = New-AccentButton -Text '清空输入' -X 16 -Y 278 -W 96 -H 28 -Primary $false
$btnClear.Anchor = 'Top, Left'
$form.Controls.Add($btnClear)

$btnOpen = New-AccentButton -Text '打开今日文件夹' -X 122 -Y 278 -W 132 -H 28 -Primary $false
$btnOpen.Anchor = 'Top, Left'
$form.Controls.Add($btnOpen)

$btnRefresh = New-AccentButton -Text '刷新列表' -X 264 -Y 278 -W 96 -H 28 -Primary $false
$btnRefresh.Anchor = 'Top, Left'
$form.Controls.Add($btnRefresh)

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Text = ''
$lblCount.ForeColor = $cAccent
$lblCount.Font = $fontBold
$lblCount.AutoSize = $true
$lblCount.Location = New-Object System.Drawing.Point(374, 283)
$form.Controls.Add($lblCount)

$lv = New-Object System.Windows.Forms.ListView
$lv.View = 'Details'
$lv.FullRowSelect = $true
$lv.GridLines = $true
$lv.HideSelection = $false
$lv.BorderStyle = 'FixedSingle'
$lv.BackColor = [System.Drawing.Color]::White
$lv.Location = New-Object System.Drawing.Point(16, 316)
$lv.Size = New-Object System.Drawing.Size(632, 250)
$lv.Anchor = 'Top, Bottom, Left, Right'
[void]$lv.Columns.Add('序号', 60)
[void]$lv.Columns.Add('单号', 200)
[void]$lv.Columns.Add('备注', 340)
$form.Controls.Add($lv)

$status = New-Object System.Windows.Forms.Label
$status.Dock = 'Bottom'
$status.Height = 28
$status.TextAlign = 'MiddleLeft'
$status.Padding = New-Object System.Windows.Forms.Padding(16, 0, 0, 0)
$status.ForeColor = $cGray
$status.BackColor = [System.Drawing.Color]::White
$status.Text = '就绪'
$form.Controls.Add($status)

function Set-ViewDate {
    param([datetime]$Date)
    $script:ViewDate = $Date.Date
    $script:Info = Get-DayInfo -RootPath $script:RootPath -Date $script:ViewDate
    $isToday = ($script:ViewDate -eq $script:TodayDate)

    $prefix = if ($isToday) { '今日文件：' } else { '历史文件：' }
    $lblPath.Text = $prefix + $script:Info.File
    $txtInput.Enabled = $isToday
    $txtNote.Enabled = $isToday
    $btnSave.Enabled = $isToday
    $btnCheck.Enabled = $isToday
    $btnOpen.Text = if ($isToday) { '打开今日文件夹' } else { '打开所选日期文件夹' }
    $lblTip.Text = if ($isToday) {
        '① 在下方粘贴单号，每行一个（也可用逗号 / 空格 / 顿号分隔），重复的会自动跳过'
    }
    else {
        '正在查看历史记录（只读）；在右上角切换回“今天”后即可继续录入'
    }
    $status.Text = if ($isToday) { '就绪' } else { '正在查看 ' + $script:Info.Date + ' 的历史记录' }
    $status.ForeColor = $cGray
    Update-List
}

function Update-DatePicker {
    param([datetime]$SelectedDate = $script:ViewDate)
    $dates = @(Get-AvailableDates -RootPath $script:RootPath -Today $script:TodayDate)
    if (-not ($dates | Where-Object { $_ -eq $SelectedDate.Date })) {
        $dates += $SelectedDate.Date
        $dates = @($dates | Sort-Object -Descending)
    }

    $script:SuppressDateChange = $true
    $cmbViewDate.BeginUpdate()
    $cmbViewDate.Items.Clear()
    foreach ($date in $dates) {
        $suffix = if ($date.Date -eq $script:TodayDate) { '（今天）' } else { '' }
        [void]$cmbViewDate.Items.Add($date.ToString('yyyy-MM-dd') + $suffix)
    }
    $selectedLabel = $SelectedDate.ToString('yyyy-MM-dd')
    for ($i = 0; $i -lt $cmbViewDate.Items.Count; $i++) {
        if (([string]$cmbViewDate.Items[$i]).StartsWith($selectedLabel)) {
            $cmbViewDate.SelectedIndex = $i
            break
        }
    }
    $cmbViewDate.EndUpdate()
    $script:SuppressDateChange = $false
    Set-ViewDate $SelectedDate
}

function Update-TodayIfNeeded {
    $actualToday = (Get-Date).Date
    if ($actualToday -eq $script:TodayDate) { return }

    $wasViewingToday = ($script:ViewDate -eq $script:TodayDate)
    $script:TodayDate = $actualToday
    $script:TodayInfo = Get-DayInfo -RootPath $script:RootPath -Date $actualToday
    $selectedDate = if ($wasViewingToday) { $actualToday } else { $script:ViewDate }
    Update-DatePicker -SelectedDate $selectedDate
}

function Update-List {
    $lv.BeginUpdate()
    $lv.Items.Clear()
    $rows = Read-Records $script:Info
    foreach ($r in $rows) {
        $it = New-Object System.Windows.Forms.ListViewItem(("" + $r.No))
        [void]$it.SubItems.Add($r.Id)
        [void]$it.SubItems.Add($r.Note)
        [void]$lv.Items.Add($it)
    }
    $lv.EndUpdate()
    if ($script:ViewDate -eq $script:TodayDate) {
        $lblCount.Text = ('今日已记录 ' + $rows.Count + ' 条')
    }
    else {
        $lblCount.Text = ($script:ViewDate.ToString('yyyy-MM-dd') + ' 已记录 ' + $rows.Count + ' 条')
    }
    if ($lv.Items.Count -gt 0) { $lv.Items[$lv.Items.Count - 1].EnsureVisible() }
}

function Do-Save {
    Update-TodayIfNeeded
    if ($script:ViewDate -ne $script:TodayDate) { return }
    $ids = Parse-Ids $txtInput.Text
    if ($ids.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('请先粘贴单号（每行一个）。', '提示', 'OK', 'Information')
        return
    }
    $res = Add-Records -Info $script:TodayInfo -Ids $ids -Note $txtNote.Text
    $txtInput.Clear()
    $txtNote.Clear()
    Update-DatePicker -SelectedDate $script:TodayDate
    $msg = '本次新增 ' + $res.Added.Count + ' 条'
    if ($res.Dups.Count -gt 0) { $msg += ('，重复跳过 ' + $res.Dups.Count + ' 条（' + (($res.Dups) -join '、') + '）') }
    $msg += '；今日累计 ' + $res.Total + ' 条'
    $status.Text = $msg
    $status.ForeColor = $cText
    if ($res.Added.Count -gt 0) { $txtInput.Focus() }
}

$btnSave.Add_Click({ Do-Save })
$btnClear.Add_Click({
    $txtInput.Clear()
    $txtNote.Clear()
    if ($script:ViewDate -eq $script:TodayDate) { $txtInput.Focus() }
})
$btnRefresh.Add_Click({
    Update-DatePicker -SelectedDate $script:ViewDate
    $status.Text = '列表已刷新'
})
$btnOpen.Add_Click({
    Ensure-DayFile $script:Info
    Start-Process explorer.exe -ArgumentList ('"' + $script:Info.Dir + '"')
})
$cmbViewDate.Add_SelectedIndexChanged({
    if ($script:SuppressDateChange -or $cmbViewDate.SelectedIndex -lt 0) { return }
    $text = [string]$cmbViewDate.SelectedItem
    if ($text.Length -lt 10) { return }
    $selectedDate = [datetime]::MinValue
    if ([datetime]::TryParseExact(
            $text.Substring(0, 10),
            'yyyy-MM-dd',
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None,
            [ref]$selectedDate)) {
        Set-ViewDate $selectedDate
    }
})
$form.Add_Activated({ Update-TodayIfNeeded })
$form.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        Do-Save
        $e.SuppressKeyPress = $true
    }
})
$lv.Add_DoubleClick({
    if ($lv.SelectedItems.Count -gt 0) {
        $v = $lv.SelectedItems[0].SubItems[1].Text
        [System.Windows.Forms.Clipboard]::SetText($v)
        $status.Text = ('已复制单号：' + $v)
        $status.ForeColor = $cText
    }
})

# ---------- 核对差异窗口 ----------
function Show-DiffDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = '核对差异（粘贴 Excel 商品ID整列）'
    $dlg.Size = New-Object System.Drawing.Size(620, 520)
    $dlg.StartPosition = 'CenterParent'
    $dlg.BackColor = $cBg
    $dlg.Font = $fontBase
    $dlg.MinimizeBox = $false

    $l1 = New-Object System.Windows.Forms.Label
    $l1.Text = '把商品ID整列粘贴到这里，点「核对」：'
    $l1.AutoSize = $true
    $l1.Location = New-Object System.Drawing.Point(14, 12)
    $dlg.Controls.Add($l1)

    $tp = New-Object System.Windows.Forms.TextBox
    $tp.Multiline = $true
    $tp.ScrollBars = 'Vertical'
    $tp.Font = $fontMono
    $tp.BorderStyle = 'FixedSingle'
    $tp.Location = New-Object System.Drawing.Point(14, 36)
    $tp.Size = New-Object System.Drawing.Size(574, 120)
    $tp.Anchor = 'Top, Left, Right'
    $dlg.Controls.Add($tp)

    $bDiff = New-AccentButton -Text '核对' -X 14 -Y 164 -W 96 -H 30 -Primary $true
    $dlg.Controls.Add($bDiff)

    $bFix = New-AccentButton -Text '把缺少的补登记' -X 120 -Y 164 -W 148 -H 30 -Primary $false
    $bFix.Enabled = $false
    $dlg.Controls.Add($bFix)

    $bClose = New-AccentButton -Text '关闭' -X 506 -Y 164 -W 82 -H 30 -Primary $false
    $bClose.Anchor = 'Top, Right'
    $dlg.Controls.Add($bClose)

    $tr = New-Object System.Windows.Forms.TextBox
    $tr.Multiline = $true
    $tr.ReadOnly = $true
    $tr.ScrollBars = 'Vertical'
    $tr.BorderStyle = 'FixedSingle'
    $tr.BackColor = [System.Drawing.Color]::White
    $tr.Location = New-Object System.Drawing.Point(14, 204)
    $tr.Size = New-Object System.Drawing.Size(574, 262)
    $tr.Anchor = 'Top, Bottom, Left, Right'
    $dlg.Controls.Add($tr)

    $script:Missing = @()

    $bDiff.Add_Click({
        $pasted = Parse-Ids $tp.Text
        $pastedUniq = New-Object System.Collections.ArrayList
        $seenP = @{}
        foreach ($p in $pasted) { if (-not $seenP.ContainsKey($p)) { $seenP[$p] = $true; [void]$pastedUniq.Add($p) } }
        $rows = Read-Records $script:Info
        $seenR = @{}
        foreach ($r in $rows) { $seenR[$r.Id] = $true }

        $missing = @()
        foreach ($p in $pastedUniq) { if (-not $seenR.ContainsKey($p)) { $missing += $p } }
        $extra = @()
        foreach ($r in $rows) { if (-not $seenP.ContainsKey($r.Id)) { $extra += $r.Id } }

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine('粘贴列表（去重后）: ' + $pastedUniq.Count + ' 个')
        [void]$sb.AppendLine('今日已记录: ' + $rows.Count + ' 个')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('缺少 ' + $missing.Count + ' 个（粘贴有、记录没有）:')
        if ($missing.Count -eq 0) { [void]$sb.AppendLine('  （无）') } else { foreach ($m in $missing) { [void]$sb.AppendLine('  ' + $m) } }
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('多余 ' + $extra.Count + ' 个（记录有、粘贴没有）:')
        if ($extra.Count -eq 0) { [void]$sb.AppendLine('  （无）') } else { foreach ($m in $extra) { [void]$sb.AppendLine('  ' + $m) } }
        $tr.Text = $sb.ToString()

        $script:Missing = $missing
        $bFix.Enabled = ($missing.Count -gt 0)
    })

    $bFix.Add_Click({
        if ($script:Missing.Count -eq 0) { return }
        $res = Add-Records -Info $script:Info -Ids $script:Missing -Note '补充（对照导出表补登）'
        $tr.Text = ('已补登记 ' + $res.Added.Count + ' 条：' + (($res.Added) -join '、') + "`r`n今日累计 " + $res.Total + ' 条')
        $bFix.Enabled = $false
        $script:Missing = @()
        Update-List
        $status.Text = '已从核对窗口补登记 ' + $res.Added.Count + ' 条'
        $status.ForeColor = $cText
    })

    $bClose.Add_Click({ $dlg.Close() })
    [void]$dlg.ShowDialog($form)
}

$btnCheck.Add_Click({ Show-DiffDialog })

$form.Add_Shown({
    Update-DatePicker -SelectedDate $script:TodayDate
    if ($script:ViewDate -eq $script:TodayDate) { $txtInput.Focus() }
})

if ($SelfTest) {
    $tomorrow = $script:TodayDate.AddDays(1)
    $tomorrowInfo = Get-DayInfo -RootPath $script:RootPath -Date $tomorrow
    $expectedDir = Join-Path $script:TodayInfo.BaseDir $tomorrow.ToString('yyyy-MM-dd')
    if ($tomorrowInfo.Dir -ne $expectedDir) {
        throw ('明日目录计算错误：' + $tomorrowInfo.Dir)
    }
    if ((Split-Path -Parent $tomorrowInfo.Dir) -ne $script:TodayInfo.BaseDir) {
        throw ('明日目录不在单号文件夹下：' + $tomorrowInfo.Dir)
    }

    $tempBase = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
    $testRoot = [System.IO.Path]::GetFullPath((Join-Path $tempBase ('danhao-selftest-' + [guid]::NewGuid().ToString('N'))))
    $safeTempName = [System.IO.Path]::GetFileName($testRoot).StartsWith('danhao-selftest-')
    if (-not $safeTempName -or -not $testRoot.StartsWith($tempBase)) {
        throw ('临时测试目录不安全：' + $testRoot)
    }
    try {
        $tempTodayInfo = Get-DayInfo -RootPath $testRoot -Date $script:TodayDate
        $tempTomorrowInfo = Get-DayInfo -RootPath $testRoot -Date $tomorrow
        $tempHistoryInfo = Get-DayInfo -RootPath $testRoot -Date $script:TodayDate.AddDays(-1)
        Ensure-DayFile $tempTodayInfo
        Ensure-DayFile $tempTomorrowInfo
        Ensure-DayFile $tempHistoryInfo

        if (-not (Test-Path -LiteralPath $tempTomorrowInfo.File -PathType Leaf)) {
            throw '模拟明日运行时未创建目录或记录文件'
        }
        $testDates = @(Get-AvailableDates -RootPath $testRoot -Today $script:TodayDate)
        if ($testDates.Count -ne 2) {
            throw ('历史日期扫描数量错误：' + $testDates.Count)
        }

        $originalRootPath = $script:RootPath
        $originalTodayInfo = $script:TodayInfo
        $originalInfo = $script:Info
        try {
            $script:RootPath = $testRoot
            $script:TodayInfo = $tempTodayInfo
            $script:Info = $tempTodayInfo
            $script:ViewDate = $script:TodayDate
            Set-ViewDate $script:TodayDate.AddDays(-1)
            if ($txtInput.Enabled -or $btnSave.Enabled -or $btnCheck.Enabled) {
                throw '历史记录查看模式未正确设置只读'
            }
            if ($script:Info.File -ne $tempHistoryInfo.File) {
                throw '历史日期切换到了错误的记录文件'
            }
        }
        finally {
            $script:RootPath = $originalRootPath
            $script:TodayInfo = $originalTodayInfo
            $script:Info = $originalInfo
            $script:ViewDate = $script:TodayDate
        }
    }
    finally {
        if (Test-Path -LiteralPath $testRoot) {
            Remove-Item -LiteralPath $testRoot -Recurse -Force
        }
    }

    Ensure-DayFile $script:TodayInfo
    Update-DatePicker -SelectedDate $script:TodayDate
    $availableDates = @(Get-AvailableDates -RootPath $script:RootPath -Today $script:TodayDate)
    if (-not ($availableDates | Where-Object { $_ -eq $script:TodayDate })) {
        throw '日期选择器未包含今天'
    }
    $form.Dispose()
    Write-Host ('SelfTest OK; history-options=' + $availableDates.Count + '; tomorrow=' + $tomorrowInfo.Dir)
    exit 0
}

[void]$form.ShowDialog()
