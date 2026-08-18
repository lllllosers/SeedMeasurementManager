#requires -Version 5.1
# =============================================================================
# 草种测定管理 v0.7.0
# -----------------------------------------------------------------------------
# 当前稳定功能：
#   1. Excel 后台连接：隐藏 Excel COM、缓存加速、保存落盘、退出释放
#   2. 今日任务：统计、搜索、DAG 筛选、双击跳转到根苗长录入
#   3. 发芽巡检：培养皿级新增发芽记录、累计发芽率、置床日期与巡检历史
#   4. 测定样本：前10个发芽样本自动进入根/苗长流程，原始坐标可选记录
#   5. 根苗长录入：3/7/14DAG 连续录入、Enter 流转、保存并下一条
#   6. 数据安全：输入校验、累计越界保护、已有测定值防覆盖、保存校验
#
# 本版变更（v0.7.0）：
#   - 新增培养皿级“发芽记录”和实时发芽率管理。
#   - 新增“试验设置”，支持总种子数、取样数、重复等实验参数。
#   - 发芽率巡检与前10个根苗长测定样本解耦。
#   - 原始坐标改为可选信息，可即时填写或后续补录。
#   - 支持0新增巡检、同日多次巡检以及0新增并下一物种。
#
# 依赖：
#   - Windows PowerShell 5.1
#   - Microsoft Excel 桌面版
#
# Excel 工作表约定：
#   根-苗长统计表：
#       D = 样本ID
#       F = 发芽时间
#       G/H/I = 3/7/14DAG 根长
#       J/K/L = 3/7/14DAG 苗长
#
#   测定时间计划表：
#       D = 样本ID
#       L = 下一次测定
#       M = 测定状态
#       N = 备注
#
# 维护建议：
#   - 日常只改“配置区 / UI区 / 对应功能区”，不要随意改 Excel 列映射。
#   - UI 美化统一从“07. UI主题与通用样式”修改，不要在各页面反复写颜色。
#   - 业务写入集中在“06. 样本查询与业务数据操作”，页面事件不要直接写 Excel。
#   - 若需要重新排查启动性能，将 $EnablePerfLog 改为 $true。
# =============================================================================


# =============================================================================
# 01. 基础环境与配置
# =============================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$ErrorActionPreference = 'Stop'

$AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogFile = Join-Path $AppDir 'error.log'
$ConfigFile = Join-Path $AppDir 'last_workbook.txt'

# 启动性能日志：平时关闭，排查性能时改为 $true
$EnablePerfLog = $false
$script:PerfStart = [System.Diagnostics.Stopwatch]::StartNew()


# =============================================================================
# 02. 日志、报错与通用工具
# =============================================================================

function Perf-Log([string]$Text) {
    if (-not $EnablePerfLog) { return }

    try {
        $ms = $script:PerfStart.ElapsedMilliseconds
        Add-Content -LiteralPath (Join-Path $PSScriptRoot 'startup_perf.log') `
            -Value "[$ms ms] $Text" `
            -Encoding UTF8
    }
    catch {}
}

function Write-Log([string]$Text) {
    try {
        $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Add-Content -LiteralPath $LogFile -Value "[$stamp] $Text" -Encoding UTF8
    }
    catch {}
}

function Show-Error([string]$Text) {
    Write-Log $Text
    [System.Windows.Forms.MessageBox]::Show(
        $Text,
        '错误',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Handle-Error([string]$Context, $Err) {
    $detail = $Err.Exception.ToString()

    if ($Err.ScriptStackTrace) {
        $detail += "`r`nPowerShell stack:`r`n" + $Err.ScriptStackTrace
    }

    Write-Log "$Context`r`n$detail"

    [System.Windows.Forms.MessageBox]::Show(
        "$Context：$($Err.Exception.Message)",
        '错误',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Release-Com($Object) {
    if ($null -eq $Object) { return }

    try {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
    catch {}
}

function Safe-Text($Value) {
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim()
}

function ExcelDate-ToText($Value) {
    if ($null -eq $Value -or $Value -eq '') { return '' }

    try {
        if ($Value -is [double] -or $Value -is [int] -or $Value -is [long]) {
            return [DateTime]::FromOADate([double]$Value).ToString('yyyy/M/d')
        }

        if ($Value -is [DateTime]) {
            return ([DateTime]$Value).ToString('yyyy/M/d')
        }
    }
    catch {}

    return [string]$Value
}

function ExcelDate-ToDateTime($Value) {
    if ($null -eq $Value -or $Value -eq '') {
        return $null
    }

    try {
        if (
            $Value -is [double] -or
            $Value -is [float] -or
            $Value -is [decimal] -or
            $Value -is [int] -or
            $Value -is [long]
        ) {
            return [DateTime]::FromOADate([double]$Value)
        }

        if ($Value -is [DateTime]) {
            return [DateTime]$Value
        }

        $parsed = [DateTime]::MinValue

        if (
            [DateTime]::TryParse(
                ([string]$Value).Trim(),
                [ref]$parsed
            )
        ) {
            return $parsed
        }
    }
    catch {}

    return $null
}

function Test-GerminatedValue($Value) {
    # 发芽时间非空即视为“已发芽”。
    if ($null -eq $Value) { return $false }

    $text = ([string]$Value).Trim()

    if ($text -eq '' -or $text -eq '0') {
        return $false
    }

    return $true
}

function Normalize-SpeciesId($Value) {
    # 统一物种编号格式。
    #
    # Excel 可能把文本 001 自动保存为数字 1。
    # 对纯数字编号统一恢复为至少三位：
    #
    # 1   -> 001
    # 12  -> 012
    # 123 -> 123
    #
    # 非纯数字编号保持原样。

    $text = Safe-Text $Value

    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }

    if ($text -match '^\d+$') {
        $number = 0

        if (
            [int]::TryParse(
                $text,
                [ref]$number
            )
        ) {
            return $number.ToString('000')
        }
    }

    return $text
}

function Get-SpeciesIdFromSampleId([string]$SampleId) {
    # 直接从样本ID解析物种编号，保留 003 这样的前导零。
    $sid = $SampleId.Trim()

    if ($sid -match '^(.*)-(\d+)$') {
        return $Matches[1]
    }

    return $sid
}

function Get-CellValue($Sheet, [int]$Row, [int]$Col) {
    $cell = $null

    try {
        $cell = $Sheet.Cells.Item($Row, $Col)
        return $cell.Value2
    }
    finally {
        Release-Com $cell
    }
}

function Set-CellValue($Sheet, [int]$Row, [int]$Col, $Value) {
    # 显式写入 .NET 类型，避免 Excel COM 的 DISP_E_TYPEMISMATCH。
    $cell = $null

    try {
        $cell = $Sheet.Cells.Item([int]$Row, [int]$Col)

        if (
            $Value -is [double] -or
            $Value -is [float] -or
            $Value -is [decimal] -or
            $Value -is [int] -or
            $Value -is [long]
        ) {
            $cell.Value2 = [double]$Value
        }
        elseif ($null -eq $Value) {
            $cell.ClearContents() | Out-Null
        }
        else {
            $cell.Value2 = [string]$Value
        }
    }
    finally {
        Release-Com $cell
    }
}


# =============================================================================
# 03. 全局状态与内存缓存
# =============================================================================

$script:Excel = $null
$script:Book = $null
$script:DataSheet = $null
$script:PlanSheet = $null
$script:GerminationLogSheet = $null
$script:SettingsSheet = $null
$script:WorkbookPath = ''

# 当前实验参数，从“试验设置”工作表读取
$script:ExperimentSettings = [ordered]@{
    TotalSeeds         = 50
    MeasureSampleCount = 10
    DefaultReplicate   = 'R1'
    MeasureDAGs        = @(3, 7, 14)
    GerminationMode    = '每日新增'
}

# 样本ID -> 样本基础信息
$script:DataCache = @{}

# 样本ID -> 测定计划信息
$script:PlanCache = @{}

# 今日需要执行的任务
$script:TodayTaskCache = @()

# “发芽记录”中的全部历史巡检记录
$script:GerminationLogCache = @()

# 培养皿当前发芽状态
# Key = 物种编号|重复，例如 001|R1
$script:GerminationStatusCache = @{}

# 物种编号 -> 发芽巡检坐标（如 E5）
$script:CoordCache = @{}

# 发芽巡检中的全部物种及前10个测定样本状态
$script:GerminationSpeciesCache = @()

# 发芽巡检当前选中的物种
$script:SelectedGerminationSpeciesId = ''

# 清除搜索条件时，避免 TextChanged / SelectedIndexChanged 重复刷新
$script:IgnoreTaskFilterEvents = $false


# =============================================================================
# 04. Excel 生命周期：连接、保存、断开
# =============================================================================


function Disconnect-Workbook {
    # 关闭软件、切换工作簿时统一走这里。
    # 目标：保存 -> 关闭工作簿 -> Quit Excel -> 释放 COM -> 清缓存。

    if ($null -ne $script:Book) {
        try {
            if (-not $script:Book.ReadOnly) {
                $script:Book.Save()
            }
        }
        catch {
            Write-Log "关闭前保存失败：$($_.Exception.Message)"
        }

        try {
            $script:Book.Close($false)
        }
        catch {
            Write-Log "关闭工作簿失败：$($_.Exception.Message)"
        }
    }

    if ($null -ne $script:DataSheet) {
        Release-Com $script:DataSheet
        $script:DataSheet = $null
    }

    if ($null -ne $script:PlanSheet) {
        Release-Com $script:PlanSheet
        $script:PlanSheet = $null
    }

    if ($null -ne $script:GerminationLogSheet) {
        Release-Com $script:GerminationLogSheet
        $script:GerminationLogSheet = $null
    }

    if ($null -ne $script:SettingsSheet) {
        Release-Com $script:SettingsSheet
        $script:SettingsSheet = $null
    }

    if ($null -ne $script:Book) {
        Release-Com $script:Book
        $script:Book = $null
    }

    if ($null -ne $script:Excel) {
        try {
            $script:Excel.DisplayAlerts = $false
            $script:Excel.Quit()
        }
        catch {
            Write-Log "退出 Excel 失败：$($_.Exception.Message)"
        }

        Release-Com $script:Excel
        $script:Excel = $null
    }

    $script:DataCache = @{}
    $script:PlanCache = @{}
    $script:TodayTaskCache = @()
    $script:GerminationLogCache = @()
    $script:GerminationStatusCache = @{}
    $script:CoordCache = @{}
    $script:GerminationSpeciesCache = @()
    $script:SelectedGerminationSpeciesId = ''
    $script:ExperimentSettings = [ordered]@{
        TotalSeeds         = 50
        MeasureSampleCount = 10
        DefaultReplicate   = 'R1'
        MeasureDAGs        = @(3, 7, 14)
        GerminationMode    = '每日新增'
    }
    $script:WorkbookPath = ''

    # 帮助 .NET 释放残余 COM Runtime Callable Wrapper
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Connect-Workbook([string]$Path) {
    Perf-Log '开始 Connect-Workbook'

    Disconnect-Workbook

    if (-not (Test-Path -LiteralPath $Path)) {
        throw 'Excel 文件不存在。'
    }

    $ext = [IO.Path]::GetExtension($Path).ToLower()

    if ($ext -notin @('.xlsx', '.xlsm', '.xlsb', '.xls')) {
        throw '请选择 Excel 工作簿文件。'
    }

    Perf-Log '准备创建 Excel COM'

    try {
        $script:Excel = New-Object -ComObject Excel.Application
    }
    catch {
        throw '无法启动 Microsoft Excel。请确认电脑已安装桌面版 Excel。'
    }

    Perf-Log 'Excel COM 创建完成'

    $script:Excel.Visible = $false
    $script:Excel.DisplayAlerts = $false
    $script:Excel.ScreenUpdating = $false
    $script:Excel.EnableEvents = $false
    $script:Excel.AskToUpdateLinks = $false

    Perf-Log '准备打开工作簿'

    try {
        # UpdateLinks = 0：不更新外部链接
        # ReadOnly = false：以可写模式打开
        $script:Book = $script:Excel.Workbooks.Open($Path, 0, $false)
    }
    catch {
        Disconnect-Workbook
        throw '无法打开工作簿。若该文件正在 Excel 中打开，请先关闭后再连接。'
    }

    Perf-Log '工作簿打开完成'

    if ($script:Book.ReadOnly) {
        Disconnect-Workbook

        throw @"
工作簿当前以只读方式打开。

可能原因：
1. 文件正在被 Excel 或另一个软件实例占用；
2. 上一次软件异常退出后仍有 EXCEL.EXE 后台进程。

请关闭占用该文件的 Excel 后重新连接。
"@
    }

    try {
        $script:DataSheet = $script:Book.Worksheets.Item('根-苗长统计表')
    }
    catch {
        Disconnect-Workbook
        throw '缺少工作表：根-苗长统计表'
    }

    try {
        $script:PlanSheet = $script:Book.Worksheets.Item('测定时间计划表')
    }
    catch {
        Disconnect-Workbook
        throw '缺少工作表：测定时间计划表'
    }

    try {
        $script:GerminationLogSheet = $script:Book.Worksheets.Item('发芽记录')
    }
    catch {
        Disconnect-Workbook
        throw '缺少工作表：发芽记录'
    }

    try {
        $script:SettingsSheet = $script:Book.Worksheets.Item('试验设置')
    }
    catch {
        Disconnect-Workbook
        throw '缺少工作表：试验设置'
    }

    $script:WorkbookPath = (Resolve-Path -LiteralPath $Path).Path

    Set-Content `
        -LiteralPath $ConfigFile `
        -Value $script:WorkbookPath `
        -Encoding UTF8

    # 只计算负责测定逻辑的计划表，不做 Excel 全量 CalculateFull。
    try {
        $script:PlanSheet.Calculate()
    }
    catch {}

    Rebuild-Cache

    Perf-Log '连接与缓存完成'
}


# =============================================================================
# 05. Excel -> 内存缓存
# =============================================================================
function Load-ExperimentSettings {
    # 从“试验设置”工作表读取实验参数。
    # A = 参数名
    # B = 当前值
    #
    # 未识别的参数暂时忽略，便于未来扩展设置表。

    $settings = [ordered]@{
        TotalSeeds         = 50
        MeasureSampleCount = 10
        DefaultReplicate   = 'R1'
        MeasureDAGs        = @(3, 7, 14)
        GerminationMode    = '每日新增'
    }

    $lastCell = $null

    try {
        $lastCell = $script:SettingsSheet.Cells.Item(
            $script:SettingsSheet.Rows.Count,
            1
        ).End(-4162)

        $lastRow = [int]$lastCell.Row
    }
    finally {
        Release-Com $lastCell
    }

    if ($lastRow -lt 2) {
        throw '“试验设置”工作表没有有效参数。'
    }

    $range = $null

    try {
        $range = $script:SettingsSheet.Range("A2:B$lastRow")
        $values = $range.Value2
    }
    finally {
        Release-Com $range
    }

    $lower = $values.GetLowerBound(0)
    $upper = $values.GetUpperBound(0)

    for ($i = $lower; $i -le $upper; $i++) {
        $name = Safe-Text ($values.GetValue($i, 1))
        $value = Safe-Text ($values.GetValue($i, 2))

        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        switch ($name) {
            '总种子数' {
                $parsed = 0

                if (
                    -not [int]::TryParse(
                        $value,
                        [ref]$parsed
                    ) -or
                    $parsed -le 0
                ) {
                    throw '试验设置“总种子数”必须为大于 0 的整数。'
                }

                $settings.TotalSeeds = $parsed
            }

            '根苗长取样数' {
                $parsed = 0

                if (
                    -not [int]::TryParse(
                        $value,
                        [ref]$parsed
                    ) -or
                    $parsed -le 0
                ) {
                    throw '试验设置“根苗长取样数”必须为大于 0 的整数。'
                }

                $settings.MeasureSampleCount = $parsed
            }

            '默认重复' {
                if ([string]::IsNullOrWhiteSpace($value)) {
                    throw '试验设置“默认重复”不能为空。'
                }

                $settings.DefaultReplicate = $value
            }

            '根苗长测定节点' {
                $dagValues = New-Object System.Collections.ArrayList

                foreach ($part in ($value -split ',')) {
                    $text = $part.Trim()
                    $dag = 0

                    if (
                        [string]::IsNullOrWhiteSpace($text) -or
                        -not [int]::TryParse(
                            $text,
                            [ref]$dag
                        ) -or
                        $dag -le 0
                    ) {
                        throw '试验设置“根苗长测定节点”格式无效，应类似：3,7,14'
                    }

                    [void]$dagValues.Add($dag)
                }

                if ($dagValues.Count -eq 0) {
                    throw '试验设置“根苗长测定节点”不能为空。'
                }

                $settings.MeasureDAGs = @($dagValues)
            }

            '发芽记录模式' {
                if ([string]::IsNullOrWhiteSpace($value)) {
                    throw '试验设置“发芽记录模式”不能为空。'
                }

                $settings.GerminationMode = $value
            }
        }
    }

    if ($settings.MeasureSampleCount -gt $settings.TotalSeeds) {
        throw '“根苗长取样数”不能大于“总种子数”。'
    }

    $script:ExperimentSettings = $settings

    Perf-Log (
        "实验设置：" +
        "总种子数=$($settings.TotalSeeds)，" +
        "取样数=$($settings.MeasureSampleCount)，" +
        "重复=$($settings.DefaultReplicate)，" +
        "DAG=$($settings.MeasureDAGs -join ',')，" +
        "发芽记录模式=$($settings.GerminationMode)"
    )
}

function Load-GerminationHistory {
    # -------------------------------------------------------------------------
    # 读取“发芽记录”工作表，并建立两个缓存：
    #
    # GerminationLogCache
    #   保存每一次巡检的原始记录。
    #
    # GerminationStatusCache
    #   按“物种编号 + 重复”汇总培养皿当前状态。
    #
    # 发芽率的唯一核心原始数据是：
    #   H = 本次新增发芽
    #
    # I（累计发芽）和 K（当前发芽率）均视为派生结果，
    # 当前阶段读取时不依赖它们。
    # -------------------------------------------------------------------------

    $script:GerminationLogCache = @()
    $script:GerminationStatusCache = @{}

    $defaultReplicate = [string]$script:ExperimentSettings.DefaultReplicate
    $defaultTotalSeeds = [int]$script:ExperimentSettings.TotalSeeds

    # -------------------------------------------------------------------------
    # 1. 先根据现有样本数据建立默认培养皿状态。
    #
    # 即使某个物种还从未进行过发芽率巡检，
    # 也应该存在：
    #
    # 001|R1 -> 0 / 50
    # -------------------------------------------------------------------------

    foreach ($entry in $script:DataCache.GetEnumerator()) {
        $seed = $entry.Value

        $speciesId = Safe-Text $seed.SpeciesId
        $speciesName = Safe-Text $seed.SpeciesName

        if ([string]::IsNullOrWhiteSpace($speciesId)) {
            continue
        }
        $speciesExists = $false

        foreach (
            $dataItem in
            $script:DataCache.Values
        ) {

            if (
                $dataItem.SpeciesId -eq
                $speciesId
            ) {

                $speciesExists = $true
                break
            }
        }

        if (-not $speciesExists) {
            throw (
                "【发芽记录】第 $excelRow 行：" +
                "物种编号 $speciesId 不存在于当前实验样本中。"
            )
        }

        $key = "$speciesId|$defaultReplicate"

        if (-not $script:GerminationStatusCache.ContainsKey($key)) {
            $script:GerminationStatusCache[$key] = [pscustomobject]@{
                Key                  = $key
                SpeciesId            = $speciesId
                SpeciesName          = $speciesName
                Replicate            = $defaultReplicate
                PlacedDate           = $null
                TotalSeeds           = $defaultTotalSeeds
                CumulativeGerminated = 0
                GerminationRate      = 0.0
                InspectionCount      = 0
                LastInspection       = $null
                LastNewGerminated    = $null
            }
        }
    }

    # -------------------------------------------------------------------------
    # 2. 找到“发芽记录”最后一行。
    # -------------------------------------------------------------------------

    $lastCell = $null

    try {
        $lastCell = $script:GerminationLogSheet.Cells.Item(
            $script:GerminationLogSheet.Rows.Count,
            2
        ).End(-4162)

        $lastRow = [int]$lastCell.Row
    }
    finally {
        Release-Com $lastCell
    }

    # 只有表头，没有历史数据。
    if ($lastRow -lt 2) {
        Perf-Log (
            "发芽记录：0 条；" +
            "培养皿状态=$($script:GerminationStatusCache.Count)"
        )

        return
    }

    # -------------------------------------------------------------------------
    # 3. 一次性读取 A:L，避免逐单元格 COM 调用。
    # -------------------------------------------------------------------------

    $range = $null

    try {
        $range = $script:GerminationLogSheet.Range(
            "A2:L$lastRow"
        )

        $values = $range.Value2
    }
    finally {
        Release-Com $range
    }

    $records = New-Object System.Collections.ArrayList

    $lower = $values.GetLowerBound(0)
    $upper = $values.GetUpperBound(0)

    # -------------------------------------------------------------------------
    # 4. 逐条解析巡检记录。
    # -------------------------------------------------------------------------

    for ($i = $lower; $i -le $upper; $i++) {
        $excelRow = 2 + ($i - $lower)

        $recordId = Safe-Text ($values.GetValue($i, 1))
        $speciesId =
        Normalize-SpeciesId (
            $values.GetValue($i, 2)
        )
        $speciesName = Safe-Text ($values.GetValue($i, 3))
        $replicate = Safe-Text ($values.GetValue($i, 4))

        $placedDateRaw = $values.GetValue($i, 5)
        $inspectionRaw = $values.GetValue($i, 6)

        $newRaw = $values.GetValue($i, 8)
        $totalSeedsRaw = $values.GetValue($i, 10)

        $note = Safe-Text ($values.GetValue($i, 12))

        # B列为空，认为这一行没有有效记录。
        if ([string]::IsNullOrWhiteSpace($speciesId)) {
            continue
        }

        if ([string]::IsNullOrWhiteSpace($replicate)) {
            $replicate = $defaultReplicate
        }

        # ---------------------------------------------------------------------
        # 本次新增发芽必须是 >= 0 的整数。
        # 0 是合法值，表示“已巡检，但无新增发芽”。
        # ---------------------------------------------------------------------

        $newGerminated = 0

        if (
            $null -eq $newRaw -or
            (Safe-Text $newRaw) -eq '' -or
            -not [int]::TryParse(
                (Safe-Text $newRaw),
                [ref]$newGerminated
            ) -or
            $newGerminated -lt 0
        ) {
            throw (
                "【发芽记录】第 $excelRow 行：" +
                "【本次新增发芽】必须为大于等于 0 的整数。"
            )
        }

        # ---------------------------------------------------------------------
        # 每条历史记录保存当时总种子数。
        # 若旧记录为空，则兼容性回退到当前试验设置。
        # ---------------------------------------------------------------------

        $totalSeeds = $defaultTotalSeeds
        $totalSeedsText = Safe-Text $totalSeedsRaw

        if (-not [string]::IsNullOrWhiteSpace($totalSeedsText)) {
            if (
                -not [int]::TryParse(
                    $totalSeedsText,
                    [ref]$totalSeeds
                ) -or
                $totalSeeds -le 0
            ) {
                throw (
                    "【发芽记录】第 $excelRow 行：" +
                    "【总种子数】必须为大于 0 的整数。"
                )
            }
        }

        $inspectionTime = ExcelDate-ToDateTime $inspectionRaw
        $placedDate = ExcelDate-ToDateTime $placedDateRaw

        if ($null -eq $inspectionTime) {
            throw (
                "【发芽记录】第 $excelRow 行：" +
                "【巡检时间】不是有效日期时间。"
            )
        }

        $key = "$speciesId|$replicate"

        # ---------------------------------------------------------------------
        # 如果未来出现 R2、R3，而默认状态中还不存在，
        # 在读到历史记录时自动建立。
        # ---------------------------------------------------------------------

        if (-not $script:GerminationStatusCache.ContainsKey($key)) {
            $script:GerminationStatusCache[$key] = [pscustomobject]@{
                Key                  = $key
                SpeciesId            = $speciesId
                SpeciesName          = $speciesName
                Replicate            = $replicate
                PlacedDate           = $null
                TotalSeeds           = $totalSeeds
                CumulativeGerminated = 0
                GerminationRate      = 0.0
                InspectionCount      = 0
                LastInspection       = $null
                LastNewGerminated    = $null
            }
        }

        $status = $script:GerminationStatusCache[$key]

        # ---------------------------------------------------------------------
        # 置床日期属于培养皿，而不是单个测定样本。
        # 第一条历史记录确定该培养皿的置床日期；
        # 后续历史记录必须保持一致。
        # ---------------------------------------------------------------------

        if ($null -eq $placedDate) {

            throw (
                "【发芽记录】第 $excelRow 行：" +
                "【置床日期】不能为空。"
            )
        }


        if ($null -eq $status.PlacedDate) {

            $status.PlacedDate =
            $placedDate.Date
        }
        elseif (
            ([DateTime]$status.PlacedDate).Date -ne
            $placedDate.Date
        ) {

            throw (
                "【发芽记录】第 $excelRow 行：" +
                "$speciesId / $replicate 的置床日期与此前记录不一致。"
            )
        }

        # 同一培养皿的总种子数在实验过程中不能改变。
        if (
            $status.InspectionCount -gt 0 -and
            [int]$status.TotalSeeds -ne $totalSeeds
        ) {
            throw (
                "【发芽记录】第 $excelRow 行：" +
                "$speciesId / $replicate 的总种子数与此前记录不一致。"
            )
        }

        if ($status.InspectionCount -eq 0) {
            $status.TotalSeeds = $totalSeeds
        }

        # 如果日志中有更完整的物种名称，则补充状态缓存。
        if (
            [string]::IsNullOrWhiteSpace($status.SpeciesName) -and
            -not [string]::IsNullOrWhiteSpace($speciesName)
        ) {
            $status.SpeciesName = $speciesName
        }

        $record = [pscustomobject]@{
            Row            = $excelRow
            RecordId       = $recordId
            SpeciesId      = $speciesId
            SpeciesName    = $speciesName
            Replicate      = $replicate
            PlacedDate     = $placedDate
            InspectionTime = $inspectionTime
            NewGerminated  = $newGerminated
            TotalSeeds     = $totalSeeds
            Note           = $note
        }

        [void]$records.Add($record)

        # ---------------------------------------------------------------------
        # 累计值只由“本次新增发芽”重新计算。
        # 不依赖 Excel I列“累计发芽”。
        # ---------------------------------------------------------------------

        $status.CumulativeGerminated += $newGerminated
        $status.InspectionCount++

        if ($status.CumulativeGerminated -gt $status.TotalSeeds) {
            throw (
                "【发芽记录】第 $excelRow 行：" +
                "$speciesId / $replicate 累计发芽数 " +
                "$($status.CumulativeGerminated) 已超过总种子数 " +
                "$($status.TotalSeeds)。"
            )
        }

        # 最近一次巡检不依赖 Excel 行顺序，而按实际时间判断。
        if (
            $null -eq $status.LastInspection -or
            $inspectionTime -gt $status.LastInspection
        ) {
            $status.LastInspection = $inspectionTime
            $status.LastNewGerminated = $newGerminated
        }
    }

    # -------------------------------------------------------------------------
    # 5. 最后统一计算当前发芽率。
    # -------------------------------------------------------------------------

    foreach ($status in $script:GerminationStatusCache.Values) {
        if ($status.TotalSeeds -gt 0) {
            $status.GerminationRate =
            [double]$status.CumulativeGerminated /
            [double]$status.TotalSeeds
        }
        else {
            $status.GerminationRate = 0.0
        }
    }

    $script:GerminationLogCache = @($records)

    Perf-Log (
        "发芽记录=$($script:GerminationLogCache.Count)，" +
        "培养皿状态=$($script:GerminationStatusCache.Count)"
    )
}

function Rebuild-Cache {
    # 性能关键点：
    # Excel 只在这里批量读取一次，之后所有查询/筛选都在内存完成。

    Perf-Log 'Rebuild-Cache 开始'

    Load-ExperimentSettings

    $script:DataCache = @{}
    $script:PlanCache = @{}
    $script:CoordCache = @{}
    $script:GerminationSpeciesCache = @()

    $todayTasks = New-Object System.Collections.ArrayList

    # -------------------------------------------------------------------------
    # 5.1 根-苗长统计表：A:F
    # -------------------------------------------------------------------------

    $lastCell = $null

    try {
        $lastCell = $script:DataSheet.Cells.Item(
            $script:DataSheet.Rows.Count,
            4
        ).End(-4162)

        $dataLastRow = [int]$lastCell.Row
    }
    finally {
        Release-Com $lastCell
    }

    if ($dataLastRow -ge 2) {
        $range = $null

        try {
            $range = $script:DataSheet.Range("A2:F$dataLastRow")
            $values = $range.Value2
        }
        finally {
            Release-Com $range
        }

        $lower = $values.GetLowerBound(0)
        $upper = $values.GetUpperBound(0)

        for ($i = $lower; $i -le $upper; $i++) {
            $sampleId = Safe-Text ($values.GetValue($i, 4))

            if ([string]::IsNullOrWhiteSpace($sampleId)) {
                continue
            }

            $excelRow = 2 + ($i - $lower)

            $script:DataCache[$sampleId] = [pscustomobject]@{
                Row         = $excelRow
                SpeciesId   = Get-SpeciesIdFromSampleId $sampleId
                SpeciesName = Safe-Text ($values.GetValue($i, 2))
                SeedNo      = Safe-Text ($values.GetValue($i, 3))
                PlacedDate  = $values.GetValue($i, 5)
                Germination = $values.GetValue($i, 6)
            }
        }
    }

    # -------------------------------------------------------------------------
    # 5.1B 发芽历史与培养皿当前状态
    # -------------------------------------------------------------------------

    Load-GerminationHistory

    # -------------------------------------------------------------------------
    # 5.2 从“测定时间计划表”备注列 N 读取原始坐标
    #
    # D = 样本ID
    # N = 备注 / 原始坐标
    #
    # 例如：
    # 001-1 -> E3
    # 001-2 -> A5
    #
    # 这里只把符合 A1～E10 格式的备注识别为坐标。
    # -------------------------------------------------------------------------

    $coordLastCell = $null

    try {

        $coordLastCell =
        $script:PlanSheet.Cells.Item(
            $script:PlanSheet.Rows.Count,
            4
        ).End(-4162)

        $coordLastRow =
        [int]$coordLastCell.Row
    }
    finally {

        Release-Com $coordLastCell
    }


    if ($coordLastRow -ge 2) {

        $coordRange = $null

        try {

            # D:N
            # D在数组第1列
            # N在数组第11列

            $coordRange =
            $script:PlanSheet.Range(
                "D2:N$coordLastRow"
            )

            $coordValues =
            $coordRange.Value2
        }
        finally {

            Release-Com $coordRange
        }


        $clower =
        $coordValues.GetLowerBound(0)

        $cupper =
        $coordValues.GetUpperBound(0)


        for (
            $i = $clower;
            $i -le $cupper;
            $i++
        ) {

            $sampleId =
            Safe-Text (
                $coordValues.GetValue(
                    $i,
                    1
                )
            )


            if (
                [string]::IsNullOrWhiteSpace(
                    $sampleId
                )
            ) {
                continue
            }


            $note =
            (
                Safe-Text (
                    $coordValues.GetValue(
                        $i,
                        11
                    )
                )
            ).ToUpperInvariant()


            # 只有A1～E10才识别为原始坐标
            if (
                $note -notmatch
                '^[A-E](10|[1-9])$'
            ) {
                continue
            }


            $speciesName = ''

            if (
                $script:DataCache.ContainsKey(
                    $sampleId
                )
            ) {

                $speciesName =
                $script:DataCache[
                $sampleId
                ].SpeciesName
            }


            $script:CoordCache[$sampleId] =
            [pscustomobject]@{

                Row         =
                2 + ($i - $clower)

                SpeciesId   =
                Get-SpeciesIdFromSampleId `
                    $sampleId

                SpeciesName =
                $speciesName

                Coordinate  =
                $note
            }
        }
    }
    
    # -------------------------------------------------------------------------
    # 5.3 建立发芽巡检物种缓存
    #
    # 注意：
    # 每个物种有50粒实际种子，
    # 但 Excel 只保留10个“测定样本槽位”。
    #
    # 因此这里统计的是：
    # 已经获得几个测定样本 / 目标10个
    # -------------------------------------------------------------------------

    $speciesMap = @{}

    foreach (
        $entry in
        $script:DataCache.GetEnumerator()
    ) {

        $sampleId =
        [string]$entry.Key

        $seed =
        $entry.Value

        $speciesId =
        Get-SpeciesIdFromSampleId $sampleId

        if (
            -not
            $speciesMap.ContainsKey(
                $speciesId
            )
        ) {

            $speciesMap[$speciesId] =
            [pscustomobject]@{

                SpeciesId         =
                $speciesId

                SpeciesName       =
                $seed.SpeciesName

                TotalCount        =
                0

                GerminatedCount   =
                0

                MissingCoordCount =
                0

                Seeds             =
                New-Object `
                    System.Collections.ArrayList
            }
        }

        $group =
        $speciesMap[$speciesId]

        $group.TotalCount++

        $isGerminated =
        Test-GerminatedValue `
            $seed.Germination

        $coordinate = ''

        if (
            $script:CoordCache.ContainsKey(
                $sampleId
            )
        ) {

            $coordinate =
            $script:CoordCache[
            $sampleId
            ].Coordinate
        }

        if ($isGerminated) {

            $group.GerminatedCount++

            if (
                [string]::IsNullOrWhiteSpace(
                    $coordinate
                )
            ) {

                $group.MissingCoordCount++
            }
        }

        [void]$group.Seeds.Add(

            [pscustomobject]@{

                SampleId    =
                $sampleId

                SeedNo      =
                $seed.SeedNo

                Germination =
                $seed.Germination

                Coordinate  =
                $coordinate
            }
        )
    }


    $allSpecies =
    New-Object `
        System.Collections.ArrayList


    foreach (
        $speciesId in
        $speciesMap.Keys
    ) {

        $group =
        $speciesMap[$speciesId]

        [void]$allSpecies.Add(

            [pscustomobject]@{

                SpeciesId         =
                $speciesId

                SpeciesName       =
                $group.SpeciesName

                TotalCount        =
                $group.TotalCount

                GerminatedCount   =
                $group.GerminatedCount

                RemainingCount    =
                [Math]::Max(
                    0,
                    $group.TotalCount -
                    $group.GerminatedCount
                )

                MissingCoordCount =
                $group.MissingCoordCount

                Seeds             =
                @($group.Seeds)
            }
        )
    }


    # v0.7：
    # 即使前10个根苗长样本已经取满，
    # 物种仍需继续参加发芽率巡检。
    $script:GerminationSpeciesCache =
    @($allSpecies)

    # -------------------------------------------------------------------------
    # 5.4 测定时间计划表：A:N
    # -------------------------------------------------------------------------

    $lastCell = $null

    try {
        $lastCell = $script:PlanSheet.Cells.Item(
            $script:PlanSheet.Rows.Count,
            4
        ).End(-4162)

        $planLastRow = [int]$lastCell.Row
    }
    finally {
        Release-Com $lastCell
    }

    if ($planLastRow -ge 2) {
        $range = $null

        try {
            $range = $script:PlanSheet.Range("A2:N$planLastRow")
            $values = $range.Value2
        }
        finally {
            Release-Com $range
        }

        $lower = $values.GetLowerBound(0)
        $upper = $values.GetUpperBound(0)

        for ($i = $lower; $i -le $upper; $i++) {
            $sampleId = Safe-Text ($values.GetValue($i, 4))

            if ([string]::IsNullOrWhiteSpace($sampleId)) {
                continue
            }

            $excelRow = 2 + ($i - $lower)
            $status = Safe-Text ($values.GetValue($i, 13))

            $item = [pscustomobject]@{
                Row         = $excelRow
                SpeciesId   = Get-SpeciesIdFromSampleId $sampleId
                SpeciesName = Safe-Text ($values.GetValue($i, 2))
                SeedNo      = Safe-Text ($values.GetValue($i, 3))
                SampleId    = $sampleId
                MeasureDate = $values.GetValue($i, 12)
                Status      = $status
                Note        = Safe-Text ($values.GetValue($i, 14))
            }

            $script:PlanCache[$sampleId] = $item

            if ($status -like '*今天测*') {
                $stage = $status.Replace('今天测', '')

                [void]$todayTasks.Add(
                    [pscustomobject]@{
                        SpeciesId   = $item.SpeciesId
                        SpeciesName = $item.SpeciesName
                        SeedNo      = $item.SeedNo
                        SampleId    = $sampleId
                        MeasureDate = ExcelDate-ToText $item.MeasureDate
                        Stage       = $stage
                        Note        = $item.Note
                    }
                )
            }
        }
    }

    $script:TodayTaskCache = @($todayTasks)

    Perf-Log "根苗长缓存：$($script:DataCache.Count)"
    Perf-Log "发芽待查物种：$($script:GerminationSpeciesCache.Count)"
    Perf-Log "今日任务缓存：$($script:TodayTaskCache.Count)"
}


# =============================================================================
# 06. 样本查询与业务数据操作
# =============================================================================

function Get-SampleInfo([string]$SampleId) {
    if ($null -eq $script:Book) {
        throw '尚未连接 Excel。'
    }

    $target = $SampleId.Trim()

    if (-not $script:DataCache.ContainsKey($target)) {
        throw "未找到样本ID：$target"
    }

    $data = $script:DataCache[$target]
    $status = ''

    if ($script:PlanCache.ContainsKey($target)) {
        $status = $script:PlanCache[$target].Status
    }

    $planRow = $null

    if ($script:PlanCache.ContainsKey($target)) {
        $planRow = $script:PlanCache[$target].Row
    }

    return [pscustomobject]@{
        DataRow     = $data.Row
        PlanRow     = $planRow
        SpeciesId   = $data.SpeciesId
        SpeciesName = $data.SpeciesName
        SeedNo      = $data.SeedNo
        Germination = ExcelDate-ToText $data.Germination
        Status      = $status
    }
}

function Save-Germination([string]$SampleId, [DateTime]$DateValue) {
    $info = Get-SampleInfo $SampleId

    Set-CellValue `
        $script:DataSheet `
        $info.DataRow `
        6 `
    ([double]$DateValue.ToOADate())

    # 设置发芽日期显示格式
    $cell = $null

    try {
        $cell = $script:DataSheet.Cells.Item([int]$info.DataRow, 6)
        $cell.NumberFormat = 'yyyy/m/d'
    }
    finally {
        Release-Com $cell
    }

    $script:PlanSheet.Calculate()
    $script:Book.Save()

    if (-not $script:Book.Saved) {
        throw 'Excel 报告工作簿尚未成功保存。'
    }

    Rebuild-Cache
}


function Normalize-GerminationCoordinate(
    [string]$Coordinate
) {

    $coord = $Coordinate.Trim().ToUpperInvariant()

    if (
        $coord -notmatch
        '^[A-E](10|[1-9])$'
    ) {

        throw (
            '种子坐标必须为 A1～E10，' +
            '例如 E5。'
        )
    }

    return $coord
}


function Split-GerminationCoordinates(
    [string]$Text
) {

    # 支持：
    #
    # E5 C7 A2
    # E5,C7,A2
    # E5，C7，A2

    $parts =
    @(
        $Text -split
        '[,\s，;；]+'
    )

    $result =
    New-Object `
        System.Collections.ArrayList

    $seen = @{}

    foreach ($part in $parts) {

        if (
            [string]::IsNullOrWhiteSpace(
                $part
            )
        ) {
            continue
        }

        $coord =
        Normalize-GerminationCoordinate `
            $part

        if ($seen.ContainsKey($coord)) {

            throw (
                '本次输入中坐标重复：' +
                $coord
            )
        }

        $seen[$coord] = $true

        [void]$result.Add($coord)
    }

    if ($result.Count -eq 0) {

        throw (
            '请输入今天新发芽种子的坐标，' +
            '例如 E5 C7。'
        )
    }

    return @($result)
}


function Get-UsedCoordinateMap(
    [string]$SpeciesId
) {

    # 坐标只要求同一个培养皿内部唯一。
    #
    # 003可以有E5，
    # 010也可以有E5。

    $used = @{}

    foreach (
        $entry in
        $script:CoordCache.GetEnumerator()
    ) {

        $record =
        $entry.Value

        if (
            $record.SpeciesId -ne
            $SpeciesId
        ) {
            continue
        }

        if (
            [string]::IsNullOrWhiteSpace(
                $record.Coordinate
            )
        ) {
            continue
        }

        $used[
        $record.Coordinate
        ] =
        [string]$entry.Key
    }

    return $used
}


function Save-CoordinateBackfill(
    [string]$SpeciesId,
    [hashtable]$Assignments
) {

    if (
        $null -eq $Assignments -or
        $Assignments.Count -eq 0
    ) {

        throw '没有需要补录的坐标。'
    }


    # 当前培养皿已经使用的物理位置
    $used =
    Get-UsedCoordinateMap `
        $SpeciesId


    # -------------------------------------------------------------------------
    # 先完整检查
    # -------------------------------------------------------------------------

    foreach (
        $sampleId in
        $Assignments.Keys
    ) {

        if (
            -not
            $script:DataCache.ContainsKey(
                $sampleId
            )
        ) {

            throw (
                '未找到样本：' +
                $sampleId
            )
        }


        $sampleSpeciesId =
        Get-SpeciesIdFromSampleId `
            $sampleId


        if (
            $sampleSpeciesId -ne
            $SpeciesId
        ) {

            throw (
                '样本不属于当前物种：' +
                $sampleId
            )
        }


        $sample =
        $script:DataCache[$sampleId]


        $isGerminated =
        Test-GerminatedValue `
            $sample.Germination


        if (-not $isGerminated) {

            throw (
                $sampleId +
                ' 尚未发芽，不能补录坐标。'
            )
        }


        $coord =
        Normalize-GerminationCoordinate `
            $Assignments[$sampleId]


        # 已经记录坐标则禁止覆盖
        if (
            $script:CoordCache.ContainsKey(
                $sampleId
            )
        ) {

            $oldCoord =
            $script:CoordCache[
            $sampleId
            ].Coordinate


            if (
                -not
                [string]::IsNullOrWhiteSpace(
                    $oldCoord
                )
            ) {

                throw (
                    $sampleId +
                    ' 已记录原始坐标 ' +
                    $oldCoord +
                    '，不允许覆盖。'
                )
            }
        }


        # 同一个培养皿内部不能重复使用同一个物理位置
        if ($used.ContainsKey($coord)) {

            throw (
                '当前培养皿中的坐标 ' +
                $coord +
                ' 已经分配给 ' +
                $used[$coord]
            )
        }


        $used[$coord] =
        $sampleId
    }


    # -------------------------------------------------------------------------
    # 正式写入测定时间计划表 N列备注
    # -------------------------------------------------------------------------

    foreach (
        $sampleId in
        $Assignments.Keys
    ) {

        $coord =
        Normalize-GerminationCoordinate `
            $Assignments[$sampleId]


        if (
            -not
            $script:PlanCache.ContainsKey(
                $sampleId
            )
        ) {

            throw (
                '测定计划表中未找到样本：' +
                $sampleId
            )
        }


        $planRow =
        [int]$script:PlanCache[
        $sampleId
        ].Row


        Set-CellValue `
            $script:PlanSheet `
            $planRow `
            14 `
            $coord
    }


    $script:Book.Save()


    if (-not $script:Book.Saved) {

        throw '原始坐标没有成功保存到Excel。'
    }


    Rebuild-Cache
}

function Get-NextGerminationRecordId {

    $maxNumber = 0

    foreach (
        $record in
        @($script:GerminationLogCache)
    ) {

        $recordId =
        Safe-Text $record.RecordId

        if (
            $recordId -match
            '^G(\d+)$'
        ) {

            $number = 0

            if (
                [int]::TryParse(
                    $Matches[1],
                    [ref]$number
                )
            ) {

                if ($number -gt $maxNumber) {
                    $maxNumber = $number
                }
            }
        }
    }

    return (
        'G{0:D6}' -f
        ($maxNumber + 1)
    )
}

function Get-GerminationSpeciesItem(
    [string]$SpeciesId
) {

    foreach (
        $item in
        @($script:GerminationSpeciesCache)
    ) {

        if (
            $item.SpeciesId -eq
            $SpeciesId
        ) {

            return $item
        }
    }

    throw (
        '发芽巡检中未找到物种：' +
        $SpeciesId
    )
}

function Get-SpeciesPlacedDate(
    [string]$SpeciesId,
    [string]$Replicate
) {

    $key =
    "$SpeciesId|$Replicate"


    # -------------------------------------------------------------------------
    # 第一优先级：
    # 已有发芽历史记录中的置床日期。
    # -------------------------------------------------------------------------

    if (
        $script:GerminationStatusCache.ContainsKey(
            $key
        )
    ) {

        $status =
        $script:GerminationStatusCache[
        $key
        ]


        if ($null -ne $status.PlacedDate) {

            return (
                [DateTime]$status.PlacedDate
            ).Date
        }
    }


    # -------------------------------------------------------------------------
    # 第二优先级：
    # 原有“根-苗长统计表”中的置床日期。
    #
    # 同一物种的测定样本共享一个置床日期，
    # 因此只要找到一个有效值即可。
    # -------------------------------------------------------------------------

    foreach (
        $seed in
        $script:DataCache.Values
    ) {

        if (
            $seed.SpeciesId -ne
            $SpeciesId
        ) {
            continue
        }


        $candidate =
        ExcelDate-ToDateTime `
            $seed.PlacedDate


        if ($null -ne $candidate) {

            return $candidate.Date
        }
    }


    return $null
}
function Save-GerminationInspection(
    [string]$SpeciesId,
    [int]$NewGerminated,
    [string[]]$Coordinates,
    [DateTime]$PlacedDateValue,
    [DateTime]$DateValue
) {

    if ($null -eq $script:Book) {
        throw '尚未连接 Excel。'
    }


    $species =
    Get-GerminationSpeciesItem `
        $SpeciesId


    $replicate =
    [string]$script:ExperimentSettings.DefaultReplicate


    $statusKey =
    "$SpeciesId|$replicate"


    if (
        -not
        $script:GerminationStatusCache.ContainsKey(
            $statusKey
        )
    ) {

        throw (
            '未找到当前培养皿的发芽状态：' +
            $statusKey
        )
    }


    $status =
    $script:GerminationStatusCache[
    $statusKey
    ]


    $totalSeeds =
    [int]$status.TotalSeeds


    $currentGerminated =
    [int]$status.CumulativeGerminated


    # -------------------------------------------------------------------------
    # 1. 校验新增发芽数
    # -------------------------------------------------------------------------

    if ($NewGerminated -lt 0) {

        throw (
            '本次新增发芽数不能小于 0。'
        )
    }


    $newCumulative =
    $currentGerminated +
    $NewGerminated


    if (
        $newCumulative -gt
        $totalSeeds
    ) {

        throw (
            '保存后累计发芽数将达到 ' +
            $newCumulative +
            ' 粒，超过总种子数 ' +
            $totalSeeds +
            ' 粒。'
        )
    }


    # -------------------------------------------------------------------------
    # 2. 找出尚未分配的根苗长测定样本槽位
    # -------------------------------------------------------------------------

    $blankSlots =
    @(
        $species.Seeds |
        Where-Object {

            $isGerminated =
            Test-GerminatedValue `
                $_.Germination

            -not $isGerminated

        } |
        Sort-Object {

            $n = 999

            [void][int]::TryParse(
                [string]$_.SeedNo,
                [ref]$n
            )

            $n
        }
    )


    # 本次需要推进的测定样本数量。
    # 坐标现在是可选信息，不再决定是否能够分配测定样本。
    $samplesToAssignCount =
    [Math]::Min(
        [int]$NewGerminated,
        [int]$blankSlots.Count
    )


    # -------------------------------------------------------------------------
    # 3. 规范并检查坐标
    # -------------------------------------------------------------------------

    $coordList =
    New-Object `
        System.Collections.ArrayList


    foreach (
        $coordRaw in
        @($Coordinates)
    ) {

        if (
            [string]::IsNullOrWhiteSpace(
                [string]$coordRaw
            )
        ) {
            continue
        }


        $coord =
        Normalize-GerminationCoordinate `
        ([string]$coordRaw)


        [void]$coordList.Add(
            $coord
        )
    }


    # 坐标允许完全不填写，也允许只填写部分。
    # 但不能比本次实际分配的测定样本更多。
    if (
        $coordList.Count -gt
        $samplesToAssignCount
    ) {

        throw (
            '本次最多只能填写 ' +
            $samplesToAssignCount +
            ' 个坐标，当前填写了 ' +
            $coordList.Count +
            ' 个。'
        )
    }


    $used =
    Get-UsedCoordinateMap `
        $SpeciesId


    $newUsed = @{}


    foreach ($coord in $coordList) {

        if ($used.ContainsKey($coord)) {

            throw (
                '当前培养皿中的坐标 ' +
                $coord +
                ' 已经用于 ' +
                $used[$coord]
            )
        }


        if ($newUsed.ContainsKey($coord)) {

            throw (
                '本次输入坐标重复：' +
                $coord
            )
        }


        $newUsed[$coord] =
        $true
    }


    # -------------------------------------------------------------------------
    # 4. 获取培养皿级置床日期
    # -------------------------------------------------------------------------

    $placedDate =
    Get-SpeciesPlacedDate `
        $SpeciesId `
        $replicate


    # 已有历史数据时，以历史置床日期为准。
    # 第一次巡检时，则采用界面中设置的置床日期。
    if ($null -eq $placedDate) {

        $placedDate =
        $PlacedDateValue.Date
    }
    else {

        $placedDate =
        ([DateTime]$placedDate).Date
    }


    $inspectionDate =
    $DateValue.Date

    if (
        $placedDate.Date -gt
        (Get-Date).Date
    ) {

        throw (
            '置床日期不能晚于今天。'
        )
    }

    if (
        $inspectionDate -lt
        $placedDate.Date
    ) {

        throw (
            '巡检日期不能早于置床日期。'
        )
    }


    # 日期来自界面；
    # 时间使用实际保存时刻。
    $inspectionTime =
    $inspectionDate.Add(
        (Get-Date).TimeOfDay
    )


    $daysAfterPlacement =
    [int](
        $inspectionDate -
        $placedDate.Date
    ).TotalDays


    $newRate =
    if ($totalSeeds -gt 0) {

        [double]$newCumulative /
        [double]$totalSeeds
    }
    else {

        0.0
    }


    $recordId =
    Get-NextGerminationRecordId


    # -------------------------------------------------------------------------
    # 5. 计算“发芽记录”下一行
    # -------------------------------------------------------------------------

    $lastCell = $null

    try {

        $lastCell =
        $script:GerminationLogSheet.Cells.Item(
            $script:GerminationLogSheet.Rows.Count,
            2
        ).End(-4162)


        $lastRow =
        [int]$lastCell.Row
    }
    finally {

        Release-Com $lastCell
    }


    if ($lastRow -lt 2) {

        $nextRow = 2
    }
    else {

        $nextRow =
        $lastRow + 1
    }


    # -------------------------------------------------------------------------
    # 6. 正式写入前10个测定样本
    # -------------------------------------------------------------------------

    $oaGerminationDate =
    [double]$inspectionDate.ToOADate()


    $assignments =
    New-Object `
        System.Collections.ArrayList


    for (
        $i = 0;
        $i -lt $samplesToAssignCount;
        $i++
    ) {

        $slot =
        $blankSlots[$i]


        $sampleId =
        [string]$slot.SampleId


        # 坐标现在是可选项。
        # 如果用户没有填写，则保留为空。
        $coord = ''

        if ($i -lt $coordList.Count) {

            $coord =
            [string]$coordList[$i]
        }


        # 根-苗长统计表 F列：发芽日期
        $dataRow =
        [int]$script:DataCache[
        $sampleId
        ].Row


        Set-CellValue `
            $script:DataSheet `
            $dataRow `
            6 `
            $oaGerminationDate


        $dateCell = $null

        try {

            $dateCell =
            $script:DataSheet.Cells.Item(
                $dataRow,
                6
            )

            $dateCell.NumberFormat =
            'yyyy/m/d'
        }
        finally {

            Release-Com $dateCell
        }


        # 测定时间计划表 N列：原始坐标
        if (
            -not
            $script:PlanCache.ContainsKey(
                $sampleId
            )
        ) {

            throw (
                '测定时间计划表中未找到样本：' +
                $sampleId
            )
        }


        $planRow =
        [int]$script:PlanCache[
        $sampleId
        ].Row


        # 只有实际填写了坐标时才写入 N 列。
        if (
            -not
            [string]::IsNullOrWhiteSpace(
                $coord
            )
        ) {

            Set-CellValue `
                $script:PlanSheet `
                $planRow `
                14 `
                $coord
        }


        [void]$assignments.Add(

            [pscustomobject]@{

                SampleId   =
                $sampleId

                Coordinate =
                $coord
            }
        )
    }


    # -------------------------------------------------------------------------
    # 7. 写入“发芽记录”A:L
    # -------------------------------------------------------------------------

    # B列强制文本，确保001不会变成1
    $speciesIdCell = $null

    try {

        $speciesIdCell =
        $script:GerminationLogSheet.Cells.Item(
            $nextRow,
            2
        )

        $speciesIdCell.NumberFormat =
        '@'
    }
    finally {

        Release-Com $speciesIdCell
    }


    Set-CellValue `
        $script:GerminationLogSheet `
        $nextRow `
        1 `
        $recordId

    Set-CellValue `
        $script:GerminationLogSheet `
        $nextRow `
        2 `
    ([string]$SpeciesId)

    Set-CellValue `
        $script:GerminationLogSheet `
        $nextRow `
        3 `
    ([string]$species.SpeciesName)

    Set-CellValue `
        $script:GerminationLogSheet `
        $nextRow `
        4 `
        $replicate

    Set-CellValue `
        $script:GerminationLogSheet `
        $nextRow `
        5 `
    ([double]$placedDate.Date.ToOADate())

    Set-CellValue `
        $script:GerminationLogSheet `
        $nextRow `
        6 `
    ([double]$inspectionTime.ToOADate())

    Set-CellValue `
        $script:GerminationLogSheet `
        $nextRow `
        7 `
        $daysAfterPlacement

    Set-CellValue `
        $script:GerminationLogSheet `
        $nextRow `
        8 `
        $NewGerminated

    Set-CellValue `
        $script:GerminationLogSheet `
        $nextRow `
        9 `
        $newCumulative

    Set-CellValue `
        $script:GerminationLogSheet `
        $nextRow `
        10 `
        $totalSeeds

    Set-CellValue `
        $script:GerminationLogSheet `
        $nextRow `
        11 `
        $newRate

    Set-CellValue `
        $script:GerminationLogSheet `
        $nextRow `
        12 `
        $null


    # 日期与百分比显示格式
    $formatCell = $null

    try {

        $formatCell =
        $script:GerminationLogSheet.Cells.Item(
            $nextRow,
            5
        )

        $formatCell.NumberFormat =
        'yyyy/m/d'
    }
    finally {

        Release-Com $formatCell
    }


    $formatCell = $null

    try {

        $formatCell =
        $script:GerminationLogSheet.Cells.Item(
            $nextRow,
            6
        )

        $formatCell.NumberFormat =
        'yyyy/m/d h:mm'
    }
    finally {

        Release-Com $formatCell
    }


    $formatCell = $null

    try {

        $formatCell =
        $script:GerminationLogSheet.Cells.Item(
            $nextRow,
            11
        )

        $formatCell.NumberFormat =
        '0.00%'
    }
    finally {

        Release-Com $formatCell
    }


    # -------------------------------------------------------------------------
    # 8. 统一计算、保存、重建缓存
    # -------------------------------------------------------------------------

    $script:PlanSheet.Calculate()

    $script:Book.Save()


    if (-not $script:Book.Saved) {

        throw (
            '本次发芽巡检数据没有成功保存到 Excel。'
        )
    }


    Rebuild-Cache


    return [pscustomobject]@{

        RecordId             =
        $recordId

        NewGerminated        =
        $NewGerminated

        CumulativeGerminated =
        $newCumulative

        TotalSeeds           =
        $totalSeeds

        GerminationRate      =
        $newRate

        AssignedSamples      =
        @($assignments)
    }
}

function Parse-Measure([string]$Text) {
    # 根/苗长输入统一校验：
    # - 不能为空
    # - 允许标准数字（含小数）
    # - 保留 NA 作为“已检查但无有效测量值”
    # - 不允许负数
    $s = $Text.Trim()

    if ($s -eq '') {
        throw '根长和苗长不能为空。'
    }

    if ($s.ToUpperInvariant() -eq 'NA') {
        return 'NA'
    }

    $num = 0.0

    # 优先按点号小数解析；失败后再按本机区域设置解析。
    if (
        -not [double]::TryParse(
            $s,
            [Globalization.NumberStyles]::Float,
            [Globalization.CultureInfo]::InvariantCulture,
            [ref]$num
        )
    ) {
        if (-not [double]::TryParse($s, [ref]$num)) {
            throw "无法识别数值：$s"
        }
    }

    if ($num -lt 0) {
        throw '测定值不能为负数。'
    }

    return $num
}

function Has-Value($Value) {
    if ($null -eq $Value) {
        return $false
    }

    if ([string]::IsNullOrWhiteSpace([string]$Value)) {
        return $false
    }

    return $true
}

function Get-ExistingMeasurement([string]$SampleId, [int]$Stage) {
    $info = Get-SampleInfo $SampleId
    $row = $info.DataRow

    switch ($Stage) {
        3 {
            $rootCol = 7
            $shootCol = 10
        }
        7 {
            $rootCol = 8
            $shootCol = 11
        }
        14 {
            $rootCol = 9
            $shootCol = 12
        }
        default {
            throw "不支持的测定阶段：$Stage"
        }
    }

    return [pscustomobject]@{
        Root  = Get-CellValue $script:DataSheet $row $rootCol
        Shoot = Get-CellValue $script:DataSheet $row $shootCol
    }
}

function Save-Measurement(
    [string]$SampleId,
    [int]$Stage,
    [string]$RootText,
    [string]$ShootText
) {
    Write-Log "保存根苗长：样本=$SampleId 阶段=$Stage 根长=$RootText 苗长=$ShootText"

    $info = Get-SampleInfo $SampleId
    $root = Parse-Measure $RootText
    $shoot = Parse-Measure $ShootText

    switch ($Stage) {
        3 {
            $rootCol = 7
            $shootCol = 10
        }
        7 {
            $rootCol = 8
            $shootCol = 11
        }
        14 {
            $rootCol = 9
            $shootCol = 12
        }
        default {
            throw '测定阶段只能是 3、7 或 14 DAG。'
        }
    }

    Set-CellValue $script:DataSheet ([int]$info.DataRow) ([int]$rootCol)  $root
    Set-CellValue $script:DataSheet ([int]$info.DataRow) ([int]$shootCol) $shoot

    # 只计算计划表，再保存到磁盘
    $script:PlanSheet.Calculate()
    $script:Book.Save()

    if (-not $script:Book.Saved) {
        throw 'Excel 报告工作簿尚未成功保存。'
    }

    Rebuild-Cache

    Write-Log '保存根苗长完成'
}


# =============================================================================
# 07. UI主题与通用样式
# =============================================================================
# 这一节只负责视觉表现，不承担任何实验业务逻辑。
# 后续若想换主色、字体、表格风格，优先修改这里即可。
#
# 设计原则：
#   - 主背景使用很浅的灰蓝色，减少大面积纯白造成的刺眼感。
#   - 内容卡片使用白色，形成清晰的信息层级。
#   - 绿色只用于“主要动作/成功状态”，蓝色用于选择，高风险信息用橙/红色。
#   - 表格统一表头、行高、交替行底色和选中颜色。
# =============================================================================

$script:UiPalette = @{
    MainBg        = [Drawing.Color]::FromArgb(246, 248, 251)
    Surface       = [Drawing.Color]::White
    Border        = [Drawing.Color]::FromArgb(220, 225, 232)
    TextPrimary   = [Drawing.Color]::FromArgb(35, 40, 47)
    TextSecondary = [Drawing.Color]::FromArgb(102, 110, 120)

    Primary       = [Drawing.Color]::FromArgb(45, 143, 91)
    PrimarySoft   = [Drawing.Color]::FromArgb(232, 246, 238)

    Blue          = [Drawing.Color]::FromArgb(44, 111, 227)
    BlueSoft      = [Drawing.Color]::FromArgb(234, 241, 253)

    Warning       = [Drawing.Color]::FromArgb(205, 126, 23)
    WarningSoft   = [Drawing.Color]::FromArgb(255, 247, 230)

    Danger        = [Drawing.Color]::FromArgb(200, 61, 73)
    DangerSoft    = [Drawing.Color]::FromArgb(253, 238, 240)

    Success       = [Drawing.Color]::FromArgb(45, 143, 91)
    GridHeader    = [Drawing.Color]::FromArgb(244, 247, 250)
    GridAltRow    = [Drawing.Color]::FromArgb(250, 251, 253)
    GridLine      = [Drawing.Color]::FromArgb(226, 230, 235)

    Dag3Soft      = [Drawing.Color]::FromArgb(255, 248, 222)
    Dag7Soft      = [Drawing.Color]::FromArgb(234, 243, 253)
    Dag14Soft     = [Drawing.Color]::FromArgb(232, 246, 238)
}

$script:UiFont = @{
    AppTitle    = (New-Object Drawing.Font('Microsoft YaHei UI', 19, [Drawing.FontStyle]::Bold))
    PageTitle   = (New-Object Drawing.Font('Microsoft YaHei UI', 16, [Drawing.FontStyle]::Bold))
    CardTitle   = (New-Object Drawing.Font('Microsoft YaHei UI', 17, [Drawing.FontStyle]::Bold))
    Body        = (New-Object Drawing.Font('Microsoft YaHei UI', 10.5, [Drawing.FontStyle]::Regular))
    BodyBold    = (New-Object Drawing.Font('Microsoft YaHei UI', 10.5, [Drawing.FontStyle]::Bold))
    Small       = (New-Object Drawing.Font('Microsoft YaHei UI', 9.5, [Drawing.FontStyle]::Regular))
    Input       = (New-Object Drawing.Font('Microsoft YaHei UI', 11, [Drawing.FontStyle]::Regular))
    InputStrong = (New-Object Drawing.Font('Microsoft YaHei UI', 12, [Drawing.FontStyle]::Bold))
}

function Set-UiPrimaryButton($Button) {
    $Button.UseVisualStyleBackColor = $false
    $Button.FlatStyle = [Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.BackColor = $script:UiPalette.Primary
    $Button.ForeColor = [Drawing.Color]::White
    $Button.Font = $script:UiFont.BodyBold
    $Button.Cursor = [Windows.Forms.Cursors]::Hand
}

function Set-UiSecondaryButton($Button) {
    $Button.UseVisualStyleBackColor = $false
    $Button.FlatStyle = [Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 1
    $Button.FlatAppearance.BorderColor = $script:UiPalette.Border
    $Button.BackColor = $script:UiPalette.Surface
    $Button.ForeColor = $script:UiPalette.TextPrimary
    $Button.Font = $script:UiFont.Body
    $Button.Cursor = [Windows.Forms.Cursors]::Hand
}

function Set-UiInput($Control) {
    $Control.BackColor = $script:UiPalette.Surface
    $Control.ForeColor = $script:UiPalette.TextPrimary
    $Control.Font = $script:UiFont.Input
}

function Set-UiGrid($Grid) {
    $Grid.EnableHeadersVisualStyles = $false
    $Grid.BackgroundColor = $script:UiPalette.Surface
    $Grid.BorderStyle = [Windows.Forms.BorderStyle]::FixedSingle
    $Grid.GridColor = $script:UiPalette.GridLine

    $Grid.ColumnHeadersDefaultCellStyle.BackColor = $script:UiPalette.GridHeader
    $Grid.ColumnHeadersDefaultCellStyle.ForeColor = $script:UiPalette.TextPrimary
    $Grid.ColumnHeadersDefaultCellStyle.Font = $script:UiFont.BodyBold
    $Grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $script:UiPalette.GridHeader
    $Grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = $script:UiPalette.TextPrimary
    $Grid.ColumnHeadersDefaultCellStyle.Alignment =
    [Windows.Forms.DataGridViewContentAlignment]::MiddleCenter

    $Grid.DefaultCellStyle.BackColor = $script:UiPalette.Surface
    $Grid.DefaultCellStyle.ForeColor = $script:UiPalette.TextPrimary
    $Grid.DefaultCellStyle.Font = $script:UiFont.Body
    $Grid.DefaultCellStyle.SelectionBackColor = $script:UiPalette.Blue
    $Grid.DefaultCellStyle.SelectionForeColor = [Drawing.Color]::White

    $Grid.AlternatingRowsDefaultCellStyle.BackColor = $script:UiPalette.GridAltRow
    $Grid.RowTemplate.Height = 34
    $Grid.ColumnHeadersHeight = 40
}


# =============================================================================
# 08. 创建主窗口与顶部区域
# =============================================================================

$form = New-Object System.Windows.Forms.Form

# 可选应用图标：图标缺失时不会影响软件主体功能
$AppIconPath = Join-Path $AppDir 'assets\app_icon.ico'
if (Test-Path -LiteralPath $AppIconPath) {
    try {
        $form.Icon = New-Object System.Drawing.Icon($AppIconPath)
    }
    catch {
        Write-Log ('加载应用图标失败：' + $_.Exception.Message)
    }
}
$form.Text = '草种测定管理 v0.7.0'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1380, 840)
$form.MinimumSize = New-Object System.Drawing.Size(1100, 700)
$form.Font = $script:UiFont.Body
$form.BackColor = $script:UiPalette.MainBg
$form.AutoScaleMode = [Windows.Forms.AutoScaleMode]::Dpi
$form.KeyPreview = $true  # 允许主窗口捕获发芽日期控件上的 Enter

# -----------------------------------------------------------------------------
# 主布局：使用两行 TableLayoutPanel，彻底避免“顶栏覆盖标签页”
#   第1行：顶部栏（固定高度）
#   第2行：标签页（自动占满剩余空间）
# -----------------------------------------------------------------------------

$rootLayout = New-Object System.Windows.Forms.TableLayoutPanel
$rootLayout.Dock = 'Fill'
$rootLayout.RowCount = 2
$rootLayout.ColumnCount = 1
$rootLayout.Margin = New-Object Windows.Forms.Padding(0)
$rootLayout.Padding = New-Object Windows.Forms.Padding(0)

$rootRowTop = New-Object System.Windows.Forms.RowStyle
$rootRowTop.SizeType = [Windows.Forms.SizeType]::Absolute
$rootRowTop.Height = 72
[void]$rootLayout.RowStyles.Add($rootRowTop)

$rootRowMain = New-Object System.Windows.Forms.RowStyle
$rootRowMain.SizeType = [Windows.Forms.SizeType]::Percent
$rootRowMain.Height = 100
[void]$rootLayout.RowStyles.Add($rootRowMain)

$form.Controls.Add($rootLayout)

# -----------------------------------------------------------------------------
# 顶栏：三列布局
#   左：软件标题
#   中：Excel连接状态与文件名
#   右：选择Excel + 刷新
# -----------------------------------------------------------------------------

$top = New-Object System.Windows.Forms.Panel
$top.Dock = 'Fill'
$top.BackColor = [Drawing.Color]::White
$top.Margin = New-Object Windows.Forms.Padding(0)
$rootLayout.Controls.Add($top, 0, 0)

$topLayout = New-Object System.Windows.Forms.TableLayoutPanel
$topLayout.Dock = 'Fill'
$topLayout.RowCount = 1
$topLayout.ColumnCount = 3
$topLayout.Margin = New-Object Windows.Forms.Padding(0)
$topLayout.Padding = New-Object Windows.Forms.Padding(0)

$topColLeft = New-Object System.Windows.Forms.ColumnStyle
$topColLeft.SizeType = [Windows.Forms.SizeType]::Absolute
$topColLeft.Width = 320
[void]$topLayout.ColumnStyles.Add($topColLeft)

$topColMiddle = New-Object System.Windows.Forms.ColumnStyle
$topColMiddle.SizeType = [Windows.Forms.SizeType]::Percent
$topColMiddle.Width = 100
[void]$topLayout.ColumnStyles.Add($topColMiddle)

$topColRight = New-Object System.Windows.Forms.ColumnStyle
$topColRight.SizeType = [Windows.Forms.SizeType]::Absolute
$topColRight.Width = 245
[void]$topLayout.ColumnStyles.Add($topColRight)

$top.Controls.Add($topLayout)

# 左：标题
$title = New-Object System.Windows.Forms.Label
$title.Text = '草种测定管理'
$title.Dock = 'Fill'
$title.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
$title.Padding = New-Object Windows.Forms.Padding(20, 0, 0, 0)
$title.Font = $script:UiFont.AppTitle
$title.ForeColor = $script:UiPalette.TextPrimary
$topLayout.Controls.Add($title, 0, 0)

# 中：连接状态（居中显示）
$conn = New-Object System.Windows.Forms.Label
$conn.Text = 'Excel：未连接'
$conn.Dock = 'Fill'
$conn.TextAlign = [Drawing.ContentAlignment]::MiddleCenter
$conn.AutoEllipsis = $true
$conn.ForeColor = $script:UiPalette.TextSecondary
$conn.Padding = New-Object Windows.Forms.Padding(10, 0, 10, 0)
$conn.Font = $script:UiFont.BodyBold
$topLayout.Controls.Add($conn, 1, 0)

# 右：按钮区
$topActions = New-Object System.Windows.Forms.Panel
$topActions.Dock = 'Fill'
$topActions.BackColor = [Drawing.Color]::White
$topLayout.Controls.Add($topActions, 2, 0)

$choose = New-Object System.Windows.Forms.Button
$choose.Text = '选择 Excel'
$choose.Size = New-Object Drawing.Size(115, 36)
$choose.Location = New-Object Drawing.Point(5, 18)
$topActions.Controls.Add($choose)

$refresh = New-Object System.Windows.Forms.Button
$refresh.Text = '刷新'
$refresh.Size = New-Object Drawing.Size(90, 36)
$refresh.Location = New-Object Drawing.Point(130, 18)
$topActions.Controls.Add($refresh)

Set-UiSecondaryButton $choose
Set-UiSecondaryButton $refresh

# 主标签页：永远位于顶栏下方，不再被覆盖
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$tabs.Padding = New-Object Drawing.Point(18, 8)
$tabs.Margin = New-Object Windows.Forms.Padding(0)
$tabs.Font = $script:UiFont.Body
$tabs.BackColor = $script:UiPalette.MainBg
$rootLayout.Controls.Add($tabs, 0, 1)


# =============================================================================
# 09. UI：今日任务页
# =============================================================================

$tabToday = New-Object System.Windows.Forms.TabPage
$tabToday.Text = '今日任务'
$tabToday.BackColor = $script:UiPalette.MainBg
$tabs.TabPages.Add($tabToday)

# -----------------------------------------------------------------------------
# 今日任务页使用两行布局：
#   第1行：日期、统计、搜索筛选
#   第2行：DataGridView
#
# 这样 DataGridView 的表头和滚动条不会再被顶部信息区遮挡。
# -----------------------------------------------------------------------------

$todayLayout = New-Object System.Windows.Forms.TableLayoutPanel
$todayLayout.Dock = 'Fill'
$todayLayout.RowCount = 2
$todayLayout.ColumnCount = 1
$todayLayout.Margin = New-Object Windows.Forms.Padding(0)
$todayLayout.Padding = New-Object Windows.Forms.Padding(0)
$tabToday.Controls.Add($todayLayout)

$todayRowTop = New-Object System.Windows.Forms.RowStyle
$todayRowTop.SizeType = [Windows.Forms.SizeType]::Absolute
$todayRowTop.Height = 150
[void]$todayLayout.RowStyles.Add($todayRowTop)

$todayRowGrid = New-Object System.Windows.Forms.RowStyle
$todayRowGrid.SizeType = [Windows.Forms.SizeType]::Percent
$todayRowGrid.Height = 100
[void]$todayLayout.RowStyles.Add($todayRowGrid)

# 顶部信息与筛选区域
$todayTop = New-Object System.Windows.Forms.Panel
$todayTop.Dock = 'Fill'
$todayTop.BackColor = $script:UiPalette.MainBg
$todayTop.Margin = New-Object Windows.Forms.Padding(0)
$todayLayout.Controls.Add($todayTop, 0, 0)

$todayHeader = New-Object System.Windows.Forms.Label
$todayHeader.Text = (Get-Date -Format 'yyyy/M/d') + ' 今日任务'
$todayHeader.Font = $script:UiFont.PageTitle
$todayHeader.ForeColor = $script:UiPalette.TextPrimary
$todayHeader.AutoSize = $true
$todayHeader.Location = New-Object Drawing.Point(22, 15)
$todayTop.Controls.Add($todayHeader)

$stats = New-Object System.Windows.Forms.Label
$stats.Text = '今日任务 0   |   3DAG 0   |   7DAG 0   |   14DAG 0'
$stats.AutoSize = $true
$stats.Location = New-Object Drawing.Point(24, 58)
$stats.Font = $script:UiFont.BodyBold
$stats.ForeColor = $script:UiPalette.TextSecondary
$todayTop.Controls.Add($stats)

$taskSearchLabel = New-Object Windows.Forms.Label
$taskSearchLabel.Text = '快速查找'
$taskSearchLabel.Location = New-Object Drawing.Point(24, 108)
$taskSearchLabel.AutoSize = $true
$todayTop.Controls.Add($taskSearchLabel)

$taskSearch = New-Object Windows.Forms.TextBox
$taskSearch.Location = New-Object Drawing.Point(100, 101)
$taskSearch.Size = New-Object Drawing.Size(280, 32)
$taskSearch.Font = $script:UiFont.Input
$todayTop.Controls.Add($taskSearch)

$stageFilter = New-Object Windows.Forms.ComboBox
$stageFilter.Location = New-Object Drawing.Point(400, 101)
$stageFilter.Size = New-Object Drawing.Size(140, 32)
$stageFilter.DropDownStyle = 'DropDownList'
[void]$stageFilter.Items.AddRange(@('全部', '3DAG', '7DAG', '14DAG'))
$stageFilter.SelectedIndex = 0
$todayTop.Controls.Add($stageFilter)

$clearSearch = New-Object Windows.Forms.Button
$clearSearch.Text = '清除'
$clearSearch.Location = New-Object Drawing.Point(558, 99)
$clearSearch.Size = New-Object Drawing.Size(80, 34)
$todayTop.Controls.Add($clearSearch)

Set-UiInput $taskSearch
Set-UiInput $stageFilter
Set-UiSecondaryButton $clearSearch

$taskCount = New-Object Windows.Forms.Label
$taskCount.Text = ''
$taskCount.Location = New-Object Drawing.Point(660, 108)
$taskCount.AutoSize = $true
$taskCount.ForeColor = $script:UiPalette.TextSecondary
$todayTop.Controls.Add($taskCount)

# -----------------------------------------------------------------------------
# 今日任务表格
#   Dock=Fill：自动适应窗口
#   ScrollBars=Both：任务过多/窗口过小时自动出现滚动条
# -----------------------------------------------------------------------------

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Fill'
$grid.Margin = New-Object Windows.Forms.Padding(0)
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.AllowUserToOrderColumns = $false
$grid.MultiSelect = $false
$grid.SelectionMode = [Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grid.RowHeadersVisible = $false
$grid.ScrollBars = [Windows.Forms.ScrollBars]::Both
$grid.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$grid.AutoSizeRowsMode = [Windows.Forms.DataGridViewAutoSizeRowsMode]::None
$grid.RowTemplate.Height = 34
$grid.ColumnHeadersHeight = 40
$grid.ColumnHeadersHeightSizeMode = [Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$grid.BackgroundColor = [Drawing.Color]::White
$grid.BorderStyle = [Windows.Forms.BorderStyle]::FixedSingle
$grid.DefaultCellStyle.WrapMode = [Windows.Forms.DataGridViewTriState]::False
$grid.ColumnHeadersDefaultCellStyle.WrapMode = [Windows.Forms.DataGridViewTriState]::False

$grid.DefaultCellStyle.Font = New-Object Drawing.Font(
    'Microsoft YaHei UI',
    10.5
)

$grid.ColumnHeadersDefaultCellStyle.Font = New-Object Drawing.Font(
    'Microsoft YaHei UI',
    10.5,
    [Drawing.FontStyle]::Bold
)

$grid.DefaultCellStyle.Alignment = [Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
$grid.ColumnHeadersDefaultCellStyle.Alignment = [Windows.Forms.DataGridViewContentAlignment]::MiddleCenter

$todayLayout.Controls.Add($grid, 0, 1)

# 列名统一，后续通过 Name 引用
[void]$grid.Columns.Add('speciesId', '物种编号')
[void]$grid.Columns.Add('speciesName', '物种名称')
[void]$grid.Columns.Add('seedNo', '种子编号')
[void]$grid.Columns.Add('sampleId', '样本ID')
[void]$grid.Columns.Add('measureDate', '测定日期')
[void]$grid.Columns.Add('stage', '测定阶段')
[void]$grid.Columns.Add('note', '备注')

# 列宽按字段重要性分配
$grid.Columns['speciesId'].Width = 90
$grid.Columns['speciesName'].Width = 220
$grid.Columns['seedNo'].Width = 85
$grid.Columns['sampleId'].Width = 120
$grid.Columns['measureDate'].Width = 115
$grid.Columns['stage'].Width = 190
$grid.Columns['note'].MinimumWidth = 180
$grid.Columns['note'].AutoSizeMode = [Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

# 结构化字段居中；物种名称、备注左对齐
foreach ($columnName in @('speciesId', 'seedNo', 'sampleId', 'measureDate', 'stage')) {
    $grid.Columns[$columnName].DefaultCellStyle.Alignment =
    [Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
}

Set-UiGrid $grid


# =============================================================================
# 10. UI：发芽巡检页
# =============================================================================

$tabG = New-Object System.Windows.Forms.TabPage
$tabG.Text = '发芽巡检'
$tabG.BackColor = $script:UiPalette.MainBg
$tabs.TabPages.Add($tabG)

# 两行布局：上方统计/搜索；下方左右分栏
$gRootLayout = New-Object Windows.Forms.TableLayoutPanel
$gRootLayout.Dock = 'Fill'
$gRootLayout.RowCount = 2
$gRootLayout.ColumnCount = 1
$gRootLayout.Margin = New-Object Windows.Forms.Padding(0)
$gRootLayout.Padding = New-Object Windows.Forms.Padding(0)
$tabG.Controls.Add($gRootLayout)

$gTopRow = New-Object Windows.Forms.RowStyle
$gTopRow.SizeType = [Windows.Forms.SizeType]::Absolute
$gTopRow.Height = 105
[void]$gRootLayout.RowStyles.Add($gTopRow)

$gMainRow = New-Object Windows.Forms.RowStyle
$gMainRow.SizeType = [Windows.Forms.SizeType]::Percent
$gMainRow.Height = 100
[void]$gRootLayout.RowStyles.Add($gMainRow)

# -------------------------------------------------------------------------
# 顶部统计 + 搜索
# -------------------------------------------------------------------------

$gTop = New-Object Windows.Forms.Panel
$gTop.Dock = 'Fill'
$gTop.BackColor = $script:UiPalette.MainBg
$gRootLayout.Controls.Add($gTop, 0, 0)

$gHeader = New-Object Windows.Forms.Label
$gHeader.Text = '发芽巡检'
$gHeader.Font = $script:UiFont.PageTitle
$gHeader.ForeColor = $script:UiPalette.TextPrimary
$gHeader.AutoSize = $true
$gHeader.Location = New-Object Drawing.Point(22, 14)
$gTop.Controls.Add($gHeader)

$gStats = New-Object Windows.Forms.Label
$gStats.Text = '待检查物种 0   |   待检查种子 0   |   今日新增 0   |   未设坐标 0'
$gStats.AutoSize = $true
$gStats.Location = New-Object Drawing.Point(24, 52)
$gStats.Font = $script:UiFont.BodyBold
$gStats.ForeColor = $script:UiPalette.TextSecondary
$gTop.Controls.Add($gStats)

$gSearchLabel = New-Object Windows.Forms.Label
$gSearchLabel.Text = '查找物种'
$gSearchLabel.AutoSize = $true
$gSearchLabel.Location = New-Object Drawing.Point(24, 82)
$gTop.Controls.Add($gSearchLabel)

$gSearch = New-Object Windows.Forms.TextBox
$gSearch.Location = New-Object Drawing.Point(100, 76)
$gSearch.Size = New-Object Drawing.Size(280, 32)
$gSearch.Font = $script:UiFont.Input
$gTop.Controls.Add($gSearch)

$gClearSearch = New-Object Windows.Forms.Button
$gClearSearch.Text = '清除'
$gClearSearch.Location = New-Object Drawing.Point(395, 74)
$gClearSearch.Size = New-Object Drawing.Size(80, 34)
$gTop.Controls.Add($gClearSearch)

Set-UiInput $gSearch
Set-UiSecondaryButton $gClearSearch

# -------------------------------------------------------------------------
# 主区域：左侧“待检查物种” + 右侧“当前物种10粒种子”
# -------------------------------------------------------------------------

$gSplit = New-Object Windows.Forms.SplitContainer
$gSplit.Dock = 'Fill'
$gSplit.Orientation = [Windows.Forms.Orientation]::Vertical
$gSplit.BackColor = $script:UiPalette.MainBg
$gSplit.SplitterWidth = 8

# 注意：
# 此时控件还没有完成窗口布局，
# 不要在这里设置 Panel1MinSize / Panel2MinSize，
# 否则 Windows PowerShell 5.1 + WinForms 可能因当前宽度不足直接报错。
$gSplit.Panel1MinSize = 100
$gSplit.Panel2MinSize = 100

$gRootLayout.Controls.Add($gSplit, 0, 1)

# -------------------- 左：物种列表 --------------------

$gSpeciesGrid = New-Object Windows.Forms.DataGridView
$gSpeciesGrid.Dock = 'Fill'
$gSpeciesGrid.ReadOnly = $true
$gSpeciesGrid.AllowUserToAddRows = $false
$gSpeciesGrid.AllowUserToDeleteRows = $false
$gSpeciesGrid.AllowUserToResizeRows = $false
$gSpeciesGrid.MultiSelect = $false
$gSpeciesGrid.SelectionMode = [Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$gSpeciesGrid.RowHeadersVisible = $false
$gSpeciesGrid.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$gSpeciesGrid.RowTemplate.Height = 32
$gSpeciesGrid.ColumnHeadersHeight = 38
$gSpeciesGrid.BackgroundColor = [Drawing.Color]::White
$gSpeciesGrid.ScrollBars = [Windows.Forms.ScrollBars]::Both
$gSplit.Panel1.Controls.Add($gSpeciesGrid)

[void]$gSpeciesGrid.Columns.Add(
    'gSpeciesId',
    '物种编号'
)

[void]$gSpeciesGrid.Columns.Add(
    'gSpeciesName',
    '物种名称'
)

[void]$gSpeciesGrid.Columns.Add(
    'gProgress',
    '发芽进度'
)

[void]$gSpeciesGrid.Columns.Add(
    'gRemaining',
    '测定样本'
)

[void]$gSpeciesGrid.Columns.Add(
    'gMissingCoord',
    '未填坐标'
)


$gSpeciesGrid.Columns[
'gSpeciesId'
].Width = 85

$gSpeciesGrid.Columns[
'gSpeciesName'
].Width = 180

$gSpeciesGrid.Columns[
'gProgress'
].Width = 135

$gSpeciesGrid.Columns[
'gRemaining'
].Width = 90

$gSpeciesGrid.Columns[
'gMissingCoord'
].Width = 85

foreach (
    $columnName in
    @(
        'gSpeciesId',
        'gProgress',
        'gRemaining',
        'gMissingCoord'
    )
) {

    $gSpeciesGrid.Columns[$columnName].DefaultCellStyle.Alignment =
    [Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
}

Set-UiGrid $gSpeciesGrid

# -------------------- 右：详情布局 --------------------

$gDetailLayout = New-Object Windows.Forms.TableLayoutPanel
$gDetailLayout.Dock = 'Fill'
$gDetailLayout.BackColor = $script:UiPalette.MainBg
$gDetailLayout.Margin = New-Object Windows.Forms.Padding(0)
$gDetailLayout.RowCount = 3
$gDetailLayout.ColumnCount = 1
$gSplit.Panel2.Controls.Add($gDetailLayout)

$gDetailRow1 = New-Object Windows.Forms.RowStyle
$gDetailRow1.SizeType = [Windows.Forms.SizeType]::Absolute
$gDetailRow1.Height = 255
[void]$gDetailLayout.RowStyles.Add($gDetailRow1)

$gDetailRow2 = New-Object Windows.Forms.RowStyle
$gDetailRow2.SizeType = [Windows.Forms.SizeType]::Percent
$gDetailRow2.Height = 100
[void]$gDetailLayout.RowStyles.Add($gDetailRow2)

$gDetailRow3 = New-Object Windows.Forms.RowStyle
$gDetailRow3.SizeType = [Windows.Forms.SizeType]::Absolute
$gDetailRow3.Height = 110
[void]$gDetailLayout.RowStyles.Add($gDetailRow3)

# 右上：当前物种信息、坐标、日期
$gDetailTop = New-Object Windows.Forms.Panel
$gDetailTop.Dock = 'Fill'
$gDetailTop.BackColor = $script:UiPalette.Surface
$gDetailTop.BorderStyle = [Windows.Forms.BorderStyle]::FixedSingle
$gDetailLayout.Controls.Add($gDetailTop, 0, 0)

$gSelectedTitle = New-Object Windows.Forms.Label
$gSelectedTitle.Text = '请选择左侧待检查物种'
$gSelectedTitle.Location = New-Object Drawing.Point(18, 12)
$gSelectedTitle.Size = New-Object Drawing.Size(620, 32)
$gSelectedTitle.Font = $script:UiFont.CardTitle
$gSelectedTitle.ForeColor = $script:UiPalette.TextPrimary
$gDetailTop.Controls.Add($gSelectedTitle)

$gSelectedStats = New-Object Windows.Forms.Label
$gSelectedStats.Text = ''
$gSelectedStats.Location = New-Object Drawing.Point(20, 48)
$gSelectedStats.Size = New-Object Drawing.Size(600, 28)
$gSelectedStats.Font = $script:UiFont.Body
$gSelectedStats.ForeColor = $script:UiPalette.TextSecondary
$gDetailTop.Controls.Add($gSelectedStats)

$gPlacedDateLabel =
New-Object Windows.Forms.Label

$gPlacedDateLabel.Text =
'置床日期'

$gPlacedDateLabel.Location =
New-Object Drawing.Point(
    20,
    91
)

$gPlacedDateLabel.AutoSize =
$true

$gDetailTop.Controls.Add(
    $gPlacedDateLabel
)


$gPlacedDate =
New-Object Windows.Forms.DateTimePicker

$gPlacedDate.Format =
'Custom'

$gPlacedDate.CustomFormat =
'yyyy/M/d'

$gPlacedDate.Value =
(Get-Date).Date

$gPlacedDate.Location =
New-Object Drawing.Point(
    155,
    84
)

$gPlacedDate.Size =
New-Object Drawing.Size(
    160,
    30
)

$gPlacedDate.Font =
$script:UiFont.Input

$gDetailTop.Controls.Add(
    $gPlacedDate
)


$gPlacedDateHint =
New-Object Windows.Forms.Label

$gPlacedDateHint.Text =
'首次巡检时设置一次'

$gPlacedDateHint.Location =
New-Object Drawing.Point(
    315,
    91
)

$gPlacedDateHint.AutoSize =
$true

$gPlacedDateHint.ForeColor =
$script:UiPalette.TextSecondary

$gPlacedDateHint.Font =
$script:UiFont.Small

$gDetailTop.Controls.Add(
    $gPlacedDateHint
)

$gNewCountLabel =
New-Object Windows.Forms.Label

$gNewCountLabel.Text =
'本次新增发芽'

$gNewCountLabel.Location =
New-Object Drawing.Point(
    20,
    132
)

$gNewCountLabel.AutoSize =
$true

$gDetailTop.Controls.Add(
    $gNewCountLabel
)


$gNewCount =
New-Object Windows.Forms.TextBox

$gNewCount.Location =
New-Object Drawing.Point(
    155,
    125
)

$gNewCount.Size =
New-Object Drawing.Size(
    90,
    30
)

$gNewCount.Font =
$script:UiFont.InputStrong

$gNewCount.TextAlign =
[Windows.Forms.HorizontalAlignment]::Center

$gDetailTop.Controls.Add(
    $gNewCount
)


$gNewCountHint =
New-Object Windows.Forms.Label

$gNewCountHint.Text =
'输入本次新发芽粒数；0 也可保存'

$gNewCountHint.Location =
New-Object Drawing.Point(
    245,
    132
)

$gNewCountHint.AutoSize =
$true

$gNewCountHint.ForeColor =
$script:UiPalette.TextSecondary

$gNewCountHint.Font =
$script:UiFont.Small

$gDetailTop.Controls.Add(
    $gNewCountHint
)

$gNewCoordLabel =
New-Object Windows.Forms.Label

$gNewCoordLabel.Text =
'样本坐标（可选）'

$gNewCoordLabel.Location =
New-Object Drawing.Point(
    20,
    173
)

$gNewCoordLabel.AutoSize =
$true

$gDetailTop.Controls.Add(
    $gNewCoordLabel
)


$gNewCoords =
New-Object Windows.Forms.TextBox

$gNewCoords.Location =
New-Object Drawing.Point(
    155,
    166
)

$gNewCoords.Size =
New-Object Drawing.Size(
    220,
    30
)

$gNewCoords.CharacterCasing =
[Windows.Forms.CharacterCasing]::Upper
$gNewCoords.Font = $script:UiFont.InputStrong
$gNewCoords.BackColor = $script:UiPalette.Surface
$gNewCoords.ForeColor = $script:UiPalette.TextPrimary

$gDetailTop.Controls.Add(
    $gNewCoords
)


$gNewCoordHint =
New-Object Windows.Forms.Label

$gNewCoordHint.Text =
'可不填写'

$gNewCoordHint.Location =
New-Object Drawing.Point(
    375,
    173
)

$gNewCoordHint.AutoSize =
$true

$gNewCoordHint.ForeColor =
$script:UiPalette.TextSecondary
$gNewCoordHint.Font = $script:UiFont.Small

$gDetailTop.Controls.Add(
    $gNewCoordHint
)

$gBatchDateLabel = New-Object Windows.Forms.Label
$gBatchDateLabel.Text = '本次巡检日期'
$gBatchDateLabel.Location = New-Object Drawing.Point(20, 214)
$gBatchDateLabel.AutoSize = $true
$gDetailTop.Controls.Add($gBatchDateLabel)

$gBatchDate = New-Object Windows.Forms.DateTimePicker
$gBatchDate.Format = 'Custom'
$gBatchDate.CustomFormat = 'yyyy/M/d'
$gBatchDate.Value = (Get-Date).Date
$gBatchDate.Location = New-Object Drawing.Point(155, 207)
$gBatchDate.Size = New-Object Drawing.Size(150, 30)
$gBatchDate.Font = $script:UiFont.Input
$gDetailTop.Controls.Add($gBatchDate)

# 右中：10粒种子的状态/勾选
$gSeedGrid = New-Object Windows.Forms.DataGridView
$gSeedGrid.Dock = 'Fill'
$gSeedGrid.AllowUserToAddRows = $false
$gSeedGrid.AllowUserToDeleteRows = $false
$gSeedGrid.AllowUserToResizeRows = $false
$gSeedGrid.MultiSelect = $false
$gSeedGrid.RowHeadersVisible = $false
$gSeedGrid.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$gSeedGrid.RowTemplate.Height = 34
$gSeedGrid.ColumnHeadersHeight = 38
$gSeedGrid.BackgroundColor = [Drawing.Color]::White
$gSeedGrid.SelectionMode = [Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$gDetailLayout.Controls.Add($gSeedGrid, 0, 1)

$gSeedGrid.ReadOnly = $false


[void]$gSeedGrid.Columns.Add(
    'gSeedNo',
    '测定编号'
)

[void]$gSeedGrid.Columns.Add(
    'gSampleId',
    '样本ID'
)

[void]$gSeedGrid.Columns.Add(
    'gGermDate',
    '发芽日期'
)

[void]$gSeedGrid.Columns.Add(
    'gOriginalCoord',
    '原始坐标'
)

[void]$gSeedGrid.Columns.Add(
    'gSeedStatus',
    '状态'
)


$gSeedGrid.Columns[
'gSeedNo'
].Width = 85

$gSeedGrid.Columns[
'gSampleId'
].Width = 115

$gSeedGrid.Columns[
'gGermDate'
].Width = 120

$gSeedGrid.Columns[
'gOriginalCoord'
].Width = 100

$gSeedGrid.Columns[
'gSeedStatus'
].AutoSizeMode =
[Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill


# 默认都不允许编辑
foreach (
    $columnName in
    @(
        'gSeedNo',
        'gSampleId',
        'gGermDate',
        'gSeedStatus'
    )
) {

    $gSeedGrid.Columns[
    $columnName
    ].ReadOnly = $true
}

foreach (
    $columnName in
    @(
        'gSeedNo',
        'gSampleId',
        'gGermDate',
        'gOriginalCoord'
    )
) {

    $gSeedGrid.Columns[$columnName].DefaultCellStyle.Alignment =
    [Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
}

Set-UiGrid $gSeedGrid


# 右下：批量保存 + 下一物种 + 状态提示
$gDetailBottom = New-Object Windows.Forms.Panel
$gDetailBottom.Dock = 'Fill'
$gDetailBottom.BackColor = $script:UiPalette.Surface
$gDetailBottom.BorderStyle = [Windows.Forms.BorderStyle]::FixedSingle
$gDetailLayout.Controls.Add($gDetailBottom, 0, 2)

$gSaveExistingCoords =
New-Object Windows.Forms.Button

$gSaveExistingCoords.Text =
'保存已有坐标'

$gSaveExistingCoords.Location =
New-Object Drawing.Point(
    18,
    12
)

$gSaveExistingCoords.Size =
New-Object Drawing.Size(
    125,
    42
)

$gDetailBottom.Controls.Add(
    $gSaveExistingCoords
)

$gRecordToday = New-Object Windows.Forms.Button
$gRecordToday.Text = '保存本次巡检'
$gRecordToday.Location = New-Object Drawing.Point(153, 12)
$gRecordToday.Size = New-Object Drawing.Size(170, 42)
$gDetailBottom.Controls.Add($gRecordToday)

$gNextSpecies = New-Object Windows.Forms.Button
$gNextSpecies.Text = '0新增并下一物种 →'
$gNextSpecies.Location = New-Object Drawing.Point(333, 12)
$gNextSpecies.Size = New-Object Drawing.Size(180, 42)
$gDetailBottom.Controls.Add($gNextSpecies)

$gInspectStatus = New-Object Windows.Forms.Label
$gInspectStatus.Text = ''
$gInspectStatus.Location = New-Object Drawing.Point(18, 65)
$gInspectStatus.Size = New-Object Drawing.Size(760, 32)
$gInspectStatus.Font = New-Object Drawing.Font(
    'Microsoft YaHei UI',
    10.5,
    [Drawing.FontStyle]::Bold
)
$gInspectStatus.ForeColor = $script:UiPalette.TextSecondary
$gDetailBottom.Controls.Add($gInspectStatus)

Set-UiSecondaryButton $gSaveExistingCoords
Set-UiPrimaryButton $gRecordToday
Set-UiSecondaryButton $gNextSpecies


# =============================================================================
# 11. UI：根苗长录入页
# =============================================================================

$tabM = New-Object System.Windows.Forms.TabPage
$tabM.Text = '根苗长录入'
$tabM.BackColor = $script:UiPalette.MainBg
$tabs.TabPages.Add($tabM)

# -----------------------------------------------------------------------------
# 11.1 样本定位
# -----------------------------------------------------------------------------
$mL1 = New-Object Windows.Forms.Label
$mL1.Text = '样本ID'
$mL1.Location = New-Object Drawing.Point(40, 40)
$mL1.AutoSize = $true
$mL1.ForeColor = $script:UiPalette.TextSecondary
$tabM.Controls.Add($mL1)

$mSid = New-Object Windows.Forms.TextBox
$mSid.Location = New-Object Drawing.Point(130, 34)
$mSid.Size = New-Object Drawing.Size(240, 35)
$mSid.Font = New-Object Drawing.Font(
    'Microsoft YaHei UI',
    13,
    [Drawing.FontStyle]::Bold
)
$tabM.Controls.Add($mSid)

$mFind = New-Object Windows.Forms.Button
$mFind.Text = '查询'
$mFind.Location = New-Object Drawing.Point(390, 33)
$mFind.Size = New-Object Drawing.Size(85, 36)
$tabM.Controls.Add($mFind)

Set-UiInput $mSid
Set-UiSecondaryButton $mFind

# -----------------------------------------------------------------------------
# 11.2 当前样本信息卡
# -----------------------------------------------------------------------------
$mCard = New-Object Windows.Forms.Panel
$mCard.Location = New-Object Drawing.Point(40, 95)
$mCard.Size = New-Object Drawing.Size(800, 105)
$mCard.BackColor = $script:UiPalette.Surface
$mCard.BorderStyle = 'FixedSingle'
$tabM.Controls.Add($mCard)

$mCardSid = New-Object Windows.Forms.Label
$mCardSid.Text = '—'
$mCardSid.Location = New-Object Drawing.Point(20, 15)
$mCardSid.AutoSize = $true
$mCardSid.Font = New-Object Drawing.Font('Microsoft YaHei UI', 22, [Drawing.FontStyle]::Bold)
$mCardSid.ForeColor = $script:UiPalette.TextPrimary
$mCard.Controls.Add($mCardSid)

$mInfo = New-Object Windows.Forms.Label
$mInfo.Text = '请输入或选择样本'
$mInfo.Location = New-Object Drawing.Point(22, 58)
$mInfo.Size = New-Object Drawing.Size(750, 30)
$mInfo.Font = $script:UiFont.Input
$mInfo.ForeColor = $script:UiPalette.TextSecondary
$mCard.Controls.Add($mInfo)

# -----------------------------------------------------------------------------
# 11.3 测定参数与根/苗长输入
# -----------------------------------------------------------------------------
$mL3 = New-Object Windows.Forms.Label
$mL3.Text = '测定阶段'
$mL3.Location = New-Object Drawing.Point(40, 230)
$mL3.AutoSize = $true
$mL3.ForeColor = $script:UiPalette.TextSecondary
$tabM.Controls.Add($mL3)

$mStage = New-Object Windows.Forms.ComboBox
$mStage.DropDownStyle = 'DropDownList'
[void]$mStage.Items.AddRange(@('3DAG', '7DAG', '14DAG'))
$mStage.SelectedIndex = 0
$mStage.Location = New-Object Drawing.Point(150, 224)
$mStage.Size = New-Object Drawing.Size(180, 35)
$mStage.Font = New-Object Drawing.Font(
    'Microsoft YaHei UI',
    12,
    [Drawing.FontStyle]::Bold
)
$tabM.Controls.Add($mStage)
Set-UiInput $mStage

# 根长
$mL4 = New-Object Windows.Forms.Label
$mL4.Text = '根长（mm）'
$mL4.Location = New-Object Drawing.Point(40, 300)
$mL4.AutoSize = $true
$mL4.ForeColor = $script:UiPalette.TextSecondary
$tabM.Controls.Add($mL4)

$mRoot = New-Object Windows.Forms.TextBox
$mRoot.Location = New-Object Drawing.Point(150, 288)
$mRoot.Size = New-Object Drawing.Size(220, 45)
$mRoot.Font = New-Object Drawing.Font(
    'Microsoft YaHei UI',
    18,
    [Drawing.FontStyle]::Bold
)
$mRoot.TextAlign = 'Center'
$mRoot.BackColor = $script:UiPalette.Surface
$mRoot.ForeColor = $script:UiPalette.TextPrimary
$tabM.Controls.Add($mRoot)

# 苗长
$mL5 = New-Object Windows.Forms.Label
$mL5.Text = '苗长（mm）'
$mL5.Location = New-Object Drawing.Point(40, 365)
$mL5.AutoSize = $true
$mL5.ForeColor = $script:UiPalette.TextSecondary
$tabM.Controls.Add($mL5)

$mShoot = New-Object Windows.Forms.TextBox
$mShoot.Location = New-Object Drawing.Point(150, 353)
$mShoot.Size = New-Object Drawing.Size(220, 45)
$mShoot.Font = New-Object Drawing.Font(
    'Microsoft YaHei UI',
    18,
    [Drawing.FontStyle]::Bold
)
$mShoot.TextAlign = 'Center'
$mShoot.BackColor = $script:UiPalette.Surface
$mShoot.ForeColor = $script:UiPalette.TextPrimary
$tabM.Controls.Add($mShoot)

# -----------------------------------------------------------------------------
# 11.4 保存操作与状态提示
# -----------------------------------------------------------------------------
$mSave = New-Object Windows.Forms.Button
$mSave.Text = '保存'
$mSave.Location = New-Object Drawing.Point(150, 435)
$mSave.Size = New-Object Drawing.Size(120, 44)
$tabM.Controls.Add($mSave)

$mSaveNext = New-Object Windows.Forms.Button
$mSaveNext.Text = '保存并下一条 →'
$mSaveNext.Location = New-Object Drawing.Point(290, 435)
$mSaveNext.Size = New-Object Drawing.Size(170, 44)
$mSaveNext.Font = New-Object Drawing.Font(
    'Microsoft YaHei UI',
    10,
    [Drawing.FontStyle]::Bold
)
$tabM.Controls.Add($mSaveNext)

Set-UiSecondaryButton $mSave
Set-UiPrimaryButton $mSaveNext

# 非弹窗式保存状态
$mStatus = New-Object Windows.Forms.Label
$mStatus.Text = ''
$mStatus.Location = New-Object Drawing.Point(150, 505)
$mStatus.Size = New-Object Drawing.Size(700, 35)
$mStatus.Font = New-Object Drawing.Font(
    'Microsoft YaHei UI',
    11,
    [Drawing.FontStyle]::Bold
)
$mStatus.ForeColor = $script:UiPalette.TextSecondary
$tabM.Controls.Add($mStatus)


# =============================================================================
# 12. UI 业务逻辑：今日任务
# =============================================================================
# 本节只负责“筛选/刷新/跳转”，不直接写 Excel。

function Get-VisibleTodayTasks {
    $query = $taskSearch.Text.Trim().ToLowerInvariant()
    $selectedStage = [string]$stageFilter.SelectedItem

    if ([string]::IsNullOrWhiteSpace($selectedStage)) {
        $selectedStage = '全部'
    }

    $result = New-Object System.Collections.ArrayList

    foreach ($task in @($script:TodayTaskCache)) {
        # DAG筛选
        if ($selectedStage -ne '全部') {
            if ($task.Stage -notlike "*$selectedStage*") {
                continue
            }
        }

        # 文本筛选：物种编号 / 物种名称 / 种子编号 / 样本ID
        if (-not [string]::IsNullOrWhiteSpace($query)) {
            $searchText = (
                "$($task.SpeciesId) " +
                "$($task.SpeciesName) " +
                "$($task.SeedNo) " +
                "$($task.SampleId)"
            ).ToLowerInvariant()

            if (-not $searchText.Contains($query)) {
                continue
            }
        }

        [void]$result.Add($task)
    }

    return @($result)
}

function Refresh-TaskGrid {
    $tasks = @(Get-VisibleTodayTasks)

    $grid.SuspendLayout()

    try {
        $grid.Rows.Clear()

        foreach ($task in $tasks) {
            $rowIndex = $grid.Rows.Add(
                $task.SpeciesId,
                $task.SpeciesName,
                $task.SeedNo,
                $task.SampleId,
                $task.MeasureDate,
                $task.Stage,
                $task.Note
            )

            $row = $grid.Rows[$rowIndex]
            $stageCell = $row.Cells['stage']

            # DAG颜色：只给“测定阶段”单元格上色，避免整表过花
            if ($task.Stage -like '*14DAG*') {
                $stageCell.Style.BackColor = $script:UiPalette.Dag14Soft
            }
            elseif ($task.Stage -like '*7DAG*') {
                $stageCell.Style.BackColor = $script:UiPalette.Dag7Soft
            }
            elseif ($task.Stage -like '*3DAG*') {
                $stageCell.Style.BackColor = $script:UiPalette.Dag3Soft
            }

            # 历史漏测：红色字体提醒
            if ($task.Stage -like '*漏测*') {
                $stageCell.Style.ForeColor = $script:UiPalette.Danger
            }
        }

        $taskCount.Text = "显示 $($tasks.Count) / $($script:TodayTaskCache.Count) 项"
    }
    finally {
        $grid.ResumeLayout()
    }
}

function Reset-TaskFilter {
    # 统一清空搜索框 + DAG筛选，避免多个事件重复触发刷新。
    $script:IgnoreTaskFilterEvents = $true

    try {
        $taskSearch.Text = ''
        $stageFilter.SelectedIndex = 0
    }
    finally {
        $script:IgnoreTaskFilterEvents = $false
    }

    if ($null -ne $script:Book) {
        Refresh-TaskGrid
    }

    $taskSearch.Focus()
}

function Refresh-Ui {
    if ($null -eq $script:Book) {
        $conn.ForeColor = $script:UiPalette.TextSecondary
        $conn.Text = 'Excel：未连接'

        $grid.Rows.Clear()
        $stats.Text = '今日任务 0   |   3DAG 0   |   7DAG 0   |   14DAG 0'
        $taskCount.Text = ''
        return
    }

    $conn.ForeColor = $script:UiPalette.Success
    $conn.Text = "● 已连接 · 可写 · $([IO.Path]::GetFileName($script:WorkbookPath))"

    $tasks = @($script:TodayTaskCache)

    $c3 = 0
    $c7 = 0
    $c14 = 0

    foreach ($task in $tasks) {
        if ($task.Stage -like '*14DAG*') {
            $c14++
        }
        elseif ($task.Stage -like '*7DAG*') {
            $c7++
        }
        elseif ($task.Stage -like '*3DAG*') {
            $c3++
        }
    }

    $stats.Text =
    "今日任务 $($tasks.Count)   |   " +
    "3DAG $c3   |   " +
    "7DAG $c7   |   " +
    "14DAG $c14"

    Refresh-TaskGrid

    if ($null -ne $gSpeciesGrid) {
        Refresh-GerminationUi
    }
}

function Refresh-FromExcel {
    if ($null -eq $script:Book) {
        return
    }

    try {
        $script:PlanSheet.Calculate()
    }
    catch {}

    Rebuild-Cache
    Refresh-Ui
}

function Open-SelectedTask {
    if ($grid.Rows.Count -eq 0) {
        return
    }

    $row = $grid.CurrentRow

    if ($null -eq $row -and $grid.Rows.Count -gt 0) {
        $row = $grid.Rows[0]
    }

    if ($null -eq $row) {
        return
    }

    $sid = [string]$row.Cells['sampleId'].Value

    if ([string]::IsNullOrWhiteSpace($sid)) {
        return
    }

    $mSid.Text = $sid
    $tabs.SelectedTab = $tabM

    Lookup-M
    $mRoot.Focus()
}


# =============================================================================
# 13. UI 业务逻辑：发芽巡检
# =============================================================================
# 本节负责页面展示与交互；实际坐标/日期写入仍调用 Section 06 的业务函数。

function Get-VisibleGerminationSpecies {

    $query =
    $gSearch.Text.Trim().ToLowerInvariant()

    $result =
    New-Object System.Collections.ArrayList

    foreach (
        $item in
        @($script:GerminationSpeciesCache)
    ) {

        if (
            -not
            [string]::IsNullOrWhiteSpace(
                $query
            )
        ) {

            $searchText =
            (
                "$($item.SpeciesId) " +
                "$($item.SpeciesName)"
            ).ToLowerInvariant()

            if (
                -not
                $searchText.Contains(
                    $query
                )
            ) {
                continue
            }
        }

        [void]$result.Add($item)
    }

    return @(
        $result |
        Sort-Object SpeciesId
    )
}


function Get-TodayNewGerminationCount {

    # v0.7：
    # 今日新增必须来自“发芽记录”的培养皿级巡检日志，
    # 不再只统计前10个根苗长测定样本。

    $today =
    (Get-Date).Date

    $count = 0

    foreach (
        $record in
        @($script:GerminationLogCache)
    ) {

        if (
            $null -eq
            $record.InspectionTime
        ) {
            continue
        }

        if (
            $record.InspectionTime.Date -eq
            $today
        ) {

            $count +=
            [int]$record.NewGerminated
        }
    }

    return $count
}


function Refresh-GerminationStats {

    $dishCount =
    $script:GerminationStatusCache.Count

    $totalSeeds = 0
    $germinatedSeeds = 0
    $missingCoordCount = 0

    foreach (
        $status in
        $script:GerminationStatusCache.Values
    ) {

        $totalSeeds +=
        [int]$status.TotalSeeds

        $germinatedSeeds +=
        [int]$status.CumulativeGerminated
    }


    foreach (
        $item in
        @($script:GerminationSpeciesCache)
    ) {

        $missingCoordCount +=
        [int]$item.MissingCoordCount
    }


    $todayNew =
    Get-TodayNewGerminationCount


    $gStats.Text =
    "培养皿 $dishCount   |   " +
    "累计发芽 $germinatedSeeds/$totalSeeds   |   " +
    "今日新增 $todayNew   |   " +
    "未填坐标 $missingCoordCount"
}


function Refresh-GerminationSpeciesGrid(
    [string]$KeepSpeciesId = ''
) {

    if ($null -eq $script:Book) {

        $gSpeciesGrid.Rows.Clear()

        return
    }

    $items =
    @(Get-VisibleGerminationSpecies)

    $gSpeciesGrid.SuspendLayout()

    try {

        $gSpeciesGrid.Rows.Clear()

        foreach ($item in $items) {

            $statusKey =
            "$($item.SpeciesId)|$($script:ExperimentSettings.DefaultReplicate)"

            $germinated = 0
            $totalSeeds =
            [int]$script:ExperimentSettings.TotalSeeds

            $rate = 0.0


            if (
                $script:GerminationStatusCache.ContainsKey(
                    $statusKey
                )
            ) {

                $germinationStatus =
                $script:GerminationStatusCache[
                $statusKey
                ]

                $germinated =
                [int]$germinationStatus.CumulativeGerminated

                $totalSeeds =
                [int]$germinationStatus.TotalSeeds

                $rate =
                [double]$germinationStatus.GerminationRate
            }


            $rateText =
            '{0:N2}%' -f ($rate * 100)


            $germinationProgress =
            "$germinated/$totalSeeds ($rateText)"


            $sampleProgress =
            "$($item.GerminatedCount)/$($item.TotalCount)"


            $rowIndex =
            $gSpeciesGrid.Rows.Add(
                $item.SpeciesId,
                $item.SpeciesName,
                $germinationProgress,
                $sampleProgress,
                $item.MissingCoordCount
            )

            $row =
            $gSpeciesGrid.Rows[$rowIndex]

            if (
                $item.MissingCoordCount -gt 0
            ) {

                $row.Cells[
                'gMissingCoord'
                ].Style.ForeColor =
                $script:UiPalette.Danger

                $row.Cells[
                'gMissingCoord'
                ].Style.BackColor =
                $script:UiPalette.DangerSoft
            }
        }
    }
    finally {

        $gSpeciesGrid.ResumeLayout()
    }


    $targetRow = $null

    if (
        -not
        [string]::IsNullOrWhiteSpace(
            $KeepSpeciesId
        )
    ) {

        foreach (
            $row in
            $gSpeciesGrid.Rows
        ) {

            $sid =
            [string]$row.Cells[
            'gSpeciesId'
            ].Value

            if ($sid -eq $KeepSpeciesId) {

                $targetRow = $row

                break
            }
        }
    }


    if (
        $null -eq $targetRow -and
        $gSpeciesGrid.Rows.Count -gt 0
    ) {

        $targetRow =
        $gSpeciesGrid.Rows[0]
    }


    if ($null -ne $targetRow) {

        $gSpeciesGrid.CurrentCell =
        $targetRow.Cells[
        'gSpeciesId'
        ]

        $speciesId =
        [string]$targetRow.Cells[
        'gSpeciesId'
        ].Value

        Load-GerminationSpeciesDetail `
            $speciesId
    }
    else {

        $script:SelectedGerminationSpeciesId =
        ''

        $gSelectedTitle.Text =
        '暂无物种数据'

        $gSelectedStats.Text =
        ''

        $gNewCoords.Clear()

        $gSeedGrid.Rows.Clear()
    }
}


function Refresh-GerminationUi {

    if ($null -eq $script:Book) {

        $gStats.Text =
        '培养皿 0   |   累计发芽 0/0   |   今日新增 0   |   未填坐标 0'

        $gSpeciesGrid.Rows.Clear()

        $gSeedGrid.Rows.Clear()

        return
    }

    Refresh-GerminationStats

    $keep =
    $script:SelectedGerminationSpeciesId

    Refresh-GerminationSpeciesGrid `
        $keep
}

function Update-GerminationCoordinateHint {

    $speciesId =
    $script:SelectedGerminationSpeciesId


    if (
        [string]::IsNullOrWhiteSpace(
            $speciesId
        )
    ) {
        return
    }


    $item = $null


    foreach (
        $candidate in
        @($script:GerminationSpeciesCache)
    ) {

        if (
            $candidate.SpeciesId -eq
            $speciesId
        ) {

            $item = $candidate

            break
        }
    }


    if ($null -eq $item) {
        return
    }


    $remaining =
    [int]$item.RemainingCount


    if ($remaining -le 0) {

        $gNewCoords.Clear()

        $gNewCoords.Enabled =
        $false

        $gNewCoordHint.Text =
        '测定样本已满，无需填写坐标'

        return
    }


    $text =
    $gNewCount.Text.Trim()


    $newCount = 0


    if (
        [string]::IsNullOrWhiteSpace(
            $text
        ) -or
        -not [int]::TryParse(
            $text,
            [ref]$newCount
        ) -or
        $newCount -lt 0
    ) {

        $gNewCoords.Enabled =
        $true

        $gNewCoordHint.Text =
        "当前还缺 $remaining 个测定样本"

        return
    }


    $required =
    [Math]::Min(
        $newCount,
        $remaining
    )


    if ($required -eq 0) {

        $gNewCoords.Clear()

        $gNewCoords.Enabled =
        $false

        $gNewCoordHint.Text =
        '本次无需填写坐标'
    }
    else {

        $gNewCoords.Enabled =
        $true

        $gNewCoordHint.Text =
        "坐标可选；本次最多填写 $required 个"
    }
}

function Load-GerminationSpeciesDetail(
    [string]$SpeciesId
) {

    if (
        [string]::IsNullOrWhiteSpace(
            $SpeciesId
        )
    ) {
        return
    }

    $item = $null

    foreach (
        $candidate in
        @($script:GerminationSpeciesCache)
    ) {

        if (
            $candidate.SpeciesId -eq
            $SpeciesId
        ) {

            $item = $candidate

            break
        }
    }

    if ($null -eq $item) {
        return
    }


    $script:SelectedGerminationSpeciesId =
    $item.SpeciesId


    $gSelectedTitle.Text =
    "$($item.SpeciesId) · " +
    "$($item.SpeciesName)"


    $statusKey =
    "$($item.SpeciesId)|$($script:ExperimentSettings.DefaultReplicate)"


    $germinated = 0

    $totalSeeds =
    [int]$script:ExperimentSettings.TotalSeeds

    $rate = 0.0


    if (
        $script:GerminationStatusCache.ContainsKey(
            $statusKey
        )
    ) {

        $germinationStatus =
        $script:GerminationStatusCache[
        $statusKey
        ]

        $germinated =
        [int]$germinationStatus.CumulativeGerminated

        $totalSeeds =
        [int]$germinationStatus.TotalSeeds

        $rate =
        [double]$germinationStatus.GerminationRate
    }


    $rateText =
    '{0:N2}%' -f ($rate * 100)


    $gSelectedStats.Text =
    "发芽 $germinated/$totalSeeds（$rateText）   |   " +
    "测定样本 $($item.GerminatedCount)/$($item.TotalCount)   |   " +
    "还需样本 $($item.RemainingCount)   |   " +
    "未填坐标 $($item.MissingCoordCount)"

    $replicate =
    [string]$script:ExperimentSettings.DefaultReplicate


    $existingPlacedDate =
    Get-SpeciesPlacedDate `
        $item.SpeciesId `
        $replicate


    if ($null -ne $existingPlacedDate) {

        # 已经确定置床日期：
        # 显示并锁定，避免后续巡检误改。
        $gPlacedDate.Value =
        ([DateTime]$existingPlacedDate).Date

        $gPlacedDate.Enabled =
        $false

        $gPlacedDateHint.Text =
        '已确定'
    }
    else {

        # 首次巡检：
        # 默认今天，但允许用户修改成实际置床日期。
        $gPlacedDate.Value =
        (Get-Date).Date

        $gPlacedDate.Enabled =
        $true

        $gPlacedDateHint.Text =
        '首次巡检，请确认置床日期'
    }

    $gBatchDate.Value =
    (Get-Date).Date

    $gNewCount.Clear()

    $gNewCoords.Clear()

    Update-GerminationCoordinateHint


    $gSeedGrid.SuspendLayout()

    try {

        $gSeedGrid.Rows.Clear()


        $seeds =
        @(
            $item.Seeds |
            Sort-Object {

                $n = 999

                [void][int]::TryParse(
                    [string]$_.SeedNo,
                    [ref]$n
                )

                $n
            }
        )


        foreach ($seed in $seeds) {

            $isGerminated =
            Test-GerminatedValue `
                $seed.Germination

            $dateText = ''
            $status = '待分配'
            $coord = $seed.Coordinate


            if ($isGerminated) {

                $dateText =
                ExcelDate-ToText `
                    $seed.Germination

                $status = '已发芽'
            }


            $rowIndex =
            $gSeedGrid.Rows.Add(
                $seed.SeedNo,
                $seed.SampleId,
                $dateText,
                $coord,
                $status
            )


            $row =
            $gSeedGrid.Rows[
            $rowIndex
            ]


            # 默认坐标不能编辑
            $row.Cells[
            'gOriginalCoord'
            ].ReadOnly = $true


            if ($isGerminated) {

                if (
                    [string]::IsNullOrWhiteSpace(
                        $coord
                    )
                ) {

                    # 历史已经发芽，
                    # 但以前没有记录原始位置：
                    # 允许人工补录。
                    $row.Cells[
                    'gOriginalCoord'
                    ].ReadOnly = $false

                    $row.Cells[
                    'gOriginalCoord'
                    ].Style.BackColor =
                    $script:UiPalette.WarningSoft

                    $row.Cells[
                    'gSeedStatus'
                    ].Value =
                    '已发芽 · 待补坐标'

                    $row.Cells[
                    'gSeedStatus'
                    ].Style.ForeColor =
                    $script:UiPalette.Warning
                }
                else {

                    $row.DefaultCellStyle.ForeColor =
                    $script:UiPalette.TextSecondary

                    $row.Cells[
                    'gSeedStatus'
                    ].Style.ForeColor =
                    $script:UiPalette.Success
                }
            }
        }
    }
    finally {

        $gSeedGrid.ResumeLayout()
    }
}


function Save-ExistingCoordinateEdits {

    try {

        $speciesId =
        $script:SelectedGerminationSpeciesId

        if (
            [string]::IsNullOrWhiteSpace(
                $speciesId
            )
        ) {

            throw '请先选择一个物种。'
        }


        $assignments = @{}


        foreach (
            $row in
            $gSeedGrid.Rows
        ) {

            $sampleId =
            Safe-Text `
                $row.Cells[
            'gSampleId'
            ].Value

            if (
                [string]::IsNullOrWhiteSpace(
                    $sampleId
                )
            ) {
                continue
            }


            # 已经有坐标的不处理
            if (
                $script:CoordCache.ContainsKey(
                    $sampleId
                )
            ) {

                $oldCoord =
                $script:CoordCache[
                $sampleId
                ].Coordinate

                if (
                    -not
                    [string]::IsNullOrWhiteSpace(
                        $oldCoord
                    )
                ) {

                    continue
                }
            }


            $coord =
            Safe-Text `
                $row.Cells[
            'gOriginalCoord'
            ].Value

            if (
                [string]::IsNullOrWhiteSpace(
                    $coord
                )
            ) {

                continue
            }


            $assignments[$sampleId] =
            $coord
        }


        if ($assignments.Count -eq 0) {

            throw (
                '没有检测到需要补录的坐标。' +
                '请在黄色坐标格中输入，例如 E5。'
            )
        }


        Save-CoordinateBackfill `
            $speciesId `
            $assignments


        $gInspectStatus.ForeColor =
        $script:UiPalette.Success

        $gInspectStatus.Text =
        "✓ 已补录 $($assignments.Count) 个原始坐标"


        Refresh-GerminationUi
    }
    catch {

        $gInspectStatus.ForeColor =
        $script:UiPalette.Danger

        $gInspectStatus.Text =
        '坐标保存失败'

        Handle-Error `
            '保存已有样本坐标失败' `
            $_
    }
}


function Record-NewGerminations {

    try {

        $speciesId =
        $script:SelectedGerminationSpeciesId


        if (
            [string]::IsNullOrWhiteSpace(
                $speciesId
            )
        ) {

            throw '请先选择一个物种。'
        }


        $newCountText =
        $gNewCount.Text.Trim()


        $newCount = 0


        if (
            [string]::IsNullOrWhiteSpace(
                $newCountText
            ) -or
            -not [int]::TryParse(
                $newCountText,
                [ref]$newCount
            ) -or
            $newCount -lt 0
        ) {

            throw (
                '本次新增发芽必须填写大于等于 0 的整数。'
            )
        }


        $coordinates = @()


        if (
            -not
            [string]::IsNullOrWhiteSpace(
                $gNewCoords.Text
            )
        ) {

            $coordinates =
            @(
                Split-GerminationCoordinates `
                    $gNewCoords.Text
            )
        }


        $result =
        Save-GerminationInspection `
            $speciesId `
            $newCount `
            $coordinates `
            $gPlacedDate.Value.Date `
            $gBatchDate.Value.Date


        $rateText =
        '{0:N2}%' -f (
            $result.GerminationRate * 100
        )


        $message =
        "✓ $($result.RecordId) · " +
        "新增 $($result.NewGerminated) · " +
        "累计 $($result.CumulativeGerminated)/$($result.TotalSeeds) " +
        "($rateText)"


        $mapping =
        New-Object `
            System.Collections.ArrayList


        foreach (
            $assignment in
            @($result.AssignedSamples)
        ) {

            [void]$mapping.Add(
                "$($assignment.SampleId)=$($assignment.Coordinate)"
            )
        }


        if ($mapping.Count -gt 0) {

            $message +=
            ' · ' +
            ($mapping -join '，')
        }


        $gInspectStatus.ForeColor =
        $script:UiPalette.Success

        $gInspectStatus.Text =
        $message


        $gNewCount.Clear()

        $gNewCoords.Clear()


        # 同时刷新今日任务和发芽巡检。
        Refresh-Ui
        return $true
    }
    catch {

        $gInspectStatus.ForeColor =
        $script:UiPalette.Danger

        $gInspectStatus.Text =
        '发芽巡检保存失败'


        Handle-Error `
            '保存本次发芽巡检失败' `
            $_
            
        return $false
    }
}

function Record-ZeroAndNextGerminationSpecies {

    # -------------------------------------------------------------------------
    # 高频操作：
    # 当前培养皿已经检查，但本次没有新发芽。
    #
    # 必须真正写入一条“新增 = 0”的巡检记录，
    # 保存成功以后才允许跳到下一物种。
    # -------------------------------------------------------------------------

    if ($null -eq $script:Book) {

        Show-Error '尚未连接 Excel。'
        return
    }


    if (
        [string]::IsNullOrWhiteSpace(
            $script:SelectedGerminationSpeciesId
        )
    ) {

        Show-Error '请先选择一个物种。'
        return
    }


    # 强制本次新增为0。
    $gNewCount.Text =
    '0'


    # 0新增不需要坐标。
    $gNewCoords.Clear()


    Update-GerminationCoordinateHint


    $saved =
    Record-NewGerminations


    # 只有真正保存成功才跳下一物种。
    if ($saved) {

        Select-NextGerminationSpecies
    }
}

function Select-NextGerminationSpecies {

    if (
        $gSpeciesGrid.Rows.Count -eq 0
    ) {
        return
    }


    $index = 0

    if (
        $null -ne
        $gSpeciesGrid.CurrentRow
    ) {

        $index =
        $gSpeciesGrid.CurrentRow.Index + 1
    }


    if (
        $index -ge
        $gSpeciesGrid.Rows.Count
    ) {

        $index = 0
    }


    $row =
    $gSpeciesGrid.Rows[$index]


    $gSpeciesGrid.CurrentCell =
    $row.Cells[
    'gSpeciesId'
    ]


    $speciesId =
    [string]$row.Cells[
    'gSpeciesId'
    ].Value


    Load-GerminationSpeciesDetail `
        $speciesId
}

# =============================================================================
# 14. UI 业务逻辑：根苗长连续录入
# =============================================================================
# 连续录入流程：定位样本 -> 校验阶段/数值 -> 防覆盖 -> 保存 -> 刷新/下一条。

function Lookup-M {
    try {
        $sid = $mSid.Text.Trim()

        if ([string]::IsNullOrWhiteSpace($sid)) {
            return
        }

        $info = Get-SampleInfo $sid

        $mCardSid.Text = $sid

        $mInfo.Text =
        "$($info.SpeciesId) · $($info.SpeciesName) · " +
        "种子 $($info.SeedNo) · 当前状态：$($info.Status)"

        # 每次查询先清空上一样本阶段，避免阶段“串样本”
        $mStage.SelectedIndex = -1

        if ($info.Status -like '*14DAG*') {
            $mStage.SelectedItem = '14DAG'
        }
        elseif ($info.Status -like '*7DAG*') {
            $mStage.SelectedItem = '7DAG'
        }
        elseif ($info.Status -like '*3DAG*') {
            $mStage.SelectedItem = '3DAG'
        }

        $mRoot.Focus()
        $mRoot.SelectAll()
    }
    catch {
        $mCardSid.Text = '—'
        $mInfo.Text = '未找到样本'
        Handle-Error '查询测定样本失败' $_
    }
}

function Get-NextTodayTask([string]$CurrentSampleId) {
    # 按今日任务缓存中的原顺序跳下一条。
    $tasks = @($script:TodayTaskCache)

    if ($tasks.Count -eq 0) {
        return $null
    }

    for ($i = 0; $i -lt $tasks.Count; $i++) {
        if ($tasks[$i].SampleId -eq $CurrentSampleId) {
            if (($i + 1) -lt $tasks.Count) {
                return $tasks[$i + 1]
            }

            return $null
        }
    }

    return $null
}

function Save-CurrentMeasurement([bool]$GoNext) {
    try {
        $sid = $mSid.Text.Trim()

        if ([string]::IsNullOrWhiteSpace($sid)) {
            throw '请先输入样本ID。'
        }

        if ($null -eq $mStage.SelectedItem) {
            throw '当前样本没有可识别的测定阶段，请检查样本状态。'
        }

        # 先校验输入，再进入覆盖判断
        [void](Parse-Measure $mRoot.Text)
        [void](Parse-Measure $mShoot.Text)

        $stageText = [string]$mStage.SelectedItem
        $stage = [int]($stageText.Replace('DAG', ''))

        # 保存前先记住下一条；保存后当前任务会从缓存中消失
        $nextTask = $null

        if ($GoNext) {
            $nextTask = Get-NextTodayTask $sid
        }

        # ---------------------------------------------------------------------
        # 已有数据防覆盖
        # ---------------------------------------------------------------------

        $existing = Get-ExistingMeasurement $sid $stage
        $hasExisting = (Has-Value $existing.Root) -or (Has-Value $existing.Shoot)

        if ($hasExisting) {
            $oldRoot = '空'
            $oldShoot = '空'

            if (Has-Value $existing.Root) {
                $oldRoot = [string]$existing.Root
            }

            if (Has-Value $existing.Shoot) {
                $oldShoot = [string]$existing.Shoot
            }

            $message = @"
样本：$sid
阶段：${stage}DAG

已有数据：
根长：$oldRoot mm
苗长：$oldShoot mm

准备写入：
根长：$($mRoot.Text) mm
苗长：$($mShoot.Text) mm

是否确认覆盖已有数据？
"@

            $answer = [System.Windows.Forms.MessageBox]::Show(
                $message,
                '确认覆盖已有数据',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning,
                [System.Windows.Forms.MessageBoxDefaultButton]::Button2
            )

            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
                $mStatus.ForeColor = $script:UiPalette.Warning
                $mStatus.Text = "已取消：$sid 的 ${stage}DAG 数据未修改"
                return
            }
        }

        Save-Measurement `
            $sid `
            $stage `
            $mRoot.Text `
            $mShoot.Text

        $mStatus.ForeColor = $script:UiPalette.Success
        $mStatus.Text = "✓ $sid · ${stage}DAG 已保存"

        $mRoot.Clear()
        $mShoot.Clear()

        Refresh-Ui

        if ($GoNext -and $null -ne $nextTask) {
            $mSid.Text = $nextTask.SampleId
            Lookup-M
            $mRoot.Focus()
        }
        elseif ($GoNext) {
            $mSid.Clear()
            $mCardSid.Text = '—'
            $mInfo.Text = '今日任务已到最后一条'
            $mStatus.Text = "✓ $sid 已保存 · 今日任务已到最后一条"
            $mSid.Focus()
        }
        else {
            Lookup-M
        }
    }
    catch {
        $mStatus.ForeColor = $script:UiPalette.Danger
        $mStatus.Text = '保存失败'
        Handle-Error '保存根苗长失败' $_
    }
}


# =============================================================================
# 15. 事件绑定
# =============================================================================

# -------------------------------------------------------------------------
# 15.1 Excel 连接与刷新
# -------------------------------------------------------------------------

$choose.Add_Click({
        $dlg = New-Object Windows.Forms.OpenFileDialog
        $dlg.Filter = 'Excel 工作簿|*.xlsx;*.xlsm;*.xlsb;*.xls|所有文件|*.*'

        if ($dlg.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
            try {
                $conn.ForeColor = $script:UiPalette.TextSecondary
                $conn.Text = 'Excel：正在连接……'
                $form.Refresh()

                Connect-Workbook $dlg.FileName
                Refresh-Ui
            }
            catch {
                Show-Error $_.Exception.Message
            }
        }
    })

$refresh.Add_Click({
        try {
            Refresh-FromExcel
        }
        catch {
            Show-Error $_.Exception.Message
        }
    })

# -------------------------------------------------------------------------
# 15.2 今日任务：搜索 / 筛选 / 清除 / 打开
# -------------------------------------------------------------------------

$clearSearch.Add_Click({
        Reset-TaskFilter
    })

$taskSearch.Add_TextChanged({
        if (
            -not $script:IgnoreTaskFilterEvents -and
            $null -ne $script:Book
        ) {
            Refresh-TaskGrid
        }
    })

$stageFilter.Add_SelectedIndexChanged({
        if (
            -not $script:IgnoreTaskFilterEvents -and
            $null -ne $script:Book
        ) {
            Refresh-TaskGrid
        }
    })

$taskSearch.Add_KeyDown({
        param($sender, $e)

        if ($e.KeyCode -eq [Windows.Forms.Keys]::Enter) {
            $e.SuppressKeyPress = $true

            if ($grid.Rows.Count -gt 0) {
                $grid.CurrentCell = $grid.Rows[0].Cells[0]
                Open-SelectedTask
            }
        }
    })

$grid.Add_CellDoubleClick({
        param($sender, $e)

        if ($e.RowIndex -lt 0) {
            return
        }

        $grid.CurrentCell = $grid.Rows[$e.RowIndex].Cells[0]
        Open-SelectedTask
    })

# -------------------------------------------------------------------------
# 15.3 发芽巡检
# -------------------------------------------------------------------------

# 搜索物种
$gSearch.Add_TextChanged({

        if ($null -ne $script:Book) {

            Refresh-GerminationSpeciesGrid `
                $script:SelectedGerminationSpeciesId
        }
    })


# 清除搜索
$gClearSearch.Add_Click({

        $gSearch.Clear()

        Refresh-GerminationSpeciesGrid `
            $script:SelectedGerminationSpeciesId

        $gSearch.Focus()
    })


# 左侧选择物种
$gSpeciesGrid.Add_SelectionChanged({

        if ($null -eq $script:Book) {
            return
        }

        if (
            $null -eq
            $gSpeciesGrid.CurrentRow
        ) {
            return
        }

        $speciesId =
        Safe-Text `
            $gSpeciesGrid.CurrentRow.Cells[
        'gSpeciesId'
        ].Value

        if (
            -not
            [string]::IsNullOrWhiteSpace(
                $speciesId
            )
        ) {

            Load-GerminationSpeciesDetail `
                $speciesId
        }
    })


# 补录以前已发芽样本的原始坐标
$gSaveExistingCoords.Add_Click({

        Save-ExistingCoordinateEdits
    })

# 新增发芽数变化时，自动提示需要填写几个测定样本坐标
$gNewCount.Add_TextChanged({

        Update-GerminationCoordinateHint
    })


# 新增数按 Enter：
# 需要坐标时跳到坐标框；
# 不需要坐标时直接保存。
$gNewCount.Add_KeyDown({

        param($sender, $e)

        if (
            $e.KeyCode -eq
            [Windows.Forms.Keys]::Enter
        ) {

            $e.SuppressKeyPress = $true
            $e.Handled = $true

            Update-GerminationCoordinateHint


            if ($gNewCoords.Enabled) {

                $gNewCoords.Focus()
                $gNewCoords.SelectAll()
            }
            else {

                Record-NewGerminations
            }
        }
    })

# 记录今天新发芽
$gRecordToday.Add_Click({

        Record-NewGerminations
    })


# 坐标输入框按Enter也可以直接保存
$gNewCoords.Add_KeyDown({

        param($sender, $e)

        if (
            $e.KeyCode -eq
            [Windows.Forms.Keys]::Enter
        ) {

            $e.SuppressKeyPress = $true
            $e.Handled = $true

            Record-NewGerminations
        }
    })


# 今天已经检查，但没有新增发芽：
# 保存一条“新增=0”的有效巡检记录，
# 成功后自动进入下一物种。
$gNextSpecies.Add_Click({

        Record-ZeroAndNextGerminationSpecies
    })
    
# -------------------------------------------------------------------------
# 15.4 根苗长录入：全键盘连续录入
# -------------------------------------------------------------------------

$mFind.Add_Click({
        Lookup-M
    })

$mSid.Add_KeyDown({
        param($sender, $e)

        if ($e.KeyCode -eq [Windows.Forms.Keys]::Enter) {
            $e.SuppressKeyPress = $true
            Lookup-M
            $mRoot.Focus()
            $mRoot.SelectAll()
        }
    })

$mRoot.Add_KeyDown({
        param($sender, $e)

        if ($e.KeyCode -eq [Windows.Forms.Keys]::Enter) {
            $e.SuppressKeyPress = $true
            $mShoot.Focus()
            $mShoot.SelectAll()
        }
    })

$mShoot.Add_KeyDown({
        param($sender, $e)

        if ($e.KeyCode -eq [Windows.Forms.Keys]::Enter) {
            $e.SuppressKeyPress = $true
            Save-CurrentMeasurement $true
        }
    })

$mSave.Add_Click({
        Save-CurrentMeasurement $false
    })

$mSaveNext.Add_Click({
        Save-CurrentMeasurement $true
    })

# -------------------------------------------------------------------------
# 15.5 软件关闭：只绑定一次，确保后台 Excel 被彻底退出
# -------------------------------------------------------------------------

$form.Add_FormClosing({
        param($sender, $e)

        try {
            $conn.ForeColor = $script:UiPalette.TextSecondary
            $conn.Text = '正在保存并关闭 Excel……'
            $form.Refresh()

            Disconnect-Workbook
        }
        catch {
            Write-Log "程序关闭清理失败：$($_.Exception.ToString())"
        }
    })


# =============================================================================
# 16. 启动：先显示窗口，再自动连接上次工作簿
# =============================================================================

$form.Add_Shown({
    
        # BeginInvoke 让窗口先出现，减少“启动后长时间没有反应”的感觉。
        $form.BeginInvoke([Action] {
                # 窗口完成布局后，再设置发芽巡检左右区域宽度
                if ($null -ne $gSplit) {

                    if ($gSplit.Width -gt 900) {
                        $gSplit.SplitterDistance = 570
                    }
                    elseif ($gSplit.Width -gt 600) {
                        $gSplit.SplitterDistance = [int]($gSplit.Width * 0.45)
                    }
                }
                try {
                    $conn.ForeColor = $script:UiPalette.TextSecondary
                    $conn.Text = 'Excel：正在连接……'
                    $form.Refresh()

                    if (Test-Path $ConfigFile) {
                        $savedPath = (Get-Content $ConfigFile -Raw).Trim()

                        if (
                            -not [string]::IsNullOrWhiteSpace($savedPath) -and
                            (Test-Path $savedPath)
                        ) {
                            Connect-Workbook $savedPath
                            Refresh-Ui
                        }
                        else {
                            $conn.Text = 'Excel：未连接'
                        }
                    }
                    else {
                        $conn.Text = 'Excel：未连接'
                    }
                }
                catch {
                    $conn.ForeColor = $script:UiPalette.Danger
                    $conn.Text = 'Excel：连接失败'

                    Write-Log $_.Exception.ToString()

                    [System.Windows.Forms.MessageBox]::Show(
                        $_.Exception.Message,
                        '连接失败',
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    ) | Out-Null
                }
            })
    })

Perf-Log '准备显示主界面'

[void]$form.ShowDialog()
