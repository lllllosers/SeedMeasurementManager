#requires -Version 5.1
# =============================================================================
# 草种测定管理 v0.6.1
# -----------------------------------------------------------------------------
# 当前稳定功能：
#   1. Excel 后台连接：隐藏 Excel COM、缓存加速、保存落盘、退出释放
#   2. 今日任务：统计、搜索、DAG 筛选、双击跳转到根苗长录入
#   3. 发芽巡检：按物种查看 10 个测定槽位、记录新发芽坐标/日期、补坐标
#   4. 根苗长录入：3/7/14DAG 连续录入、Enter 流转、保存并下一条
#   5. 数据安全：已有测定值防覆盖、非法输入拦截、只读工作簿拒绝写入
#   6. 数据同步：保存后计算测定计划表并重建内存缓存
#
# 本版变更（v0.6）：
#   - 不改业务规则和 Excel 数据结构，只整理代码导航与 UI 样式。
#   - 统一颜色、字体、按钮、表格、卡片与状态提示，提高可读性。
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


function Test-GerminatedValue($Value) {
    # 发芽时间非空即视为“已发芽”。
    if ($null -eq $Value) { return $false }

    $text = ([string]$Value).Trim()

    if ($text -eq '' -or $text -eq '0') {
        return $false
    }

    return $true
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
$script:WorkbookPath = ''

# 样本ID -> 样本基础信息
$script:DataCache = @{}

# 样本ID -> 测定计划信息
$script:PlanCache = @{}

# 今日需要执行的任务
$script:TodayTaskCache = @()

# 物种编号 -> 发芽巡检坐标（如 E5）
$script:CoordCache = @{}

# 仅包含“尚未满10粒”的物种
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
    $script:CoordCache = @{}
    $script:GerminationSpeciesCache = @()
    $script:SelectedGerminationSpeciesId = ''
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

function Rebuild-Cache {
    # 性能关键点：
    # Excel 只在这里批量读取一次，之后所有查询/筛选都在内存完成。

    Perf-Log 'Rebuild-Cache 开始'

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
                Germination = $values.GetValue($i, 6)
            }
        }
    }

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


    $incompleteSpecies =
    New-Object `
        System.Collections.ArrayList


    foreach (
        $speciesId in
        $speciesMap.Keys
    ) {

        $group =
        $speciesMap[$speciesId]

        # 满10个以后自动退出发芽巡检
        if (
            $group.GerminatedCount -ge
            $group.TotalCount
        ) {
            continue
        }

        [void]$incompleteSpecies.Add(

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
                $group.TotalCount -
                $group.GerminatedCount

                MissingCoordCount =
                $group.MissingCoordCount

                Seeds             =
                @($group.Seeds)
            }
        )
    }


    $script:GerminationSpeciesCache =
    @($incompleteSpecies)

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

function Save-GerminationsByCoordinate(
    [string]$SpeciesId,
    [string[]]$Coordinates,
    [DateTime]$DateValue
) {

    $species = $null


    foreach (
        $item in
        @($script:GerminationSpeciesCache)
    ) {

        if (
            $item.SpeciesId -eq
            $SpeciesId
        ) {

            $species = $item

            break
        }
    }


    if ($null -eq $species) {

        throw (
            '当前物种已经获得10个测定样本，' +
            '或不在发芽巡检列表中。'
        )
    }


    # -------------------------------------------------------------------------
    # 找出下一个尚未使用的测定样本槽位
    #
    # 例如：
    # 已有001-1～001-3
    #
    # 今天输入：
    # C7 E9
    #
    # 自动得到：
    # 001-4 = C7
    # 001-5 = E9
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


    if (
        $Coordinates.Count -gt
        $blankSlots.Count
    ) {

        throw (
            '当前还需要 ' +
            $blankSlots.Count +
            ' 个测定样本，但输入了 ' +
            $Coordinates.Count +
            ' 个坐标。'
        )
    }


    # -------------------------------------------------------------------------
    # 检查当前培养皿内部坐标是否重复
    #
    # 不同物种之间允许使用同一个坐标：
    # 001可以有E5
    # 009也可以有E5
    # -------------------------------------------------------------------------

    $used =
    Get-UsedCoordinateMap `
        $SpeciesId

    $newUsed = @{}


    foreach (
        $coordRaw in
        $Coordinates
    ) {

        $coord =
        Normalize-GerminationCoordinate `
            $coordRaw


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
    # 写入
    # -------------------------------------------------------------------------

    $oaDate =
    [double]$DateValue.Date.ToOADate()


    $result =
    New-Object `
        System.Collections.ArrayList


    for (
        $i = 0;
        $i -lt $Coordinates.Count;
        $i++
    ) {

        $coord =
        Normalize-GerminationCoordinate `
            $Coordinates[$i]


        $slot =
        $blankSlots[$i]


        $sampleId =
        [string]$slot.SampleId


        # -------------------------------------------------------------
        # 根-苗长统计表 F列：发芽日期
        # -------------------------------------------------------------

        $dataRow =
        [int]$script:DataCache[
        $sampleId
        ].Row


        Set-CellValue `
            $script:DataSheet `
            $dataRow `
            6 `
            $oaDate


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


        # -------------------------------------------------------------
        # 测定时间计划表 N列：原始坐标
        # -------------------------------------------------------------

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


        Set-CellValue `
            $script:PlanSheet `
            $planRow `
            14 `
            $coord


        [void]$result.Add(

            [pscustomobject]@{

                SampleId   =
                $sampleId

                Coordinate =
                $coord
            }
        )
    }


    # DAG等公式重新计算
    $script:PlanSheet.Calculate()


    $script:Book.Save()


    if (-not $script:Book.Saved) {

        throw (
            '发芽数据没有成功保存到Excel。'
        )
    }


    Rebuild-Cache


    return @($result)
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
$form.Text = '草种测定管理 v0.6.1'
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
    '已发芽'
)

[void]$gSpeciesGrid.Columns.Add(
    'gRemaining',
    '还需'
)

[void]$gSpeciesGrid.Columns.Add(
    'gMissingCoord',
    '未记坐标'
)


$gSpeciesGrid.Columns[
'gSpeciesId'
].Width = 90

$gSpeciesGrid.Columns[
'gSpeciesName'
].Width = 210

$gSpeciesGrid.Columns[
'gProgress'
].Width = 80

$gSpeciesGrid.Columns[
'gRemaining'
].Width = 70

$gSpeciesGrid.Columns[
'gMissingCoord'
].Width = 80

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
$gDetailRow1.Height = 180
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

$gNewCoordLabel =
New-Object Windows.Forms.Label

$gNewCoordLabel.Text =
'今天新发芽坐标'

$gNewCoordLabel.Location =
New-Object Drawing.Point(
    20,
    91
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
    140,
    85
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
'多个位置用空格隔开，例如 E5 C7'

$gNewCoordHint.Location =
New-Object Drawing.Point(
    370,
    91
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
$gBatchDateLabel.Text = '本次发芽日期'
$gBatchDateLabel.Location = New-Object Drawing.Point(20, 132)
$gBatchDateLabel.AutoSize = $true
$gDetailTop.Controls.Add($gBatchDateLabel)

$gBatchDate = New-Object Windows.Forms.DateTimePicker
$gBatchDate.Format = 'Custom'
$gBatchDate.CustomFormat = 'yyyy/M/d'
$gBatchDate.Value = (Get-Date).Date
$gBatchDate.Location = New-Object Drawing.Point(125, 125)
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
$gRecordToday.Text = '记录今日新发芽'
$gRecordToday.Location = New-Object Drawing.Point(153, 12)
$gRecordToday.Size = New-Object Drawing.Size(170, 42)
$gDetailBottom.Controls.Add($gRecordToday)

$gNextSpecies = New-Object Windows.Forms.Button
$gNextSpecies.Text = '下一物种 →'
$gNextSpecies.Location = New-Object Drawing.Point(333, 12)
$gNextSpecies.Size = New-Object Drawing.Size(110, 42)
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

    $today =
    (Get-Date).Date

    $count = 0

    foreach (
        $seed in
        $script:DataCache.Values
    ) {

        $isGerminated =
        Test-GerminatedValue `
            $seed.Germination

        if (-not $isGerminated) {
            continue
        }

        try {

            if (
                $seed.Germination -is
                [double]
            ) {

                $date =
                [DateTime]::FromOADate(
                    [double]$seed.Germination
                ).Date
            }
            else {

                $date =
                [DateTime]::Parse(
                    [string]$seed.Germination
                ).Date
            }

            if ($date -eq $today) {
                $count++
            }
        }
        catch {}
    }

    return $count
}


function Refresh-GerminationStats {

    $speciesCount =
    $script:GerminationSpeciesCache.Count

    $remainingCount = 0
    $missingCoordCount = 0

    foreach (
        $item in
        @($script:GerminationSpeciesCache)
    ) {

        $remainingCount +=
        $item.RemainingCount

        $missingCoordCount +=
        $item.MissingCoordCount
    }

    $todayNew =
    Get-TodayNewGerminationCount

    $gStats.Text =
    "待检查物种 $speciesCount   |   " +
    "还需发芽样本 $remainingCount   |   " +
    "今日新增 $todayNew   |   " +
    "未记坐标 $missingCoordCount"
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

            $rowIndex =
            $gSpeciesGrid.Rows.Add(
                $item.SpeciesId,
                $item.SpeciesName,
                "$($item.GerminatedCount)/$($item.TotalCount)",
                $item.RemainingCount,
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
        '暂无待检查物种'

        $gSelectedStats.Text =
        ''

        $gNewCoords.Clear()

        $gSeedGrid.Rows.Clear()
    }
}


function Refresh-GerminationUi {

    if ($null -eq $script:Book) {

        $gStats.Text =
        '待检查物种 0   |   还需发芽样本 0   |   今日新增 0   |   未记坐标 0'

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


    $gSelectedStats.Text =
    "已获得 $($item.GerminatedCount)/$($item.TotalCount) 个测定样本；" +
    "还需要 $($item.RemainingCount) 个；" +
    "已有样本未记坐标 $($item.MissingCoordCount) 个"


    $gBatchDate.Value =
    (Get-Date).Date

    $gNewCoords.Clear()


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

            throw '请先选择一个待检查物种。'
        }


        $coordinates =
        @(
            Split-GerminationCoordinates `
                $gNewCoords.Text
        )


        $result =
        @(
            Save-GerminationsByCoordinate `
                $speciesId `
                $coordinates `
                $gBatchDate.Value.Date
        )


        $mapping =
        New-Object `
            System.Collections.ArrayList


        foreach ($item in $result) {

            [void]$mapping.Add(
                "$($item.SampleId)=$($item.Coordinate)"
            )
        }


        $gInspectStatus.ForeColor =
        $script:UiPalette.Success

        $gInspectStatus.Text =
        '✓ 已记录：' +
        ($mapping -join '，')


        $gNewCoords.Clear()


        # 这里会同时刷新今日任务和发芽巡检。
        Refresh-Ui
    }
    catch {

        $gInspectStatus.ForeColor =
        $script:UiPalette.Danger

        $gInspectStatus.Text =
        '发芽记录失败'

        Handle-Error `
            '记录今日新发芽失败' `
            $_
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


# 今天没有新增，跳下一物种
$gNextSpecies.Add_Click({

        Select-NextGerminationSpecies
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
