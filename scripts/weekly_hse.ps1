# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
# Writes the weekly HSE report (HSE-FRM-00015) through Excel: copies the previous week's report (values.json "source")
# to the new file ("out"), then writes this week's numbers, carries month (when the week is in the same month) and
# project-to-date totals forward, and replaces the detail sheets. Never overwrites a file. values.json comes from weekly_hse.py.
#
#   weekly_hse.ps1 -ValuesJson values.json [-Out other.xls]
param(
    [Parameter(Mandatory)][string]$ValuesJson,
    [string]$Out = ""
)

$ErrorActionPreference = "Stop"
$v = Get-Content -LiteralPath $ValuesJson -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $Out) { $Out = $v.out }
if (Test-Path -LiteralPath $Out) { throw "Refusing to overwrite $Out" }
New-Item -ItemType Directory -Force (Split-Path $Out -Parent) | Out-Null
Copy-Item -LiteralPath $v.source -Destination $Out

function Serial([string]$ymd) { [string][int]([datetime]::ParseExact($ymd, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture).ToOADate()) }
function Num($cell) { $x = $cell.Value2; if ($null -eq $x -or "$x".Trim() -eq "" -or "$x".Trim() -eq "-") { return 0.0 } return [double]("$x".Replace(",", "")) }
function Put($cell, $val) { if (-not $cell.HasFormula) { $cell.Value2 = [string]$val } }
function Norm([string]$s) { ($s.ToLower() -replace '\s*/\s*', '/' -replace '\s+', ' ').Trim() }
function ClearCells($range) {   # not "Clear": that name is the alias of Clear-Host, which wins
    # ClearContents fails on a range that cuts through merged cells, so clear each merged area through its first cell.
    # (indexed loops: enumerating Range.Cells from PowerShell yields nothing)
    for ($i = 1; $i -le $range.Rows.Count; $i++) {
        for ($j = 1; $j -le $range.Columns.Count; $j++) {
            $c = $range.Cells.Item($i, $j)
            if (-not $c.MergeCells) { if ($null -ne $c.Value2) { $c.ClearContents() | Out-Null } }
            elseif ($c.Address() -eq $c.MergeArea.Cells.Item(1, 1).Address()) { $c.MergeArea.ClearContents() | Out-Null }
        }
    }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
$saved = $false
$log = @()
try {
    $wb = $xl.Workbooks.Open($Out)
    $ws = $wb.Worksheets.Item("Weekly HSE Statistics")
    $prevMonthHours = Num $ws.Range("G12")            # before anything changes: last week's month-to-date hours
    $ws.Range("G2").Value2 = Serial $v.dates.from
    $ws.Range("K2").Value2 = Serial $v.dates.to
    $ws.Range("K3").Value2 = Serial $v.dates.report

    # manpower: this week, month to date, project to date (typed cells only; formulas stay)
    foreach ($pair in @(@("D9", "G9", "J9"), @("E9", "H9", "K9"), @("D10", "G10", "J10"), @("E10", "H10", "K10"))) {
        $w = [double]$v.manpower.($pair[0])
        $m = if ($v.sameMonth) { (Num $ws.Range($pair[1])) + $w } else { $w }
        $p = (Num $ws.Range($pair[2])) + $w
        Put $ws.Range($pair[0]) $w; Put $ws.Range($pair[1]) $m; Put $ws.Range($pair[2]) $p
    }
    # man-hours: D12 and G12 are formulas (=D11*11, =G11*11). They stay while every day had 11 hours; otherwise the
    # real hours are written. J12 is typed, so it is carried forward.
    $xl.Calculate()
    if ([double](Num $ws.Range("D12")) -ne [double]$v.manpower.D12) {
        $ws.Range("D12").Value2 = [string]$v.manpower.D12
        $ws.Range("G12").Value2 = [string]$(if ($v.sameMonth) { $prevMonthHours + [double]$v.manpower.D12 } else { [double]$v.manpower.D12 })
        $log += "man-hours written as values (not every day had 11 hours)"
    }
    Put $ws.Range("J12") ((Num $ws.Range("J12")) + [double]$v.manpower.D12)
    Put $ws.Range("E12") $v.manpower.E12

    # incidents, monitoring, training rows: D this week, G month, J project to date
    foreach ($p in $v.rows.PSObject.Properties) {
        $r = [int]$p.Name; $w = [double]$p.Value
        $g = $ws.Range("G$r"); $j = $ws.Range("J$r")
        Put $ws.Range("D$r") $w
        Put $g $(if ($v.sameMonth) { (Num $g) + $w } else { $w })
        Put $j ((Num $j) + $w)
    }
    foreach ($c in @("D", "G")) { if (-not $ws.Range("${c}30").HasFormula) { $ws.Range("${c}30").Value2 = "=SUM(${c}20:${c}29)" } }

    # Daily Manpower
    $dm = $wb.Worksheets.Item("Daily Manpower")
    $i = 4
    foreach ($d in $v.days) {
        $dm.Cells.Item($i, 1).Value2 = [string]$d[0]
        $dm.Cells.Item($i, 2).Value2 = Serial $d[1]
        for ($k = 0; $k -lt 9; $k++) {
            $val = [int]$d[2 + $k]
            $cell = $dm.Cells.Item($i, 3 + $k)
            if ($k % 3 -eq 2 -and $val -eq 0) { $cell.Value2 = "-" } else { $cell.Value2 = [string]$val }
        }
        if ([double]$d[11] -ne 11) { $dm.Cells.Item($i, 13).Value2 = "=L$i*$([double]$d[11])" }
        $i++
    }

    # TBT: replace the rows, continuing the serial numbers
    $tb = $wb.Worksheets.Item("TBT")
    $last = 0; for ($r = 4; $r -le 26; $r++) { $a = $tb.Cells.Item($r, 1).Value2; if ("$a" -match '^\d+(\.0+)?$') { $last = [Math]::Max($last, [int]$a) } }
    ClearCells $tb.Range("A4:D26")
    $r = 4
    foreach ($x in $v.tbt) { $last++; $tb.Cells.Item($r, 1).Value2 = [string]$last; $tb.Cells.Item($r, 2).Value2 = [string]$x[0]; $tb.Cells.Item($r, 3).Value2 = [string]$x[1]; $tb.Cells.Item($r, 4).Value2 = [string]$x[2]; $r++ }

    # Trainings
    $trn = $wb.Worksheets.Item("Trainings")
    ClearCells $trn.Range("A4:E30")
    $r = 4; $n = 0
    foreach ($x in $v.trainings) { $n++; $trn.Cells.Item($r, 1).Value2 = [string]$n; $trn.Cells.Item($r, 2).Value2 = [string]$x[0]; $trn.Cells.Item($r, 3).Value2 = [string]$x[1]; $trn.Cells.Item($r, 4).Value2 = [string]$x[2]; $trn.Cells.Item($r, 5).Value2 = [string]$x[3]; $r++ }

    # Inspections (rows 4-17), walkthroughs/tours (21-27), drills (30-...)
    $ins = $wb.Worksheets.Item("Inspections")
    foreach ($blk in @(@(4, 17, $v.inspections), @(21, 27, $v.tours), @(30, 39, $v.drills))) {
        ClearCells $ins.Range("A$($blk[0]):B$($blk[1])")
        ClearCells $ins.Range("F$($blk[0]):I$($blk[1])")
        $r = $blk[0]; $n = 0
        foreach ($x in $blk[2]) {
            if ($r -gt $blk[1]) { $log += "more rows than the sheet holds at row $($blk[0]); the rest were left out"; break }
            $n++
            $ins.Cells.Item($r, 1).Value2 = [string]$n
            $ins.Cells.Item($r, 2).Value2 = [string]$x[0]
            $ins.Cells.Item($r, 6).NumberFormat = "m/d/yyyy"; $ins.Cells.Item($r, 6).Value2 = Serial $x[1]
            $ins.Cells.Item($r, 7).Value2 = [string]$x[2]
            $ins.Cells.Item($r, 9).Value2 = [string]$x[3]
            $r++
        }
    }

    # HSE Observations by topic
    $ho = $wb.Worksheets.Item("HSE Observations")
    $topics = @{}; foreach ($p in $v.observations.PSObject.Properties) { $topics[$p.Name] = $p.Value }
    $used = @{}
    for ($r = 3; $r -le 43; $r++) {
        $k = Norm ([string]$ho.Cells.Item($r, 2).Value2)
        if ($topics.ContainsKey($k)) { $ho.Cells.Item($r, 3).Value2 = [string]$topics[$k][0]; $ho.Cells.Item($r, 4).Value2 = [string]$topics[$k][1]; $used[$k] = 1 }
        else { ClearCells $ho.Range("C${r}:D${r}") }
    }
    foreach ($k in $topics.Keys) { if (-not $used.ContainsKey($k)) { $log += "observation topic not on the sheet: $k" } }

    $wb.Save()
    $saved = $true
    $xl.CalculateFull()
    "written: $Out"
    "this week: manpower $($ws.Range('F11').Text), man-hours $($ws.Range('F12').Text); month $($ws.Range('I11').Text); to date $($ws.Range('L11').Text) / $($ws.Range('L12').Text)"
    $log
}
finally {
    if ($wb) { $wb.Close($false) | Out-Null }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect()
    if (-not $saved -and (Test-Path -LiteralPath $Out)) { "not saved; the copy at $Out is the unchanged previous report and should be removed by hand" }
}
