# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
# Reads the weekly HSE reports (project.json weekly.file_glob, .xls) through Excel and writes what the monthly report needs, one entry per
# file: the dates on the report, project-to-date figures, the Daily Manpower rows, TBT, Trainings and Inspections rows,
# and the monitoring / incident / training numbers of the week. Read-only: nothing is changed.
#
#   read_weekly.ps1 -Out weekly.json [-Folder "<weekly reports folder>"] [-Since 2026-08-01]   (default: weekly.folder, else trend.folder)
param(
    [Parameter(Mandatory)][string]$Out,
    [string]$Folder = "",
    [string]$Since = ""
)

$ErrorActionPreference = "Stop"
# project folder: HSE_PROJECT_DIR, else the nearest folder up from here holding project.json, else the tool folder
$toolRoot = $env:HSE_PROJECT_DIR
if (-not ($toolRoot -and (Test-Path -LiteralPath (Join-Path $toolRoot 'project.json')))) {
    $toolRoot = $null; $d = (Get-Location).Path
    while ($d) { if (Test-Path -LiteralPath (Join-Path $d 'project.json')) { $toolRoot = $d; break }; $up = Split-Path $d -Parent; if (-not $up -or $up -eq $d) { break }; $d = $up }
    if (-not $toolRoot) { $toolRoot = Split-Path $PSScriptRoot -Parent }
}
$cfg = Get-Content -LiteralPath (Join-Path $toolRoot "project.json") -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $Folder) { $Folder = if ($cfg.weekly.folder) { [string]$cfg.weekly.folder } else { [string]$cfg.trend.folder } }
$glob = if ($cfg.weekly.file_glob) { [string]$cfg.weekly.file_glob } else { "*Weekly HSE report*.xls*" }
$files = Get-ChildItem -LiteralPath $Folder -Recurse -File -Filter $glob | Where-Object { $_.Name -notlike "~$*" }
if ($Since) { $sinceDate = [datetime]::ParseExact($Since, "yyyy-MM-dd", $null); $files = $files | Where-Object { $_.LastWriteTime -ge $sinceDate.AddDays(-10) } }

function D($v) { if ($v -is [double] -and $v -gt 30000) { return [datetime]::FromOADate($v).ToString("yyyy-MM-dd") } return [string]$v }
function N($v) { if ($null -eq $v) { return $null } $s = "$v".Replace(",", "").Trim(); if ($s -match '^-?\d+(\.\d+)?$') { return [double]$s } return $null }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$all = @()
try {
    foreach ($f in $files) {
        $wb = $xl.Workbooks.Open($f.FullName, 0, $true)
        try {
            $ws = $null
            foreach ($s in $wb.Worksheets) { if ($s.Name -like "Weekly HSE*") { $ws = $s } }
            if (-not $ws) { continue }
            $e = [ordered]@{ file = $f.FullName; from = (D $ws.Range("G2").Value2); to = (D $ws.Range("K2").Value2); reportDate = (D $ws.Range("K3").Value2) }
            $e.ptdManpower = N $ws.Range("L11").Value2
            $e.ptdHours = N $ws.Range("L12").Value2
            $e.ptdSafeHours = N $ws.Range("L14").Value2
            $e.weekManpower = N $ws.Range("F11").Value2
            $e.weekHours = N $ws.Range("F12").Value2
            $rows = [ordered]@{}
            foreach ($r in @(20..29 + 40..59 + 63..66)) { $rows["$r"] = N $ws.Range("D$r").Value2 }
            $e.rows = $rows
            $dm = @()
            try {
                $s2 = $wb.Worksheets.Item("Daily Manpower")
                for ($r = 4; $r -le 10; $r++) {
                    $dm += , @((D $s2.Cells.Item($r, 2).Value2), [string]$s2.Cells.Item($r, 1).Value2, (N $s2.Cells.Item($r, 12).Value2), (N $s2.Cells.Item($r, 13).Value2),
                               (N $s2.Cells.Item($r, 3).Value2), (N $s2.Cells.Item($r, 4).Value2), (N $s2.Cells.Item($r, 5).Value2), (N $s2.Cells.Item($r, 6).Value2),
                               (N $s2.Cells.Item($r, 7).Value2), (N $s2.Cells.Item($r, 8).Value2), (N $s2.Cells.Item($r, 9).Value2), (N $s2.Cells.Item($r, 10).Value2), (N $s2.Cells.Item($r, 11).Value2))
                }
            } catch {}
            $e.dailyManpower = $dm
            $tbt = @()
            try { $s3 = $wb.Worksheets.Item("TBT"); for ($r = 4; $r -le 40; $r++) { $t = [string]$s3.Cells.Item($r, 2).Text; if ($t.Trim()) { $tbt += , @([string]$s3.Cells.Item($r, 1).Text, $t.Trim(), [string]$s3.Cells.Item($r, 3).Text, [string]$s3.Cells.Item($r, 4).Text) } } } catch {}
            $e.tbt = $tbt
            $tr = @()
            try { $s4 = $wb.Worksheets.Item("Trainings"); for ($r = 4; $r -le 40; $r++) { $t = [string]$s4.Cells.Item($r, 2).Text; if ($t.Trim()) { $tr += , @($t.Trim(), [string]$s4.Cells.Item($r, 3).Text, [string]$s4.Cells.Item($r, 4).Text, [string]$s4.Cells.Item($r, 5).Text) } } } catch {}
            $e.trainings = $tr
            $ins = @()
            # the sheet has three blocks: inspections (rows 4-17), walkthroughs/tours (21-27) and drills (30-40)
            $tours = @(); $drills = @()
            try {
                $s5 = $wb.Worksheets.Item("Inspections")
                for ($r = 4; $r -le 40; $r++) {
                    $t = [string]$s5.Cells.Item($r, 2).Text
                    if (-not $t.Trim() -or $t -match '^(Walkthroughs|Drills|Inspection Topic)') { continue }
                    $row = @($r, $t.Trim(), (D $s5.Cells.Item($r, 6).Value2), [string]$s5.Cells.Item($r, 7).Text, [string]$s5.Cells.Item($r, 9).Text)
                    if ($r -le 17) { $ins += , $row } elseif ($r -le 27) { $tours += , $row } else { $drills += , $row }
                }
            } catch {}
            $e.inspections = $ins; $e.tours = $tours; $e.drills = $drills
            $all += [pscustomobject]$e
        }
        finally { $wb.Close($false) | Out-Null }
    }
}
finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect()
}
$all | ConvertTo-Json -Depth 6 | Out-File -LiteralPath $Out -Encoding utf8
"read $($all.Count) weekly reports -> $Out"
