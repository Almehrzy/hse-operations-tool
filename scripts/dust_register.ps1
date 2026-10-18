# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
# Writes the month's dust suppression rounds into REG-017 through Excel (rows from dust_register.py).
# Makes the month sheet from the previous month's when it is missing, adds day rows the sheet lacks (e.g. the 31st),
# fills Time and Remarks only on days that are empty (or "-"), and lists days that already hold something else.
# Backs the register up to backups\ first; stops without saving on any mismatch.
#
#   dust_register.ps1 -RowsJson rows.json [-Workbook <copy for tests>]
param(
    [Parameter(Mandatory)][string]$RowsJson,
    [string]$Workbook = "D:\Claude Projects\Quick Access\17. Registers\Dust Suppersion Register\AVK01-WP.03-HSE-REG-017 - Dust Suppersion Register.xlsx",
    [string]$BackupDir = ""
)

$ErrorActionPreference = "Stop"
# project folder: HSE_PROJECT_DIR, else the nearest folder up from here holding project.json, else the tool folder
$toolRoot = $env:HSE_PROJECT_DIR
if (-not ($toolRoot -and (Test-Path -LiteralPath (Join-Path $toolRoot 'project.json')))) {
    $toolRoot = $null; $d = (Get-Location).Path
    while ($d) { if (Test-Path -LiteralPath (Join-Path $d 'project.json')) { $toolRoot = $d; break }; $up = Split-Path $d -Parent; if (-not $up -or $up -eq $d) { break }; $d = $up }
    if (-not $toolRoot) { $toolRoot = Split-Path $PSScriptRoot -Parent }
}
if (-not $BackupDir) { $BackupDir = Join-Path $toolRoot "backups" }
$req = Get-Content -LiteralPath $RowsJson -Raw -Encoding UTF8 | ConvertFrom-Json
$y, $m = ([string]$req.month).Split("-") | ForEach-Object { [int]$_ }
$nDays = [int]$req.days_in_month

New-Item -ItemType Directory -Force $BackupDir | Out-Null
$base = [IO.Path]::GetFileNameWithoutExtension($Workbook)
$backup = Join-Path $BackupDir ("$base - backup {0}.xlsx" -f (Get-Date -Format "yyyy-MM-dd HHmmss"))
Copy-Item -LiteralPath $Workbook -Destination $backup
"backup: $backup"

function HeaderRow($ws) { for ($r = 1; $r -le 10; $r++) { if (([string]$ws.Cells.Item($r, 1).Text).Trim() -eq "S/N") { return $r } }; throw "No S/N header in sheet $($ws.Name)" }
function FooterRow($ws, $from) { for ($r = $from; $r -le $from + 45; $r++) { if (([string]$ws.Cells.Item($r, 1).Text).Trim() -like "Prepared by*") { return $r } }; throw "No 'Prepared by' footer in sheet $($ws.Name)" }
function DayRows($ws) {
    $h = HeaderRow $ws; $f = FooterRow $ws ($h + 1); $map = @{}
    for ($r = $h + 1; $r -lt $f; $r++) { $v = $ws.Cells.Item($r, 2).Value2; if ($v -is [double]) { $map[[datetime]::FromOADate($v).ToString("yyyy-MM-dd")] = $r } }
    return @{ header = $h; footer = $f; map = $map }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null; $saved = $false
try {
    $wb = $xl.Workbooks.Open($Workbook)
    if ($wb.ReadOnly) { throw "Register opened read-only. Nothing written." }
    $ws = $null
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq $req.sheet) { $ws = $s } }
    $blank = $ws -and ([string]$ws.UsedRange.Address(0, 0) -eq "A1") -and -not $ws.Range("A1").Value2
    if (-not $ws -or $blank) {
        if (-not $ws) {
            $prev = $wb.Worksheets.Item($wb.Worksheets.Count)
            $prev.Copy([Type]::Missing, $prev) | Out-Null
            $ws = $wb.Worksheets.Item($wb.Worksheets.Count)
            $ws.Name = [string]$req.sheet
        } else {                                              # an empty month sheet made in advance: fill it from the month before
            $prev = $wb.Worksheets.Item($ws.Index - 1)
            $prev.UsedRange.Copy($ws.Range("A1")) | Out-Null
            $prev.UsedRange.Copy() | Out-Null
            $ws.Range("A1").PasteSpecial(8) | Out-Null       # column widths
            for ($r = 1; $r -le $prev.UsedRange.Rows.Count; $r++) { $ws.Rows.Item($r).RowHeight = $prev.Rows.Item($r).RowHeight }
            $ws.PageSetup.Orientation = $prev.PageSetup.Orientation
        }
        $ws.Range("A4").Value2 = "Month: " + [string]$req.label
        $h = HeaderRow $ws
        if ($h -gt 5) { $ws.Range("A5").Value2 = "Recorded daily in the HSE Operations Tool: one entry per water-spraying round, with the site photo." }
        $f = FooterRow $ws ($h + 1)
        for ($r = $h + 1; $r -lt $f; $r++) { $ws.Range("D$r`:E$r").ClearContents() | Out-Null }
        $ws.Range("A$($f + 1)`:F$($f + 3)").ClearContents() | Out-Null
        $ws.Range("A$($f + 1)").Value2 = "Name:"; $ws.Range("D$($f + 1)").Value2 = "Name:"
        $ws.Range("A$($f + 2)").Value2 = "Signature:"; $ws.Range("D$($f + 2)").Value2 = "Signature:"
        $ws.Range("A$($f + 3)").Value2 = "Date:"; $ws.Range("D$($f + 3)").Value2 = "Date:"
        "sheet made: $($ws.Name) (from $($prev.Name))"
    }
    # day rows: one per day of the month, in order (dates rewritten for a new sheet; missing days added)
    $info = DayRows $ws
    $h = $info.header
    $last = $h + $nDays
    $have = $info.footer - 1 - $h          # rows between header and footer (one may be the blank spacer)
    if ([string]$ws.Cells.Item($info.footer - 1, 1).Text -eq "") { $have-- }
    while ($have -lt $nDays) { $ws.Rows.Item($h + $have + 1).Insert() | Out-Null; $ws.Rows.Item($h + $have).Copy() | Out-Null; $ws.Rows.Item($h + $have + 1).PasteSpecial(-4122) | Out-Null; $have++; "added a day row" }
    while ($have -gt $nDays) { $ws.Rows.Item($h + $have).Delete() | Out-Null; $have-- }
    for ($i = 1; $i -le $nDays; $i++) {
        $r = $h + $i
        $want = (Get-Date -Year $y -Month $m -Day $i).Date
        $cur = $ws.Cells.Item($r, 2).Value2
        if (-not ($cur -is [double]) -or [datetime]::FromOADate($cur).Date -ne $want) {
            if ($cur -is [double] -and ([string]$ws.Cells.Item($r, 4).Text).Trim() -notin @("", "-")) { throw "Row $r holds $([datetime]::FromOADate($cur).ToString('yyyy-MM-dd')) with an entry, expected $($want.ToString('yyyy-MM-dd')). Nothing written." }
            $ws.Cells.Item($r, 1).Value2 = [string]$i
            $ws.Cells.Item($r, 2).NumberFormat = $ws.Cells.Item($h + 1, 2).NumberFormat
            $ws.Cells.Item($r, 2).Value2 = [string]$want.ToOADate()
            $ws.Cells.Item($r, 3).Formula = "=B$r"
        }
    }
    $info = DayRows $ws
    $written = @(); $kept = @()
    foreach ($p in $req.days.PSObject.Properties) {
        $r = $info.map[$p.Name]
        if (-not $r) { throw "Day $($p.Name) not found in sheet $($ws.Name). Nothing written." }
        $d = ([string]$ws.Cells.Item($r, 4).Text).Trim(); $e = ([string]$ws.Cells.Item($r, 5).Text).Trim()
        if (($d -notin @("", "-")) -or ($e -notin @("", "-"))) { $kept += $p.Name; continue }
        $ws.Cells.Item($r, 4).NumberFormat = "@"
        $ws.Cells.Item($r, 4).Value2 = [string]$p.Value.time
        $ws.Cells.Item($r, 5).Value2 = [string]$p.Value.remarks
        $ws.Cells.Item($r, 4).WrapText = $true; $ws.Cells.Item($r, 5).WrapText = $true
        $ws.Rows.Item($r).AutoFit() | Out-Null
        $written += $p.Name
    }
    $wb.Save(); $saved = $true
    "written: $($written.Count) day(s) $($written -join ', ')"
    if ($kept.Count) { "kept (already had an entry): $($kept -join ', ')" }
}
finally {
    if ($wb) { $wb.Close($false) | Out-Null }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect()
    "saved: $saved"
}
