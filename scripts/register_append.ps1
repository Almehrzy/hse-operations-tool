# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
# Adds rows to the approved SVR register through Excel itself, so logos and formatting are kept.
# Never overwrites a cell that already has a value. Stops without saving on any mismatch.
#
#   register_append.ps1 -RowsJson rows.json
#
# rows.json: a list of
#   {"no": 214, "issued": "NAME\nNAME", "iqama": "...", "designation": "...", "details": "...", "company": "...",
#    "penalty": "...", "date": "2026-10-14", "reported_by": "...", "remarks": ""}
param(
    [Parameter(Mandatory)][string]$RowsJson,
    [string]$Register = "",
    [string]$BackupDir = ""
)

$ErrorActionPreference = "Stop"
$toolRoot = Split-Path $PSScriptRoot -Parent
$cfg = Get-Content -LiteralPath (Join-Path $toolRoot "project.json") -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $Register) { $Register = $cfg.svr.register }
if (-not $BackupDir) { $BackupDir = Join-Path $toolRoot "backups" }
$sheetName = $cfg.svr.register_sheet
$rows = @(Get-Content -LiteralPath $RowsJson -Raw -Encoding UTF8 | ConvertFrom-Json)
if ($rows.Count -eq 0) { throw "No rows in $RowsJson" }

New-Item -ItemType Directory -Force $BackupDir | Out-Null
$hash = (Get-FileHash -LiteralPath $Register).Hash
$existing = Get-ChildItem -LiteralPath $BackupDir -Filter "SVR LOG REGISTER 2024 - backup *.xlsx" | Where-Object { (Get-FileHash $_.FullName).Hash -eq $hash } | Select-Object -First 1
if ($existing) { "backup already holds the current register: $($existing.Name)" }
else {
    $backup = Join-Path $BackupDir ("SVR LOG REGISTER 2024 - backup {0}.xlsx" -f (Get-Date -Format "yyyy-MM-dd HHmmss"))
    Copy-Item -LiteralPath $Register -Destination $backup
    "backup: $backup"
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
$saved = $false
try {
    $wb = $xl.Workbooks.Open($Register)
    if ($wb.ReadOnly) { throw "Register opened read-only (open on another computer?). Nothing written." }
    $ws = $wb.Worksheets.Item($sheetName)

    # Pass 1: check every target row before touching anything.
    foreach ($r in $rows) {
        $row = [int]$r.no + 4
        $prev = $ws.Cells.Item($row - 1, 1).Value2
        if ($row -gt 5 -and ($null -eq $prev -or [int]$prev -ne ([int]$r.no - 1)) -and -not ($rows | Where-Object { [int]$_.no -eq ([int]$r.no - 1) })) {
            throw "SVR $([int]$r.no - 1) is missing from the row above. Numbering would have a gap. Nothing written."
        }
        $a = $ws.Cells.Item($row, 1).Value2
        if ($null -ne $a -and "$a" -ne "" -and [int]$a -ne [int]$r.no) { throw "Row $row holds SVR $a, expected $($r.no). Nothing written." }
        for ($c = 2; $c -le 10; $c++) {
            $v = $ws.Cells.Item($row, $c).Value2
            if ($null -ne $v -and "$v".Trim() -ne "") { throw "SVR $($r.no) already has data in the register (row $row, column $c). Nothing written." }
        }
    }

    # Pass 2: write.
    $maxRow = 0
    foreach ($r in $rows) {
        $row = [int]$r.no + 4
        if ($ws.Cells.Item($row, 2).Borders.Item(9).LineStyle -eq -4142) {      # no bottom border: row not formatted yet
            $ws.Rows.Item($row - 1).Copy() | Out-Null
            $ws.Rows.Item($row).PasteSpecial(-4122) | Out-Null                    # formats only
        }
        $ws.Cells.Item($row, 1).Value2 = [int]$r.no
        $ws.Cells.Item($row, 2).Value2 = [string]$r.issued
        $ws.Cells.Item($row, 3).NumberFormat = "@"
        $ws.Cells.Item($row, 3).Value2 = [string]$r.iqama
        $ws.Cells.Item($row, 4).Value2 = [string]$r.designation
        $ws.Cells.Item($row, 5).Value2 = [string]$r.details
        $ws.Cells.Item($row, 6).Value2 = [string]$r.company
        $ws.Cells.Item($row, 7).Value2 = [string]$r.penalty
        $d = [datetime]::ParseExact([string]$r.date, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture)
        $ws.Cells.Item($row, 8).NumberFormat = "dd/mm/yyyy"
        $ws.Cells.Item($row, 8).Value2 = $d.ToOADate()
        if ($r.reported_by) { $ws.Cells.Item($row, 9).Value2 = [string]$r.reported_by }
        if ($r.remarks) { $ws.Cells.Item($row, 10).Value2 = [string]$r.remarks }
        $ws.Rows.Item($row).AutoFit() | Out-Null
        $lines = [Math]::Max(([string]$r.issued).Split("`n").Count, ([string]$r.iqama).Split("`n").Count)
        $need = [Math]::Max(27, 15.75 * $lines)
        if ($ws.Rows.Item($row).RowHeight -lt $need) { $ws.Rows.Item($row).RowHeight = $need }
        if ($row -gt $maxRow) { $maxRow = $row }
    }
    $area = $ws.PageSetup.PrintArea
    if ($area -match '\$(\d+)$' -and [int]$Matches[1] -lt $maxRow) { $ws.PageSetup.PrintArea = "`$A`$1:`$J`$$maxRow" }
    $wb.Save()
    $saved = $true
    "written: $(($rows | ForEach-Object { 'SVR ' + $_.no }) -join ', ')"
}
finally {
    if ($wb) { $wb.Close($false) | Out-Null }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect()
    "saved: $saved"
}
