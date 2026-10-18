# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
# Adds issued permits to the permit register (REG-009, one sheet per month), or records their closing date,
# through Excel itself so logos and formatting are kept. Backs the register up first. Stops without saving on any mismatch.
#
#   permit_register.ps1 -RowsJson rows.json            add rows (a missing month sheet is made from the last sheet's layout)
#   permit_register.ps1 -RowsJson rows.json -Close     set PERMIT CLOSED (column E) on rows already in the register
#
# rows.json: a list of
#   {"sheet": "October", "type": "HOT WORK", "no": "139 A", "date": "2026-10-17", "start": "06:00", "end": "17:00", "company": "ACTC",
#    "area": "UP 03", "description": "CUTTING WELDING GRINDING", "issuer": "KHALID AL MOOR", "receiver": "SALAH YOUSAF",
#    "contact": "544213838", "hse": "ZUHAIB"}
# with -Close: {"sheet": "October", "type": "HOT WORK", "no": "139 A", "area": "UP 03", "closed": "2026-10-23"}
param(
    [Parameter(Mandatory)][string]$RowsJson,
    [switch]$Close,
    [string]$Register = "",
    [string]$BackupDir = ""
)

$ErrorActionPreference = "Stop"
$toolRoot = Split-Path $PSScriptRoot -Parent
$pj = Get-Content -LiteralPath (Join-Path $toolRoot "permits.json") -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $Register) { $Register = $pj.register }
if (-not $BackupDir) { $BackupDir = Join-Path $toolRoot "backups" }
$rows = @(Get-Content -LiteralPath $RowsJson -Raw -Encoding UTF8 | ConvertFrom-Json | ForEach-Object { $_ })   # unroll: PS 5.1 emits a JSON array as one object
if ($rows.Count -eq 0) { throw "No rows in $RowsJson" }

New-Item -ItemType Directory -Force $BackupDir | Out-Null
$base = [IO.Path]::GetFileNameWithoutExtension($Register)
$hash = (Get-FileHash -LiteralPath $Register).Hash
$existing = Get-ChildItem -LiteralPath $BackupDir -Filter "$base - backup *.xlsx" | Where-Object { (Get-FileHash $_.FullName).Hash -eq $hash } | Select-Object -First 1
if ($existing) { "backup already holds the current register: $($existing.Name)" }
else {
    $backup = Join-Path $BackupDir ("{0} - backup {1}.xlsx" -f $base, (Get-Date -Format "yyyy-MM-dd HHmmss"))
    Copy-Item -LiteralPath $Register -Destination $backup
    "backup: $backup"
}

function Serial([string]$ymd) { [string][int]([datetime]::ParseExact($ymd, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture).ToOADate()) }
function TimeVal([string]$hm) { $p = $hm.Split(":"); [string](([int]$p[0] * 60 + [int]$p[1]) / 1440.0).ToString([Globalization.CultureInfo]::InvariantCulture) }
function LastRow($ws) {
    # End(xlUp) is unreliable on these sheets (it returned the header row on a full sheet), so scan up from the used range.
    $r = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
    while ($r -gt 3 -and ("$($ws.Cells.Item($r, 3).Value2)".Trim() -eq "" -and "$($ws.Cells.Item($r, 2).Value2)".Trim() -eq "")) { $r-- }
    return [Math]::Max($r, 3)
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
$saved = $false
try {
    $wb = $xl.Workbooks.Open($Register)
    if ($wb.ReadOnly) { throw "Register opened read-only (open on another computer?). Nothing written." }
    $names = @(); foreach ($s in $wb.Worksheets) { $names += $s.Name }

    if ($Close) {
        foreach ($r in $rows) {
            if ($names -notcontains $r.sheet) { throw "Sheet $($r.sheet) is not in the register. Nothing written." }
            $ws = $wb.Worksheets.Item($r.sheet); $last = LastRow $ws; $hit = 0
            for ($i = 4; $i -le $last; $i++) {
                if ("$($ws.Cells.Item($i, 3).Text)".Trim() -eq $r.no -and "$($ws.Cells.Item($i, 2).Text)".Trim() -eq $r.type -and (-not $r.area -or "$($ws.Cells.Item($i, 9).Text)".Trim() -eq $r.area)) {
                    $ws.Cells.Item($i, 5).Value2 = Serial $r.closed; $hit++
                    "closed: $($r.no) $($r.type) row $i"
                }
            }
            if ($hit -ne 1) { throw "Permit $($r.no) $($r.type) $($r.area) matched $hit rows in $($r.sheet). Nothing written." }
        }
    }
    else {
        foreach ($sheet in ($rows | ForEach-Object { $_.sheet } | Select-Object -Unique)) {
            if ($names -notcontains $sheet) {
                $src = $wb.Worksheets.Item($wb.Worksheets.Count)
                $src.Copy([Type]::Missing, $src) | Out-Null
                $ws = $wb.Worksheets.Item($wb.Worksheets.Count)
                $ws.Name = $sheet
                if ($ws.FilterMode) { $ws.ShowAllData() | Out-Null }       # a filtered copy ignores ClearContents
                $last = LastRow $ws
                if ($last -ge 4) { $ws.Range("A4:N$last").ClearContents() | Out-Null }
                $names += $sheet
                "new sheet: $sheet (layout of $($src.Name))"
            }
            $ws = $wb.Worksheets.Item($sheet)
            $row = (LastRow $ws) + 1
            $sn = if ($row -gt 4) { [int]$ws.Cells.Item($row - 1, 1).Value2 } else { 0 }
            foreach ($r in ($rows | Where-Object { $_.sheet -eq $sheet })) {
                for ($i = 4; $i -lt $row; $i++) {
                    if ("$($ws.Cells.Item($i, 3).Text)".Trim() -eq $r.no -and "$($ws.Cells.Item($i, 2).Text)".Trim() -eq $r.type -and "$($ws.Cells.Item($i, 9).Text)".Trim() -eq $r.area) {
                        throw "Permit $($r.no) $($r.type) $($r.area) is already in $sheet (row $i). Nothing written."
                    }
                }
                if ($ws.Cells.Item($row, 2).Borders.Item(9).LineStyle -eq -4142 -and $row -gt 4) {
                    $ws.Rows.Item($row - 1).Copy() | Out-Null
                    $ws.Rows.Item($row).PasteSpecial(-4122) | Out-Null
                }
                $sn++
                $ws.Cells.Item($row, 1).Value2 = [string]$sn
                $ws.Cells.Item($row, 2).Value2 = [string]$r.type
                $ws.Cells.Item($row, 3).NumberFormat = "@"; $ws.Cells.Item($row, 3).Value2 = [string]$r.no
                $ws.Cells.Item($row, 4).NumberFormat = "dd/mm/yyyy"; $ws.Cells.Item($row, 4).Value2 = Serial $r.date
                $ws.Cells.Item($row, 5).NumberFormat = "dd/mm/yyyy"
                $ws.Cells.Item($row, 6).NumberFormat = "hh:mm"; $ws.Cells.Item($row, 6).Value2 = TimeVal $r.start
                $ws.Cells.Item($row, 7).NumberFormat = "hh:mm"; $ws.Cells.Item($row, 7).Value2 = TimeVal $r.end
                $ws.Cells.Item($row, 8).Value2 = [string]$r.company
                $ws.Cells.Item($row, 9).Value2 = [string]$r.area
                $ws.Cells.Item($row, 10).Value2 = [string]$r.description
                $ws.Cells.Item($row, 11).Value2 = [string]$r.issuer
                $ws.Cells.Item($row, 12).Value2 = [string]$r.receiver
                $ws.Cells.Item($row, 13).NumberFormat = "@"; $ws.Cells.Item($row, 13).Value2 = [string]$r.contact
                $ws.Cells.Item($row, 14).Value2 = [string]$r.hse
                "added: $sheet row $row  SN $sn  $($r.type) $($r.no) $($r.area)"
                $row++
            }
            $ws.Cells.Item(2, 3).Value2 = (Get-Date -Format "dd-MM-yyyy")
            $area = $ws.PageSetup.PrintArea
            if ($area -match '\$(\d+)$' -and [int]$Matches[1] -lt ($row - 1)) { $ws.PageSetup.PrintArea = $area -replace '\$\d+$', "`$$($row - 1)" }
        }
    }
    $wb.Save()
    $saved = $true
}
finally {
    if ($wb) { $wb.Close($false) | Out-Null }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect()
    "saved: $saved"
}
