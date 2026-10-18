# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
# Updates or adds equipment rows in the TUV register (REG-008, sheet "Equipment TUV") through Excel itself,
# so logos and formatting are kept. Backs the register up first. Stops without saving on any mismatch.
#
#   tuv_register.ps1 -ChangesJson changes.json
#
# changes.json (made by equipment.py regplan): a list of
#   {"no": 76, "new": false, "set": {"D": "SK10638", "F": "4951 ZAA", "G": "SMAT", "H": "INJO-...", "I": "2026-09-14", "J": "2027-09-13", "L": ""}}
# "set" maps register columns to values: B type, C serial, D sticker, E make, F plate, G certifying co., H certificate no.,
# I inspection date, J expiry date (both YYYY-MM-DD), K company, L remarks ("OFF SITE" or "").
# An existing row is found by its No. in column A; a new row must take the next No. after the last one.
param(
    [Parameter(Mandatory)][string]$ChangesJson,
    [string]$Register = "",
    [string]$BackupDir = ""
)

$ErrorActionPreference = "Stop"
$toolRoot = Split-Path $PSScriptRoot -Parent
$eq = Get-Content -LiteralPath (Join-Path $toolRoot "equipment.json") -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $Register) { $Register = $eq.register }
if (-not $BackupDir) { $BackupDir = Join-Path $toolRoot "backups" }
$sheetName = $eq.register_sheet
$changes = @(Get-Content -LiteralPath $ChangesJson -Raw -Encoding UTF8 | ConvertFrom-Json | ForEach-Object { $_ })   # unroll: PS 5.1 emits a JSON array as one object
if ($changes.Count -eq 0) { throw "No changes in $ChangesJson" }
$cols = @{ B = 2; C = 3; D = 4; E = 5; F = 6; G = 7; H = 8; I = 9; J = 10; K = 11; L = 12 }

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

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
$saved = $false
try {
    $wb = $xl.Workbooks.Open($Register)
    if ($wb.ReadOnly) { throw "Register opened read-only (open on another computer?). Nothing written." }
    $ws = $wb.Worksheets.Item($sheetName)

    # Rows by No. in column A. The register carries pre-numbered empty rows after the last entry, so the last
    # used No. is the last one with an equipment type in column B; a new entry fills its pre-numbered row.
    $rowOf = @{}; $lastRow = 5; $lastNo = 0; $lastUsed = 0
    $end = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
    for ($r = 6; $r -le $end; $r++) {
        $a = $ws.Cells.Item($r, 1).Value2
        if ($null -ne $a -and "$a" -match '^\d+(\.0+)?$') {
            $n = [int]$a; $rowOf[$n] = $r
            if ($n -gt $lastNo) { $lastNo = $n; $lastRow = $r }
            if ("$($ws.Cells.Item($r, 2).Value2)".Trim() -ne "" -and $n -gt $lastUsed) { $lastUsed = $n }
        }
    }

    # Pass 1: check.
    $next = $lastUsed
    foreach ($c in ($changes | Sort-Object { [int]$_.no })) {
        $n = [int]$c.no
        if ($c.new) {
            if ($rowOf.ContainsKey($n) -and "$($ws.Cells.Item($rowOf[$n], 2).Value2)".Trim() -ne "") { throw "No. $n already holds equipment in the register. Nothing written." }
            if ($n -ne $next + 1) { throw "New equipment No. $n does not follow the last used No. $next. Nothing written." }
            $next = $n
        }
        elseif (-not $rowOf.ContainsKey($n)) { throw "No. $n is not in the register. Nothing written." }
        foreach ($k in $c.set.PSObject.Properties.Name) { if (-not $cols.ContainsKey($k)) { throw "Unknown column $k for No. $n. Nothing written." } }
    }

    # Pass 2: write.
    foreach ($c in ($changes | Sort-Object { [int]$_.no })) {
        $n = [int]$c.no
        if ($c.new -and $rowOf.ContainsKey($n)) { $row = $rowOf[$n] }
        elseif ($c.new) {
            $row = $lastRow + 1
            $ws.Rows.Item($lastRow).Copy() | Out-Null
            $ws.Rows.Item($row).PasteSpecial(-4122) | Out-Null           # formats only
            $ws.Cells.Item($row, 1).Value2 = $n
            $lastRow = $row; $rowOf[$n] = $row
        }
        else { $row = $rowOf[$n] }
        $log = @()
        foreach ($p in $c.set.PSObject.Properties) {
            $col = $cols[$p.Name]; $cell = $ws.Cells.Item($row, $col); $old = [string]$cell.Text; $v = [string]$p.Value
            if ($p.Name -in @("I", "J") -and $v -match '^\d{4}-\d{2}-\d{2}$') {
                $d = [datetime]::ParseExact($v, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture)
                if ([string]$cell.NumberFormat -in @("General", "@")) { $cell.NumberFormat = "dd-mmm-yy" }   # the register's own date style
                $cell.Value2 = [string][int]$d.ToOADate()   # PS COM binder rejects a double here; Excel reads the serial as a number
            }
            elseif ($v -eq "") { $cell.ClearContents() | Out-Null }
            else { $cell.NumberFormat = "@"; $cell.Value2 = $v }
            $now = [string]$cell.Text; if ($old -ne $now) { $log += "$($p.Name): '$old' -> '$now'" }
        }
        "No. $n (row $row)$(if ($c.new) { ' NEW' }): $($log -join '; ')"
    }
    $area = $ws.PageSetup.PrintArea
    if ($area -match '\$(\d+)$' -and [int]$Matches[1] -lt $lastRow) { $ws.PageSetup.PrintArea = $area -replace '\$\d+$', "`$$lastRow" }
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
