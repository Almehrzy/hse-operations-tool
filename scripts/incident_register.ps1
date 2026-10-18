# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
# Writes incident rows to the incident register REG-013, the minor incident log REG-014 or the CAPA register REG-011,
# through Excel itself (formatting, formulas and the dashboard are kept). Backs the workbook up to backups\ first.
# Stops without saving on any mismatch.
#
#   incident_register.ps1 -RowsJson rows.json
#
# rows.json:
#   {"register": "013" | "014" | "011",
#    "append": [ {"A": "6", "B": "date:2026-10-14", "C": "FAC", ...}, ... ],        new rows after the last data row
#    "update": [ {"key": "6", "set": {"L": "date:2026-10-17", "P": "Closed"}} ] }   rows found by column A
# Values are text; "date:YYYY-MM-DD" is written as a date in the register's own format. A new row takes the format of
# the row above it and is inserted, so a footer below the data (REG-011 "Prepared By") moves down.
param(
    [Parameter(Mandatory)][string]$RowsJson,
    [string]$BackupDir = "",
    [string]$Workbook = ""          # another copy of the register (tests)
)

$ErrorActionPreference = "Stop"
$toolRoot = Split-Path $PSScriptRoot -Parent
$cfg = Get-Content -LiteralPath (Join-Path $toolRoot "project.json") -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $BackupDir) { $BackupDir = Join-Path $toolRoot "backups" }
$req = Get-Content -LiteralPath $RowsJson -Raw -Encoding UTF8 | ConvertFrom-Json

$spec = @{
    "013" = @{ path = $cfg.incidents.register_013; sheet = "Incident Details"; first = 3; lastCol = 16; fmt = "dd/mm/yyyy" }
    "014" = @{ path = $cfg.incidents.register_014; sheet = "Incident Register"; first = 6; lastCol = 28; fmt = "dd-mm-yyyy" }
    "011" = @{ path = $cfg.incidents.capa_register; sheet = "CAPA Register"; first = 4; lastCol = 13; fmt = "dd-mm-yyyy" }
}[[string]$req.register]
if (-not $spec) { throw "register must be 013, 014 or 011" }
if ($Workbook) { $spec.path = $Workbook }
$appends = @($req.append | Where-Object { $_ })
$updates = @($req.update | Where-Object { $_ })
if (($appends.Count + $updates.Count) -eq 0) { throw "Nothing to write in $RowsJson" }

function Col([string]$letters) { $n = 0; foreach ($ch in $letters.ToUpper().ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }; return $n }
function Put($ws, [int]$row, [string]$letters, $v) {
    $c = $ws.Cells.Item($row, (Col $letters))
    $s = [string]$v
    if ($s -match '^date:(\d{4}-\d{2}-\d{2})$') {
        $c.NumberFormat = $spec.fmt
        # a COM double does not cast into Value2 here; Excel turns the number text into the date serial
        $c.Value2 = [string][datetime]::ParseExact($Matches[1], "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture).ToOADate()
    } else { $c.Value2 = $s }
}

New-Item -ItemType Directory -Force $BackupDir | Out-Null
$base = [IO.Path]::GetFileNameWithoutExtension($spec.path)
$hash = (Get-FileHash -LiteralPath $spec.path).Hash
$existing = Get-ChildItem -LiteralPath $BackupDir -Filter "$base - backup *.xlsx" -ErrorAction SilentlyContinue | Where-Object { (Get-FileHash $_.FullName).Hash -eq $hash } | Select-Object -First 1
if ($existing) { "backup already holds the current register: $($existing.Name)" }
else {
    $backup = Join-Path $BackupDir ("$base - backup {0}.xlsx" -f (Get-Date -Format "yyyy-MM-dd HHmmss"))
    Copy-Item -LiteralPath $spec.path -Destination $backup
    "backup: $backup"
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
$saved = $false
try {
    $wb = $xl.Workbooks.Open($spec.path)
    if ($wb.ReadOnly) { throw "Register opened read-only (open on another computer?). Nothing written." }
    $ws = $wb.Worksheets.Item($spec.sheet)
    $used = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
    # last data row: the last row with something in columns A..lastCol, before a "Prepared by" footer
    $last = $spec.first - 1
    for ($r = $spec.first; $r -le $used; $r++) {
        $a = [string]$ws.Cells.Item($r, 1).Text
        if ($a -match '^(Pr[e]?p[a]?r?ed|Prpeared) By') { break }
        for ($c = 1; $c -le $spec.lastCol; $c++) { if ([string]$ws.Cells.Item($r, $c).Text -ne "") { $last = $r; break } }
    }
    # updates: find each key in column A; check before writing
    $targets = @()
    foreach ($u in $updates) {
        $row = 0
        for ($r = $spec.first; $r -le $last; $r++) { if (([string]$ws.Cells.Item($r, 1).Text).Trim() -eq ([string]$u.key).Trim()) { $row = $r; break } }
        if (-not $row) { throw "Row '$($u.key)' not found in column A. Nothing written." }
        $targets += , @($row, $u.set)
    }
    foreach ($a in $appends) {
        $k = [string]$a.A
        if ($k) { for ($r = $spec.first; $r -le $last; $r++) { if (([string]$ws.Cells.Item($r, 1).Text).Trim() -eq $k.Trim()) { throw "'$k' is already in the register (row $r). Nothing written." } } }
    }
    foreach ($t in $targets) { foreach ($p in $t[1].PSObject.Properties) { Put $ws $t[0] $p.Name $p.Value } }
    $row = $last
    foreach ($a in $appends) {
        $row++
        $ws.Rows.Item($row).Insert() | Out-Null
        $ws.Rows.Item($row - 1).Copy() | Out-Null
        $ws.Rows.Item($row).PasteSpecial(-4122) | Out-Null           # formats only
        foreach ($p in $a.PSObject.Properties) { Put $ws $row $p.Name $p.Value }
        $ws.Rows.Item($row).AutoFit() | Out-Null
    }
    $wb.Save()
    $saved = $true
    "register $($req.register): updated $($targets.Count) row(s), added $($appends.Count) row(s) after row $last"
}
finally {
    if ($wb) { $wb.Close($false) | Out-Null }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect()
    "saved: $saved"
}
