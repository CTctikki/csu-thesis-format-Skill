param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,

    [Parameter(Mandatory = $true)]
    [string]$OutputDocx,

    [Parameter(Mandatory = $true)]
    [string]$ConfigJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $InputDocx)) {
    throw "Missing input DOCX: $InputDocx"
}

if (-not (Test-Path -LiteralPath $ConfigJson)) {
    throw "Missing config JSON: $ConfigJson"
}

$config = Get-Content -LiteralPath $ConfigJson -Raw | ConvertFrom-Json
Copy-Item -LiteralPath $InputDocx -Destination $OutputDocx -Force

$glyphMap = @{
    '[RHO]'   = [char]0x03C1
    '[TIMES]' = [char]0x00D7
    '[SUM]'   = [char]0x2211
    '[ALPHA]' = [char]0x03B1
    '[BETA]'  = [char]0x03B2
    '[GAMMA]' = [char]0x03B3
    '[ETA]'   = [char]0x03B7
    '[THETA]' = [char]0x03B8
    '[LAMBDA]'= [char]0x03BB
    '[MU]'    = [char]0x03BC
    '[PI]'    = [char]0x03C0
}

function Expand-FormulaGlyphs {
    param([string]$Formula)
    $expanded = $Formula
    foreach ($key in $glyphMap.Keys) {
        $expanded = $expanded.Replace($key, $glyphMap[$key])
    }
    return $expanded
}

function Clear-CellText {
    param($Cell)
    $range = $Cell.Range
    $range.End = $range.End - 1
    $range.Text = ''
}

function Set-CellText {
    param(
        $Cell,
        [string]$Text,
        [int]$Alignment = 2
    )
    $range = $Cell.Range
    $range.End = $range.End - 1
    $range.Text = $Text
    $range.ParagraphFormat.Alignment = $Alignment
    $range.ParagraphFormat.LeftIndent = 0
}

function Set-RangeTextAsEquation {
    param(
        $Cell,
        [string]$Text,
        [int]$Alignment,
        [double]$LeftIndentPoints
    )
    $range = $Cell.Range
    $range.End = $range.End - 1
    $range.Text = (Expand-FormulaGlyphs $Text)
    $range.ParagraphFormat.Alignment = $Alignment
    $range.ParagraphFormat.LeftIndent = $LeftIndentPoints
    $range.OMaths.Add($range) | Out-Null
    $range.OMaths.Item(1).BuildUp()
}

function Ensure-TableShape {
    param(
        $Word,
        $Table,
        [int]$RowsNeeded,
        [double]$LeftWidthCm,
        [double]$FormulaWidthCm,
        [double]$NumberWidthCm
    )
    while ($Table.Rows.Count -lt $RowsNeeded) {
        $Table.Rows.Add() | Out-Null
    }
    $Table.Borders.Enable = 0
    $Table.Rows.Alignment = 1
    foreach ($row in $Table.Rows) {
        $row.Cells.Item(1).Width = $Word.CentimetersToPoints($LeftWidthCm)
        $row.Cells.Item(2).Width = $Word.CentimetersToPoints($FormulaWidthCm)
        $row.Cells.Item(3).Width = $Word.CentimetersToPoints($NumberWidthCm)
    }
}

$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open($OutputDocx, $false, $false)

    $leftWidthCm = if ($config.tableWidthsCm.left) { [double]$config.tableWidthsCm.left } else { 1.2 }
    $formulaWidthCm = if ($config.tableWidthsCm.formula) { [double]$config.tableWidthsCm.formula } else { 12.2 }
    $numberWidthCm = if ($config.tableWidthsCm.number) { [double]$config.tableWidthsCm.number } else { 1.8 }

    foreach ($entry in $config.equations) {
        $table = $doc.Tables.Item([int]$entry.tableIndex)
        Ensure-TableShape $word $table $entry.rows.Count $leftWidthCm $formulaWidthCm $numberWidthCm
        for ($i = 0; $i -lt $entry.rows.Count; $i++) {
            $rowIndex = $i + 1
            $rowConfig = $entry.rows[$i]
            Clear-CellText $table.Cell($rowIndex, 1)
            $align = if ($rowConfig.align -eq 'left') { 0 } else { 1 }
            $hasIndent = $rowConfig.PSObject.Properties.Name -contains 'leftIndentCm'
            $leftIndentCm = if ($hasIndent -and $null -ne $rowConfig.leftIndentCm) { [double]$rowConfig.leftIndentCm } else { 0.0 }
            Set-RangeTextAsEquation $table.Cell($rowIndex, 2) $rowConfig.formula $align ($word.CentimetersToPoints($leftIndentCm))
            Set-CellText $table.Cell($rowIndex, 3) ([string]$rowConfig.number)
        }
    }

    $doc.Save()
    Write-Output "[OK] saved $OutputDocx"
}
finally {
    if ($doc -ne $null) { $doc.Close() | Out-Null }
    if ($word -ne $null) { $word.Quit() | Out-Null }
    if ($doc -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null }
    if ($word -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
