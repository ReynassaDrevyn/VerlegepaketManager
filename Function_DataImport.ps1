# Function_DataImport.ps1
# Windows-PowerShell-5.1-kompatible Importversion

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$ProjectRoot = $(if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path $MyInvocation.MyCommand.Path -Parent })
$DbPath = Join-Path $ProjectRoot 'Core\db_verlegepaket.json'
$LookupPath = Join-Path $ProjectRoot 'Core\db_lookups.json'
$DataImportPresetPath = Join-Path $ProjectRoot 'Core\data_import_presets.json'
$LogsDir = Join-Path $ProjectRoot 'Logs'
$BackupDir = Join-Path $LogsDir 'Backups'
$LogFile = Join-Path $LogsDir "InitialImport_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$DatabaseSchemaVersion = 2
$DataImportPresetSchemaVersion = 1

if (!(Test-Path $LogsDir)) { New-Item -Path $LogsDir -ItemType Directory -Force | Out-Null }
if (!(Test-Path $BackupDir)) { New-Item -Path $BackupDir -ItemType Directory -Force | Out-Null }
$dbDir = Split-Path $DbPath -Parent
if (!(Test-Path $dbDir)) { New-Item -Path $dbDir -ItemType Directory -Force | Out-Null }

$global:LogBox = $null

function Write-ImportLog {
    param([string]$Message, [string]$Level = 'INFO')

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logLine = "$timestamp [$Level] $Message"

    Add-Content -Path $LogFile -Value $logLine -Encoding UTF8 -ErrorAction SilentlyContinue
    if ($global:LogBox) {
        $global:LogBox.AppendText("$logLine`r`n")
        if ($global:LogBox -is [System.Windows.Controls.TextBox]) {
            $global:LogBox.ScrollToEnd()
        }
        else {
            $global:LogBox.ScrollToCaret()
        }
    }
}

function Get-NormalizedString {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return ''
    }

    return ([string]$Value).Trim()
}

function ConvertTo-NullableString {
    param([AllowNull()][object]$Value)

    $text = Get-NormalizedString $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    return $text
}

function ConvertTo-ObjectArray {
    param([AllowNull()]$InputObject)

    if ($null -eq $InputObject) {
        return @()
    }

    return @($InputObject)
}

function Copy-DeepObject {
    param([AllowNull()]$InputObject)

    if ($null -eq $InputObject) {
        return $null
    }

    return (($InputObject | ConvertTo-Json -Depth 20) | ConvertFrom-Json)
}

function ConvertTo-NormalizedHeaderName {
    param([string]$Header)

    $normalized = Get-NormalizedString $Header
    $normalized = $normalized.ToLowerInvariant()
    $normalized = $normalized.Replace(([string][char]0x00E4), 'ae')
    $normalized = $normalized.Replace(([string][char]0x00F6), 'oe')
    $normalized = $normalized.Replace(([string][char]0x00FC), 'ue')
    $normalized = $normalized.Replace(([string][char]0x00DF), 'ss')
    $normalized = $normalized.Replace(([string][char]0x00C4), 'ae')
    $normalized = $normalized.Replace(([string][char]0x00D6), 'oe')
    $normalized = $normalized.Replace(([string][char]0x00DC), 'ue')
    $normalized = $normalized -replace '[^a-z0-9]+', ' '
    $normalized = $normalized -replace '\s+', ' '

    return $normalized.Trim()
}

function Get-HeaderDefinitions {
    return @(
        [pscustomobject]@{ Key = 'import_id'; Required = $false; Patterns = @('^id$') }
        [pscustomobject]@{ Key = 'matnr_main'; Required = $true; Patterns = @('^materialnummer.*saspf$', '^matnr.*saspf$', '^materialnummer$') }
        [pscustomobject]@{ Key = 'supplynumber'; Required = $false; Patterns = @('^versnr$', '^vers nr$', '^versorgungsnummer$') }
        [pscustomobject]@{ Key = 'mat_stat_main'; Required = $false; Patterns = @('^status.*matnr$', '^status.*materialnummer$') }
        [pscustomobject]@{ Key = 'dezentral'; Required = $false; Patterns = @('^dezent$', '^dezentral$') }
        [pscustomobject]@{ Key = 'ext_wg'; Required = $false; Patterns = @('^ext wg$', '^ext warengruppe$') }
        [pscustomobject]@{ Key = 'artnr'; Required = $false; Patterns = @('^artikel nr$', '^artikelnr$', '^art nr$', '^artnr$') }
        [pscustomobject]@{ Key = 'description'; Required = $false; Patterns = @('^materialbezeichnung$', '^bezeichnung$', '^description$') }
        [pscustomobject]@{ Key = 'technical'; Required = $false; Patterns = @('^bezeichnung technik$', '^technik$') }
        [pscustomobject]@{ Key = 'logistics'; Required = $false; Patterns = @('^bemerkung$', '^logistik$', '^logistics$') }
        [pscustomobject]@{ Key = 'unit_main'; Required = $false; Patterns = @('^bze$', '^einheit$') }
        [pscustomobject]@{ Key = 'quantity_target'; Required = $false; Patterns = @('^tlg 74$', '^menge$', '^quantity$') }
        [pscustomobject]@{ Key = 'is_dg'; Required = $false; Patterns = @('^gefstoff$', '^dg$') }
        [pscustomobject]@{ Key = 'GefStoff Verlegung'; Required = $false; Patterns = @('^gefstoff verlegung$') }
        [pscustomobject]@{ Key = 'Gefahrgut'; Required = $false; Patterns = @('^gefahrgut$', '^gefahr gut$') }
        [pscustomobject]@{ Key = 'Batterie'; Required = $false; Patterns = @('^batterie$') }
        [pscustomobject]@{ Key = 'Flight'; Required = $false; Patterns = @('^flight$') }
        [pscustomobject]@{ Key = 'Waffen'; Required = $false; Patterns = @('^waffen$') }
        [pscustomobject]@{ Key = 'Munition'; Required = $false; Patterns = @('^munition$') }
        [pscustomobject]@{ Key = 'RTS'; Required = $false; Patterns = @('^rts$') }
        [pscustomobject]@{ Key = 'AUG'; Required = $false; Patterns = @('^aug$') }
        [pscustomobject]@{ Key = 'WEF'; Required = $false; Patterns = @('^wef$') }
        [pscustomobject]@{ Key = 'BoGe'; Required = $false; Patterns = @('^boge$') }
        [pscustomobject]@{ Key = 'HFT'; Required = $false; Patterns = @('^hft$') }
        [pscustomobject]@{ Key = 'LME'; Required = $false; Patterns = @('^lme$') }
        [pscustomobject]@{ Key = 'REG'; Required = $false; Patterns = @('^reg$') }
        [pscustomobject]@{ Key = 'RNW'; Required = $false; Patterns = @('^rnw$') }
        [pscustomobject]@{ Key = 'Rad Reifen Shop'; Required = $false; Patterns = @('^rad reifen shop$') }
        [pscustomobject]@{ Key = 'IETPX Material'; Required = $false; Patterns = @('^ietpx material$') }
        [pscustomobject]@{ Key = 'GUN ON AC'; Required = $false; Patterns = @('^gun on ac$') }
        [pscustomobject]@{ Key = 'GUN OFF AC'; Required = $false; Patterns = @('^gun off ac$') }
        [pscustomobject]@{ Key = 'GUN'; Required = $false; Patterns = @('^gun$') }
        [pscustomobject]@{ Key = 'IRIS-T'; Required = $false; Patterns = @('^iris t$') }
        [pscustomobject]@{ Key = 'FLARE'; Required = $false; Patterns = @('^flare$') }
        [pscustomobject]@{ Key = 'AIM 120'; Required = $false; Patterns = @('^aim 120$', '^aim120$') }
        [pscustomobject]@{ Key = '1000 l SFT'; Required = $false; Patterns = @('^1000 l sft$', '^1000l sft$') }
        [pscustomobject]@{ Key = 'GBU 48'; Required = $false; Patterns = @('^gbu 48$', '^gbu48$') }
        [pscustomobject]@{ Key = 'Meteor'; Required = $false; Patterns = @('^meteor$') }
        [pscustomobject]@{ Key = 'LDP'; Required = $false; Patterns = @('^ldp$') }
        [pscustomobject]@{ Key = 'IWP'; Required = $false; Patterns = @('^iwp$') }
        [pscustomobject]@{ Key = 'CFP'; Required = $false; Patterns = @('^cfp$') }
        [pscustomobject]@{ Key = 'MFRL'; Required = $false; Patterns = @('^mfrl$') }
        [pscustomobject]@{ Key = 'OWP'; Required = $false; Patterns = @('^owp$') }
        [pscustomobject]@{ Key = 'CHAFF'; Required = $false; Patterns = @('^chaff$') }
        [pscustomobject]@{ Key = 'MEL'; Required = $false; Patterns = @('^mel$') }
        [pscustomobject]@{ Key = 'ITSPL'; Required = $false; Patterns = @('^itspl$') }
    )
}

function Resolve-HeaderMap {
    param(
        [Parameter(Mandatory = $true)][string[]]$Headers,
        [Parameter(Mandatory = $true)]$Definitions
    )

    $resolved = @{}
    $normalizedHeaders = @{}

    foreach ($header in $Headers) {
        $normalizedHeaders[$header] = ConvertTo-NormalizedHeaderName $header
    }

    foreach ($definition in $Definitions) {
        foreach ($header in $Headers) {
            $normalized = $normalizedHeaders[$header]
            foreach ($pattern in $definition.Patterns) {
                if ($normalized -match $pattern) {
                    if (-not $resolved.ContainsKey($definition.Key)) {
                        $resolved[$definition.Key] = $header
                    }
                    break
                }
            }

            if ($resolved.ContainsKey($definition.Key)) {
                break
            }
        }
    }

    $missingRequired = @()
    $missingOptional = @()
    foreach ($definition in $Definitions) {
        if (-not $resolved.ContainsKey($definition.Key)) {
            if ($definition.Required) {
                $missingRequired += $definition.Key
            }
            else {
                $missingOptional += $definition.Key
            }
        }
    }

    return [pscustomobject]@{
        HeaderMap       = $resolved
        MissingRequired = $missingRequired
        MissingOptional = $missingOptional
        Normalized      = $normalizedHeaders
    }
}

function Get-CellValue {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][hashtable]$HeaderMap,
        [Parameter(Mandatory = $true)][string]$Key
    )

    if (-not $HeaderMap.ContainsKey($Key)) {
        return ''
    }

    $headerName = $HeaderMap[$Key]
    $property = $Row.PSObject.Properties[$headerName]
    if ($null -eq $property) {
        return ''
    }

    return Get-NormalizedString $property.Value
}

function Test-HeaderAvailable {
    param(
        [Parameter(Mandatory = $true)][hashtable]$HeaderMap,
        [Parameter(Mandatory = $true)][string]$Key
    )

    return $HeaderMap.ContainsKey($Key)
}

function Test-HasUtf8Bom {
    param([byte[]]$Bytes)

    return $Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF
}

function Read-TextFileWithEncodingFallback {
    param([Parameter(Mandatory = $true)][string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $encoding = $null
    $text = ''

    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        $encoding = [System.Text.Encoding]::Unicode
        $text = $encoding.GetString($bytes, 2, $bytes.Length - 2)
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        $encoding = [System.Text.Encoding]::BigEndianUnicode
        $text = $encoding.GetString($bytes, 2, $bytes.Length - 2)
    }
    elseif (Test-HasUtf8Bom -Bytes $bytes) {
        $encoding = New-Object System.Text.UTF8Encoding($true)
        $text = $encoding.GetString($bytes, 3, $bytes.Length - 3)
    }
    else {
        try {
            $utf8NoBom = New-Object System.Text.UTF8Encoding($false, $true)
            $text = $utf8NoBom.GetString($bytes)
            $encoding = $utf8NoBom
        }
        catch {
            $encoding = [System.Text.Encoding]::Default
            $text = $encoding.GetString($bytes)
        }
    }

    return [pscustomobject]@{
        Text         = $text
        EncodingName = $encoding.WebName
    }
}

function Import-DelimitedCsvText {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [string]$Delimiter = ';'
    )

    return @($Text | ConvertFrom-Csv -Delimiter $Delimiter)
}

function Import-CsvData {
    param([Parameter(Mandatory = $true)][string]$Path)

    $fileData = Read-TextFileWithEncodingFallback -Path $Path
    $rows = Import-DelimitedCsvText -Text $fileData.Text

    return [pscustomobject]@{
        Rows         = $rows
        EncodingName = $fileData.EncodingName
    }
}

function Get-ImportFileTypeFromPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $extension = [System.IO.Path]::GetExtension($Path)
    if ([string]::IsNullOrWhiteSpace($extension)) {
        return ''
    }

    return $extension.TrimStart('.').ToLowerInvariant()
}

function Import-DelimitedTextWithHeaderRow {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [int]$HeaderRowIndex = 1,
        [string]$Delimiter = ';'
    )

    if ($HeaderRowIndex -lt 1) {
        throw 'HeaderRowIndex muss groesser oder gleich 1 sein.'
    }

    $lines = $Text -split "(`r`n|`n|`r)"
    if ($HeaderRowIndex -gt $lines.Count) {
        throw "HeaderRowIndex $HeaderRowIndex liegt ausserhalb der Datei."
    }

    $selectedLines = New-Object System.Collections.Generic.List[string]
    for ($lineIndex = $HeaderRowIndex - 1; $lineIndex -lt $lines.Count; $lineIndex++) {
        [void]$selectedLines.Add($lines[$lineIndex])
    }

    $trimmedLines = @($selectedLines.ToArray())
    while ($trimmedLines.Count -gt 0 -and [string]::IsNullOrWhiteSpace($trimmedLines[$trimmedLines.Count - 1])) {
        if ($trimmedLines.Count -eq 1) {
            $trimmedLines = @()
        }
        else {
            $trimmedLines = @($trimmedLines[0..($trimmedLines.Count - 2)])
        }
    }

    if ($trimmedLines.Count -lt 1) {
        return [pscustomobject]@{
            Headers      = @()
            Rows         = @()
            RowNumbers   = @()
            Worksheet    = $null
            EncodingName = $null
        }
    }

    $dataText = $trimmedLines -join [Environment]::NewLine
    $rows = if ($trimmedLines.Count -gt 1) { @($dataText | ConvertFrom-Csv -Delimiter $Delimiter) } else { @() }
    $headers = if ($rows.Count -gt 0) {
        @($rows[0].PSObject.Properties.Name)
    }
    else {
        @($trimmedLines[0] -split [regex]::Escape($Delimiter))
    }

    return [pscustomobject]@{
        Headers      = @($headers)
        Rows         = @($rows)
        RowNumbers   = @(for ($i = 0; $i -lt @($rows).Count; $i++) { $HeaderRowIndex + 1 + $i })
        Worksheet    = $null
        EncodingName = $null
    }
}

function Read-DelimitedImportSource {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$HeaderRowIndex = 1,
        [string]$Delimiter = ';'
    )

    $fileData = Read-TextFileWithEncodingFallback -Path $Path
    $result = Import-DelimitedTextWithHeaderRow -Text $fileData.Text -HeaderRowIndex $HeaderRowIndex -Delimiter $Delimiter
    $result | Add-Member -NotePropertyName EncodingName -NotePropertyValue $fileData.EncodingName -Force
    return $result
}

function Read-ExcelImportSource {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$HeaderRowIndex = 1,
        [AllowNull()][string]$WorksheetName
    )

    if ($HeaderRowIndex -lt 1) {
        throw 'HeaderRowIndex muss groesser oder gleich 1 sein.'
    }

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $usedRange = $null

    try {
        try {
            $excel = New-Object -ComObject Excel.Application
        }
        catch {
            throw "Excel COM ist fuer XLSX-Importe nicht verfuegbar. Bitte Excel installieren oder CSV/TXT verwenden."
        }

        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($Path, 0, $true)

        if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $WorksheetName))) {
            $worksheet = $workbook.Worksheets.Item(1)
        }
        else {
            try {
                $worksheet = $workbook.Worksheets.Item($WorksheetName)
            }
            catch {
                throw "Worksheet '$WorksheetName' wurde in '$Path' nicht gefunden."
            }
        }

        $usedRange = $worksheet.UsedRange
        $firstUsedRow = [int]$usedRange.Row
        $firstUsedColumn = [int]$usedRange.Column
        $lastUsedRow = $firstUsedRow + [int]$usedRange.Rows.Count - 1
        $lastUsedColumn = $firstUsedColumn + [int]$usedRange.Columns.Count - 1

        if ($HeaderRowIndex -gt $lastUsedRow) {
            throw "HeaderRowIndex $HeaderRowIndex liegt ausserhalb des Worksheets '$($worksheet.Name)'."
        }

        $headers = New-Object System.Collections.Generic.List[string]
        for ($columnIndex = $firstUsedColumn; $columnIndex -le $lastUsedColumn; $columnIndex++) {
            [void]$headers.Add((Get-NormalizedString $worksheet.Cells.Item($HeaderRowIndex, $columnIndex).Text))
        }

        $rows = New-Object System.Collections.Generic.List[object]
        $rowNumbers = New-Object System.Collections.Generic.List[int]
        for ($rowIndex = $HeaderRowIndex + 1; $rowIndex -le $lastUsedRow; $rowIndex++) {
            $ordered = [ordered]@{}
            $hasData = $false

            for ($columnIndex = $firstUsedColumn; $columnIndex -le $lastUsedColumn; $columnIndex++) {
                $header = $headers[$columnIndex - $firstUsedColumn]
                if ([string]::IsNullOrWhiteSpace($header)) {
                    continue
                }

                $cellText = Get-NormalizedString $worksheet.Cells.Item($rowIndex, $columnIndex).Text
                if (-not [string]::IsNullOrWhiteSpace($cellText)) {
                    $hasData = $true
                }

                $ordered[$header] = $cellText
            }

            if ($hasData) {
                [void]$rows.Add([pscustomobject]$ordered)
                [void]$rowNumbers.Add($rowIndex)
            }
        }

        return [pscustomobject]@{
            Headers      = @($headers.ToArray() | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            Rows         = @($rows.ToArray())
            RowNumbers   = @($rowNumbers.ToArray())
            Worksheet    = $worksheet.Name
            EncodingName = $null
        }
    }
    finally {
        if ($null -ne $usedRange) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange)
        }

        if ($null -ne $worksheet) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet)
        }

        if ($null -ne $workbook) {
            $workbook.Close($false)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        }

        if ($null -ne $excel) {
            $excel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Read-GenericImportSource {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$FileType,
        [int]$HeaderRowIndex = 1,
        [AllowNull()][string]$WorksheetName,
        [string]$Delimiter = ';'
    )

    $resolvedFileType = Get-NormalizedString $FileType
    if ([string]::IsNullOrWhiteSpace($resolvedFileType)) {
        $resolvedFileType = Get-ImportFileTypeFromPath -Path $Path
    }

    $resolvedFileType = $resolvedFileType.ToLowerInvariant()
    switch ($resolvedFileType) {
        'xlsx' {
            return (Read-ExcelImportSource -Path $Path -HeaderRowIndex $HeaderRowIndex -WorksheetName $WorksheetName)
        }
        'csv' {
            return (Read-DelimitedImportSource -Path $Path -HeaderRowIndex $HeaderRowIndex -Delimiter $Delimiter)
        }
        'txt' {
            return (Read-DelimitedImportSource -Path $Path -HeaderRowIndex $HeaderRowIndex -Delimiter $Delimiter)
        }
        default {
            throw "Dateityp '$resolvedFileType' wird nicht unterstuetzt. Erwartet: csv, txt oder xlsx."
        }
    }
}

function ConvertTo-ImportBooleanParseResult {
    param([AllowNull()][string]$Value)

    $text = Get-NormalizedString $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return [pscustomobject]@{ Success = $true; Value = $false; IsBlank = $true }
    }

    $normalized = $text.ToLowerInvariant()
    if (@('x', '1', 'true', 'ja', 'yes', 'y', 'j') -contains $normalized) {
        return [pscustomobject]@{ Success = $true; Value = $true; IsBlank = $false }
    }

    if (@('0', 'false', 'nein', 'no', 'n') -contains $normalized) {
        return [pscustomobject]@{ Success = $true; Value = $false; IsBlank = $false }
    }

    return [pscustomobject]@{ Success = $false; Value = $false; IsBlank = $false }
}

function Test-Flag {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][hashtable]$HeaderMap,
        [Parameter(Mandatory = $true)][string]$Key
    )

    $parsed = ConvertTo-ImportBooleanParseResult (Get-CellValue -Row $Row -HeaderMap $HeaderMap -Key $Key)
    return $parsed.Success -and $parsed.Value
}

function ConvertTo-ImportNumberParseResult {
    param([AllowNull()][string]$Value)

    $text = Get-NormalizedString $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return [pscustomobject]@{ Success = $true; Value = 0.0; IsBlank = $true }
    }

    $parsed = 0.0
    $deCulture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
    if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Any, $deCulture, [ref]$parsed)) {
        return [pscustomobject]@{ Success = $true; Value = $parsed; IsBlank = $false }
    }

    $invariantCulture = [System.Globalization.CultureInfo]::InvariantCulture
    if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Any, $invariantCulture, [ref]$parsed)) {
        return [pscustomobject]@{ Success = $true; Value = $parsed; IsBlank = $false }
    }

    return [pscustomobject]@{ Success = $false; Value = 0.0; IsBlank = $false }
}
function Get-ResponsibilityDefinitions {
    return @(
        [pscustomobject]@{ Key = 'Flight'; Code = 'flight'; Label = 'Flight' }
        [pscustomobject]@{ Key = 'Waffen'; Code = 'waffen'; Label = 'Waffen' }
        [pscustomobject]@{ Key = 'Munition'; Code = 'munition'; Label = 'Munition' }
        [pscustomobject]@{ Key = 'RTS'; Code = 'rts'; Label = 'RTS' }
        [pscustomobject]@{ Key = 'AUG'; Code = 'aug'; Label = 'AUG' }
        [pscustomobject]@{ Key = 'WEF'; Code = 'wef'; Label = 'WEF' }
        [pscustomobject]@{ Key = 'BoGe'; Code = 'boge'; Label = 'BoGe' }
        [pscustomobject]@{ Key = 'HFT'; Code = 'hft'; Label = 'HFT' }
        [pscustomobject]@{ Key = 'LME'; Code = 'lme'; Label = 'LME' }
        [pscustomobject]@{ Key = 'REG'; Code = 'reg'; Label = 'REG' }
        [pscustomobject]@{ Key = 'RNW'; Code = 'rnw'; Label = 'RNW' }
        [pscustomobject]@{ Key = 'Rad Reifen Shop'; Code = 'rad_reifen'; Label = 'RadReifen' }
    )
}

function Get-AssignmentTagDefinitions {
    return @(
        [pscustomobject]@{ Key = 'IETPX Material'; Code = 'ietpx_material'; Label = 'IETPX Material' }
        [pscustomobject]@{ Key = 'GUN'; Code = 'gun'; Label = 'GUN' }
        [pscustomobject]@{ Key = 'GUN ON AC'; Code = 'gun_on_ac'; Label = 'GUN ON AC' }
        [pscustomobject]@{ Key = 'GUN OFF AC'; Code = 'gun_off_ac'; Label = 'GUN OFF AC' }
        [pscustomobject]@{ Key = 'IRIS-T'; Code = 'iris_t'; Label = 'IRIS-T' }
        [pscustomobject]@{ Key = 'FLARE'; Code = 'flare'; Label = 'FLARE' }
        [pscustomobject]@{ Key = 'AIM 120'; Code = 'aim_120'; Label = 'AIM 120' }
        [pscustomobject]@{ Key = '1000 l SFT'; Code = 'sft_1000_l'; Label = '1000 l SFT' }
        [pscustomobject]@{ Key = 'GBU 48'; Code = 'gbu_48'; Label = 'GBU 48' }
        [pscustomobject]@{ Key = 'Meteor'; Code = 'meteor'; Label = 'Meteor' }
        [pscustomobject]@{ Key = 'LDP'; Code = 'ldp'; Label = 'LDP' }
        [pscustomobject]@{ Key = 'IWP'; Code = 'iwp'; Label = 'IWP' }
        [pscustomobject]@{ Key = 'CFP'; Code = 'cfp'; Label = 'CFP' }
        [pscustomobject]@{ Key = 'MFRL'; Code = 'mfrl'; Label = 'MFRL' }
        [pscustomobject]@{ Key = 'OWP'; Code = 'owp'; Label = 'OWP' }
        [pscustomobject]@{ Key = 'CHAFF'; Code = 'chaff'; Label = 'CHAFF' }
        [pscustomobject]@{ Key = 'MEL'; Code = 'mel'; Label = 'MEL' }
        [pscustomobject]@{ Key = 'ITSPL'; Code = 'itspl'; Label = 'ITSPL' }
    )
}

function Get-HazmatFlagDefinitions {
    return @(
        [pscustomobject]@{ Key = 'GefStoff Verlegung'; Code = 'gefstoff_verlegung'; Label = 'GefStoff Verlegung' }
        [pscustomobject]@{ Key = 'Gefahrgut'; Code = 'gefahrgut'; Label = 'Gefahrgut' }
        [pscustomobject]@{ Key = 'Batterie'; Code = 'batterie'; Label = 'Batterie' }
    )
}

function Get-CanonicalIdentifierValue {
    param([string]$Value)

    $normalized = Get-NormalizedString $Value
    $normalized = $normalized.ToLowerInvariant()
    $normalized = $normalized -replace '\s+', ' '
    return $normalized.Trim()
}

function Get-IdentifierTypeFromValue {
    param([string]$Value)

    return 'matnr'
}

function Get-CanonicalKey {
    param(
        [Parameter(Mandatory = $true)][string]$Type,
        [Parameter(Mandatory = $true)][string]$Value
    )

    return '{0}:{1}' -f $Type, (Get-CanonicalIdentifierValue $Value)
}

function ConvertTo-UniqueStringArray {
    param([AllowNull()][object[]]$Values)

    $seen = @{}
    $result = New-Object System.Collections.Generic.List[string]
    foreach ($value in @($Values)) {
        $text = Get-NormalizedString $value
        if ([string]::IsNullOrWhiteSpace($text)) {
            continue
        }

        if (-not $seen.ContainsKey($text)) {
            $seen[$text] = $true
            [void]$result.Add($text)
        }
    }

    return @($result)
}

function Merge-DefinedTagValues {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][hashtable]$HeaderMap,
        [Parameter(Mandatory = $true)][object[]]$Definitions,
        [AllowNull()][string[]]$ExistingValues
    )

    $existingSet = @{}
    foreach ($existingValue in (ConvertTo-UniqueStringArray $ExistingValues)) {
        $existingSet[$existingValue] = $true
    }

    $merged = New-Object System.Collections.Generic.List[string]
    foreach ($definition in $Definitions) {
        if (Test-HeaderAvailable -HeaderMap $HeaderMap -Key $definition.Key) {
            if (Test-Flag -Row $Row -HeaderMap $HeaderMap -Key $definition.Key) {
                [void]$merged.Add($definition.Code)
            }
        }
        elseif ($existingSet.ContainsKey($definition.Code)) {
            [void]$merged.Add($definition.Code)
        }
    }

    return ConvertTo-UniqueStringArray $merged
}

function ConvertTo-ImportIdParseResult {
    param([AllowNull()][string]$Value)

    $text = Get-NormalizedString $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return [pscustomobject]@{ Success = $false; HasValue = $false; Value = 0 }
    }

    $parsed = 0
    if ([int]::TryParse($text, [ref]$parsed) -and $parsed -ge 0) {
        return [pscustomobject]@{ Success = $true; HasValue = $true; Value = $parsed }
    }

    return [pscustomobject]@{ Success = $false; HasValue = $true; Value = 0 }
}

function Read-LookupFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path $Path)) {
        throw "Lookup-Datei nicht gefunden: $Path"
    }

    $raw = Get-Content -Path $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) {
        throw "Lookup-Datei ist leer: $Path"
    }

    $parsed = $raw | ConvertFrom-Json
    $requiredProperties = @('responsibility_codes', 'assignment_tags', 'hazmat_flags', 'identifier_types', 'unit_codes')
    foreach ($propertyName in $requiredProperties) {
        if (-not $parsed.PSObject.Properties[$propertyName]) {
            throw "Lookup-Datei enthaelt kein Feld '$propertyName'"
        }
    }

    return $parsed
}

function ConvertTo-CodeSet {
    param([AllowNull()][object[]]$Entries)

    $codeSet = @{}
    foreach ($entry in @($Entries)) {
        if ($null -eq $entry) {
            continue
        }

        $code = Get-NormalizedString $entry.code
        if ([string]::IsNullOrWhiteSpace($code)) {
            continue
        }

        $codeSet[$code] = $true
    }

    return $codeSet
}

function Test-LookupCodesMatchDefinitions {
    param(
        [Parameter(Mandatory = $true)][string]$LookupName,
        [Parameter(Mandatory = $true)][hashtable]$LookupCodes,
        [Parameter(Mandatory = $true)][object[]]$Definitions
    )

    $definitionCodes = @{}
    foreach ($definition in $Definitions) {
        $definitionCodes[$definition.Code] = $true
    }

    $missingCodes = New-Object System.Collections.Generic.List[string]
    foreach ($definitionCode in $definitionCodes.Keys) {
        if (-not $LookupCodes.ContainsKey($definitionCode)) {
            [void]$missingCodes.Add($definitionCode)
        }
    }

    $unknownCodes = New-Object System.Collections.Generic.List[string]
    foreach ($lookupCode in $LookupCodes.Keys) {
        if (-not $definitionCodes.ContainsKey($lookupCode)) {
            [void]$unknownCodes.Add($lookupCode)
        }
    }

    return [pscustomobject]@{
        LookupName   = $LookupName
        MissingCodes = @($missingCodes)
        UnknownCodes = @($unknownCodes)
        IsValid      = ($missingCodes.Count -eq 0 -and $unknownCodes.Count -eq 0)
    }
}

function Get-NextMaterialId {
    param([AllowEmptyCollection()][object[]]$Materials)

    $maxId = 999
    if ($null -eq $Materials) {
        return $maxId
    }

    foreach ($material in $Materials) {
        $idText = Get-NormalizedString $material.id
        $idValue = 0
        if ([int]::TryParse($idText, [ref]$idValue) -and $idValue -gt $maxId) {
            $maxId = $idValue
        }
    }

    return $maxId
}

function Get-ExistingMaterialLookup {
    param([AllowEmptyCollection()][object[]]$Materials)

    $lookup = @{}
    if ($null -eq $Materials) {
        return $lookup
    }

    for ($i = 0; $i -lt $Materials.Count; $i++) {
        $material = $Materials[$i]
        $canonicalKey = Get-NormalizedString $material.canonical_key
        if ([string]::IsNullOrWhiteSpace($canonicalKey)) {
            $materialNumber = Get-NormalizedString $material.identifiers.matnr
            if ([string]::IsNullOrWhiteSpace($materialNumber)) {
                $materialNumber = Get-NormalizedString $material.primary_identifier.value
            }

            $canonicalKey = if ([string]::IsNullOrWhiteSpace($materialNumber)) { '' } else { Get-CanonicalKey -Type 'matnr' -Value $materialNumber }
        }

        if (-not [string]::IsNullOrWhiteSpace($canonicalKey) -and -not $lookup.ContainsKey($canonicalKey)) {
            $lookup[$canonicalKey] = $i
        }
    }

    return $lookup
}

function ConvertTo-NormalizedAlternateRecord {
    param($Alternate)

    $positionValue = 0
    [void][int]::TryParse((Get-NormalizedString $Alternate.position), [ref]$positionValue)

    return [pscustomobject][ordered]@{
        position             = $positionValue
        identifier           = [pscustomobject][ordered]@{
            type  = 'matnr'
            value = Get-NormalizedString $Alternate.identifier.value
        }
        material_status_code = Get-NormalizedString $Alternate.material_status_code
        preferred_unit_code  = $(if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $Alternate.preferred_unit_code))) { $null } else { Get-NormalizedString $Alternate.preferred_unit_code })
    }
}

function ConvertTo-NormalizedMaterialRecord {
    param([Parameter(Mandatory = $true)]$Material)

    $materialNumber = Get-NormalizedString $Material.identifiers.matnr
    if ([string]::IsNullOrWhiteSpace($materialNumber)) {
        $materialNumber = Get-NormalizedString $Material.primary_identifier.value
    }

    $normalizedAlternateUnits = New-Object System.Collections.Generic.List[object]
    foreach ($alternateUnit in @($Material.quantity.alternate_units)) {
        [void]$normalizedAlternateUnits.Add([pscustomobject][ordered]@{
                unit_code          = $(if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $alternateUnit.unit_code))) { $null } else { Get-NormalizedString $alternateUnit.unit_code })
                conversion_to_base = [double]$alternateUnit.conversion_to_base
            })
    }

    $normalizedAlternates = New-Object System.Collections.Generic.List[object]
    foreach ($alternate in @($Material.alternates)) {
        [void]$normalizedAlternates.Add((ConvertTo-NormalizedAlternateRecord -Alternate $alternate))
    }

    return [pscustomobject][ordered]@{
        id                 = [int]$Material.id
        canonical_key      = $(if ([string]::IsNullOrWhiteSpace($materialNumber)) { '' } else { Get-CanonicalKey -Type 'matnr' -Value $materialNumber })
        primary_identifier = [pscustomobject][ordered]@{
            type  = 'matnr'
            value = $materialNumber
        }
        identifiers        = [pscustomobject][ordered]@{
            matnr             = $(if ([string]::IsNullOrWhiteSpace($materialNumber)) { $null } else { $materialNumber })
            supply_number     = $(if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $Material.identifiers.supply_number))) { $null } else { Get-NormalizedString $Material.identifiers.supply_number })
            article_number    = $(if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $Material.identifiers.article_number))) { $null } else { Get-NormalizedString $Material.identifiers.article_number })
            nato_stock_number = $(if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $Material.identifiers.nato_stock_number))) { $null } else { Get-NormalizedString $Material.identifiers.nato_stock_number })
        }
        status             = [pscustomobject][ordered]@{
            material_status_code = $(if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $Material.status.material_status_code))) { 'XX' } else { Get-NormalizedString $Material.status.material_status_code })
        }
        texts              = [pscustomobject][ordered]@{
            short_description = Get-NormalizedString $Material.texts.short_description
            technical_note    = Get-NormalizedString $Material.texts.technical_note
            logistics_note    = Get-NormalizedString $Material.texts.logistics_note
        }
        classification     = [pscustomobject][ordered]@{
            ext_wg       = Get-NormalizedString $Material.classification.ext_wg
            is_decentral = [bool]$Material.classification.is_decentral
            creditor     = $(if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $Material.classification.creditor))) { $null } else { Get-NormalizedString $Material.classification.creditor })
        }
        hazmat             = [pscustomobject][ordered]@{
            is_hazardous = [bool]$Material.hazmat.is_hazardous
            un_number    = $(if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $Material.hazmat.un_number))) { $null } else { Get-NormalizedString $Material.hazmat.un_number })
            flags        = @(ConvertTo-UniqueStringArray $Material.hazmat.flags)
        }
        quantity           = [pscustomobject][ordered]@{
            base_unit       = $(if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $Material.quantity.base_unit))) { 'EA' } else { Get-NormalizedString $Material.quantity.base_unit })
            target          = [double]$Material.quantity.target
            alternate_units = @($normalizedAlternateUnits.ToArray())
        }
        alternates         = @($normalizedAlternates.ToArray())
        assignments        = [pscustomobject][ordered]@{
            responsibility_codes = @(ConvertTo-UniqueStringArray $Material.assignments.responsibility_codes)
            assignment_tags      = @(ConvertTo-UniqueStringArray $Material.assignments.assignment_tags)
        }
    }
}

function Get-ExistingMaterialIdLookup {
    param([AllowEmptyCollection()][object[]]$Materials)

    $lookup = @{}
    if ($null -eq $Materials) {
        return $lookup
    }

    for ($i = 0; $i -lt $Materials.Count; $i++) {
        $material = $Materials[$i]
        $idText = Get-NormalizedString $material.id
        $idValue = 0
        if ([int]::TryParse($idText, [ref]$idValue) -and -not $lookup.ContainsKey($idValue)) {
            $lookup[$idValue] = $i
        }
    }

    return $lookup
}

function New-MaterialRecord {
    param(
        [int]$Id,
        [string]$PrimaryIdentifierType,
        [string]$PrimaryIdentifierValue,
        [AllowNull()][string]$Matnr,
        [AllowNull()][string]$SupplyNumber,
        [AllowNull()][string]$ArticleNumber,
        [AllowNull()][string]$NatoStockNumber,
        [string]$MaterialStatusCode,
        [string]$ShortDescription,
        [string]$TechnicalNote,
        [string]$LogisticsNote,
        [string]$ExtWg,
        [bool]$IsDecentral,
        [AllowNull()][string]$Creditor,
        [bool]$IsHazardous,
        [AllowNull()][string]$UnNumber,
        [object[]]$HazmatFlags,
        [string]$BaseUnit,
        [double]$TargetQuantity,
        [object[]]$AlternateUnits,
        [object[]]$Alternates,
        [object[]]$ResponsibilityCodes,
        [object[]]$AssignmentTags
    )

    $resolvedMatnr = Get-NormalizedString $Matnr
    if ([string]::IsNullOrWhiteSpace($resolvedMatnr)) {
        $resolvedMatnr = Get-NormalizedString $PrimaryIdentifierValue
    }

    return [pscustomobject][ordered]@{
        id                 = $Id
        canonical_key      = Get-CanonicalKey -Type 'matnr' -Value $resolvedMatnr
        primary_identifier = [pscustomobject][ordered]@{
            type  = 'matnr'
            value = $resolvedMatnr
        }
        identifiers        = [pscustomobject][ordered]@{
            matnr             = $(if ([string]::IsNullOrWhiteSpace($resolvedMatnr)) { $null } else { $resolvedMatnr })
            supply_number     = $(if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $SupplyNumber))) { $null } else { $SupplyNumber })
            article_number    = $(if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $ArticleNumber))) { $null } else { $ArticleNumber })
            nato_stock_number = $(if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $NatoStockNumber))) { $null } else { $NatoStockNumber })
        }
        status             = [pscustomobject][ordered]@{
            material_status_code = $MaterialStatusCode
        }
        texts              = [pscustomobject][ordered]@{
            short_description = $ShortDescription
            technical_note    = $TechnicalNote
            logistics_note    = $LogisticsNote
        }
        classification     = [pscustomobject][ordered]@{
            ext_wg       = $ExtWg
            is_decentral = $IsDecentral
            creditor     = $(if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $Creditor))) { $null } else { $Creditor })
        }
        hazmat             = [pscustomobject][ordered]@{
            is_hazardous = $IsHazardous
            un_number    = $(if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $UnNumber))) { $null } else { $UnNumber })
            flags        = @(ConvertTo-UniqueStringArray $HazmatFlags)
        }
        quantity           = [pscustomobject][ordered]@{
            base_unit       = $BaseUnit
            target          = $TargetQuantity
            alternate_units = @($AlternateUnits)
        }
        alternates         = @($Alternates | ForEach-Object { ConvertTo-NormalizedAlternateRecord -Alternate $_ })
        assignments        = [pscustomobject][ordered]@{
            responsibility_codes = @(ConvertTo-UniqueStringArray $ResponsibilityCodes)
            assignment_tags      = @(ConvertTo-UniqueStringArray $AssignmentTags)
        }
    }
}

function Read-ExistingDatabase {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path $Path)) {
        return [pscustomobject]@{
            Materials = @()
            Lookup    = @{}
            MaxId     = 999
        }
    }

    $raw = Get-Content -Path $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return [pscustomobject]@{
            Materials = @()
            Lookup    = @{}
            MaxId     = 999
        }
    }

    $parsed = $raw | ConvertFrom-Json
    if (-not $parsed.PSObject.Properties['schema_version'] -or -not $parsed.PSObject.Properties['materials']) {
        throw "Nicht unterstuetztes Datenbankschema in '$Path'. Erwartet wird schema_version $DatabaseSchemaVersion."
    }

    $materials = @($parsed.materials | ForEach-Object { ConvertTo-NormalizedMaterialRecord -Material $_ })
    return [pscustomobject]@{
        Materials = $materials
        Lookup    = Get-ExistingMaterialLookup -Materials $materials
        MaxId     = Get-NextMaterialId -Materials $materials
    }
}

function Get-ImportedScalarValue {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][hashtable]$HeaderMap,
        [Parameter(Mandatory = $true)][string]$Key,
        [string]$ExistingValue = ''
    )

    if (Test-HeaderAvailable -HeaderMap $HeaderMap -Key $Key) {
        return Get-CellValue -Row $Row -HeaderMap $HeaderMap -Key $Key
    }

    return Get-NormalizedString $ExistingValue
}
function Convert-RowToImportData {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][hashtable]$HeaderMap,
        [Parameter(Mandatory = $true)][int]$RowNumber,
        $ExistingMaterial = $null
    )

    $warnings = New-Object System.Collections.Generic.List[string]

    $primaryValue = Get-CellValue -Row $Row -HeaderMap $HeaderMap -Key 'matnr_main'
    if ([string]::IsNullOrWhiteSpace($primaryValue)) {
        return [pscustomobject]@{
            ShouldSkip = $true
            Warnings   = @("Zeile $RowNumber - matnr_main leer, uebersprungen")
        }
    }

    $primaryType = 'matnr'
    $canonicalKey = Get-CanonicalKey -Type 'matnr' -Value $primaryValue
    $importIdResult = ConvertTo-ImportIdParseResult (Get-CellValue -Row $Row -HeaderMap $HeaderMap -Key 'import_id')

    $existingQuantityTarget = 0.0
    if ($ExistingMaterial) {
        $existingQuantityTarget = [double]$ExistingMaterial.quantity.target
    }

    if (Test-HeaderAvailable -HeaderMap $HeaderMap -Key 'quantity_target') {
        $qtyRaw = Get-CellValue -Row $Row -HeaderMap $HeaderMap -Key 'quantity_target'
        $qtyParse = ConvertTo-ImportNumberParseResult $qtyRaw
        $quantityTarget = $qtyParse.Value
        if (-not $qtyParse.Success) {
            if ($ExistingMaterial) {
                $quantityTarget = $existingQuantityTarget
                [void]$warnings.Add("Zeile $RowNumber - Menge '$qtyRaw' ungueltig, vorhandenen Wert beibehalten")
            }
            else {
                $quantityTarget = 0.0
                [void]$warnings.Add("Zeile $RowNumber - Menge '$qtyRaw' ungueltig, Standardwert 0 verwendet")
            }
        }
    }
    else {
        $quantityTarget = $existingQuantityTarget
    }

    $existingIsDecentral = $false
    if ($ExistingMaterial) {
        $existingIsDecentral = [bool]$ExistingMaterial.classification.is_decentral
    }

    if (Test-HeaderAvailable -HeaderMap $HeaderMap -Key 'dezentral') {
        $dezentRaw = Get-CellValue -Row $Row -HeaderMap $HeaderMap -Key 'dezentral'
        $dezentParse = ConvertTo-ImportBooleanParseResult $dezentRaw
        if ($dezentParse.Success) {
            $isDecentral = [bool]$dezentParse.Value
        }
        elseif ($ExistingMaterial) {
            $isDecentral = $existingIsDecentral
            [void]$warnings.Add("Zeile $RowNumber - Dezentralwert '$dezentRaw' ungueltig, vorhandenen Wert beibehalten")
        }
        else {
            $isDecentral = $false
            [void]$warnings.Add("Zeile $RowNumber - Dezentralwert '$dezentRaw' ungueltig, Standardwert false verwendet")
        }
    }
    else {
        $isDecentral = $existingIsDecentral
    }

    $existingIsHazardous = $false
    if ($ExistingMaterial) {
        $existingIsHazardous = [bool]$ExistingMaterial.hazmat.is_hazardous
    }

    if (Test-HeaderAvailable -HeaderMap $HeaderMap -Key 'is_dg') {
        $isDgRaw = Get-CellValue -Row $Row -HeaderMap $HeaderMap -Key 'is_dg'
        $isDgParse = ConvertTo-ImportBooleanParseResult $isDgRaw
        if ($isDgParse.Success) {
            $explicitIsHazardous = [bool]$isDgParse.Value
        }
        elseif ($ExistingMaterial) {
            $explicitIsHazardous = $existingIsHazardous
            [void]$warnings.Add("Zeile $RowNumber - Gefahrstoffwert '$isDgRaw' ungueltig, vorhandenen Wert beibehalten")
        }
        else {
            $explicitIsHazardous = $false
            [void]$warnings.Add("Zeile $RowNumber - Gefahrstoffwert '$isDgRaw' ungueltig, Standardwert false verwendet")
        }
    }
    else {
        $explicitIsHazardous = $existingIsHazardous
    }

    $hazmatFlags = Merge-DefinedTagValues -Row $Row -HeaderMap $HeaderMap -Definitions (Get-HazmatFlagDefinitions) -ExistingValues $ExistingMaterial.hazmat.flags
    $responsibilityCodes = Merge-DefinedTagValues -Row $Row -HeaderMap $HeaderMap -Definitions (Get-ResponsibilityDefinitions) -ExistingValues $ExistingMaterial.assignments.responsibility_codes
    $assignmentTags = Merge-DefinedTagValues -Row $Row -HeaderMap $HeaderMap -Definitions (Get-AssignmentTagDefinitions) -ExistingValues $ExistingMaterial.assignments.assignment_tags

    $materialStatusCode = 'XX'
    if ($ExistingMaterial) {
        $materialStatusCode = Get-NormalizedString $ExistingMaterial.status.material_status_code
        if ([string]::IsNullOrWhiteSpace($materialStatusCode)) {
            $materialStatusCode = 'XX'
        }
    }

    if (Test-HeaderAvailable -HeaderMap $HeaderMap -Key 'mat_stat_main') {
        $matStatRaw = Get-CellValue -Row $Row -HeaderMap $HeaderMap -Key 'mat_stat_main'
        if (-not [string]::IsNullOrWhiteSpace($matStatRaw)) {
            $trimmedMatStat = $matStatRaw.Trim()
            if ($trimmedMatStat.Length -eq 2) {
                $materialStatusCode = $trimmedMatStat
            }
            else {
                [void]$warnings.Add("Zeile $RowNumber - MatStatus '$matStatRaw' ungueltig, Wert beibehalten")
            }
        }
    }

    $existingSupplyNumber = if ($ExistingMaterial) { Get-NormalizedString $ExistingMaterial.identifiers.supply_number } else { '' }
    $existingArticleNumber = if ($ExistingMaterial) { Get-NormalizedString $ExistingMaterial.identifiers.article_number } else { '' }
    $existingNatoStockNumber = if ($ExistingMaterial) { Get-NormalizedString $ExistingMaterial.identifiers.nato_stock_number } else { '' }

    $artNrValue = Get-ImportedScalarValue -Row $Row -HeaderMap $HeaderMap -Key 'artnr' -ExistingValue $existingArticleNumber
    $matnrValue = $primaryValue
    $articleNumber = $artNrValue

    $existingAlternateUnits = @()
    $existingAlternates = @()
    $existingCreditor = ''
    $existingUnNumber = ''
    if ($ExistingMaterial) {
        $existingAlternateUnits = @($ExistingMaterial.quantity.alternate_units)
        $existingAlternates = @($ExistingMaterial.alternates)
        $existingCreditor = Get-NormalizedString $ExistingMaterial.classification.creditor
        $existingUnNumber = Get-NormalizedString $ExistingMaterial.hazmat.un_number
    }

    $resolvedId = 0
    if ($importIdResult.Success) {
        $resolvedId = [int]$importIdResult.Value
    }
    elseif ($importIdResult.HasValue) {
        if ($ExistingMaterial) {
            $resolvedId = [int]$ExistingMaterial.id
            [void]$warnings.Add("Zeile $RowNumber - ID '$($Row.PSObject.Properties[$HeaderMap['import_id']].Value)' ungueltig, vorhandene ID beibehalten")
        }
        else {
            [void]$warnings.Add("Zeile $RowNumber - ID '$($Row.PSObject.Properties[$HeaderMap['import_id']].Value)' ungueltig, neue ID wird vergeben")
        }
    }
    elseif ($ExistingMaterial) {
        $resolvedId = [int]$ExistingMaterial.id
    }

    return [pscustomobject]@{
        ShouldSkip            = $false
        Warnings              = @($warnings)
        ResolvedId            = $resolvedId
        CanonicalKey          = $canonicalKey
        PrimaryIdentifierType = $primaryType
        PrimaryIdentifierValue = $primaryValue
        Matnr                 = $matnrValue
        SupplyNumber          = Get-ImportedScalarValue -Row $Row -HeaderMap $HeaderMap -Key 'supplynumber' -ExistingValue $existingSupplyNumber
        ArticleNumber         = $articleNumber
        NatoStockNumber       = $(if ([string]::IsNullOrWhiteSpace($existingNatoStockNumber)) { $null } else { $existingNatoStockNumber })
        MaterialStatusCode    = $materialStatusCode
        ShortDescription      = Get-ImportedScalarValue -Row $Row -HeaderMap $HeaderMap -Key 'description' -ExistingValue $ExistingMaterial.texts.short_description
        TechnicalNote         = Get-ImportedScalarValue -Row $Row -HeaderMap $HeaderMap -Key 'technical' -ExistingValue $ExistingMaterial.texts.technical_note
        LogisticsNote         = Get-ImportedScalarValue -Row $Row -HeaderMap $HeaderMap -Key 'logistics' -ExistingValue $ExistingMaterial.texts.logistics_note
        ExtWg                 = Get-ImportedScalarValue -Row $Row -HeaderMap $HeaderMap -Key 'ext_wg' -ExistingValue $ExistingMaterial.classification.ext_wg
        IsDecentral           = $isDecentral
        Creditor              = $existingCreditor
        IsHazardous           = ($explicitIsHazardous -or $hazmatFlags.Count -gt 0)
        UnNumber              = $existingUnNumber
        HazmatFlags           = $hazmatFlags
        BaseUnit              = Get-ImportedScalarValue -Row $Row -HeaderMap $HeaderMap -Key 'unit_main' -ExistingValue $ExistingMaterial.quantity.base_unit
        TargetQuantity        = [double]$quantityTarget
        AlternateUnits        = @($existingAlternateUnits)
        Alternates            = @($existingAlternates)
        ResponsibilityCodes   = $responsibilityCodes
        AssignmentTags        = $assignmentTags
    }
}

function Merge-MaterialRecord {
    param(
        [Parameter(Mandatory = $true)]$ImportData,
        [Parameter(Mandatory = $true)][int]$Id
    )

    return New-MaterialRecord `
        -Id $Id `
        -PrimaryIdentifierType $ImportData.PrimaryIdentifierType `
        -PrimaryIdentifierValue $ImportData.PrimaryIdentifierValue `
        -Matnr $ImportData.Matnr `
        -SupplyNumber $ImportData.SupplyNumber `
        -ArticleNumber $ImportData.ArticleNumber `
        -NatoStockNumber $ImportData.NatoStockNumber `
        -MaterialStatusCode $ImportData.MaterialStatusCode `
        -ShortDescription $ImportData.ShortDescription `
        -TechnicalNote $ImportData.TechnicalNote `
        -LogisticsNote $ImportData.LogisticsNote `
        -ExtWg $ImportData.ExtWg `
        -IsDecentral ([bool]$ImportData.IsDecentral) `
        -Creditor $ImportData.Creditor `
        -IsHazardous ([bool]$ImportData.IsHazardous) `
        -UnNumber $ImportData.UnNumber `
        -HazmatFlags $ImportData.HazmatFlags `
        -BaseUnit $ImportData.BaseUnit `
        -TargetQuantity ([double]$ImportData.TargetQuantity) `
        -AlternateUnits $ImportData.AlternateUnits `
        -Alternates $ImportData.Alternates `
        -ResponsibilityCodes $ImportData.ResponsibilityCodes `
        -AssignmentTags $ImportData.AssignmentTags
}

function Test-DuplicateImportKeys {
    param(
        [Parameter(Mandatory = $true)][object[]]$Rows,
        [Parameter(Mandatory = $true)][hashtable]$HeaderMap
    )

    $seen = @{}
    $duplicates = New-Object System.Collections.Generic.List[string]

    for ($index = 0; $index -lt $Rows.Count; $index++) {
        $rowNumber = $index + 2
        $primaryValue = Get-CellValue -Row $Rows[$index] -HeaderMap $HeaderMap -Key 'matnr_main'
        if ([string]::IsNullOrWhiteSpace($primaryValue)) {
            continue
        }

        $canonicalKey = Get-CanonicalKey -Type 'matnr' -Value $primaryValue
        if ($seen.ContainsKey($canonicalKey)) {
            [void]$duplicates.Add("Schluessel '$canonicalKey' in Zeile $($seen[$canonicalKey]) und Zeile $rowNumber")
        }
        else {
            $seen[$canonicalKey] = $rowNumber
        }
    }

    return @($duplicates)
}

function Test-DuplicateImportIds {
    param(
        [Parameter(Mandatory = $true)][object[]]$Rows,
        [Parameter(Mandatory = $true)][hashtable]$HeaderMap
    )

    if (-not (Test-HeaderAvailable -HeaderMap $HeaderMap -Key 'import_id')) {
        return @()
    }

    $seen = @{}
    $duplicates = New-Object System.Collections.Generic.List[string]

    for ($index = 0; $index -lt $Rows.Count; $index++) {
        $rowNumber = $index + 2
        $idResult = ConvertTo-ImportIdParseResult (Get-CellValue -Row $Rows[$index] -HeaderMap $HeaderMap -Key 'import_id')
        if (-not $idResult.Success) {
            continue
        }

        $importId = [int]$idResult.Value
        if ($seen.ContainsKey($importId)) {
            [void]$duplicates.Add("ID '$importId' in Zeile $($seen[$importId]) und Zeile $rowNumber")
        }
        else {
            $seen[$importId] = $rowNumber
        }
    }

    return @($duplicates)
}

function Backup-DatabaseFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path $Path)) {
        return $null
    }

    $backupName = "db_verlegepaket_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $backupPath = Join-Path $BackupDir $backupName
    Copy-Item -Path $Path -Destination $backupPath -Force
    return $backupPath
}

function Save-DatabaseFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object[]]$Materials
    )

    $database = [pscustomobject][ordered]@{
        schema_version = $DatabaseSchemaVersion
        materials      = @($Materials)
    }

    $database | ConvertTo-Json -Depth 12 | Out-File -FilePath $Path -Encoding UTF8
}

function Get-SapImportFieldDefinitions {
    return @(
        [pscustomobject]@{ Key = 'matnr'; Label = 'Materialnummer (MATNR)'; Required = $true; ValueType = 'text' }
        [pscustomobject]@{ Key = 'supply_number'; Label = 'Versorgungsnummer'; Required = $false; ValueType = 'text' }
        [pscustomobject]@{ Key = 'nato_stock_number'; Label = 'NATO Stock Number'; Required = $false; ValueType = 'text' }
        [pscustomobject]@{ Key = 'material_status_code'; Label = 'Materialstatus'; Required = $false; ValueType = 'status' }
        [pscustomobject]@{ Key = 'short_description'; Label = 'Kurzbezeichnung'; Required = $false; ValueType = 'text' }
        [pscustomobject]@{ Key = 'ext_wg'; Label = 'Ext WG'; Required = $false; ValueType = 'text' }
        [pscustomobject]@{ Key = 'is_decentral'; Label = 'Dezentral'; Required = $false; ValueType = 'bool' }
        [pscustomobject]@{ Key = 'base_unit'; Label = 'Basiseinheit'; Required = $false; ValueType = 'unit' }
        [pscustomobject]@{ Key = 'target_quantity'; Label = 'Zielmenge'; Required = $false; ValueType = 'number' }
    )
}

function New-DefaultDataImportPresetStore {
    return [pscustomobject][ordered]@{
        schema_version = $DataImportPresetSchemaVersion
        presets        = @()
    }
}

function ConvertTo-NormalizedDataImportPreset {
    param(
        [Parameter(Mandatory = $true)]$Preset,
        [string]$DefaultFileType = ''
    )

    $definitions = @(Get-SapImportFieldDefinitions)
    $name = ConvertTo-NullableString $Preset.Name
    $fileType = Get-NormalizedString $Preset.FileType
    if ([string]::IsNullOrWhiteSpace($fileType)) {
        $fileType = Get-NormalizedString $DefaultFileType
    }

    if (-not [string]::IsNullOrWhiteSpace($fileType)) {
        $fileType = $fileType.ToLowerInvariant()
    }

    if ([string]::IsNullOrWhiteSpace($fileType) -or @('xlsx', 'csv', 'txt') -notcontains $fileType) {
        throw "Preset '$name' hat einen ungueltigen Dateityp. Erwartet: xlsx, csv oder txt."
    }

    $headerRowIndex = 0
    [void][int]::TryParse((Get-NormalizedString $Preset.HeaderRowIndex), [ref]$headerRowIndex)
    if ($headerRowIndex -le 0) {
        $headerRowIndex = 1
    }

    $worksheetName = ConvertTo-NullableString $Preset.WorksheetName
    $delimiter = ConvertTo-NullableString $Preset.Delimiter
    if ($fileType -eq 'xlsx') {
        $delimiter = $null
    }
    elseif ([string]::IsNullOrWhiteSpace((Get-NormalizedString $delimiter))) {
        $delimiter = ';'
    }

    $columnSource = $Preset.Columns
    $columns = [ordered]@{}
    $usedHeaders = @{}
    foreach ($definition in $definitions) {
        $headerName = $null
        if ($null -ne $columnSource -and $null -ne $columnSource.PSObject.Properties[$definition.Key]) {
            $headerName = ConvertTo-NullableString $columnSource.PSObject.Properties[$definition.Key].Value
        }

        if ($definition.Required -and [string]::IsNullOrWhiteSpace((Get-NormalizedString $headerName))) {
            throw "Preset '$name' fehlt die Pflichtzuordnung '$($definition.Key)'."
        }

        if (-not [string]::IsNullOrWhiteSpace((Get-NormalizedString $headerName))) {
            if ($usedHeaders.ContainsKey($headerName)) {
                throw "Preset '$name' weist die Importspalte '$headerName' mehrfach zu."
            }

            $usedHeaders[$headerName] = $definition.Key
        }

        $columns[$definition.Key] = $headerName
    }

    return [pscustomobject][ordered]@{
        Name           = $name
        FileType       = $fileType
        WorksheetName  = $worksheetName
        HeaderRowIndex = $headerRowIndex
        Delimiter      = $delimiter
        Columns        = [pscustomobject]$columns
    }
}

function Read-DataImportPresetStore {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path $Path)) {
        return (New-DefaultDataImportPresetStore)
    }

    $raw = Get-Content -Path $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return (New-DefaultDataImportPresetStore)
    }

    $parsed = $raw | ConvertFrom-Json
    $schemaVersion = 0
    if (-not [int]::TryParse((Get-NormalizedString $parsed.schema_version), [ref]$schemaVersion) -or $schemaVersion -ne $DataImportPresetSchemaVersion) {
        throw "Preset-Speicher '$Path' hat schema_version '$($parsed.schema_version)'. Erwartet: $DataImportPresetSchemaVersion."
    }

    $presets = New-Object System.Collections.Generic.List[object]
    $nameSet = @{}
    foreach ($preset in @(ConvertTo-ObjectArray $parsed.presets)) {
        $normalized = ConvertTo-NormalizedDataImportPreset -Preset $preset
        $presetName = Get-NormalizedString $normalized.Name
        if ([string]::IsNullOrWhiteSpace($presetName)) {
            throw "Preset-Speicher '$Path' enthaelt ein Preset ohne Namen."
        }

        if ($nameSet.ContainsKey($presetName)) {
            throw "Preset-Speicher '$Path' enthaelt das Preset '$presetName' mehrfach."
        }

        $nameSet[$presetName] = $true
        [void]$presets.Add($normalized)
    }

    return [pscustomobject][ordered]@{
        schema_version = $DataImportPresetSchemaVersion
        presets        = @($presets.ToArray())
    }
}

function Write-DataImportPresetStore {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$Store
    )

    $normalizedPresets = @(
        @(ConvertTo-ObjectArray $Store.presets) | ForEach-Object {
            ConvertTo-NormalizedDataImportPreset -Preset $_
        }
    )

    $resolvedStore = [pscustomobject][ordered]@{
        schema_version = $DataImportPresetSchemaVersion
        presets        = @($normalizedPresets)
    }

    $resolvedStore | ConvertTo-Json -Depth 10 | Out-File -FilePath $Path -Encoding UTF8
}

function Set-DataImportPreset {
    param(
        [Parameter(Mandatory = $true)]$Preset,
        [string]$Path = $DataImportPresetPath
    )

    $normalizedPreset = ConvertTo-NormalizedDataImportPreset -Preset $Preset
    $presetName = Get-NormalizedString $normalizedPreset.Name
    if ([string]::IsNullOrWhiteSpace($presetName)) {
        throw 'Preset-Name ist erforderlich.'
    }

    $store = Read-DataImportPresetStore -Path $Path
    $presets = New-Object System.Collections.Generic.List[object]
    $replaced = $false
    foreach ($existingPreset in @(ConvertTo-ObjectArray $store.presets)) {
        if ((Get-NormalizedString $existingPreset.Name) -eq $presetName) {
            [void]$presets.Add($normalizedPreset)
            $replaced = $true
        }
        else {
            [void]$presets.Add($existingPreset)
        }
    }

    if (-not $replaced) {
        [void]$presets.Add($normalizedPreset)
    }

    Write-DataImportPresetStore -Path $Path -Store ([pscustomobject][ordered]@{
            schema_version = $DataImportPresetSchemaVersion
            presets        = @($presets.ToArray())
        })

    return $normalizedPreset
}

function Remove-DataImportPreset {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [string]$Path = $DataImportPresetPath
    )

    $presetName = Get-NormalizedString $Name
    if ([string]::IsNullOrWhiteSpace($presetName)) {
        throw 'Preset-Name ist erforderlich.'
    }

    $store = Read-DataImportPresetStore -Path $Path
    $presets = New-Object System.Collections.Generic.List[object]
    $removed = $false
    foreach ($existingPreset in @(ConvertTo-ObjectArray $store.presets)) {
        if ((Get-NormalizedString $existingPreset.Name) -eq $presetName) {
            $removed = $true
            continue
        }

        [void]$presets.Add($existingPreset)
    }

    if (-not $removed) {
        throw "Preset '$presetName' wurde nicht gefunden."
    }

    Write-DataImportPresetStore -Path $Path -Store ([pscustomobject][ordered]@{
            schema_version = $DataImportPresetSchemaVersion
            presets        = @($presets.ToArray())
        })
}

function Get-SourceColumnValue {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)]$ColumnMap,
        [Parameter(Mandatory = $true)][string]$Key
    )

    if ($null -eq $ColumnMap -or $null -eq $ColumnMap.PSObject.Properties[$Key]) {
        return ''
    }

    $headerName = Get-NormalizedString $ColumnMap.PSObject.Properties[$Key].Value
    if ([string]::IsNullOrWhiteSpace($headerName)) {
        return ''
    }

    $property = $Row.PSObject.Properties[$headerName]
    if ($null -eq $property) {
        return ''
    }

    return Get-NormalizedString $property.Value
}

function Test-SapImportColumnMappings {
    param(
        [Parameter(Mandatory = $true)]$Preset,
        [Parameter(Mandatory = $true)][string[]]$Headers
    )

    $missingMappings = New-Object System.Collections.Generic.List[string]
    $missingHeaders = New-Object System.Collections.Generic.List[string]
    foreach ($definition in @(Get-SapImportFieldDefinitions)) {
        $mappedHeader = Get-NormalizedString $Preset.Columns.PSObject.Properties[$definition.Key].Value
        if ($definition.Required -and [string]::IsNullOrWhiteSpace($mappedHeader)) {
            [void]$missingMappings.Add($definition.Label)
            continue
        }

        if (-not [string]::IsNullOrWhiteSpace($mappedHeader) -and @($Headers) -notcontains $mappedHeader) {
            [void]$missingHeaders.Add("$($definition.Label) -> $mappedHeader")
        }
    }

    return [pscustomobject]@{
        MissingMappings = @($missingMappings)
        MissingHeaders  = @($missingHeaders)
        IsValid         = ($missingMappings.Count -eq 0 -and $missingHeaders.Count -eq 0)
    }
}

function Get-DefaultUnitCodeFromLookup {
    param($LookupFile)

    foreach ($entry in @(ConvertTo-ObjectArray $LookupFile.unit_codes)) {
        $code = Get-NormalizedString $entry.code
        if (-not [string]::IsNullOrWhiteSpace($code)) {
            return $code
        }
    }

    return 'EA'
}

function Convert-SapRowToImportData {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)]$Preset,
        [Parameter(Mandatory = $true)][int]$RowNumber,
        $ExistingMaterial = $null,
        [Parameter(Mandatory = $true)][string]$DefaultUnitCode,
        [Parameter(Mandatory = $true)][hashtable]$ValidUnitCodes
    )

    $warnings = New-Object System.Collections.Generic.List[string]
    $columnMap = $Preset.Columns
    $matnr = Get-SourceColumnValue -Row $Row -ColumnMap $columnMap -Key 'matnr'
    if ([string]::IsNullOrWhiteSpace($matnr)) {
        return [pscustomobject]@{
            ShouldSkip = $true
            Warnings   = @("Zeile $RowNumber - MATNR leer, uebersprungen")
        }
    }

    $existingSupplyNumber = if ($ExistingMaterial) { Get-NormalizedString $ExistingMaterial.identifiers.supply_number } else { '' }
    $existingArticleNumber = if ($ExistingMaterial) { Get-NormalizedString $ExistingMaterial.identifiers.article_number } else { '' }
    $existingNatoStockNumber = if ($ExistingMaterial) { Get-NormalizedString $ExistingMaterial.identifiers.nato_stock_number } else { '' }
    $existingMaterialStatusCode = if ($ExistingMaterial) { Get-NormalizedString $ExistingMaterial.status.material_status_code } else { 'XX' }
    if ([string]::IsNullOrWhiteSpace($existingMaterialStatusCode)) {
        $existingMaterialStatusCode = 'XX'
    }

    $existingShortDescription = if ($ExistingMaterial) { Get-NormalizedString $ExistingMaterial.texts.short_description } else { '' }
    $existingTechnicalNote = if ($ExistingMaterial) { Get-NormalizedString $ExistingMaterial.texts.technical_note } else { '' }
    $existingLogisticsNote = if ($ExistingMaterial) { Get-NormalizedString $ExistingMaterial.texts.logistics_note } else { '' }
    $existingExtWg = if ($ExistingMaterial) { Get-NormalizedString $ExistingMaterial.classification.ext_wg } else { '' }
    $existingIsDecentral = if ($ExistingMaterial) { [bool]$ExistingMaterial.classification.is_decentral } else { $false }
    $existingCreditor = if ($ExistingMaterial) { Get-NormalizedString $ExistingMaterial.classification.creditor } else { '' }
    $existingIsHazardous = if ($ExistingMaterial) { [bool]$ExistingMaterial.hazmat.is_hazardous } else { $false }
    $existingUnNumber = if ($ExistingMaterial) { Get-NormalizedString $ExistingMaterial.hazmat.un_number } else { '' }
    $existingHazmatFlags = if ($ExistingMaterial) { @(ConvertTo-ObjectArray $ExistingMaterial.hazmat.flags) } else { @() }
    $existingBaseUnit = if ($ExistingMaterial) { Get-NormalizedString $ExistingMaterial.quantity.base_unit } else { $DefaultUnitCode }
    if ([string]::IsNullOrWhiteSpace($existingBaseUnit)) {
        $existingBaseUnit = $DefaultUnitCode
    }

    $existingTargetQuantity = if ($ExistingMaterial) { [double]$ExistingMaterial.quantity.target } else { 0.0 }
    $existingAlternateUnits = if ($ExistingMaterial) { @(ConvertTo-ObjectArray $ExistingMaterial.quantity.alternate_units) } else { @() }
    $existingAlternates = if ($ExistingMaterial) { @(ConvertTo-ObjectArray $ExistingMaterial.alternates) } else { @() }
    $existingResponsibilityCodes = if ($ExistingMaterial) { @(ConvertTo-ObjectArray $ExistingMaterial.assignments.responsibility_codes) } else { @() }
    $existingAssignmentTags = if ($ExistingMaterial) { @(ConvertTo-ObjectArray $ExistingMaterial.assignments.assignment_tags) } else { @() }

    $supplyNumber = $existingSupplyNumber
    $supplyNumberRaw = Get-SourceColumnValue -Row $Row -ColumnMap $columnMap -Key 'supply_number'
    if (-not [string]::IsNullOrWhiteSpace($supplyNumberRaw)) {
        $supplyNumber = $supplyNumberRaw
    }

    $natoStockNumber = $existingNatoStockNumber
    $natoRaw = Get-SourceColumnValue -Row $Row -ColumnMap $columnMap -Key 'nato_stock_number'
    if (-not [string]::IsNullOrWhiteSpace($natoRaw)) {
        $natoStockNumber = $natoRaw
    }

    $materialStatusCode = $existingMaterialStatusCode
    $statusRaw = Get-SourceColumnValue -Row $Row -ColumnMap $columnMap -Key 'material_status_code'
    if (-not [string]::IsNullOrWhiteSpace($statusRaw)) {
        $trimmedStatus = $statusRaw.Trim()
        if ($trimmedStatus.Length -eq 2) {
            $materialStatusCode = $trimmedStatus
        }
        else {
            [void]$warnings.Add("Zeile $RowNumber - Materialstatus '$statusRaw' ungueltig, vorhandenen Wert beibehalten")
        }
    }

    $shortDescription = $existingShortDescription
    $shortDescriptionRaw = Get-SourceColumnValue -Row $Row -ColumnMap $columnMap -Key 'short_description'
    if (-not [string]::IsNullOrWhiteSpace($shortDescriptionRaw)) {
        $shortDescription = $shortDescriptionRaw
    }

    $extWg = $existingExtWg
    $extWgRaw = Get-SourceColumnValue -Row $Row -ColumnMap $columnMap -Key 'ext_wg'
    if (-not [string]::IsNullOrWhiteSpace($extWgRaw)) {
        $extWg = $extWgRaw
    }

    $isDecentral = $existingIsDecentral
    $decentralRaw = Get-SourceColumnValue -Row $Row -ColumnMap $columnMap -Key 'is_decentral'
    if (-not [string]::IsNullOrWhiteSpace($decentralRaw)) {
        $decentralParse = ConvertTo-ImportBooleanParseResult $decentralRaw
        if ($decentralParse.Success) {
            $isDecentral = [bool]$decentralParse.Value
        }
        else {
            [void]$warnings.Add("Zeile $RowNumber - Dezentralwert '$decentralRaw' ungueltig, vorhandenen Wert beibehalten")
        }
    }

    $baseUnit = $existingBaseUnit
    $baseUnitRaw = Get-SourceColumnValue -Row $Row -ColumnMap $columnMap -Key 'base_unit'
    if (-not [string]::IsNullOrWhiteSpace($baseUnitRaw)) {
        if ($ValidUnitCodes.ContainsKey($baseUnitRaw)) {
            $baseUnit = $baseUnitRaw
        }
        else {
            [void]$warnings.Add("Zeile $RowNumber - Einheit '$baseUnitRaw' ungueltig, vorhandenen Wert beibehalten")
        }
    }

    $targetQuantity = $existingTargetQuantity
    $targetQuantityRaw = Get-SourceColumnValue -Row $Row -ColumnMap $columnMap -Key 'target_quantity'
    if (-not [string]::IsNullOrWhiteSpace($targetQuantityRaw)) {
        $numberParse = ConvertTo-ImportNumberParseResult $targetQuantityRaw
        if ($numberParse.Success) {
            $targetQuantity = [double]$numberParse.Value
        }
        else {
            [void]$warnings.Add("Zeile $RowNumber - Zielmenge '$targetQuantityRaw' ungueltig, vorhandenen Wert beibehalten")
        }
    }

    return [pscustomobject]@{
        ShouldSkip             = $false
        Warnings               = @($warnings)
        CanonicalKey           = Get-CanonicalKey -Type 'matnr' -Value $matnr
        PrimaryIdentifierType  = 'matnr'
        PrimaryIdentifierValue = $matnr
        Matnr                  = $matnr
        SupplyNumber           = $supplyNumber
        ArticleNumber          = $(if ([string]::IsNullOrWhiteSpace($existingArticleNumber)) { $null } else { $existingArticleNumber })
        NatoStockNumber        = $(if ([string]::IsNullOrWhiteSpace($natoStockNumber)) { $null } else { $natoStockNumber })
        MaterialStatusCode     = $materialStatusCode
        ShortDescription       = $shortDescription
        TechnicalNote          = $existingTechnicalNote
        LogisticsNote          = $existingLogisticsNote
        ExtWg                  = $extWg
        IsDecentral            = [bool]$isDecentral
        Creditor               = $(if ([string]::IsNullOrWhiteSpace($existingCreditor)) { $null } else { $existingCreditor })
        IsHazardous            = [bool]$existingIsHazardous
        UnNumber               = $(if ([string]::IsNullOrWhiteSpace($existingUnNumber)) { $null } else { $existingUnNumber })
        HazmatFlags            = @($existingHazmatFlags)
        BaseUnit               = $baseUnit
        TargetQuantity         = [double]$targetQuantity
        AlternateUnits         = @($existingAlternateUnits)
        Alternates             = @($existingAlternates)
        ResponsibilityCodes    = @($existingResponsibilityCodes)
        AssignmentTags         = @($existingAssignmentTags)
    }
}

function Test-DuplicateSapImportMatnrs {
    param(
        [Parameter(Mandatory = $true)][object[]]$Rows,
        [Parameter(Mandatory = $true)]$Preset,
        [AllowNull()][int[]]$RowNumbers
    )

    $seen = @{}
    $duplicates = New-Object System.Collections.Generic.List[string]
    for ($index = 0; $index -lt $Rows.Count; $index++) {
        $row = $Rows[$index]
        $rowNumber = if ($null -ne $RowNumbers -and $RowNumbers.Count -gt $index) { [int]$RowNumbers[$index] } else { $index + 2 }
        $matnr = Get-SourceColumnValue -Row $row -ColumnMap $Preset.Columns -Key 'matnr'
        if ([string]::IsNullOrWhiteSpace($matnr)) {
            continue
        }

        $canonicalKey = Get-CanonicalKey -Type 'matnr' -Value $matnr
        if ($seen.ContainsKey($canonicalKey)) {
            [void]$duplicates.Add("Schluessel '$canonicalKey' in Zeile $($seen[$canonicalKey]) und Zeile $rowNumber")
        }
        else {
            $seen[$canonicalKey] = $rowNumber
        }
    }

    return @($duplicates)
}

function Start-SapMaintenanceImport {
    param(
        [Parameter(Mandatory = $true)][string]$SourceFile,
        [Parameter(Mandatory = $true)]$Preset,
        [switch]$SuppressSuccessMessage
    )

    if (-not (Test-Path $SourceFile)) {
        Write-ImportLog "Datei nicht gefunden: $SourceFile" 'ERROR'
        return
    }

    $resolvedPreset = ConvertTo-NormalizedDataImportPreset -Preset $Preset -DefaultFileType (Get-ImportFileTypeFromPath -Path $SourceFile)
    $actualFileType = Get-ImportFileTypeFromPath -Path $SourceFile
    if ($resolvedPreset.FileType -ne $actualFileType) {
        Write-ImportLog "Dateityp-Konflikt: Preset erwartet '$($resolvedPreset.FileType)', Datei ist '$actualFileType'" 'ERROR'
        return
    }

    Write-ImportLog "Starte SAP-Wartungsimport aus $SourceFile ..." 'INFO'

    $sourceData = $null
    try {
        $sourceData = Read-GenericImportSource -Path $SourceFile -FileType $resolvedPreset.FileType -HeaderRowIndex ([int]$resolvedPreset.HeaderRowIndex) -WorksheetName $resolvedPreset.WorksheetName -Delimiter $resolvedPreset.Delimiter
        Write-ImportLog "Quelldatei geladen - $($sourceData.Rows.Count) Zeilen gefunden" 'INFO'
        if (-not [string]::IsNullOrWhiteSpace((Get-NormalizedString $sourceData.Worksheet))) {
            Write-ImportLog "Worksheet: $($sourceData.Worksheet)" 'INFO'
        }
        if (-not [string]::IsNullOrWhiteSpace((Get-NormalizedString $sourceData.EncodingName))) {
            Write-ImportLog "Encoding: $($sourceData.EncodingName)" 'INFO'
        }
    }
    catch {
        Write-ImportLog "Quelldatei konnte nicht gelesen werden: $($_.Exception.Message)" 'ERROR'
        return
    }

    if ($null -eq $sourceData.Rows -or $sourceData.Rows.Count -eq 0) {
        Write-ImportLog 'Keine Datenzeilen gefunden' 'ERROR'
        return
    }

    $headers = @($sourceData.Headers)
    Write-ImportLog "Header erkannt: $($headers -join ' | ')" 'INFO'

    $columnValidation = Test-SapImportColumnMappings -Preset $resolvedPreset -Headers $headers
    if (-not $columnValidation.IsValid) {
        if ($columnValidation.MissingMappings.Count -gt 0) {
            Write-ImportLog "Pflichtzuordnungen fehlen: $($columnValidation.MissingMappings -join ', ')" 'ERROR'
        }

        if ($columnValidation.MissingHeaders.Count -gt 0) {
            Write-ImportLog "Zuordnungen verweisen auf fehlende Importspalten: $($columnValidation.MissingHeaders -join ', ')" 'ERROR'
        }

        return
    }

    $duplicateKeys = Test-DuplicateSapImportMatnrs -Rows $sourceData.Rows -Preset $resolvedPreset -RowNumbers $sourceData.RowNumbers
    if ($duplicateKeys.Count -gt 0) {
        foreach ($duplicate in $duplicateKeys) {
            Write-ImportLog "Doppelte MATNR im Import gefunden: $duplicate" 'ERROR'
        }

        Write-ImportLog 'Import abgebrochen - doppelte MATNR im SAP-Import' 'ERROR'
        return
    }

    $lookupFile = $null
    try {
        $lookupFile = Read-LookupFile -Path $LookupPath
        Write-ImportLog "Lookup-Datei geladen: $LookupPath" 'INFO'
    }
    catch {
        Write-ImportLog "Lookup-Datei nicht lesbar: $($_.Exception.Message)" 'ERROR'
        return
    }

    $validUnitCodes = ConvertTo-CodeSet $lookupFile.unit_codes
    $defaultUnitCode = Get-DefaultUnitCodeFromLookup -LookupFile $lookupFile

    $existingDb = $null
    try {
        $existingDb = Read-ExistingDatabase -Path $DbPath
        Write-ImportLog "Vorhandene Datenbank geladen - $($existingDb.Materials.Count) Materialien" 'INFO'
    }
    catch {
        Write-ImportLog "Vorhandene Datenbank nicht lesbar: $($_.Exception.Message)" 'ERROR'
        return
    }

    $materials = New-Object System.Collections.Generic.List[object]
    foreach ($material in $existingDb.Materials) {
        [void]$materials.Add($material)
    }

    $lookup = @{}
    foreach ($key in $existingDb.Lookup.Keys) {
        $lookup[$key] = $existingDb.Lookup[$key]
    }

    $idLookup = Get-ExistingMaterialIdLookup -Materials $existingDb.Materials
    $nextId = [int]$existingDb.MaxId
    $insertedCount = 0
    $updatedCount = 0
    $skippedCount = 0
    $warningCount = 0
    $errorCount = 0

    for ($index = 0; $index -lt $sourceData.Rows.Count; $index++) {
        $row = $sourceData.Rows[$index]
        $rowNumber = if ($sourceData.RowNumbers.Count -gt $index) { [int]$sourceData.RowNumbers[$index] } else { $index + 2 }
        $matnr = Get-SourceColumnValue -Row $row -ColumnMap $resolvedPreset.Columns -Key 'matnr'
        $canonicalKey = if ([string]::IsNullOrWhiteSpace($matnr)) { '' } else { Get-CanonicalKey -Type 'matnr' -Value $matnr }
        $existingMaterial = $null
        if (-not [string]::IsNullOrWhiteSpace($canonicalKey) -and $lookup.ContainsKey($canonicalKey)) {
            $existingMaterial = $materials[$lookup[$canonicalKey]]
        }

        $importData = Convert-SapRowToImportData -Row $row -Preset $resolvedPreset -RowNumber $rowNumber -ExistingMaterial $existingMaterial -DefaultUnitCode $defaultUnitCode -ValidUnitCodes $validUnitCodes
        foreach ($warning in @($importData.Warnings)) {
            Write-ImportLog $warning 'WARNING'
            $warningCount++
        }

        if ($importData.ShouldSkip) {
            $skippedCount++
            continue
        }

        if ($existingMaterial) {
            $existingIndex = [int]$lookup[$canonicalKey]
            $id = [int]$existingMaterial.id
            $mergedRecord = Merge-MaterialRecord -ImportData $importData -Id $id
            $materials[$existingIndex] = $mergedRecord
            $idLookup[$id] = $existingIndex
            $updatedCount++
            Write-ImportLog "Zeile $rowNumber - aktualisiert: $($importData.CanonicalKey) (ID $id)" 'INFO'
        }
        else {
            $nextId++
            while ($idLookup.ContainsKey($nextId)) {
                $nextId++
            }

            $id = [int]$nextId
            $newRecord = Merge-MaterialRecord -ImportData $importData -Id $id
            [void]$materials.Add($newRecord)
            $newIndex = $materials.Count - 1
            $lookup[$importData.CanonicalKey] = $newIndex
            $idLookup[$id] = $newIndex
            $insertedCount++
            Write-ImportLog "Zeile $rowNumber - importiert: $($importData.CanonicalKey) (ID $id)" 'INFO'
        }
    }

    try {
        $backupPath = Backup-DatabaseFile -Path $DbPath
        if ($backupPath) {
            Write-ImportLog "Backup erstellt: $(Split-Path $backupPath -Leaf)" 'INFO'
        }

        Save-DatabaseFile -Path $DbPath -Materials $materials.ToArray()
        Write-ImportLog "SAP-Wartungsimport abgeschlossen - $($materials.Count) Materialien gespeichert" 'SUCCESS'
        Write-ImportLog "Zusammenfassung: Neu=$insertedCount, Aktualisiert=$updatedCount, Uebersprungen=$skippedCount, Warnungen=$warningCount, Fehler=$errorCount" 'INFO'

        if (-not $SuppressSuccessMessage) {
            [System.Windows.Forms.MessageBox]::Show("SAP-Wartungsimport abgeschlossen!`nNeu: $insertedCount`nAktualisiert: $updatedCount`nUebersprungen: $skippedCount`nWarnungen: $warningCount`nLog: $LogFile", 'Erfolg', 'OK', 'Information')
        }
    }
    catch {
        $errorCount++
        Write-ImportLog "Fehler beim Speichern der JSON: $($_.Exception.Message)" 'ERROR'
    }
}

function Start-InitialImport {
    param(
        [string]$SourceFile,
        [switch]$SuppressSuccessMessage
    )

    if (!(Test-Path $SourceFile)) {
        Write-ImportLog "Datei nicht gefunden: $SourceFile" 'ERROR'
        return
    }

    Write-ImportLog "Starte Import aus $SourceFile ..." 'INFO'

    $csvData = $null
    try {
        $csvData = Import-CsvData -Path $SourceFile
        Write-ImportLog "CSV geladen - $($csvData.Rows.Count) Zeilen gefunden (Encoding: $($csvData.EncodingName))" 'INFO'
    }
    catch {
        Write-ImportLog "CSV-Ladefehler: $($_.Exception.Message)" 'ERROR'
        return
    }

    $lookupFile = $null
    try {
        $lookupFile = Read-LookupFile -Path $LookupPath
        Write-ImportLog "Lookup-Datei geladen: $LookupPath" 'INFO'
    }
    catch {
        Write-ImportLog "Lookup-Datei nicht lesbar: $($_.Exception.Message)" 'ERROR'
        return
    }

    $lookupValidationResults = @(
        (Test-LookupCodesMatchDefinitions -LookupName 'responsibility_codes' -LookupCodes (ConvertTo-CodeSet $lookupFile.responsibility_codes) -Definitions (Get-ResponsibilityDefinitions))
        (Test-LookupCodesMatchDefinitions -LookupName 'assignment_tags' -LookupCodes (ConvertTo-CodeSet $lookupFile.assignment_tags) -Definitions (Get-AssignmentTagDefinitions))
        (Test-LookupCodesMatchDefinitions -LookupName 'hazmat_flags' -LookupCodes (ConvertTo-CodeSet $lookupFile.hazmat_flags) -Definitions (Get-HazmatFlagDefinitions))
    )
    foreach ($validationResult in $lookupValidationResults) {
        if (-not $validationResult.IsValid) {
            if ($validationResult.MissingCodes.Count -gt 0) {
                Write-ImportLog "Lookup-Datei unvollstaendig ($($validationResult.LookupName)): fehlende Codes: $($validationResult.MissingCodes -join ', ')" 'ERROR'
            }

            if ($validationResult.UnknownCodes.Count -gt 0) {
                Write-ImportLog "Lookup-Datei enthaelt unbekannte Codes ($($validationResult.LookupName)): $($validationResult.UnknownCodes -join ', ')" 'ERROR'
            }

            return
        }
    }

    if ($null -eq $csvData.Rows -or $csvData.Rows.Count -eq 0) {
        Write-ImportLog 'Keine Datenzeilen' 'ERROR'
        return
    }

    $headers = @($csvData.Rows[0].PSObject.Properties.Name)
    Write-ImportLog "Header erkannt: $($headers -join ' | ')" 'INFO'

    $headerResolution = Resolve-HeaderMap -Headers $headers -Definitions (Get-HeaderDefinitions)
    Write-ImportLog "Gefundene Headerzuordnung: $($headerResolution.HeaderMap.Keys -join ', ')" 'DEBUG'

    if ($headerResolution.MissingRequired.Count -gt 0) {
        Write-ImportLog "Kritischer Fehler: Pflichtspalten fehlen: $($headerResolution.MissingRequired -join ', ')" 'ERROR'
        return
    }

    if ($headerResolution.MissingOptional.Count -gt 0) {
        Write-ImportLog "Optionale Spalten fehlen: $($headerResolution.MissingOptional -join ', ')" 'WARNING'
    }

    $duplicateKeys = Test-DuplicateImportKeys -Rows $csvData.Rows -HeaderMap $headerResolution.HeaderMap
    if ($duplicateKeys.Count -gt 0) {
        foreach ($duplicate in $duplicateKeys) {
            Write-ImportLog "Doppelte CSV-Schluessel gefunden: $duplicate" 'ERROR'
        }
        Write-ImportLog 'Import abgebrochen - doppelte Schluessel im CSV' 'ERROR'
        return
    }

    $duplicateIds = Test-DuplicateImportIds -Rows $csvData.Rows -HeaderMap $headerResolution.HeaderMap
    if ($duplicateIds.Count -gt 0) {
        foreach ($duplicate in $duplicateIds) {
            Write-ImportLog "Doppelte CSV-IDs gefunden: $duplicate" 'ERROR'
        }
        Write-ImportLog 'Import abgebrochen - doppelte IDs im CSV' 'ERROR'
        return
    }

    $existingDb = $null
    try {
        $existingDb = Read-ExistingDatabase -Path $DbPath
        Write-ImportLog "Vorhandene Datenbank geladen - $($existingDb.Materials.Count) Materialien" 'INFO'
    }
    catch {
        Write-ImportLog "Vorhandene Datenbank nicht lesbar: $($_.Exception.Message)" 'ERROR'
        return
    }

    $materials = New-Object System.Collections.Generic.List[object]
    foreach ($material in $existingDb.Materials) {
        [void]$materials.Add($material)
    }

    $lookup = @{}
    foreach ($key in $existingDb.Lookup.Keys) {
        $lookup[$key] = $existingDb.Lookup[$key]
    }

    $idLookup = Get-ExistingMaterialIdLookup -Materials $existingDb.Materials

    $nextId = [int]$existingDb.MaxId
    $insertedCount = 0
    $updatedCount = 0
    $skippedCount = 0
    $warningCount = 0
    $errorCount = 0

    for ($index = 0; $index -lt $csvData.Rows.Count; $index++) {
        $rowNumber = $index + 2
        $row = $csvData.Rows[$index]
        $primaryValue = Get-CellValue -Row $row -HeaderMap $headerResolution.HeaderMap -Key 'matnr_main'
        $canonicalKey = Get-CanonicalKey -Type 'matnr' -Value $primaryValue
        $existingMaterial = $null
        if ($lookup.ContainsKey($canonicalKey)) {
            $existingMaterial = $materials[$lookup[$canonicalKey]]
        }

        $importData = Convert-RowToImportData -Row $row -HeaderMap $headerResolution.HeaderMap -RowNumber $rowNumber -ExistingMaterial $existingMaterial
        foreach ($warning in $importData.Warnings) {
            Write-ImportLog $warning 'WARNING'
            $warningCount++
        }

        if ($importData.ShouldSkip) {
            $skippedCount++
            continue
        }

        if ($existingMaterial) {
            $id = if ([int]$importData.ResolvedId -gt 0) { [int]$importData.ResolvedId } else { [int]$existingMaterial.id }
            $existingIndex = [int]$lookup[$canonicalKey]
            if ($idLookup.ContainsKey($id) -and [int]$idLookup[$id] -ne $existingIndex) {
                Write-ImportLog "Zeile $rowNumber - ID-Konflikt: ID $id ist bereits einem anderen Material zugewiesen" 'ERROR'
                $errorCount++
                $skippedCount++
                continue
            }

            $previousId = [int]$existingMaterial.id
            $mergedRecord = Merge-MaterialRecord -ImportData $importData -Id $id
            $materials[$existingIndex] = $mergedRecord
            if ($previousId -ne $id -and $idLookup.ContainsKey($previousId)) {
                [void]$idLookup.Remove($previousId)
            }
            $idLookup[$id] = $existingIndex
            if ($id -gt $nextId) {
                $nextId = $id
            }
            $updatedCount++
            Write-ImportLog "Zeile $rowNumber - aktualisiert: $($importData.CanonicalKey) (ID $id)" 'INFO'
        }
        else {
            $id = [int]$importData.ResolvedId
            if ($id -gt 0 -and $idLookup.ContainsKey($id)) {
                Write-ImportLog "Zeile $rowNumber - ID-Konflikt: ID $id ist bereits vorhanden" 'ERROR'
                $errorCount++
                $skippedCount++
                continue
            }

            if ($id -le 0) {
                $nextId++
                while ($idLookup.ContainsKey($nextId)) {
                    $nextId++
                }
                $id = $nextId
            }
            elseif ($id -gt $nextId) {
                $nextId = $id
            }

            $newRecord = Merge-MaterialRecord -ImportData $importData -Id $id
            [void]$materials.Add($newRecord)
            $newIndex = $materials.Count - 1
            $lookup[$importData.CanonicalKey] = $newIndex
            $idLookup[$id] = $newIndex
            $insertedCount++
            Write-ImportLog "Zeile $rowNumber - importiert: $($importData.CanonicalKey) (ID $id)" 'INFO'
        }
    }

    try {
        $backupPath = Backup-DatabaseFile -Path $DbPath
        if ($backupPath) {
            Write-ImportLog "Backup erstellt: $(Split-Path $backupPath -Leaf)" 'INFO'
        }

        Save-DatabaseFile -Path $DbPath -Materials $materials.ToArray()
        Write-ImportLog "Import abgeschlossen - $($materials.Count) Materialien gespeichert" 'SUCCESS'
        Write-ImportLog "Zusammenfassung: Neu=$insertedCount, Aktualisiert=$updatedCount, Uebersprungen=$skippedCount, Warnungen=$warningCount, Fehler=$errorCount" 'INFO'

        if (-not $SuppressSuccessMessage) {
            [System.Windows.Forms.MessageBox]::Show("Import abgeschlossen!`nNeu: $insertedCount`nAktualisiert: $updatedCount`nUebersprungen: $skippedCount`nWarnungen: $warningCount`nLog: $LogFile", 'Erfolg', 'OK', 'Information')
        }
    }
    catch {
        $errorCount++
        Write-ImportLog "Fehler beim Speichern der JSON: $($_.Exception.Message)" 'ERROR'
    }
}

function Get-ImportToolUiXaml {
    return @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Verlegepaket Datenimport (PS 5.1)"
        Height="900"
        Width="1120"
        WindowStartupLocation="CenterScreen"
        Background="#F5F5F5"
        FontFamily="Segoe UI"
        FontSize="11">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="320"/>
        </Grid.RowDefinitions>

        <TextBlock Text="Verlegepaket Datenimport"
                   FontSize="24"
                   FontWeight="Bold"
                   Foreground="#2C3E50"
                   Margin="0,0,0,10"/>

        <TabControl Grid.Row="1"
                    Margin="0,15,0,15"
                    Background="Transparent"
                    BorderBrush="#D6DBDF">
            <TabItem Header="Legacy Initial Import">
                <Grid Margin="15">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <Border BorderThickness="1"
                            BorderBrush="#E0E0E0"
                            CornerRadius="5"
                            Padding="15"
                            Background="White"
                            Margin="0,0,0,15">
                        <StackPanel>
                            <TextBlock Text="Quelldatei" FontWeight="Bold" Foreground="#34495E"/>
                            <Grid Margin="0,10,0,0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="130"/>
                                </Grid.ColumnDefinitions>
                                <TextBox x:Name="txtFile"
                                         Grid.Column="0"
                                         Padding="10"
                                         Background="White"
                                         BorderThickness="1"
                                         BorderBrush="#BDC3C7"
                                         IsReadOnly="True"
                                         Foreground="#2C3E50"/>
                                <Button x:Name="btnBrowse"
                                        Grid.Column="1"
                                        Margin="10,0,0,0"
                                        Content="Durchsuchen..."
                                        Background="#3498DB"
                                        Foreground="White"
                                        FontWeight="Bold"
                                        Cursor="Hand"
                                        Padding="10"/>
                            </Grid>
                        </StackPanel>
                    </Border>

                    <Button Grid.Row="1"
                            x:Name="btnImport"
                            Content="Legacy-Import starten"
                            Background="#27AE60"
                            Foreground="White"
                            FontWeight="Bold"
                            FontSize="13"
                            Padding="15,12"
                            Height="45"
                            Width="240"
                            HorizontalAlignment="Left"
                            Cursor="Hand"
                            Margin="0,0,0,15"/>

                    <Border Grid.Row="2"
                            BorderThickness="1"
                            BorderBrush="#E0E0E0"
                            CornerRadius="5"
                            Padding="15"
                            Background="White">
                        <TextBlock Text="Beibehaltener Spezialimport fuer die bisherige Initialliste mit fester Header-Erkennung und unveraendertem Merge-Verhalten."
                                   Foreground="#5D6D7E"
                                   TextWrapping="Wrap"/>
                    </Border>
                </Grid>
            </TabItem>

            <TabItem Header="SAP Maintenance Import">
                <Grid Margin="15">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <Border Grid.Row="0"
                            BorderThickness="1"
                            BorderBrush="#E0E0E0"
                            CornerRadius="5"
                            Padding="15"
                            Background="White"
                            Margin="0,0,0,15">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="130"/>
                                <ColumnDefinition Width="140"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.ColumnSpan="3">
                                <TextBlock Text="SAP-Quelldatei" FontWeight="Bold" Foreground="#34495E"/>
                                <TextBlock Text="Unterstuetzt CSV, TXT und XLSX. MATNR ist der einzige Match-Schluessel."
                                           Foreground="#5D6D7E"
                                           Margin="0,3,0,0"/>
                            </StackPanel>
                            <TextBox x:Name="txtSapFile"
                                     Grid.Column="0"
                                     Margin="0,42,0,0"
                                     Padding="10"
                                     Background="White"
                                     BorderThickness="1"
                                     BorderBrush="#BDC3C7"
                                     IsReadOnly="True"
                                     Foreground="#2C3E50"/>
                            <Button x:Name="btnSapBrowse"
                                    Grid.Column="1"
                                    Margin="10,42,0,0"
                                    Content="Durchsuchen..."
                                    Background="#3498DB"
                                    Foreground="White"
                                    FontWeight="Bold"
                                    Cursor="Hand"
                                    Padding="10"/>
                            <Button x:Name="btnSapLoadHeaders"
                                    Grid.Column="2"
                                    Margin="10,42,0,0"
                                    Content="Spalten lesen"
                                    Background="#5D6D7E"
                                    Foreground="White"
                                    FontWeight="Bold"
                                    Cursor="Hand"
                                    Padding="10"/>
                        </Grid>
                    </Border>

                    <Border Grid.Row="1"
                            BorderThickness="1"
                            BorderBrush="#E0E0E0"
                            CornerRadius="5"
                            Padding="15"
                            Background="White"
                            Margin="0,0,0,15">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="120"/>
                                <ColumnDefinition Width="160"/>
                                <ColumnDefinition Width="120"/>
                                <ColumnDefinition Width="160"/>
                                <ColumnDefinition Width="120"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>

                            <TextBlock Grid.Row="0" Grid.Column="0" VerticalAlignment="Center" Foreground="#34495E" Text="Preset"/>
                            <ComboBox x:Name="cmbSapPreset"
                                      Grid.Row="0"
                                      Grid.Column="1"
                                      Margin="0,0,10,8"/>
                            <Button x:Name="btnSapLoadPreset"
                                    Grid.Row="0"
                                    Grid.Column="2"
                                    Margin="0,0,10,8"
                                    Content="Preset laden"
                                    Background="#5D6D7E"
                                    Foreground="White"
                                    FontWeight="Bold"
                                    Cursor="Hand"
                                    Padding="8,6"/>
                            <TextBlock Grid.Row="0" Grid.Column="3" VerticalAlignment="Center" Foreground="#34495E" Text="Preset-Name"/>
                            <TextBox x:Name="txtSapPresetName"
                                     Grid.Row="0"
                                     Grid.Column="4"
                                     Grid.ColumnSpan="2"
                                     Margin="0,0,0,8"
                                     Padding="8"/>

                            <TextBlock Grid.Row="1" Grid.Column="0" VerticalAlignment="Center" Foreground="#34495E" Text="Header-Zeile"/>
                            <TextBox x:Name="txtSapHeaderRow"
                                     Grid.Row="1"
                                     Grid.Column="1"
                                     Margin="0,0,10,8"
                                     Padding="8"
                                     Text="1"/>
                            <TextBlock Grid.Row="1" Grid.Column="2" VerticalAlignment="Center" Foreground="#34495E" Text="Worksheet"/>
                            <TextBox x:Name="txtSapWorksheet"
                                     Grid.Row="1"
                                     Grid.Column="3"
                                     Margin="0,0,10,8"
                                     Padding="8"/>
                            <TextBlock Grid.Row="1" Grid.Column="4" VerticalAlignment="Center" Foreground="#34495E" Text="Delimiter"/>
                            <TextBox x:Name="txtSapDelimiter"
                                     Grid.Row="1"
                                     Grid.Column="5"
                                     Margin="0,0,0,8"
                                     Padding="8"
                                     Text=";"/>

                            <StackPanel Grid.Row="2"
                                        Grid.Column="0"
                                        Grid.ColumnSpan="6"
                                        Orientation="Horizontal">
                                <Button x:Name="btnSapPresetEditor"
                                        Content="Preset aus Datei bearbeiten"
                                        Background="#2874A6"
                                        Foreground="White"
                                        FontWeight="Bold"
                                        Cursor="Hand"
                                        Padding="10,8"
                                        Margin="0,0,10,0"/>
                                <Button x:Name="btnSapSavePreset"
                                        Content="Preset speichern"
                                        Background="#27AE60"
                                        Foreground="White"
                                        FontWeight="Bold"
                                        Cursor="Hand"
                                        Padding="10,8"
                                        Margin="0,0,10,0"/>
                                <Button x:Name="btnSapDeletePreset"
                                        Content="Preset loeschen"
                                        Background="#C0392B"
                                        Foreground="White"
                                        FontWeight="Bold"
                                        Cursor="Hand"
                                        Padding="10,8"/>
                            </StackPanel>
                        </Grid>
                    </Border>

                    <Border Grid.Row="2"
                            BorderThickness="1"
                            BorderBrush="#E0E0E0"
                            CornerRadius="5"
                            Padding="12"
                            Background="White"
                            Margin="0,0,0,15">
                        <TextBlock x:Name="txtSapSourceInfo"
                                   Foreground="#5D6D7E"
                                   Text="Noch keine SAP-Datei eingelesen."
                                   TextWrapping="Wrap"/>
                    </Border>

                    <Border Grid.Row="3"
                            BorderThickness="1"
                            BorderBrush="#E0E0E0"
                            CornerRadius="5"
                            Padding="12"
                            Background="White"
                            Margin="0,0,0,15">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <TextBlock Text="Spaltenzuordnung"
                                       FontWeight="Bold"
                                       Foreground="#34495E"
                                       Margin="0,0,0,10"/>
                            <DataGrid x:Name="gridSapMappings"
                                      Grid.Row="1"
                                      AutoGenerateColumns="False"
                                      CanUserAddRows="False"
                                      CanUserDeleteRows="False"
                                      CanUserResizeRows="False"
                                      HeadersVisibility="Column"
                                      RowHeaderWidth="0"
                                      SelectionMode="Single"
                                      GridLinesVisibility="Horizontal"
                                      BorderThickness="1"
                                      BorderBrush="#D5DBDB"
                                      Background="White">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Datenbankfeld"
                                                        Binding="{Binding TargetLabel}"
                                                        IsReadOnly="True"
                                                        Width="260"/>
                                    <DataGridTextColumn Header="Pflicht"
                                                        Binding="{Binding RequiredLabel}"
                                                        IsReadOnly="True"
                                                        Width="80"/>
                                    <DataGridTemplateColumn Header="Importspalte"
                                                            Width="*">
                                        <DataGridTemplateColumn.CellTemplate>
                                            <DataTemplate>
                                                <TextBlock Text="{Binding SelectedHeader}"
                                                           VerticalAlignment="Center"/>
                                            </DataTemplate>
                                        </DataGridTemplateColumn.CellTemplate>
                                        <DataGridTemplateColumn.CellEditingTemplate>
                                            <DataTemplate>
                                                <ComboBox ItemsSource="{Binding Options}"
                                                          SelectedItem="{Binding SelectedHeader, UpdateSourceTrigger=PropertyChanged}"
                                                          IsEditable="False"/>
                                            </DataTemplate>
                                        </DataGridTemplateColumn.CellEditingTemplate>
                                    </DataGridTemplateColumn>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Grid>
                    </Border>

                    <Button Grid.Row="4"
                            x:Name="btnSapImport"
                            Content="SAP-Wartungsimport starten"
                            Background="#1E8449"
                            Foreground="White"
                            FontWeight="Bold"
                            FontSize="13"
                            Padding="15,12"
                            Height="45"
                            Width="280"
                            HorizontalAlignment="Left"
                            Cursor="Hand"/>
                </Grid>
            </TabItem>
        </TabControl>

        <Border Grid.Row="2"
                BorderThickness="1"
                BorderBrush="#E0E0E0"
                CornerRadius="5"
                Padding="15"
                Background="White">
            <StackPanel>
                <TextBlock Text="Importprotokoll" FontWeight="Bold" Foreground="#34495E"/>
                <TextBox x:Name="txtLog"
                         Height="260"
                         Padding="10"
                         Background="#ECF0F1"
                         Foreground="#2C3E50"
                         FontFamily="Consolas"
                         FontSize="10"
                         IsReadOnly="True"
                         VerticalScrollBarVisibility="Auto"
                         TextWrapping="Wrap"
                         BorderThickness="1"
                         BorderBrush="#BDC3C7"
                         Margin="0,10,0,0"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
}

function Start-ImportToolUi {
    $reader = New-Object System.Xml.XmlNodeReader([xml](Get-ImportToolUiXaml))
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $txtFile = $window.FindName('txtFile')
    $btnBrowse = $window.FindName('btnBrowse')
    $btnImport = $window.FindName('btnImport')
    $txtLog = $window.FindName('txtLog')
    $txtSapFile = $window.FindName('txtSapFile')
    $btnSapBrowse = $window.FindName('btnSapBrowse')
    $btnSapLoadHeaders = $window.FindName('btnSapLoadHeaders')
    $cmbSapPreset = $window.FindName('cmbSapPreset')
    $btnSapLoadPreset = $window.FindName('btnSapLoadPreset')
    $txtSapPresetName = $window.FindName('txtSapPresetName')
    $txtSapHeaderRow = $window.FindName('txtSapHeaderRow')
    $txtSapWorksheet = $window.FindName('txtSapWorksheet')
    $txtSapDelimiter = $window.FindName('txtSapDelimiter')
    $btnSapPresetEditor = $window.FindName('btnSapPresetEditor')
    $btnSapSavePreset = $window.FindName('btnSapSavePreset')
    $btnSapDeletePreset = $window.FindName('btnSapDeletePreset')
    $txtSapSourceInfo = $window.FindName('txtSapSourceInfo')
    $gridSapMappings = $window.FindName('gridSapMappings')
    $btnSapImport = $window.FindName('btnSapImport')

    $global:LogBox = $txtLog

    $sapState = @{
        Headers = @()
        FileType = ''
        ApplyingPreset = $false
    }

    $newMappingRows = {
        $collection = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
        foreach ($definition in @(Get-SapImportFieldDefinitions)) {
            [void]$collection.Add([pscustomobject]@{
                    TargetKey      = $definition.Key
                    TargetLabel    = $definition.Label
                    RequiredLabel  = $(if ($definition.Required) { 'Ja' } else { '' })
                    SelectedHeader = ''
                    Options        = @('')
                })
        }

        return ,$collection
    }

    $gridSapMappings.ItemsSource = & $newMappingRows

    $getCurrentMappingRows = {
        return @($gridSapMappings.ItemsSource)
    }

    $refreshMappingGrid = {
        $gridSapMappings.Items.Refresh()
    }

    $applyHeadersToMappingRows = {
        param([string[]]$Headers)

        $sapState.Headers = @($Headers)
        foreach ($row in @(& $getCurrentMappingRows)) {
            $selectedHeader = Get-NormalizedString $row.SelectedHeader
            $options = New-Object System.Collections.Generic.List[string]
            [void]$options.Add('')
            foreach ($header in @($Headers)) {
                [void]$options.Add((Get-NormalizedString $header))
            }

            if (-not [string]::IsNullOrWhiteSpace($selectedHeader) -and $options -notcontains $selectedHeader) {
                [void]$options.Add($selectedHeader)
            }

            $row.Options = @($options.ToArray() | Select-Object -Unique)
            if (-not [string]::IsNullOrWhiteSpace($selectedHeader) -and @($row.Options) -contains $selectedHeader) {
                $row.SelectedHeader = $selectedHeader
            }
            else {
                $row.SelectedHeader = ''
            }
        }

        & $refreshMappingGrid
    }

    $getColumnMapFromGrid = {
        $columns = [ordered]@{}
        foreach ($definition in @(Get-SapImportFieldDefinitions)) {
            $row = @((& $getCurrentMappingRows) | Where-Object { $_.TargetKey -eq $definition.Key }) | Select-Object -First 1
            $columns[$definition.Key] = $(if ($null -eq $row) { $null } else { ConvertTo-NullableString $row.SelectedHeader })
        }

        return [pscustomobject]$columns
    }

    $applyColumnMapToGrid = {
        param($ColumnMap)

        foreach ($definition in @(Get-SapImportFieldDefinitions)) {
            $row = @((& $getCurrentMappingRows) | Where-Object { $_.TargetKey -eq $definition.Key }) | Select-Object -First 1
            if ($null -eq $row) {
                continue
            }

            $selectedHeader = $null
            if ($null -ne $ColumnMap -and $null -ne $ColumnMap.PSObject.Properties[$definition.Key]) {
                $selectedHeader = ConvertTo-NullableString $ColumnMap.PSObject.Properties[$definition.Key].Value
            }

            $row.SelectedHeader = $(if ($null -eq $selectedHeader) { '' } else { $selectedHeader })
        }

        & $applyHeadersToMappingRows -Headers $sapState.Headers
    }

    $refreshSapPresetList = {
        param([AllowNull()][string]$PreferredName)

        $store = Read-DataImportPresetStore -Path $DataImportPresetPath
        $names = @($store.presets | Sort-Object Name | ForEach-Object { Get-NormalizedString $_.Name })
        $cmbSapPreset.ItemsSource = $names
        if (-not [string]::IsNullOrWhiteSpace((Get-NormalizedString $PreferredName)) -and $names -contains $PreferredName) {
            $cmbSapPreset.SelectedItem = $PreferredName
        }
        elseif ($names.Count -gt 0) {
            $cmbSapPreset.SelectedIndex = 0
        }
        else {
            $cmbSapPreset.SelectedItem = $null
        }
    }

    $getSapPresetByName = {
        param([AllowNull()][string]$Name)

        $presetName = Get-NormalizedString $Name
        if ([string]::IsNullOrWhiteSpace($presetName)) {
            return $null
        }

        $store = Read-DataImportPresetStore -Path $DataImportPresetPath
        return (@($store.presets | Where-Object { (Get-NormalizedString $_.Name) -eq $presetName }) | Select-Object -First 1)
    }

    $buildSapPresetFromUi = {
        param([switch]$IncludeName)

        $sourcePath = Get-NormalizedString $txtSapFile.Text
        if ([string]::IsNullOrWhiteSpace($sourcePath)) {
            throw 'Bitte zuerst eine SAP-Datei auswaehlen.'
        }

        $fileType = Get-ImportFileTypeFromPath -Path $sourcePath
        if (@('xlsx', 'csv', 'txt') -notcontains $fileType) {
            throw "Dateityp '$fileType' wird nicht unterstuetzt."
        }

        $preset = [pscustomobject][ordered]@{
            Name           = $(if ($IncludeName) { $txtSapPresetName.Text } else { $null })
            FileType       = $fileType
            WorksheetName  = ConvertTo-NullableString $txtSapWorksheet.Text
            HeaderRowIndex = Get-NormalizedString $txtSapHeaderRow.Text
            Delimiter      = ConvertTo-NullableString $txtSapDelimiter.Text
            Columns        = & $getColumnMapFromGrid
        }

        return (ConvertTo-NormalizedDataImportPreset -Preset $preset -DefaultFileType $fileType)
    }

    $buildSapDraftFromUi = {
        $sourcePath = Get-NormalizedString $txtSapFile.Text
        if ([string]::IsNullOrWhiteSpace($sourcePath)) {
            throw 'Bitte zuerst eine SAP-Datei auswaehlen.'
        }

        $fileType = Get-ImportFileTypeFromPath -Path $sourcePath
        return [pscustomobject][ordered]@{
            Name           = $null
            FileType       = $fileType
            WorksheetName  = ConvertTo-NullableString $txtSapWorksheet.Text
            HeaderRowIndex = Get-NormalizedString $txtSapHeaderRow.Text
            Delimiter      = ConvertTo-NullableString $txtSapDelimiter.Text
            Columns        = & $getColumnMapFromGrid
        }
    }

    $loadSapHeaders = {
        $sourcePath = Get-NormalizedString $txtSapFile.Text
        if ([string]::IsNullOrWhiteSpace($sourcePath)) {
            throw 'Bitte zuerst eine SAP-Datei auswaehlen.'
        }

        $fileType = Get-ImportFileTypeFromPath -Path $sourcePath
        if (@('xlsx', 'csv', 'txt') -notcontains $fileType) {
            throw "Dateityp '$fileType' wird nicht unterstuetzt."
        }

        $headerRowIndex = 0
        if (-not [int]::TryParse((Get-NormalizedString $txtSapHeaderRow.Text), [ref]$headerRowIndex) -or $headerRowIndex -le 0) {
            throw 'Header-Zeile muss eine positive Zahl sein.'
        }

        $delimiter = Get-NormalizedString $txtSapDelimiter.Text
        if ([string]::IsNullOrWhiteSpace($delimiter)) {
            $delimiter = ';'
        }

        $sourceData = Read-GenericImportSource -Path $sourcePath -FileType $fileType -HeaderRowIndex $headerRowIndex -WorksheetName (ConvertTo-NullableString $txtSapWorksheet.Text) -Delimiter $delimiter
        $sapState.FileType = $fileType
        & $applyHeadersToMappingRows -Headers @($sourceData.Headers)

        $infoParts = New-Object System.Collections.Generic.List[string]
        [void]$infoParts.Add("Dateityp: $fileType")
        [void]$infoParts.Add("Header-Zeile: $headerRowIndex")
        [void]$infoParts.Add("Spalten: $($sourceData.Headers.Count)")
        [void]$infoParts.Add("Datenzeilen: $($sourceData.Rows.Count)")
        if (-not [string]::IsNullOrWhiteSpace((Get-NormalizedString $sourceData.Worksheet))) {
            [void]$infoParts.Add("Worksheet: $($sourceData.Worksheet)")
        }
        if (-not [string]::IsNullOrWhiteSpace((Get-NormalizedString $sourceData.EncodingName))) {
            [void]$infoParts.Add("Encoding: $($sourceData.EncodingName)")
        }

        $txtSapSourceInfo.Text = $infoParts -join '  |  '
        Write-ImportLog "SAP-Quelldatei eingelesen: $sourcePath" 'INFO'
        return $sourceData
    }

    $applyPresetToUi = {
        param($Preset)

        if ($null -eq $Preset) {
            return
        }

        $sapState.ApplyingPreset = $true
        try {
            $txtSapPresetName.Text = Get-NormalizedString $Preset.Name
            $txtSapHeaderRow.Text = [string]([int]$Preset.HeaderRowIndex)
            $txtSapWorksheet.Text = Get-NormalizedString $Preset.WorksheetName
            $txtSapDelimiter.Text = $(if ($Preset.FileType -eq 'xlsx') { '' } else { Get-NormalizedString $Preset.Delimiter })
            & $applyColumnMapToGrid -ColumnMap $Preset.Columns
        }
        finally {
            $sapState.ApplyingPreset = $false
        }
    }

    $OpenSapPresetEditorDialog = {
        $sourcePath = Get-NormalizedString $txtSapFile.Text
        if ([string]::IsNullOrWhiteSpace($sourcePath)) {
            [System.Windows.MessageBox]::Show('Bitte zuerst eine SAP-Datei auswaehlen.', 'Preset bearbeiten', 'OK', 'Warning') | Out-Null
            return
        }

        if ($sapState.Headers.Count -eq 0) {
            try {
                [void](& $loadSapHeaders)
            }
            catch {
                Write-ImportLog "Preset-Editor konnte nicht geoeffnet werden: $($_.Exception.Message)" 'ERROR'
                [System.Windows.MessageBox]::Show($_.Exception.Message, 'Preset bearbeiten', 'OK', 'Warning') | Out-Null
                return
            }
        }

        $selectedPresetName = Get-NormalizedString $cmbSapPreset.SelectedItem
        $initialDraft = & $getSapPresetByName -Name $selectedPresetName
        if ($null -eq $initialDraft) {
            $initialDraft = & $buildSapDraftFromUi
        }
        else {
            $initialDraft = Copy-DeepObject $initialDraft
        }

        $dialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SAP-Preset Editor"
        Height="760"
        Width="1080"
        WindowStartupLocation="CenterOwner"
        Background="#F5F5F5"
        FontFamily="Segoe UI"
        FontSize="11"
        ResizeMode="CanResize">
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Border Grid.Row="0"
                BorderThickness="1"
                BorderBrush="#D5DBDB"
                CornerRadius="5"
                Padding="12"
                Background="White"
                Margin="0,0,0,12">
            <DockPanel LastChildFill="True">
                <StackPanel DockPanel.Dock="Left">
                    <TextBlock Text="Preset auf Basis der geladenen Datei erstellen oder bearbeiten"
                               FontSize="18"
                               FontWeight="Bold"
                               Foreground="#2C3E50"/>
                    <TextBlock x:Name="txtPresetEditorSourceInfo"
                               Margin="0,4,0,0"
                               Foreground="#5D6D7E"
                               TextWrapping="Wrap"/>
                </StackPanel>
            </DockPanel>
        </Border>

        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="250"/>
                <ColumnDefinition Width="16"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <Border Grid.Column="0"
                    BorderThickness="1"
                    BorderBrush="#D5DBDB"
                    CornerRadius="5"
                    Padding="12"
                    Background="White">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                        <Button x:Name="btnPresetDialogNew"
                                Content="Neuer Entwurf"
                                Background="#2874A6"
                                Foreground="White"
                                FontWeight="Bold"
                                Padding="8,6"
                                Margin="0,0,8,0"/>
                        <Button x:Name="btnPresetDialogReload"
                                Content="Neu laden"
                                Background="#5D6D7E"
                                Foreground="White"
                                FontWeight="Bold"
                                Padding="8,6"/>
                    </StackPanel>
                    <ListBox x:Name="lbPresetDialogPresets"
                             Grid.Row="1"
                             DisplayMemberPath="Name"
                             BorderThickness="1"
                             BorderBrush="#D5DBDB"/>
                </Grid>
            </Border>

            <Border Grid.Column="2"
                    BorderThickness="1"
                    BorderBrush="#D5DBDB"
                    CornerRadius="5"
                    Padding="12"
                    Background="White">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <Grid Grid.Row="0" Margin="0,0,0,10">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="120"/>
                            <ColumnDefinition Width="180"/>
                            <ColumnDefinition Width="120"/>
                            <ColumnDefinition Width="180"/>
                            <ColumnDefinition Width="120"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>

                        <TextBlock Grid.Row="0" Grid.Column="0" VerticalAlignment="Center" Foreground="#34495E" Text="Preset-Name"/>
                        <TextBox x:Name="txtPresetDialogName" Grid.Row="0" Grid.Column="1" Margin="0,0,12,8" Padding="8"/>
                        <TextBlock Grid.Row="0" Grid.Column="2" VerticalAlignment="Center" Foreground="#34495E" Text="Dateityp"/>
                        <ComboBox x:Name="cmbPresetDialogFileType" Grid.Row="0" Grid.Column="3" Margin="0,0,12,8"/>
                        <TextBlock Grid.Row="0" Grid.Column="4" VerticalAlignment="Center" Foreground="#34495E" Text="Header-Zeile"/>
                        <TextBox x:Name="txtPresetDialogHeaderRow" Grid.Row="0" Grid.Column="5" Margin="0,0,0,8" Padding="8"/>

                        <TextBlock Grid.Row="1" Grid.Column="0" VerticalAlignment="Center" Foreground="#34495E" Text="Worksheet"/>
                        <TextBox x:Name="txtPresetDialogWorksheet" Grid.Row="1" Grid.Column="1" Margin="0,0,12,0" Padding="8"/>
                        <TextBlock Grid.Row="1" Grid.Column="2" VerticalAlignment="Center" Foreground="#34495E" Text="Delimiter"/>
                        <TextBox x:Name="txtPresetDialogDelimiter" Grid.Row="1" Grid.Column="3" Margin="0,0,12,0" Padding="8"/>
                        <Button x:Name="btnPresetDialogLoadHeaders"
                                Grid.Row="1"
                                Grid.Column="4"
                                Grid.ColumnSpan="2"
                                HorizontalAlignment="Left"
                                Content="Spalten aus Datei lesen"
                                Background="#5D6D7E"
                                Foreground="White"
                                FontWeight="Bold"
                                Padding="10,8"/>
                    </Grid>

                    <TextBlock x:Name="txtPresetDialogHeaderInfo"
                               Grid.Row="1"
                               Margin="0,0,0,10"
                               Foreground="#5D6D7E"
                               TextWrapping="Wrap"/>

                    <DataGrid x:Name="gridPresetDialogMappings"
                              Grid.Row="2"
                              AutoGenerateColumns="False"
                              CanUserAddRows="False"
                              CanUserDeleteRows="False"
                              CanUserResizeRows="False"
                              HeadersVisibility="Column"
                              RowHeaderWidth="0"
                              SelectionMode="Single"
                              GridLinesVisibility="Horizontal"
                              BorderThickness="1"
                              BorderBrush="#D5DBDB"
                              Background="White">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Datenbankfeld"
                                                Binding="{Binding TargetLabel}"
                                                IsReadOnly="True"
                                                Width="240"/>
                            <DataGridTextColumn Header="Pflicht"
                                                Binding="{Binding RequiredLabel}"
                                                IsReadOnly="True"
                                                Width="80"/>
                            <DataGridTemplateColumn Header="Importspalte"
                                                    Width="*">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBlock Text="{Binding SelectedHeader}"
                                                   VerticalAlignment="Center"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                                <DataGridTemplateColumn.CellEditingTemplate>
                                    <DataTemplate>
                                        <ComboBox ItemsSource="{Binding Options}"
                                                  SelectedItem="{Binding SelectedHeader, UpdateSourceTrigger=PropertyChanged}"
                                                  IsEditable="False"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellEditingTemplate>
                            </DataGridTemplateColumn>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </Border>
        </Grid>

        <DockPanel Grid.Row="2" LastChildFill="False" Margin="0,12,0,0">
            <TextBlock x:Name="txtPresetDialogStatus"
                       DockPanel.Dock="Left"
                       VerticalAlignment="Center"
                       Foreground="#5D6D7E"
                       Text="Bereit"/>
            <Button x:Name="btnPresetDialogClose"
                    DockPanel.Dock="Right"
                    Width="120"
                    Content="Schliessen"/>
            <Button x:Name="btnPresetDialogSave"
                    DockPanel.Dock="Right"
                    Width="140"
                    Margin="0,0,10,0"
                    Content="Preset speichern"
                    Background="#27AE60"
                    Foreground="White"
                    FontWeight="Bold"/>
        </DockPanel>
    </Grid>
</Window>
"@

        $dialogReader = New-Object System.Xml.XmlNodeReader([xml]$dialogXaml)
        $dialog = [Windows.Markup.XamlReader]::Load($dialogReader)
        $dialog.Owner = $window

        $txtPresetEditorSourceInfo = $dialog.FindName('txtPresetEditorSourceInfo')
        $btnPresetDialogNew = $dialog.FindName('btnPresetDialogNew')
        $btnPresetDialogReload = $dialog.FindName('btnPresetDialogReload')
        $lbPresetDialogPresets = $dialog.FindName('lbPresetDialogPresets')
        $txtPresetDialogName = $dialog.FindName('txtPresetDialogName')
        $cmbPresetDialogFileType = $dialog.FindName('cmbPresetDialogFileType')
        $txtPresetDialogHeaderRow = $dialog.FindName('txtPresetDialogHeaderRow')
        $txtPresetDialogWorksheet = $dialog.FindName('txtPresetDialogWorksheet')
        $txtPresetDialogDelimiter = $dialog.FindName('txtPresetDialogDelimiter')
        $btnPresetDialogLoadHeaders = $dialog.FindName('btnPresetDialogLoadHeaders')
        $txtPresetDialogHeaderInfo = $dialog.FindName('txtPresetDialogHeaderInfo')
        $gridPresetDialogMappings = $dialog.FindName('gridPresetDialogMappings')
        $txtPresetDialogStatus = $dialog.FindName('txtPresetDialogStatus')
        $btnPresetDialogClose = $dialog.FindName('btnPresetDialogClose')
        $btnPresetDialogSave = $dialog.FindName('btnPresetDialogSave')

        $cmbPresetDialogFileType.ItemsSource = @('csv', 'xlsx', 'txt')

        $dialogState = @{
            Headers = @($sapState.Headers)
            ApplyingPreset = $false
        }

        $setDialogStatus = {
            param([string]$Message, [string]$Level = 'Info')

            $txtPresetDialogStatus.Text = $Message
            switch ($Level) {
                'Error' { $txtPresetDialogStatus.Foreground = '#B91C1C' }
                'Success' { $txtPresetDialogStatus.Foreground = '#0F766E' }
                default { $txtPresetDialogStatus.Foreground = '#5D6D7E' }
            }
        }

        $newDialogMappingRows = {
            param([string[]]$Headers, $ColumnMap)

            $collection = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
            foreach ($definition in @(Get-SapImportFieldDefinitions)) {
                $selectedHeader = ''
                if ($null -ne $ColumnMap -and $null -ne $ColumnMap.PSObject.Properties[$definition.Key]) {
                    $selectedHeader = Get-NormalizedString $ColumnMap.PSObject.Properties[$definition.Key].Value
                }

                $options = New-Object System.Collections.Generic.List[string]
                [void]$options.Add('')
                foreach ($header in @($Headers)) {
                    [void]$options.Add((Get-NormalizedString $header))
                }

                if (-not [string]::IsNullOrWhiteSpace($selectedHeader) -and $options -notcontains $selectedHeader) {
                    [void]$options.Add($selectedHeader)
                }

                [void]$collection.Add([pscustomobject]@{
                        TargetKey      = $definition.Key
                        TargetLabel    = $definition.Label
                        RequiredLabel  = $(if ($definition.Required) { 'Ja' } else { '' })
                        SelectedHeader = $selectedHeader
                        Options        = @($options.ToArray() | Select-Object -Unique)
                    })
            }

            return ,$collection
        }

        $getDialogColumnMap = {
            $columns = [ordered]@{}
            foreach ($definition in @(Get-SapImportFieldDefinitions)) {
                $row = @($gridPresetDialogMappings.ItemsSource | Where-Object { $_.TargetKey -eq $definition.Key }) | Select-Object -First 1
                $columns[$definition.Key] = $(if ($null -eq $row) { $null } else { ConvertTo-NullableString $row.SelectedHeader })
            }

            return [pscustomobject]$columns
        }

        $refreshDialogHeaderInfo = {
            if ($dialogState.Headers.Count -eq 0) {
                $txtPresetDialogHeaderInfo.Text = 'Noch keine Importspalten geladen.'
                return
            }

            $txtPresetDialogHeaderInfo.Text = "Geladene Importspalten: $($dialogState.Headers.Count)  |  $($dialogState.Headers -join ' | ')"
        }

        $applyDialogHeadersToRows = {
            param([string[]]$Headers)

            $dialogState.Headers = @($Headers)
            foreach ($row in @($gridPresetDialogMappings.ItemsSource)) {
                $selectedHeader = Get-NormalizedString $row.SelectedHeader
                $options = New-Object System.Collections.Generic.List[string]
                [void]$options.Add('')
                foreach ($header in @($Headers)) {
                    [void]$options.Add((Get-NormalizedString $header))
                }

                if (-not [string]::IsNullOrWhiteSpace($selectedHeader) -and $options -notcontains $selectedHeader) {
                    [void]$options.Add($selectedHeader)
                }

                $row.Options = @($options.ToArray() | Select-Object -Unique)
                if (-not [string]::IsNullOrWhiteSpace($selectedHeader) -and @($row.Options) -contains $selectedHeader) {
                    $row.SelectedHeader = $selectedHeader
                }
                else {
                    $row.SelectedHeader = ''
                }
            }

            $gridPresetDialogMappings.Items.Refresh()
            & $refreshDialogHeaderInfo
        }

        $applyDraftToDialog = {
            param(
                $Draft,
                [switch]$BlankName
            )

            if ($null -eq $Draft) {
                return
            }

            $dialogState.ApplyingPreset = $true
            try {
                $resolvedHeaderRow = 1
                $parsedHeaderRow = 0
                if ([int]::TryParse((Get-NormalizedString $Draft.HeaderRowIndex), [ref]$parsedHeaderRow) -and $parsedHeaderRow -gt 0) {
                    $resolvedHeaderRow = $parsedHeaderRow
                }

                $txtPresetDialogName.Text = $(if ($BlankName) { '' } else { Get-NormalizedString $Draft.Name })
                $cmbPresetDialogFileType.SelectedItem = Get-NormalizedString $Draft.FileType
                $txtPresetDialogHeaderRow.Text = [string]$resolvedHeaderRow
                $txtPresetDialogWorksheet.Text = Get-NormalizedString $Draft.WorksheetName
                $txtPresetDialogDelimiter.Text = $(if ((Get-NormalizedString $Draft.FileType) -eq 'xlsx') { '' } else { Get-NormalizedString $Draft.Delimiter })
                $gridPresetDialogMappings.ItemsSource = & $newDialogMappingRows -Headers $dialogState.Headers -ColumnMap $Draft.Columns
                & $refreshDialogHeaderInfo
            }
            finally {
                $dialogState.ApplyingPreset = $false
            }
        }

        $refreshDialogPresetList = {
            $store = Read-DataImportPresetStore -Path $DataImportPresetPath
            $lbPresetDialogPresets.ItemsSource = @($store.presets | Sort-Object Name)
            $lbPresetDialogPresets.DisplayMemberPath = 'Name'
        }

        $buildDialogPreset = {
            $preset = [pscustomobject][ordered]@{
                Name           = Get-NormalizedString $txtPresetDialogName.Text
                FileType       = Get-NormalizedString $cmbPresetDialogFileType.Text
                HeaderRowIndex = Get-NormalizedString $txtPresetDialogHeaderRow.Text
                WorksheetName  = ConvertTo-NullableString $txtPresetDialogWorksheet.Text
                Delimiter      = ConvertTo-NullableString $txtPresetDialogDelimiter.Text
                Columns        = & $getDialogColumnMap
            }

            return (ConvertTo-NormalizedDataImportPreset -Preset $preset -DefaultFileType (Get-ImportFileTypeFromPath -Path $sourcePath))
        }

        $reloadDialogHeaders = {
            $fileType = Get-NormalizedString $cmbPresetDialogFileType.Text
            if ([string]::IsNullOrWhiteSpace($fileType)) {
                $fileType = Get-ImportFileTypeFromPath -Path $sourcePath
            }

            $headerRowIndex = 0
            if (-not [int]::TryParse((Get-NormalizedString $txtPresetDialogHeaderRow.Text), [ref]$headerRowIndex) -or $headerRowIndex -le 0) {
                throw 'Header-Zeile muss eine positive Zahl sein.'
            }

            $delimiter = Get-NormalizedString $txtPresetDialogDelimiter.Text
            if ([string]::IsNullOrWhiteSpace($delimiter)) {
                $delimiter = ';'
            }

            $sourceData = Read-GenericImportSource -Path $sourcePath -FileType $fileType -HeaderRowIndex $headerRowIndex -WorksheetName (ConvertTo-NullableString $txtPresetDialogWorksheet.Text) -Delimiter $delimiter
            & $applyDialogHeadersToRows -Headers @($sourceData.Headers)
            & $setDialogStatus -Message "Spalten aus Datei geladen ($($sourceData.Headers.Count))." -Level 'Success'
        }

        $sourceInfoParts = New-Object System.Collections.Generic.List[string]
        [void]$sourceInfoParts.Add("Datei: $sourcePath")
        [void]$sourceInfoParts.Add("Aktuelle Spalten: $($sapState.Headers.Count)")
        $txtPresetEditorSourceInfo.Text = $sourceInfoParts -join '  |  '

        $gridPresetDialogMappings.ItemsSource = & $newDialogMappingRows -Headers $dialogState.Headers -ColumnMap $initialDraft.Columns
        & $refreshDialogHeaderInfo
        & $refreshDialogPresetList
        & $applyDraftToDialog -Draft $initialDraft
        & $setDialogStatus -Message 'Preset-Editor bereit.'

        $lbPresetDialogPresets.Add_SelectionChanged({
                if ($dialogState.ApplyingPreset) {
                    return
                }

                $selected = $lbPresetDialogPresets.SelectedItem
                if ($null -eq $selected) {
                    return
                }

                & $applyDraftToDialog -Draft $selected
                & $setDialogStatus -Message "Preset '$((Get-NormalizedString $selected.Name))' geladen."
            })

        $btnPresetDialogNew.Add_Click({
                $draft = & $buildSapDraftFromUi
                $lbPresetDialogPresets.SelectedItem = $null
                & $applyDraftToDialog -Draft $draft -BlankName
                & $setDialogStatus -Message 'Neuer Entwurf auf Basis der aktuellen SAP-Zuordnung.'
            })

        $btnPresetDialogReload.Add_Click({
                try {
                    & $refreshDialogPresetList
                    & $setDialogStatus -Message 'Preset-Liste neu geladen.' -Level 'Success'
                }
                catch {
                    & $setDialogStatus -Message $_.Exception.Message -Level 'Error'
                }
            })

        $btnPresetDialogLoadHeaders.Add_Click({
                try {
                    & $reloadDialogHeaders
                }
                catch {
                    & $setDialogStatus -Message $_.Exception.Message -Level 'Error'
                }
            })

        $btnPresetDialogSave.Add_Click({
                try {
                    $savedPreset = Set-DataImportPreset -Preset (& $buildDialogPreset) -Path $DataImportPresetPath
                    & $refreshDialogPresetList
                    $lbPresetDialogPresets.SelectedItem = @($lbPresetDialogPresets.ItemsSource | Where-Object { (Get-NormalizedString $_.Name) -eq (Get-NormalizedString $savedPreset.Name) }) | Select-Object -First 1
                    & $refreshSapPresetList -PreferredName (Get-NormalizedString $savedPreset.Name)
                    & $applyPresetToUi -Preset $savedPreset
                    Write-ImportLog "Preset gespeichert (Dialog): $($savedPreset.Name)" 'INFO'
                    & $setDialogStatus -Message "Preset '$($savedPreset.Name)' gespeichert." -Level 'Success'
                }
                catch {
                    & $setDialogStatus -Message $_.Exception.Message -Level 'Error'
                }
            })

        $btnPresetDialogClose.Add_Click({
                $dialog.Close()
            })

        [void]$dialog.ShowDialog()
    }

    $btnBrowse.Add_Click({
            $ofd = New-Object System.Windows.Forms.OpenFileDialog
            $ofd.Filter = 'CSV/Text-Dateien (*.csv;*.txt)|*.csv;*.txt|Alle Dateien (*.*)|*.*'
            if ($ofd.ShowDialog() -eq 'OK') {
                $txtFile.Text = $ofd.FileName
            }
        })

    $btnImport.Add_Click({
            if ([string]::IsNullOrWhiteSpace($txtFile.Text)) {
                [System.Windows.MessageBox]::Show('Bitte eine Datei auswaehlen!', 'Hinweis', 'OK', 'Warning') | Out-Null
                return
            }

            $btnImport.IsEnabled = $false
            try {
                Start-InitialImport -SourceFile $txtFile.Text
            }
            finally {
                $btnImport.IsEnabled = $true
            }
        })

    $btnSapBrowse.Add_Click({
            $ofd = New-Object System.Windows.Forms.OpenFileDialog
            $ofd.Filter = 'SAP-Dateien (*.xlsx;*.csv;*.txt)|*.xlsx;*.csv;*.txt|Alle Dateien (*.*)|*.*'
            if ($ofd.ShowDialog() -eq 'OK') {
                $txtSapFile.Text = $ofd.FileName
                $fileType = Get-ImportFileTypeFromPath -Path $ofd.FileName
                if ($fileType -eq 'xlsx') {
                    $txtSapDelimiter.Text = ''
                }
                elseif ([string]::IsNullOrWhiteSpace((Get-NormalizedString $txtSapDelimiter.Text))) {
                    $txtSapDelimiter.Text = ';'
                }

                try {
                    [void](& $loadSapHeaders)
                }
                catch {
                    Write-ImportLog "SAP-Quelldatei konnte nicht eingelesen werden: $($_.Exception.Message)" 'ERROR'
                    $txtSapSourceInfo.Text = $_.Exception.Message
                }
            }
        })

    $btnSapLoadHeaders.Add_Click({
            try {
                [void](& $loadSapHeaders)
            }
            catch {
                Write-ImportLog "Spalten konnten nicht eingelesen werden: $($_.Exception.Message)" 'ERROR'
                [System.Windows.MessageBox]::Show($_.Exception.Message, 'SAP-Datei lesen', 'OK', 'Warning') | Out-Null
            }
        })

    $cmbSapPreset.Add_SelectionChanged({
            if ($sapState.ApplyingPreset) {
                return
            }

            if ($null -ne $cmbSapPreset.SelectedItem) {
                $txtSapPresetName.Text = Get-NormalizedString $cmbSapPreset.SelectedItem
            }
        })

    $btnSapLoadPreset.Add_Click({
            $presetName = Get-NormalizedString $cmbSapPreset.SelectedItem
            if ([string]::IsNullOrWhiteSpace($presetName)) {
                [System.Windows.MessageBox]::Show('Bitte zuerst ein Preset auswaehlen.', 'Preset laden', 'OK', 'Warning') | Out-Null
                return
            }

            try {
                $store = Read-DataImportPresetStore -Path $DataImportPresetPath
                $preset = @($store.presets | Where-Object { (Get-NormalizedString $_.Name) -eq $presetName }) | Select-Object -First 1
                if ($null -eq $preset) {
                    throw "Preset '$presetName' wurde nicht gefunden."
                }

                & $applyPresetToUi -Preset $preset
                Write-ImportLog "Preset geladen: $presetName" 'INFO'
            }
            catch {
                Write-ImportLog "Preset konnte nicht geladen werden: $($_.Exception.Message)" 'ERROR'
                [System.Windows.MessageBox]::Show($_.Exception.Message, 'Preset laden', 'OK', 'Error') | Out-Null
            }
        })

    $btnSapPresetEditor.Add_Click({
            & $OpenSapPresetEditorDialog
        })

    $btnSapSavePreset.Add_Click({
            try {
                $preset = & $buildSapPresetFromUi -IncludeName
                $savedPreset = Set-DataImportPreset -Preset $preset -Path $DataImportPresetPath
                & $refreshSapPresetList -PreferredName (Get-NormalizedString $savedPreset.Name)
                & $applyPresetToUi -Preset $savedPreset
                Write-ImportLog "Preset gespeichert: $($savedPreset.Name)" 'INFO'
            }
            catch {
                Write-ImportLog "Preset konnte nicht gespeichert werden: $($_.Exception.Message)" 'ERROR'
                [System.Windows.MessageBox]::Show($_.Exception.Message, 'Preset speichern', 'OK', 'Error') | Out-Null
            }
        })

    $btnSapDeletePreset.Add_Click({
            $presetName = Get-NormalizedString $txtSapPresetName.Text
            if ([string]::IsNullOrWhiteSpace($presetName)) {
                [System.Windows.MessageBox]::Show('Bitte einen Preset-Namen angeben oder auswaehlen.', 'Preset loeschen', 'OK', 'Warning') | Out-Null
                return
            }

            $result = [System.Windows.MessageBox]::Show("Preset '$presetName' loeschen?", 'Preset loeschen', 'YesNo', 'Warning')
            if ($result -ne 'Yes') {
                return
            }

            try {
                Remove-DataImportPreset -Name $presetName -Path $DataImportPresetPath
                $txtSapPresetName.Text = ''
                & $refreshSapPresetList -PreferredName $null
                Write-ImportLog "Preset geloescht: $presetName" 'INFO'
            }
            catch {
                Write-ImportLog "Preset konnte nicht geloescht werden: $($_.Exception.Message)" 'ERROR'
                [System.Windows.MessageBox]::Show($_.Exception.Message, 'Preset loeschen', 'OK', 'Error') | Out-Null
            }
        })

    $btnSapImport.Add_Click({
            try {
                $preset = & $buildSapPresetFromUi
                $btnSapImport.IsEnabled = $false
                Start-SapMaintenanceImport -SourceFile $txtSapFile.Text -Preset $preset
            }
            catch {
                Write-ImportLog "SAP-Wartungsimport konnte nicht gestartet werden: $($_.Exception.Message)" 'ERROR'
                [System.Windows.MessageBox]::Show($_.Exception.Message, 'SAP-Wartungsimport', 'OK', 'Warning') | Out-Null
            }
            finally {
                $btnSapImport.IsEnabled = $true
            }
        })

    Write-ImportLog "Tool gestartet - $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')" 'INFO'
    Write-ImportLog "Logdatei: $LogFile" 'INFO'
    Write-ImportLog "Datenbank: $DbPath" 'INFO'
    Write-ImportLog "Lookup-Datei: $LookupPath" 'INFO'
    Write-ImportLog "Preset-Datei SAP-Wartungsimport: $DataImportPresetPath" 'INFO'

    try {
        if (-not (Test-Path $DataImportPresetPath)) {
            Write-DataImportPresetStore -Path $DataImportPresetPath -Store (New-DefaultDataImportPresetStore)
        }
    }
    catch {
        Write-ImportLog "Preset-Datei konnte nicht initialisiert werden: $($_.Exception.Message)" 'ERROR'
    }

    & $refreshSapPresetList -PreferredName $null
    & $applyHeadersToMappingRows -Headers @()

    $window.ShowDialog() | Out-Null
}

if ($MyInvocation.InvocationName -ne '.') {
    Start-ImportToolUi
}
