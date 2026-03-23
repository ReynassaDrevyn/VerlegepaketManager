# Function_StockComparison.ps1
# Standalone SOLL/IST comparison engine for the material database.

$Script:ProjectRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path $MyInvocation.MyCommand.Path -Parent }
$Script:DefaultDatabasePath = Join-Path $Script:ProjectRoot 'Core\db_verlegepaket.json'
$Script:DefaultPresetStorePath = Join-Path $Script:ProjectRoot 'Core\compare_presets.json'
$Script:DatabaseSchemaVersion = 2
$Script:PresetSchemaVersion = 1

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

function ConvertTo-UniqueStringArray {
    param([AllowNull()][object[]]$Values)

    $seen = @{}
    $result = New-Object System.Collections.Generic.List[string]
    foreach ($value in @(ConvertTo-ObjectArray $Values)) {
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

function Get-DeepPropertyValue {
    param(
        [AllowNull()]$Object,
        [Parameter(Mandatory = $true)][string]$Path,
        $Default = $null
    )

    $current = $Object
    foreach ($segment in ($Path -split '\.')) {
        if ($null -eq $current) {
            return $Default
        }

        $property = $current.PSObject.Properties[$segment]
        if ($null -eq $property) {
            return $Default
        }

        $current = $property.Value
    }

    if ($null -eq $current) {
        return $Default
    }

    return $current
}

function ConvertTo-JsonString {
    param(
        [AllowNull()]$InputObject,
        [int]$Depth = 20,
        [switch]$Compress
    )

    if ($Compress) {
        return ($InputObject | ConvertTo-Json -Depth $Depth -Compress)
    }

    return ($InputObject | ConvertTo-Json -Depth $Depth)
}

function Copy-DeepObject {
    param([AllowNull()]$InputObject)

    if ($null -eq $InputObject) {
        return $null
    }

    return ((ConvertTo-JsonString -InputObject $InputObject -Depth 20) | ConvertFrom-Json)
}

function ConvertTo-NumberParseResult {
    param([AllowNull()][object]$Value)

    $text = Get-NormalizedString $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return [pscustomobject]@{ Success = $false; Value = 0.0; IsBlank = $true }
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

function Get-CanonicalIdentifierValue {
    param([string]$Value)

    $normalized = Get-NormalizedString $Value
    $normalized = $normalized.ToLowerInvariant()
    $normalized = $normalized -replace '\s+', ' '
    return $normalized.Trim()
}

function Get-CanonicalKey {
    param(
        [Parameter(Mandatory = $true)][string]$Type,
        [Parameter(Mandatory = $true)][string]$Value
    )

    return '{0}:{1}' -f (Get-NormalizedString $Type), (Get-CanonicalIdentifierValue $Value)
}

function Get-CanonicalUnitCode {
    param([AllowNull()][object]$Value)

    $text = Get-NormalizedString $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }

    return $text.ToUpperInvariant()
}

function Get-RolePropertyName {
    param(
        [Parameter(Mandatory = $true)][string]$RoleName,
        [Parameter(Mandatory = $true)][hashtable]$ExistingNames
    )

    $candidate = (Get-NormalizedString $RoleName) -replace '[^A-Za-z0-9]+', '_'
    $candidate = $candidate.Trim('_')
    if ([string]::IsNullOrWhiteSpace($candidate)) {
        $candidate = 'Role'
    }

    $base = "Role_$candidate"
    $resolved = $base
    $counter = 1
    while ($ExistingNames.ContainsKey($resolved)) {
        $counter++
        $resolved = '{0}_{1}' -f $base, $counter
    }

    $ExistingNames[$resolved] = $true
    return $resolved
}

function ConvertTo-NormalizedAlternateUnit {
    param($AlternateUnit)

    return [pscustomobject][ordered]@{
        unit_code          = ConvertTo-NullableString (Get-DeepPropertyValue $AlternateUnit 'unit_code')
        conversion_to_base = [double](Get-DeepPropertyValue $AlternateUnit 'conversion_to_base' 0.0)
    }
}

function ConvertTo-NormalizedAlternate {
    param($Alternate)

    $positionValue = 0
    [void][int]::TryParse((Get-NormalizedString (Get-DeepPropertyValue $Alternate 'position' 0)), [ref]$positionValue)

    return [pscustomobject][ordered]@{
        position             = $positionValue
        identifier           = [pscustomobject][ordered]@{
            type  = 'matnr'
            value = Get-NormalizedString (Get-DeepPropertyValue $Alternate 'identifier.value')
        }
        material_status_code = Get-NormalizedString (Get-DeepPropertyValue $Alternate 'material_status_code')
        preferred_unit_code  = ConvertTo-NullableString (Get-DeepPropertyValue $Alternate 'preferred_unit_code')
    }
}

function ConvertTo-NormalizedMaterial {
    param(
        [Parameter(Mandatory = $true)]$Material,
        [string]$DefaultUnitCode = 'EA'
    )

    $idValue = 0
    [void][int]::TryParse((Get-NormalizedString (Get-DeepPropertyValue $Material 'id' 0)), [ref]$idValue)

    $alternateUnits = New-Object System.Collections.Generic.List[object]
    foreach ($alternateUnit in @(ConvertTo-ObjectArray (Get-DeepPropertyValue $Material 'quantity.alternate_units' @()))) {
        [void]$alternateUnits.Add((ConvertTo-NormalizedAlternateUnit -AlternateUnit $alternateUnit))
    }

    $alternates = New-Object System.Collections.Generic.List[object]
    foreach ($alternate in @(ConvertTo-ObjectArray (Get-DeepPropertyValue $Material 'alternates' @()))) {
        [void]$alternates.Add((ConvertTo-NormalizedAlternate -Alternate $alternate))
    }

    $resolvedBaseUnit = Get-NormalizedString (Get-DeepPropertyValue $Material 'quantity.base_unit')
    if ([string]::IsNullOrWhiteSpace($resolvedBaseUnit)) {
        $resolvedBaseUnit = $DefaultUnitCode
    }

    $identifierMatnr = ConvertTo-NullableString (Get-DeepPropertyValue $Material 'identifiers.matnr')
    $legacyPrimaryValue = Get-NormalizedString (Get-DeepPropertyValue $Material 'primary_identifier.value')
    if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $identifierMatnr)) -and -not [string]::IsNullOrWhiteSpace($legacyPrimaryValue)) {
        $identifierMatnr = $legacyPrimaryValue
    }

    $primaryValue = Get-NormalizedString $identifierMatnr
    $canonicalKey = if ([string]::IsNullOrWhiteSpace($primaryValue)) { '' } else { Get-CanonicalKey -Type 'matnr' -Value $primaryValue }
    $statusCode = Get-NormalizedString (Get-DeepPropertyValue $Material 'status.material_status_code')
    if ([string]::IsNullOrWhiteSpace($statusCode)) {
        $statusCode = 'XX'
    }

    return [pscustomobject][ordered]@{
        id                 = $idValue
        canonical_key      = $canonicalKey
        primary_identifier = [pscustomobject][ordered]@{
            type  = 'matnr'
            value = $primaryValue
        }
        identifiers        = [pscustomobject][ordered]@{
            matnr             = $identifierMatnr
            supply_number     = ConvertTo-NullableString (Get-DeepPropertyValue $Material 'identifiers.supply_number')
            article_number    = ConvertTo-NullableString (Get-DeepPropertyValue $Material 'identifiers.article_number')
            nato_stock_number = ConvertTo-NullableString (Get-DeepPropertyValue $Material 'identifiers.nato_stock_number')
        }
        status             = [pscustomobject][ordered]@{
            material_status_code = $statusCode
        }
        texts              = [pscustomobject][ordered]@{
            short_description = Get-NormalizedString (Get-DeepPropertyValue $Material 'texts.short_description')
            technical_note    = Get-NormalizedString (Get-DeepPropertyValue $Material 'texts.technical_note')
            logistics_note    = Get-NormalizedString (Get-DeepPropertyValue $Material 'texts.logistics_note')
        }
        classification     = [pscustomobject][ordered]@{
            ext_wg       = Get-NormalizedString (Get-DeepPropertyValue $Material 'classification.ext_wg')
            is_decentral = [bool](Get-DeepPropertyValue $Material 'classification.is_decentral' $false)
            creditor     = ConvertTo-NullableString (Get-DeepPropertyValue $Material 'classification.creditor')
        }
        hazmat             = [pscustomobject][ordered]@{
            is_hazardous = [bool](Get-DeepPropertyValue $Material 'hazmat.is_hazardous' $false)
            un_number    = ConvertTo-NullableString (Get-DeepPropertyValue $Material 'hazmat.un_number')
            flags        = @(ConvertTo-UniqueStringArray (ConvertTo-ObjectArray (Get-DeepPropertyValue $Material 'hazmat.flags' @())))
        }
        quantity           = [pscustomobject][ordered]@{
            base_unit       = $resolvedBaseUnit
            target          = [double](Get-DeepPropertyValue $Material 'quantity.target' 0.0)
            alternate_units = $alternateUnits.ToArray()
        }
        alternates         = $alternates.ToArray()
        assignments        = [pscustomobject][ordered]@{
            responsibility_codes = @(ConvertTo-UniqueStringArray (ConvertTo-ObjectArray (Get-DeepPropertyValue $Material 'assignments.responsibility_codes' @())))
            assignment_tags      = @(ConvertTo-UniqueStringArray (ConvertTo-ObjectArray (Get-DeepPropertyValue $Material 'assignments.assignment_tags' @())))
        }
    }
}

function Read-DatabaseFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$DefaultUnitCode = 'EA'
    )

    if (-not (Test-Path $Path)) {
        throw "Database file not found: $Path"
    }

    $raw = Get-Content -Path $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) {
        throw "Database file is empty: $Path"
    }

    $parsed = $raw | ConvertFrom-Json
    if (-not $parsed.PSObject.Properties['schema_version'] -or -not $parsed.PSObject.Properties['materials']) {
        throw "Unsupported database schema in '$Path'. Expected schema_version $Script:DatabaseSchemaVersion."
    }

    $schemaVersion = 0
    if (-not [int]::TryParse((Get-NormalizedString $parsed.schema_version), [ref]$schemaVersion) -or $schemaVersion -ne $Script:DatabaseSchemaVersion) {
        throw "Unsupported schema_version '$($parsed.schema_version)'. Expected $Script:DatabaseSchemaVersion."
    }

    $materials = New-Object System.Collections.Generic.List[object]
    foreach ($material in @(ConvertTo-ObjectArray $parsed.materials)) {
        [void]$materials.Add((ConvertTo-NormalizedMaterial -Material $material -DefaultUnitCode $DefaultUnitCode))
    }

    return [pscustomobject]@{
        schema_version = $schemaVersion
        lookup_file    = Get-NormalizedString $parsed.lookup_file
        materials      = $materials.ToArray()
    }
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

function Import-DelimitedTextWithHeaderRow {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [int]$HeaderRowIndex = 1,
        [string]$Delimiter = ';'
    )

    if ($HeaderRowIndex -lt 1) {
        throw 'HeaderRowIndex must be greater than or equal to 1.'
    }

    $lines = $Text -split "(`r`n|`n|`r)"
    if ($HeaderRowIndex -gt $lines.Count) {
        throw "HeaderRowIndex $HeaderRowIndex is outside the file."
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
        Headers    = @($headers)
        Rows       = @($rows)
        RowNumbers = @(for ($i = 0; $i -lt @($rows).Count; $i++) { $HeaderRowIndex + 1 + $i })
    }
}

function Read-DelimitedSource {
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

function Read-ExcelWorksheetSource {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$HeaderRowIndex = 1,
        [AllowNull()][string]$WorksheetName
    )

    if ($HeaderRowIndex -lt 1) {
        throw 'HeaderRowIndex must be greater than or equal to 1.'
    }

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $usedRange = $null

    try {
        $excel = New-Object -ComObject Excel.Application
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
                throw "Worksheet '$WorksheetName' was not found in '$Path'."
            }
        }

        $usedRange = $worksheet.UsedRange
        $firstUsedRow = [int]$usedRange.Row
        $firstUsedColumn = [int]$usedRange.Column
        $lastUsedRow = $firstUsedRow + [int]$usedRange.Rows.Count - 1
        $lastUsedColumn = $firstUsedColumn + [int]$usedRange.Columns.Count - 1

        if ($HeaderRowIndex -gt $lastUsedRow) {
            throw "HeaderRowIndex $HeaderRowIndex is outside worksheet '$($worksheet.Name)'."
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

function New-DefaultPresetStore {
    return [pscustomobject][ordered]@{
        schema_version = $Script:PresetSchemaVersion
        presets        = @()
    }
}

function ConvertTo-NormalizedComparePreset {
    param(
        [Parameter(Mandatory = $true)]$Preset,
        [string]$DefaultFileType = ''
    )

    $name = Get-NormalizedString (Get-DeepPropertyValue $Preset 'Name')
    $fileType = Get-NormalizedString (Get-DeepPropertyValue $Preset 'FileType')
    if ([string]::IsNullOrWhiteSpace($fileType)) {
        $fileType = Get-NormalizedString $DefaultFileType
    }

    if (-not [string]::IsNullOrWhiteSpace($fileType)) {
        $fileType = $fileType.ToLowerInvariant()
    }

    $headerRowIndex = 0
    [void][int]::TryParse((Get-NormalizedString (Get-DeepPropertyValue $Preset 'HeaderRowIndex' 0)), [ref]$headerRowIndex)
    if ($headerRowIndex -le 0) {
        $headerRowIndex = 1
    }

    $columnSource = Get-DeepPropertyValue $Preset 'Columns' $null
    $columnKeys = @('material_number', 'quantity', 'unit', 'description', 'status', 'storage_location', 'batch', 'note')
    $columns = [ordered]@{}
    foreach ($columnKey in $columnKeys) {
        $columns[$columnKey] = ConvertTo-NullableString (Get-DeepPropertyValue $columnSource $columnKey)
    }

    if ([string]::IsNullOrWhiteSpace($fileType) -or @('xlsx', 'csv', 'txt') -notcontains $fileType) {
        throw "Preset '$name' has an invalid FileType. Expected xlsx, csv, or txt."
    }

    foreach ($requiredKey in @('material_number', 'quantity', 'unit')) {
        if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $columns[$requiredKey]))) {
            throw "Preset '$name' is missing required column mapping '$requiredKey'."
        }
    }

    return [pscustomobject][ordered]@{
        Name           = $(if ([string]::IsNullOrWhiteSpace($name)) { $null } else { $name })
        FileType       = $fileType
        HeaderRowIndex = $headerRowIndex
        WorksheetName  = ConvertTo-NullableString (Get-DeepPropertyValue $Preset 'WorksheetName')
        Delimiter      = ConvertTo-NullableString (Get-DeepPropertyValue $Preset 'Delimiter')
        Columns        = [pscustomobject]$columns
    }
}

function Read-ComparePresetStore {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path $Path)) {
        return (New-DefaultPresetStore)
    }

    $raw = Get-Content -Path $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return (New-DefaultPresetStore)
    }

    $parsed = $raw | ConvertFrom-Json
    if (-not $parsed.PSObject.Properties['schema_version'] -or -not $parsed.PSObject.Properties['presets']) {
        throw "Unsupported preset store schema in '$Path'."
    }

    $schemaVersion = 0
    if (-not [int]::TryParse((Get-NormalizedString $parsed.schema_version), [ref]$schemaVersion) -or $schemaVersion -ne $Script:PresetSchemaVersion) {
        throw "Unsupported preset store schema_version '$($parsed.schema_version)'. Expected $Script:PresetSchemaVersion."
    }

    $normalizedPresets = New-Object System.Collections.Generic.List[object]
    $seenNames = @{}
    foreach ($preset in @(ConvertTo-ObjectArray $parsed.presets)) {
        $normalized = ConvertTo-NormalizedComparePreset -Preset $preset
        $presetName = Get-NormalizedString $normalized.Name
        if ([string]::IsNullOrWhiteSpace($presetName)) {
            throw "Preset store '$Path' contains a preset without a Name."
        }

        if ($seenNames.ContainsKey($presetName)) {
            throw "Preset store '$Path' contains duplicate preset '$presetName'."
        }

        $seenNames[$presetName] = $true
        [void]$normalizedPresets.Add($normalized)
    }

    return [pscustomobject][ordered]@{
        schema_version = $schemaVersion
        presets        = @($normalizedPresets.ToArray())
    }
}

function Write-ComparePresetStore {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$Store
    )

    $directory = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path $directory)) {
        New-Item -Path $directory -ItemType Directory -Force | Out-Null
    }

    $normalizedStore = [pscustomobject][ordered]@{
        schema_version = $Script:PresetSchemaVersion
        presets        = @(
            @(ConvertTo-ObjectArray (Get-DeepPropertyValue $Store 'presets' @())) | ForEach-Object {
                ConvertTo-NormalizedComparePreset -Preset $_
            }
        )
    }

    $normalizedStore | ConvertTo-Json -Depth 10 | Out-File -FilePath $Path -Encoding UTF8
}

function Get-StockComparisonPresetStore {
    param([string]$Path = $Script:DefaultPresetStorePath)

    return (Read-ComparePresetStore -Path $Path)
}

function Get-StockComparisonPreset {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [string]$Path = $Script:DefaultPresetStorePath
    )

    $store = Read-ComparePresetStore -Path $Path
    foreach ($preset in @(ConvertTo-ObjectArray $store.presets)) {
        if ((Get-NormalizedString $preset.Name) -eq (Get-NormalizedString $Name)) {
            return $preset
        }
    }

    throw "Preset '$Name' was not found in '$Path'."
}

function Set-StockComparisonPreset {
    param(
        [Parameter(Mandatory = $true)]$Preset,
        [string]$Path = $Script:DefaultPresetStorePath
    )

    $normalizedPreset = ConvertTo-NormalizedComparePreset -Preset $Preset
    $presetName = Get-NormalizedString $normalizedPreset.Name
    if ([string]::IsNullOrWhiteSpace($presetName)) {
        throw 'Preset Name is required.'
    }

    $store = Read-ComparePresetStore -Path $Path
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

    Write-ComparePresetStore -Path $Path -Store ([pscustomobject][ordered]@{
            schema_version = $Script:PresetSchemaVersion
            presets        = $presets.ToArray()
        })

    return $normalizedPreset
}

function Resolve-ComparisonPreset {
    param(
        [Parameter(Mandatory = $true)]$SourceSpec,
        [Parameter(Mandatory = $true)][string]$PresetStorePath
    )

    $inlinePreset = Get-DeepPropertyValue $SourceSpec 'Preset' $null
    if ($null -ne $inlinePreset) {
        $sourcePath = Get-NormalizedString (Get-DeepPropertyValue $SourceSpec 'Path')
        $defaultFileType = [System.IO.Path]::GetExtension($sourcePath).TrimStart('.').ToLowerInvariant()
        $preset = ConvertTo-NormalizedComparePreset -Preset $inlinePreset -DefaultFileType $defaultFileType
    }
    else {
        $presetName = Get-NormalizedString (Get-DeepPropertyValue $SourceSpec 'PresetName')
        if ([string]::IsNullOrWhiteSpace($presetName)) {
            throw 'Each SourceSpec must contain PresetName or inline Preset.'
        }

        $preset = Copy-DeepObject (Get-StockComparisonPreset -Name $presetName -Path $PresetStorePath)
    }

    $worksheetOverride = ConvertTo-NullableString (Get-DeepPropertyValue $SourceSpec 'WorksheetName')
    $delimiterOverride = ConvertTo-NullableString (Get-DeepPropertyValue $SourceSpec 'Delimiter')
    if (-not [string]::IsNullOrWhiteSpace($worksheetOverride)) {
        $preset.WorksheetName = $worksheetOverride
    }

    if (-not [string]::IsNullOrWhiteSpace($delimiterOverride)) {
        $preset.Delimiter = $delimiterOverride
    }

    return $preset
}

function Read-SourceRows {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$Preset
    )

    $fileType = Get-NormalizedString $Preset.FileType
    switch ($fileType) {
        'xlsx' {
            return (Read-ExcelWorksheetSource -Path $Path -HeaderRowIndex ([int]$Preset.HeaderRowIndex) -WorksheetName $Preset.WorksheetName)
        }
        'csv' {
            $delimiter = Get-NormalizedString $Preset.Delimiter
            if ([string]::IsNullOrWhiteSpace($delimiter)) {
                $delimiter = ';'
            }

            return (Read-DelimitedSource -Path $Path -HeaderRowIndex ([int]$Preset.HeaderRowIndex) -Delimiter $delimiter)
        }
        'txt' {
            $delimiter = Get-NormalizedString $Preset.Delimiter
            if ([string]::IsNullOrWhiteSpace($delimiter)) {
                $delimiter = ';'
            }

            return (Read-DelimitedSource -Path $Path -HeaderRowIndex ([int]$Preset.HeaderRowIndex) -Delimiter $delimiter)
        }
        default {
            throw "Unsupported file type '$fileType' for '$Path'."
        }
    }
}

function Test-HasColumnHeader {
    param(
        [AllowNull()][object[]]$Headers,
        [Parameter(Mandatory = $true)][string]$HeaderName
    )

    foreach ($header in @(ConvertTo-ObjectArray $Headers)) {
        if ((Get-NormalizedString $header) -eq (Get-NormalizedString $HeaderName)) {
            return $true
        }
    }

    return $false
}

function Get-ColumnValue {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [AllowNull()][string]$HeaderName
    )

    if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $HeaderName))) {
        return ''
    }

    $property = $Row.PSObject.Properties[$HeaderName]
    if ($null -eq $property) {
        return ''
    }

    return Get-NormalizedString $property.Value
}

function New-AliasIndex {
    param([Parameter(Mandatory = $true)][object[]]$Materials)

    $aliasLookup = @{}
    $duplicateMap = @{}

    foreach ($material in @(ConvertTo-ObjectArray $Materials)) {
        $anchorMatnr = Get-NormalizedString $material.identifiers.matnr
        if ([string]::IsNullOrWhiteSpace($anchorMatnr)) {
            continue
        }

        $anchorEntries = New-Object System.Collections.Generic.List[object]
        [void]$anchorEntries.Add([pscustomobject][ordered]@{
                AliasValue        = $anchorMatnr
                IsAnchor          = $true
                PreferredUnitCode = $null
                AnchorMaterial    = $material
            })

        foreach ($alternate in @(ConvertTo-ObjectArray $material.alternates)) {
            $alternateValue = Get-NormalizedString $alternate.identifier.value
            if ([string]::IsNullOrWhiteSpace($alternateValue)) {
                continue
            }

            [void]$anchorEntries.Add([pscustomobject][ordered]@{
                    AliasValue        = $alternateValue
                    IsAnchor          = $false
                    PreferredUnitCode = ConvertTo-NullableString $alternate.preferred_unit_code
                    AnchorMaterial    = $material
                })
        }

        foreach ($entry in @($anchorEntries.ToArray())) {
            $canonicalKey = Get-CanonicalKey -Type 'matnr' -Value $entry.AliasValue
            if (-not $aliasLookup.ContainsKey($canonicalKey)) {
                $aliasLookup[$canonicalKey] = $entry
            }
            else {
                if (-not $duplicateMap.ContainsKey($canonicalKey)) {
                    $bucket = New-Object System.Collections.Generic.List[object]
                    [void]$bucket.Add($aliasLookup[$canonicalKey])
                    $duplicateMap[$canonicalKey] = $bucket
                }

                [void]$duplicateMap[$canonicalKey].Add($entry)
            }
        }
    }

    foreach ($duplicateKey in @($duplicateMap.Keys)) {
        [void]$aliasLookup.Remove($duplicateKey)
    }

    $diagnostics = New-Object System.Collections.Generic.List[object]
    foreach ($duplicateKey in @($duplicateMap.Keys | Sort-Object)) {
        $entries = @($duplicateMap[$duplicateKey].ToArray())
        $materialIds = New-Object System.Collections.Generic.List[int]
        $materialNumbers = New-Object System.Collections.Generic.List[string]
        foreach ($entry in $entries) {
            [void]$materialIds.Add([int]$entry.AnchorMaterial.id)
            [void]$materialNumbers.Add((Get-NormalizedString $entry.AnchorMaterial.identifiers.matnr))
        }

        [void]$diagnostics.Add([pscustomobject][ordered]@{
                AliasKey        = $duplicateKey
                AliasValues     = @(ConvertTo-UniqueStringArray ($entries | ForEach-Object { $_.AliasValue }))
                MaterialIds     = @($materialIds.ToArray() | Sort-Object -Unique)
                MaterialNumbers = @($materialNumbers.ToArray() | Sort-Object -Unique)
                EntryCount      = $entries.Count
            })
    }

    return [pscustomobject]@{
        Lookup                    = $aliasLookup
        DuplicateKeys             = @($duplicateMap.Keys)
        DuplicateAliasDiagnostics = @($diagnostics.ToArray())
    }
}

function Get-QuantityConversionInfo {
    param(
        [Parameter(Mandatory = $true)]$AnchorMaterial,
        [Parameter(Mandatory = $true)]$AliasEntry,
        [AllowNull()][string]$UnitCode
    )

    $baseUnit = Get-CanonicalUnitCode $AnchorMaterial.quantity.base_unit
    $resolvedUnit = Get-CanonicalUnitCode $UnitCode

    if ([string]::IsNullOrWhiteSpace($resolvedUnit)) {
        if (-not [bool]$AliasEntry.IsAnchor) {
            $preferredUnit = Get-CanonicalUnitCode $AliasEntry.PreferredUnitCode
            if (-not [string]::IsNullOrWhiteSpace($preferredUnit)) {
                $resolvedUnit = $preferredUnit
            }
        }

        if ([string]::IsNullOrWhiteSpace($resolvedUnit)) {
            $resolvedUnit = $baseUnit
        }
    }

    if ($resolvedUnit -eq $baseUnit) {
        return [pscustomobject]@{
            Success          = $true
            ResolvedUnitCode = $resolvedUnit
            ConversionFactor = 1.0
            BaseUnitCode     = $baseUnit
            Reason           = 'base_unit'
        }
    }

    foreach ($alternateUnit in @(ConvertTo-ObjectArray $AnchorMaterial.quantity.alternate_units)) {
        $alternateUnitCode = Get-CanonicalUnitCode $alternateUnit.unit_code
        if ($alternateUnitCode -eq $resolvedUnit) {
            return [pscustomobject]@{
                Success          = $true
                ResolvedUnitCode = $resolvedUnit
                ConversionFactor = [double]$alternateUnit.conversion_to_base
                BaseUnitCode     = $baseUnit
                Reason           = 'alternate_unit'
            }
        }
    }

    return [pscustomobject]@{
        Success          = $false
        ResolvedUnitCode = $resolvedUnit
        ConversionFactor = 0.0
        BaseUnitCode     = $baseUnit
        Reason           = 'unknown_unit'
    }
}

function New-ComparisonAccumulator {
    param(
        [Parameter(Mandatory = $true)]$Material,
        [Parameter(Mandatory = $true)][string[]]$RoleNames
    )

    $roleTotals = [ordered]@{}
    foreach ($roleName in @($RoleNames)) {
        $roleTotals[$roleName] = 0.0
    }

    return [pscustomobject][ordered]@{
        Material          = $Material
        RoleTotals        = [pscustomobject]$roleTotals
        MatchedAliases    = New-Object System.Collections.Generic.List[string]
        MatchedSourceRows = 0
        MatchedRowDetails = New-Object System.Collections.Generic.List[object]
    }
}

function Add-MatchedAlias {
    param(
        [Parameter(Mandatory = $true)]$Accumulator,
        [Parameter(Mandatory = $true)][string]$AliasValue
    )

    $normalizedAlias = Get-NormalizedString $AliasValue
    if ([string]::IsNullOrWhiteSpace($normalizedAlias)) {
        return
    }

    foreach ($existingAlias in @($Accumulator.MatchedAliases.ToArray())) {
        if ((Get-NormalizedString $existingAlias) -eq $normalizedAlias) {
            return
        }
    }

    [void]$Accumulator.MatchedAliases.Add($normalizedAlias)
}

function ConvertTo-GridRow {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][hashtable]$RolePropertyMap
    )

    $ordered = [ordered]@{
        Id                      = [int]$Row.Id
        MaterialNumber          = Get-NormalizedString $Row.MaterialNumber
        Description             = Get-NormalizedString $Row.Description
        SupplyNumber            = Get-NormalizedString $Row.SupplyNumber
        NatoStockNumber         = Get-NormalizedString $Row.NatoStockNumber
        BaseUnit                = Get-NormalizedString $Row.BaseUnit
        TargetQuantityBase      = [double]$Row.TargetQuantityBase
        StockQuantityBase       = [double]$Row.StockQuantityBase
        InboundQuantityBase     = [double]$Row.InboundQuantityBase
        AvailableQuantityBase   = [double]$Row.AvailableQuantityBase
        MissingToOrderBase      = [double]$Row.MissingToOrderBase
        SurplusAfterInboundBase = [double]$Row.SurplusAfterInboundBase
        StockGapBase            = [double]$Row.StockGapBase
        MatchState              = Get-NormalizedString $Row.MatchState
        MatchedAliasList        = @($Row.MatchedAliases) -join ', '
        MatchedSourceRowCount   = [int]$Row.MatchedSourceRowCount
    }

    foreach ($roleName in @($RolePropertyMap.Keys | Sort-Object)) {
        $ordered[$RolePropertyMap[$roleName]] = [double](Get-DeepPropertyValue $Row.RoleQuantities $roleName 0.0)
    }

    return [pscustomobject]$ordered
}

function Invoke-StockComparison {
    [CmdletBinding()]
    param(
        [string]$DatabasePath = $Script:DefaultDatabasePath,
        [Parameter(Mandatory = $true)][object[]]$SourceSpecs,
        [Parameter(Mandatory = $true)][string]$StockRoleName,
        [string]$PresetStorePath = $Script:DefaultPresetStorePath
    )

    if (@(ConvertTo-ObjectArray $SourceSpecs).Count -eq 0) {
        throw 'SourceSpecs must contain at least one source.'
    }

    $database = Read-DatabaseFile -Path $DatabasePath
    $materials = @($database.materials)

    $normalizedSourceSpecs = New-Object System.Collections.Generic.List[object]
    $seenRoleNames = @{}
    foreach ($sourceSpec in @(ConvertTo-ObjectArray $SourceSpecs)) {
        $roleName = Get-NormalizedString (Get-DeepPropertyValue $sourceSpec 'RoleName')
        if ([string]::IsNullOrWhiteSpace($roleName)) {
            throw 'Each SourceSpec must contain RoleName.'
        }

        if ($seenRoleNames.ContainsKey($roleName)) {
            throw "RoleName '$roleName' is duplicated in SourceSpecs."
        }

        $seenRoleNames[$roleName] = $true
        $path = Get-NormalizedString (Get-DeepPropertyValue $sourceSpec 'Path')
        if ([string]::IsNullOrWhiteSpace($path)) {
            throw "SourceSpec '$roleName' is missing Path."
        }

        if (-not (Test-Path $path)) {
            throw "Source path not found for role '$roleName': $path"
        }

        $preset = Resolve-ComparisonPreset -SourceSpec $sourceSpec -PresetStorePath $PresetStorePath
        [void]$normalizedSourceSpecs.Add([pscustomobject][ordered]@{
                RoleName = $roleName
                Path     = $path
                Preset   = $preset
            })
    }

    $roleNames = @($normalizedSourceSpecs.ToArray() | ForEach-Object { $_.RoleName })
    if (-not ($roleNames -contains $StockRoleName)) {
        throw "StockRoleName '$StockRoleName' was not found in SourceSpecs."
    }

    $aliasIndex = New-AliasIndex -Materials $materials

    $rolePropertyMap = @{}
    $usedRolePropertyNames = @{}
    foreach ($roleName in @($roleNames)) {
        $rolePropertyMap[$roleName] = Get-RolePropertyName -RoleName $roleName -ExistingNames $usedRolePropertyNames
    }

    $accumulatorById = @{}
    foreach ($material in @($materials)) {
        $accumulatorById[[int]$material.id] = New-ComparisonAccumulator -Material $material -RoleNames $roleNames
    }

    $unknownSapMaterials = New-Object System.Collections.Generic.List[object]
    $invalidUnits = New-Object System.Collections.Generic.List[object]
    $invalidRows = New-Object System.Collections.Generic.List[object]
    $sourceSummaries = New-Object System.Collections.Generic.List[object]

    foreach ($source in @($normalizedSourceSpecs.ToArray())) {
        $roleName = $source.RoleName
        $preset = $source.Preset
        $sourceData = Read-SourceRows -Path $source.Path -Preset $preset
        $sourceHeaders = @($sourceData.Headers)
        $sourceRows = @($sourceData.Rows)
        $sourceRowNumbers = @($sourceData.RowNumbers)

        foreach ($columnKey in @('material_number', 'quantity', 'unit', 'description', 'status', 'storage_location', 'batch', 'note')) {
            $headerName = Get-NormalizedString (Get-DeepPropertyValue $preset.Columns $columnKey)
            if (-not [string]::IsNullOrWhiteSpace($headerName) -and -not (Test-HasColumnHeader -Headers $sourceHeaders -HeaderName $headerName)) {
                throw "Source '$roleName' is missing mapped header '$headerName' for column '$columnKey'."
            }
        }

        $matchedRowCount = 0
        $unknownRowCount = 0
        $invalidRowCount = 0
        $invalidUnitCount = 0

        for ($rowIndex = 0; $rowIndex -lt $sourceRows.Count; $rowIndex++) {
            $row = $sourceRows[$rowIndex]
            $sourceRowNumber = if ($rowIndex -lt $sourceRowNumbers.Count) { [int]$sourceRowNumbers[$rowIndex] } else { $rowIndex + 1 }
            $materialNumber = Get-ColumnValue -Row $row -HeaderName $preset.Columns.material_number
            $quantityRaw = Get-ColumnValue -Row $row -HeaderName $preset.Columns.quantity
            $unitRaw = Get-ColumnValue -Row $row -HeaderName $preset.Columns.unit
            $description = Get-ColumnValue -Row $row -HeaderName $preset.Columns.description
            $statusText = Get-ColumnValue -Row $row -HeaderName $preset.Columns.status
            $storageLocation = Get-ColumnValue -Row $row -HeaderName $preset.Columns.storage_location
            $batch = Get-ColumnValue -Row $row -HeaderName $preset.Columns.batch
            $note = Get-ColumnValue -Row $row -HeaderName $preset.Columns.note

            if ([string]::IsNullOrWhiteSpace($materialNumber)) {
                $invalidRowCount++
                [void]$invalidRows.Add([pscustomobject][ordered]@{
                        SourceRole      = $roleName
                        SourcePath      = $source.Path
                        RowNumber       = $sourceRowNumber
                        MaterialNumber  = $materialNumber
                        QuantityRaw     = $quantityRaw
                        UnitRaw         = $unitRaw
                        Description     = $description
                        Status          = $statusText
                        StorageLocation = $storageLocation
                        Batch           = $batch
                        Note            = $note
                        Reason          = 'material_number_missing'
                    })
                continue
            }

            $quantityParse = ConvertTo-NumberParseResult $quantityRaw
            if (-not $quantityParse.Success) {
                $invalidRowCount++
                [void]$invalidRows.Add([pscustomobject][ordered]@{
                        SourceRole      = $roleName
                        SourcePath      = $source.Path
                        RowNumber       = $sourceRowNumber
                        MaterialNumber  = $materialNumber
                        QuantityRaw     = $quantityRaw
                        UnitRaw         = $unitRaw
                        Description     = $description
                        Status          = $statusText
                        StorageLocation = $storageLocation
                        Batch           = $batch
                        Note            = $note
                        Reason          = 'quantity_invalid'
                    })
                continue
            }

            $canonicalAlias = Get-CanonicalKey -Type 'matnr' -Value $materialNumber
            if ($aliasIndex.DuplicateKeys -contains $canonicalAlias) {
                $invalidRowCount++
                [void]$invalidRows.Add([pscustomobject][ordered]@{
                        SourceRole      = $roleName
                        SourcePath      = $source.Path
                        RowNumber       = $sourceRowNumber
                        MaterialNumber  = $materialNumber
                        QuantityRaw     = $quantityRaw
                        UnitRaw         = $unitRaw
                        Description     = $description
                        Status          = $statusText
                        StorageLocation = $storageLocation
                        Batch           = $batch
                        Note            = $note
                        Reason          = 'duplicate_alias_blocked'
                    })
                continue
            }

            if (-not $aliasIndex.Lookup.ContainsKey($canonicalAlias)) {
                $unknownRowCount++
                [void]$unknownSapMaterials.Add([pscustomobject][ordered]@{
                        SourceRole      = $roleName
                        SourcePath      = $source.Path
                        RowNumber       = $sourceRowNumber
                        MaterialNumber  = $materialNumber
                        QuantityRaw     = $quantityRaw
                        UnitRaw         = $unitRaw
                        Description     = $description
                        Status          = $statusText
                        StorageLocation = $storageLocation
                        Batch           = $batch
                        Note            = $note
                    })
                continue
            }

            $aliasEntry = $aliasIndex.Lookup[$canonicalAlias]
            $anchorMaterial = $aliasEntry.AnchorMaterial
            $conversionInfo = Get-QuantityConversionInfo -AnchorMaterial $anchorMaterial -AliasEntry $aliasEntry -UnitCode $unitRaw
            if (-not $conversionInfo.Success) {
                $invalidUnitCount++
                [void]$invalidUnits.Add([pscustomobject][ordered]@{
                        SourceRole       = $roleName
                        SourcePath       = $source.Path
                        RowNumber        = $sourceRowNumber
                        MaterialNumber   = $materialNumber
                        QuantityRaw      = $quantityRaw
                        UnitRaw          = $unitRaw
                        ResolvedUnitCode = $conversionInfo.ResolvedUnitCode
                        BaseUnitCode     = $conversionInfo.BaseUnitCode
                        AllowedUnitCodes = @(
                            @(
                                $anchorMaterial.quantity.base_unit
                                @(ConvertTo-ObjectArray $anchorMaterial.quantity.alternate_units | ForEach-Object { $_.unit_code })
                            ) | Where-Object { -not [string]::IsNullOrWhiteSpace((Get-NormalizedString $_)) } | Select-Object -Unique
                        )
                        Description      = $description
                        Status           = $statusText
                        StorageLocation  = $storageLocation
                        Batch            = $batch
                        Note             = $note
                    })
                continue
            }

            $quantityBase = [double]$quantityParse.Value * [double]$conversionInfo.ConversionFactor
            $accumulator = $accumulatorById[[int]$anchorMaterial.id]
            $accumulator.RoleTotals.PSObject.Properties[$roleName].Value = [double]$accumulator.RoleTotals.PSObject.Properties[$roleName].Value + $quantityBase
            $accumulator.MatchedSourceRows++
            Add-MatchedAlias -Accumulator $accumulator -AliasValue $materialNumber
            [void]$accumulator.MatchedRowDetails.Add([pscustomobject][ordered]@{
                    SourceRole       = $roleName
                    SourcePath       = $source.Path
                    RowNumber        = $sourceRowNumber
                    MatchedAlias     = $materialNumber
                    IsAnchorAlias    = [bool]$aliasEntry.IsAnchor
                    QuantityRaw      = Get-NormalizedString $quantityRaw
                    QuantityBase     = $quantityBase
                    UnitRaw          = $unitRaw
                    ResolvedUnitCode = $conversionInfo.ResolvedUnitCode
                    ConversionFactor = [double]$conversionInfo.ConversionFactor
                    Description      = $description
                    Status           = $statusText
                    StorageLocation  = $storageLocation
                    Batch            = $batch
                    Note             = $note
                })
            $matchedRowCount++
        }

        [void]$sourceSummaries.Add([pscustomobject][ordered]@{
                RoleName         = $roleName
                IsStockRole      = ($roleName -eq $StockRoleName)
                Path             = $source.Path
                FileType         = Get-NormalizedString $preset.FileType
                PresetName       = Get-NormalizedString $preset.Name
                HeaderRowIndex   = [int]$preset.HeaderRowIndex
                WorksheetName    = Get-NormalizedString $preset.WorksheetName
                Delimiter        = Get-NormalizedString $preset.Delimiter
                EncodingName     = Get-NormalizedString $sourceData.EncodingName
                RowCount         = $sourceRows.Count
                MatchedRowCount  = $matchedRowCount
                UnknownRowCount  = $unknownRowCount
                InvalidRowCount  = $invalidRowCount
                InvalidUnitCount = $invalidUnitCount
            })
    }

    $rows = New-Object System.Collections.Generic.List[object]
    $gridRows = New-Object System.Collections.Generic.List[object]
    foreach ($material in @($materials | Sort-Object { [int]$_.id })) {
        $accumulator = $accumulatorById[[int]$material.id]
        $stockQuantityBase = [double](Get-DeepPropertyValue $accumulator.RoleTotals $StockRoleName 0.0)
        $inboundQuantityBase = 0.0
        foreach ($roleName in @($roleNames)) {
            if ($roleName -ne $StockRoleName) {
                $inboundQuantityBase += [double](Get-DeepPropertyValue $accumulator.RoleTotals $roleName 0.0)
            }
        }

        $targetQuantityBase = [double]$material.quantity.target
        $availableQuantityBase = $stockQuantityBase + $inboundQuantityBase
        $missingToOrderBase = [Math]::Max($targetQuantityBase - $availableQuantityBase, 0.0)
        $surplusAfterInboundBase = [Math]::Max($availableQuantityBase - $targetQuantityBase, 0.0)
        $stockGapBase = [Math]::Max($targetQuantityBase - $stockQuantityBase, 0.0)

        $matchState = if ($missingToOrderBase -gt 0) {
            'missing'
        }
        elseif ($surplusAfterInboundBase -gt 0) {
            'surplus'
        }
        else {
            'balanced'
        }

        $roleTotalsOrdered = [ordered]@{}
        foreach ($roleName in @($roleNames)) {
            $roleTotalsOrdered[$roleName] = [double](Get-DeepPropertyValue $accumulator.RoleTotals $roleName 0.0)
        }

        $row = [pscustomobject][ordered]@{
            Id                      = [int]$material.id
            MaterialNumber          = Get-NormalizedString $material.identifiers.matnr
            Description             = Get-NormalizedString $material.texts.short_description
            SupplyNumber            = Get-NormalizedString $material.identifiers.supply_number
            NatoStockNumber         = Get-NormalizedString $material.identifiers.nato_stock_number
            BaseUnit                = Get-NormalizedString $material.quantity.base_unit
            TargetQuantityBase      = $targetQuantityBase
            StockQuantityBase       = $stockQuantityBase
            InboundQuantityBase     = $inboundQuantityBase
            AvailableQuantityBase   = $availableQuantityBase
            MissingToOrderBase      = $missingToOrderBase
            SurplusAfterInboundBase = $surplusAfterInboundBase
            StockGapBase            = $stockGapBase
            MatchState              = $matchState
            RoleQuantities          = [pscustomobject]$roleTotalsOrdered
            MatchedAliases          = @(ConvertTo-UniqueStringArray $accumulator.MatchedAliases.ToArray())
            MatchedSourceRowCount   = [int]$accumulator.MatchedSourceRows
            MatchedRows             = @($accumulator.MatchedRowDetails.ToArray())
            MaterialRef             = $material
        }

        [void]$rows.Add($row)
        [void]$gridRows.Add((ConvertTo-GridRow -Row $row -RolePropertyMap $rolePropertyMap))
    }

    $summaryByRole = New-Object System.Collections.Generic.List[object]
    foreach ($roleName in @($roleNames)) {
        $roleTotal = 0.0
        foreach ($row in @($rows.ToArray())) {
            $roleTotal += [double](Get-DeepPropertyValue $row.RoleQuantities $roleName 0.0)
        }

        [void]$summaryByRole.Add([pscustomobject][ordered]@{
                RoleName          = $roleName
                IsStockRole       = ($roleName -eq $StockRoleName)
                QuantityBaseTotal = $roleTotal
            })
    }

    $summary = [pscustomobject][ordered]@{
        MaterialCount                = @($rows.ToArray()).Count
        MatchedMaterialCount         = @(@($rows.ToArray()) | Where-Object { $_.MatchedSourceRowCount -gt 0 }).Count
        MissingMaterialCount         = @(@($rows.ToArray()) | Where-Object { $_.MissingToOrderBase -gt 0 }).Count
        SurplusMaterialCount         = @(@($rows.ToArray()) | Where-Object { $_.SurplusAfterInboundBase -gt 0 }).Count
        BalancedMaterialCount        = @(@($rows.ToArray()) | Where-Object { $_.MissingToOrderBase -eq 0 -and $_.SurplusAfterInboundBase -eq 0 }).Count
        TargetQuantityBaseTotal      = (@($rows.ToArray()) | Measure-Object -Property TargetQuantityBase -Sum).Sum
        StockQuantityBaseTotal       = (@($rows.ToArray()) | Measure-Object -Property StockQuantityBase -Sum).Sum
        InboundQuantityBaseTotal     = (@($rows.ToArray()) | Measure-Object -Property InboundQuantityBase -Sum).Sum
        AvailableQuantityBaseTotal   = (@($rows.ToArray()) | Measure-Object -Property AvailableQuantityBase -Sum).Sum
        MissingToOrderBaseTotal      = (@($rows.ToArray()) | Measure-Object -Property MissingToOrderBase -Sum).Sum
        SurplusAfterInboundBaseTotal = (@($rows.ToArray()) | Measure-Object -Property SurplusAfterInboundBase -Sum).Sum
        StockGapBaseTotal            = (@($rows.ToArray()) | Measure-Object -Property StockGapBase -Sum).Sum
        RoleTotals                   = @($summaryByRole.ToArray())
        StockRoleName                = $StockRoleName
        RolePropertyMap              = [pscustomobject]$rolePropertyMap
    }

    return [pscustomobject][ordered]@{
        Rows        = @($rows.ToArray())
        GridRows    = @($gridRows.ToArray())
        Summary     = $summary
        Diagnostics = [pscustomobject][ordered]@{
            unknown_sap_materials = @($unknownSapMaterials.ToArray())
            invalid_units         = @($invalidUnits.ToArray())
            invalid_rows          = @($invalidRows.ToArray())
            duplicate_aliases     = @($aliasIndex.DuplicateAliasDiagnostics)
        }
        Sources     = @($sourceSummaries.ToArray())
    }
}
