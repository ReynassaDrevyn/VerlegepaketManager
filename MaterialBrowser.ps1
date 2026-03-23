# MaterialBrowser.ps1
# PowerShell 5.1 WPF editor for the material database.

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$Script:ProjectRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path $MyInvocation.MyCommand.Path -Parent }
$Script:DbPath = Join-Path $Script:ProjectRoot 'Core\db_verlegepaket.json'
$Script:LookupPath = Join-Path $Script:ProjectRoot 'Core\db_lookups.json'
$Script:LogsDir = Join-Path $Script:ProjectRoot 'Logs'
$Script:BackupDir = Join-Path $Script:LogsDir 'Backups'
$Script:DatabaseSchemaVersion = 2
$Script:LookupFileName = 'db_lookups.json'

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

function ConvertTo-LabelTextArray {
    param(
        [AllowNull()][object[]]$Values,
        [AllowNull()][hashtable]$LabelMap
    )

    $result = New-Object System.Collections.Generic.List[string]
    foreach ($value in @(ConvertTo-ObjectArray $Values)) {
        $text = Get-NormalizedString $value
        if ([string]::IsNullOrWhiteSpace($text)) {
            continue
        }

        $label = $text
        if ($null -ne $LabelMap -and $LabelMap.ContainsKey($text)) {
            $label = Get-NormalizedString $LabelMap[$text]
        }

        [void]$result.Add($label)
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

function ConvertTo-IntParseResult {
    param([AllowNull()][object]$Value)

    $text = Get-NormalizedString $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return [pscustomobject]@{ Success = $false; Value = 0; IsBlank = $true }
    }

    $parsed = 0
    if ([int]::TryParse($text, [ref]$parsed)) {
        return [pscustomobject]@{ Success = $true; Value = $parsed; IsBlank = $false }
    }

    return [pscustomobject]@{ Success = $false; Value = 0; IsBlank = $false }
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

function New-ObservableCollection {
    param([AllowNull()][object[]]$Items)

    $collection = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
    foreach ($item in @(ConvertTo-ObjectArray $Items)) {
        [void]$collection.Add($item)
    }

    return ,$collection
}

function New-DefaultMaterial {
    param(
        [int]$Id = 0,
        [string]$DefaultIdentifierType = 'matnr',
        [string]$DefaultUnitCode = 'EA'
    )

    return [pscustomobject][ordered]@{
        id                 = $Id
        canonical_key      = ''
        primary_identifier = [pscustomobject][ordered]@{
            type  = $DefaultIdentifierType
            value = ''
        }
        identifiers        = [pscustomobject][ordered]@{
            matnr             = $null
            supply_number     = $null
            article_number    = $null
            nato_stock_number = $null
        }
        status             = [pscustomobject][ordered]@{
            material_status_code = 'XX'
        }
        texts              = [pscustomobject][ordered]@{
            short_description = ''
            technical_note    = ''
            logistics_note    = ''
        }
        classification     = [pscustomobject][ordered]@{
            ext_wg       = ''
            is_decentral = $false
            creditor     = $null
        }
        hazmat             = [pscustomobject][ordered]@{
            is_hazardous = $false
            un_number    = $null
            flags        = @()
        }
        quantity           = [pscustomobject][ordered]@{
            base_unit       = $DefaultUnitCode
            target          = 0.0
            alternate_units = @()
        }
        alternates         = @()
        assignments        = [pscustomobject][ordered]@{
            responsibility_codes = @()
            assignment_tags      = @()
        }
    }
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
        [string]$DefaultIdentifierType = 'matnr',
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

    $primaryType = 'matnr'
    $primaryValue = Get-NormalizedString $identifierMatnr
    $canonicalKey = if ([string]::IsNullOrWhiteSpace($primaryValue)) { '' } else { Get-CanonicalKey -Type $primaryType -Value $primaryValue }

    $identifierSupply = ConvertTo-NullableString (Get-DeepPropertyValue $Material 'identifiers.supply_number')
    $identifierArticle = ConvertTo-NullableString (Get-DeepPropertyValue $Material 'identifiers.article_number')
    $identifierNato = ConvertTo-NullableString (Get-DeepPropertyValue $Material 'identifiers.nato_stock_number')
    $statusCode = Get-NormalizedString (Get-DeepPropertyValue $Material 'status.material_status_code')
    if ([string]::IsNullOrWhiteSpace($statusCode)) {
        $statusCode = 'XX'
    }

    $textShort = Get-NormalizedString (Get-DeepPropertyValue $Material 'texts.short_description')
    $textTechnical = Get-NormalizedString (Get-DeepPropertyValue $Material 'texts.technical_note')
    $textLogistics = Get-NormalizedString (Get-DeepPropertyValue $Material 'texts.logistics_note')
    $classificationExtWg = Get-NormalizedString (Get-DeepPropertyValue $Material 'classification.ext_wg')
    $classificationIsDecentral = [bool](Get-DeepPropertyValue $Material 'classification.is_decentral' $false)
    $classificationCreditor = ConvertTo-NullableString (Get-DeepPropertyValue $Material 'classification.creditor')
    $hazmatIsHazardous = [bool](Get-DeepPropertyValue $Material 'hazmat.is_hazardous' $false)
    $hazmatUnNumber = ConvertTo-NullableString (Get-DeepPropertyValue $Material 'hazmat.un_number')
    $hazmatFlags = ConvertTo-UniqueStringArray (ConvertTo-ObjectArray (Get-DeepPropertyValue $Material 'hazmat.flags' @()))
    $responsibilityCodes = ConvertTo-UniqueStringArray (ConvertTo-ObjectArray (Get-DeepPropertyValue $Material 'assignments.responsibility_codes' @()))
    $assignmentTags = ConvertTo-UniqueStringArray (ConvertTo-ObjectArray (Get-DeepPropertyValue $Material 'assignments.assignment_tags' @()))
    $quantityTarget = [double](Get-DeepPropertyValue $Material 'quantity.target' 0.0)

    $primaryIdentifier = [pscustomobject][ordered]@{
        type  = $primaryType
        value = $primaryValue
    }
    $identifiers = [pscustomobject][ordered]@{
        matnr             = $identifierMatnr
        supply_number     = $identifierSupply
        article_number    = $identifierArticle
        nato_stock_number = $identifierNato
    }
    $status = [pscustomobject][ordered]@{
        material_status_code = $statusCode
    }
    $texts = [pscustomobject][ordered]@{
        short_description = $textShort
        technical_note    = $textTechnical
        logistics_note    = $textLogistics
    }
    $classification = [pscustomobject][ordered]@{
        ext_wg       = $classificationExtWg
        is_decentral = $classificationIsDecentral
        creditor     = $classificationCreditor
    }
    $hazmat = [pscustomobject][ordered]@{
        is_hazardous = $hazmatIsHazardous
        un_number    = $hazmatUnNumber
        flags        = @($hazmatFlags)
    }
    $quantity = [pscustomobject][ordered]@{
        base_unit       = $resolvedBaseUnit
        target          = $quantityTarget
        alternate_units = $alternateUnits.ToArray()
    }
    $assignments = [pscustomobject][ordered]@{
        responsibility_codes = @($responsibilityCodes)
        assignment_tags      = @($assignmentTags)
    }

    return [pscustomobject][ordered]@{
        id                 = $idValue
        canonical_key      = $canonicalKey
        primary_identifier = $primaryIdentifier
        identifiers        = $identifiers
        status             = $status
        texts              = $texts
        classification     = $classification
        hazmat             = $hazmat
        quantity           = $quantity
        alternates         = $alternates.ToArray()
        assignments        = $assignments
    }
}

function Read-LookupFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path $Path)) {
        throw "Lookup file not found: $Path"
    }

    $raw = Get-Content -Path $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) {
        throw "Lookup file is empty: $Path"
    }

    $parsed = $raw | ConvertFrom-Json
    $requiredProperties = @('responsibility_codes', 'assignment_tags', 'hazmat_flags', 'identifier_types', 'unit_codes')
    foreach ($propertyName in $requiredProperties) {
        if (-not $parsed.PSObject.Properties[$propertyName]) {
            throw "Lookup file is missing '$propertyName'"
        }
    }

    return $parsed
}

function Read-DatabaseFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$DefaultIdentifierType = 'matnr',
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
        [void]$materials.Add((ConvertTo-NormalizedMaterial -Material $material -DefaultIdentifierType $DefaultIdentifierType -DefaultUnitCode $DefaultUnitCode))
    }

    $resolvedLookupFile = Get-NormalizedString $parsed.lookup_file
    if ([string]::IsNullOrWhiteSpace($resolvedLookupFile)) {
        $resolvedLookupFile = $Script:LookupFileName
    }

    return [pscustomobject]@{
        schema_version = $schemaVersion
        lookup_file    = $resolvedLookupFile
        materials      = $materials.ToArray()
    }
}

function Get-NextMaterialId {
    param([AllowNull()][object[]]$Materials)

    $maxId = 999
    foreach ($material in @(ConvertTo-ObjectArray $Materials)) {
        $idValue = 0
        if ([int]::TryParse((Get-NormalizedString $material.id), [ref]$idValue) -and $idValue -gt $maxId) {
            $maxId = $idValue
        }
    }

    return ($maxId + 1)
}

function Get-UniqueCloneIdentifierValue {
    param(
        [AllowNull()][string]$BaseValue,
        [Parameter(Mandatory = $true)][string]$IdentifierType,
        [AllowNull()][object[]]$Materials,
        [int]$SuggestedId = 0
    )

    $normalizedBase = Get-NormalizedString $BaseValue
    $baseCandidate = if ([string]::IsNullOrWhiteSpace($normalizedBase)) {
        if ($SuggestedId -gt 0) { "copy-$SuggestedId" } else { 'copy' }
    }
    else {
        "$normalizedBase-copy"
    }

    $counter = 1
    while ($true) {
        $candidate = if ($counter -eq 1) { $baseCandidate } else { "$baseCandidate-$counter" }
        $candidateKey = Get-CanonicalKey -Type $IdentifierType -Value $candidate
        $exists = $false
        foreach ($material in @(ConvertTo-ObjectArray $Materials)) {
            if ((Get-NormalizedString $material.canonical_key) -eq $candidateKey) {
                $exists = $true
                break
            }
        }

        if (-not $exists) {
            return $candidate
        }

        $counter++
    }
}

function Get-MaterialSummary {
    param(
        [Parameter(Mandatory = $true)]$Material,
        [AllowNull()][hashtable]$HazmatFlagLabelMap,
        [AllowNull()][hashtable]$ResponsibilityCodeLabelMap,
        [AllowNull()][hashtable]$AssignmentTagLabelMap
    )

    $description = Get-NormalizedString (Get-DeepPropertyValue $Material 'texts.short_description')
    if ([string]::IsNullOrWhiteSpace($description)) {
        $description = '(no description)'
    }

    $materialNumber = Get-NormalizedString (Get-DeepPropertyValue $Material 'identifiers.matnr')
    $status = Get-NormalizedString (Get-DeepPropertyValue $Material 'status.material_status_code')
    $matnr = Get-NormalizedString (Get-DeepPropertyValue $Material 'identifiers.matnr')
    $supplyNumber = Get-NormalizedString (Get-DeepPropertyValue $Material 'identifiers.supply_number')
    $articleNumber = Get-NormalizedString (Get-DeepPropertyValue $Material 'identifiers.article_number')
    $natoStockNumber = Get-NormalizedString (Get-DeepPropertyValue $Material 'identifiers.nato_stock_number')
    $extWg = Get-NormalizedString (Get-DeepPropertyValue $Material 'classification.ext_wg')
    $creditor = Get-NormalizedString (Get-DeepPropertyValue $Material 'classification.creditor')
    $isHazardous = [bool](Get-DeepPropertyValue $Material 'hazmat.is_hazardous' $false)
    $isDecentral = [bool](Get-DeepPropertyValue $Material 'classification.is_decentral' $false)
    $unNumber = Get-NormalizedString (Get-DeepPropertyValue $Material 'hazmat.un_number')
    $baseUnit = Get-NormalizedString (Get-DeepPropertyValue $Material 'quantity.base_unit')
    $targetQuantity = [double](Get-DeepPropertyValue $Material 'quantity.target' 0.0)
    $technicalNote = Get-NormalizedString (Get-DeepPropertyValue $Material 'texts.technical_note')
    $logisticsNote = Get-NormalizedString (Get-DeepPropertyValue $Material 'texts.logistics_note')
    $alternateCount = (@(ConvertTo-ObjectArray (Get-DeepPropertyValue $Material 'alternates' @()))).Count
    $hazmatFlags = ConvertTo-UniqueStringArray (ConvertTo-ObjectArray (Get-DeepPropertyValue $Material 'hazmat.flags' @()))
    $responsibilityCodes = ConvertTo-UniqueStringArray (ConvertTo-ObjectArray (Get-DeepPropertyValue $Material 'assignments.responsibility_codes' @()))
    $assignmentTags = ConvertTo-UniqueStringArray (ConvertTo-ObjectArray (Get-DeepPropertyValue $Material 'assignments.assignment_tags' @()))
    $hazmatFlagLabels = ConvertTo-LabelTextArray -Values $hazmatFlags -LabelMap $HazmatFlagLabelMap
    $responsibilityLabels = ConvertTo-LabelTextArray -Values $responsibilityCodes -LabelMap $ResponsibilityCodeLabelMap
    $assignmentLabels = ConvertTo-LabelTextArray -Values $assignmentTags -LabelMap $AssignmentTagLabelMap
    $searchIndex = @(
        [int](Get-DeepPropertyValue $Material 'id' 0)
        $materialNumber
        $description
        $status
        $matnr
        $supplyNumber
        $articleNumber
        $natoStockNumber
        $extWg
        $creditor
        $unNumber
        $baseUnit
        $targetQuantity
        $technicalNote
        $logisticsNote
        ($hazmatFlagLabels -join ' ')
        ($responsibilityLabels -join ' ')
        ($assignmentLabels -join ' ')
    ) -join ' '

    return [pscustomobject]@{
        Id                  = [int](Get-DeepPropertyValue $Material 'id' 0)
        MaterialNumber      = $materialNumber
        Description         = $description
        Status              = $status
        SupplyNumber        = $supplyNumber
        ArticleNumber       = $articleNumber
        NatoStockNumber     = $natoStockNumber
        ExtWg               = $extWg
        Creditor            = $creditor
        Hazardous           = $(if ($isHazardous) { 'Ja' } else { '-' })
        Decentral           = $(if ($isDecentral) { 'Ja' } else { '-' })
        Alternates          = $alternateCount
        IsHazardous         = $isHazardous
        IsDecentral         = $isDecentral
        HasAlternates       = ($alternateCount -gt 0)
        UnNumber            = $unNumber
        BaseUnit            = $baseUnit
        TargetQuantity      = $targetQuantity
        TechnicalNote       = $technicalNote
        LogisticsNote       = $logisticsNote
        HazmatFlags         = @($hazmatFlagLabels)
        ResponsibilityCodes = @($responsibilityLabels)
        AssignmentTags      = @($assignmentLabels)
        SearchIndex         = $searchIndex.ToLowerInvariant()
        MaterialRef         = $Material
    }
}

function Test-MaterialSummaryRule {
    param(
        [Parameter(Mandatory = $true)]$Summary,
        [Parameter(Mandatory = $true)]$Rule
    )

    $propertyName = Get-NormalizedString $Rule.PropertyName
    $operator = Get-NormalizedString $Rule.Operator
    $ruleType = Get-NormalizedString $Rule.Type
    $ruleValue = Get-NormalizedString $Rule.Value
    $summaryValue = Get-DeepPropertyValue $Summary $propertyName

    switch ($ruleType) {
        'bool' {
            switch ($operator) {
                'is yes' { return [bool]$summaryValue }
                'is no' { return (-not [bool]$summaryValue) }
                default { return $true }
            }
        }
        'number' {
            $parseResult = ConvertTo-NumberParseResult $ruleValue
            if (-not $parseResult.Success) {
                return $false
            }

            $numericValue = [double]$summaryValue
            switch ($operator) {
                '=' { return ($numericValue -eq [double]$parseResult.Value) }
                '>=' { return ($numericValue -ge [double]$parseResult.Value) }
                '<=' { return ($numericValue -le [double]$parseResult.Value) }
                default { return $true }
            }
        }
        'tags' {
            $normalizedRuleValue = $ruleValue.ToLowerInvariant()
            if ([string]::IsNullOrWhiteSpace($normalizedRuleValue)) {
                return $true
            }

            foreach ($entry in @(ConvertTo-ObjectArray $summaryValue)) {
                $entryText = Get-NormalizedString $entry
                if (-not [string]::IsNullOrWhiteSpace($entryText) -and $entryText.ToLowerInvariant().Contains($normalizedRuleValue)) {
                    return $true
                }
            }

            return $false
        }
        default {
            $normalizedSummaryValue = (Get-NormalizedString $summaryValue).ToLowerInvariant()
            $normalizedRuleValue = $ruleValue.ToLowerInvariant()
            switch ($operator) {
                'equals' { return ($normalizedSummaryValue -eq $normalizedRuleValue) }
                'starts with' { return $normalizedSummaryValue.StartsWith($normalizedRuleValue) }
                default { return $normalizedSummaryValue.Contains($normalizedRuleValue) }
            }
        }
    }
}

function Get-FilteredMaterialSummaries {
    param(
        [AllowNull()][object[]]$Materials,
        [string]$SearchText,
        [switch]$HazardousOnly,
        [switch]$DecentralOnly,
        [AllowNull()][object[]]$AdvancedRules,
        [AllowNull()][hashtable]$HazmatFlagLabelMap,
        [AllowNull()][hashtable]$ResponsibilityCodeLabelMap,
        [AllowNull()][hashtable]$AssignmentTagLabelMap
    )

    $normalizedSearch = (Get-NormalizedString $SearchText).ToLowerInvariant()
    $summaries = New-Object System.Collections.Generic.List[object]

    foreach ($material in @(ConvertTo-ObjectArray $Materials)) {
        $summary = Get-MaterialSummary `
            -Material $material `
            -HazmatFlagLabelMap $HazmatFlagLabelMap `
            -ResponsibilityCodeLabelMap $ResponsibilityCodeLabelMap `
            -AssignmentTagLabelMap $AssignmentTagLabelMap

        if ($HazardousOnly -and -not $summary.IsHazardous) {
            continue
        }

        if ($DecentralOnly -and -not $summary.IsDecentral) {
            continue
        }

        if (-not [string]::IsNullOrWhiteSpace($normalizedSearch) -and -not $summary.SearchIndex.Contains($normalizedSearch)) {
            continue
        }

        $matchesAdvancedRules = $true
        foreach ($rule in @(ConvertTo-ObjectArray $AdvancedRules)) {
            if (-not (Test-MaterialSummaryRule -Summary $summary -Rule $rule)) {
                $matchesAdvancedRules = $false
                break
            }
        }

        if ($matchesAdvancedRules) {
            [void]$summaries.Add($summary)
        }
    }

    return @($summaries | Sort-Object MaterialNumber, Id)
}

function Backup-DatabaseFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (!(Test-Path $Script:LogsDir)) { New-Item -Path $Script:LogsDir -ItemType Directory -Force | Out-Null }
    if (!(Test-Path $Script:BackupDir)) { New-Item -Path $Script:BackupDir -ItemType Directory -Force | Out-Null }

    $backupName = "db_verlegepaket_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $backupPath = Join-Path $Script:BackupDir $backupName
    Copy-Item -Path $Path -Destination $backupPath -Force
    return $backupPath
}

function Save-DatabaseFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object[]]$Materials
    )

    $database = [pscustomobject][ordered]@{
        schema_version = $Script:DatabaseSchemaVersion
        lookup_file    = $Script:LookupFileName
        materials      = @($Materials)
    }

    $database | ConvertTo-Json -Depth 20 | Out-File -FilePath $Path -Encoding UTF8
}

function Start-MaterialBrowserUi {
    $lookupData = Read-LookupFile -Path $Script:LookupPath
    $defaultIdentifierType = Get-NormalizedString $lookupData.identifier_types[0].code
    if ([string]::IsNullOrWhiteSpace($defaultIdentifierType)) {
        $defaultIdentifierType = 'matnr'
    }

    $defaultUnitCode = Get-NormalizedString $lookupData.unit_codes[0].code
    if ([string]::IsNullOrWhiteSpace($defaultUnitCode)) {
        $defaultUnitCode = 'EA'
    }

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Material Browser"
        Height="980"
        Width="1680"
        MinHeight="860"
        MinWidth="1380"
        WindowStartupLocation="CenterScreen"
        Background="#F3F5F7"
        FontFamily="Segoe UI"
        FontSize="11">
    <Window.Resources>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="Margin" Value="0,4,0,10"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="White"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Margin" Value="0,4,0,10"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="White"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Margin" Value="0,4,12,4"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="CanUserAddRows" Value="False"/>
            <Setter Property="CanUserDeleteRows" Value="False"/>
            <Setter Property="CanUserResizeRows" Value="False"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="RowBackground" Value="White"/>
            <Setter Property="AlternatingRowBackground" Value="#F8FAFC"/>
            <Setter Property="BorderBrush" Value="#E2E8F0"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="EnableRowVirtualization" Value="True"/>
            <Setter Property="EnableColumnVirtualization" Value="True"/>
            <Setter Property="VirtualizingPanel.IsVirtualizing" Value="True"/>
            <Setter Property="VirtualizingPanel.VirtualizationMode" Value="Recycling"/>
            <Setter Property="Margin" Value="0,8,0,0"/>
        </Style>
        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Padding" Value="14,10"/>
            <Setter Property="Margin" Value="0,0,10,0"/>
            <Setter Property="Background" Value="#0F766E"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderBrush" Value="#0F766E"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style x:Key="SecondaryButton" TargetType="Button">
            <Setter Property="Padding" Value="14,10"/>
            <Setter Property="Margin" Value="0,0,10,0"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="Foreground" Value="#0F172A"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style x:Key="DangerButton" TargetType="Button">
            <Setter Property="Padding" Value="14,10"/>
            <Setter Property="Margin" Value="0,0,10,0"/>
            <Setter Property="Background" Value="#BE123C"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderBrush" Value="#BE123C"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style x:Key="PanelTitle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#0F172A"/>
            <Setter Property="FontSize" Value="16"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="Padding" Value="14,8"/>
        </Style>
    </Window.Resources>
    <DockPanel LastChildFill="True">
        <Border DockPanel.Dock="Top" Background="#0F172A" Padding="22,18,22,18">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                    <TextBlock Text="Material Browser" FontFamily="Bahnschrift SemiBold" FontSize="28" Foreground="White"/>
                    <TextBlock Text="Fast editor for db_verlegepaket.json with full-schema validation and backup-on-save." Margin="0,4,0,0" Foreground="#CBD5E1"/>
                    <TextBlock x:Name="txtDbPath" Margin="0,8,0,0" Foreground="#93C5FD"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top">
                    <Button x:Name="btnNewMaterial" Content="Neu" Style="{StaticResource SecondaryButton}"/>
                    <Button x:Name="btnSaveDatabase" Content="Speichern" Style="{StaticResource PrimaryButton}"/>
                    <Button x:Name="btnReloadDatabase" Content="Neu laden" Style="{StaticResource SecondaryButton}" Margin="0"/>
                </StackPanel>
            </Grid>
        </Border>

        <Border DockPanel.Dock="Bottom" Background="White" BorderBrush="#E2E8F0" BorderThickness="1,1,0,0" Padding="18,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="16"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="txtDirtyState" Grid.Column="0" FontWeight="SemiBold" Foreground="#475569" VerticalAlignment="Center" Text="Saved"/>
                <TextBlock x:Name="txtStatus" Grid.Column="2" Foreground="#475569" VerticalAlignment="Center" Text="Ready"/>
                <TextBlock Grid.Column="3" Foreground="#94A3B8" VerticalAlignment="Center" Text="Ctrl+S Speichern  |  Ctrl+N Neu  |  Ctrl+D Duplizieren  |  Esc Verwerfen  |  Delete Loeschen"/>
            </Grid>
        </Border>

        <Grid Margin="20">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="430"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <Border Grid.Column="0" Background="White" BorderBrush="#E2E8F0" BorderThickness="1" CornerRadius="10" Padding="18">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <StackPanel Grid.Row="0">
                        <TextBlock Style="{StaticResource PanelTitle}" Text="Materialliste"/>
                        <TextBlock Margin="0,4,0,0" Foreground="#64748B" Text="Fast filter path, denser overview, and staged edits for larger databases."/>
                    </StackPanel>

                    <Grid Grid.Row="1" Margin="0,16,0,12">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="110"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="txtSearch" Grid.Column="0" Margin="0" ToolTip="Search by ID, material number, description, status, supply number, or article number."/>
                        <Button x:Name="btnOpenFilterMenu" Grid.Column="1" Content="Filter..." Style="{StaticResource SecondaryButton}" Margin="10,0,0,0"/>
                        <Button x:Name="btnOpenColumnMenu" Grid.Column="2" Content="Columns..." Style="{StaticResource SecondaryButton}" Margin="10,0,0,0"/>
                        <Button x:Name="btnClearSearch" Grid.Column="3" Content="Reset" Style="{StaticResource SecondaryButton}" Margin="10,0,0,0"/>
                    </Grid>

                    <WrapPanel Grid.Row="2" Margin="0,0,0,12">
                        <CheckBox x:Name="chkFilterHazardous" Content="Gefahrgut"/>
                        <CheckBox x:Name="chkFilterDecentral" Content="Dezentral"/>
                    </WrapPanel>

                    <Grid Grid.Row="3">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <TextBlock x:Name="txtListMeta" Foreground="#64748B" Text="0 materials"/>
                        <DataGrid x:Name="dgMaterials" Grid.Row="1" Margin="0,8,0,0" SelectionMode="Single" SelectionUnit="FullRow" IsReadOnly="True" RowHeaderWidth="0">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Materialnummer" Binding="{Binding MaterialNumber}" Width="1.2*" />
                                <DataGridTextColumn Header="Beschreibung" Binding="{Binding Description}" Width="1.5*" />
                                <DataGridTextColumn Header="ID" Binding="{Binding Id}" Width="76"/>
                                <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="76"/>
                                <DataGridTextColumn Header="Supply Number" Binding="{Binding SupplyNumber}" Width="1.1*"/>
                                <DataGridTextColumn Header="Article Number" Binding="{Binding ArticleNumber}" Width="1.1*"/>
                                <DataGridTextColumn Header="NATO Stock Number" Binding="{Binding NatoStockNumber}" Width="1.1*"/>
                                <DataGridTextColumn Header="Ext WG" Binding="{Binding ExtWg}" Width="88"/>
                                <DataGridTextColumn Header="Creditor" Binding="{Binding Creditor}" Width="100"/>
                                <DataGridTextColumn Header="Gefahrgut" Binding="{Binding Hazardous}" Width="78"/>
                                <DataGridTextColumn Header="Dezentral" Binding="{Binding Decentral}" Width="88"/>
                                <DataGridTextColumn Header="Alternativen" Binding="{Binding Alternates}" Width="92"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Grid>
                </Grid>
            </Border>

            <GridSplitter Grid.Column="1" Width="8" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Background="Transparent"/>
            <Border Grid.Column="2" Background="White" BorderBrush="#E2E8F0" BorderThickness="1" CornerRadius="10" Padding="18">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <Grid Grid.Row="0">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0">
                            <TextBlock x:Name="txtEditorHeadline" FontFamily="Bahnschrift SemiBold" FontSize="24" Foreground="#0F172A" Text="No material selected"/>
                            <TextBlock x:Name="txtEditorSubheadline" Margin="0,4,0,0" Foreground="#64748B" Text="Select a material or create a new one."/>
                        </StackPanel>
                        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top">
                            <Button x:Name="btnCloneMaterial" Content="Duplizieren" Style="{StaticResource SecondaryButton}"/>
                            <Button x:Name="btnRevertCurrent" Content="Aenderungen verwerfen" Style="{StaticResource SecondaryButton}"/>
                            <Button x:Name="btnDeleteMaterial" Content="Loeschen" Style="{StaticResource DangerButton}" Margin="0"/>
                        </StackPanel>
                    </Grid>

                    <TabControl x:Name="tabEditor" Grid.Row="1" Margin="0,18,0,0">
                        <TabItem Header="General">
                            <ScrollViewer VerticalScrollBarVisibility="Auto">
                                <Grid Margin="6">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="1.15*"/>
                                        <ColumnDefinition Width="24"/>
                                        <ColumnDefinition Width="1*"/>
                                    </Grid.ColumnDefinitions>
                                    <StackPanel Grid.Column="0">
                                        <TextBlock Foreground="#334155" FontWeight="SemiBold" Text="Material"/>
                                        <Grid Margin="0,10,0,0"><Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions><TextBlock Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="ID"/><TextBox x:Name="txtId" Grid.Column="1"/></Grid>
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions><TextBlock Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Materialnummer"/><TextBox x:Name="txtMaterialNumber" Grid.Column="1"/></Grid>
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions><TextBlock Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Supply number"/><TextBox x:Name="txtSupplyNumber" Grid.Column="1"/></Grid>
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions><TextBlock Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="NATO stock number"/><TextBox x:Name="txtNatoStockNumber" Grid.Column="1"/></Grid>
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions><TextBlock Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Ext WG"/><TextBox x:Name="txtExtWg" Grid.Column="1"/></Grid>
                                        <TextBlock Margin="0,18,0,0" Foreground="#334155" FontWeight="SemiBold" Text="Dezentral"/>
                                        <CheckBox x:Name="chkIsDecentral" Margin="0,10,0,0" Content="Is decentral"/>
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions><TextBlock Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Article number"/><TextBox x:Name="txtArticleNumber" Grid.Column="1"/></Grid>
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions><TextBlock Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Creditor"/><TextBox x:Name="txtCreditor" Grid.Column="1"/></Grid>
                                    </StackPanel>
                                    <StackPanel Grid.Column="2">
                                        <TextBlock Foreground="#334155" FontWeight="SemiBold" Text="Status and Notes"/>
                                        <Grid Margin="0,10,0,0"><Grid.ColumnDefinitions><ColumnDefinition Width="170"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions><TextBlock Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Material status"/><TextBox x:Name="txtMaterialStatus" Grid.Column="1"/></Grid>
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="170"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions><TextBlock Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Short description"/><TextBox x:Name="txtShortDescription" Grid.Column="1"/></Grid>
                                        <TextBlock Margin="0,10,0,0" Foreground="#475569" Text="Technical note"/>
                                        <TextBox x:Name="txtTechnicalNote" Height="110" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"/>
                                        <TextBlock Foreground="#475569" Text="Logistics note"/>
                                        <TextBox x:Name="txtLogisticsNote" Height="110" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"/>
                                    </StackPanel>
                                </Grid>
                            </ScrollViewer>
                        </TabItem>
                        <TabItem Header="Hazmat / Assignments">
                            <ScrollViewer VerticalScrollBarVisibility="Auto">
                                <Grid Margin="6">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="1*"/>
                                        <ColumnDefinition Width="24"/>
                                        <ColumnDefinition Width="1.2*"/>
                                    </Grid.ColumnDefinitions>
                                    <StackPanel Grid.Column="0">
                                        <TextBlock Foreground="#334155" FontWeight="SemiBold" Text="Hazmat"/>
                                        <CheckBox x:Name="chkIsHazardous" Margin="0,12,0,4" Content="Is hazardous"/>
                                        <Grid Margin="0,6,0,0"><Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions><TextBlock Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="UN number"/><TextBox x:Name="txtUnNumber" Grid.Column="1"/></Grid>
                                        <TextBlock Margin="0,10,0,6" Foreground="#334155" FontWeight="SemiBold" Text="Hazmat flags"/>
                                        <ScrollViewer Height="260" VerticalScrollBarVisibility="Auto"><WrapPanel x:Name="pnlHazmatFlags"/></ScrollViewer>
                                    </StackPanel>
                                    <Grid Grid.Column="2">
                                        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                                        <TextBlock Grid.Row="0" Foreground="#334155" FontWeight="SemiBold" Text="Responsibility codes"/>
                                        <ScrollViewer Grid.Row="1" Height="170" VerticalScrollBarVisibility="Auto" Margin="0,8,0,12"><WrapPanel x:Name="pnlResponsibilityCodes"/></ScrollViewer>
                                        <TextBlock Grid.Row="2" Foreground="#334155" FontWeight="SemiBold" Text="Assignment tags"/>
                                        <ScrollViewer Grid.Row="3" Height="250" VerticalScrollBarVisibility="Auto" Margin="0,8,0,0"><WrapPanel x:Name="pnlAssignmentTags"/></ScrollViewer>
                                    </Grid>
                                </Grid>
                            </ScrollViewer>
                        </TabItem>
                        <TabItem Header="Quantity / Alternates">
                            <ScrollViewer VerticalScrollBarVisibility="Auto">
                                <StackPanel Margin="6">
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="220"/><ColumnDefinition Width="180"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                        <StackPanel Grid.Column="0"><TextBlock Foreground="#334155" FontWeight="SemiBold" Text="Base unit"/><ComboBox x:Name="cmbBaseUnit" DisplayMemberPath="label" SelectedValuePath="code"/></StackPanel>
                                        <StackPanel Grid.Column="1" Margin="16,0,0,0"><TextBlock Foreground="#334155" FontWeight="SemiBold" Text="Target quantity"/><TextBox x:Name="txtQuantityTarget"/></StackPanel>
                                        <Border Grid.Column="2" Margin="16,0,0,0" Background="#F8FAFC" BorderBrush="#E2E8F0" BorderThickness="1" CornerRadius="8" Padding="14"><TextBlock Foreground="#475569" TextWrapping="Wrap" Text="Use alternate units for pack conversions and alternates for replacement materials. Nested grids are staged in memory and validated before save."/></Border>
                                    </Grid>
                                    <Grid Margin="0,18,0,0">
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="1*"/><ColumnDefinition Width="24"/><ColumnDefinition Width="1.3*"/></Grid.ColumnDefinitions>
                                        <StackPanel Grid.Column="0">
                                            <DockPanel LastChildFill="False"><TextBlock DockPanel.Dock="Left" Foreground="#334155" FontWeight="SemiBold" Text="Alternate units"/><Button x:Name="btnAddAlternateUnit" DockPanel.Dock="Right" Content="+ Einheit" Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"/><Button x:Name="btnRemoveAlternateUnit" DockPanel.Dock="Right" Content="- Einheit" Style="{StaticResource SecondaryButton}" Margin="0"/></DockPanel>
                                            <DataGrid x:Name="gridAlternateUnits" Height="300" CanUserResizeColumns="True" RowHeaderWidth="0">
                                                <DataGrid.Columns>
                                                    <DataGridComboBoxColumn Header="Unit code" Width="150" DisplayMemberPath="label" SelectedValuePath="code" ItemsSource="{Binding DataContext.UnitCodeOptions, RelativeSource={RelativeSource AncestorType=Window}}" SelectedValueBinding="{Binding unit_code, UpdateSourceTrigger=PropertyChanged}"/>
                                                    <DataGridTextColumn Header="Conversion to base" Width="*" Binding="{Binding conversion_to_base, UpdateSourceTrigger=PropertyChanged}"/>
                                                </DataGrid.Columns>
                                            </DataGrid>
                                        </StackPanel>
                                        <StackPanel Grid.Column="2">
                                            <DockPanel LastChildFill="False"><TextBlock DockPanel.Dock="Left" Foreground="#334155" FontWeight="SemiBold" Text="Alternates"/><Button x:Name="btnAddAlternate" DockPanel.Dock="Right" Content="+ Alternate" Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"/><Button x:Name="btnRemoveAlternate" DockPanel.Dock="Right" Content="- Alternate" Style="{StaticResource SecondaryButton}" Margin="0"/></DockPanel>
                                            <DataGrid x:Name="gridAlternates" Height="300" CanUserResizeColumns="True" RowHeaderWidth="0">
                                                <DataGrid.Columns>
                                                    <DataGridTextColumn Header="Pos" Width="72" Binding="{Binding position, UpdateSourceTrigger=PropertyChanged}"/>
                                                    <DataGridTextColumn Header="Materialnummer" Width="*" Binding="{Binding identifier_value, UpdateSourceTrigger=PropertyChanged}"/>
                                                    <DataGridTextColumn Header="Mat status" Width="110" Binding="{Binding material_status_code, UpdateSourceTrigger=PropertyChanged}"/>
                                                    <DataGridComboBoxColumn Header="Preferred unit" Width="150" DisplayMemberPath="label" SelectedValuePath="code" ItemsSource="{Binding DataContext.UnitCodeOptions, RelativeSource={RelativeSource AncestorType=Window}}" SelectedValueBinding="{Binding preferred_unit_code, UpdateSourceTrigger=PropertyChanged}"/>
                                                </DataGrid.Columns>
                                            </DataGrid>
                                        </StackPanel>
                                    </Grid>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>
                    </TabControl>
                </Grid>
            </Border>
        </Grid>
    </DockPanel>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader([xml]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $txtDbPath = $window.FindName('txtDbPath')
    $btnNewMaterial = $window.FindName('btnNewMaterial')
    $btnSaveDatabase = $window.FindName('btnSaveDatabase')
    $btnReloadDatabase = $window.FindName('btnReloadDatabase')
    $txtDirtyState = $window.FindName('txtDirtyState')
    $txtStatus = $window.FindName('txtStatus')
    $txtSearch = $window.FindName('txtSearch')
    $btnOpenFilterMenu = $window.FindName('btnOpenFilterMenu')
    $btnOpenColumnMenu = $window.FindName('btnOpenColumnMenu')
    $btnClearSearch = $window.FindName('btnClearSearch')
    $chkFilterHazardous = $window.FindName('chkFilterHazardous')
    $chkFilterDecentral = $window.FindName('chkFilterDecentral')
    $txtListMeta = $window.FindName('txtListMeta')
    $dgMaterials = $window.FindName('dgMaterials')
    $txtEditorHeadline = $window.FindName('txtEditorHeadline')
    $txtEditorSubheadline = $window.FindName('txtEditorSubheadline')
    $btnCloneMaterial = $window.FindName('btnCloneMaterial')
    $btnRevertCurrent = $window.FindName('btnRevertCurrent')
    $btnDeleteMaterial = $window.FindName('btnDeleteMaterial')
    $txtId = $window.FindName('txtId')
    $txtMaterialNumber = $window.FindName('txtMaterialNumber')
    $txtSupplyNumber = $window.FindName('txtSupplyNumber')
    $txtArticleNumber = $window.FindName('txtArticleNumber')
    $txtNatoStockNumber = $window.FindName('txtNatoStockNumber')
    $txtMaterialStatus = $window.FindName('txtMaterialStatus')
    $txtShortDescription = $window.FindName('txtShortDescription')
    $txtTechnicalNote = $window.FindName('txtTechnicalNote')
    $txtLogisticsNote = $window.FindName('txtLogisticsNote')
    $txtExtWg = $window.FindName('txtExtWg')
    $chkIsDecentral = $window.FindName('chkIsDecentral')
    $txtCreditor = $window.FindName('txtCreditor')
    $chkIsHazardous = $window.FindName('chkIsHazardous')
    $txtUnNumber = $window.FindName('txtUnNumber')
    $pnlHazmatFlags = $window.FindName('pnlHazmatFlags')
    $pnlResponsibilityCodes = $window.FindName('pnlResponsibilityCodes')
    $pnlAssignmentTags = $window.FindName('pnlAssignmentTags')
    $cmbBaseUnit = $window.FindName('cmbBaseUnit')
    $txtQuantityTarget = $window.FindName('txtQuantityTarget')
    $btnAddAlternateUnit = $window.FindName('btnAddAlternateUnit')
    $btnRemoveAlternateUnit = $window.FindName('btnRemoveAlternateUnit')
    $gridAlternateUnits = $window.FindName('gridAlternateUnits')
    $btnAddAlternate = $window.FindName('btnAddAlternate')
    $btnRemoveAlternate = $window.FindName('btnRemoveAlternate')
    $gridAlternates = $window.FindName('gridAlternates')

    $txtDbPath.Text = $Script:DbPath
    $window.DataContext = [pscustomobject]@{
        UnitCodeOptions = @($lookupData.unit_codes)
    }

    $cmbBaseUnit.ItemsSource = @($lookupData.unit_codes)

    $state = [ordered]@{
        LookupData              = $lookupData
        Materials               = New-Object System.Collections.ArrayList
        CurrentMaterial         = $null
        CurrentSummary          = $null
        DatabaseDirty           = $false
        EditorDirty             = $false
        PopulatingEditor        = $false
        SuppressSelectionChange = $false
        FilterHazardousOnly     = $false
        FilterDecentralOnly     = $false
        ActiveFilterRules       = @()
        VisibleColumnKeys       = @('material_number', 'description')
    }

    $hazmatFlagLabelMap = @{}
    foreach ($entry in @($lookupData.hazmat_flags)) {
        $hazmatFlagLabelMap[(Get-NormalizedString $entry.code)] = Get-NormalizedString $entry.label
    }

    $responsibilityCodeLabelMap = @{}
    foreach ($entry in @($lookupData.responsibility_codes)) {
        $responsibilityCodeLabelMap[(Get-NormalizedString $entry.code)] = Get-NormalizedString $entry.label
    }

    $assignmentTagLabelMap = @{}
    foreach ($entry in @($lookupData.assignment_tags)) {
        $assignmentTagLabelMap[(Get-NormalizedString $entry.code)] = Get-NormalizedString $entry.label
    }

    $listColumnDefinitions = @(
        [pscustomobject]@{ Key = 'material_number'; Label = 'Materialnummer'; Index = 0 }
        [pscustomobject]@{ Key = 'description'; Label = 'Beschreibung'; Index = 1 }
        [pscustomobject]@{ Key = 'id'; Label = 'ID'; Index = 2 }
        [pscustomobject]@{ Key = 'status'; Label = 'Status'; Index = 3 }
        [pscustomobject]@{ Key = 'supply_number'; Label = 'Supply Number'; Index = 4 }
        [pscustomobject]@{ Key = 'article_number'; Label = 'Article Number'; Index = 5 }
        [pscustomobject]@{ Key = 'nato_stock_number'; Label = 'NATO Stock Number'; Index = 6 }
        [pscustomobject]@{ Key = 'ext_wg'; Label = 'Ext WG'; Index = 7 }
        [pscustomobject]@{ Key = 'creditor'; Label = 'Creditor'; Index = 8 }
        [pscustomobject]@{ Key = 'hazardous'; Label = 'Gefahrgut'; Index = 9 }
        [pscustomobject]@{ Key = 'decentral'; Label = 'Dezentral'; Index = 10 }
        [pscustomobject]@{ Key = 'alternates'; Label = 'Alternativen'; Index = 11 }
    )

    $filterDefinitions = @(
        [pscustomobject]@{ Key = 'material_number'; Label = 'Materialnummer'; PropertyName = 'MaterialNumber'; Type = 'text' }
        [pscustomobject]@{ Key = 'description'; Label = 'Beschreibung'; PropertyName = 'Description'; Type = 'text' }
        [pscustomobject]@{ Key = 'id'; Label = 'ID'; PropertyName = 'Id'; Type = 'number' }
        [pscustomobject]@{ Key = 'status'; Label = 'Status'; PropertyName = 'Status'; Type = 'text' }
        [pscustomobject]@{ Key = 'supply_number'; Label = 'Supply Number'; PropertyName = 'SupplyNumber'; Type = 'text' }
        [pscustomobject]@{ Key = 'article_number'; Label = 'Article Number'; PropertyName = 'ArticleNumber'; Type = 'text' }
        [pscustomobject]@{ Key = 'nato_stock_number'; Label = 'NATO Stock Number'; PropertyName = 'NatoStockNumber'; Type = 'text' }
        [pscustomobject]@{ Key = 'ext_wg'; Label = 'Ext WG'; PropertyName = 'ExtWg'; Type = 'text' }
        [pscustomobject]@{ Key = 'creditor'; Label = 'Creditor'; PropertyName = 'Creditor'; Type = 'text' }
        [pscustomobject]@{ Key = 'is_hazardous'; Label = 'Gefahrgut'; PropertyName = 'IsHazardous'; Type = 'bool' }
        [pscustomobject]@{ Key = 'is_decentral'; Label = 'Dezentral'; PropertyName = 'IsDecentral'; Type = 'bool' }
        [pscustomobject]@{ Key = 'alternates'; Label = 'Alternativen'; PropertyName = 'Alternates'; Type = 'number' }
        [pscustomobject]@{ Key = 'un_number'; Label = 'UN Number'; PropertyName = 'UnNumber'; Type = 'text' }
        [pscustomobject]@{ Key = 'base_unit'; Label = 'Base Unit'; PropertyName = 'BaseUnit'; Type = 'text' }
        [pscustomobject]@{ Key = 'target_quantity'; Label = 'Target Quantity'; PropertyName = 'TargetQuantity'; Type = 'number' }
        [pscustomobject]@{ Key = 'technical_note'; Label = 'Technical Note'; PropertyName = 'TechnicalNote'; Type = 'text' }
        [pscustomobject]@{ Key = 'logistics_note'; Label = 'Logistics Note'; PropertyName = 'LogisticsNote'; Type = 'text' }
        [pscustomobject]@{ Key = 'hazmat_flags'; Label = 'Hazmat Flags'; PropertyName = 'HazmatFlags'; Type = 'tags' }
        [pscustomobject]@{ Key = 'responsibility_codes'; Label = 'Responsibility Codes'; PropertyName = 'ResponsibilityCodes'; Type = 'tags' }
        [pscustomobject]@{ Key = 'assignment_tags'; Label = 'Assignment Tags'; PropertyName = 'AssignmentTags'; Type = 'tags' }
    )

    $filterDefinitionLookup = @{}
    foreach ($definition in $filterDefinitions) {
        $filterDefinitionLookup[$definition.Key] = $definition
    }

    $SetStatus = $null
    $UpdateEditorActions = $null
    $RefreshList = $null
    $PopulateEditor = $null
    $CommitCurrentEditor = $null
    $LoadDatabase = $null
    $SaveDatabase = $null
    $ConfirmClose = $null
    $ApplyListColumnVisibility = $null
    $GetFilterRuleLabel = $null
    $OpenFilterDialog = $null
    $OpenColumnDialog = $null

    $SetStatus = {
        param(
            [string]$Message,
            [string]$Level = 'Info'
        )

        $txtStatus.Text = $Message
        switch ($Level) {
            'Error' { $txtStatus.Foreground = '#B91C1C' }
            'Warning' { $txtStatus.Foreground = '#B45309' }
            'Success' { $txtStatus.Foreground = '#0F766E' }
            default { $txtStatus.Foreground = '#475569' }
        }
    }

    $ApplyListColumnVisibility = {
        foreach ($columnDefinition in $listColumnDefinitions) {
            $dgMaterials.Columns[$columnDefinition.Index].Visibility = if ($state.VisibleColumnKeys -contains $columnDefinition.Key) {
                [System.Windows.Visibility]::Visible
            }
            else {
                [System.Windows.Visibility]::Collapsed
            }
        }
    }

    $GetFilterRuleLabel = {
        param($Rule)

        $label = Get-NormalizedString $Rule.Label
        $operator = Get-NormalizedString $Rule.Operator
        $value = Get-NormalizedString $Rule.Value
        if ((Get-NormalizedString $Rule.Type) -eq 'bool') {
            return "$label $operator"
        }

        return "$label $operator $value"
    }

    $OpenFilterDialog = {
        $dialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Advanced Filters"
        SizeToContent="WidthAndHeight"
        MinWidth="900"
        ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner"
        Background="#F8FAFC">
    <Grid Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0">
            <TextBlock FontFamily="Bahnschrift SemiBold" FontSize="22" Foreground="#0F172A" Text="Advanced Filters"/>
            <TextBlock Margin="0,6,0,0" Foreground="#64748B" Text="Combine multiple rules with AND. Text supports contains, equals, and starts with."/>
        </StackPanel>
        <ScrollViewer Grid.Row="1" Margin="0,16,0,0" Height="280" VerticalScrollBarVisibility="Auto">
            <StackPanel x:Name="pnlRuleRows"/>
        </ScrollViewer>
        <DockPanel Grid.Row="2" Margin="0,16,0,0" LastChildFill="False">
            <Button x:Name="btnAddRule" DockPanel.Dock="Left" Width="110" Content="Add rule"/>
            <Button x:Name="btnClearRules" DockPanel.Dock="Left" Width="110" Margin="10,0,0,0" Content="Clear all"/>
            <Button x:Name="btnCancelRules" DockPanel.Dock="Right" Width="110" Margin="10,0,0,0" Content="Cancel"/>
            <Button x:Name="btnApplyRules" DockPanel.Dock="Right" Width="110" Content="Apply"/>
        </DockPanel>
    </Grid>
</Window>
"@

        $dialogReader = New-Object System.Xml.XmlNodeReader([xml]$dialogXaml)
        $dialog = [Windows.Markup.XamlReader]::Load($dialogReader)
        $dialog.Owner = $window
        $pnlRuleRows = $dialog.FindName('pnlRuleRows')
        $btnAddRule = $dialog.FindName('btnAddRule')
        $btnClearRules = $dialog.FindName('btnClearRules')
        $btnCancelRules = $dialog.FindName('btnCancelRules')
        $btnApplyRules = $dialog.FindName('btnApplyRules')
        $rowStates = New-Object System.Collections.Generic.List[object]

        $setRowOperators = {
            param($rowState, [AllowNull()][string]$PreferredOperator)

            $selectedDefinition = $rowState.PropertyCombo.SelectedItem
            $operators = switch ((Get-NormalizedString $selectedDefinition.Type)) {
                'bool' { @('is yes', 'is no') }
                'number' { @('=', '>=', '<=') }
                'tags' { @('contains') }
                default { @('contains', 'equals', 'starts with') }
            }

            $rowState.OperatorCombo.ItemsSource = @($operators)
            $selectedOperator = if (-not [string]::IsNullOrWhiteSpace((Get-NormalizedString $PreferredOperator)) -and ($operators -contains $PreferredOperator)) {
                $PreferredOperator
            }
            else {
                $operators[0]
            }

            $rowState.OperatorCombo.SelectedItem = $selectedOperator
            $rowState.ValueBox.IsEnabled = ((Get-NormalizedString $selectedDefinition.Type) -ne 'bool')
            if (-not $rowState.ValueBox.IsEnabled) {
                $rowState.ValueBox.Text = ''
            }
        }

        $addRuleRow = {
            param([AllowNull()]$ExistingRule)

            $rowGrid = New-Object System.Windows.Controls.Grid
            $rowGrid.Margin = '0,0,0,10'
            $columnDefinition = New-Object System.Windows.Controls.ColumnDefinition
            $columnDefinition.Width = [System.Windows.GridLength]::new(240)
            [void]$rowGrid.ColumnDefinitions.Add($columnDefinition)
            $columnDefinition = New-Object System.Windows.Controls.ColumnDefinition
            $columnDefinition.Width = [System.Windows.GridLength]::new(140)
            [void]$rowGrid.ColumnDefinitions.Add($columnDefinition)
            $columnDefinition = New-Object System.Windows.Controls.ColumnDefinition
            $columnDefinition.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
            [void]$rowGrid.ColumnDefinitions.Add($columnDefinition)
            $columnDefinition = New-Object System.Windows.Controls.ColumnDefinition
            $columnDefinition.Width = [System.Windows.GridLength]::new(90)
            [void]$rowGrid.ColumnDefinitions.Add($columnDefinition)

            $propertyCombo = New-Object System.Windows.Controls.ComboBox
            $propertyCombo.DisplayMemberPath = 'Label'
            $propertyCombo.SelectedValuePath = 'Key'
            $propertyCombo.ItemsSource = @($filterDefinitions)
            $propertyCombo.Margin = '0,0,10,0'

            $operatorCombo = New-Object System.Windows.Controls.ComboBox
            $operatorCombo.Margin = '0,0,10,0'
            [System.Windows.Controls.Grid]::SetColumn($operatorCombo, 1)

            $valueBox = New-Object System.Windows.Controls.TextBox
            $valueBox.Margin = '0,0,10,0'
            [System.Windows.Controls.Grid]::SetColumn($valueBox, 2)

            $removeButton = New-Object System.Windows.Controls.Button
            $removeButton.Content = 'Remove'
            [System.Windows.Controls.Grid]::SetColumn($removeButton, 3)

            [void]$rowGrid.Children.Add($propertyCombo)
            [void]$rowGrid.Children.Add($operatorCombo)
            [void]$rowGrid.Children.Add($valueBox)
            [void]$rowGrid.Children.Add($removeButton)
            [void]$pnlRuleRows.Children.Add($rowGrid)

            $rowState = [pscustomobject]@{
                Container    = $rowGrid
                PropertyCombo = $propertyCombo
                OperatorCombo = $operatorCombo
                ValueBox      = $valueBox
            }
            [void]$rowStates.Add($rowState)

            $updateRowOperators = $setRowOperators
            $currentRowState = $rowState
            $ruleRowsPanel = $pnlRuleRows
            $ruleStates = $rowStates

            $propertyCombo.Add_SelectionChanged({
                    & $updateRowOperators -rowState $currentRowState -PreferredOperator $null
                }.GetNewClosure())

            $removeButton.Add_Click({
                    [void]$ruleRowsPanel.Children.Remove($currentRowState.Container)
                    [void]$ruleStates.Remove($currentRowState)
                }.GetNewClosure())

            if ($null -ne $ExistingRule -and $filterDefinitionLookup.ContainsKey((Get-NormalizedString $ExistingRule.Key))) {
                $propertyCombo.SelectedItem = $filterDefinitionLookup[(Get-NormalizedString $ExistingRule.Key)]
                & $setRowOperators -rowState $rowState -PreferredOperator (Get-NormalizedString $ExistingRule.Operator)
                if ((Get-NormalizedString $ExistingRule.Type) -ne 'bool') {
                    $valueBox.Text = Get-NormalizedString $ExistingRule.Value
                }
            }
            else {
                $propertyCombo.SelectedItem = $filterDefinitions[0]
                & $setRowOperators -rowState $rowState -PreferredOperator $null
            }
        }

        if (@($state.ActiveFilterRules).Count -gt 0) {
            foreach ($rule in @($state.ActiveFilterRules)) {
                & $addRuleRow -ExistingRule $rule
            }
        }
        else {
            & $addRuleRow -ExistingRule $null
        }

        $btnAddRule.Add_Click({ & $addRuleRow -ExistingRule $null })
        $btnClearRules.Add_Click({
                $pnlRuleRows.Children.Clear()
                $rowStates.Clear()
            })
        $btnCancelRules.Add_Click({
                $dialog.DialogResult = $false
                $dialog.Close()
            })
        $btnApplyRules.Add_Click({
                $appliedRules = New-Object System.Collections.Generic.List[object]
                foreach ($rowState in @($rowStates.ToArray())) {
                    $selectedDefinition = $rowState.PropertyCombo.SelectedItem
                    if ($null -eq $selectedDefinition) {
                        continue
                    }

                    $selectedType = Get-NormalizedString $selectedDefinition.Type
                    $ruleValue = Get-NormalizedString $rowState.ValueBox.Text
                    if ($selectedType -ne 'bool' -and [string]::IsNullOrWhiteSpace($ruleValue)) {
                        [System.Windows.MessageBox]::Show("Filter '$($selectedDefinition.Label)' requires a value.", 'Advanced filters', 'OK', 'Warning') | Out-Null
                        return
                    }

                    [void]$appliedRules.Add([pscustomobject]@{
                            Key          = Get-NormalizedString $selectedDefinition.Key
                            Label        = Get-NormalizedString $selectedDefinition.Label
                            PropertyName = Get-NormalizedString $selectedDefinition.PropertyName
                            Type         = $selectedType
                            Operator     = Get-NormalizedString $rowState.OperatorCombo.SelectedItem
                            Value        = $ruleValue
                        })
                }

                $state.ActiveFilterRules = @($appliedRules.ToArray())
                $dialog.DialogResult = $true
                $dialog.Close()
            })

        return [bool]$dialog.ShowDialog()
    }

    $OpenColumnDialog = {
        $dialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Visible Columns"
        SizeToContent="WidthAndHeight"
        MinWidth="360"
        ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner"
        Background="#F8FAFC">
    <Grid Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0">
            <TextBlock FontFamily="Bahnschrift SemiBold" FontSize="22" Foreground="#0F172A" Text="Visible Columns"/>
            <TextBlock Margin="0,6,0,0" Foreground="#64748B" Text="Choose which columns are shown in the material list for this session."/>
        </StackPanel>
        <StackPanel x:Name="pnlColumns" Grid.Row="1" Margin="0,16,0,0"/>
        <DockPanel Grid.Row="2" Margin="0,16,0,0" LastChildFill="False">
            <Button x:Name="btnDefaultColumns" DockPanel.Dock="Left" Width="120" Content="Defaults"/>
            <Button x:Name="btnCancelColumns" DockPanel.Dock="Right" Width="110" Margin="10,0,0,0" Content="Cancel"/>
            <Button x:Name="btnApplyColumns" DockPanel.Dock="Right" Width="110" Content="Apply"/>
        </DockPanel>
    </Grid>
</Window>
"@

        $dialogReader = New-Object System.Xml.XmlNodeReader([xml]$dialogXaml)
        $dialog = [Windows.Markup.XamlReader]::Load($dialogReader)
        $dialog.Owner = $window
        $pnlColumns = $dialog.FindName('pnlColumns')
        $btnDefaultColumns = $dialog.FindName('btnDefaultColumns')
        $btnCancelColumns = $dialog.FindName('btnCancelColumns')
        $btnApplyColumns = $dialog.FindName('btnApplyColumns')
        $checkBoxStates = New-Object System.Collections.Generic.List[object]

        foreach ($columnDefinition in $listColumnDefinitions) {
            $checkBox = New-Object System.Windows.Controls.CheckBox
            $checkBox.Content = $columnDefinition.Label
            $checkBox.Margin = '0,0,0,8'
            $checkBox.IsChecked = ($state.VisibleColumnKeys -contains $columnDefinition.Key)
            [void]$pnlColumns.Children.Add($checkBox)
            [void]$checkBoxStates.Add([pscustomobject]@{ Key = $columnDefinition.Key; CheckBox = $checkBox })
        }

        $btnDefaultColumns.Add_Click({
                foreach ($entry in @($checkBoxStates.ToArray())) {
                    $entry.CheckBox.IsChecked = ($entry.Key -in @('material_number', 'description'))
                }
            })
        $btnCancelColumns.Add_Click({
                $dialog.DialogResult = $false
                $dialog.Close()
            })
        $btnApplyColumns.Add_Click({
                $selectedKeys = New-Object System.Collections.Generic.List[string]
                foreach ($entry in @($checkBoxStates.ToArray())) {
                    if ([bool]$entry.CheckBox.IsChecked) {
                        [void]$selectedKeys.Add($entry.Key)
                    }
                }

                if ($selectedKeys.Count -eq 0) {
                    [System.Windows.MessageBox]::Show('Select at least one visible column.', 'Visible columns', 'OK', 'Warning') | Out-Null
                    return
                }

                $state.VisibleColumnKeys = @($selectedKeys.ToArray())
                & $ApplyListColumnVisibility
                $dialog.DialogResult = $true
                $dialog.Close()
            })

        return [bool]$dialog.ShowDialog()
    }

    $UpdateDirtyState = {
        if ($state.DatabaseDirty -or $state.EditorDirty) {
            $txtDirtyState.Text = 'Unsaved changes'
            $txtDirtyState.Foreground = '#B45309'
            $window.Title = 'Material Browser *'
        }
        else {
            $txtDirtyState.Text = 'Saved'
            $txtDirtyState.Foreground = '#0F766E'
            $window.Title = 'Material Browser'
        }

        & $UpdateEditorActions
    }

    $UpdateEditorActions = {
        $hasMaterial = ($null -ne $state.CurrentMaterial)
        $btnCloneMaterial.IsEnabled = $hasMaterial
        $btnDeleteMaterial.IsEnabled = $hasMaterial
        $btnRevertCurrent.IsEnabled = ($hasMaterial -and $state.EditorDirty)
    }

    $CreateCheckBoxes = {
        param(
            [Parameter(Mandatory = $true)][System.Windows.Controls.Panel]$Panel,
            [Parameter(Mandatory = $true)][object[]]$Entries
        )

        $Panel.Children.Clear()
        foreach ($entry in $Entries) {
            $checkBox = New-Object System.Windows.Controls.CheckBox
            $checkBox.Content = (Get-NormalizedString $entry.label)
            $checkBox.Tag = (Get-NormalizedString $entry.code)
            $checkBox.MinWidth = 140
            $checkBox.Foreground = '#1E293B'
            $checkBox.Add_Checked({
                    if (-not $state.PopulatingEditor) {
                        $state.EditorDirty = $true
                        & $UpdateDirtyState
                    }
                })
            $checkBox.Add_Unchecked({
                    if (-not $state.PopulatingEditor) {
                        $state.EditorDirty = $true
                        & $UpdateDirtyState
                    }
                })
            [void]$Panel.Children.Add($checkBox)
        }
    }

    $GetCheckedCodes = {
        param([Parameter(Mandatory = $true)][System.Windows.Controls.Panel]$Panel)

        $codes = New-Object System.Collections.Generic.List[string]
        foreach ($child in $Panel.Children) {
            if ($child -is [System.Windows.Controls.CheckBox] -and $child.IsChecked) {
                [void]$codes.Add((Get-NormalizedString $child.Tag))
            }
        }

        return @(ConvertTo-UniqueStringArray $codes)
    }

    $SetCheckedCodes = {
        param(
            [Parameter(Mandatory = $true)][System.Windows.Controls.Panel]$Panel,
            [AllowNull()][string[]]$Codes
        )

        $selected = @{}
        foreach ($code in @(ConvertTo-UniqueStringArray $Codes)) {
            $selected[$code] = $true
        }

        foreach ($child in $Panel.Children) {
            if ($child -is [System.Windows.Controls.CheckBox]) {
                $child.IsChecked = $selected.ContainsKey((Get-NormalizedString $child.Tag))
            }
        }
    }

    & $CreateCheckBoxes -Panel $pnlHazmatFlags -Entries @($lookupData.hazmat_flags)
    & $CreateCheckBoxes -Panel $pnlResponsibilityCodes -Entries @($lookupData.responsibility_codes)
    & $CreateCheckBoxes -Panel $pnlAssignmentTags -Entries @($lookupData.assignment_tags)

    $BuildEditorCandidate = {
        $gridAlternateUnits.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true) | Out-Null
        $gridAlternateUnits.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true) | Out-Null
        $gridAlternates.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true) | Out-Null
        $gridAlternates.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true) | Out-Null

        $errors = New-Object System.Collections.Generic.List[string]

        $idResult = ConvertTo-IntParseResult $txtId.Text
        if (-not $idResult.Success) {
            [void]$errors.Add('ID must be a whole number.')
        }

        $quantityResult = ConvertTo-NumberParseResult $txtQuantityTarget.Text
        if (-not $quantityResult.Success) {
            [void]$errors.Add('Target quantity is not a valid number.')
        }

        $alternateUnits = New-Object System.Collections.Generic.List[object]
        $rowIndex = 0
        foreach ($row in @(ConvertTo-ObjectArray $gridAlternateUnits.ItemsSource)) {
            $rowIndex++
            $conversionResult = ConvertTo-NumberParseResult $row.conversion_to_base
            if (-not $conversionResult.Success) {
                [void]$errors.Add("Alternate unit row $rowIndex has an invalid conversion.")
            }

            [void]$alternateUnits.Add([pscustomobject][ordered]@{
                    unit_code          = ConvertTo-NullableString $row.unit_code
                    conversion_to_base = [double]$conversionResult.Value
                })
        }

        $alternates = New-Object System.Collections.Generic.List[object]
        $alternateIndex = 0
        foreach ($row in @(ConvertTo-ObjectArray $gridAlternates.ItemsSource)) {
            $alternateIndex++
            $positionResult = ConvertTo-IntParseResult $row.position
            if (-not $positionResult.Success) {
                [void]$errors.Add("Alternate row $alternateIndex has an invalid position.")
            }

            [void]$alternates.Add([pscustomobject][ordered]@{
                    position             = [int]$positionResult.Value
                    identifier           = [pscustomobject][ordered]@{
                        type  = 'matnr'
                        value = Get-NormalizedString $row.identifier_value
                    }
                    material_status_code = Get-NormalizedString $row.material_status_code
                    preferred_unit_code  = ConvertTo-NullableString $row.preferred_unit_code
                })
        }

        $materialNumber = Get-NormalizedString $txtMaterialNumber.Text
        $canonicalKey = if ([string]::IsNullOrWhiteSpace($materialNumber)) { '' } else { Get-CanonicalKey -Type 'matnr' -Value $materialNumber }

        $candidate = [pscustomobject][ordered]@{
            id                 = [int]$idResult.Value
            canonical_key      = $canonicalKey
            primary_identifier = [pscustomobject][ordered]@{
                type  = 'matnr'
                value = $materialNumber
            }
            identifiers        = [pscustomobject][ordered]@{
                matnr             = ConvertTo-NullableString $materialNumber
                supply_number     = ConvertTo-NullableString $txtSupplyNumber.Text
                article_number    = ConvertTo-NullableString $txtArticleNumber.Text
                nato_stock_number = ConvertTo-NullableString $txtNatoStockNumber.Text
            }
            status             = [pscustomobject][ordered]@{
                material_status_code = $(if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $txtMaterialStatus.Text))) { 'XX' } else { Get-NormalizedString $txtMaterialStatus.Text })
            }
            texts              = [pscustomobject][ordered]@{
                short_description = Get-NormalizedString $txtShortDescription.Text
                technical_note    = Get-NormalizedString $txtTechnicalNote.Text
                logistics_note    = Get-NormalizedString $txtLogisticsNote.Text
            }
            classification     = [pscustomobject][ordered]@{
                ext_wg       = Get-NormalizedString $txtExtWg.Text
                is_decentral = [bool]$chkIsDecentral.IsChecked
                creditor     = ConvertTo-NullableString $txtCreditor.Text
            }
            hazmat             = [pscustomobject][ordered]@{
                is_hazardous = [bool]$chkIsHazardous.IsChecked
                un_number    = ConvertTo-NullableString $txtUnNumber.Text
                flags        = @(& $GetCheckedCodes -Panel $pnlHazmatFlags)
            }
            quantity           = [pscustomobject][ordered]@{
                base_unit       = $(if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $cmbBaseUnit.SelectedValue))) { $defaultUnitCode } else { Get-NormalizedString $cmbBaseUnit.SelectedValue })
                target          = [double]$quantityResult.Value
                alternate_units = $alternateUnits.ToArray()
            }
            alternates         = $alternates.ToArray()
            assignments        = [pscustomobject][ordered]@{
                responsibility_codes = @(& $GetCheckedCodes -Panel $pnlResponsibilityCodes)
                assignment_tags      = @(& $GetCheckedCodes -Panel $pnlAssignmentTags)
            }
        }

        return [pscustomobject]@{
            Candidate = $candidate
            Errors    = @($errors)
        }
    }

    $ValidateCandidate = {
        param(
            [Parameter(Mandatory = $true)]$Candidate,
            [AllowNull()]$CurrentMaterial
        )

        $messages = New-Object System.Collections.Generic.List[string]
        $validUnitCodes = @($lookupData.unit_codes | ForEach-Object { Get-NormalizedString $_.code })
        $validHazmatFlags = @($lookupData.hazmat_flags | ForEach-Object { Get-NormalizedString $_.code })
        $validResponsibilityCodes = @($lookupData.responsibility_codes | ForEach-Object { Get-NormalizedString $_.code })
        $validAssignmentTags = @($lookupData.assignment_tags | ForEach-Object { Get-NormalizedString $_.code })

        if ([int]$Candidate.id -le 0) {
            [void]$messages.Add('ID must be greater than 0.')
        }

        if ((Get-NormalizedString $Candidate.primary_identifier.type) -ne 'matnr') {
            [void]$messages.Add('Primary identifier type must be matnr.')
        }

        if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $Candidate.primary_identifier.value))) {
            [void]$messages.Add('Materialnummer is required.')
        }

        $canonicalKey = Get-NormalizedString $Candidate.canonical_key
        if ([string]::IsNullOrWhiteSpace($canonicalKey)) {
            [void]$messages.Add('Canonical key could not be generated.')
        }

        foreach ($material in @($state.Materials.ToArray())) {
            if ($null -ne $CurrentMaterial -and [object]::ReferenceEquals($material, $CurrentMaterial)) {
                continue
            }

            if ([int]$material.id -eq [int]$Candidate.id) {
                [void]$messages.Add("ID $($Candidate.id) already exists.")
                break
            }
        }

        foreach ($material in @($state.Materials.ToArray())) {
            if ($null -ne $CurrentMaterial -and [object]::ReferenceEquals($material, $CurrentMaterial)) {
                continue
            }

            if ((Get-NormalizedString $material.canonical_key) -eq $canonicalKey) {
                [void]$messages.Add("Canonical key '$canonicalKey' already exists.")
                break
            }
        }

        if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $Candidate.quantity.base_unit)) -or -not ($validUnitCodes -contains (Get-NormalizedString $Candidate.quantity.base_unit))) {
            [void]$messages.Add('Base unit is required and must exist in the lookup.')
        }

        if ([double]$Candidate.quantity.target -lt 0) {
            [void]$messages.Add('Target quantity must be 0 or greater.')
        }

        $seenAltUnits = @{}
        $altUnitIndex = 0
        foreach ($alternateUnit in @(ConvertTo-ObjectArray $Candidate.quantity.alternate_units)) {
            $altUnitIndex++
            $unitCode = Get-NormalizedString $alternateUnit.unit_code
            if ([string]::IsNullOrWhiteSpace($unitCode)) {
                [void]$messages.Add("Alternate unit row $altUnitIndex requires a unit code.")
            }
            elseif (-not ($validUnitCodes -contains $unitCode)) {
                [void]$messages.Add("Alternate unit '$unitCode' is not in the lookup.")
            }
            elseif ($seenAltUnits.ContainsKey($unitCode)) {
                [void]$messages.Add("Alternate unit '$unitCode' is duplicated.")
            }
            else {
                $seenAltUnits[$unitCode] = $true
            }

            if ([double]$alternateUnit.conversion_to_base -le 0) {
                [void]$messages.Add("Alternate unit '$unitCode' must have a conversion greater than 0.")
            }
        }

        $seenPositions = @{}
        $alternateRowIndex = 0
        foreach ($alternate in @(ConvertTo-ObjectArray $Candidate.alternates)) {
            $alternateRowIndex++
            $position = [int]$alternate.position
            if ($position -le 0) {
                [void]$messages.Add("Alternate row $alternateRowIndex must have a positive position.")
            }
            elseif ($seenPositions.ContainsKey($position)) {
                [void]$messages.Add("Alternate position $position is duplicated.")
            }
            else {
                $seenPositions[$position] = $true
            }

            if ((Get-NormalizedString $alternate.identifier.type) -ne 'matnr') {
                [void]$messages.Add("Alternate row $alternateRowIndex must use matnr identifiers.")
            }

            if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $alternate.identifier.value))) {
                [void]$messages.Add("Alternate row $alternateRowIndex requires an identifier value.")
            }

            $preferredUnit = Get-NormalizedString $alternate.preferred_unit_code
            if (-not [string]::IsNullOrWhiteSpace($preferredUnit) -and -not ($validUnitCodes -contains $preferredUnit)) {
                [void]$messages.Add("Alternate row $alternateRowIndex has an invalid preferred unit.")
            }
        }

        foreach ($flag in @(ConvertTo-ObjectArray $Candidate.hazmat.flags)) {
            if (-not ($validHazmatFlags -contains (Get-NormalizedString $flag))) {
                [void]$messages.Add("Hazmat flag '$flag' is not in the lookup.")
            }
        }

        foreach ($code in @(ConvertTo-ObjectArray $Candidate.assignments.responsibility_codes)) {
            if (-not ($validResponsibilityCodes -contains (Get-NormalizedString $code))) {
                [void]$messages.Add("Responsibility code '$code' is not in the lookup.")
            }
        }

        foreach ($code in @(ConvertTo-ObjectArray $Candidate.assignments.assignment_tags)) {
            if (-not ($validAssignmentTags -contains (Get-NormalizedString $code))) {
                [void]$messages.Add("Assignment tag '$code' is not in the lookup.")
            }
        }

        return [pscustomobject]@{
            IsValid      = ($messages.Count -eq 0)
            Messages     = @($messages)
            CanonicalKey = $canonicalKey
        }
    }

    $PopulateEditor = {
        param([AllowNull()]$Material)

        $state.PopulatingEditor = $true
        try {
            if ($null -eq $Material) {
                $txtEditorHeadline.Text = 'No material selected'
                $txtEditorSubheadline.Text = 'Select a material or create a new one.'
                $txtId.Text = ''
                $txtMaterialNumber.Text = ''
                $txtSupplyNumber.Text = ''
                $txtArticleNumber.Text = ''
                $txtNatoStockNumber.Text = ''
                $txtMaterialStatus.Text = 'XX'
                $txtShortDescription.Text = ''
                $txtTechnicalNote.Text = ''
                $txtLogisticsNote.Text = ''
                $txtExtWg.Text = ''
                $chkIsDecentral.IsChecked = $false
                $txtCreditor.Text = ''
                $chkIsHazardous.IsChecked = $false
                $txtUnNumber.Text = ''
                $cmbBaseUnit.SelectedValue = $defaultUnitCode
                $txtQuantityTarget.Text = '0'
                $gridAlternateUnits.ItemsSource = New-ObservableCollection
                $gridAlternates.ItemsSource = New-ObservableCollection
                & $SetCheckedCodes -Panel $pnlHazmatFlags -Codes @()
                & $SetCheckedCodes -Panel $pnlResponsibilityCodes -Codes @()
                & $SetCheckedCodes -Panel $pnlAssignmentTags -Codes @()
                $state.CurrentMaterial = $null
                $state.CurrentSummary = $null
            }
            else {
                $txtEditorHeadline.Text = "Material #$($Material.id)"
                $alternateCount = (@(ConvertTo-ObjectArray $Material.alternates)).Count
                $txtEditorSubheadline.Text = "$(Get-NormalizedString $Material.identifiers.matnr)  |  $(Get-NormalizedString $Material.texts.short_description)  |  $alternateCount alternates"
                $txtId.Text = [string]$Material.id
                $txtMaterialNumber.Text = Get-NormalizedString $Material.identifiers.matnr
                $txtSupplyNumber.Text = Get-NormalizedString $Material.identifiers.supply_number
                $txtArticleNumber.Text = Get-NormalizedString $Material.identifiers.article_number
                $txtNatoStockNumber.Text = Get-NormalizedString $Material.identifiers.nato_stock_number
                $txtMaterialStatus.Text = Get-NormalizedString $Material.status.material_status_code
                $txtShortDescription.Text = Get-NormalizedString $Material.texts.short_description
                $txtTechnicalNote.Text = Get-NormalizedString $Material.texts.technical_note
                $txtLogisticsNote.Text = Get-NormalizedString $Material.texts.logistics_note
                $txtExtWg.Text = Get-NormalizedString $Material.classification.ext_wg
                $chkIsDecentral.IsChecked = [bool]$Material.classification.is_decentral
                $txtCreditor.Text = Get-NormalizedString $Material.classification.creditor
                $chkIsHazardous.IsChecked = [bool]$Material.hazmat.is_hazardous
                $txtUnNumber.Text = Get-NormalizedString $Material.hazmat.un_number
                $cmbBaseUnit.SelectedValue = Get-NormalizedString $Material.quantity.base_unit
                $txtQuantityTarget.Text = [string]([double]$Material.quantity.target)

                $alternateUnitRows = New-Object System.Collections.Generic.List[object]
                foreach ($alternateUnit in @(ConvertTo-ObjectArray $Material.quantity.alternate_units)) {
                    [void]$alternateUnitRows.Add([pscustomobject]@{
                            unit_code          = Get-NormalizedString $alternateUnit.unit_code
                            conversion_to_base = [double]$alternateUnit.conversion_to_base
                        })
                }
                $gridAlternateUnits.ItemsSource = New-ObservableCollection -Items $alternateUnitRows.ToArray()

                $alternateRows = New-Object System.Collections.Generic.List[object]
                foreach ($alternate in @(ConvertTo-ObjectArray $Material.alternates)) {
                    [void]$alternateRows.Add([pscustomobject]@{
                            position             = [int]$alternate.position
                            identifier_value     = Get-NormalizedString $alternate.identifier.value
                            material_status_code = Get-NormalizedString $alternate.material_status_code
                            preferred_unit_code  = Get-NormalizedString $alternate.preferred_unit_code
                        })
                }
                $gridAlternates.ItemsSource = New-ObservableCollection -Items $alternateRows.ToArray()

                & $SetCheckedCodes -Panel $pnlHazmatFlags -Codes @($Material.hazmat.flags)
                & $SetCheckedCodes -Panel $pnlResponsibilityCodes -Codes @($Material.assignments.responsibility_codes)
                & $SetCheckedCodes -Panel $pnlAssignmentTags -Codes @($Material.assignments.assignment_tags)

                $state.CurrentMaterial = $Material
            }
        }
        finally {
            $state.PopulatingEditor = $false
            $state.EditorDirty = $false
            & $UpdateDirtyState
        }
    }

    $RefreshList = {
        param([AllowNull()]$PreferredMaterial)

        $summaries = @(Get-FilteredMaterialSummaries `
            -Materials $state.Materials.ToArray() `
            -SearchText $txtSearch.Text `
            -HazardousOnly:$state.FilterHazardousOnly `
            -DecentralOnly:$state.FilterDecentralOnly `
            -AdvancedRules $state.ActiveFilterRules `
            -HazmatFlagLabelMap $hazmatFlagLabelMap `
            -ResponsibilityCodeLabelMap $responsibilityCodeLabelMap `
            -AssignmentTagLabelMap $assignmentTagLabelMap)
        $dgMaterials.ItemsSource = New-ObservableCollection -Items $summaries

        $filteredCount = @($summaries).Count
        $totalCount = $state.Materials.Count
        $hazardousCount = (@($state.Materials.ToArray() | Where-Object { $_.hazmat.is_hazardous })).Count
        $decentralCount = (@($state.Materials.ToArray() | Where-Object { $_.classification.is_decentral })).Count
        $activeFilters = New-Object System.Collections.Generic.List[string]
        if ($state.FilterHazardousOnly) { [void]$activeFilters.Add('Gefahrgut') }
        if ($state.FilterDecentralOnly) { [void]$activeFilters.Add('Dezentral') }
        foreach ($rule in @($state.ActiveFilterRules)) {
            [void]$activeFilters.Add((& $GetFilterRuleLabel -Rule $rule))
        }
        $filterLabel = if ($activeFilters.Count -gt 0) { $activeFilters -join ', ' } else { 'none' }
        $txtListMeta.Text = "$filteredCount visible / $totalCount total  |  $hazardousCount hazardous  |  $decentralCount decentral  |  filters: $filterLabel"
        if ($null -ne $PreferredMaterial) {
            $match = $summaries | Where-Object { [object]::ReferenceEquals($_.MaterialRef, $PreferredMaterial) } | Select-Object -First 1
            $state.SuppressSelectionChange = $true
            $dgMaterials.SelectedItem = $match
            $state.SuppressSelectionChange = $false
            if ($null -ne $match) {
                $dgMaterials.ScrollIntoView($match)
                $state.CurrentSummary = $match
            }
        }
    }

    $CommitCurrentEditor = {
        param([switch]$ShowErrors)

        if ($null -eq $state.CurrentMaterial -or -not $state.EditorDirty) {
            return $true
        }

        $buildResult = & $BuildEditorCandidate
        $validationResult = & $ValidateCandidate -Candidate $buildResult.Candidate -CurrentMaterial $state.CurrentMaterial

        $allErrors = New-Object System.Collections.Generic.List[string]
        foreach ($validationMessage in @($buildResult.Errors + $validationResult.Messages)) {
            if (-not [string]::IsNullOrWhiteSpace((Get-NormalizedString $validationMessage))) {
                [void]$allErrors.Add((Get-NormalizedString $validationMessage))
            }
        }

        if ($allErrors.Count -gt 0) {
            $message = ($allErrors | Select-Object -Unique) -join [Environment]::NewLine
            & $SetStatus -Message 'Validation failed. Fix the current material before continuing.' -Level 'Error'
            if ($ShowErrors) {
                [System.Windows.MessageBox]::Show($message, 'Validation error', 'OK', 'Warning') | Out-Null
            }
            return $false
        }

        $candidate = $buildResult.Candidate
        $candidate.canonical_key = $validationResult.CanonicalKey
        $currentJson = ConvertTo-JsonString -InputObject $state.CurrentMaterial -Depth 20 -Compress
        $candidateJson = ConvertTo-JsonString -InputObject $candidate -Depth 20 -Compress
        if ($currentJson -ne $candidateJson) {
            $state.CurrentMaterial.id = $candidate.id
            $state.CurrentMaterial.canonical_key = $candidate.canonical_key
            $state.CurrentMaterial.primary_identifier = $candidate.primary_identifier
            $state.CurrentMaterial.identifiers = $candidate.identifiers
            $state.CurrentMaterial.status = $candidate.status
            $state.CurrentMaterial.texts = $candidate.texts
            $state.CurrentMaterial.classification = $candidate.classification
            $state.CurrentMaterial.hazmat = $candidate.hazmat
            $state.CurrentMaterial.quantity = $candidate.quantity
            $state.CurrentMaterial.alternates = $candidate.alternates
            $state.CurrentMaterial.assignments = $candidate.assignments
            $state.DatabaseDirty = $true
        }

        $state.EditorDirty = $false
        & $UpdateDirtyState
        & $RefreshList -PreferredMaterial $state.CurrentMaterial
        & $PopulateEditor -Material $state.CurrentMaterial
        & $SetStatus -Message 'Material staged in memory.' -Level 'Success'
        return $true
    }

    $LoadDatabase = {
        $database = Read-DatabaseFile -Path $Script:DbPath -DefaultIdentifierType $defaultIdentifierType -DefaultUnitCode $defaultUnitCode
        $state.Materials.Clear()
        foreach ($material in @($database.materials)) {
            [void]$state.Materials.Add($material)
        }
        $state.DatabaseDirty = $false
        $state.EditorDirty = $false
        & $UpdateDirtyState
        & $RefreshList -PreferredMaterial $null
        if ($state.Materials.Count -gt 0) {
            & $PopulateEditor -Material $state.Materials[0]
            & $RefreshList -PreferredMaterial $state.Materials[0]
        }
        else {
            & $PopulateEditor -Material $null
        }
        & $SetStatus -Message "Loaded $($state.Materials.Count) materials." -Level 'Success'
    }

    $SaveDatabase = {
        if (-not (& $CommitCurrentEditor -ShowErrors)) {
            return $false
        }

        try {
            $backupPath = Backup-DatabaseFile -Path $Script:DbPath
            Save-DatabaseFile -Path $Script:DbPath -Materials $state.Materials.ToArray()
            $state.DatabaseDirty = $false
            $state.EditorDirty = $false
            & $UpdateDirtyState
            & $SetStatus -Message "Database saved. Backup: $(Split-Path $backupPath -Leaf)" -Level 'Success'
            return $true
        }
        catch {
            & $SetStatus -Message "Save failed: $($_.Exception.Message)" -Level 'Error'
            [System.Windows.MessageBox]::Show($_.Exception.Message, 'Save error', 'OK', 'Error') | Out-Null
            return $false
        }
    }

    $ConfirmClose = {
        if (-not ($state.DatabaseDirty -or $state.EditorDirty)) {
            return $true
        }

        $result = [System.Windows.MessageBox]::Show('There are unsaved changes. Save before closing?', 'Unsaved changes', 'YesNoCancel', 'Warning')
        switch ($result) {
            'Yes' { return (& $SaveDatabase) }
            'No' { return $true }
            default { return $false }
        }
    }

    $MarkEditorDirty = {
        if ($state.PopulatingEditor) {
            return
        }

        $state.EditorDirty = $true
        & $UpdateDirtyState
    }

    $txtSearchTimer = New-Object System.Windows.Threading.DispatcherTimer
    $txtSearchTimer.Interval = [TimeSpan]::FromMilliseconds(180)
    $txtSearchTimer.Add_Tick({
            $txtSearchTimer.Stop()
            & $RefreshList -PreferredMaterial $state.CurrentMaterial
            & $SetStatus -Message 'Search updated.' -Level 'Info'
        })

    foreach ($control in @($txtId, $txtMaterialNumber, $txtSupplyNumber, $txtArticleNumber, $txtNatoStockNumber, $txtMaterialStatus, $txtShortDescription, $txtTechnicalNote, $txtLogisticsNote, $txtExtWg, $txtCreditor, $txtUnNumber, $txtQuantityTarget)) {
        $control.Add_TextChanged({ & $MarkEditorDirty })
    }

    $cmbBaseUnit.Add_SelectionChanged({ & $MarkEditorDirty })
    $chkIsDecentral.Add_Checked({ & $MarkEditorDirty })
    $chkIsDecentral.Add_Unchecked({ & $MarkEditorDirty })
    $chkIsHazardous.Add_Checked({ & $MarkEditorDirty })
    $chkIsHazardous.Add_Unchecked({ & $MarkEditorDirty })
    $gridAlternateUnits.Add_CellEditEnding({ if (-not $state.PopulatingEditor) { $state.EditorDirty = $true; & $UpdateDirtyState } })
    $gridAlternates.Add_CellEditEnding({ if (-not $state.PopulatingEditor) { $state.EditorDirty = $true; & $UpdateDirtyState } })

    $txtSearch.Add_TextChanged({
            $txtSearchTimer.Stop()
            $txtSearchTimer.Start()
        })

    $btnClearSearch.Add_Click({
            $txtSearch.Text = ''
            & $RefreshList -PreferredMaterial $state.CurrentMaterial
        })

    $btnOpenFilterMenu.Add_Click({
            if (& $OpenFilterDialog) {
                & $RefreshList -PreferredMaterial $state.CurrentMaterial
                & $SetStatus -Message 'Advanced filters updated.' -Level 'Info'
            }
        })

    $btnOpenColumnMenu.Add_Click({
            if (& $OpenColumnDialog) {
                & $SetStatus -Message 'Visible columns updated.' -Level 'Info'
            }
        })

    & $ApplyListColumnVisibility

    foreach ($filterControl in @($chkFilterHazardous, $chkFilterDecentral)) {
        $filterControl.Add_Checked({
                $state.FilterHazardousOnly = [bool]$chkFilterHazardous.IsChecked
                $state.FilterDecentralOnly = [bool]$chkFilterDecentral.IsChecked
                & $RefreshList -PreferredMaterial $state.CurrentMaterial
                & $SetStatus -Message 'List filters updated.' -Level 'Info'
            })
        $filterControl.Add_Unchecked({
                $state.FilterHazardousOnly = [bool]$chkFilterHazardous.IsChecked
                $state.FilterDecentralOnly = [bool]$chkFilterDecentral.IsChecked
                & $RefreshList -PreferredMaterial $state.CurrentMaterial
                & $SetStatus -Message 'List filters updated.' -Level 'Info'
            })
    }

    $dgMaterials.Add_SelectionChanged({
            if ($state.SuppressSelectionChange) {
                return
            }

            $selectedSummary = $dgMaterials.SelectedItem
            if ($null -eq $selectedSummary) {
                return
            }

            if ($null -ne $state.CurrentMaterial -and -not (& $CommitCurrentEditor -ShowErrors)) {
                $state.SuppressSelectionChange = $true
                $dgMaterials.SelectedItem = $state.CurrentSummary
                $state.SuppressSelectionChange = $false
                return
            }

            $state.CurrentSummary = $selectedSummary
            & $PopulateEditor -Material $selectedSummary.MaterialRef
            & $SetStatus -Message "Selected material $($selectedSummary.Id)." -Level 'Info'
        })

    $btnNewMaterial.Add_Click({
            if ($null -ne $state.CurrentMaterial -and -not (& $CommitCurrentEditor -ShowErrors)) {
                return
            }

            $newMaterial = New-DefaultMaterial -Id (Get-NextMaterialId -Materials $state.Materials.ToArray()) -DefaultIdentifierType $defaultIdentifierType -DefaultUnitCode $defaultUnitCode
            $newMaterial.canonical_key = ''
            [void]$state.Materials.Add($newMaterial)
            $state.DatabaseDirty = $true
            $txtSearch.Text = ''
            & $RefreshList -PreferredMaterial $newMaterial
            & $PopulateEditor -Material $newMaterial
            $state.EditorDirty = $true
            & $UpdateDirtyState
            & $SetStatus -Message 'New material created in memory.' -Level 'Success'
        })

        $btnCloneMaterial.Add_Click({
            if ($null -eq $state.CurrentMaterial) {
                return
            }

            if (-not (& $CommitCurrentEditor -ShowErrors)) {
                return
            }

            $sourceId = [int]$state.CurrentMaterial.id
            $cloneId = Get-NextMaterialId -Materials $state.Materials.ToArray()
            $clone = ConvertTo-NormalizedMaterial -Material (Copy-DeepObject $state.CurrentMaterial) -DefaultIdentifierType $defaultIdentifierType -DefaultUnitCode $defaultUnitCode
            $clone.id = $cloneId
            $clone.primary_identifier.type = 'matnr'
            $clone.primary_identifier.value = Get-UniqueCloneIdentifierValue -BaseValue $clone.primary_identifier.value -IdentifierType 'matnr' -Materials $state.Materials.ToArray() -SuggestedId $cloneId
            $clone.identifiers.matnr = $clone.primary_identifier.value
            $clone.canonical_key = Get-CanonicalKey -Type 'matnr' -Value $clone.primary_identifier.value
            $clone.texts.short_description = if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $clone.texts.short_description))) { 'Copy' } else { "$(Get-NormalizedString $clone.texts.short_description) (Copy)" }
            [void]$state.Materials.Add($clone)
            $state.DatabaseDirty = $true
            $txtSearch.Text = ''
            & $RefreshList -PreferredMaterial $clone
            & $PopulateEditor -Material $clone
            $state.EditorDirty = $true
            & $UpdateDirtyState
            & $SetStatus -Message "Material #$sourceId duplicated as #$cloneId." -Level 'Success'
        })

    $btnRevertCurrent.Add_Click({
            if ($null -eq $state.CurrentMaterial -or -not $state.EditorDirty) {
                return
            }

            $result = [System.Windows.MessageBox]::Show('Discard uncommitted editor changes for the current material?', 'Discard changes', 'YesNo', 'Warning')
            if ($result -ne 'Yes') {
                return
            }

            & $PopulateEditor -Material $state.CurrentMaterial
            & $SetStatus -Message 'Editor changes reverted to the last staged version.' -Level 'Warning'
        })

    $btnDeleteMaterial.Add_Click({
            if ($null -eq $state.CurrentMaterial) {
                return
            }

            $result = [System.Windows.MessageBox]::Show("Delete material #$($state.CurrentMaterial.id)? This removes it from the JSON on next save.", 'Delete material', 'YesNo', 'Warning')
            if ($result -ne 'Yes') {
                return
            }

            $currentIndex = $state.Materials.IndexOf($state.CurrentMaterial)
            [void]$state.Materials.Remove($state.CurrentMaterial)
            $state.DatabaseDirty = $true
            $state.EditorDirty = $false
            $nextMaterial = $null
            if ($state.Materials.Count -gt 0) {
                if ($currentIndex -ge $state.Materials.Count) {
                    $currentIndex = $state.Materials.Count - 1
                }
                $nextMaterial = $state.Materials[$currentIndex]
            }

            & $RefreshList -PreferredMaterial $nextMaterial
            & $PopulateEditor -Material $nextMaterial
            & $UpdateDirtyState
            & $SetStatus -Message 'Material deleted in memory.' -Level 'Warning'
        })

    $btnReloadDatabase.Add_Click({
            if ($state.DatabaseDirty -or $state.EditorDirty) {
                $result = [System.Windows.MessageBox]::Show('Discard current unsaved changes and reload the database from disk?', 'Reload database', 'YesNo', 'Warning')
                if ($result -ne 'Yes') {
                    return
                }
            }

            try {
                & $LoadDatabase
            }
            catch {
                & $SetStatus -Message "Reload failed: $($_.Exception.Message)" -Level 'Error'
                [System.Windows.MessageBox]::Show($_.Exception.Message, 'Load error', 'OK', 'Error') | Out-Null
            }
        })

    $btnSaveDatabase.Add_Click({ [void](& $SaveDatabase) })

    $btnAddAlternateUnit.Add_Click({
            if ($null -eq $gridAlternateUnits.ItemsSource) {
                $gridAlternateUnits.ItemsSource = New-ObservableCollection
            }

            [void]$gridAlternateUnits.ItemsSource.Add([pscustomobject]@{ unit_code = $defaultUnitCode; conversion_to_base = 1.0 })
            $state.EditorDirty = $true
            & $UpdateDirtyState
        })

    $btnRemoveAlternateUnit.Add_Click({
            if ($null -eq $gridAlternateUnits.SelectedItem) {
                return
            }

            [void]$gridAlternateUnits.ItemsSource.Remove($gridAlternateUnits.SelectedItem)
            $state.EditorDirty = $true
            & $UpdateDirtyState
        })

        $btnAddAlternate.Add_Click({
            if ($null -eq $gridAlternates.ItemsSource) {
                $gridAlternates.ItemsSource = New-ObservableCollection
            }

            $nextPosition = (@(ConvertTo-ObjectArray $gridAlternates.ItemsSource) | Measure-Object).Count + 1
            [void]$gridAlternates.ItemsSource.Add([pscustomobject]@{ position = $nextPosition; identifier_value = ''; material_status_code = ''; preferred_unit_code = $defaultUnitCode })
            $state.EditorDirty = $true
            & $UpdateDirtyState
        })

    $btnRemoveAlternate.Add_Click({
            if ($null -eq $gridAlternates.SelectedItem) {
                return
            }

            [void]$gridAlternates.ItemsSource.Remove($gridAlternates.SelectedItem)
            $state.EditorDirty = $true
            & $UpdateDirtyState
        })

    $window.Add_PreviewKeyDown({
            param($windowSender, $windowEventArgs)

            if (($windowEventArgs.KeyboardDevice.Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -and $windowEventArgs.Key -eq [System.Windows.Input.Key]::S) {
                $windowEventArgs.Handled = $true
                [void](& $SaveDatabase)
                return
            }

            if (($windowEventArgs.KeyboardDevice.Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -and $windowEventArgs.Key -eq [System.Windows.Input.Key]::N) {
                $windowEventArgs.Handled = $true
                $btnNewMaterial.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
                return
            }

            if (($windowEventArgs.KeyboardDevice.Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -and $windowEventArgs.Key -eq [System.Windows.Input.Key]::D) {
                $windowEventArgs.Handled = $true
                $btnCloneMaterial.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
                return
            }

            if ($windowEventArgs.Key -eq [System.Windows.Input.Key]::Escape -and $state.EditorDirty) {
                $windowEventArgs.Handled = $true
                $btnRevertCurrent.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
                return
            }

            if ($windowEventArgs.Key -eq [System.Windows.Input.Key]::Delete -and $dgMaterials.IsKeyboardFocusWithin) {
                $windowEventArgs.Handled = $true
                $btnDeleteMaterial.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
            }
        })

    $window.Add_Closing({
            param($windowSender, $windowEventArgs)

            if (-not (& $ConfirmClose)) {
                $windowEventArgs.Cancel = $true
            }
        })

    & $SetStatus -Message 'Loading database...' -Level 'Info'
    try {
        & $LoadDatabase
    }
    catch {
        & $SetStatus -Message "Failed to load database: $($_.Exception.Message)" -Level 'Error'
        [System.Windows.MessageBox]::Show($_.Exception.Message, 'Load error', 'OK', 'Error') | Out-Null
    }

    $window.ShowDialog() | Out-Null
}

if ($MyInvocation.InvocationName -ne '.') {
    Start-MaterialBrowserUi
}
