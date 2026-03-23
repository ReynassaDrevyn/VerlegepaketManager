# GridViewer.ps1
# Standalone WPF grid editor for the material database.

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

    return [pscustomobject][ordered]@{
        id                 = $idValue
        canonical_key      = $canonicalKey
        primary_identifier = [pscustomobject][ordered]@{
            type  = $primaryType
            value = $primaryValue
        }
        identifiers        = [pscustomobject][ordered]@{
            matnr             = $identifierMatnr
            supply_number     = $identifierSupply
            article_number    = $identifierArticle
            nato_stock_number = $identifierNato
        }
        status             = [pscustomobject][ordered]@{
            material_status_code = $statusCode
        }
        texts              = [pscustomobject][ordered]@{
            short_description = $textShort
            technical_note    = $textTechnical
            logistics_note    = $textLogistics
        }
        classification     = [pscustomobject][ordered]@{
            ext_wg       = $classificationExtWg
            is_decentral = $classificationIsDecentral
            creditor     = $classificationCreditor
        }
        hazmat             = [pscustomobject][ordered]@{
            is_hazardous = $hazmatIsHazardous
            un_number    = $hazmatUnNumber
            flags        = @($hazmatFlags)
        }
        quantity           = [pscustomobject][ordered]@{
            base_unit       = $resolvedBaseUnit
            target          = $quantityTarget
            alternate_units = $alternateUnits.ToArray()
        }
        alternates         = $alternates.ToArray()
        assignments        = [pscustomobject][ordered]@{
            responsibility_codes = @($responsibilityCodes)
            assignment_tags      = @($assignmentTags)
        }
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

function Get-GridColumnDefinitions {
    $definitions = New-Object System.Collections.Generic.List[object]

    $orderedDefinitions = @(
        [pscustomobject]@{ Key = 'import_id'; Label = 'ID'; PropertyName = 'ImportId'; Type = 'number'; Width = 90; DefaultVisible = $true; Editable = $true }
        [pscustomobject]@{ Key = 'supplynumber'; Label = 'VersNr'; PropertyName = 'SupplyNumber'; Type = 'text'; Width = 150; DefaultVisible = $true; Editable = $true }
        [pscustomobject]@{ Key = 'matnr_main'; Label = 'Materialnummer SASPF'; PropertyName = 'MaterialNumber'; Type = 'text'; Width = 190; DefaultVisible = $true; Editable = $true }
        [pscustomobject]@{ Key = 'mat_stat_main'; Label = 'Status MatNr'; PropertyName = 'MaterialStatus'; Type = 'text'; Width = 110; DefaultVisible = $true; Editable = $true }
        [pscustomobject]@{ Key = 'nato_stock_number'; Label = 'NATO Stock Number'; PropertyName = 'NatoStockNumber'; Type = 'text'; Width = 150; DefaultVisible = $false; Editable = $true }
        [pscustomobject]@{ Key = 'dezentral'; Label = 'Dezent'; PropertyName = 'IsDecentral'; Type = 'bool'; Width = 80; DefaultVisible = $true; Editable = $true }
        [pscustomobject]@{ Key = 'ext_wg'; Label = 'Ext WG'; PropertyName = 'ExtWg'; Type = 'text'; Width = 100; DefaultVisible = $true; Editable = $true }
        [pscustomobject]@{ Key = 'creditor'; Label = 'Creditor'; PropertyName = 'Creditor'; Type = 'text'; Width = 110; DefaultVisible = $false; Editable = $true }
        [pscustomobject]@{ Key = 'artnr'; Label = 'Artikel Nr'; PropertyName = 'ArticleNumber'; Type = 'text'; Width = 170; DefaultVisible = $true; Editable = $true }
        [pscustomobject]@{ Key = 'description'; Label = 'Materialbezeichnung'; PropertyName = 'Description'; Type = 'text'; Width = 240; DefaultVisible = $true; Editable = $true }
        [pscustomobject]@{ Key = 'technical'; Label = 'Bezeichnung Technik'; PropertyName = 'TechnicalNote'; Type = 'text'; Width = 220; DefaultVisible = $true; Editable = $true }
        [pscustomobject]@{ Key = 'logistics'; Label = 'Bemerkung'; PropertyName = 'LogisticsNote'; Type = 'text'; Width = 220; DefaultVisible = $true; Editable = $true }
        [pscustomobject]@{ Key = 'unit_main'; Label = 'BZE'; PropertyName = 'BaseUnit'; Type = 'unit'; Width = 90; DefaultVisible = $true; Editable = $true }
        [pscustomobject]@{ Key = 'quantity_target'; Label = 'TLG 74'; PropertyName = 'TargetQuantity'; Type = 'number'; Width = 90; DefaultVisible = $true; Editable = $true }
        [pscustomobject]@{ Key = 'is_dg'; Label = 'GefStoff'; PropertyName = 'IsHazardous'; Type = 'bool'; Width = 90; DefaultVisible = $true; Editable = $true }
        [pscustomobject]@{ Key = 'un_number'; Label = 'UN Number'; PropertyName = 'UnNumber'; Type = 'text'; Width = 110; DefaultVisible = $false; Editable = $true }
        [pscustomobject]@{ Key = 'hazmat_gefstoff_verlegung'; Label = 'GefStoff Verlegung'; PropertyName = 'HazmatGefstoffVerlegung'; Type = 'bool'; Width = 130; DefaultVisible = $true; Editable = $true; Code = 'gefstoff_verlegung' }
        [pscustomobject]@{ Key = 'hazmat_gefahrgut'; Label = 'Gefahrgut'; PropertyName = 'HazmatGefahrgut'; Type = 'bool'; Width = 95; DefaultVisible = $true; Editable = $true; Code = 'gefahrgut' }
        [pscustomobject]@{ Key = 'hazmat_batterie'; Label = 'Batterie'; PropertyName = 'HazmatBatterie'; Type = 'bool'; Width = 90; DefaultVisible = $true; Editable = $true; Code = 'batterie' }
        [pscustomobject]@{ Key = 'resp_flight'; Label = 'Flight'; PropertyName = 'RespFlight'; Type = 'bool'; Width = 80; DefaultVisible = $true; Editable = $true; Code = 'flight' }
        [pscustomobject]@{ Key = 'resp_waffen'; Label = 'Waffen'; PropertyName = 'RespWaffen'; Type = 'bool'; Width = 85; DefaultVisible = $true; Editable = $true; Code = 'waffen' }
        [pscustomobject]@{ Key = 'resp_munition'; Label = 'Munition'; PropertyName = 'RespMunition'; Type = 'bool'; Width = 90; DefaultVisible = $true; Editable = $true; Code = 'munition' }
        [pscustomobject]@{ Key = 'resp_rts'; Label = 'RTS'; PropertyName = 'RespRts'; Type = 'bool'; Width = 70; DefaultVisible = $true; Editable = $true; Code = 'rts' }
        [pscustomobject]@{ Key = 'resp_aug'; Label = 'AUG'; PropertyName = 'RespAug'; Type = 'bool'; Width = 70; DefaultVisible = $true; Editable = $true; Code = 'aug' }
        [pscustomobject]@{ Key = 'resp_wef'; Label = 'WEF'; PropertyName = 'RespWef'; Type = 'bool'; Width = 70; DefaultVisible = $true; Editable = $true; Code = 'wef' }
        [pscustomobject]@{ Key = 'resp_boge'; Label = 'BoGe'; PropertyName = 'RespBoge'; Type = 'bool'; Width = 70; DefaultVisible = $true; Editable = $true; Code = 'boge' }
        [pscustomobject]@{ Key = 'resp_hft'; Label = 'HFT'; PropertyName = 'RespHft'; Type = 'bool'; Width = 70; DefaultVisible = $true; Editable = $true; Code = 'hft' }
        [pscustomobject]@{ Key = 'resp_lme'; Label = 'LME'; PropertyName = 'RespLme'; Type = 'bool'; Width = 70; DefaultVisible = $true; Editable = $true; Code = 'lme' }
        [pscustomobject]@{ Key = 'resp_reg'; Label = 'REG'; PropertyName = 'RespReg'; Type = 'bool'; Width = 70; DefaultVisible = $true; Editable = $true; Code = 'reg' }
        [pscustomobject]@{ Key = 'resp_rnw'; Label = 'RNW'; PropertyName = 'RespRnw'; Type = 'bool'; Width = 70; DefaultVisible = $true; Editable = $true; Code = 'rnw' }
        [pscustomobject]@{ Key = 'resp_rad_reifen'; Label = 'Rad Reifen Shop'; PropertyName = 'RespRadReifen'; Type = 'bool'; Width = 125; DefaultVisible = $true; Editable = $true; Code = 'rad_reifen' }
        [pscustomobject]@{ Key = 'assign_ietpx_material'; Label = 'IETPX Material'; PropertyName = 'AssignIetpxMaterial'; Type = 'bool'; Width = 120; DefaultVisible = $true; Editable = $true; Code = 'ietpx_material' }
        [pscustomobject]@{ Key = 'assign_gun_on_ac'; Label = 'GUN ON AC'; PropertyName = 'AssignGunOnAc'; Type = 'bool'; Width = 105; DefaultVisible = $true; Editable = $true; Code = 'gun_on_ac' }
        [pscustomobject]@{ Key = 'assign_gun_off_ac'; Label = 'GUN OFF AC'; PropertyName = 'AssignGunOffAc'; Type = 'bool'; Width = 110; DefaultVisible = $true; Editable = $true; Code = 'gun_off_ac' }
        [pscustomobject]@{ Key = 'assign_gun'; Label = 'GUN'; PropertyName = 'AssignGun'; Type = 'bool'; Width = 70; DefaultVisible = $true; Editable = $true; Code = 'gun' }
        [pscustomobject]@{ Key = 'assign_iris_t'; Label = 'IRIS-T'; PropertyName = 'AssignIrisT'; Type = 'bool'; Width = 80; DefaultVisible = $true; Editable = $true; Code = 'iris_t' }
        [pscustomobject]@{ Key = 'assign_flare'; Label = 'FLARE'; PropertyName = 'AssignFlare'; Type = 'bool'; Width = 80; DefaultVisible = $true; Editable = $true; Code = 'flare' }
        [pscustomobject]@{ Key = 'assign_aim_120'; Label = 'AIM 120'; PropertyName = 'AssignAim120'; Type = 'bool'; Width = 90; DefaultVisible = $true; Editable = $true; Code = 'aim_120' }
        [pscustomobject]@{ Key = 'assign_sft_1000_l'; Label = '1000 l SFT'; PropertyName = 'AssignSft1000L'; Type = 'bool'; Width = 100; DefaultVisible = $true; Editable = $true; Code = 'sft_1000_l' }
        [pscustomobject]@{ Key = 'assign_gbu_48'; Label = 'GBU 48'; PropertyName = 'AssignGbu48'; Type = 'bool'; Width = 85; DefaultVisible = $true; Editable = $true; Code = 'gbu_48' }
        [pscustomobject]@{ Key = 'assign_meteor'; Label = 'Meteor'; PropertyName = 'AssignMeteor'; Type = 'bool'; Width = 80; DefaultVisible = $true; Editable = $true; Code = 'meteor' }
        [pscustomobject]@{ Key = 'assign_ldp'; Label = 'LDP'; PropertyName = 'AssignLdp'; Type = 'bool'; Width = 70; DefaultVisible = $true; Editable = $true; Code = 'ldp' }
        [pscustomobject]@{ Key = 'assign_iwp'; Label = 'IWP'; PropertyName = 'AssignIwp'; Type = 'bool'; Width = 70; DefaultVisible = $true; Editable = $true; Code = 'iwp' }
        [pscustomobject]@{ Key = 'assign_cfp'; Label = 'CFP'; PropertyName = 'AssignCfp'; Type = 'bool'; Width = 70; DefaultVisible = $true; Editable = $true; Code = 'cfp' }
        [pscustomobject]@{ Key = 'assign_mfrl'; Label = 'MFRL'; PropertyName = 'AssignMfrl'; Type = 'bool'; Width = 75; DefaultVisible = $true; Editable = $true; Code = 'mfrl' }
        [pscustomobject]@{ Key = 'assign_owp'; Label = 'OWP'; PropertyName = 'AssignOwp'; Type = 'bool'; Width = 70; DefaultVisible = $true; Editable = $true; Code = 'owp' }
        [pscustomobject]@{ Key = 'assign_chaff'; Label = 'CHAFF'; PropertyName = 'AssignChaff'; Type = 'bool'; Width = 80; DefaultVisible = $true; Editable = $true; Code = 'chaff' }
        [pscustomobject]@{ Key = 'assign_mel'; Label = 'MEL'; PropertyName = 'AssignMel'; Type = 'bool'; Width = 70; DefaultVisible = $true; Editable = $true; Code = 'mel' }
        [pscustomobject]@{ Key = 'assign_itspl'; Label = 'ITSPL'; PropertyName = 'AssignItspl'; Type = 'bool'; Width = 75; DefaultVisible = $true; Editable = $true; Code = 'itspl' }
        [pscustomobject]@{ Key = 'alternate_units_count'; Label = 'Alt. Einheiten'; PropertyName = 'AlternateUnitsCount'; Type = 'number'; Width = 110; DefaultVisible = $true; Editable = $false }
        [pscustomobject]@{ Key = 'alternates_count'; Label = 'Alternativen'; PropertyName = 'AlternatesCount'; Type = 'number'; Width = 100; DefaultVisible = $true; Editable = $false }
    )

    $index = 0
    foreach ($definition in $orderedDefinitions) {
        $index++
        $copy = [pscustomobject][ordered]@{
            Key            = $definition.Key
            Label          = $definition.Label
            PropertyName   = $definition.PropertyName
            Type           = $definition.Type
            Width          = $definition.Width
            DefaultVisible = [bool]$definition.DefaultVisible
            Editable       = [bool]$definition.Editable
            Index          = ($index - 1)
            Code           = $(if ($definition.PSObject.Properties['Code']) { Get-NormalizedString $definition.Code } else { '' })
        }
        [void]$definitions.Add($copy)
    }

    return $definitions.ToArray()
}

function Update-GridRowDerivedValues {
    param([Parameter(Mandatory = $true)]$Row)

    $Row.AlternateUnitsCount = (@(ConvertTo-ObjectArray $Row.AlternateUnits)).Count
    $Row.AlternatesCount = (@(ConvertTo-ObjectArray $Row.Alternates)).Count
}

function Convert-MaterialToGridRow {
    param([Parameter(Mandatory = $true)]$Material)

    $hazmatFlags = @{}
    foreach ($flag in @(ConvertTo-ObjectArray $Material.hazmat.flags)) {
        $hazmatFlags[(Get-NormalizedString $flag)] = $true
    }

    $responsibilities = @{}
    foreach ($code in @(ConvertTo-ObjectArray $Material.assignments.responsibility_codes)) {
        $responsibilities[(Get-NormalizedString $code)] = $true
    }

    $assignments = @{}
    foreach ($code in @(ConvertTo-ObjectArray $Material.assignments.assignment_tags)) {
        $assignments[(Get-NormalizedString $code)] = $true
    }

    $alternateUnitRows = New-Object System.Collections.Generic.List[object]
    foreach ($alternateUnit in @(ConvertTo-ObjectArray $Material.quantity.alternate_units)) {
        [void]$alternateUnitRows.Add([pscustomobject]@{
                unit_code          = Get-NormalizedString $alternateUnit.unit_code
                conversion_to_base = [double]$alternateUnit.conversion_to_base
            })
    }

    $alternateRows = New-Object System.Collections.Generic.List[object]
    foreach ($alternate in @(ConvertTo-ObjectArray $Material.alternates)) {
        [void]$alternateRows.Add([pscustomobject]@{
                position             = [int]$alternate.position
                identifier_value     = Get-NormalizedString $alternate.identifier.value
                material_status_code = Get-NormalizedString $alternate.material_status_code
                preferred_unit_code  = Get-NormalizedString $alternate.preferred_unit_code
            })
    }

    $row = [pscustomobject][ordered]@{
        ImportId                 = [int]$Material.id
        SupplyNumber             = Get-NormalizedString $Material.identifiers.supply_number
        MaterialNumber           = Get-NormalizedString $Material.identifiers.matnr
        MaterialStatus           = Get-NormalizedString $Material.status.material_status_code
        NatoStockNumber          = Get-NormalizedString $Material.identifiers.nato_stock_number
        IsDecentral              = [bool]$Material.classification.is_decentral
        ExtWg                    = Get-NormalizedString $Material.classification.ext_wg
        Creditor                 = Get-NormalizedString $Material.classification.creditor
        ArticleNumber            = Get-NormalizedString $Material.identifiers.article_number
        Description              = Get-NormalizedString $Material.texts.short_description
        TechnicalNote            = Get-NormalizedString $Material.texts.technical_note
        LogisticsNote            = Get-NormalizedString $Material.texts.logistics_note
        BaseUnit                 = Get-NormalizedString $Material.quantity.base_unit
        TargetQuantity           = [double]$Material.quantity.target
        IsHazardous              = [bool]$Material.hazmat.is_hazardous
        UnNumber                 = Get-NormalizedString $Material.hazmat.un_number
        HazmatGefstoffVerlegung  = $hazmatFlags.ContainsKey('gefstoff_verlegung')
        HazmatGefahrgut          = $hazmatFlags.ContainsKey('gefahrgut')
        HazmatBatterie           = $hazmatFlags.ContainsKey('batterie')
        RespFlight               = $responsibilities.ContainsKey('flight')
        RespWaffen               = $responsibilities.ContainsKey('waffen')
        RespMunition             = $responsibilities.ContainsKey('munition')
        RespRts                  = $responsibilities.ContainsKey('rts')
        RespAug                  = $responsibilities.ContainsKey('aug')
        RespWef                  = $responsibilities.ContainsKey('wef')
        RespBoge                 = $responsibilities.ContainsKey('boge')
        RespHft                  = $responsibilities.ContainsKey('hft')
        RespLme                  = $responsibilities.ContainsKey('lme')
        RespReg                  = $responsibilities.ContainsKey('reg')
        RespRnw                  = $responsibilities.ContainsKey('rnw')
        RespRadReifen            = $responsibilities.ContainsKey('rad_reifen')
        AssignIetpxMaterial      = $assignments.ContainsKey('ietpx_material')
        AssignGunOnAc            = $assignments.ContainsKey('gun_on_ac')
        AssignGunOffAc           = $assignments.ContainsKey('gun_off_ac')
        AssignGun                = $assignments.ContainsKey('gun')
        AssignIrisT              = $assignments.ContainsKey('iris_t')
        AssignFlare              = $assignments.ContainsKey('flare')
        AssignAim120             = $assignments.ContainsKey('aim_120')
        AssignSft1000L           = $assignments.ContainsKey('sft_1000_l')
        AssignGbu48              = $assignments.ContainsKey('gbu_48')
        AssignMeteor             = $assignments.ContainsKey('meteor')
        AssignLdp                = $assignments.ContainsKey('ldp')
        AssignIwp                = $assignments.ContainsKey('iwp')
        AssignCfp                = $assignments.ContainsKey('cfp')
        AssignMfrl               = $assignments.ContainsKey('mfrl')
        AssignOwp                = $assignments.ContainsKey('owp')
        AssignChaff              = $assignments.ContainsKey('chaff')
        AssignMel                = $assignments.ContainsKey('mel')
        AssignItspl              = $assignments.ContainsKey('itspl')
        AlternateUnits           = New-ObservableCollection -Items $alternateUnitRows.ToArray()
        Alternates               = New-ObservableCollection -Items $alternateRows.ToArray()
        AlternateUnitsCount      = 0
        AlternatesCount          = 0
    }

    Update-GridRowDerivedValues -Row $row
    return $row
}

function Get-GridRowLabel {
    param([Parameter(Mandatory = $true)]$Row)

    $materialNumber = Get-NormalizedString $Row.MaterialNumber
    if (-not [string]::IsNullOrWhiteSpace($materialNumber)) {
        return "Material '$materialNumber'"
    }

    $idText = Get-NormalizedString $Row.ImportId
    if (-not [string]::IsNullOrWhiteSpace($idText)) {
        return "ID '$idText'"
    }

    return 'new row'
}

function Convert-GridRowToMaterialBuildResult {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][object[]]$ColumnDefinitions,
        [string]$DefaultUnitCode = 'EA'
    )

    Update-GridRowDerivedValues -Row $Row

    $errors = New-Object System.Collections.Generic.List[string]
    $rowLabel = Get-GridRowLabel -Row $Row

    $idResult = ConvertTo-IntParseResult $Row.ImportId
    if (-not $idResult.Success) {
        [void]$errors.Add("${rowLabel}: ID must be a whole number.")
    }

    $quantityResult = ConvertTo-NumberParseResult $Row.TargetQuantity
    if (-not $quantityResult.Success) {
        [void]$errors.Add("${rowLabel}: TLG 74 is not a valid number.")
    }

    $alternateUnits = New-Object System.Collections.Generic.List[object]
    $altUnitIndex = 0
    foreach ($alternateUnit in @(ConvertTo-ObjectArray $Row.AlternateUnits)) {
        $altUnitIndex++
        $conversionResult = ConvertTo-NumberParseResult $alternateUnit.conversion_to_base
        if (-not $conversionResult.Success) {
            [void]$errors.Add("${rowLabel}: alternate unit row $altUnitIndex has an invalid conversion.")
        }

        [void]$alternateUnits.Add([pscustomobject][ordered]@{
                unit_code          = ConvertTo-NullableString $alternateUnit.unit_code
                conversion_to_base = [double]$conversionResult.Value
            })
    }

    $alternates = New-Object System.Collections.Generic.List[object]
    $alternateIndex = 0
    foreach ($alternate in @(ConvertTo-ObjectArray $Row.Alternates)) {
        $alternateIndex++
        $positionResult = ConvertTo-IntParseResult $alternate.position
        if (-not $positionResult.Success) {
            [void]$errors.Add("${rowLabel}: alternate row $alternateIndex has an invalid position.")
        }

        [void]$alternates.Add([pscustomobject][ordered]@{
                position             = [int]$positionResult.Value
                identifier           = [pscustomobject][ordered]@{
                    type  = 'matnr'
                    value = Get-NormalizedString $alternate.identifier_value
                }
                material_status_code = Get-NormalizedString $alternate.material_status_code
                preferred_unit_code  = ConvertTo-NullableString $alternate.preferred_unit_code
            })
    }

    $hazmatFlags = New-Object System.Collections.Generic.List[string]
    $responsibilityCodes = New-Object System.Collections.Generic.List[string]
    $assignmentTags = New-Object System.Collections.Generic.List[string]

    foreach ($columnDefinition in @($ColumnDefinitions)) {
        if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $columnDefinition.Code))) {
            continue
        }

        $propertyValue = [bool](Get-DeepPropertyValue $Row $columnDefinition.PropertyName $false)
        if (-not $propertyValue) {
            continue
        }

        switch -Regex ((Get-NormalizedString $columnDefinition.Key)) {
            '^hazmat_' { [void]$hazmatFlags.Add((Get-NormalizedString $columnDefinition.Code)) }
            '^resp_' { [void]$responsibilityCodes.Add((Get-NormalizedString $columnDefinition.Code)) }
            '^assign_' { [void]$assignmentTags.Add((Get-NormalizedString $columnDefinition.Code)) }
        }
    }

    $materialNumber = Get-NormalizedString $Row.MaterialNumber
    $canonicalKey = if ([string]::IsNullOrWhiteSpace($materialNumber)) { '' } else { Get-CanonicalKey -Type 'matnr' -Value $materialNumber }
    $baseUnit = Get-NormalizedString $Row.BaseUnit
    if ([string]::IsNullOrWhiteSpace($baseUnit)) {
        $baseUnit = $DefaultUnitCode
    }

    $candidate = [pscustomobject][ordered]@{
        id                 = [int]$idResult.Value
        canonical_key      = $canonicalKey
        primary_identifier = [pscustomobject][ordered]@{
            type  = 'matnr'
            value = $materialNumber
        }
        identifiers        = [pscustomobject][ordered]@{
            matnr             = ConvertTo-NullableString $materialNumber
            supply_number     = ConvertTo-NullableString $Row.SupplyNumber
            article_number    = ConvertTo-NullableString $Row.ArticleNumber
            nato_stock_number = ConvertTo-NullableString $Row.NatoStockNumber
        }
        status             = [pscustomobject][ordered]@{
            material_status_code = $(if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $Row.MaterialStatus))) { 'XX' } else { Get-NormalizedString $Row.MaterialStatus })
        }
        texts              = [pscustomobject][ordered]@{
            short_description = Get-NormalizedString $Row.Description
            technical_note    = Get-NormalizedString $Row.TechnicalNote
            logistics_note    = Get-NormalizedString $Row.LogisticsNote
        }
        classification     = [pscustomobject][ordered]@{
            ext_wg       = Get-NormalizedString $Row.ExtWg
            is_decentral = [bool]$Row.IsDecentral
            creditor     = ConvertTo-NullableString $Row.Creditor
        }
        hazmat             = [pscustomobject][ordered]@{
            is_hazardous = [bool]$Row.IsHazardous
            un_number    = ConvertTo-NullableString $Row.UnNumber
            flags        = @(ConvertTo-UniqueStringArray $hazmatFlags.ToArray())
        }
        quantity           = [pscustomobject][ordered]@{
            base_unit       = $baseUnit
            target          = [double]$quantityResult.Value
            alternate_units = $alternateUnits.ToArray()
        }
        alternates         = $alternates.ToArray()
        assignments        = [pscustomobject][ordered]@{
            responsibility_codes = @(ConvertTo-UniqueStringArray $responsibilityCodes.ToArray())
            assignment_tags      = @(ConvertTo-UniqueStringArray $assignmentTags.ToArray())
        }
    }

    return [pscustomobject]@{
        Row       = $Row
        RowLabel  = $rowLabel
        Candidate = $candidate
        Errors    = @($errors)
    }
}

function Test-GridMaterialCandidates {
    param(
        [Parameter(Mandatory = $true)][object[]]$BuildResults,
        [Parameter(Mandatory = $true)]$LookupData
    )

    $messages = New-Object System.Collections.Generic.List[string]
    $validUnitCodes = @($LookupData.unit_codes | ForEach-Object { Get-NormalizedString $_.code })
    $validHazmatFlags = @($LookupData.hazmat_flags | ForEach-Object { Get-NormalizedString $_.code })
    $validResponsibilityCodes = @($LookupData.responsibility_codes | ForEach-Object { Get-NormalizedString $_.code })
    $validAssignmentTags = @($LookupData.assignment_tags | ForEach-Object { Get-NormalizedString $_.code })
    $seenIds = @{}
    $seenCanonicalKeys = @{}

    foreach ($buildResult in @($BuildResults)) {
        foreach ($buildError in @($buildResult.Errors)) {
            if (-not [string]::IsNullOrWhiteSpace((Get-NormalizedString $buildError))) {
                [void]$messages.Add((Get-NormalizedString $buildError))
            }
        }

        $candidate = $buildResult.Candidate
        $rowLabel = $buildResult.RowLabel

        if ([int]$candidate.id -le 0) {
            [void]$messages.Add("${rowLabel}: ID must be greater than 0.")
        }
        elseif ($seenIds.ContainsKey([int]$candidate.id)) {
            [void]$messages.Add("${rowLabel}: ID $($candidate.id) is duplicated.")
        }
        else {
            $seenIds[[int]$candidate.id] = $true
        }

        if ((Get-NormalizedString $candidate.primary_identifier.type) -ne 'matnr') {
            [void]$messages.Add("${rowLabel}: primary identifier type must be matnr.")
        }

        if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $candidate.primary_identifier.value))) {
            [void]$messages.Add("${rowLabel}: Materialnummer SASPF is required.")
        }

        if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $candidate.canonical_key))) {
            [void]$messages.Add("${rowLabel}: canonical key could not be generated.")
        }
        elseif ($seenCanonicalKeys.ContainsKey((Get-NormalizedString $candidate.canonical_key))) {
            [void]$messages.Add("${rowLabel}: canonical key '$($candidate.canonical_key)' is duplicated.")
        }
        else {
            $seenCanonicalKeys[(Get-NormalizedString $candidate.canonical_key)] = $true
        }

        if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $candidate.quantity.base_unit)) -or -not ($validUnitCodes -contains (Get-NormalizedString $candidate.quantity.base_unit))) {
            [void]$messages.Add("${rowLabel}: BZE must exist in the lookup.")
        }

        if ([double]$candidate.quantity.target -lt 0) {
            [void]$messages.Add("${rowLabel}: TLG 74 must be 0 or greater.")
        }

        $seenAltUnits = @{}
        $altUnitIndex = 0
        foreach ($alternateUnit in @(ConvertTo-ObjectArray $candidate.quantity.alternate_units)) {
            $altUnitIndex++
            $unitCode = Get-NormalizedString $alternateUnit.unit_code
            if ([string]::IsNullOrWhiteSpace($unitCode)) {
                [void]$messages.Add("${rowLabel}: alternate unit row $altUnitIndex requires a unit code.")
            }
            elseif (-not ($validUnitCodes -contains $unitCode)) {
                [void]$messages.Add("${rowLabel}: alternate unit '$unitCode' is not in the lookup.")
            }
            elseif ($seenAltUnits.ContainsKey($unitCode)) {
                [void]$messages.Add("${rowLabel}: alternate unit '$unitCode' is duplicated.")
            }
            else {
                $seenAltUnits[$unitCode] = $true
            }

            if ([double]$alternateUnit.conversion_to_base -le 0) {
                [void]$messages.Add("${rowLabel}: alternate unit '$unitCode' must have a conversion greater than 0.")
            }
        }

        $seenPositions = @{}
        $alternateRowIndex = 0
        foreach ($alternate in @(ConvertTo-ObjectArray $candidate.alternates)) {
            $alternateRowIndex++
            $position = [int]$alternate.position
            if ($position -le 0) {
                [void]$messages.Add("${rowLabel}: alternate row $alternateRowIndex must have a positive position.")
            }
            elseif ($seenPositions.ContainsKey($position)) {
                [void]$messages.Add("${rowLabel}: alternate position $position is duplicated.")
            }
            else {
                $seenPositions[$position] = $true
            }

            if ((Get-NormalizedString $alternate.identifier.type) -ne 'matnr') {
                [void]$messages.Add("${rowLabel}: alternate row $alternateRowIndex must use matnr identifiers.")
            }

            if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $alternate.identifier.value))) {
                [void]$messages.Add("${rowLabel}: alternate row $alternateRowIndex requires an identifier value.")
            }

            $preferredUnit = Get-NormalizedString $alternate.preferred_unit_code
            if (-not [string]::IsNullOrWhiteSpace($preferredUnit) -and -not ($validUnitCodes -contains $preferredUnit)) {
                [void]$messages.Add("${rowLabel}: alternate row $alternateRowIndex has an invalid preferred unit.")
            }
        }

        foreach ($flag in @(ConvertTo-ObjectArray $candidate.hazmat.flags)) {
            if (-not ($validHazmatFlags -contains (Get-NormalizedString $flag))) {
                [void]$messages.Add("${rowLabel}: hazmat flag '$flag' is not in the lookup.")
            }
        }

        foreach ($code in @(ConvertTo-ObjectArray $candidate.assignments.responsibility_codes)) {
            if (-not ($validResponsibilityCodes -contains (Get-NormalizedString $code))) {
                [void]$messages.Add("${rowLabel}: responsibility code '$code' is not in the lookup.")
            }
        }

        foreach ($code in @(ConvertTo-ObjectArray $candidate.assignments.assignment_tags)) {
            if (-not ($validAssignmentTags -contains (Get-NormalizedString $code))) {
                [void]$messages.Add("${rowLabel}: assignment tag '$code' is not in the lookup.")
            }
        }
    }

    return [pscustomobject]@{
        IsValid  = ($messages.Count -eq 0)
        Messages = @($messages | Select-Object -Unique)
    }
}

function Get-GridRowValueForDisplay {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)]$ColumnDefinition
    )

    $value = Get-DeepPropertyValue $Row $ColumnDefinition.PropertyName
    switch ((Get-NormalizedString $ColumnDefinition.Type)) {
        'bool' { return $(if ([bool]$value) { 'TRUE' } else { 'FALSE' }) }
        'number' { return $(if ($null -eq $value) { '' } else { [string]$value }) }
        default { return Get-NormalizedString $value }
    }
}

function Get-GridRowSearchIndex {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][object[]]$ColumnDefinitions
    )

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($columnDefinition in @($ColumnDefinitions)) {
        [void]$parts.Add((Get-GridRowValueForDisplay -Row $Row -ColumnDefinition $columnDefinition))
    }

    [void]$parts.Add(([string]$Row.AlternateUnitsCount))
    [void]$parts.Add(([string]$Row.AlternatesCount))
    return (($parts.ToArray() -join ' ').ToLowerInvariant())
}

function Test-GridRowRule {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)]$Rule
    )

    $propertyName = Get-NormalizedString $Rule.PropertyName
    $operator = Get-NormalizedString $Rule.Operator
    $ruleType = Get-NormalizedString $Rule.Type
    $ruleValue = Get-NormalizedString $Rule.Value
    $rowValue = Get-DeepPropertyValue $Row $propertyName

    switch ($ruleType) {
        'bool' {
            switch ($operator) {
                'is yes' { return [bool]$rowValue }
                'is no' { return (-not [bool]$rowValue) }
                default { return $true }
            }
        }
        'number' {
            $parseResult = ConvertTo-NumberParseResult $ruleValue
            if (-not $parseResult.Success) {
                return $false
            }

            $numericValue = [double]$rowValue
            switch ($operator) {
                '=' { return ($numericValue -eq [double]$parseResult.Value) }
                '>=' { return ($numericValue -ge [double]$parseResult.Value) }
                '<=' { return ($numericValue -le [double]$parseResult.Value) }
                default { return $true }
            }
        }
        default {
            $normalizedRowValue = (Get-NormalizedString $rowValue).ToLowerInvariant()
            $normalizedRuleValue = $ruleValue.ToLowerInvariant()
            switch ($operator) {
                'equals' { return ($normalizedRowValue -eq $normalizedRuleValue) }
                'starts with' { return $normalizedRowValue.StartsWith($normalizedRuleValue) }
                default { return $normalizedRowValue.Contains($normalizedRuleValue) }
            }
        }
    }
}

function Get-FilteredGridRows {
    param(
        [AllowNull()][object[]]$Rows,
        [Parameter(Mandatory = $true)][object[]]$ColumnDefinitions,
        [string]$SearchText,
        [switch]$HazardousOnly,
        [switch]$DecentralOnly,
        [AllowNull()][object[]]$AdvancedRules
    )

    $normalizedSearch = (Get-NormalizedString $SearchText).ToLowerInvariant()
    $results = New-Object System.Collections.Generic.List[object]

    foreach ($row in @(ConvertTo-ObjectArray $Rows)) {
        Update-GridRowDerivedValues -Row $row

        if ($HazardousOnly -and -not [bool]$row.IsHazardous) {
            continue
        }

        if ($DecentralOnly -and -not [bool]$row.IsDecentral) {
            continue
        }

        if (-not [string]::IsNullOrWhiteSpace($normalizedSearch)) {
            $searchIndex = Get-GridRowSearchIndex -Row $row -ColumnDefinitions $ColumnDefinitions
            if (-not $searchIndex.Contains($normalizedSearch)) {
                continue
            }
        }

        $matchesAdvancedRules = $true
        foreach ($rule in @(ConvertTo-ObjectArray $AdvancedRules)) {
            if (-not (Test-GridRowRule -Row $row -Rule $rule)) {
                $matchesAdvancedRules = $false
                break
            }
        }

        if ($matchesAdvancedRules) {
            [void]$results.Add($row)
        }
    }

    return @($results | Sort-Object MaterialNumber, ImportId)
}

function Get-ExportValue {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)]$ColumnDefinition
    )

    $value = Get-DeepPropertyValue $Row $ColumnDefinition.PropertyName
    switch ((Get-NormalizedString $ColumnDefinition.Type)) {
        'bool' { return $(if ([bool]$value) { 'TRUE' } else { 'FALSE' }) }
        'number' { return $(if ($null -eq $value) { '' } else { [string]$value }) }
        default { return Get-NormalizedString $value }
    }
}

function Convert-GridRowsToExportObjects {
    param(
        [Parameter(Mandatory = $true)][object[]]$Rows,
        [Parameter(Mandatory = $true)][object[]]$ColumnDefinitions
    )

    $exportRows = New-Object System.Collections.Generic.List[object]
    foreach ($row in @($Rows)) {
        $ordered = [ordered]@{}
        foreach ($columnDefinition in @($ColumnDefinitions)) {
            $ordered[$columnDefinition.Label] = Get-ExportValue -Row $row -ColumnDefinition $columnDefinition
        }

        [void]$exportRows.Add([pscustomobject]$ordered)
    }

    return $exportRows.ToArray()
}

function Convert-GridRowsToNestedExportObjects {
    param([Parameter(Mandatory = $true)][object[]]$Rows)

    $nestedRows = New-Object System.Collections.Generic.List[object]
    foreach ($row in @($Rows)) {
        foreach ($alternateUnit in @(ConvertTo-ObjectArray $row.AlternateUnits)) {
            [void]$nestedRows.Add([pscustomobject][ordered]@{
                    'Nested Type'              = 'alternate_unit'
                    'ID'                       = [string]$row.ImportId
                    'Materialnummer SASPF'     = Get-NormalizedString $row.MaterialNumber
                    'Materialbezeichnung'      = Get-NormalizedString $row.Description
                    'Status MatNr'             = Get-NormalizedString $row.MaterialStatus
                    'Unit code'                = Get-NormalizedString $alternateUnit.unit_code
                    'Conversion to base'       = [string]$alternateUnit.conversion_to_base
                    'Position'                 = ''
                    'Alternate Materialnummer' = ''
                    'Alternate Status'         = ''
                    'Preferred unit'           = ''
                })
        }

        foreach ($alternate in @(ConvertTo-ObjectArray $row.Alternates)) {
            [void]$nestedRows.Add([pscustomobject][ordered]@{
                    'Nested Type'              = 'alternate'
                    'ID'                       = [string]$row.ImportId
                    'Materialnummer SASPF'     = Get-NormalizedString $row.MaterialNumber
                    'Materialbezeichnung'      = Get-NormalizedString $row.Description
                    'Status MatNr'             = Get-NormalizedString $row.MaterialStatus
                    'Unit code'                = ''
                    'Conversion to base'       = ''
                    'Position'                 = [string]$alternate.position
                    'Alternate Materialnummer' = Get-NormalizedString $alternate.identifier_value
                    'Alternate Status'         = Get-NormalizedString $alternate.material_status_code
                    'Preferred unit'           = Get-NormalizedString $alternate.preferred_unit_code
                })
        }
    }

    return $nestedRows.ToArray()
}

function Export-GridRowsToCsv {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object[]]$Rows,
        [Parameter(Mandatory = $true)][object[]]$ColumnDefinitions,
        [switch]$IncludeNestedData
    )

    $mainExportRows = Convert-GridRowsToExportObjects -Rows $Rows -ColumnDefinitions $ColumnDefinitions
    $mainExportRows | Export-Csv -Path $Path -Delimiter ';' -NoTypeInformation -Encoding UTF8

    if ($IncludeNestedData) {
        $directory = Split-Path -Path $Path -Parent
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        $nestedPath = Join-Path $directory ($baseName + '_nested.csv')
        $nestedRows = Convert-GridRowsToNestedExportObjects -Rows $Rows
        $nestedRows | Export-Csv -Path $nestedPath -Delimiter ';' -NoTypeInformation -Encoding UTF8
        return $nestedPath
    }

    return $null
}

function Set-ExcelWorksheetData {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [Parameter(Mandatory = $true)][string]$SheetName,
        [Parameter(Mandatory = $true)][object[]]$Rows
    )

    $Worksheet.Name = $SheetName

    if (@($Rows).Count -eq 0) {
        $Worksheet.Cells.Item(1, 1) = 'No data'
        return
    }

    $headers = @($Rows[0].PSObject.Properties.Name)
    for ($columnIndex = 0; $columnIndex -lt $headers.Count; $columnIndex++) {
        $Worksheet.Cells.Item(1, $columnIndex + 1) = $headers[$columnIndex]
    }

    $rowIndex = 2
    foreach ($row in @($Rows)) {
        for ($columnIndex = 0; $columnIndex -lt $headers.Count; $columnIndex++) {
            $header = $headers[$columnIndex]
            $Worksheet.Cells.Item($rowIndex, $columnIndex + 1) = Get-NormalizedString $row.PSObject.Properties[$header].Value
        }
        $rowIndex++
    }

    [void]$Worksheet.Columns.AutoFit()
}

function Export-GridRowsToXlsx {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object[]]$Rows,
        [Parameter(Mandatory = $true)][object[]]$ColumnDefinitions,
        [switch]$IncludeNestedData
    )

    $mainExportRows = Convert-GridRowsToExportObjects -Rows $Rows -ColumnDefinitions $ColumnDefinitions
    $nestedExportRows = if ($IncludeNestedData) { @(Convert-GridRowsToNestedExportObjects -Rows $Rows) } else { @() }

    $excel = $null
    $workbook = $null
    $materialsWorksheet = $null
    $nestedWorksheet = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()

        $materialsWorksheet = $workbook.Worksheets.Item(1)
        Set-ExcelWorksheetData -Worksheet $materialsWorksheet -SheetName 'Materials' -Rows $mainExportRows

        if ($IncludeNestedData) {
            $nestedWorksheet = if ($workbook.Worksheets.Count -ge 2) { $workbook.Worksheets.Item(2) } else { $workbook.Worksheets.Add() }
            Set-ExcelWorksheetData -Worksheet $nestedWorksheet -SheetName 'Nested Data' -Rows $nestedExportRows
        }

        while ($workbook.Worksheets.Count -gt $(if ($IncludeNestedData) { 2 } else { 1 })) {
            $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
        }

        $workbook.SaveAs($Path, 51)
    }
    finally {
        if ($null -ne $workbook) {
            $workbook.Close($true)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        }

        if ($null -ne $materialsWorksheet) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($materialsWorksheet)
        }

        if ($null -ne $nestedWorksheet) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($nestedWorksheet)
        }

        if ($null -ne $excel) {
            $excel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Start-GridViewerUi {
    $lookupData = Read-LookupFile -Path $Script:LookupPath
    $defaultIdentifierType = Get-NormalizedString $lookupData.identifier_types[0].code
    if ([string]::IsNullOrWhiteSpace($defaultIdentifierType)) {
        $defaultIdentifierType = 'matnr'
    }

    $defaultUnitCode = Get-NormalizedString $lookupData.unit_codes[0].code
    if ([string]::IsNullOrWhiteSpace($defaultUnitCode)) {
        $defaultUnitCode = 'EA'
    }

    $columnDefinitions = @(Get-GridColumnDefinitions)
    $filterDefinitions = @($columnDefinitions | ForEach-Object {
            [pscustomobject]@{
                Key          = $_.Key
                Label        = $_.Label
                PropertyName = $_.PropertyName
                Type         = $(if ($_.Type -eq 'unit') { 'text' } else { $_.Type })
            }
        })
    $filterDefinitionLookup = @{}
    foreach ($definition in @($filterDefinitions)) {
        $filterDefinitionLookup[$definition.Key] = $definition
    }

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Grid Viewer"
        Height="980"
        Width="1820"
        MinHeight="860"
        MinWidth="1460"
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
        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Margin" Value="0,0,10,0"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Background" Value="#0F766E"/>
            <Setter Property="BorderBrush" Value="#0F766E"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style x:Key="SecondaryButton" TargetType="Button">
            <Setter Property="Margin" Value="0,0,10,0"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="Foreground" Value="#0F172A"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="CanUserAddRows" Value="False"/>
            <Setter Property="CanUserDeleteRows" Value="False"/>
            <Setter Property="CanUserReorderColumns" Value="False"/>
            <Setter Property="CanUserResizeRows" Value="False"/>
            <Setter Property="RowHeaderWidth" Value="0"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="AlternatingRowBackground" Value="#F8FAFC"/>
            <Setter Property="EnableColumnVirtualization" Value="True"/>
            <Setter Property="EnableRowVirtualization" Value="True"/>
        </Style>
    </Window.Resources>
    <DockPanel>
        <Border DockPanel.Dock="Top" Background="#0F172A" Padding="18,16">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                    <TextBlock Text="Grid Viewer" FontFamily="Bahnschrift SemiBold" FontSize="28" Foreground="White"/>
                    <TextBlock Margin="0,4,0,0" Foreground="#CBD5E1" Text="Editable database grid with import-style columns, nested alternates, export, and lookup-aware save validation."/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top">
                    <Button x:Name="btnReloadDatabase" Content="Neu laden" Style="{StaticResource SecondaryButton}"/>
                    <Button x:Name="btnSaveDatabase" Content="Speichern" Style="{StaticResource PrimaryButton}"/>
                </StackPanel>
            </Grid>
        </Border>

        <Border DockPanel.Dock="Bottom" Background="White" BorderBrush="#E2E8F0" BorderThickness="1,1,0,0" Padding="18,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="16"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="txtDirtyState" Grid.Column="0" FontWeight="SemiBold" Foreground="#475569" VerticalAlignment="Center" Text="Saved"/>
                <TextBlock x:Name="txtStatus" Grid.Column="2" Foreground="#475569" VerticalAlignment="Center" Text="Ready"/>
                <TextBlock Grid.Column="3" Foreground="#94A3B8" VerticalAlignment="Center" Text="Ctrl+S Speichern | Ctrl+N Neu | Ctrl+D Duplizieren | Delete Loeschen"/>
            </Grid>
        </Border>

        <Grid Margin="18">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <Border Grid.Row="0" Background="White" BorderBrush="#E2E8F0" BorderThickness="1" CornerRadius="10" Padding="16">
                <StackPanel>
                    <WrapPanel>
                        <Button x:Name="btnNewMaterial" Content="Neu" Style="{StaticResource PrimaryButton}"/>
                        <Button x:Name="btnCloneMaterial" Content="Duplizieren" Style="{StaticResource SecondaryButton}"/>
                        <Button x:Name="btnDeleteMaterial" Content="Loeschen" Style="{StaticResource SecondaryButton}"/>
                        <Button x:Name="btnEditAlternates" Content="Alternates..." Style="{StaticResource SecondaryButton}"/>
                        <Button x:Name="btnExportCsv" Content="Export CSV" Style="{StaticResource SecondaryButton}"/>
                        <Button x:Name="btnExportXlsx" Content="Export XLSX" Style="{StaticResource SecondaryButton}"/>
                        <CheckBox x:Name="chkExportNested" Margin="8,6,0,0" Content="Nested data exportieren"/>
                    </WrapPanel>
                    <TextBlock Margin="0,8,0,0" Foreground="#64748B" TextWrapping="Wrap" Text="Die Haupttabelle ist direkt editierbar. Alternate units und Alternates werden ueber einen separaten Dialog pro Material bearbeitet und beim Speichern in die bestehende JSON-Struktur zurueckgeschrieben."/>
                </StackPanel>
            </Border>

            <Border Grid.Row="1" Margin="0,16,0,16" Background="White" BorderBrush="#E2E8F0" BorderThickness="1" CornerRadius="10" Padding="16">
                <StackPanel>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="txtSearch" Grid.Column="0" Margin="0" ToolTip="Search across all database-backed columns."/>
                        <Button x:Name="btnClearSearch" Grid.Column="1" Content="Clear" Style="{StaticResource SecondaryButton}" Margin="10,0,0,0"/>
                        <Button x:Name="btnOpenFilterMenu" Grid.Column="2" Content="Filter..." Style="{StaticResource SecondaryButton}" Margin="10,0,0,0"/>
                        <Button x:Name="btnOpenColumnMenu" Grid.Column="3" Content="Columns..." Style="{StaticResource SecondaryButton}" Margin="10,0,0,0"/>
                        <StackPanel Grid.Column="4" Orientation="Horizontal" Margin="12,0,0,0" VerticalAlignment="Center">
                            <CheckBox x:Name="chkFilterHazardous" Content="Nur GefStoff"/>
                            <CheckBox x:Name="chkFilterDecentral" Content="Nur Dezent"/>
                        </StackPanel>
                    </Grid>
                    <TextBlock x:Name="txtGridMeta" Margin="0,12,0,0" Foreground="#64748B" Text="0 visible / 0 total"/>
                </StackPanel>
            </Border>

            <Border Grid.Row="2" Background="White" BorderBrush="#E2E8F0" BorderThickness="1" CornerRadius="10" Padding="0">
                <DataGrid x:Name="dgGridViewer"
                          Margin="0"
                          SelectionMode="Single"
                          FrozenColumnCount="2"
                          IsReadOnly="False"
                          CanUserSortColumns="True">
                </DataGrid>
            </Border>
        </Grid>
    </DockPanel>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader([xml]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $window.DataContext = [pscustomobject]@{
        UnitCodeOptions = @($lookupData.unit_codes)
    }

    $btnReloadDatabase = $window.FindName('btnReloadDatabase')
    $btnSaveDatabase = $window.FindName('btnSaveDatabase')
    $btnNewMaterial = $window.FindName('btnNewMaterial')
    $btnCloneMaterial = $window.FindName('btnCloneMaterial')
    $btnDeleteMaterial = $window.FindName('btnDeleteMaterial')
    $btnEditAlternates = $window.FindName('btnEditAlternates')
    $btnExportCsv = $window.FindName('btnExportCsv')
    $btnExportXlsx = $window.FindName('btnExportXlsx')
    $chkExportNested = $window.FindName('chkExportNested')
    $txtSearch = $window.FindName('txtSearch')
    $btnClearSearch = $window.FindName('btnClearSearch')
    $btnOpenFilterMenu = $window.FindName('btnOpenFilterMenu')
    $btnOpenColumnMenu = $window.FindName('btnOpenColumnMenu')
    $chkFilterHazardous = $window.FindName('chkFilterHazardous')
    $chkFilterDecentral = $window.FindName('chkFilterDecentral')
    $txtGridMeta = $window.FindName('txtGridMeta')
    $txtDirtyState = $window.FindName('txtDirtyState')
    $txtStatus = $window.FindName('txtStatus')
    $dgGridViewer = $window.FindName('dgGridViewer')

    $state = [ordered]@{
        LookupData              = $lookupData
        AllRows                 = New-Object System.Collections.ArrayList
        FilteredRows            = New-ObservableCollection
        SelectedRow             = $null
        DatabaseDirty           = $false
        Loading                 = $false
        FilterHazardousOnly     = $false
        FilterDecentralOnly     = $false
        ActiveFilterRules       = @()
        VisibleColumnKeys       = @($columnDefinitions | ForEach-Object { $_.Key })
        ColumnDefinitions       = $columnDefinitions
    }

    $SetStatus = $null
    $UpdateDirtyState = $null
    $RefreshGrid = $null
    $SelectRow = $null
    $UpdateActionState = $null
    $ApplyColumnVisibility = $null
    $BuildGridColumns = $null
    $OpenAlternateEditor = $null
    $OpenFilterDialog = $null
    $OpenColumnDialog = $null
    $MarkDirty = $null
    $CommitActiveEdits = $null
    $LoadDatabase = $null
    $SaveDatabase = $null
    $ConfirmClose = $null
    $ValidateAndBuildMaterials = $null
    $ExportCurrentGrid = $null

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

    $UpdateDirtyState = {
        if ($state.DatabaseDirty) {
            $txtDirtyState.Text = 'Unsaved changes'
            $txtDirtyState.Foreground = '#B45309'
            $window.Title = 'Grid Viewer *'
        }
        else {
            $txtDirtyState.Text = 'Saved'
            $txtDirtyState.Foreground = '#0F766E'
            $window.Title = 'Grid Viewer'
        }

        & $UpdateActionState
    }

    $UpdateActionState = {
        $hasSelection = ($null -ne $state.SelectedRow)
        $btnCloneMaterial.IsEnabled = $hasSelection
        $btnDeleteMaterial.IsEnabled = $hasSelection
        $btnEditAlternates.IsEnabled = $hasSelection
        $btnExportCsv.IsEnabled = ($state.AllRows.Count -gt 0)
        $btnExportXlsx.IsEnabled = ($state.AllRows.Count -gt 0)
    }

    $BuildGridColumns = {
        $dgGridViewer.Columns.Clear()
        foreach ($columnDefinition in @($columnDefinitions)) {
            $column = $null
            switch ((Get-NormalizedString $columnDefinition.Type)) {
                'bool' {
                    $binding = New-Object System.Windows.Data.Binding($columnDefinition.PropertyName)
                    $binding.UpdateSourceTrigger = [System.Windows.Data.UpdateSourceTrigger]::PropertyChanged
                    $column = New-Object System.Windows.Controls.DataGridCheckBoxColumn
                    $column.Binding = $binding
                }
                'unit' {
                    $binding = New-Object System.Windows.Data.Binding($columnDefinition.PropertyName)
                    $binding.UpdateSourceTrigger = [System.Windows.Data.UpdateSourceTrigger]::PropertyChanged
                    $column = New-Object System.Windows.Controls.DataGridComboBoxColumn
                    $column.DisplayMemberPath = 'label'
                    $column.SelectedValuePath = 'code'
                    $column.ItemsSource = @($lookupData.unit_codes)
                    $column.SelectedValueBinding = $binding
                }
                default {
                    $binding = New-Object System.Windows.Data.Binding($columnDefinition.PropertyName)
                    $binding.UpdateSourceTrigger = [System.Windows.Data.UpdateSourceTrigger]::PropertyChanged
                    $column = New-Object System.Windows.Controls.DataGridTextColumn
                    $column.Binding = $binding
                }
            }

            $column.Header = $columnDefinition.Label
            $column.Width = $columnDefinition.Width
            $column.IsReadOnly = (-not [bool]$columnDefinition.Editable)
            [void]$dgGridViewer.Columns.Add($column)
        }
    }

    $ApplyColumnVisibility = {
        foreach ($columnDefinition in @($columnDefinitions)) {
            $dgGridViewer.Columns[$columnDefinition.Index].Visibility = if ($state.VisibleColumnKeys -contains $columnDefinition.Key) {
                [System.Windows.Visibility]::Visible
            }
            else {
                [System.Windows.Visibility]::Collapsed
            }
        }
    }

    $SelectRow = {
        param([AllowNull()]$PreferredRow)

        if ($null -eq $PreferredRow) {
            $dgGridViewer.SelectedItem = $null
            $state.SelectedRow = $null
            & $UpdateActionState
            return
        }

        $state.SelectedRow = $PreferredRow
        $dgGridViewer.SelectedItem = $PreferredRow
        if ($null -ne $dgGridViewer.SelectedItem) {
            $dgGridViewer.ScrollIntoView($PreferredRow)
        }
        & $UpdateActionState
    }

    $RefreshGrid = {
        param([AllowNull()]$PreferredRow)

        $filteredRows = @(Get-FilteredGridRows `
                -Rows $state.AllRows.ToArray() `
                -ColumnDefinitions $columnDefinitions `
                -SearchText $txtSearch.Text `
                -HazardousOnly:$state.FilterHazardousOnly `
                -DecentralOnly:$state.FilterDecentralOnly `
                -AdvancedRules $state.ActiveFilterRules)
        $state.FilteredRows = New-ObservableCollection -Items $filteredRows
        $dgGridViewer.ItemsSource = $state.FilteredRows

        $filteredCount = @($filteredRows).Count
        $totalCount = $state.AllRows.Count
        $hazardousCount = (@($state.AllRows.ToArray() | Where-Object { $_.IsHazardous })).Count
        $decentralCount = (@($state.AllRows.ToArray() | Where-Object { $_.IsDecentral })).Count
        $activeFilters = New-Object System.Collections.Generic.List[string]
        if ($state.FilterHazardousOnly) { [void]$activeFilters.Add('GefStoff') }
        if ($state.FilterDecentralOnly) { [void]$activeFilters.Add('Dezent') }
        foreach ($rule in @($state.ActiveFilterRules)) {
            $label = Get-NormalizedString $rule.Label
            $operator = Get-NormalizedString $rule.Operator
            $value = Get-NormalizedString $rule.Value
            if ((Get-NormalizedString $rule.Type) -eq 'bool') {
                [void]$activeFilters.Add("$label $operator")
            }
            else {
                [void]$activeFilters.Add("$label $operator $value")
            }
        }
        $filterLabel = if ($activeFilters.Count -gt 0) { $activeFilters -join ', ' } else { 'none' }
        $txtGridMeta.Text = "$filteredCount visible / $totalCount total | $hazardousCount hazardous | $decentralCount decentral | filters: $filterLabel"

        if ($null -ne $PreferredRow -and ($filteredRows -contains $PreferredRow)) {
            & $SelectRow -PreferredRow $PreferredRow
        }
        elseif ($filteredRows.Count -gt 0) {
            if ($null -ne $state.SelectedRow -and ($filteredRows -contains $state.SelectedRow)) {
                & $SelectRow -PreferredRow $state.SelectedRow
            }
            else {
                & $SelectRow -PreferredRow $filteredRows[0]
            }
        }
        else {
            & $SelectRow -PreferredRow $null
        }
    }

    $MarkDirty = {
        if ($state.Loading) {
            return
        }

        $state.DatabaseDirty = $true
        & $UpdateDirtyState
    }

    $CommitActiveEdits = {
        $dgGridViewer.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true) | Out-Null
        $dgGridViewer.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true) | Out-Null
    }

    $OpenAlternateEditor = {
        param([Parameter(Mandatory = $true)]$Row)

        $dialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Alternates"
        Height="700"
        Width="1040"
        MinHeight="620"
        MinWidth="900"
        WindowStartupLocation="CenterOwner"
        Background="#F8FAFC">
    <Grid Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="18"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0">
            <TextBlock FontFamily="Bahnschrift SemiBold" FontSize="22" Foreground="#0F172A" Text="Alternate Editor"/>
            <TextBlock x:Name="txtAlternateEditorMeta" Margin="0,6,0,0" Foreground="#64748B" Text=""/>
        </StackPanel>
        <DockPanel Grid.Row="1" LastChildFill="False">
            <TextBlock DockPanel.Dock="Left" Foreground="#334155" FontWeight="SemiBold" Text="Alternate units"/>
            <Button x:Name="btnAddAlternateUnit" DockPanel.Dock="Right" Content="+ Einheit" Width="110" Margin="0,0,8,0"/>
            <Button x:Name="btnRemoveAlternateUnit" DockPanel.Dock="Right" Content="- Einheit" Width="110"/>
        </DockPanel>
        <DataGrid x:Name="gridAlternateUnits" Grid.Row="3" Height="180" CanUserResizeColumns="True" RowHeaderWidth="0"/>
        <DockPanel Grid.Row="4" Margin="0,18,0,0" LastChildFill="False" VerticalAlignment="Top">
            <TextBlock DockPanel.Dock="Left" Foreground="#334155" FontWeight="SemiBold" Text="Alternates"/>
            <Button x:Name="btnAddAlternate" DockPanel.Dock="Right" Content="+ Alternate" Width="110" Margin="0,0,8,0"/>
            <Button x:Name="btnRemoveAlternate" DockPanel.Dock="Right" Content="- Alternate" Width="110"/>
        </DockPanel>
        <Grid Grid.Row="4" Margin="0,46,0,0">
            <DataGrid x:Name="gridAlternates" Height="300" CanUserResizeColumns="True" RowHeaderWidth="0"/>
        </Grid>
        <DockPanel Grid.Row="5" Margin="0,18,0,0" LastChildFill="False">
            <Button x:Name="btnCancelAlternates" DockPanel.Dock="Right" Width="110" Margin="10,0,0,0" Content="Cancel"/>
            <Button x:Name="btnApplyAlternates" DockPanel.Dock="Right" Width="110" Content="Apply"/>
        </DockPanel>
    </Grid>
</Window>
"@

        $dialogReader = New-Object System.Xml.XmlNodeReader([xml]$dialogXaml)
        $dialog = [Windows.Markup.XamlReader]::Load($dialogReader)
        $dialog.Owner = $window
        $dialog.DataContext = [pscustomobject]@{
            UnitCodeOptions = @($lookupData.unit_codes)
        }

        $txtAlternateEditorMeta = $dialog.FindName('txtAlternateEditorMeta')
        $gridAlternateUnits = $dialog.FindName('gridAlternateUnits')
        $gridAlternates = $dialog.FindName('gridAlternates')
        $btnAddAlternateUnit = $dialog.FindName('btnAddAlternateUnit')
        $btnRemoveAlternateUnit = $dialog.FindName('btnRemoveAlternateUnit')
        $btnAddAlternate = $dialog.FindName('btnAddAlternate')
        $btnRemoveAlternate = $dialog.FindName('btnRemoveAlternate')
        $btnCancelAlternates = $dialog.FindName('btnCancelAlternates')
        $btnApplyAlternates = $dialog.FindName('btnApplyAlternates')

        $txtAlternateEditorMeta.Text = "{0} | {1}" -f (Get-NormalizedString $Row.MaterialNumber), (Get-NormalizedString $Row.Description)

        $tempAlternateUnits = New-Object System.Collections.Generic.List[object]
        foreach ($alternateUnit in @(ConvertTo-ObjectArray $Row.AlternateUnits)) {
            [void]$tempAlternateUnits.Add([pscustomobject]@{
                    unit_code          = Get-NormalizedString $alternateUnit.unit_code
                    conversion_to_base = [double]$alternateUnit.conversion_to_base
                })
        }

        $tempAlternates = New-Object System.Collections.Generic.List[object]
        foreach ($alternate in @(ConvertTo-ObjectArray $Row.Alternates)) {
            [void]$tempAlternates.Add([pscustomobject]@{
                    position             = [int]$alternate.position
                    identifier_value     = Get-NormalizedString $alternate.identifier_value
                    material_status_code = Get-NormalizedString $alternate.material_status_code
                    preferred_unit_code  = Get-NormalizedString $alternate.preferred_unit_code
                })
        }

        $gridAlternateUnits.ItemsSource = New-ObservableCollection -Items $tempAlternateUnits.ToArray()
        $gridAlternates.ItemsSource = New-ObservableCollection -Items $tempAlternates.ToArray()

        $alternateUnitColumn1 = New-Object System.Windows.Controls.DataGridComboBoxColumn
        $alternateUnitColumn1.Header = 'Unit code'
        $alternateUnitColumn1.Width = 150
        $alternateUnitColumn1.DisplayMemberPath = 'label'
        $alternateUnitColumn1.SelectedValuePath = 'code'
        $alternateUnitColumn1.ItemsSource = @($lookupData.unit_codes)
        $altUnitBinding = New-Object System.Windows.Data.Binding('unit_code')
        $altUnitBinding.UpdateSourceTrigger = [System.Windows.Data.UpdateSourceTrigger]::PropertyChanged
        $alternateUnitColumn1.SelectedValueBinding = $altUnitBinding
        [void]$gridAlternateUnits.Columns.Add($alternateUnitColumn1)

        $alternateUnitColumn2 = New-Object System.Windows.Controls.DataGridTextColumn
        $alternateUnitColumn2.Header = 'Conversion to base'
        $alternateUnitColumn2.Width = '*'
        $alternateUnitBinding2 = New-Object System.Windows.Data.Binding('conversion_to_base')
        $alternateUnitBinding2.UpdateSourceTrigger = [System.Windows.Data.UpdateSourceTrigger]::PropertyChanged
        $alternateUnitColumn2.Binding = $alternateUnitBinding2
        [void]$gridAlternateUnits.Columns.Add($alternateUnitColumn2)

        $alternateColumn1 = New-Object System.Windows.Controls.DataGridTextColumn
        $alternateColumn1.Header = 'Pos'
        $alternateColumn1.Width = 72
        $altBinding1 = New-Object System.Windows.Data.Binding('position')
        $altBinding1.UpdateSourceTrigger = [System.Windows.Data.UpdateSourceTrigger]::PropertyChanged
        $alternateColumn1.Binding = $altBinding1
        [void]$gridAlternates.Columns.Add($alternateColumn1)

        $alternateColumn2 = New-Object System.Windows.Controls.DataGridTextColumn
        $alternateColumn2.Header = 'Materialnummer'
        $alternateColumn2.Width = '*'
        $altBinding2 = New-Object System.Windows.Data.Binding('identifier_value')
        $altBinding2.UpdateSourceTrigger = [System.Windows.Data.UpdateSourceTrigger]::PropertyChanged
        $alternateColumn2.Binding = $altBinding2
        [void]$gridAlternates.Columns.Add($alternateColumn2)

        $alternateColumn3 = New-Object System.Windows.Controls.DataGridTextColumn
        $alternateColumn3.Header = 'Mat status'
        $alternateColumn3.Width = 110
        $altBinding3 = New-Object System.Windows.Data.Binding('material_status_code')
        $altBinding3.UpdateSourceTrigger = [System.Windows.Data.UpdateSourceTrigger]::PropertyChanged
        $alternateColumn3.Binding = $altBinding3
        [void]$gridAlternates.Columns.Add($alternateColumn3)

        $alternateColumn4 = New-Object System.Windows.Controls.DataGridComboBoxColumn
        $alternateColumn4.Header = 'Preferred unit'
        $alternateColumn4.Width = 150
        $alternateColumn4.DisplayMemberPath = 'label'
        $alternateColumn4.SelectedValuePath = 'code'
        $alternateColumn4.ItemsSource = @($lookupData.unit_codes)
        $altBinding4 = New-Object System.Windows.Data.Binding('preferred_unit_code')
        $altBinding4.UpdateSourceTrigger = [System.Windows.Data.UpdateSourceTrigger]::PropertyChanged
        $alternateColumn4.SelectedValueBinding = $altBinding4
        [void]$gridAlternates.Columns.Add($alternateColumn4)

        $btnAddAlternateUnit.Add_Click({
                [void]$gridAlternateUnits.ItemsSource.Add([pscustomobject]@{ unit_code = $defaultUnitCode; conversion_to_base = 1.0 })
            })
        $btnRemoveAlternateUnit.Add_Click({
                if ($null -eq $gridAlternateUnits.SelectedItem) {
                    return
                }
                [void]$gridAlternateUnits.ItemsSource.Remove($gridAlternateUnits.SelectedItem)
            })
        $btnAddAlternate.Add_Click({
                $nextPosition = (@(ConvertTo-ObjectArray $gridAlternates.ItemsSource)).Count + 1
                [void]$gridAlternates.ItemsSource.Add([pscustomobject]@{
                        position             = $nextPosition
                        identifier_value     = ''
                        material_status_code = ''
                        preferred_unit_code  = $defaultUnitCode
                    })
            })
        $btnRemoveAlternate.Add_Click({
                if ($null -eq $gridAlternates.SelectedItem) {
                    return
                }
                [void]$gridAlternates.ItemsSource.Remove($gridAlternates.SelectedItem)
            })
        $btnCancelAlternates.Add_Click({
                $dialog.DialogResult = $false
                $dialog.Close()
            })
        $btnApplyAlternates.Add_Click({
                $gridAlternateUnits.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true) | Out-Null
                $gridAlternateUnits.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true) | Out-Null
                $gridAlternates.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true) | Out-Null
                $gridAlternates.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true) | Out-Null

                $appliedAlternateUnits = New-Object System.Collections.Generic.List[object]
                foreach ($alternateUnit in @(ConvertTo-ObjectArray $gridAlternateUnits.ItemsSource)) {
                    [void]$appliedAlternateUnits.Add([pscustomobject]@{
                            unit_code          = Get-NormalizedString $alternateUnit.unit_code
                            conversion_to_base = $alternateUnit.conversion_to_base
                        })
                }

                $appliedAlternates = New-Object System.Collections.Generic.List[object]
                foreach ($alternate in @(ConvertTo-ObjectArray $gridAlternates.ItemsSource)) {
                    [void]$appliedAlternates.Add([pscustomobject]@{
                            position             = $alternate.position
                            identifier_value     = Get-NormalizedString $alternate.identifier_value
                            material_status_code = Get-NormalizedString $alternate.material_status_code
                            preferred_unit_code  = Get-NormalizedString $alternate.preferred_unit_code
                        })
                }

                $Row.AlternateUnits = New-ObservableCollection -Items $appliedAlternateUnits.ToArray()
                $Row.Alternates = New-ObservableCollection -Items $appliedAlternates.ToArray()
                Update-GridRowDerivedValues -Row $Row
                & $MarkDirty
                $dialog.DialogResult = $true
                $dialog.Close()
            })

        return [bool]$dialog.ShowDialog()
    }

    $ValidateAndBuildMaterials = {
        & $CommitActiveEdits

        $buildResults = New-Object System.Collections.Generic.List[object]
        foreach ($row in @($state.AllRows.ToArray())) {
            [void]$buildResults.Add((Convert-GridRowToMaterialBuildResult -Row $row -ColumnDefinitions $columnDefinitions -DefaultUnitCode $defaultUnitCode))
        }

        $validationResult = Test-GridMaterialCandidates -BuildResults $buildResults.ToArray() -LookupData $lookupData
        return [pscustomobject]@{
            IsValid      = $validationResult.IsValid
            Messages     = @($validationResult.Messages)
            BuildResults = @($buildResults.ToArray())
            Materials    = @($buildResults.ToArray() | ForEach-Object {
                    ConvertTo-NormalizedMaterial -Material $_.Candidate -DefaultIdentifierType $defaultIdentifierType -DefaultUnitCode $defaultUnitCode
                })
        }
    }

    $LoadDatabase = {
        $database = Read-DatabaseFile -Path $Script:DbPath -DefaultIdentifierType $defaultIdentifierType -DefaultUnitCode $defaultUnitCode
        $state.Loading = $true
        try {
            $state.AllRows.Clear()
            foreach ($material in @($database.materials)) {
                [void]$state.AllRows.Add((Convert-MaterialToGridRow -Material $material))
            }
            $state.DatabaseDirty = $false
            & $UpdateDirtyState
            & $RefreshGrid -PreferredRow $null
            & $SetStatus -Message "Loaded $($state.AllRows.Count) materials." -Level 'Success'
        }
        finally {
            $state.Loading = $false
        }
    }

    $SaveDatabase = {
        $buildResult = & $ValidateAndBuildMaterials
        if (-not $buildResult.IsValid) {
            $message = ($buildResult.Messages | Select-Object -Unique) -join [Environment]::NewLine
            & $SetStatus -Message 'Validation failed. Fix the grid before saving.' -Level 'Error'
            [System.Windows.MessageBox]::Show($message, 'Validation error', 'OK', 'Warning') | Out-Null
            return $false
        }

        try {
            $backupPath = Backup-DatabaseFile -Path $Script:DbPath
            Save-DatabaseFile -Path $Script:DbPath -Materials $buildResult.Materials
            $state.DatabaseDirty = $false
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
        if (-not $state.DatabaseDirty) {
            return $true
        }

        $result = [System.Windows.MessageBox]::Show('There are unsaved changes. Save before closing?', 'Unsaved changes', 'YesNoCancel', 'Warning')
        switch ($result) {
            'Yes' { return (& $SaveDatabase) }
            'No' { return $true }
            default { return $false }
        }
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
            <TextBlock Margin="0,6,0,0" Foreground="#64748B" Text="Combine rules with AND. Text supports contains, equals, and starts with."/>
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
            foreach ($width in @(240, 140, '*', 90)) {
                $columnDefinition = New-Object System.Windows.Controls.ColumnDefinition
                if ($width -eq '*') {
                    $columnDefinition.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
                }
                else {
                    $columnDefinition.Width = [System.Windows.GridLength]::new([double]$width)
                }
                [void]$rowGrid.ColumnDefinitions.Add($columnDefinition)
            }

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
            <TextBlock Margin="0,6,0,0" Foreground="#64748B" Text="Choose which columns are shown in the grid for this session."/>
        </StackPanel>
        <ScrollViewer Grid.Row="1" Margin="0,16,0,0" MaxHeight="520" VerticalScrollBarVisibility="Auto">
            <StackPanel x:Name="pnlColumns"/>
        </ScrollViewer>
        <DockPanel Grid.Row="2" Margin="0,16,0,0" LastChildFill="False">
            <Button x:Name="btnDefaultColumns" DockPanel.Dock="Left" Width="120" Content="Defaults"/>
            <Button x:Name="btnAllColumns" DockPanel.Dock="Left" Width="120" Margin="10,0,0,0" Content="Select all"/>
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
        $btnAllColumns = $dialog.FindName('btnAllColumns')
        $btnCancelColumns = $dialog.FindName('btnCancelColumns')
        $btnApplyColumns = $dialog.FindName('btnApplyColumns')
        $checkBoxStates = New-Object System.Collections.Generic.List[object]

        foreach ($columnDefinition in @($columnDefinitions)) {
            $checkBox = New-Object System.Windows.Controls.CheckBox
            $checkBox.Content = $columnDefinition.Label
            $checkBox.Margin = '0,0,0,8'
            $checkBox.IsChecked = ($state.VisibleColumnKeys -contains $columnDefinition.Key)
            [void]$pnlColumns.Children.Add($checkBox)
            [void]$checkBoxStates.Add([pscustomobject]@{ Key = $columnDefinition.Key; DefaultVisible = $columnDefinition.DefaultVisible; CheckBox = $checkBox })
        }

        $btnDefaultColumns.Add_Click({
                foreach ($entry in @($checkBoxStates.ToArray())) {
                    $entry.CheckBox.IsChecked = [bool]$entry.DefaultVisible
                }
            })
        $btnAllColumns.Add_Click({
                foreach ($entry in @($checkBoxStates.ToArray())) {
                    $entry.CheckBox.IsChecked = $true
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
                & $ApplyColumnVisibility
                $dialog.DialogResult = $true
                $dialog.Close()
            })

        return [bool]$dialog.ShowDialog()
    }

    $ExportCurrentGrid = {
        param([ValidateSet('csv', 'xlsx')][string]$Format)

        $visibleRows = @($dgGridViewer.ItemsSource)
        if (@($visibleRows).Count -eq 0) {
            [System.Windows.MessageBox]::Show('There are no visible rows to export.', 'Export', 'OK', 'Information') | Out-Null
            return
        }

        $visibleColumns = @($columnDefinitions | Where-Object { $state.VisibleColumnKeys -contains $_.Key })
        if (@($visibleColumns).Count -eq 0) {
            $visibleColumns = @($columnDefinitions)
        }

        if ($Format -eq 'csv') {
            $dialog = New-Object System.Windows.Forms.SaveFileDialog
            $dialog.Filter = 'CSV-Dateien (*.csv)|*.csv'
            $dialog.FileName = 'db_export.csv'
            if ($dialog.ShowDialog() -ne 'OK') {
                return
            }

            try {
                $nestedPath = Export-GridRowsToCsv -Path $dialog.FileName -Rows $visibleRows -ColumnDefinitions $visibleColumns -IncludeNestedData:([bool]$chkExportNested.IsChecked)
                if ($null -ne $nestedPath) {
                    & $SetStatus -Message "CSV export completed. Nested data: $(Split-Path $nestedPath -Leaf)" -Level 'Success'
                }
                else {
                    & $SetStatus -Message 'CSV export completed.' -Level 'Success'
                }
            }
            catch {
                & $SetStatus -Message "CSV export failed: $($_.Exception.Message)" -Level 'Error'
                [System.Windows.MessageBox]::Show($_.Exception.Message, 'CSV export error', 'OK', 'Error') | Out-Null
            }
            return
        }

        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Filter = 'Excel-Dateien (*.xlsx)|*.xlsx'
        $dialog.FileName = 'db_export.xlsx'
        if ($dialog.ShowDialog() -ne 'OK') {
            return
        }

        try {
            Export-GridRowsToXlsx -Path $dialog.FileName -Rows $visibleRows -ColumnDefinitions $visibleColumns -IncludeNestedData:([bool]$chkExportNested.IsChecked)
            & $SetStatus -Message 'XLSX export completed.' -Level 'Success'
        }
        catch {
            & $SetStatus -Message "XLSX export failed: $($_.Exception.Message)" -Level 'Error'
            [System.Windows.MessageBox]::Show($_.Exception.Message, 'XLSX export error', 'OK', 'Error') | Out-Null
        }
    }

    & $BuildGridColumns
    & $ApplyColumnVisibility
    & $UpdateActionState

    $txtSearchTimer = New-Object System.Windows.Threading.DispatcherTimer
    $txtSearchTimer.Interval = [TimeSpan]::FromMilliseconds(180)
    $txtSearchTimer.Add_Tick({
            $txtSearchTimer.Stop()
            & $RefreshGrid -PreferredRow $state.SelectedRow
            & $SetStatus -Message 'Search updated.' -Level 'Info'
        })

    $txtSearch.Add_TextChanged({
            $txtSearchTimer.Stop()
            $txtSearchTimer.Start()
        })

    $btnClearSearch.Add_Click({
            $txtSearch.Text = ''
            & $RefreshGrid -PreferredRow $state.SelectedRow
        })

    $btnOpenFilterMenu.Add_Click({
            if (& $OpenFilterDialog) {
                & $RefreshGrid -PreferredRow $state.SelectedRow
                & $SetStatus -Message 'Advanced filters updated.' -Level 'Info'
            }
        })

    $btnOpenColumnMenu.Add_Click({
            if (& $OpenColumnDialog) {
                & $SetStatus -Message 'Visible columns updated.' -Level 'Info'
            }
        })

    foreach ($filterControl in @($chkFilterHazardous, $chkFilterDecentral)) {
        $filterControl.Add_Checked({
                $state.FilterHazardousOnly = [bool]$chkFilterHazardous.IsChecked
                $state.FilterDecentralOnly = [bool]$chkFilterDecentral.IsChecked
                & $RefreshGrid -PreferredRow $state.SelectedRow
                & $SetStatus -Message 'List filters updated.' -Level 'Info'
            })
        $filterControl.Add_Unchecked({
                $state.FilterHazardousOnly = [bool]$chkFilterHazardous.IsChecked
                $state.FilterDecentralOnly = [bool]$chkFilterDecentral.IsChecked
                & $RefreshGrid -PreferredRow $state.SelectedRow
                & $SetStatus -Message 'List filters updated.' -Level 'Info'
            })
    }

    $dgGridViewer.Add_SelectionChanged({
            $state.SelectedRow = $dgGridViewer.SelectedItem
            & $UpdateActionState
        })

    $dgGridViewer.Add_CellEditEnding({
            if (-not $state.Loading) {
                & $MarkDirty
                & $SetStatus -Message 'Grid row updated in memory.' -Level 'Info'
            }
        })

    $btnNewMaterial.Add_Click({
            $existingMaterials = @()
            foreach ($row in @($state.AllRows.ToArray())) {
                $buildResult = Convert-GridRowToMaterialBuildResult -Row $row -ColumnDefinitions $columnDefinitions -DefaultUnitCode $defaultUnitCode
                $existingMaterials += $buildResult.Candidate
            }

            $newMaterial = New-DefaultMaterial -Id (Get-NextMaterialId -Materials $existingMaterials) -DefaultIdentifierType $defaultIdentifierType -DefaultUnitCode $defaultUnitCode
            $newRow = Convert-MaterialToGridRow -Material $newMaterial
            [void]$state.AllRows.Add($newRow)
            $txtSearch.Text = ''
            & $RefreshGrid -PreferredRow $newRow
            & $MarkDirty
            & $SetStatus -Message 'New material created in memory.' -Level 'Success'
        })

    $btnCloneMaterial.Add_Click({
            if ($null -eq $state.SelectedRow) {
                return
            }

            $currentBuildResult = Convert-GridRowToMaterialBuildResult -Row $state.SelectedRow -ColumnDefinitions $columnDefinitions -DefaultUnitCode $defaultUnitCode
            $existingMaterials = @()
            foreach ($row in @($state.AllRows.ToArray())) {
                $buildResult = Convert-GridRowToMaterialBuildResult -Row $row -ColumnDefinitions $columnDefinitions -DefaultUnitCode $defaultUnitCode
                $existingMaterials += $buildResult.Candidate
            }

            $sourceId = [int]$currentBuildResult.Candidate.id
            $cloneId = Get-NextMaterialId -Materials $existingMaterials
            $clone = ConvertTo-NormalizedMaterial -Material (Copy-DeepObject $currentBuildResult.Candidate) -DefaultIdentifierType $defaultIdentifierType -DefaultUnitCode $defaultUnitCode
            $clone.id = $cloneId
            $clone.primary_identifier.type = 'matnr'
            $clone.primary_identifier.value = Get-UniqueCloneIdentifierValue -BaseValue $clone.primary_identifier.value -IdentifierType 'matnr' -Materials $existingMaterials -SuggestedId $cloneId
            $clone.identifiers.matnr = $clone.primary_identifier.value
            $clone.canonical_key = Get-CanonicalKey -Type 'matnr' -Value $clone.primary_identifier.value
            $clone.texts.short_description = if ([string]::IsNullOrWhiteSpace((Get-NormalizedString $clone.texts.short_description))) { 'Copy' } else { "$(Get-NormalizedString $clone.texts.short_description) (Copy)" }

            $cloneRow = Convert-MaterialToGridRow -Material $clone
            [void]$state.AllRows.Add($cloneRow)
            $txtSearch.Text = ''
            & $RefreshGrid -PreferredRow $cloneRow
            & $MarkDirty
            & $SetStatus -Message "Material #$sourceId duplicated as #$cloneId." -Level 'Success'
        })

    $btnDeleteMaterial.Add_Click({
            if ($null -eq $state.SelectedRow) {
                return
            }

            $result = [System.Windows.MessageBox]::Show("Delete material #$($state.SelectedRow.ImportId)? This removes it from the JSON on next save.", 'Delete material', 'YesNo', 'Warning')
            if ($result -ne 'Yes') {
                return
            }

            $currentIndex = $state.AllRows.IndexOf($state.SelectedRow)
            [void]$state.AllRows.Remove($state.SelectedRow)
            $nextRow = $null
            if ($state.AllRows.Count -gt 0) {
                if ($currentIndex -ge $state.AllRows.Count) {
                    $currentIndex = $state.AllRows.Count - 1
                }
                $nextRow = $state.AllRows[$currentIndex]
            }

            & $RefreshGrid -PreferredRow $nextRow
            & $MarkDirty
            & $SetStatus -Message 'Material deleted in memory.' -Level 'Warning'
        })

    $btnEditAlternates.Add_Click({
            if ($null -eq $state.SelectedRow) {
                return
            }

            if (& $OpenAlternateEditor -Row $state.SelectedRow) {
                & $RefreshGrid -PreferredRow $state.SelectedRow
                & $SetStatus -Message 'Alternates updated in memory.' -Level 'Success'
            }
        })

    $btnReloadDatabase.Add_Click({
            if ($state.DatabaseDirty) {
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
    $btnExportCsv.Add_Click({ & $ExportCurrentGrid -Format 'csv' })
    $btnExportXlsx.Add_Click({ & $ExportCurrentGrid -Format 'xlsx' })

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

            if ($windowEventArgs.Key -eq [System.Windows.Input.Key]::Delete -and $dgGridViewer.IsKeyboardFocusWithin) {
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
    Start-GridViewerUi
}
