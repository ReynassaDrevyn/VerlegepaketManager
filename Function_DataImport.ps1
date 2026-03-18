# Function_DataImport.ps1
# Windows-PowerShell-5.1-kompatible Importversion

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$ProjectRoot = $(if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path $MyInvocation.MyCommand.Path -Parent })
$DbPath = Join-Path $ProjectRoot 'Core\db_verlegepaket.json'
$LogsDir = Join-Path $ProjectRoot 'Logs'
$BackupDir = Join-Path $LogsDir 'Backups'
$LogFile = Join-Path $LogsDir "InitialImport_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

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

function Get-DerivedNatoStockNumber {
    param(
        [string]$MatnrMain,
        [string]$SupplyNumber
    )

    $trimmedMatnr = Get-NormalizedString $MatnrMain
    $trimmedSupply = Get-NormalizedString $SupplyNumber

    if ([string]::IsNullOrWhiteSpace($trimmedMatnr) -or [string]::IsNullOrWhiteSpace($trimmedSupply)) {
        return ''
    }

    $supplyDigits = $trimmedSupply -replace '[^0-9]', ''
    if ($trimmedMatnr -match '^\d+$' -and $supplyDigits -and $trimmedMatnr -eq $supplyDigits) {
        return $trimmedMatnr
    }

    return ''
}

function ConvertTo-MaterialArray {
    param([AllowNull()]$InputObject)

    if ($null -eq $InputObject) {
        return @()
    }

    if ($InputObject -is [System.Array]) {
        return @($InputObject)
    }

    if ($InputObject.PSObject.Properties['material_class']) {
        return @($InputObject)
    }

    return @($InputObject)
}

function Get-ExistingMaterialLookup {
    param([Parameter(Mandatory = $true)][object[]]$Materials)

    $lookup = @{}
    for ($i = 0; $i -lt $Materials.Count; $i++) {
        $matnr = Get-NormalizedString $Materials[$i].material_class.matnr_main
        if (-not [string]::IsNullOrWhiteSpace($matnr) -and -not $lookup.ContainsKey($matnr)) {
            $lookup[$matnr] = $i
        }
    }

    return $lookup
}

function Get-NextMaterialId {
    param([Parameter(Mandatory = $true)][object[]]$Materials)

    $maxId = 999
    foreach ($material in $Materials) {
        $idText = Get-NormalizedString $material.material_class.id
        $idValue = 0
        if ([int]::TryParse($idText, [ref]$idValue) -and $idValue -gt $maxId) {
            $maxId = $idValue
        }
    }

    return $maxId
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
    $materials = ConvertTo-MaterialArray $parsed

    return [pscustomobject]@{
        Materials = $materials
        Lookup    = Get-ExistingMaterialLookup -Materials $materials
        MaxId     = Get-NextMaterialId -Materials $materials
    }
}

function New-MaterialRecord {
    param(
        [int]$Id,
        [string]$MatnrMain,
        [string]$Description,
        [string]$NatoStockNumber,
        [string]$SupplyNumber,
        [string]$MatStatMain,
        [string]$ExtWg,
        [bool]$Dezentral,
        [string]$ArtNr,
        [string]$Creditor,
        [bool]$IsDg,
        [string]$UnNum,
        [object[]]$DangerousTags,
        [string]$UnitMain,
        [double]$QuantityTarget,
        [object[]]$AltUnits,
        [object[]]$AltMaterial,
        [object[]]$WtgWaStff,
        [object[]]$InstElo,
        [string]$Logistics,
        [string]$Technical,
        [object[]]$MiscTags
    )

    return [pscustomobject][ordered]@{
        material_class = [pscustomobject][ordered]@{
            id                = $Id
            matnr_main        = $MatnrMain
            description       = $Description
            nato_stock_number = $NatoStockNumber
            supplynumber      = $SupplyNumber
            mat_stat_main     = $MatStatMain
            properties        = [pscustomobject][ordered]@{
                productgroup   = [pscustomobject][ordered]@{
                    ext_wg    = $ExtWg
                    dezentral = $Dezentral
                    artnr     = $ArtNr
                    creditor  = $Creditor
                }
                dangerous_good = [pscustomobject][ordered]@{
                    is_dg  = $IsDg
                    un_num = $UnNum
                    tags   = @($DangerousTags)
                }
                quantity       = [pscustomobject][ordered]@{
                    unit_main       = $UnitMain
                    quantity_target = $QuantityTarget
                    alt_units       = @($AltUnits)
                }
            }
            alt_material      = @($AltMaterial)
            mat_ref           = [pscustomobject][ordered]@{
                WtgWaStff  = @($WtgWaStff)
                'Inst/Elo' = @($InstElo)
            }
            comments          = [pscustomobject][ordered]@{
                logistics = $Logistics
                technical = $Technical
            }
            misc              = [pscustomobject][ordered]@{
                tags = @($MiscTags)
            }
        }
    }
}

function Get-ImportedTagValues {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][hashtable]$HeaderMap,
        [Parameter(Mandatory = $true)][string[]]$Keys
    )

    $tags = New-Object System.Collections.Generic.List[string]
    foreach ($key in $Keys) {
        if (Test-Flag -Row $Row -HeaderMap $HeaderMap -Key $key) {
            [void]$tags.Add($key)
        }
    }

    return @($tags)
}

function Get-ExistingTagSet {
    param(
        $ExistingMaterial,
        [Parameter(Mandatory = $true)][string]$Category
    )

    if (-not $ExistingMaterial) {
        return @()
    }

    switch ($Category) {
        'DangerousTags' { return @($ExistingMaterial.material_class.properties.dangerous_good.tags) }
        'WtgWaStff' { return @($ExistingMaterial.material_class.mat_ref.WtgWaStff) }
        'InstElo' { return @($ExistingMaterial.material_class.mat_ref.'Inst/Elo') }
        'MiscTags' { return @($ExistingMaterial.material_class.misc.tags) }
        default { return @() }
    }
}

function Merge-TagValues {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][hashtable]$HeaderMap,
        [Parameter(Mandatory = $true)][string[]]$Keys,
        [Parameter(Mandatory = $true)][string]$Category,
        $ExistingMaterial = $null
    )

    $existingSet = @{}
    foreach ($tag in (Get-ExistingTagSet -ExistingMaterial $ExistingMaterial -Category $Category)) {
        $existingSet[(Get-NormalizedString $tag)] = $true
    }

    $merged = New-Object System.Collections.Generic.List[string]
    foreach ($key in $Keys) {
        if (Test-HeaderAvailable -HeaderMap $HeaderMap -Key $key) {
            if (Test-Flag -Row $Row -HeaderMap $HeaderMap -Key $key) {
                [void]$merged.Add($key)
            }
        }
        elseif ($existingSet.ContainsKey($key)) {
            [void]$merged.Add($key)
        }
    }

    return @($merged)
}

function Get-ImportedScalarValue {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][hashtable]$HeaderMap,
        [Parameter(Mandatory = $true)][string]$Key,
        $ExistingMaterial = $null,
        [string]$ExistingValue = ''
    )

    if (Test-HeaderAvailable -HeaderMap $HeaderMap -Key $Key) {
        return Get-CellValue -Row $Row -HeaderMap $HeaderMap -Key $Key
    }

    if ($ExistingMaterial) {
        return Get-NormalizedString $ExistingValue
    }

    return ''
}

function Convert-RowToImportData {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][hashtable]$HeaderMap,
        [Parameter(Mandatory = $true)][int]$RowNumber,
        $ExistingMaterial = $null
    )

    $warnings = New-Object System.Collections.Generic.List[string]

    $matnr = Get-CellValue -Row $Row -HeaderMap $HeaderMap -Key 'matnr_main'
    if ([string]::IsNullOrWhiteSpace($matnr)) {
        return [pscustomobject]@{
            ShouldSkip = $true
            Warnings   = @("Zeile $RowNumber - matnr_main leer, uebersprungen")
        }
    }

    $existingQuantityTarget = 0.0
    if ($ExistingMaterial) {
        $existingQuantityTarget = [double]$ExistingMaterial.material_class.properties.quantity.quantity_target
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
                [void]$warnings.Add("Zeile $RowNumber - Menge '$qtyRaw' ungueltig, Standardwert 0 verwendet")
            }
        }
    }
    else {
        $quantityTarget = $existingQuantityTarget
    }

    $existingDezentral = $false
    if ($ExistingMaterial) {
        $existingDezentral = [bool]$ExistingMaterial.material_class.properties.productgroup.dezentral
    }

    if (Test-HeaderAvailable -HeaderMap $HeaderMap -Key 'dezentral') {
        $dezentRaw = Get-CellValue -Row $Row -HeaderMap $HeaderMap -Key 'dezentral'
        $dezentParse = ConvertTo-ImportBooleanParseResult $dezentRaw
        $dezentralValue = $false
        if ($dezentParse.Success) {
            $dezentralValue = [bool]$dezentParse.Value
        }
        elseif ($ExistingMaterial) {
            $dezentralValue = $existingDezentral
            [void]$warnings.Add("Zeile $RowNumber - Dezentralwert '$dezentRaw' ungueltig, vorhandenen Wert beibehalten")
        }
        else {
            [void]$warnings.Add("Zeile $RowNumber - Dezentralwert '$dezentRaw' ungueltig, Standardwert false verwendet")
        }
    }
    else {
        $dezentralValue = $existingDezentral
    }

    $existingIsDg = $false
    if ($ExistingMaterial) {
        $existingIsDg = [bool]$ExistingMaterial.material_class.properties.dangerous_good.is_dg
    }

    if (Test-HeaderAvailable -HeaderMap $HeaderMap -Key 'is_dg') {
        $isDgRaw = Get-CellValue -Row $Row -HeaderMap $HeaderMap -Key 'is_dg'
        $isDgParse = ConvertTo-ImportBooleanParseResult $isDgRaw
        $explicitIsDg = $false
        if ($isDgParse.Success) {
            $explicitIsDg = [bool]$isDgParse.Value
        }
        elseif ($ExistingMaterial) {
            $explicitIsDg = $existingIsDg
            [void]$warnings.Add("Zeile $RowNumber - Gefahrstoffwert '$isDgRaw' ungueltig, vorhandenen Wert beibehalten")
        }
        else {
            [void]$warnings.Add("Zeile $RowNumber - Gefahrstoffwert '$isDgRaw' ungueltig, Standardwert false verwendet")
        }
    }
    else {
        $explicitIsDg = $existingIsDg
    }

    $dangerousTags = Merge-TagValues -Row $Row -HeaderMap $HeaderMap -Keys @('GefStoff Verlegung', 'Gefahrgut', 'Batterie') -Category 'DangerousTags' -ExistingMaterial $ExistingMaterial
    $wtgWaStff = Merge-TagValues -Row $Row -HeaderMap $HeaderMap -Keys @('Flight', 'Waffen', 'Munition') -Category 'WtgWaStff' -ExistingMaterial $ExistingMaterial
    $instElo = Merge-TagValues -Row $Row -HeaderMap $HeaderMap -Keys @('RTS', 'AUG', 'WEF', 'BoGe', 'HFT', 'LME', 'REG', 'RNW', 'Rad Reifen Shop') -Category 'InstElo' -ExistingMaterial $ExistingMaterial
    $miscTags = Merge-TagValues -Row $Row -HeaderMap $HeaderMap -Keys @('IETPX Material', 'GUN', 'GUN ON AC', 'GUN OFF AC', 'IRIS-T', 'FLARE', 'AIM 120', '1000 l SFT', 'GBU 48', 'Meteor', 'LDP', 'IWP', 'CFP', 'MFRL', 'OWP', 'CHAFF', 'MEL', 'ITSPL') -Category 'MiscTags' -ExistingMaterial $ExistingMaterial

    $matStatValue = 'XX'
    if ($ExistingMaterial) {
        $matStatValue = Get-NormalizedString $ExistingMaterial.material_class.mat_stat_main
        if ([string]::IsNullOrWhiteSpace($matStatValue)) {
            $matStatValue = 'XX'
        }
    }

    if (Test-HeaderAvailable -HeaderMap $HeaderMap -Key 'mat_stat_main') {
        $matStatRaw = Get-CellValue -Row $Row -HeaderMap $HeaderMap -Key 'mat_stat_main'
        if (-not [string]::IsNullOrWhiteSpace($matStatRaw)) {
            $trimmedMatStat = $matStatRaw.Trim()
            if ($trimmedMatStat.Length -eq 2) {
                $matStatValue = $trimmedMatStat
            }
            else {
                $matStatValue = if ($ExistingMaterial) { $matStatValue } else { 'XX' }
                [void]$warnings.Add("Zeile $RowNumber - MatStatus '$matStatRaw' ungueltig, Wert beibehalten")
            }
        }
        elseif (-not $ExistingMaterial) {
            $matStatValue = 'XX'
        }
    }

    $existingSupplyNumber = ''
    if ($ExistingMaterial) {
        $existingSupplyNumber = Get-NormalizedString $ExistingMaterial.material_class.supplynumber
    }
    $supplyNumber = Get-ImportedScalarValue -Row $Row -HeaderMap $HeaderMap -Key 'supplynumber' -ExistingMaterial $ExistingMaterial -ExistingValue $existingSupplyNumber
    $derivedNato = Get-DerivedNatoStockNumber -MatnrMain $matnr -SupplyNumber $supplyNumber
    if ([string]::IsNullOrWhiteSpace($derivedNato) -and $ExistingMaterial) {
        $derivedNato = Get-NormalizedString $ExistingMaterial.material_class.nato_stock_number
    }

    return [pscustomobject]@{
        ShouldSkip      = $false
        Warnings        = @($warnings)
        MatnrMain       = $matnr
        Description     = Get-ImportedScalarValue -Row $Row -HeaderMap $HeaderMap -Key 'description' -ExistingMaterial $ExistingMaterial -ExistingValue $ExistingMaterial.material_class.description
        NatoStockNumber = $derivedNato
        SupplyNumber    = $supplyNumber
        MatStatMain     = $matStatValue
        ExtWg           = Get-ImportedScalarValue -Row $Row -HeaderMap $HeaderMap -Key 'ext_wg' -ExistingMaterial $ExistingMaterial -ExistingValue $ExistingMaterial.material_class.properties.productgroup.ext_wg
        Dezentral       = $dezentralValue
        ArtNr           = Get-ImportedScalarValue -Row $Row -HeaderMap $HeaderMap -Key 'artnr' -ExistingMaterial $ExistingMaterial -ExistingValue $ExistingMaterial.material_class.properties.productgroup.artnr
        IsDg            = ($explicitIsDg -or $dangerousTags.Count -gt 0)
        DangerousTags   = $dangerousTags
        UnitMain        = Get-ImportedScalarValue -Row $Row -HeaderMap $HeaderMap -Key 'unit_main' -ExistingMaterial $ExistingMaterial -ExistingValue $ExistingMaterial.material_class.properties.quantity.unit_main
        QuantityTarget  = [double]$quantityTarget
        WtgWaStff       = $wtgWaStff
        InstElo         = $instElo
        Logistics       = Get-ImportedScalarValue -Row $Row -HeaderMap $HeaderMap -Key 'logistics' -ExistingMaterial $ExistingMaterial -ExistingValue $ExistingMaterial.material_class.comments.logistics
        Technical       = Get-ImportedScalarValue -Row $Row -HeaderMap $HeaderMap -Key 'technical' -ExistingMaterial $ExistingMaterial -ExistingValue $ExistingMaterial.material_class.comments.technical
        MiscTags        = $miscTags
    }
}

function Merge-MaterialRecord {
    param(
        [Parameter(Mandatory = $true)]$ImportData,
        $ExistingMaterial = $null,
        [Parameter(Mandatory = $true)][int]$Id
    )

    $creditor = ''
    $unNum = ''
    $altUnits = @()
    $altMaterial = @()

    if ($ExistingMaterial) {
        $creditor = Get-NormalizedString $ExistingMaterial.material_class.properties.productgroup.creditor
        $unNum = Get-NormalizedString $ExistingMaterial.material_class.properties.dangerous_good.un_num
        $altUnits = @($ExistingMaterial.material_class.properties.quantity.alt_units)
        $altMaterial = @($ExistingMaterial.material_class.alt_material)
    }

    return New-MaterialRecord `
        -Id $Id `
        -MatnrMain $ImportData.MatnrMain `
        -Description $ImportData.Description `
        -NatoStockNumber $ImportData.NatoStockNumber `
        -SupplyNumber $ImportData.SupplyNumber `
        -MatStatMain $ImportData.MatStatMain `
        -ExtWg $ImportData.ExtWg `
        -Dezentral ([bool]$ImportData.Dezentral) `
        -ArtNr $ImportData.ArtNr `
        -Creditor $creditor `
        -IsDg ([bool]$ImportData.IsDg) `
        -UnNum $unNum `
        -DangerousTags $ImportData.DangerousTags `
        -UnitMain $ImportData.UnitMain `
        -QuantityTarget ([double]$ImportData.QuantityTarget) `
        -AltUnits $altUnits `
        -AltMaterial $altMaterial `
        -WtgWaStff $ImportData.WtgWaStff `
        -InstElo $ImportData.InstElo `
        -Logistics $ImportData.Logistics `
        -Technical $ImportData.Technical `
        -MiscTags $ImportData.MiscTags
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
        $matnr = Get-CellValue -Row $Rows[$index] -HeaderMap $HeaderMap -Key 'matnr_main'
        if ([string]::IsNullOrWhiteSpace($matnr)) {
            continue
        }

        if ($seen.ContainsKey($matnr)) {
            [void]$duplicates.Add("matnr_main '$matnr' in Zeile $($seen[$matnr]) und Zeile $rowNumber")
        }
        else {
            $seen[$matnr] = $rowNumber
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

    $Materials | ConvertTo-Json -Depth 10 | Out-File -FilePath $Path -Encoding UTF8
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
        Write-ImportLog 'Import abgebrochen - doppelte matnr_main-Werte im CSV' 'ERROR'
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

    $nextId = [int]$existingDb.MaxId
    $insertedCount = 0
    $updatedCount = 0
    $skippedCount = 0
    $warningCount = 0
    $errorCount = 0

    for ($index = 0; $index -lt $csvData.Rows.Count; $index++) {
        $rowNumber = $index + 2
        $row = $csvData.Rows[$index]
        $matnr = Get-CellValue -Row $row -HeaderMap $headerResolution.HeaderMap -Key 'matnr_main'
        $existingMaterial = $null
        if ($lookup.ContainsKey($matnr)) {
            $existingMaterial = $materials[$lookup[$matnr]]
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
            $id = [int]$existingMaterial.material_class.id
            $mergedRecord = Merge-MaterialRecord -ImportData $importData -ExistingMaterial $existingMaterial -Id $id
            $materials[$lookup[$matnr]] = $mergedRecord
            $updatedCount++
            Write-ImportLog "Zeile $rowNumber - aktualisiert: $matnr (ID $id)" 'INFO'
        }
        else {
            $nextId++
            $newRecord = Merge-MaterialRecord -ImportData $importData -Id $nextId
            [void]$materials.Add($newRecord)
            $lookup[$matnr] = $materials.Count - 1
            $insertedCount++
            Write-ImportLog "Zeile $rowNumber - importiert: $matnr (ID $nextId)" 'INFO'
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

function Start-ImportToolUi {
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Initial-Import Verlegepaket (PS 5.1)"
        Height="700"
        Width="900"
        WindowStartupLocation="CenterScreen"
        Background="#F5F5F5"
        FontFamily="Segoe UI"
        FontSize="11">
    <Grid>
        <StackPanel Margin="20">
            <TextBlock Text="Verlegepaket Datenimport"
                       FontSize="24"
                       FontWeight="Bold"
                       Foreground="#2C3E50"
                       Margin="0,0,0,10"/>

            <Border BorderThickness="1" BorderBrush="#E0E0E0" CornerRadius="5" Padding="15" Background="White" Margin="0,0,0,15">
                <StackPanel>
                    <TextBlock Text="Quelldatei" FontWeight="Bold" Foreground="#34495E"/>
                    <Grid Margin="0,10,0,0">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="110"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="txtFile"
                                 Grid.Column="0"
                                 Padding="10"
                                 Background="White"
                                 BorderThickness="1"
                                 BorderBrush="#BDC3C7"
                                 IsReadOnly="True"
                                 Foreground="#7F8C8D"/>
                        <Button Content="Durchsuchen..."
                                Grid.Column="1"
                                Margin="10,0,0,0"
                                x:Name="btnBrowse"
                                Background="#3498DB"
                                Foreground="White"
                                FontWeight="Bold"
                                Cursor="Hand"
                                Padding="10"/>
                    </Grid>
                </StackPanel>
            </Border>

            <Button Content="Import starten"
                    x:Name="btnImport"
                    Background="#27AE60"
                    Foreground="White"
                    FontWeight="Bold"
                    FontSize="13"
                    Padding="15,12"
                    Height="45"
                    Cursor="Hand"
                    Margin="0,0,0,15"/>

            <Border BorderThickness="1" BorderBrush="#E0E0E0" CornerRadius="5" Padding="15" Background="White">
                <StackPanel>
                    <TextBlock Text="Importprotokoll" FontWeight="Bold" Foreground="#34495E"/>
                    <TextBox x:Name="txtLog"
                             Height="450"
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
        </StackPanel>
    </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader([xml]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $txtFile = $window.FindName('txtFile')
    $btnBrowse = $window.FindName('btnBrowse')
    $btnImport = $window.FindName('btnImport')
    $txtLog = $window.FindName('txtLog')

    $global:LogBox = $txtLog

    $btnBrowse.Add_Click({
            $ofd = New-Object System.Windows.Forms.OpenFileDialog
            $ofd.Filter = 'CSV/Text-Dateien (*.csv;*.txt)|*.csv;*.txt|Alle Dateien (*.*)|*.*'
            if ($ofd.ShowDialog() -eq 'OK') {
                $txtFile.Text = $ofd.FileName
            }
        })

    $btnImport.Add_Click({
            if ([string]::IsNullOrWhiteSpace($txtFile.Text)) {
                [System.Windows.MessageBox]::Show('Bitte eine Datei auswaehlen!', 'Hinweis', 'OK', 'Warning')
                return
            }

            $btnImport.IsEnabled = $false
            Start-InitialImport -SourceFile $txtFile.Text
            $btnImport.IsEnabled = $true
        })

    Write-ImportLog "Tool gestartet - $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')" 'INFO'
    Write-ImportLog "Logdatei: $LogFile" 'INFO'
    Write-ImportLog "Datenbank: $DbPath" 'INFO'

    $window.ShowDialog() | Out-Null
}

if ($MyInvocation.InvocationName -ne '.') {
    Start-ImportToolUi
}
