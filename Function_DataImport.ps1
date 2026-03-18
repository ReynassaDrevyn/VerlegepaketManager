# Function_DataImport.ps1
# Windows-PowerShell-5.1-kompatible Importversion

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Konfiguration
$ProjectRoot = $(if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path $MyInvocation.MyCommand.Path -Parent })
$DbPath = Join-Path $ProjectRoot 'Core\db_verlegepaket.json'
$LogsDir = Join-Path $ProjectRoot 'Logs'
$BackupDir = Join-Path $LogsDir 'Backups'
$LogFile = Join-Path $LogsDir "InitialImport_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Ordner anlegen
if (!(Test-Path $LogsDir)) { New-Item -Path $LogsDir -ItemType Directory -Force | Out-Null }
if (!(Test-Path $BackupDir)) { New-Item -Path $BackupDir -ItemType Directory -Force | Out-Null }
$dbDir = Split-Path $DbPath -Parent
if (!(Test-Path $dbDir)) { New-Item -Path $dbDir -ItemType Directory -Force | Out-Null }

# Logging
$global:LogBox = $null

function Write-ImportLog {
    param([string]$Message, [string]$Level = 'INFO')

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logLine = "$timestamp [$Level] $Message"

    Add-Content -Path $LogFile -Value $logLine -Encoding UTF8 -ErrorAction SilentlyContinue
    if ($global:LogBox) {
        $global:LogBox.AppendText("$logLine`r`n")
        $global:LogBox.ScrollToCaret()
    }
}

function Get-CellValue {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][hashtable]$IndexMap,
        [Parameter(Mandatory = $true)][string]$Key
    )

    if ($IndexMap.ContainsKey($Key)) {
        return [string]$Row.PSObject.Properties.Value[$IndexMap[$Key]]
    }

    return ''
}

function Test-Flag {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][hashtable]$IndexMap,
        [Parameter(Mandatory = $true)][string]$Key,
        [string]$Pattern = '^[xX1]$'
    )

    return (Get-CellValue -Row $Row -IndexMap $IndexMap -Key $Key) -match $Pattern
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
        $Dezentral,
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

# Hauptfunktion
function Start-InitialImport {
    param([string]$SourceFile)

    if (!(Test-Path $SourceFile)) {
        Write-ImportLog "Datei nicht gefunden: $SourceFile" 'ERROR'
        return
    }

    Write-ImportLog "Starte Import aus $SourceFile ..." 'INFO'

    if (Test-Path $DbPath) {
        $backup = "db_verlegepaket_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
        Copy-Item -Path $DbPath -Destination (Join-Path $BackupDir $backup) -Force
        Write-ImportLog "Backup erstellt: $backup" 'INFO'
    }

    $data = $null
    try {
        $data = Import-Csv -Path $SourceFile -Delimiter ';' -Encoding Default
        Write-ImportLog "CSV geladen - $($data.Count) Zeilen gefunden" 'INFO'
    }
    catch {
        Write-ImportLog "CSV-Ladefehler: $($_.Exception.Message)" 'ERROR'
        return
    }

    if ($null -eq $data -or $data.Count -eq 0) {
        Write-ImportLog 'Keine Datenzeilen' 'ERROR'
        return
    }

    $headers = $data[0].PSObject.Properties.Name
    Write-ImportLog "Header erkannt: $($headers -join ' | ')" 'INFO'

    $idx = @{}
    foreach ($col in $headers) {
        $clean = $col.Trim() -replace '[äöüÄÖÜß]', '?'

        if ($clean -match 'Materialnummer.*SASPF') { $idx['matnr_main'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^VersNr$') { $idx['supplynumber'] = [array]::IndexOf($headers, $col) }
        if ($clean -match 'Status.*MatNr') { $idx['mat_stat_main'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^Dezent$') { $idx['dezentral'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^Ext WG$') { $idx['ext_wg'] = [array]::IndexOf($headers, $col) }
        if ($clean -match 'Artikel Nr') { $idx['artnr'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^Materialbezeichnung$') { $idx['description'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^Bezeichnung Technik$') { $idx['technical'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^Bemerkung$') { $idx['logistics'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^BZE$') { $idx['unit_main'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^TLG 74$') { $idx['quantity_target'] = [array]::IndexOf($headers, $col) }

        if ($clean -match '^GefStoff Verlegung$') { $idx['GefStoff Verlegung'] = [array]::IndexOf($headers, $col) }
        elseif ($clean -match '^GefStoff$') { $idx['is_dg'] = [array]::IndexOf($headers, $col) }

        if ($clean -match '^Gefahrgut$') { $idx['Gefahrgut'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^Batterie$') { $idx['Batterie'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^Flight$') { $idx['Flight'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^Waffen$') { $idx['Waffen'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^Munition$') { $idx['Munition'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^RTS$') { $idx['RTS'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^AUG$') { $idx['AUG'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^WEF$') { $idx['WEF'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^BoGe$') { $idx['BoGe'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^HFT$') { $idx['HFT'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^LME$') { $idx['LME'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^REG$') { $idx['REG'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^RNW$') { $idx['RNW'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^Rad Reifen Shop$') { $idx['Rad Reifen Shop'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^IETPX Material$') { $idx['IETPX Material'] = [array]::IndexOf($headers, $col) }

        if ($clean -match '^GUN ON AC$') { $idx['GUN ON AC'] = [array]::IndexOf($headers, $col) }
        elseif ($clean -match '^GUN OFF AC$') { $idx['GUN OFF AC'] = [array]::IndexOf($headers, $col) }
        elseif ($clean -match '^GUN$') { $idx['GUN'] = [array]::IndexOf($headers, $col) }

        if ($clean -match '^IRIS-T$') { $idx['IRIS-T'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^FLARE$') { $idx['FLARE'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^AIM 120$') { $idx['AIM 120'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^1000 l SFT$') { $idx['1000 l SFT'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^GBU 48$') { $idx['GBU 48'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^Meteor$') { $idx['Meteor'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^LDP$') { $idx['LDP'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^IWP$') { $idx['IWP'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^CFP$') { $idx['CFP'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^MFRL$') { $idx['MFRL'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^OWP$') { $idx['OWP'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^CHAFF$') { $idx['CHAFF'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^MEL$') { $idx['MEL'] = [array]::IndexOf($headers, $col) }
        if ($clean -match '^ITSPL$') { $idx['ITSPL'] = [array]::IndexOf($headers, $col) }
    }

    Write-ImportLog "Gefundene Indizes: $($idx.Keys -join ', ')" 'DEBUG'

    if (-not $idx.ContainsKey('matnr_main')) {
        Write-ImportLog 'Kritischer Fehler: matnr_main nicht gefunden!' 'ERROR'
        return
    }

    $materials = @()
    $maxId = 999

    if (Test-Path $DbPath) {
        try {
            $oldDb = Get-Content -Path $DbPath -Encoding UTF8 | ConvertFrom-Json
            $maxId = ($oldDb | ForEach-Object { [int]$_.material_class.id } | Measure-Object -Maximum).Maximum
            if (!$maxId) { $maxId = 999 }
        }
        catch {
            Write-ImportLog 'Alte DB nicht lesbar - starte bei 1000' 'WARNING'
        }
    }

    $lineNum = 2
    foreach ($row in $data) {
        $lineNum++

        $matnr = Get-CellValue -Row $row -IndexMap $idx -Key 'matnr_main'
        if ([string]::IsNullOrWhiteSpace($matnr)) {
            Write-ImportLog "Zeile $lineNum - matnr_main leer, uebersprungen" 'WARNING'
            continue
        }

        $desc = Get-CellValue -Row $row -IndexMap $idx -Key 'description'
        $supply = Get-CellValue -Row $row -IndexMap $idx -Key 'supplynumber'
        $matStat = Get-CellValue -Row $row -IndexMap $idx -Key 'mat_stat_main'
        $dezentRaw = Get-CellValue -Row $row -IndexMap $idx -Key 'dezentral'
        $extWg = Get-CellValue -Row $row -IndexMap $idx -Key 'ext_wg'
        $artNr = Get-CellValue -Row $row -IndexMap $idx -Key 'artnr'
        $technical = Get-CellValue -Row $row -IndexMap $idx -Key 'technical'
        $logistics = Get-CellValue -Row $row -IndexMap $idx -Key 'logistics'
        $unitMain = Get-CellValue -Row $row -IndexMap $idx -Key 'unit_main'
        $qtyTargetRaw = Get-CellValue -Row $row -IndexMap $idx -Key 'quantity_target'
        $isDgRaw = Get-CellValue -Row $row -IndexMap $idx -Key 'is_dg'

        $qtyTarget = 0.0
        $parsed = 0.0
        $culture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
        if ([string]::IsNullOrWhiteSpace($qtyTargetRaw) -eq $false) {
            [void][double]::TryParse($qtyTargetRaw.Trim(), [System.Globalization.NumberStyles]::Any, $culture, [ref]$parsed)
            $qtyTarget = $parsed
        }

        $dangerousTags = @()
        if (Test-Flag -Row $row -IndexMap $idx -Key 'GefStoff Verlegung') { $dangerousTags += 'GefStoff Verlegung' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'Gefahrgut') { $dangerousTags += 'Gefahrgut' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'Batterie') { $dangerousTags += 'Batterie' }

        $wtgWaStff = @()
        if (Test-Flag -Row $row -IndexMap $idx -Key 'Flight' -Pattern '^[xX]$') { $wtgWaStff += 'Flight' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'Waffen' -Pattern '^[xX]$') { $wtgWaStff += 'Waffen' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'Munition' -Pattern '^[xX]$') { $wtgWaStff += 'Munition' }

        $instElo = @()
        if (Test-Flag -Row $row -IndexMap $idx -Key 'RTS' -Pattern '^[xX]$') { $instElo += 'RTS' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'AUG' -Pattern '^[xX]$') { $instElo += 'AUG' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'WEF' -Pattern '^[xX]$') { $instElo += 'WEF' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'BoGe' -Pattern '^[xX]$') { $instElo += 'BoGe' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'HFT' -Pattern '^[xX]$') { $instElo += 'HFT' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'LME' -Pattern '^[xX]$') { $instElo += 'LME' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'REG' -Pattern '^[xX]$') { $instElo += 'REG' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'RNW' -Pattern '^[xX]$') { $instElo += 'RNW' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'Rad Reifen Shop' -Pattern '^[xX]$') { $instElo += 'Rad Reifen Shop' }

        $miscTags = @()
        if (Test-Flag -Row $row -IndexMap $idx -Key 'IETPX Material') { $miscTags += 'IETPX Material' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'GUN') { $miscTags += 'GUN' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'GUN ON AC') { $miscTags += 'GUN ON AC' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'GUN OFF AC') { $miscTags += 'GUN OFF AC' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'IRIS-T') { $miscTags += 'IRIS-T' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'FLARE') { $miscTags += 'FLARE' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'AIM 120') { $miscTags += 'AIM 120' }
        if (Test-Flag -Row $row -IndexMap $idx -Key '1000 l SFT') { $miscTags += '1000 l SFT' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'GBU 48') { $miscTags += 'GBU 48' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'Meteor') { $miscTags += 'Meteor' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'LDP') { $miscTags += 'LDP' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'IWP') { $miscTags += 'IWP' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'CFP') { $miscTags += 'CFP' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'MFRL') { $miscTags += 'MFRL' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'OWP') { $miscTags += 'OWP' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'CHAFF') { $miscTags += 'CHAFF' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'MEL') { $miscTags += 'MEL' }
        if (Test-Flag -Row $row -IndexMap $idx -Key 'ITSPL') { $miscTags += 'ITSPL' }

        $matStatValue = 'XX'
        if (-not [string]::IsNullOrWhiteSpace($matStat)) {
            $trimmedMatStat = $matStat.Trim()
            if ($trimmedMatStat.Length -eq 2) {
                $matStatValue = $trimmedMatStat
            }
        }

        $dezentralValue = ''
        if (-not [string]::IsNullOrWhiteSpace($dezentRaw)) {
            $dezentCandidate = $dezentRaw.Trim().ToLowerInvariant()
            if ($dezentCandidate -eq 'true') {
                $dezentralValue = $true
            }
            elseif ($dezentCandidate -eq 'false') {
                $dezentralValue = $false
            }
        }

        $isDgValue = $false
        if (-not [string]::IsNullOrWhiteSpace($isDgRaw) -and $isDgRaw.Trim() -eq '1') {
            $isDgValue = $true
        }

        $natoStockNumber = ''
        $trimmedMatnr = $matnr.Trim()
        $supplyDigits = ''
        if (-not [string]::IsNullOrWhiteSpace($supply)) {
            $supplyDigits = ($supply -replace '[^0-9]', '')
        }
        if ($trimmedMatnr -match '^\d+$' -and $supplyDigits -and $trimmedMatnr -eq $supplyDigits) {
            $natoStockNumber = $trimmedMatnr
        }

        $obj = New-MaterialRecord `
            -Id (++$maxId) `
            -MatnrMain $matnr.TrimEnd() `
            -Description $desc `
            -NatoStockNumber $natoStockNumber `
            -SupplyNumber $supply `
            -MatStatMain $matStatValue `
            -ExtWg $extWg `
            -Dezentral $dezentralValue `
            -ArtNr $artNr `
            -Creditor '' `
            -IsDg $isDgValue `
            -UnNum '' `
            -DangerousTags $dangerousTags `
            -UnitMain $unitMain `
            -QuantityTarget $qtyTarget `
            -AltUnits @() `
            -AltMaterial @() `
            -WtgWaStff $wtgWaStff `
            -InstElo $instElo `
            -Logistics $logistics `
            -Technical $technical `
            -MiscTags $miscTags

        $materials += $obj
        Write-ImportLog "Zeile $lineNum - importiert: $matnr (ID $maxId)" 'INFO'
    }

    try {
        $materials | ConvertTo-Json -Depth 10 | Out-File -FilePath $DbPath -Encoding UTF8
        Write-ImportLog "Import abgeschlossen - $($materials.Count) Materialien gespeichert" 'SUCCESS'
        [System.Windows.Forms.MessageBox]::Show("Import abgeschlossen!`n$($materials.Count) Eintraege`nLog: $LogFile", 'Erfolg', 'OK', 'Information')
    }
    catch {
        Write-ImportLog "Fehler beim Speichern der JSON: $($_.Exception.Message)" 'ERROR'
    }
}

# GUI
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Initial-Import Verlegepaket (PS 5.1)'
$form.Size = New-Object System.Drawing.Size(780, 620)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedSingle'

$lblFile = New-Object System.Windows.Forms.Label
$lblFile.Text = 'Quelldatei:'
$lblFile.Location = New-Object System.Drawing.Point(12, 18)
$form.Controls.Add($lblFile)

$txtFile = New-Object System.Windows.Forms.TextBox
$txtFile.Location = New-Object System.Drawing.Point(100, 15)
$txtFile.Size = New-Object System.Drawing.Size(550, 24)
$txtFile.ReadOnly = $true
$form.Controls.Add($txtFile)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = 'Durchsuchen...'
$btnBrowse.Location = New-Object System.Drawing.Point(660, 14)
$btnBrowse.Size = New-Object System.Drawing.Size(100, 28)
$btnBrowse.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = 'CSV/Text-Dateien (*.csv;*.txt)|*.csv;*.txt|Alle Dateien (*.*)|*.*'
        if ($ofd.ShowDialog() -eq 'OK') {
            $txtFile.Text = $ofd.FileName
        }
    })
$form.Controls.Add($btnBrowse)

$btnImport = New-Object System.Windows.Forms.Button
$btnImport.Text = 'Import starten'
$btnImport.Location = New-Object System.Drawing.Point(12, 50)
$btnImport.Size = New-Object System.Drawing.Size(180, 40)
$btnImport.BackColor = [System.Drawing.Color]::LightGreen
$btnImport.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtFile.Text)) {
            [System.Windows.Forms.MessageBox]::Show('Bitte eine Datei auswaehlen!', 'Hinweis', 'OK', 'Warning')
            return
        }

        $btnImport.Enabled = $false
        Start-InitialImport -SourceFile $txtFile.Text
        $btnImport.Enabled = $true
    })
$form.Controls.Add($btnImport)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::White
$txtLog.Location = New-Object System.Drawing.Point(12, 100)
$txtLog.Size = New-Object System.Drawing.Size(740, 460)
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtLog)

$global:LogBox = $txtLog

Write-ImportLog "Tool gestartet - $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')" 'INFO'
Write-ImportLog "Logdatei: $LogFile" 'INFO'
Write-ImportLog "Datenbank: $DbPath" 'INFO'

$form.ShowDialog() | Out-Null