# Function_DataImport.ps1
# Windows-PowerShell-5.1-kompatible Importversion

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

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
        if ($global:LogBox -is [System.Windows.Controls.TextBox]) {
            $global:LogBox.ScrollToEnd()
        }
        else {
            $global:LogBox.ScrollToCaret()
        }
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

# GUI - WPF Modern Design
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
            <!-- Header -->
            <TextBlock Text="Verlegepaket Datenimport" 
                       FontSize="24" 
                       FontWeight="Bold" 
                       Foreground="#2C3E50"
                       Margin="0,0,0,10"/>

            <!-- File Selection Section -->
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

            <!-- Import Button -->
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

            <!-- Log Section -->
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

# Load XAML
$reader = New-Object System.Xml.XmlNodeReader([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get controls
$txtFile = $window.FindName("txtFile")
$btnBrowse = $window.FindName("btnBrowse")
$btnImport = $window.FindName("btnImport")
$txtLog = $window.FindName("txtLog")

$global:LogBox = $txtLog

# Browse button click
$btnBrowse.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = 'CSV/Text-Dateien (*.csv;*.txt)|*.csv;*.txt|Alle Dateien (*.*)|*.*'
        if ($ofd.ShowDialog() -eq 'OK') {
            $txtFile.Text = $ofd.FileName
        }
    })

# Import button click
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