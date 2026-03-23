# StockComparisonUi.ps1
# Standalone WPF workbench for Invoke-StockComparison.

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$Script:ProjectRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path $MyInvocation.MyCommand.Path -Parent }
$Script:CompareScriptPath = Join-Path $Script:ProjectRoot 'Function_StockComparison.ps1'
$Script:DefaultDatabasePath = Join-Path $Script:ProjectRoot 'Core\db_verlegepaket.json'
$Script:DefaultPresetStorePath = Join-Path $Script:ProjectRoot 'Core\compare_presets.json'

if (-not (Test-Path $Script:CompareScriptPath)) {
    throw "Comparison engine not found: $Script:CompareScriptPath"
}

. $Script:CompareScriptPath

function New-UiSourceRow {
    param(
        [bool]$IsStockRole = $false,
        [string]$RoleName = '',
        [string]$Path = '',
        [string]$PresetName = '',
        [string]$WorksheetName = '',
        [string]$Delimiter = ''
    )

    return [pscustomobject]@{
        IsStockRole   = $IsStockRole
        RoleName      = $RoleName
        Path          = $Path
        PresetName    = $PresetName
        WorksheetName = $WorksheetName
        Delimiter     = $Delimiter
    }
}

function ConvertTo-SummaryRows {
    param([AllowNull()]$Summary)

    if ($null -eq $Summary) {
        return @()
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($propertyName in @(
            'MaterialCount',
            'MatchedMaterialCount',
            'MissingMaterialCount',
            'SurplusMaterialCount',
            'BalancedMaterialCount',
            'TargetQuantityBaseTotal',
            'StockQuantityBaseTotal',
            'InboundQuantityBaseTotal',
            'AvailableQuantityBaseTotal',
            'MissingToOrderBaseTotal',
            'SurplusAfterInboundBaseTotal',
            'StockGapBaseTotal',
            'StockRoleName'
        )) {
        $property = $Summary.PSObject.Properties[$propertyName]
        if ($null -eq $property) {
            continue
        }

        [void]$rows.Add([pscustomobject]@{
                Property = $propertyName
                Value    = Get-NormalizedString $property.Value
            })
    }

    return @($rows.ToArray())
}

function ConvertTo-DiagnosticTypeOptions {
    return @(
        [pscustomobject]@{ Key = 'unknown_sap_materials'; Label = 'Unknown SAP materials' }
        [pscustomobject]@{ Key = 'invalid_units'; Label = 'Invalid units' }
        [pscustomobject]@{ Key = 'invalid_rows'; Label = 'Invalid rows' }
        [pscustomobject]@{ Key = 'duplicate_aliases'; Label = 'Duplicate aliases' }
    )
}

function Convert-FlatRowsToExportObjects {
    param([AllowNull()][object[]]$Rows)

    $exportRows = New-Object System.Collections.Generic.List[object]
    foreach ($row in @(ConvertTo-ObjectArray $Rows)) {
        $ordered = [ordered]@{}
        foreach ($property in $row.PSObject.Properties) {
            $value = $property.Value
            if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
                $ordered[$property.Name] = ((@(ConvertTo-ObjectArray $value) | ForEach-Object { Get-NormalizedString $_ }) -join ', ')
            }
            else {
                $ordered[$property.Name] = Get-NormalizedString $value
            }
        }

        [void]$exportRows.Add([pscustomobject]$ordered)
    }

    return @($exportRows.ToArray())
}

function Export-FlatRowsToCsv {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [AllowNull()][object[]]$Rows
    )

    $exportRows = Convert-FlatRowsToExportObjects -Rows $Rows
    $exportRows | Export-Csv -Path $Path -Delimiter ';' -NoTypeInformation -Encoding UTF8
}

function Set-ExcelWorksheetData {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [Parameter(Mandatory = $true)][string]$SheetName,
        [AllowNull()][object[]]$Rows
    )

    $Worksheet.Name = $SheetName
    $resolvedRows = @(ConvertTo-ObjectArray $Rows)
    if ($resolvedRows.Count -eq 0) {
        $Worksheet.Cells.Item(1, 1) = 'No data'
        return
    }

    $headers = @($resolvedRows[0].PSObject.Properties.Name)
    for ($columnIndex = 0; $columnIndex -lt $headers.Count; $columnIndex++) {
        $Worksheet.Cells.Item(1, $columnIndex + 1) = $headers[$columnIndex]
    }

    $rowIndex = 2
    foreach ($row in $resolvedRows) {
        for ($columnIndex = 0; $columnIndex -lt $headers.Count; $columnIndex++) {
            $header = $headers[$columnIndex]
            $Worksheet.Cells.Item($rowIndex, $columnIndex + 1) = Get-NormalizedString $row.PSObject.Properties[$header].Value
        }
        $rowIndex++
    }

    [void]$Worksheet.Columns.AutoFit()
}

function Export-FlatRowsToXlsx {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [AllowNull()][object[]]$Rows,
        [string]$SheetName = 'Comparison Results'
    )

    $exportRows = Convert-FlatRowsToExportObjects -Rows $Rows
    $excel = $null
    $workbook = $null
    $worksheet = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        Set-ExcelWorksheetData -Worksheet $worksheet -SheetName $SheetName -Rows $exportRows
        while ($workbook.Worksheets.Count -gt 1) {
            $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
        }

        $workbook.SaveAs($Path, 51)
    }
    finally {
        if ($null -ne $workbook) {
            $workbook.Close($true)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        }

        if ($null -ne $worksheet) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet)
        }

        if ($null -ne $excel) {
            $excel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }
    }
}

function Start-StockComparisonUi {
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Stock Comparison"
        Height="980"
        Width="1680"
        MinHeight="860"
        MinWidth="1400"
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
            <Setter Property="Margin" Value="0,6,12,6"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>
        <Style TargetType="DataGrid">
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
                    <TextBlock Text="Stock Comparison" FontFamily="Bahnschrift SemiBold" FontSize="28" Foreground="White"/>
                    <TextBlock Text="Configure SAP source files, manage presets, run SOLL/IST comparison, and inspect diagnostics." Margin="0,4,0,0" Foreground="#CBD5E1"/>
                    <TextBlock x:Name="txtDbPath" Margin="0,8,0,0" Foreground="#93C5FD"/>
                    <TextBlock x:Name="txtPresetStorePath" Margin="0,2,0,0" Foreground="#93C5FD"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top">
                    <Button x:Name="btnManagePresets" Content="Presets..." Style="{StaticResource SecondaryButton}"/>
                    <Button x:Name="btnReloadPresets" Content="Reload presets" Style="{StaticResource SecondaryButton}"/>
                    <Button x:Name="btnRunComparison" Content="Vergleich starten" Style="{StaticResource PrimaryButton}" Margin="0"/>
                </StackPanel>
            </Grid>
        </Border>

        <Border DockPanel.Dock="Bottom" Background="White" BorderBrush="#E2E8F0" BorderThickness="1,1,0,0" Padding="18,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="txtStatus" Grid.Column="0" Foreground="#475569" VerticalAlignment="Center" Text="Ready"/>
                <TextBlock Grid.Column="1" Foreground="#94A3B8" VerticalAlignment="Center" Text="Thin GUI wrapper around Function_StockComparison.ps1"/>
            </Grid>
        </Border>

        <Grid Margin="20">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="500"/>
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
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <StackPanel Grid.Row="0">
                        <TextBlock Style="{StaticResource PanelTitle}" Text="Configuration"/>
                        <TextBlock Margin="0,4,0,0" Foreground="#64748B" Text="Maintain source definitions, choose the stock role, and assign saved SAP presets."/>
                    </StackPanel>

                    <Grid Grid.Row="1" Margin="0,16,0,10">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="110"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="txtDatabasePath" Grid.Column="0" Margin="0" ToolTip="Database JSON path"/>
                        <Button x:Name="btnBrowseDatabase" Grid.Column="1" Content="DB..." Style="{StaticResource SecondaryButton}" Margin="10,0,0,0"/>
                    </Grid>

                    <Grid Grid.Row="2" Margin="0,0,0,12">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="110"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="txtPresetPath" Grid.Column="0" Margin="0" ToolTip="Preset store JSON path"/>
                        <Button x:Name="btnBrowsePresetPath" Grid.Column="1" Content="Presets..." Style="{StaticResource SecondaryButton}" Margin="10,0,0,0"/>
                    </Grid>

                    <Grid Grid.Row="3">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="260"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <DockPanel Grid.Row="0" LastChildFill="False">
                            <TextBlock DockPanel.Dock="Left" Style="{StaticResource PanelTitle}" Text="Source List"/>
                            <Button x:Name="btnNewSource" DockPanel.Dock="Right" Content="New source" Style="{StaticResource SecondaryButton}" Margin="0"/>
                            <Button x:Name="btnRemoveSource" DockPanel.Dock="Right" Content="Remove" Style="{StaticResource SecondaryButton}"/>
                            <Button x:Name="btnApplySource" DockPanel.Dock="Right" Content="Apply" Style="{StaticResource PrimaryButton}"/>
                        </DockPanel>
                        <DataGrid x:Name="dgSources" Grid.Row="1" SelectionMode="Single" SelectionUnit="FullRow" IsReadOnly="True" RowHeaderWidth="0">
                            <DataGrid.Columns>
                                <DataGridCheckBoxColumn Header="Stock" Binding="{Binding IsStockRole}" Width="64"/>
                                <DataGridTextColumn Header="Role" Binding="{Binding RoleName}" Width="120"/>
                                <DataGridTextColumn Header="Path" Binding="{Binding Path}" Width="*"/>
                                <DataGridTextColumn Header="Preset" Binding="{Binding PresetName}" Width="130"/>
                            </DataGrid.Columns>
                        </DataGrid>

                        <TextBlock Grid.Row="2" Margin="0,16,0,0" Style="{StaticResource PanelTitle}" Text="Source Editor"/>
                        <Grid Grid.Row="3" Margin="0,8,0,0">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="140"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="110"/>
                            </Grid.ColumnDefinitions>

                            <TextBlock Grid.Row="0" Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Stock role"/>
                            <CheckBox x:Name="chkSourceIsStock" Grid.Row="0" Grid.Column="1" Content="Use this source as the physical stock list"/>

                            <TextBlock Grid.Row="1" Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Role name"/>
                            <TextBox x:Name="txtSourceRoleName" Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="2"/>

                            <TextBlock Grid.Row="2" Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Source file"/>
                            <TextBox x:Name="txtSourcePath" Grid.Row="2" Grid.Column="1"/>
                            <Button x:Name="btnBrowseSourceFile" Grid.Row="2" Grid.Column="2" Content="Browse..." Style="{StaticResource SecondaryButton}" Margin="10,4,0,10"/>

                            <TextBlock Grid.Row="3" Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Preset"/>
                            <ComboBox x:Name="cmbSourcePreset" Grid.Row="3" Grid.Column="1"/>
                            <Button x:Name="btnRefreshPresetCombo" Grid.Row="3" Grid.Column="2" Content="Refresh" Style="{StaticResource SecondaryButton}" Margin="10,4,0,10"/>

                            <TextBlock Grid.Row="4" Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Worksheet / Delimiter"/>
                            <Grid Grid.Row="4" Grid.Column="1" Grid.ColumnSpan="2">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="16"/>
                                    <ColumnDefinition Width="120"/>
                                </Grid.ColumnDefinitions>
                                <TextBox x:Name="txtSourceWorksheet" Grid.Column="0" ToolTip="Optional worksheet name"/>
                                <TextBox x:Name="txtSourceDelimiter" Grid.Column="2" ToolTip="Optional delimiter override"/>
                            </Grid>
                        </Grid>
                    </Grid>

                    <TextBlock Grid.Row="4" x:Name="txtSourceMeta" Margin="0,16,0,0" Foreground="#64748B" Text="0 sources configured"/>
                </Grid>
            </Border>

            <GridSplitter Grid.Column="1" Width="8" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Background="Transparent"/>

            <Border Grid.Column="2" Background="White" BorderBrush="#E2E8F0" BorderThickness="1" CornerRadius="10" Padding="18">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <StackPanel Grid.Row="0">
                        <TextBlock x:Name="txtResultsHeadline" FontFamily="Bahnschrift SemiBold" FontSize="24" Foreground="#0F172A" Text="Comparison Results"/>
                        <TextBlock x:Name="txtResultsMeta" Margin="0,4,0,0" Foreground="#64748B" Text="Run a comparison to populate summary, results, diagnostics, and source execution details."/>
                    </StackPanel>

                    <TabControl Grid.Row="1" Margin="0,18,0,0">
                        <TabItem Header="Summary">
                            <Grid Margin="6">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <TextBlock Grid.Row="0" Foreground="#334155" FontWeight="SemiBold" Text="Overall totals"/>
                                <DataGrid x:Name="dgSummary" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True" RowHeaderWidth="0">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Property" Binding="{Binding Property}" Width="260"/>
                                        <DataGridTextColumn Header="Value" Binding="{Binding Value}" Width="*"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                                <TextBlock Grid.Row="2" Margin="0,16,0,0" Foreground="#334155" FontWeight="SemiBold" Text="Role totals"/>
                                <DataGrid x:Name="dgRoleSummary" Grid.Row="3" AutoGenerateColumns="True" IsReadOnly="True" RowHeaderWidth="0"/>
                            </Grid>
                        </TabItem>

                        <TabItem Header="Results">
                            <Grid Margin="6">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <Grid Grid.Row="0">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="110"/>
                                        <ColumnDefinition Width="110"/>
                                        <ColumnDefinition Width="110"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox x:Name="txtResultSearch" Grid.Column="0" Margin="0" ToolTip="Search visible results"/>
                                    <Button x:Name="btnClearResultSearch" Grid.Column="1" Content="Clear" Style="{StaticResource SecondaryButton}" Margin="10,0,0,0"/>
                                    <Button x:Name="btnExportResultsCsv" Grid.Column="2" Content="Export CSV" Style="{StaticResource SecondaryButton}" Margin="10,0,0,0"/>
                                    <Button x:Name="btnExportResultsXlsx" Grid.Column="3" Content="Export XLSX" Style="{StaticResource SecondaryButton}" Margin="10,0,0,0"/>
                                </Grid>
                                <TextBlock Grid.Row="1" x:Name="txtResultGridMeta" Margin="0,12,0,0" Foreground="#64748B" Text="0 rows"/>
                                <DataGrid x:Name="dgResults" Grid.Row="2" AutoGenerateColumns="True" IsReadOnly="True" RowHeaderWidth="0"/>
                            </Grid>
                        </TabItem>

                        <TabItem Header="Diagnostics">
                            <Grid Margin="6">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <Grid Grid.Row="0">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="260"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <ComboBox x:Name="cmbDiagnosticType" Grid.Column="0" DisplayMemberPath="Label" SelectedValuePath="Key"/>
                                </Grid>
                                <TextBlock Grid.Row="1" x:Name="txtDiagnosticsMeta" Margin="0,12,0,0" Foreground="#64748B" Text="0 rows"/>
                                <DataGrid x:Name="dgDiagnostics" Grid.Row="2" AutoGenerateColumns="True" IsReadOnly="True" RowHeaderWidth="0"/>
                            </Grid>
                        </TabItem>

                        <TabItem Header="Sources">
                            <Grid Margin="6">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <TextBlock Grid.Row="0" x:Name="txtSourcesMeta" Foreground="#64748B" Text="0 source executions"/>
                                <DataGrid x:Name="dgSourceResults" Grid.Row="1" AutoGenerateColumns="True" IsReadOnly="True" RowHeaderWidth="0"/>
                            </Grid>
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
    $txtPresetStorePath = $window.FindName('txtPresetStorePath')
    $btnManagePresets = $window.FindName('btnManagePresets')
    $btnReloadPresets = $window.FindName('btnReloadPresets')
    $btnRunComparison = $window.FindName('btnRunComparison')
    $txtStatus = $window.FindName('txtStatus')
    $txtDatabasePath = $window.FindName('txtDatabasePath')
    $btnBrowseDatabase = $window.FindName('btnBrowseDatabase')
    $txtPresetPath = $window.FindName('txtPresetPath')
    $btnBrowsePresetPath = $window.FindName('btnBrowsePresetPath')
    $btnNewSource = $window.FindName('btnNewSource')
    $btnRemoveSource = $window.FindName('btnRemoveSource')
    $btnApplySource = $window.FindName('btnApplySource')
    $dgSources = $window.FindName('dgSources')
    $chkSourceIsStock = $window.FindName('chkSourceIsStock')
    $txtSourceRoleName = $window.FindName('txtSourceRoleName')
    $txtSourcePath = $window.FindName('txtSourcePath')
    $btnBrowseSourceFile = $window.FindName('btnBrowseSourceFile')
    $cmbSourcePreset = $window.FindName('cmbSourcePreset')
    $btnRefreshPresetCombo = $window.FindName('btnRefreshPresetCombo')
    $txtSourceWorksheet = $window.FindName('txtSourceWorksheet')
    $txtSourceDelimiter = $window.FindName('txtSourceDelimiter')
    $txtSourceMeta = $window.FindName('txtSourceMeta')
    $txtResultsHeadline = $window.FindName('txtResultsHeadline')
    $txtResultsMeta = $window.FindName('txtResultsMeta')
    $dgSummary = $window.FindName('dgSummary')
    $dgRoleSummary = $window.FindName('dgRoleSummary')
    $txtResultSearch = $window.FindName('txtResultSearch')
    $btnClearResultSearch = $window.FindName('btnClearResultSearch')
    $btnExportResultsCsv = $window.FindName('btnExportResultsCsv')
    $btnExportResultsXlsx = $window.FindName('btnExportResultsXlsx')
    $txtResultGridMeta = $window.FindName('txtResultGridMeta')
    $dgResults = $window.FindName('dgResults')
    $cmbDiagnosticType = $window.FindName('cmbDiagnosticType')
    $txtDiagnosticsMeta = $window.FindName('txtDiagnosticsMeta')
    $dgDiagnostics = $window.FindName('dgDiagnostics')
    $txtSourcesMeta = $window.FindName('txtSourcesMeta')
    $dgSourceResults = $window.FindName('dgSourceResults')

    $state = [ordered]@{
        SourceRows            = New-Object System.Collections.ArrayList
        SelectedSourceRow     = $null
        PresetStore           = $null
        PresetNames           = @()
        Result                = $null
        FilteredResultRows    = @()
        SelectedDiagnosticKey = 'unknown_sap_materials'
    }

    $txtDatabasePath.Text = $Script:DefaultDatabasePath
    $txtPresetPath.Text = $Script:DefaultPresetStorePath
    $txtDbPath.Text = "Database: $($txtDatabasePath.Text)"
    $txtPresetStorePath.Text = "Preset store: $($txtPresetPath.Text)"
    $cmbDiagnosticType.ItemsSource = @(ConvertTo-DiagnosticTypeOptions)
    $cmbDiagnosticType.SelectedValue = 'unknown_sap_materials'

    $SetStatus = $null
    $RefreshPresetStore = $null
    $RefreshSourceEditor = $null
    $RefreshSourceList = $null
    $CommitSourceEditor = $null
    $RefreshResultTabs = $null
    $ApplyResultFilter = $null
    $OpenPresetManagerDialog = $null
    $ExportResults = $null
    $BuildSourceSpecs = $null
    $RunComparison = $null

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

    $RefreshPresetStore = {
        try {
            $state.PresetStore = Get-StockComparisonPresetStore -Path $txtPresetPath.Text
            $state.PresetNames = @($state.PresetStore.presets | ForEach-Object { Get-NormalizedString $_.Name } | Sort-Object)
            $cmbSourcePreset.ItemsSource = @($state.PresetNames)
            if ($null -ne $state.SelectedSourceRow) {
                $currentPresetName = Get-NormalizedString $state.SelectedSourceRow.PresetName
                if (-not [string]::IsNullOrWhiteSpace($currentPresetName)) {
                    $cmbSourcePreset.SelectedItem = $currentPresetName
                }
            }

            & $SetStatus -Message "Loaded $($state.PresetNames.Count) presets." -Level 'Success'
        }
        catch {
            $state.PresetStore = $null
            $state.PresetNames = @()
            $cmbSourcePreset.ItemsSource = @()
            & $SetStatus -Message "Failed to load presets: $($_.Exception.Message)" -Level 'Error'
        }
    }

    $RefreshSourceEditor = {
        $selected = $state.SelectedSourceRow
        if ($null -eq $selected) {
            $chkSourceIsStock.IsChecked = $false
            $txtSourceRoleName.Text = ''
            $txtSourcePath.Text = ''
            $cmbSourcePreset.SelectedItem = $null
            $txtSourceWorksheet.Text = ''
            $txtSourceDelimiter.Text = ''
            return
        }

        $chkSourceIsStock.IsChecked = [bool]$selected.IsStockRole
        $txtSourceRoleName.Text = Get-NormalizedString $selected.RoleName
        $txtSourcePath.Text = Get-NormalizedString $selected.Path
        $cmbSourcePreset.SelectedItem = Get-NormalizedString $selected.PresetName
        $txtSourceWorksheet.Text = Get-NormalizedString $selected.WorksheetName
        $txtSourceDelimiter.Text = Get-NormalizedString $selected.Delimiter
    }

    $RefreshSourceList = {
        $dgSources.ItemsSource = $null
        $dgSources.ItemsSource = @($state.SourceRows.ToArray())
        $txtSourceMeta.Text = '{0} sources configured' -f $state.SourceRows.Count
        $btnRemoveSource.IsEnabled = ($null -ne $state.SelectedSourceRow)
    }

    $CommitSourceEditor = {
        $roleName = Get-NormalizedString $txtSourceRoleName.Text
        $path = Get-NormalizedString $txtSourcePath.Text
        $presetName = Get-NormalizedString $cmbSourcePreset.Text
        $worksheetName = Get-NormalizedString $txtSourceWorksheet.Text
        $delimiter = Get-NormalizedString $txtSourceDelimiter.Text
        $isStockRole = [bool]$chkSourceIsStock.IsChecked

        if ([string]::IsNullOrWhiteSpace($roleName)) {
            throw 'Role name is required.'
        }

        if ([string]::IsNullOrWhiteSpace($path)) {
            throw 'Source path is required.'
        }

        if ([string]::IsNullOrWhiteSpace($presetName)) {
            throw 'Preset name is required.'
        }

        $row = if ($null -ne $state.SelectedSourceRow) {
            $state.SelectedSourceRow
        }
        else {
            $newRow = New-UiSourceRow
            [void]$state.SourceRows.Add($newRow)
            $newRow
        }

        if ($isStockRole) {
            foreach ($existingRow in @($state.SourceRows.ToArray())) {
                $existingRow.IsStockRole = $false
            }
        }

        $row.IsStockRole = $isStockRole
        $row.RoleName = $roleName
        $row.Path = $path
        $row.PresetName = $presetName
        $row.WorksheetName = $worksheetName
        $row.Delimiter = $delimiter
        $state.SelectedSourceRow = $row
        & $RefreshSourceList
        $dgSources.SelectedItem = $row
        & $SetStatus -Message "Source '$roleName' staged." -Level 'Success'
    }

    $ApplyResultFilter = {
        if ($null -eq $state.Result) {
            $state.FilteredResultRows = @()
            $dgResults.ItemsSource = @()
            $txtResultGridMeta.Text = '0 rows'
            return
        }

        $search = Get-NormalizedString $txtResultSearch.Text
        $rows = @($state.Result.GridRows)
        if (-not [string]::IsNullOrWhiteSpace($search)) {
            $needle = $search.ToLowerInvariant()
            $rows = @(
                foreach ($row in @($state.Result.GridRows)) {
                    $searchText = ((@($row.PSObject.Properties | ForEach-Object { Get-NormalizedString $_.Value })) -join ' ').ToLowerInvariant()
                    if ($searchText.Contains($needle)) {
                        $row
                    }
                }
            )
        }

        $state.FilteredResultRows = @($rows)
        $dgResults.ItemsSource = $null
        $dgResults.ItemsSource = @($state.FilteredResultRows)
        $txtResultGridMeta.Text = '{0} visible rows' -f $state.FilteredResultRows.Count
    }

    $RefreshResultTabs = {
        if ($null -eq $state.Result) {
            $txtResultsMeta.Text = 'Run a comparison to populate summary, results, diagnostics, and source execution details.'
            $dgSummary.ItemsSource = @()
            $dgRoleSummary.ItemsSource = @()
            $dgDiagnostics.ItemsSource = @()
            $dgSourceResults.ItemsSource = @()
            $txtDiagnosticsMeta.Text = '0 rows'
            $txtSourcesMeta.Text = '0 source executions'
            & $ApplyResultFilter
            return
        }

        $summaryRows = ConvertTo-SummaryRows -Summary $state.Result.Summary
        $dgSummary.ItemsSource = @($summaryRows)
        $dgRoleSummary.ItemsSource = @($state.Result.Summary.RoleTotals)
        $dgSourceResults.ItemsSource = @($state.Result.Sources)
        $txtResultsMeta.Text = 'Materials: {0} | Missing: {1} | Surplus: {2}' -f `
            $state.Result.Summary.MaterialCount, `
            $state.Result.Summary.MissingMaterialCount, `
            $state.Result.Summary.SurplusMaterialCount
        $txtSourcesMeta.Text = '{0} source executions' -f @($state.Result.Sources).Count

        $selectedDiagnosticKey = Get-NormalizedString $cmbDiagnosticType.SelectedValue
        if ([string]::IsNullOrWhiteSpace($selectedDiagnosticKey)) {
            $selectedDiagnosticKey = 'unknown_sap_materials'
        }

        $diagnosticRows = @(Get-DeepPropertyValue $state.Result.Diagnostics $selectedDiagnosticKey @())
        $dgDiagnostics.ItemsSource = @($diagnosticRows)
        $txtDiagnosticsMeta.Text = '{0} rows' -f $diagnosticRows.Count
        & $ApplyResultFilter
    }

    $BuildSourceSpecs = {
        if ($state.SourceRows.Count -eq 0) {
            throw 'At least one source is required.'
        }

        $stockRows = @(@($state.SourceRows.ToArray()) | Where-Object { $_.IsStockRole })
        if ($stockRows.Count -ne 1) {
            throw 'Exactly one source must be marked as the stock role.'
        }

        $sourceSpecs = New-Object System.Collections.Generic.List[object]
        foreach ($row in @($state.SourceRows.ToArray())) {
            $roleName = Get-NormalizedString $row.RoleName
            $path = Get-NormalizedString $row.Path
            $presetName = Get-NormalizedString $row.PresetName
            if ([string]::IsNullOrWhiteSpace($roleName) -or [string]::IsNullOrWhiteSpace($path) -or [string]::IsNullOrWhiteSpace($presetName)) {
                throw 'Every source must have RoleName, Path, and PresetName.'
            }

            [void]$sourceSpecs.Add([pscustomobject]@{
                    RoleName      = $roleName
                    Path          = $path
                    PresetName    = $presetName
                    WorksheetName = ConvertTo-NullableString $row.WorksheetName
                    Delimiter     = ConvertTo-NullableString $row.Delimiter
                })
        }

        return [pscustomobject]@{
            StockRoleName = Get-NormalizedString $stockRows[0].RoleName
            SourceSpecs   = @($sourceSpecs.ToArray())
        }
    }

    $OpenPresetManagerDialog = {
        $dialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Preset Manager"
        Height="760"
        Width="1160"
        MinHeight="700"
        MinWidth="980"
        WindowStartupLocation="CenterOwner"
        Background="#F8FAFC">
    <Grid Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0">
            <TextBlock FontFamily="Bahnschrift SemiBold" FontSize="22" Foreground="#0F172A" Text="Preset Manager"/>
            <TextBlock Margin="0,6,0,0" Foreground="#64748B" Text="Create, edit, save, and reload SAP source presets."/>
        </StackPanel>
        <Grid Grid.Row="1" Margin="0,18,0,18">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="260"/>
                <ColumnDefinition Width="20"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0">
                <DockPanel LastChildFill="False">
                    <TextBlock DockPanel.Dock="Left" Foreground="#334155" FontWeight="SemiBold" Text="Presets"/>
                    <Button x:Name="btnPresetNew" DockPanel.Dock="Right" Width="90" Content="New"/>
                    <Button x:Name="btnPresetReload" DockPanel.Dock="Right" Width="90" Margin="10,0,10,0" Content="Reload"/>
                </DockPanel>
                <ListBox x:Name="lbPresets" Margin="0,10,0,0"/>
            </StackPanel>
            <ScrollViewer Grid.Column="2" VerticalScrollBarVisibility="Auto">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="170"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Row="0" Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Name"/>
                    <TextBox x:Name="txtPresetName" Grid.Row="0" Grid.Column="1"/>
                    <TextBlock Grid.Row="1" Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="File type"/>
                    <ComboBox x:Name="cmbPresetFileType" Grid.Row="1" Grid.Column="1"/>
                    <TextBlock Grid.Row="2" Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Header row index"/>
                    <TextBox x:Name="txtPresetHeaderRow" Grid.Row="2" Grid.Column="1"/>
                    <TextBlock Grid.Row="3" Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Worksheet name"/>
                    <TextBox x:Name="txtPresetWorksheet" Grid.Row="3" Grid.Column="1"/>
                    <TextBlock Grid.Row="4" Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Delimiter"/>
                    <TextBox x:Name="txtPresetDelimiter" Grid.Row="4" Grid.Column="1"/>
                    <TextBlock Grid.Row="5" Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Material column"/>
                    <TextBox x:Name="txtColMaterial" Grid.Row="5" Grid.Column="1"/>
                    <TextBlock Grid.Row="6" Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Quantity column"/>
                    <TextBox x:Name="txtColQuantity" Grid.Row="6" Grid.Column="1"/>
                    <TextBlock Grid.Row="7" Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Unit column"/>
                    <TextBox x:Name="txtColUnit" Grid.Row="7" Grid.Column="1"/>
                    <TextBlock Grid.Row="8" Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Description column"/>
                    <TextBox x:Name="txtColDescription" Grid.Row="8" Grid.Column="1"/>
                    <TextBlock Grid.Row="9" Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Status column"/>
                    <TextBox x:Name="txtColStatus" Grid.Row="9" Grid.Column="1"/>
                    <TextBlock Grid.Row="10" Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Storage / Batch"/>
                    <Grid Grid.Row="10" Grid.Column="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="12"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="txtColStorage" Grid.Column="0"/>
                        <TextBox x:Name="txtColBatch" Grid.Column="2"/>
                    </Grid>
                    <TextBlock Grid.Row="11" Grid.Column="0" VerticalAlignment="Center" Foreground="#475569" Text="Note column"/>
                    <TextBox x:Name="txtColNote" Grid.Row="11" Grid.Column="1"/>
                </Grid>
            </ScrollViewer>
        </Grid>
        <DockPanel Grid.Row="2" LastChildFill="False">
            <TextBlock x:Name="txtPresetDialogStatus" DockPanel.Dock="Left" VerticalAlignment="Center" Foreground="#475569" Text="Ready"/>
            <Button x:Name="btnPresetClose" DockPanel.Dock="Right" Width="110" Content="Close"/>
            <Button x:Name="btnPresetSave" DockPanel.Dock="Right" Width="110" Margin="10,0,10,0" Content="Save"/>
        </DockPanel>
    </Grid>
</Window>
"@

        $dialogReader = New-Object System.Xml.XmlNodeReader([xml]$dialogXaml)
        $dialog = [Windows.Markup.XamlReader]::Load($dialogReader)
        $dialog.Owner = $window

        $lbPresets = $dialog.FindName('lbPresets')
        $btnPresetNew = $dialog.FindName('btnPresetNew')
        $btnPresetReload = $dialog.FindName('btnPresetReload')
        $btnPresetClose = $dialog.FindName('btnPresetClose')
        $btnPresetSave = $dialog.FindName('btnPresetSave')
        $txtPresetDialogStatus = $dialog.FindName('txtPresetDialogStatus')
        $txtPresetName = $dialog.FindName('txtPresetName')
        $cmbPresetFileType = $dialog.FindName('cmbPresetFileType')
        $txtPresetHeaderRow = $dialog.FindName('txtPresetHeaderRow')
        $txtPresetWorksheet = $dialog.FindName('txtPresetWorksheet')
        $txtPresetDelimiter = $dialog.FindName('txtPresetDelimiter')
        $txtColMaterial = $dialog.FindName('txtColMaterial')
        $txtColQuantity = $dialog.FindName('txtColQuantity')
        $txtColUnit = $dialog.FindName('txtColUnit')
        $txtColDescription = $dialog.FindName('txtColDescription')
        $txtColStatus = $dialog.FindName('txtColStatus')
        $txtColStorage = $dialog.FindName('txtColStorage')
        $txtColBatch = $dialog.FindName('txtColBatch')
        $txtColNote = $dialog.FindName('txtColNote')
        $cmbPresetFileType.ItemsSource = @('csv', 'xlsx', 'txt')

        $setDialogStatus = {
            param([string]$Message, [string]$Level = 'Info')
            $txtPresetDialogStatus.Text = $Message
            switch ($Level) {
                'Error' { $txtPresetDialogStatus.Foreground = '#B91C1C' }
                'Success' { $txtPresetDialogStatus.Foreground = '#0F766E' }
                default { $txtPresetDialogStatus.Foreground = '#475569' }
            }
        }

        $populatePresetEditor = {
            param($Preset)
            if ($null -eq $Preset) {
                $txtPresetName.Text = ''
                $cmbPresetFileType.SelectedItem = 'csv'
                $txtPresetHeaderRow.Text = '1'
                $txtPresetWorksheet.Text = ''
                $txtPresetDelimiter.Text = ';'
                $txtColMaterial.Text = ''
                $txtColQuantity.Text = ''
                $txtColUnit.Text = ''
                $txtColDescription.Text = ''
                $txtColStatus.Text = ''
                $txtColStorage.Text = ''
                $txtColBatch.Text = ''
                $txtColNote.Text = ''
                return
            }

            $txtPresetName.Text = Get-NormalizedString $Preset.Name
            $cmbPresetFileType.SelectedItem = Get-NormalizedString $Preset.FileType
            $txtPresetHeaderRow.Text = [string][int]$Preset.HeaderRowIndex
            $txtPresetWorksheet.Text = Get-NormalizedString $Preset.WorksheetName
            $txtPresetDelimiter.Text = Get-NormalizedString $Preset.Delimiter
            $txtColMaterial.Text = Get-NormalizedString $Preset.Columns.material_number
            $txtColQuantity.Text = Get-NormalizedString $Preset.Columns.quantity
            $txtColUnit.Text = Get-NormalizedString $Preset.Columns.unit
            $txtColDescription.Text = Get-NormalizedString $Preset.Columns.description
            $txtColStatus.Text = Get-NormalizedString $Preset.Columns.status
            $txtColStorage.Text = Get-NormalizedString $Preset.Columns.storage_location
            $txtColBatch.Text = Get-NormalizedString $Preset.Columns.batch
            $txtColNote.Text = Get-NormalizedString $Preset.Columns.note
        }

        $refreshPresetList = {
            $store = Get-StockComparisonPresetStore -Path $txtPresetPath.Text
            $lbPresets.ItemsSource = @($store.presets | Sort-Object Name)
            $lbPresets.DisplayMemberPath = 'Name'
        }

        & $refreshPresetList
        & $populatePresetEditor -Preset $null
        & $setDialogStatus -Message 'Preset list loaded.'

        $lbPresets.Add_SelectionChanged({
                & $populatePresetEditor -Preset $lbPresets.SelectedItem
            })
        $btnPresetNew.Add_Click({
                $lbPresets.SelectedItem = $null
                & $populatePresetEditor -Preset $null
                & $setDialogStatus -Message 'New preset draft.'
            })
        $btnPresetReload.Add_Click({
                try {
                    & $refreshPresetList
                    & $setDialogStatus -Message 'Preset list reloaded.' -Level 'Success'
                }
                catch {
                    & $setDialogStatus -Message $_.Exception.Message -Level 'Error'
                }
            })
        $btnPresetSave.Add_Click({
                try {
                    $preset = [pscustomobject]@{
                        Name           = Get-NormalizedString $txtPresetName.Text
                        FileType       = Get-NormalizedString $cmbPresetFileType.Text
                        HeaderRowIndex = Get-NormalizedString $txtPresetHeaderRow.Text
                        WorksheetName  = Get-NormalizedString $txtPresetWorksheet.Text
                        Delimiter      = Get-NormalizedString $txtPresetDelimiter.Text
                        Columns        = [pscustomobject]@{
                            material_number = Get-NormalizedString $txtColMaterial.Text
                            quantity        = Get-NormalizedString $txtColQuantity.Text
                            unit            = Get-NormalizedString $txtColUnit.Text
                            description     = Get-NormalizedString $txtColDescription.Text
                            status          = Get-NormalizedString $txtColStatus.Text
                            storage_location = Get-NormalizedString $txtColStorage.Text
                            batch           = Get-NormalizedString $txtColBatch.Text
                            note            = Get-NormalizedString $txtColNote.Text
                        }
                    }

                    $savedPreset = Set-StockComparisonPreset -Preset $preset -Path $txtPresetPath.Text
                    & $refreshPresetList
                    $lbPresets.SelectedItem = @($lbPresets.ItemsSource | Where-Object { (Get-NormalizedString $_.Name) -eq (Get-NormalizedString $savedPreset.Name) })[0]
                    & $setDialogStatus -Message "Preset '$($savedPreset.Name)' saved." -Level 'Success'
                    & $RefreshPresetStore
                }
                catch {
                    & $setDialogStatus -Message $_.Exception.Message -Level 'Error'
                }
            })
        $btnPresetClose.Add_Click({
                $dialog.Close()
            })

        [void]$dialog.ShowDialog()
    }

    $ExportResults = {
        param([ValidateSet('csv', 'xlsx')][string]$Format)

        $visibleRows = @($state.FilteredResultRows)
        if ($visibleRows.Count -eq 0) {
            [System.Windows.MessageBox]::Show('There are no visible result rows to export.', 'Export', 'OK', 'Information') | Out-Null
            return
        }

        if ($Format -eq 'csv') {
            $dialog = New-Object System.Windows.Forms.SaveFileDialog
            $dialog.Filter = 'CSV-Dateien (*.csv)|*.csv'
            $dialog.FileName = 'stock_comparison.csv'
            if ($dialog.ShowDialog() -ne 'OK') { return }

            try {
                Export-FlatRowsToCsv -Path $dialog.FileName -Rows $visibleRows
                & $SetStatus -Message 'CSV export completed.' -Level 'Success'
            }
            catch {
                & $SetStatus -Message "CSV export failed: $($_.Exception.Message)" -Level 'Error'
            }
            return
        }

        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Filter = 'Excel-Dateien (*.xlsx)|*.xlsx'
        $dialog.FileName = 'stock_comparison.xlsx'
        if ($dialog.ShowDialog() -ne 'OK') { return }

        try {
            Export-FlatRowsToXlsx -Path $dialog.FileName -Rows $visibleRows
            & $SetStatus -Message 'XLSX export completed.' -Level 'Success'
        }
        catch {
            & $SetStatus -Message "XLSX export failed: $($_.Exception.Message)" -Level 'Error'
        }
    }

    $RunComparison = {
        try {
            if (
                $null -ne $state.SelectedSourceRow -or
                -not [string]::IsNullOrWhiteSpace((Get-NormalizedString $txtSourceRoleName.Text)) -or
                -not [string]::IsNullOrWhiteSpace((Get-NormalizedString $txtSourcePath.Text)) -or
                -not [string]::IsNullOrWhiteSpace((Get-NormalizedString $cmbSourcePreset.Text))
            ) {
                & $CommitSourceEditor
            }

            $config = & $BuildSourceSpecs
            & $SetStatus -Message 'Running comparison...' -Level 'Info'
            $btnRunComparison.IsEnabled = $false
            $result = Invoke-StockComparison -DatabasePath $txtDatabasePath.Text -PresetStorePath $txtPresetPath.Text -StockRoleName $config.StockRoleName -SourceSpecs $config.SourceSpecs
            $state.Result = $result
            & $RefreshResultTabs
            & $SetStatus -Message 'Comparison completed.' -Level 'Success'
        }
        catch {
            & $SetStatus -Message "Comparison failed: $($_.Exception.Message)" -Level 'Error'
            [System.Windows.MessageBox]::Show($_.Exception.Message, 'Comparison failed', 'OK', 'Error') | Out-Null
        }
        finally {
            $btnRunComparison.IsEnabled = $true
        }
    }

    $btnBrowseDatabase.Add_Click({
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Filter = 'JSON-Dateien (*.json)|*.json|Alle Dateien (*.*)|*.*'
            if ($dialog.ShowDialog() -eq 'OK') {
                $txtDatabasePath.Text = $dialog.FileName
                $txtDbPath.Text = "Database: $($txtDatabasePath.Text)"
                & $SetStatus -Message 'Database path updated.'
            }
        })
    $btnBrowsePresetPath.Add_Click({
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Filter = 'JSON-Dateien (*.json)|*.json|Alle Dateien (*.*)|*.*'
            if ($dialog.ShowDialog() -eq 'OK') {
                $txtPresetPath.Text = $dialog.FileName
                $txtPresetStorePath.Text = "Preset store: $($txtPresetPath.Text)"
                & $RefreshPresetStore
            }
        })
    $btnBrowseSourceFile.Add_Click({
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Filter = 'SAP-Dateien (*.xlsx;*.csv;*.txt)|*.xlsx;*.csv;*.txt|Alle Dateien (*.*)|*.*'
            if ($dialog.ShowDialog() -eq 'OK') {
                $txtSourcePath.Text = $dialog.FileName
            }
        })
    $btnNewSource.Add_Click({
            $state.SelectedSourceRow = $null
            & $RefreshSourceEditor
            $dgSources.SelectedItem = $null
            & $SetStatus -Message 'New source draft.'
        })
    $btnApplySource.Add_Click({
            try { & $CommitSourceEditor } catch {
                & $SetStatus -Message $_.Exception.Message -Level 'Error'
                [System.Windows.MessageBox]::Show($_.Exception.Message, 'Source validation', 'OK', 'Warning') | Out-Null
            }
        })
    $btnRemoveSource.Add_Click({
            if ($null -eq $state.SelectedSourceRow) { return }
            [void]$state.SourceRows.Remove($state.SelectedSourceRow)
            $state.SelectedSourceRow = $null
            & $RefreshSourceEditor
            & $RefreshSourceList
            & $SetStatus -Message 'Source removed.' -Level 'Warning'
        })
    $btnManagePresets.Add_Click({ & $OpenPresetManagerDialog })
    $btnReloadPresets.Add_Click({ & $RefreshPresetStore })
    $btnRefreshPresetCombo.Add_Click({ & $RefreshPresetStore })
    $btnRunComparison.Add_Click({ & $RunComparison })
    $txtResultSearch.Add_TextChanged({ & $ApplyResultFilter })
    $btnClearResultSearch.Add_Click({
            $txtResultSearch.Text = ''
            & $ApplyResultFilter
        })
    $btnExportResultsCsv.Add_Click({ & $ExportResults -Format 'csv' })
    $btnExportResultsXlsx.Add_Click({ & $ExportResults -Format 'xlsx' })
    $cmbDiagnosticType.Add_SelectionChanged({
            $state.SelectedDiagnosticKey = Get-NormalizedString $cmbDiagnosticType.SelectedValue
            & $RefreshResultTabs
        })
    $dgSources.Add_SelectionChanged({
            $state.SelectedSourceRow = $dgSources.SelectedItem
            & $RefreshSourceEditor
            & $RefreshSourceList
        })

    & $RefreshPresetStore
    [void]$state.SourceRows.Add((New-UiSourceRow -IsStockRole $true -RoleName 'stock'))
    & $RefreshSourceList
    $dgSummary.ItemsSource = @()
    $dgRoleSummary.ItemsSource = @()
    $dgDiagnostics.ItemsSource = @()
    $dgSourceResults.ItemsSource = @()
    & $RefreshSourceEditor
    & $RefreshResultTabs
    & $SetStatus -Message 'Ready'
    $window.ShowDialog() | Out-Null
}

if ($MyInvocation.InvocationName -ne '.') {
    Start-StockComparisonUi
}
