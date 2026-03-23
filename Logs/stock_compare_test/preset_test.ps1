. 'd:\VPManagerGitHub\VerlegepaketManager\Function_StockComparison.ps1'
$presetPath = 'd:\VPManagerGitHub\VerlegepaketManager\Logs\stock_compare_test\compare_presets_test.json'
$null = Set-StockComparisonPreset -Path $presetPath -Preset ([pscustomobject]@{
    Name = 'sap-stock'
    FileType = 'csv'
    HeaderRowIndex = 2
    Delimiter = ';'
    Columns = [pscustomobject]@{
        material_number = 'MATNR'
        quantity = 'LABST'
        unit = 'MEINS'
        note = 'NOTE'
    }
})
$preset = Get-StockComparisonPreset -Path $presetPath -Name 'sap-stock'
[pscustomobject][ordered]@{
    Name = $preset.Name
    FileType = $preset.FileType
    HeaderRowIndex = $preset.HeaderRowIndex
    MaterialColumn = $preset.Columns.material_number
    NoteColumn = $preset.Columns.note
} | ConvertTo-Json -Compress
