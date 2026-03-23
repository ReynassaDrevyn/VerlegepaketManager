. 'd:\VPManagerGitHub\VerlegepaketManager\Function_StockComparison.ps1'
$orderedPath = 'd:\VPManagerGitHub\VerlegepaketManager\Logs\stock_compare_test\ordered_clean.csv'
$preset = [pscustomobject]@{ FileType='csv'; HeaderRowIndex=1; Delimiter=';'; Columns=[pscustomobject]@{ material_number='MAT'; quantity='QTY'; unit='UNIT' } }
$sourceData = Read-SourceRows -Path $orderedPath -Preset $preset
[pscustomobject][ordered]@{
    Headers = ($sourceData.Headers -join ',')
    RowCount = $sourceData.Rows.Count
    FirstMaterial = (Get-ColumnValue -Row $sourceData.Rows[0] -HeaderName $preset.Columns.material_number)
    FirstQty = (Get-ColumnValue -Row $sourceData.Rows[0] -HeaderName $preset.Columns.quantity)
    FirstUnit = (Get-ColumnValue -Row $sourceData.Rows[0] -HeaderName $preset.Columns.unit)
} | ConvertTo-Json -Compress
