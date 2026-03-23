. 'd:\VPManagerGitHub\VerlegepaketManager\Function_StockComparison.ps1'
$orderedPath = 'd:\VPManagerGitHub\VerlegepaketManager\Logs\stock_compare_test\ordered_clean.csv'
$preset = [pscustomobject]@{ FileType='csv'; HeaderRowIndex=1; Delimiter=';'; Columns=[pscustomobject]@{ material_number='MAT'; quantity='QTY'; unit='UNIT' } }
$sourceData = Read-SourceRows -Path $orderedPath -Preset $preset
[pscustomobject][ordered]@{
    RowsType = $sourceData.Rows.GetType().FullName
    IsArray = ($sourceData.Rows -is [array])
    CountProp = $sourceData.Rows.PSObject.Properties['Count'] -ne $null
    CountValue = $sourceData.Rows.Count
    RowNumbersType = $sourceData.RowNumbers.GetType().FullName
    RowNumbersCount = $sourceData.RowNumbers.Count
} | ConvertTo-Json -Compress
