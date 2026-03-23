$root = 'd:\VPManagerGitHub\VerlegepaketManager\Logs\stock_compare_test'
New-Item -Path $root -ItemType Directory -Force | Out-Null
$dbPath = Join-Path $root 'db.json'
$stockPath = Join-Path $root 'stock.csv'
$orderedPath = Join-Path $root 'ordered.csv'

$db = [pscustomobject][ordered]@{
    schema_version = 2
    lookup_file = 'db_lookups.json'
    materials = @(
        [pscustomobject][ordered]@{
            id = 1001
            canonical_key = 'matnr:100'
            primary_identifier = [pscustomobject][ordered]@{ type='matnr'; value='100' }
            identifiers = [pscustomobject][ordered]@{
                matnr = '100'
                supply_number = 'SUP-100'
                article_number = $null
                nato_stock_number = 'NSN-100'
            }
            status = [pscustomobject][ordered]@{ material_status_code = '11' }
            texts = [pscustomobject][ordered]@{ short_description='Hydraulic oil'; technical_note=''; logistics_note='' }
            classification = [pscustomobject][ordered]@{ ext_wg=''; is_decentral=$false; creditor=$null }
            hazmat = [pscustomobject][ordered]@{ is_hazardous=$false; un_number=$null; flags=@() }
            quantity = [pscustomobject][ordered]@{
                base_unit = 'EA'
                target = 50.0
                alternate_units = @(
                    [pscustomobject][ordered]@{ unit_code='BX'; conversion_to_base=10.0 }
                )
            }
            alternates = @(
                [pscustomobject][ordered]@{
                    position = 1
                    identifier = [pscustomobject][ordered]@{ type='matnr'; value='101' }
                    material_status_code = '12'
                    preferred_unit_code = 'BX'
                }
            )
            assignments = [pscustomobject][ordered]@{ responsibility_codes=@(); assignment_tags=@() }
        }
    )
}
$db | ConvertTo-Json -Depth 20 | Set-Content -Path $dbPath -Encoding UTF8

$stockCsv = @"
MAT;QTY;UNIT;DESC
101;2;BX;alt box row
101;1;;alt blank unit row
101;1;ZZ;alt invalid unit row
999;9;EA;unknown row
"@
Set-Content -Path $stockPath -Value $stockCsv -Encoding UTF8

$orderedCsv = @"
MAT;QTY;UNIT
100;15;EA
"@
Set-Content -Path $orderedPath -Value $orderedCsv -Encoding UTF8

. 'd:\VPManagerGitHub\VerlegepaketManager\Function_StockComparison.ps1'
$result = Invoke-StockComparison -DatabasePath $dbPath -StockRoleName 'stock' -SourceSpecs @(
    [pscustomobject]@{
        RoleName = 'stock'
        Path = $stockPath
        Preset = [pscustomobject]@{
            Name = 'inline-stock'
            FileType = 'csv'
            HeaderRowIndex = 1
            Delimiter = ';'
            Columns = [pscustomobject]@{
                material_number = 'MAT'
                quantity = 'QTY'
                unit = 'UNIT'
                description = 'DESC'
            }
        }
    },
    [pscustomobject]@{
        RoleName = 'ordered'
        Path = $orderedPath
        Preset = [pscustomobject]@{
            Name = 'inline-ordered'
            FileType = 'csv'
            HeaderRowIndex = 1
            Delimiter = ';'
            Columns = [pscustomobject]@{
                material_number = 'MAT'
                quantity = 'QTY'
                unit = 'UNIT'
            }
        }
    }
)
$row = $result.Rows[0]
Write-Output (([pscustomobject][ordered]@{
    Parse = 'OK'
    Stock = $row.StockQuantityBase
    Inbound = $row.InboundQuantityBase
    Available = $row.AvailableQuantityBase
    Missing = $row.MissingToOrderBase
    UnknownCount = $result.Diagnostics.unknown_sap_materials.Count
    InvalidUnitCount = $result.Diagnostics.invalid_units.Count
    MatchAliases = ($row.MatchedAliases -join ',')
    MatchedRows = $row.MatchedSourceRowCount
}) | ConvertTo-Json -Compress)
