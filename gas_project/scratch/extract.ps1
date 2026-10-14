$csv = Import-Csv -Path 'sheet_1609045579.csv'
$seen = @{}
$list = [System.Collections.Generic.List[PSObject]]::new()
foreach ($row in $csv) {
    $c = $row.'ITEM CODE'
    $d = $row.'ITEM DESCRIPTION'
    if ($c -and $d) {
        $c = $c.Trim()
        $d = $d.Trim()
        if (-not $seen.ContainsKey($c)) {
            $seen[$c] = $true
            $list.Add([PSCustomObject]@{ code = $c; desc = $d })
        }
    }
}
Write-Output "Unique items count: $($list.Count)"
$list | ConvertTo-Json -Compress | Set-Content -Path 'gas_project\scratch\mock_catalog.json' -Encoding UTF8
