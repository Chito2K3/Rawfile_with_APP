$csv = Import-Csv -Path 'sheet_1609045579.csv'
$propNames = @($csv[0].psobject.Properties | ForEach-Object { $_.Name })

$codeCol = @($propNames | Where-Object { $_ -match 'CODE' })[0]
$descCol = @($propNames | Where-Object { $_ -match 'DESC' })[0]
$qtyCol  = @($propNames | Where-Object { $_ -match 'QTY' })[0]
$expCol  = @($propNames | Where-Object { $_ -match 'EXP' })[0]

Write-Host "Matched column names:"
Write-Host "  Code: '$codeCol'"
Write-Host "  Desc: '$descCol'"
Write-Host "  Qty:  '$qtyCol'"
Write-Host "  Exp:  '$expCol'"

$data = @()
foreach ($row in $csv) {
    $code = "$($row.$codeCol)".Trim()
    if ([string]::IsNullOrWhiteSpace($code)) { continue }
    
    $desc = "$($row.$descCol)".Trim()
    $rawQty = "$($row.$qtyCol)"
    $rawExp = "$($row.$expCol)".Trim()
    
    $cleanQty = $rawQty -replace '[^0-9.]', ''
    $qty = 0
    if ($cleanQty -ne '') {
        $qty = [double]$cleanQty
    }
    
    $yr = "NO EXPIRY / UNRECORDED"
    if ($rawExp -match '([A-Za-z]+)-(\d{2})') {
        $yr = "20" + $matches[2]
    } elseif ($rawExp -match '20(\d{2})') {
        $yr = "20" + $matches[1]
    }
    
    $data += [PSCustomObject]@{
        ItemCode = $code
        ItemDesc = $desc
        Qty = $qty
        ExpDate = if ($rawExp) { $rawExp } else { "N/A" }
        ExpYear = $yr
    }
}

Write-Host "`nTotal Valid Records: $($data.Count)"
$sumQty = ($data | Measure-Object -Property Qty -Sum).Sum
Write-Host "Total Storage Inventory Quantity: $($sumQty.ToString('N0')) units"

Write-Host "`n======================================================="
Write-Host "1. BREAKDOWN BY EXPIRY YEAR"
Write-Host "======================================================="
$byYear = $data | Group-Object ExpYear | Sort-Object Name
$byYear | Select-Object @{N="Expiry Year";E={$_.Name}}, @{N="Line Items (Batches)";E={$_.Count}}, @{N="Total Quantity";E={ (($_.Group | Measure-Object -Property Qty -Sum).Sum).ToString("N0") }}, @{N="% of Total Qty";E={ "{0:P2}" -f ((($_.Group | Measure-Object -Property Qty -Sum).Sum) / $sumQty) }} | Format-Table -AutoSize

Write-Host "`n======================================================="
Write-Host "2. CRITICAL ALERT: ALL ITEMS EXPIRING IN 2026"
Write-Host "======================================================="
$exp2026 = $data | Where-Object { $_.ExpYear -eq '2026' } | Sort-Object ExpDate, Qty -Descending
$exp2026 | Select-Object ExpDate, ItemCode, ItemDesc, @{N="Qty (Units)";E={$_.Qty.ToString("N0")}} | Format-Table -AutoSize

Write-Host "`n======================================================="
Write-Host "3. HIGH RISK: ITEMS EXPIRING IN 2027 (TOP 20 BY QUANTITY)"
Write-Host "======================================================="
$exp2027 = $data | Where-Object { $_.ExpYear -eq '2027' } | Sort-Object Qty -Descending | Select-Object -First 20
$exp2027 | Select-Object ExpDate, ItemCode, ItemDesc, @{N="Qty (Units)";E={$_.Qty.ToString("N0")}} | Format-Table -AutoSize

Write-Host "`n======================================================="
Write-Host "4. TOP 15 INVENTORY ITEMS BY TOTAL QUANTITY ACROSS ALL BATCHES"
Write-Host "======================================================="
$topStock = $data | Group-Object ItemDesc | Select-Object @{N="Item Description";E={$_.Name}}, @{N="Total Qty";E={ ($_.Group | Measure-Object -Property Qty -Sum).Sum }}, @{N="Batch Count";E={$_.Count}} | Sort-Object "Total Qty" -Descending | Select-Object -First 15
$topStock | Select-Object "Item Description", @{N="Total Qty (Units)";E={$_.("Total Qty").ToString("N0")}}, "Batch Count" | Format-Table -AutoSize

Write-Host "`n======================================================="
Write-Host "5. INVENTORY CLASSIFICATION (DONATION VS REGULAR HOSPITAL PROCUREMENT)"
Write-Host "======================================================="
$data | Group-Object {
    if ($_.ItemCode -like "DMDON*") { "Donation / DOH Program (DMDON)" }
    elseif ($_.ItemCode -like "DMR*") { "Regular Hospital Stock (DMR)" }
    else { "Other Stock Code" }
} | Select-Object @{N="Category";E={$_.Name}}, @{N="Line Items";E={$_.Count}}, @{N="Total Qty (Units)";E={ (($_.Group | Measure-Object -Property Qty -Sum).Sum).ToString("N0") }}, @{N="% of Total Qty";E={ "{0:P2}" -f ((($_.Group | Measure-Object -Property Qty -Sum).Sum) / $sumQty) }} | Format-Table -AutoSize

Write-Host "`n======================================================="
Write-Host "6. ITEMS WITH MULTIPLE BATCHES / SPLIT EXPIRY DATES"
Write-Host ""
$multiBatch = $data | Group-Object ItemCode | Where-Object { $_.Count -gt 1 }
Write-Host "Found $($multiBatch.Count) items with multiple batches / entries in storage:"
foreach ($item in $multiBatch) {
    Write-Host "  * [$($item.Name)] $($item.Group[0].ItemDesc) - $($item.Count) batches:"
    foreach ($b in $item.Group) {
        Write-Host "      - Qty: $($b.Qty.ToString('N0')) | Expiry: $($b.ExpDate)"
    }
}

Write-Host "`n======================================================="
Write-Host "7. ITEMS MISSING EXPIRY DATE"
Write-Host "======================================================="
$noExp = $data | Where-Object { $_.ExpYear -eq 'NO EXPIRY / UNRECORDED' }
Write-Host "Found $($noExp.Count) items without recorded expiry date:"
$noExp | Select-Object ItemCode, ItemDesc, @{N="Qty (Units)";E={$_.Qty.ToString("N0")}} | Format-Table -AutoSize
