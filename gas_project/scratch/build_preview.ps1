$index = Get-Content 'gas_project\Index.html' -Raw -Encoding UTF8
$styles = Get-Content 'gas_project\Styles.html' -Raw -Encoding UTF8
$js = Get-Content 'gas_project\JavaScript.html' -Raw -Encoding UTF8

$combined = $index.Replace("<?!= include('Styles'); ?>", $styles).Replace("<?!= include('JavaScript'); ?>", $js)
Set-Content -Path 'gas_project\scratch\preview.html' -Value $combined -Encoding UTF8
Write-Output "Preview HTML compiled at gas_project\scratch\preview.html"
