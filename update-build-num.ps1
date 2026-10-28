param([string]$Version, [string]$BuildNo, [string]$RcFile)
$full = "$Version.$BuildNo"
$commas = $full -replace '\.', ','
$content = Get-Content $RcFile -Raw
$content = $content -replace '(FILEVERSION|PRODUCTVERSION) \d+,\d+,\d+,\d+', "`$1 $commas"
$content = $content -replace '"(FileVersion|ProductVersion)", "[\d.]+"', "`"`$1`", `"$full`""
$content | Set-Content $RcFile -NoNewline
Write-Host "Updated to $full"
