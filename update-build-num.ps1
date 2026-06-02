param([string]$BuildNo, [string]$RcFile)
$content = Get-Content $RcFile -Raw
$content = $content -replace '1,0,16,\d+', "1,0,16,$BuildNo"
$content | Set-Content $RcFile -NoNewline
Write-Host "Updated to $BuildNo"