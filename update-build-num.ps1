param([string]$BuildNo, [string]$RcFile)
$content = Get-Content $RcFile -Raw
$content = $content -replace '1,0,15,\d+', "1,0,15,$BuildNo"
$content | Set-Content $RcFile -NoNewline
Write-Host "Updated to $BuildNo"