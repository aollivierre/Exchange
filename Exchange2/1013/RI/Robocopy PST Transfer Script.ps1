$sourceDir = "D:\ExchangeArchives"
$destDir = "\\nti-ri-vmhost2\F$\ExchangeArchives"
$logFile = "$destDir\RobocopyLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

# Create destination if it doesn't exist
if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Force -Path $destDir }

# Robocopy with:
# /E - Copy subdirectories including empty ones
# /Z - Restartable mode for better reliability
# /ZB - Use restartable mode; if access denied use backup mode
# /R:3 - Retry 3 times
# /W:5 - Wait 5 seconds between retries
# /MT:8 - 8 threads
# /LOG - Save output to log file
# /TEE - Display output in console and log file
# /NP - No progress indicator (cleaner log)
# /NDL - No directory list (cleaner log)

robocopy $sourceDir $destDir *.pst /E /Z /ZB /R:3 /W:5 /MT:8 /LOG:$logFile /TEE /NP /NDL

if ($LASTEXITCODE -gt 7) {
    Write-Error "Robocopy failed with exit code $LASTEXITCODE"
} else {
    Write-Host "Transfer completed successfully. Log file: $logFile"
}