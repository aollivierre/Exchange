# Set up transcript logging
$logPath = "C:\Code\AD\Exports\Logs\$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')_Disable_Computers.log"
if (!(Test-Path $logPath)) {
    New-Item -ItemType File -Path $logPath -Force | Out-Null
}
Start-Transcript -Path $logPath

# Import the CSV file
$csvFile = "C:\Code\AD\exports\2024-06-04\premove_inactive_computers.csv"
$inactiveComputers = Import-Csv $csvFile

# Disable each computer in the CSV file
$disabledCount = 0
foreach ($computer in $inactiveComputers) {
    try {
        $adComputer = Get-ADComputer -Identity $computer.Name
        Disable-ADAccount -Identity $adComputer
        Write-Host ("{0} - Successfully disabled {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $adComputer.Name) -ForegroundColor Green
        $disabledCount++
    } catch {
        Write-Host ("{0} - Failed to disable {1}: {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $adComputer.Name, $_) -ForegroundColor Red
    }
}

# Output the total number of disabled computers to the console
Write-Host ("{0} - Total computers disabled: {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $disabledCount) -ForegroundColor Cyan

# Stop transcript logging
Stop-Transcript