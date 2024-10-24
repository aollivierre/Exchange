function Move-InactiveComputers {
    param (
        [string]$Inputcsvfile = "C:\Code\AD\exports\2024-06-04\premove_inactive_computers.csv",
        [string]$logPath = "C:\Code\AD\exports\Logs\$(Get-Date -Format 'yyyy-MM-dd-HH-mm-ss')\"
    )

    # Import the ActiveDirectory module
    Import-Module ActiveDirectory

    # Start transcript
    if (!(Test-Path $logPath)) {
        New-Item -ItemType Directory -Path $logPath | Out-Null
    }
    $logFile = "${logPath}Move_Computers.log"
    Start-Transcript -Path $logFile

    # Get the root domain distinguished name
    $rootDomain = (Get-ADDomain).DistinguishedName

    # Define the target container's distinguished name
    $targetContainer = "CN=Disabled Computers - container,$rootDomain"

    # Import the CSV file
    $computers = Import-Csv $Inputcsvfile

    # Get the total number of computers before the move
    $computersInOUBefore = (Get-ADComputer -Filter *).Count

    # Move the computers listed in the CSV file to the target container and count moved computers
    $movedCount = 0
    foreach ($computer in $computers) {
        try {
            $computerName = $computer.Name
            $computerObject = Get-ADComputer $computerName

            # Set ProtectedFromAccidentalDeletion to $false
            Set-ADObject -Identity $computerObject.DistinguishedName -ProtectedFromAccidentalDeletion $false

            Move-ADObject -Identity $computerObject -TargetPath $targetContainer

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Moved $computerName to Disabled Computers container." -ForegroundColor Green
            $movedCount++
        } catch {
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Failed to move $computerName. Error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Get the total number of computers after the move
    $computersInOUAfter = (Get-ADComputer -Filter *).Count

    # Output summary to the console
    Write-Host "`nSummary:" -ForegroundColor Yellow
    Write-Host "Total computers in OU before move: $computersInOUBefore" -ForegroundColor Cyan
    Write-Host "Total computers moved: $movedCount" -ForegroundColor Green
    Write-Host "Total computers in OU after move: $computersInOUAfter" -ForegroundColor Cyan

    # Stop transcript
    Stop-Transcript
}


Move-InactiveComputers