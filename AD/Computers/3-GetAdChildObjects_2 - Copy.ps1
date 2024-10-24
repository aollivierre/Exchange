function Get-ADComputerChildObjects {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$CSVPath
    )

    if (!(Test-Path $CSVPath)) {
        Write-Host "CSV file not found at the specified path." -ForegroundColor Red
        return
    }

    Import-Module ActiveDirectory

    $tempFolderPath = [System.IO.Path]::GetTempPath()
    $logFilePath = Join-Path -Path $tempFolderPath -ChildPath "ADComputerChildObjects.log"

    $exportsFolderPath = "C:\Code\CB\AD\Exports"
    $csvOutputPath = Join-Path -Path $exportsFolderPath -ChildPath "ADComputerChildObjects_post_removing.csv"

    Start-Transcript -Path $logFilePath -Append

    $computers = Import-Csv -Path $CSVPath
    $results = @()

    foreach ($computer in $computers) {
        $computerName = $computer.Name

        try {
            $adComputer = Get-ADComputer -Identity $computerName -ErrorAction Stop
            $childObjects = Get-ADObject -Filter * -SearchBase ($adComputer.DistinguishedName) -ErrorAction Stop

            if ($childObjects) {
                Write-Host "Child objects for ($computerName):" -ForegroundColor Cyan
                $childObjects | Format-Table -Property Name, ObjectClass, DistinguishedName

                foreach ($childObject in $childObjects) {
                    $results += [PSCustomObject]@{
                        ComputerName = $computerName
                        ChildObjectName = $childObject.Name
                        ChildObjectClass = $childObject.ObjectClass
                        ChildObjectDistinguishedName = $childObject.DistinguishedName
                    }
                }
            } else {
                Write-Host "No child objects found for $computerName." -ForegroundColor Green
            }
        } catch [System.Exception] {
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Error processing ($computerName): $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    $results | Export-Csv -Path $csvOutputPath -NoTypeInformation

    Stop-Transcript
}

# Get-ADComputerChildObjects -CSVPath "C:\Code\CB\AD\Exports\Disable\AD_LHC_post_removing_computers_report.csv"
Get-ADComputerChildObjects -CSVPath "C:\Code\AD\exports\2024-06-04\premove_inactive_computers.csv"
