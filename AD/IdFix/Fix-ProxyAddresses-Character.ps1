function Fix-ProxyAddresses {
    param (
        [string]$CsvPath
    )

    $data = Import-Csv -Path $CsvPath

    foreach ($row in $data) {
        if ($row.ATTRIBUTE -eq "proxyAddresses" -and $row.ERROR -eq "Character") {
            $dn = $row.DISTINGUISHEDNAME
            $proxyAddresses = (Get-ADObject -Identity $dn -Properties proxyAddresses).proxyAddresses

            $newProxyAddresses = @()
            foreach ($address in $proxyAddresses) {
                $newAddress = $address -replace "'", ""
                $newProxyAddresses += $newAddress
            }

            Write-Host "Updating proxyAddresses for: $dn" -ForegroundColor Yellow
            Set-ADObject -Identity $dn -Replace @{proxyAddresses = $newProxyAddresses}
        }
    }

    Write-Host "Completed updating proxyAddresses attributes." -ForegroundColor Green
}


# Import the functions
# . .\PathToYourScript.ps1

# Fix proxyAddresses with character errors
# Fix-ProxyAddresses -CsvPath "C:\Code\AD\IdFix\IdFix-ARH-May-24-2024-report.csv"
Fix-ProxyAddresses -CsvPath "C:\Code\AD\IdFix\IdFix-ARH-May-24-2024-report-character.csv"

