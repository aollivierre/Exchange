function Fix-BlankDisplayName {
    param (
        [string]$CsvPath
    )

    $data = Import-Csv -Path $CsvPath

    foreach ($row in $data) {
        if ($row.ATTRIBUTE -eq "displayName" -and $row.ERROR -eq "Blank") {
            $dn = $row.DISTINGUISHEDNAME
            $cn = $row.COMMONNAME
            Write-Host "Updating displayName for: $dn" -ForegroundColor Yellow
            Set-ADObject -Identity $dn -Replace @{displayName = $cn}
        }
    }

    Write-Host "Completed updating displayName attributes." -ForegroundColor Green
}


# # Import the functions
# . .\PathToYourScript.ps1

# Fix blank displayName attributes
Fix-BlankDisplayName -CsvPath "C:\Code\AD\IdFix\IdFix-ARH-May-24-2024-report.csv"

