function Export-OnPremMailUsersWithDomain {
    # Define the CSV file path
    $csvPath = "C:\Code\CB\Exchange\CHFC\Exports2\CHFC_OnPremMailUsers_with_Primary_Domain_Feb_06_2024.csv"

    # Retrieve on-premises mail users with the primary domain @chfcanada.coop
    $onPremMailUsers = Get-Recipient -ResultSize Unlimited | Where-Object {
        $_.RecipientTypeDetails -eq "RemoteUserMailbox" -and
        $_.PrimarySmtpAddress -like "*@chfcanada.coop"
    } | Select-Object DisplayName,PrimarySmtpAddress, @{Name='EmailAddresses';Expression={$_.EmailAddresses | Where-Object {$_ -like "SMTP:*"} | ForEach-Object { $_ -replace "SMTP:","" }}}, RecipientTypeDetails, WhenCreated

    # Output to GridView
    $onPremMailUsers | Out-GridView -Title "On-Premises Mail Users with Domain @chfcanada.coop"

    # Export to CSV
    $onPremMailUsers | Export-Csv -Path $csvPath -NoTypeInformation

    # Display total count
    $totalCount = $onPremMailUsers.Count
    Write-Host "Total On-Premises Mail Users Exported: $totalCount"
    
    # Return total count for further use if needed
    return $totalCount
}
Export-OnPremMailUsersWithDomain