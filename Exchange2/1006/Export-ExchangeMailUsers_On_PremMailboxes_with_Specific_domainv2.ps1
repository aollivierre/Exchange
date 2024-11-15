function Export-EOMailUsers {
    # Define the CSV file path
    $csvPath = "C:\Code\CB\Exchange\CHFC\Exports2\CHFC_MailUsers_Feb_06_2024_Primary_domain.csv"

    # Retrieve mail users, filter by primary SMTP address domain, and select relevant properties
    $mailUsers = Get-MailUser -ResultSize Unlimited | Where-Object {
        $_.PrimarySmtpAddress -like "*@chfcanada.coop"
    } | Select-Object DisplayName, PrimarySmtpAddress, @{Name='EmailAddresses';Expression={
        $_.EmailAddresses | Where-Object {$_ -like "SMTP:*"} | ForEach-Object { $_ -replace "SMTP:","" }
    }}, ExternalEmailAddress, WhenCreated

    # Output to GridView
    $mailUsers | Out-GridView -Title "Exchange Online Mail Users with Domain @chfcanada.coop"

    # Export to CSV
    $mailUsers | Export-Csv -Path $csvPath -NoTypeInformation

    # Display total count
    $totalCount = $mailUsers.Count
    Write-Host "Total Mail Users with Domain @chfcanada.coop Exported: $totalCount"
    
    # Return total count for further use if needed
    return $totalCount
}
Export-EOMailUsers