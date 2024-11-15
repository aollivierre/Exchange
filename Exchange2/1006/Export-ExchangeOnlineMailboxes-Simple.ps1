# Connect-ExchangeOnline

function Export-EOmailboxes {
    # Define the CSV file path
    $csvPath = "C:\Code\CB\Exchange\CHFC\Exports2\CHFC_EXOMailboxes_Feb_06_2024.csv"
    
    # Retrieve mailboxes and select relevant properties
    $mailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object DisplayName,PrimarySmtpAddress, @{Name='EmailAddresses';Expression={$_.EmailAddresses | Where-Object {$_ -like "SMTP:*"} | ForEach-Object { $_ -replace "SMTP:","" }}}, ArchiveStatus, ProhibitSendQuota

    # Output to GridView
    $mailboxes | Out-GridView -Title "Exchange Online Mailboxes"

    # Export to CSV
    $mailboxes | Export-Csv -Path $csvPath -NoTypeInformation

    # Display total count
    $totalCount = $mailboxes.Count
    Write-Host "Total Mailboxes Exported: $totalCount"
    
    # Return total count for further use if needed
    return $totalCount
}


function Export-EOMailUsers {
    # Define the CSV file path
    $csvPath = "C:\Code\CB\Exchange\CHFC\Exports2\CHFC_MailUsers_Feb_06_2024.csv"

    # Retrieve mail users and select relevant properties
    $mailUsers = Get-MailUser -ResultSize Unlimited | Select-Object DisplayName,PrimarySmtpAddress, @{Name='EmailAddresses';Expression={$_.EmailAddresses | Where-Object {$_ -like "SMTP:*"} | ForEach-Object { $_ -replace "SMTP:","" }}}, ExternalEmailAddress, WhenCreated

    # Output to GridView
    $mailUsers | Out-GridView -Title "Exchange Online Mail Users"

    # Export to CSV
    $mailUsers | Export-Csv -Path $csvPath -NoTypeInformation

    # Display total count
    $totalCount = $mailUsers.Count
    Write-Host "Total Mail Users Exported: $totalCount"
    
    # Return total count for further use if needed
    return $totalCount
}



function Export-EOMailContacts {
    # Define the CSV file path
    $csvPath = "C:\Code\CB\Exchange\CHFC\Exports2\CHFC_Feb_06_2024_MailContacts.csv"

    # Retrieve mail contacts and select relevant properties
    $mailContacts = Get-MailContact -ResultSize Unlimited | Select-Object DisplayName,PrimarySmtpAddress, @{Name='EmailAddresses';Expression={$_.EmailAddresses | Where-Object {$_ -like "SMTP:*"} | ForEach-Object { $_ -replace "SMTP:","" }}}, @{Name='ContactType';Expression={'MailContact'}}, WhenCreated

    # Output to GridView
    $mailContacts | Out-GridView -Title "Exchange Online Mail Contacts"

    # Export to CSV
    $mailContacts | Export-Csv -Path $csvPath -NoTypeInformation

    # Display total count
    $totalCount = $mailContacts.Count
    Write-Host "Total Mail Contacts Exported: $totalCount"
    
    # Return total count for further use if needed
    return $totalCount
}


Export-EOmailboxes
Export-EOMailUsers

Export-EOMailContacts


