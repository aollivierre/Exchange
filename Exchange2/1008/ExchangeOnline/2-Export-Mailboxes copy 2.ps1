$DateCutoff = (Get-Date).AddYears(-4)
$StartDate = (Get-Date).AddYears(-10)
Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Where-Object {$_.WhenMailboxCreated -lt $DateCutoff -and $_.WhenMailboxCreated -ge $StartDate} | Select-Object DisplayName, UserPrincipalName, PrimarySmtpAddress, WhenMailboxCreated | Export-Csv -Path "$env:USERPROFILE\Desktop\OldSharedMailboxes4.csv" -NoTypeInformation
