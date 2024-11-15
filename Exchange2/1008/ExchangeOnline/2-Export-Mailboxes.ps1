Get-User -ResultSize Unlimited | Select-Object UserPrincipalName | Export-Csv -Path "$env:USERPROFILE\Desktop\UserAccounts.csv" -NoTypeInformation


$DateCutoff = (Get-Date).AddYears(-4)
Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Where-Object {$_.WhenMailboxCreated -lt $DateCutoff} | Select-Object DisplayName, UserPrincipalName, PrimarySmtpAddress, WhenMailboxCreated | Export-Csv -Path "$env:USERPROFILE\Desktop\OldSharedMailboxes2.csv" -NoTypeInformation
