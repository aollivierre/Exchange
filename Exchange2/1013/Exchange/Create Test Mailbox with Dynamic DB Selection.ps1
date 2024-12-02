# Get most populated database
$targetDB = Get-MailboxDatabase | Sort-Object DatabaseSize -Descending | Select-Object -First 1 -ExpandProperty Name

# Create unique identifier
$uniqueId = "BTM" + (Get-Date -Format "MMddHHmm")
$mailboxParams = @{
    Name = "BitTitanTest_$uniqueId"
    FirstName = "BitTitan"
    LastName = "Test_$uniqueId"
    DisplayName = "BitTitan Test User $uniqueId"
    UserPrincipalName = "bittitantest_$uniqueId@tunngavik.com"
    Password = (ConvertTo-SecureString -String "P@ssw0rd123!" -AsPlainText -Force)
    ResetPasswordOnNextLogon = $false
    OrganizationalUnit = "ott.nti.local/NTIOTT/Staff/Test_OU"
}

# Create mailbox in most populated DB
New-Mailbox @mailboxParams -Database $targetDB

# Configure mailbox
Set-Mailbox "BitTitanTest_$uniqueId" -EmailAddresses @{
    add = "SMTP:bittitantest_$uniqueId@tunngavik.com"
} -IssueWarningQuota 9GB -ProhibitSendQuota 10GB -ProhibitSendReceiveQuota 11GB

# Enable Archive
Enable-Mailbox "BitTitanTest_$uniqueId" -Archive

# Add test emails
1..10 | ForEach-Object {
    Send-MailMessage -To "bittitantest_$uniqueId@tunngavik.com" -From "sender@tunngavik.com" `
    -Subject "Test Email $_" -Body "This is test email number $_" -SmtpServer "localhost"
}