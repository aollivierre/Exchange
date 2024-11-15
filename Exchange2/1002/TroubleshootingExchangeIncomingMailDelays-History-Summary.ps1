# The PowerShell history contains a mix of commands related to managing and querying Exchange Server settings, certificates, message tracking logs, mail queues, and other related tasks.

# Let's first extract unique commands and then summarize them for clarity:

# 1. Connection and Script Execution:

# Connecting to the Exchange Server.
# Running various scripts like Backup-RemoveOldExhangeCerts.ps1, Backup-RemoveOldExhangeCerts2.ps1, and so on.
# Changing directory and listing directory contents.
# 2. Message Tracking Logs:

# Retrieving message tracking logs filtered by recipients or senders.
# Formatting the output for specific columns like Timestamp, EventId, Source, etc.
# 3. Mail Queue Management:

# Retrieving mail queues.
# Retrieving messages in specific queues.
# 4. Exchange Certificate Management:

# Retrieving Exchange certificates.
# Extracting certificate details based on thumbprint.
# Checking the services enabled for a certificate.
# Backing up and removing old Exchange certificates.
# 5. Send Connector Management:

# Retrieving and setting properties of Send Connectors.
# Removing a Send Connector.
# 6. Receive Connector Management:

# Retrieving properties of Receive Connectors.
# 7. Mailbox and Remote Mailbox Queries:

# Retrieving mailboxes and remote mailboxes.
# 8. OWA (Outlook Web App) Configuration:

# Retrieving OWA virtual directory information.
# 9. Database Availability and Mailbox Database:

# Retrieving Database Availability Groups.
# Querying mailbox databases on specific servers.
# Now, let's extract the unique commands from the history:





# Connect to Exchange Server
. 'F:\Program Files\Microsoft\Exchange Server\V15in\RemoteExchange.ps1'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell

# Message Tracking Logs
Get-MessageTrackingLog -Recipients novatest@ambico.caom
Get-MessageTrackingLog -Recipients novatest@ambico.com
Get-MessageTrackingLog -Sender aollivierre@novanetworks.com
Get-MessageTrackingLog -Sender aollivierre@novanetworks.com | select eventID
Get-MessageTrackingLog -Recipients user@example.com -Start "mm/dd/yyyy hh:mm am/pm" -End "mm/dd/yyyy hh:mm am/pm" | ft Timestamp, Source, EventId, MessageSubject, Sender, Recipients -AutoSize
Get-MessageTrackingLog -Sender aollivierre@novanetworks.com -Start (Get-Date).AddDays(-30) -End (Get-Date) | Format-List Timestamp, EventId, Source, Sender, Recipients, MessageSubject, MessageInfo, ClientIp, S...

# Mail Queue
Get-Queue | ft Identity, MessageCount, Status, LastError -AutoSize

# Exchange Certificate Management
Get-ExchangeCertificate
Get-ExchangeCertificate | Format-List Thumbprint,Subject,NotAfter,Services
Get-ExchangeCertificate | Format-List Thumbprint,Subject,NotBefore,NotAfter,Services,IsSelfSigned,CertificateDomains
Get-ExchangeCertificate | Format-List Thumbprint,Subject,NotBefore,NotAfter,Services,IsSelfSigned,CertificateDomains,issuer

# Send Connector Management
Get-SendConnector "Outbound to Office 365 - 184c8e2f-b4ef-4a58-8e56-1ef196f9b744"
Get-SendConnector "Outbound to Office 365 - 184c8e2f-b4ef-4a58-8e56-1ef196f9b744" | Format-List * 
Get-SendConnector "Outbound to Office 365 - 184c8e2f-b4ef-4a58-8e56-1ef196f9b744" | Format-List Name, TlsCertificateName
Get-SendConnector "Outbound to Office 365 - 184c8e2f-b4ef-4a58-8e56-1ef196f9b744" | Format-List TlsCertificateName
Get-SendConnector "Outbound to Office 365 - 184c8e2f-b4ef-4a58-8e56-1ef196f9b744" | select * 
Set-SendConnector "Outbound to Office 365 - 184c8e2f-b4ef-4a58-8e56-1ef196f9b744" -TlsCertificateName "7FD683C1D53F595B33BD4FC8A87497EB2E659CEA"
Remove-SendConnector "Outbound to Office 365 - 184c8e2f-b4ef-4a58-8e56-1ef196f9b744"

# Receive Connector Management
Get-ReceiveConnector | Format-List Name, TlsCertificateName

# Mailbox and Remote Mailbox Queries
Get-mailbox
Get-RemoteMailbox

# OWA Configuration
Get-OwaVirtualDirectory | fl server, name, externalurl

# Database Availability
Get-DatabaseAvailabilityGroup

# Exchange Server & Mailbox Database
Get-ExchangeServer | Where-Object { $_.ServerRole -like "*Mailbox*" } | ForEach-Object { [PSCustomObject]@{ ServerName = $_.Name; DAG = $_.DatabaseAvailabilityGroup } } | Format-Table -AutoSize
Get-MailboxDatabase -Server "AMBICO-MAIL"

# Directory Commands
cd c:\code
ls

# Script Execution
.\Backup-RemoveOldExhangeCerts.ps1

# Certificate Management
$TLSCert = Get-ExchangeCertificate -Thumbprint "7FD683C1D53F595B33BD4FC8A87497EB2E659CEA"
$TLSCertName = "<I>" + $newCert.Issuer + "<S>" + $newCert.Subject
$TLSCertName = "<I>$($TLSCert.Issuer)<S>$($TLSCert.Subject)"
$certName = "<I>$($validCert.Issuer)<S>$($validCert.Subject)"
$certificate = Get-ExchangeCertificate -Thumbprint $thumbprint
$newCert = Get-ExchangeCertificate | Where-Object {$_.Thumbprint -eq "7FD683C1D53F595B33BD4FC8A87497EB2E659CEA"}
$thumbprint = "7FD683C1D53F595B33BD4FC8A87497EB2E659CEA"
$validCert = Get-ExchangeCertificate | Where-Object {$_.Thumbprint -eq "7FD683C1D53F595B33BD4FC8A87497EB2E659CEA"}
if ($certificate.Services -like "*SMTP*") {...}

# History Retrieval
Get-History
