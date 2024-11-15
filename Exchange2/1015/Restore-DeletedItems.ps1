# https://learn.microsoft.com/en-us/powershell/module/exchange/restore-recoverableitems?view=exchange-ps

# It's worth noting that the cmdlet is only available in the Mailbox Import Export role, which isn't assigned to any role groups by default. To use it, one needs to add the Mailbox Import Export role to a role group.

# It can take some time (about 1hr) for the permissions to apply. Then Go to ECA > Recipient > Mailboxes > Select the mailbox for which you want to recover deleted messages and click on the display name > under Others > Recover deleted items.

# This command offers a range of filtering options, including by item type, date, subject, and others. This flexibility allows administrators to recover specific items based on a variety of criteria.

# The provided examples show various applications of the cmdlet, from restoring a specific email in a single mailbox, to restoring a specific email in multiple mailboxes, and even to bulk restoring all recoverable items for a given administrator.

# The recent changes you're pointing to represent an improvement in the ability to recover deleted items in Exchange Online, providing administrators more power and flexibility in handling data loss situations. Thanks for bringing this new cmdlet to my attention. I'm sure it will be very useful for many Exchange Online administrators.


#Install required module
Install-Module ExchangeOnlineManagement

# Import required module
Import-Module ExchangeOnlineManagement

# Store your admin credentials
# $UserCredential = Get-Credential

# Connect to Exchange Online
# Connect-ExchangeOnline -Credential $UserCredential -ShowBanner:$false
Connect-ExchangeOnline

# Specify the mailbox, date range, item type and subject for the items you want to restore
$Identity = "tzimmer@rmhctoronto.ca"
$FilterStartTime = "5/1/2023 12:00:00 AM"
$FilterEndTime = "6/9/2023 11:59:59 PM"
$FilterItemType = "IPM.Appointment"
# $SubjectContains = "FY18 Accounting"

# Run the Restore-RecoverableItems cmdlet to restore the specified items
# Restore-RecoverableItems -Identity $Identity -FilterItemType $FilterItemType -SubjectContains $SubjectContains -FilterStartTime $FilterStartTime -FilterEndTime $FilterEndTime
Restore-RecoverableItems -Identity $Identity -FilterItemType $FilterItemType -FilterStartTime $FilterStartTime -FilterEndTime $FilterEndTime

# Disconnect the session
# Disconnect-ExchangeOnline