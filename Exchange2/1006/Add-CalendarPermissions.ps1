# Read the input CSV file
$members = Import-Csv -Path "C:\Exports\_External_chfc_board_members.csv"

# Add external users as members to Azure AD without sending invitation emails
foreach ($member in $members) {
    New-AzureADMSInvitation -InvitedUserEmailAddress $member.PrimarySmtpAddress -SendInvitationMessage $false -InvitedUserType "Guest"
}

# AFTER inviting the following 4 users as External Guest users in Entra.microsoft.com
# WAIT for 15 minutes or then disconnect and re-connect to Exchange Online

Connect-ExchangeOnline
Add-MailboxFolderPermission -Identity "BoardRoomBookings@chfcanada.coop:\calendar" -User "lizette@rooftops.ca" -AccessRights Editor
Add-MailboxFolderPermission -Identity "BoardRoomBookings@chfcanada.coop:\calendar" -User "rebecca@rooftops.ca" -AccessRights Editor
Add-MailboxFolderPermission -Identity "BoardRoomBookings@chfcanada.coop:\calendar" -User "efemena@rooftops.ca" -AccessRights Editor
Add-MailboxFolderPermission -Identity "BoardRoomBookings@chfcanada.coop:\calendar" -User "giselle@rooftops.ca" -AccessRights Editor


# then we gave my admin account in CHFC full admin delegate permissions from Exchange Online into the Room Mailbox then access the Mailbox from OWA (open another mailbox top right corner under the profile name)
# then clicked calendars
# then removed all 4 external guest users who were added as view only
# then left the ones we added through PowerShell as edit only


#example if you do NOT wait enough time after inviting the external guest users to Azure AD/Entra ID

# C:\Code> Add-MailboxFolderPermission -Identity "BoardRoomBookings@chfcanada.coop:\calendar" -User "giselle@rooftops.ca" -AccessRights Editor
# Add-MailboxFolderPermission: |Microsoft.Exchange.Management.StoreTasks.InvalidExternalUserIdException|The user "giselle@rooftops.ca" is either not valid
# SMTP address, or there is no matching information.



#example if you wait enough time after inviting the external guest users to Azure AD/Entra ID


# PS C:\Code> Add-MailboxFolderPermission -Identity "BoardRoomBookings@chfcanada.coop:\calendar" -User "efemena@rooftops.ca" -AccessRights Editor

# FolderName           User                 AccessRights                               SharingPermissionFlags
# ----------           ----                 ------------                               ----------------------
# Calendar             Efemena Ozhuga       {Editor}

# PS C:\Code> Add-MailboxFolderPermission -Identity "BoardRoomBookings@chfcanada.coop:\calendar" -User "giselle@rooftops.ca" -AccessRights Editor

# FolderName           User                 AccessRights                               SharingPermissionFlags
# ----------           ----                 ------------                               ----------------------
# Calendar             Giselle Del Rosario  {Editor}
