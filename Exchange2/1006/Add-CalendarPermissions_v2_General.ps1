# INC 2035741 in Ivanti

# https://learn.microsoft.com/en-us/answers/questions/432464/how-to-give-access-to-a-shared-mailbox-for-a-guest?comment=question#newest-question-comment


# Looks like its possible for Room mailboxes. Happy PS!

# This script helps in inviting users as External Guests to an Azure AD domain and then provides them with calendar permissions.

# Step 1: Import members from a CSV file
# Replace the path below with the path to your CSV file
$members = Import-Csv -Path "C:\Exports\_External_abc_organization_members.csv"

# Step 2: Invite each member to Azure AD without sending an invitation email
foreach ($member in $members) {
    New-AzureADMSInvitation -InvitedUserEmailAddress $member.PrimaryEmail -SendInvitationMessage $false -InvitedUserType "Guest"
}

# Step 3: After inviting all users, connect to Exchange Online
# It's recommended to wait for a while after the previous step before executing this one.
Connect-ExchangeOnline

# Step 4: Add calendar permissions for specific users
# Replace the email addresses below with the emails of the members you want to provide permissions to
$usersToProvidePermissions = @("john@sampledomain.com", "jane@sampledomain.com", "doe@sampledomain.com", "alice@sampledomain.com")

foreach ($user in $usersToProvidePermissions) {
    Add-MailboxFolderPermission -Identity "MeetingRoomBookings@abcorg.com:\calendar" -User $user -AccessRights Editor
}

# Optional Steps:
# If you need to assign admin delegate permissions and then access the mailbox, follow the steps below.
# 1. Provide your admin account full admin delegate permissions from Exchange Online to the Room Mailbox.
# 2. Access the mailbox from OWA (open another mailbox from the top right corner under the profile name).
# 3. Click on calendars.
# 4. Remove any unwanted external guest users who were previously added.
# 5. Ensure the ones added through PowerShell are retained.

# Note:
# If you see errors similar to "InvalidExternalUserIdException" for a user, it usually indicates that there was not enough waiting time after inviting the user to Azure AD.
# It's recommended to ensure adequate waiting time to prevent such issues.