# # Step 1: Verify Hybrid Identity Status with Microsoft Graph API

# # Connect to Microsoft Graph
# Connect-MgGraph -Scopes "User.Read.All"

# # Get the user details
# $user = Get-MgUser -UserId "enter email address"

# if ($user) {
#     Write-Output "User found in Entra ID: $($user.DisplayName)"
# } else {
#     Write-Output "User not found in Entra ID."
#     exit
# }

# Step 2: Check On-Premises Exchange Management Server
$mailbox = Get-Mailbox -Identity "enter email address" -ErrorAction SilentlyContinue
$mailUser = Get-MailUser -Identity "enter email address" -ErrorAction SilentlyContinue

if ($mailbox) {
    Write-Output "Mailbox found: $($mailbox.PrimarySmtpAddress)"
} elseif ($mailUser) {
    Write-Output "Mail user found: $($mailUser.PrimarySmtpAddress)"
} else {
    Write-Output "No mailbox or mail user found for enter email address."
}

# Step 3: Identify Disabled Mailboxes (correct the recipient type)
$disabledMailbox = Get-User -RecipientTypeDetails DisabledUser -Identity "enter email address" -ErrorAction SilentlyContinue

if ($disabledMailbox) {
    Write-Output "Disabled user found: $($disabledMailbox.PrimarySmtpAddress)"
    Enable-Mailbox -Identity "enter email address"
    Write-Output "Mailbox enabled."
} else {
    Write-Output "No disabled mailbox found."
}

# Step 4: Identify Existing Object and Resolve Conflict
if (-not $mailbox -and -not $mailUser) {
    # Find the existing object with the UPN
    $existingObject = Get-Recipient -Filter "EmailAddresses -eq 'SMTP:enter email address'"

    if ($existingObject) {
        Write-Output "Existing object found with UPN: $($existingObject.DisplayName), Type: $($existingObject.RecipientType)"

        # Remove the existing object if appropriate
        if ($existingObject.RecipientType -eq 'MailUser' -or $existingObject.RecipientType -eq 'MailContact') {
            try {
                Remove-MailUser -Identity $existingObject.Identity -Confirm:$false
                Write-Output "Removed existing MailUser or MailContact: $($existingObject.DisplayName)"
            } catch {
                Write-Output "Failed to remove MailUser or MailContact: $($_.Exception.Message)"
                exit
            }
        } elseif ($existingObject.RecipientType -eq 'User') {
            try {
                Disable-Mailbox -Identity $existingObject.Identity -Confirm:$false
                Write-Output "Disabled existing User mailbox: $($existingObject.DisplayName)"
            } catch {
                Write-Output "Failed to disable User mailbox: $($_.Exception.Message)"
                exit
            }
        } else {
            Write-Output "Cannot automatically resolve conflict for RecipientType: $($existingObject.RecipientType). Manual intervention required."
            exit
        }
    }

    # Create a new remote mailbox
    try {
        New-RemoteMailbox -Name "JHeuser" -Alias "JHeuser" -UserPrincipalName "enter email address" -PrimarySmtpAddress "enter email address" -RemoteRoutingAddress "JHeuser@contoso.com.mail.onmicrosoft.com"
        Write-Output "New remote mailbox created."
    } catch {
        Write-Output "Failed to create new remote mailbox: $($_.Exception.Message)"
    }
}

# # Step 5: Verify the Mailbox in Exchange Online
# $mailboxOnline = Get-Mailbox -Identity "enter email address" -ErrorAction SilentlyContinue

# if ($mailboxOnline) {
#     Write-Output "Mailbox found in Exchange Online: $($mailboxOnline.PrimarySmtpAddress)"
# } else {
#     Write-Output "No mailbox found in Exchange Online."
# }

# # Step 6: Import PST File Using Microsoft Purview
# # Ensure you have the PST file available and the necessary permissions
# $importRequest = New-MailboxImportRequest -Mailbox "enter email address" -FilePath "\\path\to\pst\file.pst"
# Write-Output "Mailbox import request created. Request ID: $($importRequest.RequestGuid)"
