# Function to ensure replication across all domain controllers using Repadmin
function Ensure-ADReplication {
    Write-host "Forcing AD replication across all domain controllers..."
    Get-ADDomainController -Filter * | ForEach-Object {
        $dc = $_.Name
        Write-host "Replicating changes to $dc..."
        Invoke-Expression "Repadmin /syncall $dc /AeD"
    }
}

# Function to remove the existing object from all domain controllers
function Remove-ExistingObject {
    param (
        [string]$ObjectDN
    )
    Write-host "Removing existing object from all domain controllers..."
    Get-ADDomainController -Filter * | ForEach-Object {
        $dc = $_.Name
        Write-host "Attempting to remove object from $dc..."
        try {
            Remove-ADObject -Identity $ObjectDN -Server $dc -Confirm:$false -ErrorAction Stop
            Write-host "Successfully removed object from $dc."
        } catch {
            Write-host "Failed to remove object from $dc $($_.Exception.Message)"
        }
    }
}

# # Step 1: Verify Hybrid Identity Status with Microsoft Graph API

# # Connect to Microsoft Graph
# Connect-MgGraph -Scopes "User.Read.All"

# # Get the user details
# $user = Get-MgUser -UserId "enter email address"

# if ($user) {
#     Write-host "User found in Entra ID: $($user.DisplayName)"
# } else {
#     Write-host "User not found in Entra ID."
#     exit
# }

# Step 2: Check On-Premises Exchange Management Server
$mailbox = Get-Mailbox -Identity "enter email address" -ErrorAction SilentlyContinue
$mailUser = Get-MailUser -Identity "enter email address" -ErrorAction SilentlyContinue

if ($mailbox) {
    Write-host "Mailbox found: $($mailbox.PrimarySmtpAddress)"
} elseif ($mailUser) {
    Write-host "Mail user found: $($mailUser.PrimarySmtpAddress)"
} else {
    Write-host "No mailbox or mail user found for enter email address."
}

# Step 3: Identify Disabled Mailboxes (correct the recipient type)
$disabledMailbox = Get-User -RecipientTypeDetails DisabledUser -Identity "enter email address" -ErrorAction SilentlyContinue

if ($disabledMailbox) {
    Write-host "Disabled user found: $($disabledMailbox.PrimarySmtpAddress)"
    Enable-Mailbox -Identity "enter email address"
    Write-host "Mailbox enabled."
} else {
    Write-host "No disabled mailbox found."
}

# Step 4: Identify Existing Object and Resolve Conflict
if (-not $mailbox -and -not $mailUser) {
    # Find the existing object with the UPN
    $existingObject = Get-Recipient -Filter "EmailAddresses -eq 'SMTP:enter email address'"

    if ($existingObject) {
        Write-host "Existing object found with UPN: $($existingObject.DisplayName), Type: $($existingObject.RecipientType)"

        # Remove the existing object if appropriate
        if ($existingObject.RecipientType -eq 'MailUser' -or $existingObject.RecipientType -eq 'MailContact' -or $existingObject.RecipientType -eq 'User') {
            try {
                $ObjectDN = $existingObject.DistinguishedName
                Remove-ExistingObject -ObjectDN $ObjectDN
                Write-host "Removed existing object: $($existingObject.DisplayName)"
            } catch {
                Write-host "Failed to remove existing object: $($_.Exception.Message)"
                exit
            }
        } else {
            Write-host "Cannot automatically resolve conflict for RecipientType: $($existingObject.RecipientType). Manual intervention required."
            exit
        }
    }

    # Ensure replication across all domain controllers
    Ensure-ADReplication

    # Recheck and remove existing object if it still exists
    $existingObject = Get-Recipient -Filter "EmailAddresses -eq 'SMTP:enter email address'"
    if ($existingObject) {
        Write-host "Existing object still found after replication. Manual intervention required."
        exit
    }

    # Create a new remote mailbox
    try {
        New-RemoteMailbox -Name "JHeuser" -Alias "JHeuser" -UserPrincipalName "enter email address" -PrimarySmtpAddress "enter email address" -RemoteRoutingAddress "JHeuser@arnpriorhealth.mail.onmicrosoft.com"
        Write-host "New remote mailbox created."
    } catch {
        Write-host "Failed to create new remote mailbox: $($_.Exception.Message)"
    }
}

# # Step 5: Verify the Mailbox in Exchange Online
# $mailboxOnline = Get-Mailbox -Identity "enter email address" -ErrorAction SilentlyContinue

# if ($mailboxOnline) {
#     Write-host "Mailbox found in Exchange Online: $($mailboxOnline.PrimarySmtpAddress)"
# } else {
#     Write-host "No mailbox found in Exchange Online."
# }

# # Step 6: Import PST File Using Microsoft Purview
# # Ensure you have the PST file available and the necessary permissions
# $importRequest = New-MailboxImportRequest -Mailbox "enter email address" -FilePath "\\path\to\pst\file.pst"
# Write-host "Mailbox import request created. Request ID: $($importRequest.RequestGuid)"
