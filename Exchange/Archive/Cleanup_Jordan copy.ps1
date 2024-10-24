# Function to ensure replication across all domain controllers using Repadmin
function Ensure-ADReplication {
    Write-Host "Forcing AD replication across all domain controllers..." -ForegroundColor Yellow
    Get-ADDomainController -Filter * | ForEach-Object {
        $dc = $_.Name
        Write-Host "Replicating changes to $dc..." -ForegroundColor Yellow
        Invoke-Expression "Repadmin /syncall $dc /AeD"
    }
}

# Function to remove the existing object from all domain controllers
function Remove-ExistingObject {
    param (
        [string]$ObjectDN
    )
    Write-Host "Removing existing object from all domain controllers..." -ForegroundColor Yellow
    Get-ADDomainController -Filter * | ForEach-Object {
        $dc = $_.Name
        Write-Host "Attempting to remove object from $dc..." -ForegroundColor Yellow
        try {
            Remove-ADObject -Identity $ObjectDN -Server $dc -Confirm:$false -ErrorAction Stop
            Write-Host "Successfully removed object from $dc." -ForegroundColor Green
        } catch {
            Write-Host "Failed to remove object from $dc $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Step 1: Verify Hybrid Identity Status with Microsoft Graph API

# # Connect to Microsoft Graph
# Connect-MgGraph -Scopes "User.Read.All"

# # Get the user details
# $user = Get-MgUser -UserId "JHeuser@arnpriorhealth.ca"

# if ($user) {
#     Write-Host "User found in Entra ID: $($user.DisplayName)" -ForegroundColor Green
# } else {
#     Write-Host "User not found in Entra ID." -ForegroundColor Red
#     exit
# }

# Step 2: Check On-Premises Exchange Management Server
$mailbox = Get-Mailbox -Identity "JHeuser@arnpriorhealth.ca" -ErrorAction SilentlyContinue
$mailUser = Get-MailUser -Identity "JHeuser@arnpriorhealth.ca" -ErrorAction SilentlyContinue

if ($mailbox) {
    Write-Host "Mailbox found: $($mailbox.PrimarySmtpAddress)" -ForegroundColor Green
} elseif ($mailUser) {
    Write-Host "Mail user found: $($mailUser.PrimarySmtpAddress)" -ForegroundColor Green
} else {
    Write-Host "No mailbox or mail user found for JHeuser@arnpriorhealth.ca." -ForegroundColor Red
}

# Step 3: Identify Disabled Mailboxes (correct the recipient type)
$disabledMailbox = Get-User -RecipientTypeDetails DisabledUser -Identity "JHeuser@arnpriorhealth.ca" -ErrorAction SilentlyContinue

if ($disabledMailbox) {
    Write-Host "Disabled user found: $($disabledMailbox.PrimarySmtpAddress)" -ForegroundColor Yellow
    Enable-Mailbox -Identity "JHeuser@arnpriorhealth.ca"
    Write-Host "Mailbox enabled." -ForegroundColor Green
} else {
    Write-Host "No disabled mailbox found." -ForegroundColor Red
}

# Step 4: Identify Existing Object and Resolve Conflict
if (-not $mailbox -and -not $mailUser) {
    # Find the existing object with the UPN
    $existingObject = Get-Recipient -Filter "EmailAddresses -eq 'SMTP:JHeuser@arnpriorhealth.ca'"

    if ($existingObject) {
        Write-Host "Existing object found with UPN: $($existingObject.DisplayName), Type: $($existingObject.RecipientType)" -ForegroundColor Yellow

        # Remove the existing object if appropriate
        if ($existingObject.RecipientType -eq 'MailUser' -or $existingObject.RecipientType -eq 'MailContact' -or $existingObject.RecipientType -eq 'User') {
            try {
                $ObjectDN = $existingObject.DistinguishedName
                Remove-ExistingObject -ObjectDN $ObjectDN
                Write-Host "Removed existing object: $($existingObject.DisplayName)" -ForegroundColor Green
            } catch {
                Write-Host "Failed to remove existing object: $($_.Exception.Message)" -ForegroundColor Red
                exit
            }
        } else {
            Write-Host "Cannot automatically resolve conflict for RecipientType: $($existingObject.RecipientType). Manual intervention required." -ForegroundColor Red
            exit
        }
    }

    # Ensure replication across all domain controllers
    Ensure-ADReplication

    # Recheck and remove existing object if it still exists
    $existingObject = Get-Recipient -Filter "EmailAddresses -eq 'SMTP:JHeuser@arnpriorhealth.ca'"
    if ($existingObject) {
        Write-Host "Existing object still found after replication. Manual intervention required." -ForegroundColor Red
        exit
    }

    # Create a new remote mailbox
    try {
        New-RemoteMailbox -Name "JHeuser" -Alias "JHeuser" -UserPrincipalName "JHeuser@arnpriorhealth.ca" -PrimarySmtpAddress "JHeuser@arnpriorhealth.ca" -RemoteRoutingAddress "JHeuser@arnpriorhealth.mail.onmicrosoft.com"
        Write-Host "New remote mailbox created." -ForegroundColor Green
    } catch {
        Write-Host "Failed to create new remote mailbox: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# # Step 5: Verify the Mailbox in Exchange Online
# $mailboxOnline = Get-Mailbox -Identity "JHeuser@arnpriorhealth.ca" -ErrorAction SilentlyContinue

# if ($mailboxOnline) {
#     Write-Host "Mailbox found in Exchange Online: $($mailboxOnline.PrimarySmtpAddress)" -ForegroundColor Green
# } else {
#     Write-Host "No mailbox found in Exchange Online." -ForegroundColor Red
# }

# # Step 6: Import PST File Using Microsoft Purview
# # Ensure you have the PST file available and the necessary permissions
# $importRequest = New-MailboxImportRequest -Mailbox "JHeuser@arnpriorhealth.ca" -FilePath "\\path\to\pst\file.pst"
# Write-Host "Mailbox import request created. Request ID: $($importRequest.RequestGuid)" -ForegroundColor Green
