# Function to ensure replication across all domain controllers using Repadmin
function Ensure-ADReplication {
    Write-Host "Forcing AD replication across all domain controllers..." -ForegroundColor Yellow
    Get-ADDomainController -Filter * | ForEach-Object {
        $dc = $_.Name
        Write-Host "Replicating changes to $dc..." -ForegroundColor Yellow
        Invoke-Expression "Repadmin /syncall $dc /AeD"
    }
}

# Step 2: Check On-Premises Exchange Management Server
$mailbox = Get-Mailbox -Identity "enter email address" -ErrorAction SilentlyContinue
$mailUser = Get-MailUser -Identity "enter email address" -ErrorAction SilentlyContinue

if ($mailbox) {
    Write-Host "Mailbox found: $($mailbox.PrimarySmtpAddress)" -ForegroundColor Green
} elseif ($mailUser) {
    Write-Host "Mail user found: $($mailUser.PrimarySmtpAddress)" -ForegroundColor Green
} else {
    Write-Host "No mailbox or mail user found for enter email address." -ForegroundColor Red
}

# Step 3: Identify Disabled Mailboxes (correct the recipient type)
$disabledMailbox = Get-User -RecipientTypeDetails DisabledUser -Identity "enter email address" -ErrorAction SilentlyContinue

if ($disabledMailbox) {
    Write-Host "Disabled user found: $($disabledMailbox.PrimarySmtpAddress)" -ForegroundColor Yellow
    Enable-Mailbox -Identity "enter email address"
    Write-Host "Mailbox enabled." -ForegroundColor Green
} else {
    Write-Host "No disabled mailbox found." -ForegroundColor Red
}

# Step 4: Create a New Remote Mailbox if No Existing Mailbox or Mail User Found
if (-not $mailbox -and -not $mailUser -and -not $disabledMailbox) {
    try {
        New-RemoteMailbox -Name "Jordan Heuser" -Alias "JHeuser" -UserPrincipalName "enter email address" -PrimarySmtpAddress "enter email address" -RemoteRoutingAddress "JHeuser@contoso.com.mail.onmicrosoft.com"
        Write-Host "New remote mailbox created." -ForegroundColor Green
    } catch {
        Write-Host "Failed to create new remote mailbox: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Ensure replication across all domain controllers
Ensure-ADReplication