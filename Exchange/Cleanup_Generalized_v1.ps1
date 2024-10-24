# Exchange Mailbox Management and Migration Script

# Function to trigger and monitor AD Connect sync
function Start-ADSyncAndWait {
    param (
        [string]$Operation = "Operation"
    )
    Write-Host "Initiating Entra Connect Sync for $Operation..." -ForegroundColor Yellow
    try {
        # Start the sync
        Start-ADSyncSyncCycle -PolicyType Delta
        
        # Wait for initial sync to kick off
        Start-Sleep -Seconds 5
        
        # Monitor sync status
        $attempts = 0
        $maxAttempts = 12  # 2 minutes total waiting time
        
        do {
            $syncStatus = Get-ADSyncConnectorRunStatus
            
            if ($null -eq $syncStatus -or $syncStatus.Count -eq 0) {
                # No status means sync is complete
                Write-Host "Entra Connect Sync completed successfully." -ForegroundColor Green
                break
            } elseif ($syncStatus.RunState -eq "Busy") {
                $attempts++
                Write-Host "Sync in progress... Attempt $attempts of $maxAttempts" -ForegroundColor Yellow
                Start-Sleep -Seconds 10
            }
            
            if ($attempts -ge $maxAttempts) {
                Write-Host "Warning: Sync monitoring timed out. Please verify sync status manually." -ForegroundColor Yellow
                break
            }
        } while ($true)
        
    } catch {
        Write-Host "Error triggering Entra Connect Sync: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Function to ensure replication across all domain controllers using Repadmin
function Ensure-ADReplication {
    Write-Host "Forcing AD replication across all domain controllers..." -ForegroundColor Yellow
    Get-ADDomainController -Filter * | ForEach-Object {
        $dc = $_.Name
        Write-Host "Replicating changes to $dc..." -ForegroundColor Yellow
        Invoke-Expression "Repadmin /syncall $dc /AeD"
    }
}

# Function to remove the existing object from all domain controllers with retry
function Remove-ExistingObject {
    param (
        [string]$ObjectDN
    )
    Write-Host "Removing existing object from all domain controllers..." -ForegroundColor Yellow
    Get-ADDomainController -Filter * | ForEach-Object {
        $dc = $_.Name
        $attempts = 0
        $maxAttempts = 3
        while ($attempts -lt $maxAttempts) {
            try {
                Write-Host "Attempting to remove object from $dc (Attempt $($attempts + 1)/$maxAttempts)..." -ForegroundColor Yellow
                Remove-ADObject -Identity $ObjectDN -Server $dc -Confirm:$false -ErrorAction Stop
                Write-Host "Successfully removed object from $dc." -ForegroundColor Green
                break
            } catch {
                Write-Host "Failed to remove object from $dc $($_.Exception.Message)" -ForegroundColor Red
                $attempts++
                if ($attempts -ge $maxAttempts) {
                    Write-Host "Max attempts reached for $dc. Moving on..." -ForegroundColor Red
                }
            }
        }
    }
}


# Function to export user details in multiple formats
function Export-UserDetails {
    param (
        [string]$UserEmail,
        [switch]$ShowGridView = $true
    )
    
    try {
        # Get AD User details with all properties
        $adUser = Get-ADUser -Filter {(UserPrincipalName -eq $UserEmail) -or (mail -eq $UserEmail)} -Properties * -ErrorAction SilentlyContinue
        
        $userDetails = @{
            # Exchange Details
            Mailbox = Get-Mailbox -Identity $UserEmail -ErrorAction SilentlyContinue
            MailUser = Get-MailUser -Identity $UserEmail -ErrorAction SilentlyContinue
            User = Get-User -Identity $UserEmail -ErrorAction SilentlyContinue
            Recipient = Get-Recipient -Identity $UserEmail -ErrorAction SilentlyContinue
            
            # AD Profile Details
            ProfileDetails = [PSCustomObject]@{
                HomeDirectory = $adUser.HomeDirectory
                HomeDrive = $adUser.HomeDrive
                ProfilePath = $adUser.ProfilePath
                ScriptPath = $adUser.ScriptPath
                TerminalServicesHomeDirectory = $adUser.TerminalServicesHomeDirectory
                TerminalServicesHomeDrive = $adUser.TerminalServicesHomeDrive
                TerminalServicesProfilePath = $adUser.TerminalServicesProfilePath
                UserPrincipalName = $adUser.UserPrincipalName
                DistinguishedName = $adUser.DistinguishedName
                SamAccountName = $adUser.SamAccountName
                Department = $adUser.Department
                Title = $adUser.Title
                Manager = $adUser.Manager
                Company = $adUser.Company
                Office = $adUser.Office
                StreetAddress = $adUser.StreetAddress
                City = $adUser.City
                State = $adUser.State
                PostalCode = $adUser.PostalCode
                Country = $adUser.Country
                telephoneNumber = $adUser.telephoneNumber
                MobilePhone = $adUser.MobilePhone
                Fax = $adUser.Fax
                Enabled = $adUser.Enabled
                LastLogonDate = $adUser.LastLogonDate
                Created = $adUser.Created
                Modified = $adUser.Modified
                GroupMemberships = ($adUser.memberOf -join '; ')
            }
        }

        # Create timestamp and base filename
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $baseFilename = "UserDetails-$($UserEmail.Split('@')[0])-$timestamp"
        
        # Export to XML
        $xmlPath = "$baseFilename.xml"
        $userDetails | Export-Clixml -Path $xmlPath
        Write-Host "User details exported to XML: $xmlPath" -ForegroundColor Green

        # Create CSV-friendly object
        $csvObject = [PSCustomObject]@{
            # User Information
            UserPrincipalName = $adUser.UserPrincipalName
            SamAccountName = $adUser.SamAccountName
            DisplayName = $adUser.DisplayName
            Enabled = $adUser.Enabled
            
            # Profile Paths
            HomeDirectory = $adUser.HomeDirectory
            HomeDrive = $adUser.HomeDrive
            ProfilePath = $adUser.ProfilePath
            ScriptPath = $adUser.ScriptPath
            
            # Terminal Services
            TSHomeDirectory = $adUser.TerminalServicesHomeDirectory
            TSHomeDrive = $adUser.TerminalServicesHomeDrive
            TSProfilePath = $adUser.TerminalServicesProfilePath
            
            # Contact Information
            Department = $adUser.Department
            Title = $adUser.Title
            Manager = $adUser.Manager
            Company = $adUser.Company
            Office = $adUser.Office
            StreetAddress = $adUser.StreetAddress
            City = $adUser.City
            State = $adUser.State
            PostalCode = $adUser.PostalCode
            Country = $adUser.Country
            PhoneNumber = $adUser.telephoneNumber
            MobilePhone = $adUser.MobilePhone
            Fax = $adUser.Fax
            
            # Account Details
            LastLogonDate = $adUser.LastLogonDate
            Created = $adUser.Created
            Modified = $adUser.Modified
            DistinguishedName = $adUser.DistinguishedName
            GroupMemberships = ($adUser.memberOf -join '; ')
            
            # Exchange Details
            MailboxEnabled = ($null -ne $userDetails.Mailbox)
            MailUserEnabled = ($null -ne $userDetails.MailUser)
            RecipientType = $userDetails.Recipient.RecipientType
        }

        # Export to CSV
        $csvPath = "$baseFilename.csv"
        $csvObject | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Host "User details exported to CSV: $csvPath" -ForegroundColor Green

        # Display in GridView if requested
        if ($ShowGridView) {
            $properties = $csvObject | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
            $gridViewData = @()
            
            foreach ($prop in $properties) {
                $gridViewData += [PSCustomObject]@{
                    Property = $prop
                    Value = $csvObject.$prop
                }
            }
            
            $gridViewData | Out-GridView -Title "User Details for $UserEmail"
        }

        # Display to console
        Write-Host "`nUser Details:" -ForegroundColor Cyan
        
        # Display Exchange details
        Write-Host "`nExchange Details:" -ForegroundColor Yellow
        $userDetails.GetEnumerator() | Where-Object { $_.Key -ne 'ProfileDetails' } | ForEach-Object {
            Write-Host "`n$($_.Key):" -ForegroundColor Yellow
            $_.Value | Format-List
        }
        
        # Display Profile details
        Write-Host "`nProfile Details:" -ForegroundColor Yellow
        $userDetails.ProfileDetails | Format-List
        
        # Display Group Memberships
        Write-Host "`nGroup Memberships:" -ForegroundColor Yellow
        $adUser.memberOf | ForEach-Object {
            Write-Host "  $_"
        }
        
        return $userDetails
    }
    catch {
        Write-Host "Error exporting user details: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# Example usage
$userEmail = Read-Host "Enter user email address"
Write-Host "Exporting user details for $userEmail..." -ForegroundColor Yellow
Export-UserDetails -UserEmail $userEmail -ShowGridView $true


# Prompt for user information
$example = @"
Example input:
Email: jgauthier@contoso.com
Name: John Gauthier
Alias: jgauthier
Domain: contoso.com
"@

Write-Host "`n$example`n" -ForegroundColor Cyan
Write-Host "Please enter the user information:" -ForegroundColor Green

$userEmail = Read-Host "Email address"
$userName = Read-Host "Full Name"
$userAlias = Read-Host "Alias"
$domain = $userEmail.Split('@')[1]
$remoteRoutingAddress = "$userAlias@$($domain.Split('.')[0]).mail.onmicrosoft.com"

# Export existing user details if available
$existingUserDetails = Export-UserDetails -UserEmail $userEmail

# Step 1: Check On-Premises Exchange Management Server
$mailbox = Get-Mailbox -Identity $userEmail -ErrorAction SilentlyContinue
$mailUser = Get-MailUser -Identity $userEmail -ErrorAction SilentlyContinue

if ($mailbox) {
    Write-Host "Mailbox found: $($mailbox.PrimarySmtpAddress)" -ForegroundColor Green
} elseif ($mailUser) {
    Write-Host "Mail user found: $($mailUser.PrimarySmtpAddress)" -ForegroundColor Green
} else {
    Write-Host "No mailbox or mail user found for $userEmail." -ForegroundColor Red
}

# Step 2: Identify Disabled Mailboxes
$disabledMailbox = Get-User -RecipientTypeDetails DisabledUser -Identity $userEmail -ErrorAction SilentlyContinue

if ($disabledMailbox) {
    Write-Host "Disabled user found: $($disabledMailbox.PrimarySmtpAddress)" -ForegroundColor Yellow
    $enableConfirm = Read-Host "Do you want to enable this mailbox? (Y/N)"
    if ($enableConfirm -eq 'Y') {
        Enable-Mailbox -Identity $userEmail
        Write-Host "Mailbox enabled." -ForegroundColor Green
        
        # Trigger AD sync after enabling mailbox
        Start-ADSyncAndWait -Operation "mailbox enabling"
    }
} else {
    Write-Host "No disabled mailbox found." -ForegroundColor Red
}

# Step 3: Identify Existing Object and Resolve Conflict
if (-not $mailbox -and -not $mailUser) {
    $existingObject = Get-Recipient -Filter "EmailAddresses -eq 'SMTP:$userEmail'"

    if ($existingObject) {
        Write-Host "Existing object found with UPN: $($existingObject.DisplayName), Type: $($existingObject.RecipientType)" -ForegroundColor Yellow
        $removeConfirm = Read-Host "Do you want to remove this existing object? (Y/N)"
        
        if ($removeConfirm -eq 'Y') {
            try {
                $ObjectDN = $existingObject.DistinguishedName
                Remove-ExistingObject -ObjectDN $ObjectDN
                Write-Host "Removed existing object: $($existingObject.DisplayName)" -ForegroundColor Green
                
                # Trigger AD sync after removal
                Start-ADSyncAndWait -Operation "object removal"
            } catch {
                Write-Host "Failed to remove existing object: $($_.Exception.Message)" -ForegroundColor Red
                exit
            }
        } else {
            Write-Host "Operation cancelled by user." -ForegroundColor Yellow
            exit
        }
    }

    # Ensure replication across all domain controllers
    Ensure-ADReplication

    # Recheck and remove existing object if it still exists
    $existingObject = Get-Recipient -Filter "EmailAddresses -eq 'SMTP:$userEmail'"
    if ($existingObject) {
        Write-Host "Existing object still found after replication. Manual intervention required." -ForegroundColor Red
        exit
    }

    # Create a new remote mailbox
    $createConfirm = Read-Host "Do you want to create a new remote mailbox? (Y/N)"
    if ($createConfirm -eq 'Y') {
        try {
            New-RemoteMailbox -Name $userName -Alias $userAlias -UserPrincipalName $userEmail -PrimarySmtpAddress $userEmail -RemoteRoutingAddress $remoteRoutingAddress
            Write-Host "New remote mailbox created." -ForegroundColor Green
            
            # Trigger AD sync after creation
            Start-ADSyncAndWait -Operation "mailbox creation"
            
            # Export new user details
            Write-Host "`nExporting new user details..." -ForegroundColor Yellow
            Export-UserDetails -UserEmail $userEmail
        } catch {
            Write-Host "Failed to create new remote mailbox: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

Write-Host "`nScript execution completed." -ForegroundColor Green