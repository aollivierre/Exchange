# Function for AD replication
function Ensure-ADReplication {
    [CmdletBinding()]
    param()
    
    Write-Verbose "Forcing AD replication across all domain controllers..."
    Get-ADDomainController -Filter * | ForEach-Object {
        $dc = $_.Name
        Write-Verbose "Replicating changes to $dc..."
        Invoke-Expression "Repadmin /syncall $dc /AeD"
    }
}

# Function to get OU suggestions
function Get-OUSuggestions {
    [CmdletBinding()]
    param()
    
    $ous = Get-ADOrganizationalUnit -Filter * | ForEach-Object {
        $enabledUsers = (Get-ADUser -Filter {Enabled -eq $true} -SearchBase $_.DistinguishedName).Count
        [PSCustomObject]@{
            Name = $_.Name
            DistinguishedName = $_.DistinguishedName
            EnabledUsers = $enabledUsers
        }
    }
    return $ous | Sort-Object EnabledUsers -Descending
}

# Function to get mailbox statistics
function Get-MailboxStats {
    [CmdletBinding()]
    param()
    
    $stats = @{
        TotalMailboxes = @(Get-Mailbox -ResultSize Unlimited).Count
        RemoteMailboxes = @(Get-RemoteMailbox -ResultSize Unlimited).Count
        RemoteSharedMailboxes = @(Get-RemoteMailbox -ResultSize Unlimited | 
            Where-Object { $_.RecipientTypeDetails -eq 'RemoteSharedMailbox' }).Count
    }
    return $stats
}

# Main function to create remote mailbox
function New-EnhancedRemoteMailbox {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Name,
        
        [Parameter(Mandatory = $true)]
        [string]$FirstName,
        
        [Parameter(Mandatory = $true)]
        [string]$LastName,
        
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName,
        
        [Parameter(Mandatory = $false)]
        [switch]$IsShared,
        
        [Parameter(Mandatory = $false)]
        [string]$OUName = "Users",  # Default OU
        
        [Parameter(Mandatory = $false)]
        [switch]$Force,

        [Parameter(Mandatory = $false)]
        [SecureString]$Password
    )

    # Verify Exchange cmdlets
    if (-not (Get-Command -Name New-RemoteMailbox -ErrorAction SilentlyContinue)) {
        throw "Exchange cmdlets are not available. Please run this script on an Exchange server."
    }

    try {
        # Get initial statistics
        $initialStats = Get-MailboxStats
        Write-Verbose "Initial mailbox statistics:`n$(($initialStats | ConvertTo-Json))"

        # Handle OU
        $ou = Get-ADOrganizationalUnit -Filter "Name -eq '$OUName'"
        if (-not $ou) {
            throw "OU '$OUName' not found"
        }

        # Check for existing objects
        $existingObject = Get-Recipient -Filter "EmailAddresses -eq 'SMTP:$UserPrincipalName'" -ErrorAction SilentlyContinue

        if ($existingObject -and -not $Force) {
            throw "An object with email address $UserPrincipalName already exists. Use -Force to remove existing object."
        }
        elseif ($existingObject -and $Force) {
            Write-Verbose "Removing existing object: $($existingObject.DisplayName)"
            
            $removeParams = @{
                Identity = $existingObject.Identity
                Confirm = $false
            }

            switch ($existingObject.RecipientType) {
                'MailUser' { Remove-MailUser @removeParams }
                'MailContact' { Remove-MailContact @removeParams }
                'UserMailbox' { Disable-Mailbox @removeParams }
                default { throw "Cannot automatically handle RecipientType: $($existingObject.RecipientType)" }
            }
            
            Ensure-ADReplication
            Start-Sleep -Seconds 5
        }

        # Create remote mailbox
        $domain = $UserPrincipalName -split "@" | Select-Object -Last 1
        $remoteRoutingAddress = "$Name@$domain.mail.onmicrosoft.com"
        
        $mailboxParams = @{
            Name = $Name
            FirstName = $FirstName
            LastName = $LastName
            UserPrincipalName = $UserPrincipalName
            OnPremisesOrganizationalUnit = $ou.DistinguishedName
            RemoteRoutingAddress = $remoteRoutingAddress
        }

        if ($IsShared) {
            $mailboxParams.Add('Shared', $true)
        }
        elseif ($Password) {
            $mailboxParams.Add('Password', $Password)
        }
        else {
            throw "Password is required for non-shared mailboxes"
        }

        $newMailbox = New-RemoteMailbox @mailboxParams
        
        # Verify creation
        $retryCount = 0
        $maxRetries = 5
        $success = $false

        do {
            Start-Sleep -Seconds 10
            $createdUser = Get-ADUser -Filter "UserPrincipalName -eq '$UserPrincipalName'" -ErrorAction SilentlyContinue
            $createdMailbox = Get-RemoteMailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue
            
            if ($createdUser -and $createdMailbox) {
                $success = $true
                break
            }
            
            $retryCount++
        } while ($retryCount -lt $maxRetries)

        if ($success) {
            $finalStats = Get-MailboxStats
            
            Write-Host "`nRemote Mailbox Creation Successful!" -ForegroundColor Green
            Write-Host "User GUID: $($createdUser.ObjectGUID)" -ForegroundColor Green
            
            $mailboxInfo = [PSCustomObject]@{
                DisplayName = $createdMailbox.DisplayName
                PrimarySmtpAddress = $createdMailbox.PrimarySmtpAddress
                UserPrincipalName = $createdMailbox.UserPrincipalName
                RemoteRoutingAddress = $createdMailbox.RemoteRoutingAddress
                RecipientType = $createdMailbox.RecipientTypeDetails
                EmailAddresses = $createdMailbox.EmailAddresses -join "; "
            }
            
            Write-Host "`nMailbox Details:" -ForegroundColor Cyan
            $mailboxInfo | Format-List
            
            Write-Host "`nMailbox Statistics:" -ForegroundColor Cyan
            Write-Host "Total Mailboxes: $($finalStats.TotalMailboxes) (Changed by: $($finalStats.TotalMailboxes - $initialStats.TotalMailboxes))"
            Write-Host "Remote Mailboxes: $($finalStats.RemoteMailboxes) (Changed by: $($finalStats.RemoteMailboxes - $initialStats.RemoteMailboxes))"
            Write-Host "Remote Shared Mailboxes: $($finalStats.RemoteSharedMailboxes) (Changed by: $($finalStats.RemoteSharedMailboxes - $initialStats.RemoteSharedMailboxes))"
        }
        else {
            throw "Failed to verify mailbox creation after $maxRetries attempts"
        }
    }
    catch {
        Write-Error "Error creating remote mailbox: $_"
        throw
    }
}

# Script to create the requested mailboxes
function Create-RequestedMailboxes {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$OUName = "Users",
        
        [Parameter(Mandatory = $true)]
        [SecureString]$SMTPPassword
    )

    # Create SMTP user mailbox
    Write-Host "Creating SMTP user mailbox..." -ForegroundColor Cyan
    
    $smtpParams = @{
        Name = "SMTP"
        FirstName = "SMTP"
        LastName = "Service"
        UserPrincipalName = "SMTP@tunngavik.com"
        OUName = $OUName
        Force = $true
        Password = $SMTPPassword
    }
    
    New-EnhancedRemoteMailbox @smtpParams

    # Create DMARC shared mailbox
    Write-Host "`nCreating DMARC shared mailbox..." -ForegroundColor Cyan
    
    $dmarcParams = @{
        Name = "DMARC"
        FirstName = "DMARC"
        LastName = "Reports"
        UserPrincipalName = "DMARC@tunngavik.com"
        OUName = $OUName
        IsShared = $true
        Force = $true
    }
    
    New-EnhancedRemoteMailbox @dmarcParams
}

# Example usage:
<#
# For secure password handling
$securePassword = Read-Host -AsSecureString -Prompt "Enter password for SMTP mailbox"

# Create both mailboxes using default OU
Create-RequestedMailboxes -SMTPPassword $securePassword

# Or specify a different OU
Create-RequestedMailboxes -OUName "CustomOU" -SMTPPassword $securePassword

# Or create them individually:
$mailboxParams = @{
    Name = "SMTP"
    FirstName = "SMTP"
    LastName = "Service"
    UserPrincipalName = "SMTP@tunngavik.com"
    Password = $securePassword
    Force = $true
}
New-EnhancedRemoteMailbox @mailboxParams

$sharedParams = @{
    Name = "DMARC"
    FirstName = "DMARC"
    LastName = "Reports"
    UserPrincipalName = "DMARC@tunngavik.com"
    IsShared = $true
    Force = $true
}
New-EnhancedRemoteMailbox @sharedParams
#>


# For secure password handling
$securePassword = Read-Host -AsSecureString -Prompt "Enter password for SMTP mailbox"

# Using default OU
Create-RequestedMailboxes -SMTPPassword $securePassword

# Or with custom OU
# Create-RequestedMailboxes -OUName "CustomOU" -SMTPPassword $securePassword