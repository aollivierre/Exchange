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
        [string]$OUName,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    # Verify Exchange cmdlets
    if (-not (Get-Command -Name New-RemoteMailbox -ErrorAction SilentlyContinue)) {
        throw "Exchange cmdlets are not available. Please run this script on an Exchange server."
    }

    try {
        # Get initial statistics
        $initialStats = Get-MailboxStats
        Write-Verbose "Initial mailbox statistics:`n$(($initialStats | ConvertTo-Json))"

        # Handle OU selection
        if (-not $OUName) {
            Write-Host "`nAvailable OUs (sorted by number of enabled users):" -ForegroundColor Cyan
            $ouSuggestions = Get-OUSuggestions
            $ouSuggestions | Format-Table Name, EnabledUsers -AutoSize
            $OUName = Read-Host "Enter the desired OU name from the list above"
        }

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
            switch ($existingObject.RecipientType) {
                'MailUser' { Remove-MailUser -Identity $existingObject.Identity -Confirm:$false }
                'MailContact' { Remove-MailContact -Identity $existingObject.Identity -Confirm:$false }
                'UserMailbox' { Disable-Mailbox -Identity $existingObject.Identity -Confirm:$false }
                default { throw "Cannot automatically handle RecipientType: $($existingObject.RecipientType)" }
            }
            
            Ensure-ADReplication
            Start-Sleep -Seconds 5
        }

        # Create remote mailbox
        $domain = $UserPrincipalName -split "@" | Select-Object -Last 1
        $remoteRoutingAddress = "$Name@$domain.mail.onmicrosoft.com"
        
        $params = @{
            Name = $Name
            FirstName = $FirstName
            LastName = $LastName
            UserPrincipalName = $UserPrincipalName
            OnPremisesOrganizationalUnit = $ou.DistinguishedName
            RemoteRoutingAddress = $remoteRoutingAddress
        }

        if ($IsShared) {
            $params.Add('Shared', $true)
        }

        $newMailbox = New-RemoteMailbox @params
        
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
        [Parameter(Mandatory = $true)]
        [string]$OUName
    )

    # Create SMTP user mailbox
    Write-Host "Creating SMTP user mailbox..." -ForegroundColor Cyan
    New-EnhancedRemoteMailbox `
        -Name "SMTP" `
        -FirstName "SMTP" `
        -LastName "Service" `
        -UserPrincipalName "SMTP@tunngavik.com" `
        -OUName $OUName `
        -Force

    # Create DMARC shared mailbox
    Write-Host "`nCreating DMARC shared mailbox..." -ForegroundColor Cyan
    New-EnhancedRemoteMailbox `
        -Name "DMARC" `
        -FirstName "DMARC" `
        -LastName "Reports" `
        -UserPrincipalName "DMARC@tunngavik.com" `
        -OUName $OUName `
        -IsShared `
        -Force
}

# Example usage - uncomment and modify the OU name as needed:
Create-RequestedMailboxes -OUName "YourOUName"




# # Create a regular remote mailbox
# New-EnhancedRemoteMailbox -Name "jsmith" -FirstName "John" -LastName "Smith" -UserPrincipalName "jsmith@contoso.com"

# # Create a shared remote mailbox
# New-EnhancedRemoteMailbox -Name "accounting" -FirstName "Accounting" -LastName "Department" -UserPrincipalName "accounting@contoso.com" -IsShared

# # Force creation even if object exists
# New-EnhancedRemoteMailbox -Name "jsmith" -FirstName "John" -LastName "Smith" -UserPrincipalName "jsmith@contoso.com" -Force