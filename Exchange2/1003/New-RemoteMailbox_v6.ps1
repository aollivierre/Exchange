<#
.SYNOPSIS
Creates a remote shared mailbox in Exchange.

.DESCRIPTION
This function creates a remote shared mailbox in Exchange. It takes several parameters, including the name of the mailbox, the user principal name, and the organizational unit where the mailbox will be created.

.PARAMETER Name
The name of the mailbox.

.PARAMETER FirstName
The first name of the mailbox user.

.PARAMETER LastName
The last name of the mailbox user.

.PARAMETER UserPrincipalName
The user principal name of the mailbox.

.PARAMETER OnPremisesOrganizationalUnit
The organizational unit where the mailbox will be created.

.PARAMETER RemoteRoutingAddress
The remote routing address of the mailbox.

.EXAMPLE
CreateRemoteSharedMailbox -Name "DMARC" -FirstName "DMARC" -UserPrincipalName "DMARC@anteagroupus.onmicrosoft.com" -OnPremisesOrganizationalUnit "OU=Service Accounts,OU=CORP,DC=AnteaGroup,DC=US" -RemoteRoutingAddress "DMARC@anteagroupus.mail.onmicrosoft.com"

This example creates a remote shared mailbox named "DMARC" with the specified parameters.

.NOTES
Ensure that you are running this function on an Exchange server.
#>
function CreateRemoteSharedMailbox {
    param (
        [string]$Name = "DMARC",
        [string]$FirstName = "DMARC",
        [string]$LastName = "",
        [string]$UserPrincipalName = "DMARC@anteagroupus.onmicrosoft.com",
        [string]$OnPremisesOrganizationalUnit = "OU=Service Accounts,OU=CORP,DC=AnteaGroup,DC=US",
        [string]$RemoteRoutingAddress = "DMARC@anteagroupus.mail.onmicrosoft.com"
    )

    # Ensure the Exchange cmdlets are loaded
    if (-not (Get-Command -Name New-RemoteMailbox -ErrorAction Ignore)) {
        Write-Error "The New-RemoteMailbox cmdlet is not available. Ensure you are running this on an Exchange server."
        return
    }

    # Function to get the number of users in the OU
    function Get-NumberOfUsersInOU {
        param (
            [string]$OU
        )
        return (Get-ADUser -Filter * -SearchBase $OU).Count
    }

    function Get-MailboxCounts {
        $totalMailboxes = @(Get-Mailbox -ResultSize Unlimited).Count
        $totalRemoteMailboxes = @(Get-RemoteMailbox -ResultSize Unlimited).Count
        $totalRemoteSharedMailboxes = @(Get-RemoteMailbox -ResultSize Unlimited | Where-Object { $_.RecipientTypeDetails -eq 'RemoteSharedMailbox' }).Count
        return $totalMailboxes, $totalRemoteMailboxes, $totalRemoteSharedMailboxes
    }

    try {
        # Output initial stats
        $initialCount = Get-NumberOfUsersInOU -OU $OnPremisesOrganizationalUnit
        $initialTotalMailboxes, $initialTotalRemoteMailboxes, $initialTotalRemoteSharedMailboxes = Get-MailboxCounts
        Write-Host ("[" + (Get-Date) + "] Initial number of users in OU `"$OnPremisesOrganizationalUnit`": $initialCount") -ForegroundColor Cyan
        Write-Host ("[" + (Get-Date) + "] Initial total number of mailboxes: $initialTotalMailboxes") -ForegroundColor Cyan
        Write-Host ("[" + (Get-Date) + "] Initial total number of remote mailboxes: $initialTotalRemoteMailboxes") -ForegroundColor Cyan
        Write-Host ("[" + (Get-Date) + "] Initial total number of remote shared mailboxes: $initialTotalRemoteSharedMailboxes") -ForegroundColor Cyan

        # Create the remote shared mailbox
        New-RemoteMailbox -Shared -Name $Name -Firstname $FirstName -LastName $LastName `
            -UserPrincipalName $UserPrincipalName -OnPremisesOrganizationalUnit $OnPremisesOrganizationalUnit `
            -RemoteRoutingAddress $RemoteRoutingAddress

        # Check if the user and mailbox were created
        $retryCount = 0
        $createdUser = $null
        $createdMailbox = $null
        while ($null -eq $createdUser -or $null -eq $createdMailbox -and $retryCount -lt 5) {
            $createdUser = Get-ADUser -Identity $Name -ErrorAction Ignore
            $createdMailbox = Get-RemoteMailbox -Identity $Name -ErrorAction Ignore
            if ($null -ne $createdUser -and $null -ne $createdMailbox) {
                break
            }
            Start-Sleep -Seconds 10
            $retryCount++
        }

        if ($null -ne $createdUser -and $null -ne $createdMailbox) {
            # Output the user's GUID
            Write-Host ("[" + (Get-Date) + "] User GUID: " + $createdUser.ObjectGUID) -ForegroundColor Green

            # Output final stats
            $finalCount = Get-NumberOfUsersInOU -OU $OnPremisesOrganizationalUnit
            $finalTotalMailboxes, $finalTotalRemoteMailboxes, $finalTotalRemoteSharedMailboxes = Get-MailboxCounts
            Write-Host ("[" + (Get-Date) + "] Final number of users in OU `"$OnPremisesOrganizationalUnit`": $finalCount") -ForegroundColor Green
            Write-Host ("[" + (Get-Date) + "] Final total number of mailboxes: $finalTotalMailboxes") -ForegroundColor Green
            Write-Host ("[" + (Get-Date) + "] Final total number of remote mailboxes: $finalTotalRemoteMailboxes") -ForegroundColor Green
            Write-Host ("[" + (Get-Date) + "] Final total number of remote shared mailboxes: $finalTotalRemoteSharedMailboxes") -ForegroundColor Green
            Write-Host ("[" + (Get-Date) + "] Remote shared mailbox $Name created successfully.") -ForegroundColor Green
        }
        else {
            Write-Host ("[" + (Get-Date) + "] Failed to create remote shared mailbox $Name.") -ForegroundColor Red
        }
    }
    catch {
        Write-Error "An error occurred: $_"
    }
}

# Call the function
CreateRemoteSharedMailbox
