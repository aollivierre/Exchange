<#
.SYNOPSIS
Creates a remote shared mailbox in Exchange.

.DESCRIPTION

1. The code defines a PowerShell function called `CreateRemoteSharedMailbox`.
2. The function prompts the user for input, including the name, first name, last name, user principal name, domain, and organizational unit (OU) name.
3. The function automatically determines the OU distinguished name from the provided OU name.
4. The function checks if the Exchange cmdlets are loaded and if the New-RemoteMailbox cmdlet is available.
5. The function outputs initial statistics, including the number of users in the OU and the total number of mailboxes, remote mailboxes, and remote shared mailboxes.
6. The function creates the remote shared mailbox using the New-RemoteMailbox cmdlet.
7. The function checks if the user and mailbox were created by retrying a few times and waiting 10 seconds between retries.
8. If the user and mailbox were created successfully, the function outputs the user's GUID and final statistics, including the number of users in the OU and the total number of mailboxes, remote mailboxes, and remote shared mailboxes.
9. If the creation was not successful, the function outputs an error message.

.PARAMETER None

.EXAMPLE
CreateRemoteSharedMailbox

This example creates a remote shared mailbox in Exchange.

.NOTES
Author: AOllivierre
Date: Nov 15, 2023
#>
function CreateRemoteSharedMailbox {
    # Interactive prompts for user input
    $Name = Read-Host -Prompt "Enter the name for the mailbox"
    $FirstName = Read-Host -Prompt "Enter the first name (e.g., Accounting) "
    $LastName = Read-Host -Prompt "Enter the last name (e.g., Department) "
    $UserPrincipalName = Read-Host -Prompt "Enter the User Principal Name (e.g., user@domain.com)"
    $domain = $UserPrincipalName -split "@" | Select-Object -Last 1
    $OUName = Read-Host -Prompt "Enter the Organizational Unit (OU) name"

    # Automatically determine OU DN from provided OU name
    $OnPremisesOrganizationalUnit = (Get-ADOrganizationalUnit -Filter "Name -eq '$OUName'").DistinguishedName
    $RemoteRoutingAddress = "$Name@$domain.mail.onmicrosoft.com"

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