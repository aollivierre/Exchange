<#
.SYNOPSIS
    Removes a remote shared mailbox and its associated user account.

.DESCRIPTION
    This script removes a remote shared mailbox and its associated user account. It first attempts to get the remote mailbox using the provided name. If the mailbox is found, it is removed. Then, it waits for a while to ensure that the mailbox is removed. Next, it attempts to get the user using the provided name. If the user is found, it is removed.

.PARAMETER Name
    The name of the mailbox and user to remove.

.EXAMPLE
    RemoveRemoteSharedMailbox -Name "SharedMailbox01"
#>

function RemoveRemoteSharedMailbox {
    $Name = Read-Host -Prompt "Enter the name for the mailbox"

    # Ensure the Exchange cmdlets are loaded
    if (-not (Get-Command -Name Remove-RemoteMailbox -ErrorAction Ignore)) {
        Write-Error "The Remove-RemoteMailbox cmdlet is not available. Ensure you are running this on an Exchange server."
        return
    }

    try {
        # Attempt to get the remote mailbox
        $remoteMailbox = Get-RemoteMailbox -Identity $Name -ErrorAction Ignore
        if ($null -ne $remoteMailbox) {
            # Remove the remote mailbox
            Remove-RemoteMailbox -Identity $Name -Confirm:$false
            Write-Host ("[" + (Get-Date) + "] Removed remote shared mailbox $Name successfully.") -ForegroundColor Green
        }
        else {
            Write-Host ("[" + (Get-Date) + "] No remote shared mailbox found with name $Name.") -ForegroundColor Yellow
        }

        # Wait for a while to ensure that the mailbox is removed
        Start-Sleep -Seconds 10

        # Attempt to get the user
        $user = Get-ADUser -Identity $Name -ErrorAction Ignore
        if ($null -ne $user) {
            # Remove the user
            Remove-ADUser -Identity $Name -Confirm:$false
            Write-Host ("[" + (Get-Date) + "] Removed user $Name successfully.") -ForegroundColor Green
        }
        else {
            Write-Host ("[" + (Get-Date) + "] No user found with name $Name.") -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "An error occurred: $_"
    }
}

# Call the function
RemoveRemoteSharedMailbox
