function RemoveRemoteSharedMailbox {
    param (
        [string]$Name = "DMARC",
        [string]$OnPremisesOrganizationalUnit = "OU=Service Accounts,OU=CORP,DC=AnteaGroup,DC=US"
    )

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
