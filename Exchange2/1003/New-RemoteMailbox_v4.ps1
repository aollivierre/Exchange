function CreateRemoteSharedMailbox {
    param (
        [string]$Name = "DMARC",
        [string]$FirstName = "DMARC",
        [string]$LastName = "Mailbox",
        [string]$UserPrincipalName = "dmarc@anteagroupus.onmicrosoft.com",
        [string]$OnPremisesOrganizationalUnit = "OU=Service Accounts,OU=CORP,DC=AnteaGroup,DC=US",
        [string]$RemoteRoutingAddress = "dmarc@anteagroupus.mail.onmicrosoft.com"
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

    try {
        # Output initial stats
        $initialCount = Get-NumberOfUsersInOU -OU $OnPremisesOrganizationalUnit
        Write-Host ("[" + (Get-Date) + "] Initial number of users in OU `"$OnPremisesOrganizationalUnit`": $initialCount") -ForegroundColor Cyan

        # Create the remote shared mailbox
        New-RemoteMailbox -Shared -Name $Name -Firstname $FirstName -LastName $LastName `
            -UserPrincipalName $UserPrincipalName -OnPremisesOrganizationalUnit $OnPremisesOrganizationalUnit `
            -RemoteRoutingAddress $RemoteRoutingAddress

        # Wait for a while to ensure that the user and mailbox are created
        Start-Sleep -Seconds 10

        # Check if the user and mailbox were created
        $createdUser = Get-ADUser -Identity $Name -ErrorAction Ignore
        $createdMailbox = Get-RemoteMailbox -Identity $Name -ErrorAction Ignore

        if ($null -ne $createdUser -and $null -ne $createdMailbox) {
            # Output final stats
            $finalCount = Get-NumberOfUsersInOU -OU $OnPremisesOrganizationalUnit
            Write-Host ("[" + (Get-Date) + "] Final number of users in OU `"$OnPremisesOrganizationalUnit`": $finalCount") -ForegroundColor Green
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
