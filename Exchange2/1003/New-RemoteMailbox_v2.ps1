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

    try {
        # Create the remote shared mailbox
        New-RemoteMailbox -Shared -Name $Name -Firstname $FirstName -LastName $LastName `
            -UserPrincipalName $UserPrincipalName -OnPremisesOrganizationalUnit $OnPremisesOrganizationalUnit `
            -RemoteRoutingAddress $RemoteRoutingAddress

        Write-Output "Remote shared mailbox $Name created successfully."
    }
    catch {
        Write-Error "An error occurred: $_"
    }
}

# Call the function
CreateRemoteSharedMailbox
