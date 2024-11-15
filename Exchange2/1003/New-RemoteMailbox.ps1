function CreateRemoteMailbox {
    param (
        [string]$MailboxName = "dmarc",
        [string]$Database,
        [string]$RemoteRoutingAddress
    )

    # Import the Exchange cmdlets if not already loaded
    if (-not (Get-Command -Name New-RemoteMailbox -ErrorAction Ignore)) {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
    }

    try {
        # Check if the mailbox already exists
        $mailbox = Get-RemoteMailbox -Identity $MailboxName -ErrorAction Ignore
        if ($null -ne $mailbox) {
            Write-Output "Mailbox $MailboxName already exists."
            return
        }

        # Create the remote mailbox
        New-RemoteMailbox -Name $MailboxName -Database $Database -RemoteRoutingAddress $RemoteRoutingAddress

        Write-Output "Remote mailbox $MailboxName created successfully."
    }
    catch {
        Write-Error "An error occurred: $_"
    }
}
CreateRemoteMailbox -Database "YourDatabase" -RemoteRoutingAddress "anteagroupus.onmicrosoft.com"