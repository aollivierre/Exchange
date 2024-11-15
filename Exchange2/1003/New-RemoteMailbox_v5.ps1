# Check if running as Administrator
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    # Get the ID and security principal of the current user account
    $myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
    $myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)

    # Get the security principal for the administrator role
    $adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator

    # Check to see if we are currently running "as Administrator"
    if ($myWindowsPrincipal.IsInRole($adminRole))
    {
        # We are running "as Administrator" - so change the title and background color to indicate this
        $Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)"
        $Host.UI.RawUI.BackgroundColor = "DarkBlue"
        clear-host
    }
    else {
        # We are not running "as Administrator" - so relaunch as administrator
        # Create a new process object that starts PowerShell
        $newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell";
        # Specify the current script path and name as a parameter
        $newProcess.Arguments = $myInvocation.MyCommand.Definition;
        # Indicate that the process should be elevated
        $newProcess.Verb = "runas";
        # Start the new process
        [System.Diagnostics.Process]::Start($newProcess);
        # Exit from the current, unelevated, process
        exit
    }
}

# Your script code goes here
Write-Output "Running as Administrator"


function CreateRemoteSharedMailbox {
    param (
        [string]$Name = "DMARC3",
        [string]$FirstName = "DMARC3",
        [string]$LastName = "Mailbox3",
        [string]$UserPrincipalName = "dmarc3@anteagroupus.onmicrosoft.com",
        [string]$OnPremisesOrganizationalUnit = "OU=Service Accounts,OU=CORP,DC=AnteaGroup,DC=US",
        [string]$RemoteRoutingAddress = "dmarc3@anteagroupus.mail.onmicrosoft.com"
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

        # Wait for a while to ensure that the user and mailbox are created
        Start-Sleep -Seconds 10

        # Check if the user and mailbox were created
        $createdUser = Get-ADUser -Identity $Name -ErrorAction Ignore
        $createdMailbox = Get-RemoteMailbox -Identity $Name -ErrorAction Ignore

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
