Get-MailboxDatabase | Format-Table -Auto Name, Server, EdbFilePath
$Database = (Get-MailboxDatabase "AGH-DB-01").Name



function CreateOnPremMailbox {
    param(
        [string]$UserPrincipalName = "FortiMailLDAPTest001@agh-fvm.com",
        [string]$Name = "FortiMailLDAPTest001",
        [string]$PasswordText = "YOUR STRONG Password HERE",
        [string]$OrganizationalUnit = "OU=Network_Services,OU=AGHUsers,DC=AGH,DC=com",
        [string]$Database = "AGH-DB-01"
    )

    # Helper function to write color-coded, timestamped messages
    function Write-TimestampedMessage {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Message,
            [Parameter(Mandatory = $false)]
            [ConsoleColor]$Color = 'White'
        )

        $timestamp = Get-Date -Format "HH:mm:ss"
        Write-Host "[$timestamp] $Message" -ForegroundColor $Color
    }

    # Convert password to SecureString
    $Password = ConvertTo-SecureString -String $PasswordText -AsPlainText -Force

    # Load Exchange module - Only needed if you're not already in the Exchange Management Shell
    # Import-Module $env:ExchangeInstallPath\bin\RemoteExchange.ps1
    # Connect-ExchangeServer -auto

    # Try to create the mailbox
    try {
        Write-TimestampedMessage "Attempting to create mailbox: $UserPrincipalName..." -Color Cyan
        New-Mailbox -UserPrincipalName $UserPrincipalName -Name $Name -Password $Password -OrganizationalUnit $OrganizationalUnit -Database $Database

        # Confirm creation
        $mailbox = Get-Mailbox -Identity $UserPrincipalName
        if ($mailbox) {
            Write-TimestampedMessage "Successfully created mailbox: $UserPrincipalName." -Color Green
        } else {
            Write-TimestampedMessage "Failed to confirm creation of mailbox: $UserPrincipalName." -Color Red
        }
    } catch {
        Write-TimestampedMessage "Error: $_" -Color Red
    }
}

# Call the function
CreateOnPremMailbox
