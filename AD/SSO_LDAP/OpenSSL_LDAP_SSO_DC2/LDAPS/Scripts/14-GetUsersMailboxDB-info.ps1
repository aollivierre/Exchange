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

# Load Exchange module - Only needed if you're not already in the Exchange Management Shell
# Import-Module $env:ExchangeInstallPath\bin\RemoteExchange.ps1
# Connect-ExchangeServer -auto

Write-TimestampedMessage "Retrieving all mailboxes and their databases..." -Color Cyan

$mailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object DisplayName, Database

# Export to CSV
$csvPath = "C:\Code\Exchange\Exports\all_mailboxes.csv" # Change path_to_save to your desired location

$mailboxes | Out-GridView
$mailboxes | Export-Csv $csvPath -NoTypeInformation

Write-TimestampedMessage "Data exported to $csvPath" -Color Green