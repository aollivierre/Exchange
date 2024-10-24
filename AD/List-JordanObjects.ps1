# Function to display detailed information of AD objects
function Display-ADObjectDetails {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Identity
    )

    try {
        $adUser = Get-ADUser -Filter { Name -like "*Jordan Heuser*" -or UserPrincipalName -like "*Jordan Heuser*"} -Properties *
        if ($adUser) {
            Write-Host "AD User:" -ForegroundColor Green
            $adUser | Format-List | Out-String | Write-Host
        } else {
            Write-Host "No AD User found for $Identity" -ForegroundColor Red
        }

        $mailUser = Get-ADObject -Filter { ObjectClass -eq "user" -and ProxyAddresses -like "*Jordan Heuser*"} -Properties *
        if ($mailUser) {
            Write-Host "AD Mail User:" -ForegroundColor Green
            $mailUser | Format-List | Out-String | Write-Host
        } else {
            Write-Host "No AD Mail User found for $Identity" -ForegroundColor Red
        }

        $adContact = Get-ADObject -Filter { ObjectClass -eq "contact" -and ProxyAddresses -like "*Jordan Heuser*"} -Properties *
        if ($adContact) {
            Write-Host "AD Contact:" -ForegroundColor Green
            $adContact | Format-List | Out-String | Write-Host
        } else {
            Write-Host "No AD Contact found for $Identity" -ForegroundColor Red
        }
    } catch {
        Write-Host "Error retrieving details for $Identity $($_.Exception.Message)" -ForegroundColor Red
    }
}

# List all AD objects for Jordan Heuser
$identity = "Jordan Heuser"
Display-ADObjectDetails -Identity $identity
