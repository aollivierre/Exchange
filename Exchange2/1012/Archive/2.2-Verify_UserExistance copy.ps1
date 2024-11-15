$users = Import-Csv -Path "C:\Code\CB\Exchange\Glebe\Exports\Glebe_Exchange_Remote_Move_10pilotusers_Import.csv"

foreach ($user in $users) {
    $emailAddress = $user.EmailAddress  

    $exchangeUser = $null
    $ErrorActionPreference = 'Stop'
    $err = $null

    try {
        $exchangeUser = Get-ExoMailbox -Identity $emailAddress
    } catch {
        Write-Output "Encountered an error: $($err[0])"
    }
    $ErrorActionPreference = 'Continue'

    if ($exchangeUser) {
        Write-Output "User $emailAddress exists in Exchange Online."
    } else {
        Write-Output "User $emailAddress does not exist in Exchange Online."
    }
}