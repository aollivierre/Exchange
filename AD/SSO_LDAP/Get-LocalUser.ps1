#because lusrmgr.msc will return an error on a domain controller, you can instead use the following PowerShell Script

# Get all local users
$localUsers = Get-LocalUser

# Get the user details and sort by date created
$sortedUsers = $localUsers | ForEach-Object {
    $userDetails = [PSCustomObject]@{
        Name         = $_.Name
        Enabled      = $_.Enabled
        AccountExpires = $_.AccountExpires
        Description  = $_.Description
        FullName     = $_.FullName
        LastLogon    = $_.LastLogon
        PasswordChangeableDate = $_.PasswordChangeableDate
        PasswordExpires = $_.PasswordExpires
        PrincipalSource = $_.PrincipalSource
        SID          = $_.SID
        UserFlags    = $_.UserFlags
        Created      = $_.PasswordLastSet
    }
    $userDetails
} | Sort-Object Created

# Export the sorted data to a CSV file on the desktop
$csvPath = [Environment]::GetFolderPath("Desktop") + "\LocalUsers.csv"
$sortedUsers | Export-Csv -Path $csvPath -NoTypeInformation
