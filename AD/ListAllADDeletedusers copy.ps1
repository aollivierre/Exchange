# Import the Active Directory module
Import-Module ActiveDirectory

# Check if the Active Directory Recycle Bin is enabled
$recycleBinEnabled = Get-ADOptionalFeature -Filter { Name -like "Recycle Bin Feature" }

if (-not $recycleBinEnabled) {
    Write-Host "The Active Directory Recycle Bin is not enabled. Please enable it to retrieve deleted objects."
    exit
}

# Retrieve all deleted user objects from the AD Recycle Bin
$deletedUsers = Get-ADObject -Filter {
    IsDeleted -eq $true -and
    ObjectClass -eq "user"
} -IncludeDeletedObjects -Property *

# Check if any deleted users were found
if ($deletedUsers.Count -eq 0) {
    Write-Host "No deleted users found."
} else {
    # Generate a unique filename with timestamp
    $timestamp = Get-Date -Format "yyyyMMddHHmmss"
    $outputPath = "C:\Code\AD\exports\DeletedUsers_$timestamp.xml"

    # Export the deleted users' attributes to XML
    $deletedUsers | Export-Clixml -Path $outputPath

    Write-Host "Deleted users' attributes have been exported to: $outputPath"
}
