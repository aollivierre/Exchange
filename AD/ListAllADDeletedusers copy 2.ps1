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

# Filter for users whose DistinguishedName contains the last name "Gauthier"
$filteredUsers = $deletedUsers | Where-Object { $_.DistinguishedName -like "*Gauthier*" }

# Check if any filtered users were found
if ($filteredUsers.Count -eq 0) {
    Write-Host "No deleted users found with the last name 'Gauthier'."
} else {
    # Generate a unique filename with timestamp
    $timestamp = Get-Date -Format "yyyyMMddHHmmss"
    $outputPath = "C:\Code\AD\exports\DeletedUsers_Gauthier_$timestamp.xml"

    # Export the filtered users' attributes to XML
    $filteredUsers | Export-Clixml -Path $outputPath

    Write-Host "Filtered deleted users' attributes have been exported to: $outputPath"
}
