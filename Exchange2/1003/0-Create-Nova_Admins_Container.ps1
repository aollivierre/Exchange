# Import the ActiveDirectory module
Import-Module ActiveDirectory

# Get the root domain distinguished name
$rootDomain = (Get-ADDomain).DistinguishedName

# Find the distinguished name of the "Nova_Admins_Do_NOT_AADSync" container
$Nova_Admins_Do_NOT_AADSyncContainerDN = (Get-ADObject -Filter {Name -eq "Nova_Admins_Do_NOT_AADSync" -and ObjectClass -eq "container"}).DistinguishedName

# Check if the "Nova_Admins_Do_NOT_AADSync" container exists
if (!$Nova_Admins_Do_NOT_AADSyncContainerDN) {
    Write-Host "The 'Nova_Admins_Do_NOT_AADSync' container does not exist in the root domain. Creating it now..."
    # Create the new container called "Nova_Admins_Do_NOT_AADSync" in the root domain
    New-ADObject -Type container -Name "Nova_Admins_Do_NOT_AADSync" -Path $rootDomain

    # Get the distinguished name of the "Nova_Admins_Do_NOT_AADSync" container
    $Nova_Admins_Do_NOT_AADSyncContainerDN = (Get-ADObject -Filter {Name -eq "Nova_Admins_Do_NOT_AADSync" -and ObjectClass -eq "container"}).DistinguishedName
}