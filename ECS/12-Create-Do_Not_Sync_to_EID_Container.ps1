# Import the ActiveDirectory module
Import-Module ActiveDirectory

# Get the root domain distinguished name
$rootDomain = (Get-ADDomain).DistinguishedName

# Find the distinguished name of the "DoNotSynctoEID" container
$DoNotSynctoEIDContainerDN = (Get-ADObject -Filter {Name -eq "DoNotSynctoEID" -and ObjectClass -eq "container"}).DistinguishedName

# Check if the "DoNotSynctoEID" container exists
if (!$DoNotSynctoEIDContainerDN) {
    Write-Host "The 'DoNotSynctoEID' container does not exist in the root domain. Creating it now..."
    # Create the new container called "DoNotSynctoEID" in the root domain
    New-ADObject -Type container -Name "DoNotSynctoEID" -Path $rootDomain

    # Get the distinguished name of the "DoNotSynctoEID" container
    $DoNotSynctoEIDContainerDN = (Get-ADObject -Filter {Name -eq "DoNotSynctoEID" -and ObjectClass -eq "container"}).DistinguishedName
}