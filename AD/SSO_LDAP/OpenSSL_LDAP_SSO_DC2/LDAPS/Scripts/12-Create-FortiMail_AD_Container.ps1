# Import the ActiveDirectory module
Import-Module ActiveDirectory

# Get the root domain distinguished name
$rootDomain = (Get-ADDomain).DistinguishedName

# Find the distinguished name of the "FortiMail" container
$FortiMailContainerDN = (Get-ADObject -Filter {Name -eq "FortiMail" -and ObjectClass -eq "container"}).DistinguishedName

# Check if the "FortiMail" container exists
if (!$FortiMailContainerDN) {
    Write-Host "The 'FortiMail' container does not exist in the root domain. Creating it now..."
    # Create the new container called "FortiMail" in the root domain
    New-ADObject -Type container -Name "FortiMail" -Path $rootDomain

    # Get the distinguished name of the "FortiMail" container
    $FortiMailContainerDN = (Get-ADObject -Filter {Name -eq "FortiMail" -and ObjectClass -eq "container"}).DistinguishedName
}