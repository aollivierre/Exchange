# Import the ActiveDirectory module
Import-Module ActiveDirectory

# Define the distinguished name (DN) of the target OU
$targetOUDN = "OU=Computers,OU=RAC,DC=RAILCAN,DC=CA"

# Retrieve all enabled computers running Windows 10 or Windows 11 from the domain
$targetComputers = Get-ADComputer -Filter {(Enabled -eq $True) -and (OperatingSystem -like "Windows 10*" -or OperatingSystem -like "Windows 11*")}

# Initialize a counter for the number of successfully moved computers
$movedCount = 0

foreach ($computer in $targetComputers) {
    try {
        # Move each computer to the target OU
        Move-ADObject -Identity $computer.DistinguishedName -TargetPath $targetOUDN
        Write-Host "Moved computer $($computer.Name) to the OU Computers/RAC." -ForegroundColor Green
        $movedCount++
    } catch {
        Write-Host "Failed to move computer $($computer.Name): $_" -ForegroundColor Red
    }
}

# Output the total count of successfully moved computers
Write-Host "Total number of enabled computers running Windows 10 or Windows 11 successfully moved: $movedCount" -ForegroundColor Blue
