# Function to remove duplicate "old" in the Name, CN, and DisplayName attributes
function Remove-DuplicateOld {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Identity
    )

    try {
        # Get the AD user with duplicate "old"
        $user = Get-ADUser -Filter { Name -like "$Identity-old-old" } -Properties Name, DisplayName, CN
        if ($user) {
            $newName = $user.Name -replace '-old-old', '-old'
            $newDisplayName = $user.DisplayName -replace '-old-old', '-old'
            $newCN = $user.CN -replace '-old-old', '-old'

            # Rename the user
            Rename-ADObject -Identity $user.DistinguishedName -NewName $newCN

            # Update the DisplayName and UserPrincipalName
            Set-ADUser -Identity $user.DistinguishedName -DisplayName $newDisplayName

            Write-Host "Successfully renamed AD object to remove duplicate 'old'."
            Write-Host "New Name: $newName" -ForegroundColor Green
            Write-Host "New DisplayName: $newDisplayName" -ForegroundColor Green
            Write-Host "New CN: $newCN" -ForegroundColor Green
        } else {
            Write-Host "No AD User found with duplicate 'old' for $Identity" -ForegroundColor Red
        }
    } catch {
        Write-Host "Error updating details for $Identity $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Call the function to remove duplicate "old"
$identity = "Jordan Heuser"
Remove-DuplicateOld -Identity $identity
