# Define target OU
$targetOU = Get-ADOrganizationalUnit -Filter "Name -eq 'NotSyncedtoEID'" | Select-Object -ExpandProperty DistinguishedName

if (-not $targetOU) {
    Write-Host "Target OU 'NotSyncedtoEID' not found!" -ForegroundColor Red
    exit
}

# Array of distinguished names from the XML with permission issues
$problemUsers = @(
    "CN=James Sandy (Admin),OU=Users,OU=Communications,OU=1-Departmental-Units,DC=ri,DC=nti,DC=local",
    "CN=James Sandy,OU=Users,OU=Communications,OU=1-Departmental-Units,DC=ri,DC=nti,DC=local"
)

foreach ($userDN in $problemUsers) {
    Write-Host "`nProcessing user: $userDN" -ForegroundColor Yellow
    
    try {
        # Get user details before moving
        $user = Get-ADObject -Identity $userDN -Properties ObjectClass, DisplayName
        
        Write-Host "Found user:" -ForegroundColor Green
        Write-Host "Display Name: $($user.DisplayName)" -ForegroundColor Green
        Write-Host "Object Class: $($user.ObjectClass)" -ForegroundColor Green
        Write-Host "Current Location: $($user.DistinguishedName)" -ForegroundColor Green

        # Check if already in NotSyncedtoEID
        if ($user.DistinguishedName -notlike "*$targetOU*") {
            # Move the user
            Move-ADObject -Identity $userDN -TargetPath $targetOU 
            Write-Host "Successfully moved user to NotSyncedtoEID" -ForegroundColor Green
            
            # Verify the move
            $movedUser = Get-ADObject -Identity $user.ObjectGUID -Properties DistinguishedName
            Write-Host "New location: $($movedUser.DistinguishedName)" -ForegroundColor Green
        }
        else {
            Write-Host "User already in NotSyncedtoEID OU" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Error processing user $userDN : $_" -ForegroundColor Red
    }
}

Write-Host "`nAll users with permission issues have been processed." -ForegroundColor Cyan