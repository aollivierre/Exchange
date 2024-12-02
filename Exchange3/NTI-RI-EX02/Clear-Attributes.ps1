# Full Distinguished Name from your previous output
$userDN = "CN=James Sandy,OU=Users,OU=Communications,OU=1-Departmental-Units,DC=ri,DC=nti,DC=local"

# Create a DirectoryEntry for low-level attribute manipulation
try {
    $userDE = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$userDN")
    Write-Host "Connected to user object" -ForegroundColor Green

    # Clear msDS-KeyCredentialLink
    if ($userDE."msDS-KeyCredentialLink") {
        Write-Host "Clearing msDS-KeyCredentialLink" -ForegroundColor Yellow
        $userDE."msDS-KeyCredentialLink".Clear()
    }
    else {
        Write-Host "msDS-KeyCredentialLink is already empty" -ForegroundColor Cyan
    }
    
    # Clear msDS-ExternalDirectoryObjectId
    if ($userDE."msDS-ExternalDirectoryObjectId") {
        Write-Host "Clearing msDS-ExternalDirectoryObjectId" -ForegroundColor Yellow
        $userDE."msDS-ExternalDirectoryObjectId".Clear()
    }
    else {
        Write-Host "msDS-ExternalDirectoryObjectId is already empty" -ForegroundColor Cyan
    }
    
    $userDE.CommitChanges()
    Write-Host "Successfully processed attributes" -ForegroundColor Green
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
}
finally {
    if ($userDE) {
        $userDE.Dispose()
    }
}