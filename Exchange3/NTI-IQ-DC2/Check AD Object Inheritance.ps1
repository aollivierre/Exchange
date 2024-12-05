# Import the Active Directory module
Import-Module ActiveDirectory

# User's Distinguished Name from the screenshot - corrected with proper DC components
$userDN = "CN=Nathaniel Alexander,OU=Users,OU=Information Technology and Systems,OU=2-Departmental,DC=iq,DC=nti,DC=local"

# Get the user object and report if not found
try {
    $user = Get-ADUser -Identity $userDN -ErrorAction Stop
    Write-Host "User found: $($user.Name)" -ForegroundColor Green
} catch {
    Write-Host "Error finding user. Trying to search by name..." -ForegroundColor Yellow
    try {
        $user = Get-ADUser -Filter "Name -eq 'Nathaniel Alexander'" -ErrorAction Stop
        Write-Host "User found by name search. Full DN: $($user.DistinguishedName)" -ForegroundColor Green
        $userDN = $user.DistinguishedName
    } catch {
        Write-Host "Could not find user by name either. Please verify the username." -ForegroundColor Red
        exit
    }
}

# Get the user's ACL
try {
    $acl = Get-Acl -Path "AD:\$userDN"

    # Check if inheritance is enabled
    Write-Host "`nChecking inheritance status for: $($user.Name)" -ForegroundColor Cyan
    Write-Host "Inheritance Enabled: $(-not $acl.AreAccessRulesProtected)" -ForegroundColor Yellow

    # Display inheritance information
    Write-Host "`nDetailed Access Rules:" -ForegroundColor Cyan
    $acl.Access | Select-Object @{
        Name='IdentityReference'
        Expression={$_.IdentityReference}
    }, @{
        Name='AccessControlType'
        Expression={$_.AccessControlType}
    }, @{
        Name='IsInherited'
        Expression={$_.IsInherited}
    }, @{
        Name='InheritanceFlags'
        Expression={$_.InheritanceFlags}
    }, @{
        Name='PropagationFlags'
        Expression={$_.PropagationFlags}
    } | Format-Table -AutoSize

    # Check if AADSync account has proper permissions
    $aadSyncAccount = Get-ADUser -Filter {Name -like "MSOL_*"} | Select-Object -First 1
    if ($aadSyncAccount) {
        Write-Host "`nChecking AADSync account permissions:" -ForegroundColor Cyan
        $acl.Access | Where-Object {$_.IdentityReference -like "*$($aadSyncAccount.SamAccountName)*"} | 
        Format-Table IdentityReference, AccessControlType, IsInherited -AutoSize
    } else {
        Write-Host "`nNo AADSync account found (looking for MSOL_* account)" -ForegroundColor Yellow
    }

    # Check Enterprise Key Admins membership
    Write-Host "`nChecking Enterprise Key Admins membership:" -ForegroundColor Cyan
    try {
        $keyAdmins = Get-ADGroupMember "Enterprise Key Admins" -ErrorAction Stop
        $aadSyncInKeyAdmins = $keyAdmins | Where-Object {$_.SamAccountName -like "MSOL_*"}
        if ($aadSyncInKeyAdmins) {
            Write-Host "AADSync account is a member of Enterprise Key Admins" -ForegroundColor Green
        } else {
            Write-Host "AADSync account is NOT a member of Enterprise Key Admins" -ForegroundColor Red
            Write-Host "Consider adding the AADSync account to Enterprise Key Admins group"
        }
    } catch {
        Write-Host "Enterprise Key Admins group not found - this requires at least one Windows Server 2016 DC" -ForegroundColor Red
        
        # Check Domain Functional Level
        $domainLevel = (Get-ADDomain).DomainMode
        Write-Host "`nDomain Functional Level: $domainLevel" -ForegroundColor Yellow
        
        if ($domainLevel -lt "Windows2016Domain") {
            Write-Host "Domain needs to be at least Windows Server 2016 functional level" -ForegroundColor Red
        }
    }
} catch {
    Write-Host "Error accessing ACL: $_" -ForegroundColor Red
}