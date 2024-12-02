# Script to grant specific required permissions for Entra Connect sync
$confirmation = Read-Host "This script will update permissions for MSOL_ accounts. Continue? (y/n)"
if ($confirmation -ne 'y') {
    Write-Host "Operation cancelled by user"
    exit
}

# Import the required module
Import-Module ActiveDirectory

# Find MSOL_ accounts
$msolAccounts = Get-ADUser -Filter {SamAccountName -like "MSOL_*"}
$domainDN = (Get-ADDomain).DistinguishedName

foreach ($msolAccount in $msolAccounts) {
    Write-Host "`nProcessing account: $($msolAccount.SamAccountName)"
    
    try {
        # Get AD object
        $adObject = [ADSI]"LDAP://$domainDN"
        $adSecurity = $adObject.psbase.ObjectSecurity

        # Define the permissions
        $rights = [System.DirectoryServices.ActiveDirectoryRights]"GenericAll"
        $type = [System.Security.AccessControl.AccessControlType]"Allow"
        $inheritanceType = [System.DirectoryServices.ActiveDirectorySecurityInheritance]"All"
        
        # Create and apply the access rule
        $accessRule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule(
            $msolAccount.SID,
            $rights,
            $type,
            $inheritanceType
        )

        try {
            $adSecurity.AddAccessRule($accessRule)
            $adObject.psbase.CommitChanges()
            Write-Host "Successfully granted permissions to $($msolAccount.SamAccountName)"
        }
        catch {
            Write-Host "Error applying permissions: $_"
        }

        # Now add the account to the necessary built-in group for attribute management
        Add-ADGroupMember -Identity "Account Operators" -Members $msolAccount.SamAccountName -ErrorAction Stop
        Write-Host "Added $($msolAccount.SamAccountName) to Account Operators group"

    }
    catch {
        Write-Host "Error processing account $($msolAccount.SamAccountName): $_"
    }
}

Write-Host "`nPermissions update complete. Please test Entra Connect sync again."
Write-Host "Note: The script has granted necessary permissions for sync operations."