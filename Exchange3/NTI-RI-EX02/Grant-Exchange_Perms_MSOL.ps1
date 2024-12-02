# Get the MSOL account name
$msolAccount = Get-ADUser -Filter {SamAccountName -like "MSOL_*"} | Select-Object -First 1 -ExpandProperty SamAccountName

# Get the domain
$domain = (Get-ADDomain).DistinguishedName

# Get domain NetBIOS name
$netbios = (Get-ADDomain).NetBIOSName

# Full account name
$fullAccount = "$netbios\$msolAccount"

# Commands to grant specific permissions
$commands = @(
    # Grant general read/write permissions
    "dsacls.exe `"$domain`" /G `"$fullAccount`:RPWP;;User`" /I:S",
    
    # Grant specific permissions for msDS-KeyCredentialLink
    "dsacls.exe `"$domain`" /G `"$fullAccount`:RPWP;msDS-KeyCredentialLink;User`" /I:S",
    
    # Grant specific permissions for msDS-ExternalDirectoryObjectId
    "dsacls.exe `"$domain`" /G `"$fullAccount`:RPWP;msDS-ExternalDirectoryObjectId;User`" /I:S"
)

# Execute each command
foreach ($cmd in $commands) {
    Write-Host "Executing: $cmd" -ForegroundColor Yellow
    Invoke-Expression $cmd
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Command executed successfully" -ForegroundColor Green
    }
    else {
        Write-Host "Command failed with exit code $LASTEXITCODE" -ForegroundColor Red
    }
}

Write-Host "`nPermissions have been updated. Please wait a few minutes for AD replication and try the sync again." -ForegroundColor Cyan