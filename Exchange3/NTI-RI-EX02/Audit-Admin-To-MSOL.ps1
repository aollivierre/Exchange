# Script to audit MSOL_ accounts and their privileged group memberships

# First prompt for confirmation
$confirmation = Read-Host "This script will AUDIT MSOL_ accounts and their group memberships. Continue? (y/n)"
if ($confirmation -ne 'y') {
    Write-Host "Operation cancelled by user"
    exit
}

# Get all MSOL_ accounts
$msolAccounts = Get-ADUser -Filter {SamAccountName -like "MSOL_*"} -Properties MemberOf

# Initialize report array
$report = @()

# Check each account
foreach ($account in $msolAccounts) {
    # Get current group memberships
    $groups = $account.MemberOf | ForEach-Object {
        (Get-ADGroup $_).Name
    }
    
    # Check privileged group memberships
    $inDomainAdmins = $groups -contains "Domain Admins"
    $inEnterpriseAdmins = $groups -contains "Enterprise Admins"
    $inSchemaAdmins = $groups -contains "Schema Admins"
    
    # Add to report
    $report += [PSCustomObject]@{
        AccountName = $account.SamAccountName
        InDomainAdmins = $inDomainAdmins
        InEnterpriseAdmins = $inEnterpriseAdmins
        InSchemaAdmins = $inSchemaAdmins
        CurrentGroups = ($groups -join ", ")
    }
}

# Display report
Write-Host "`nMSOL_ Account Privilege Audit Report"
Write-Host "================================="
$report | Format-Table -AutoSize

# Export to CSV
$datetime = Get-Date -Format "yyyyMMdd-HHmmss"
$report | Export-Csv -Path "MSOL_Account_Audit_$datetime.csv" -NoTypeInformation
Write-Host "`nReport exported to MSOL_Account_Audit_$datetime.csv"