# Import required module
Import-Module ActiveDirectory

# Function to create the target OU if it doesn't exist
function New-NotSyncedOU {
    param (
        [string]$OUName = "NotSyncedtoEID",
        [string]$Description = "Users not synced to Entra ID"
    )
    
    try {
        # Get domain DN
        $domainDN = (Get-ADDomain).DistinguishedName
        $ouDN = "OU=$OUName,$domainDN"
        
        # Check if OU exists
        if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$ouDN'" -ErrorAction SilentlyContinue)) {
            Write-Host "Creating OU: $OUName" -ForegroundColor Yellow
            New-ADOrganizationalUnit -Name $OUName -Path $domainDN -Description $Description
            Write-Host "OU created successfully" -ForegroundColor Green
            return $ouDN
        }
        else {
            Write-Host "OU $OUName already exists" -ForegroundColor Cyan
            return $ouDN
        }
    }
    catch {
        Write-Error "Failed to create OU: $_"
        exit
    }
}

# Function to convert ImmutableID to ObjectGUID
function Convert-ImmutableIDToGuid {
    param (
        [string]$ImmutableID
    )
    
    try {
        $guidBytes = [System.Convert]::FromBase64String($ImmutableID)
        $guid = [System.Guid]::new($guidBytes)
        return $guid
    }
    catch {
        Write-Warning "Failed to convert ImmutableID: $ImmutableID"
        return $null
    }
}

# Initialize result tracking
$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$reportPath = "C:\code\exports\NTI\MoveReport_$timestamp.csv"

# Verify CSV exists
$csvPath = "C:\code\exports\NTI\UnlicensedUsersToMoveToNonSyncedOU.csv"
if (-not (Test-Path $csvPath)) {
    Write-Error "CSV file not found at: $csvPath"
    exit
}

# Create target OU
$targetOUDN = New-NotSyncedOU

# Read CSV and prepare move list
Write-Host "`nReading CSV file..." -ForegroundColor Cyan
$usersToMove = Import-Csv $csvPath

Write-Host "`nPreparing move list..." -ForegroundColor Cyan
$moveList = foreach ($user in $usersToMove) {
    $guid = Convert-ImmutableIDToGuid -ImmutableID $user.ImmutableId
    if ($guid) {
        try {
            $adUser = Get-ADUser -Identity $guid -Properties DistinguishedName
            if ($adUser) {
                [PSCustomObject]@{
                    ImmutableId = $user.ImmutableId
                    UserPrincipalName = $user.UserPrincipalName
                    ObjectGUID = $guid
                    CurrentOU = ($adUser.DistinguishedName -split ',', 2)[1]
                    ADUser = $adUser
                }
            }
        }
        catch {
            Write-Warning "Failed to find AD user for: $($user.UserPrincipalName)"
        }
    }
}

# Show preview and get confirmation
Write-Host "`nUsers to be moved:" -ForegroundColor Yellow
$moveList | Format-Table -AutoSize @{
    Label = "UPN"
    Expression = { $_.UserPrincipalName }
}, @{
    Label = "Current OU"
    Expression = { $_.CurrentOU }
}

$totalToMove = $moveList.Count
Write-Host "`nTotal users to move: $totalToMove" -ForegroundColor Cyan

$confirm = Read-Host "`nDo you want to proceed with moving these users? (Y/N)"
if ($confirm -ne 'Y') {
    Write-Host "Operation cancelled by user" -ForegroundColor Yellow
    exit
}

# Perform moves
Write-Host "`nMoving users..." -ForegroundColor Cyan
$successful = 0
$failed = 0

foreach ($user in $moveList) {
    try {
        $oldOU = $user.CurrentOU
        Move-ADObject -Identity $user.ADUser.DistinguishedName -TargetPath $targetOUDN -ErrorAction Stop
        $successful++
        
        # Verify new location
        $newLocation = (Get-ADUser -Identity $user.ObjectGUID).DistinguishedName
        
        $results.Add([PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            ImmutableId = $user.ImmutableId
            Status = "Success"
            OldOU = $oldOU
            NewOU = ($newLocation -split ',', 2)[1]
            ErrorMessage = ""
        })
        
        Write-Host "Moved: $($user.UserPrincipalName)" -ForegroundColor Green
    }
    catch {
        $failed++
        $results.Add([PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            ImmutableId = $user.ImmutableId
            Status = "Failed"
            OldOU = $oldOU
            NewOU = "N/A"
            ErrorMessage = $_.Exception.Message
        })
        
        Write-Warning "Failed to move $($user.UserPrincipalName): $_"
    }
}

# Generate summary report
$summaryReport = @"
==========================================
User Move Operation Summary
Generated on: $(Get-Date)
==========================================

Operation Statistics:
-------------------
Total Users Processed: $totalToMove
Successfully Moved: $successful
Failed Moves: $failed

Target OU: $targetOUDN

Detailed Results:
---------------
$(($results | Format-Table -AutoSize | Out-String).Trim())
"@

# Display and export results
Write-Host "`n$summaryReport"

# Export detailed results to CSV
$results | Export-Csv -Path $reportPath -NoTypeInformation

Write-Host "`nOperation completed!" -ForegroundColor Green
Write-Host "Detailed report saved to: $reportPath" -ForegroundColor Cyan
Write-Host "Successfully moved: $successful users" -ForegroundColor $(if ($successful -eq $totalToMove) { 'Green' } else { 'Yellow' })
if ($failed -gt 0) {
    Write-Host "Failed to move: $failed users" -ForegroundColor Red
}