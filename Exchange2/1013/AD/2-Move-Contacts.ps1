# Import required module
Import-Module ActiveDirectory

# Function to create the target OU if it doesn't exist
function New-NotSyncedOU {
    param (
        [string]$OUName = "NotSyncedContactsEID",
        [string]$Description = "Contacts not synced to Entra ID"
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
$reportPath = "C:\code\exports\NTI\ContactMoveReport_$timestamp.csv"

# Verify CSV exists
$csvPath = "C:\Code\Exports\NTI\Directory Sync Errors - 2024-11-14_09-55-01-Import-Contacts-To-Move.csv"
if (-not (Test-Path $csvPath)) {
    Write-Error "CSV file not found at: $csvPath"
    exit
}

# Create target OU
$targetOUDN = New-NotSyncedOU

# Read CSV and prepare move list
Write-Host "`nReading CSV file..." -ForegroundColor Cyan
$contactsToMove = Import-Csv $csvPath

Write-Host "`nPreparing move list..." -ForegroundColor Cyan
$moveList = foreach ($contact in $contactsToMove) {
    $guid = Convert-ImmutableIDToGuid -ImmutableID $contact.ImmutableId
    if ($guid) {
        try {
            # Use Get-ADObject to find contacts
            $adContact = Get-ADObject -Filter "ObjectGUID -eq '$guid'" -Properties DistinguishedName, mail, proxyAddresses
            if ($adContact) {
                [PSCustomObject]@{
                    ImmutableId = $contact.ImmutableId
                    Mail = $adContact.mail
                    ObjectGUID = $guid
                    CurrentOU = ($adContact.DistinguishedName -split ',', 2)[1]
                    ADContact = $adContact
                }
            }
        }
        catch {
            Write-Warning "Failed to find AD contact with GUID: $guid"
        }
    }
}

# Show preview and get confirmation
Write-Host "`nContacts to be moved:" -ForegroundColor Yellow
$moveList | Format-Table -AutoSize @{
    Label = "Email"
    Expression = { $_.Mail }
}, @{
    Label = "Current OU"
    Expression = { $_.CurrentOU }
}

$totalToMove = $moveList.Count
Write-Host "`nTotal contacts to move: $totalToMove" -ForegroundColor Cyan

$confirm = Read-Host "`nDo you want to proceed with moving these contacts? (Y/N)"
if ($confirm -ne 'Y') {
    Write-Host "Operation cancelled by user" -ForegroundColor Yellow
    exit
}

# Perform moves
Write-Host "`nMoving contacts..." -ForegroundColor Cyan
$successful = 0
$failed = 0

foreach ($contact in $moveList) {
    try {
        $oldOU = $contact.CurrentOU
        Move-ADObject -Identity $contact.ADContact.DistinguishedName -TargetPath $targetOUDN -ErrorAction Stop
        $successful++
        
        # Verify new location
        $newLocation = (Get-ADObject -Filter "ObjectGUID -eq '$($contact.ObjectGUID)'").DistinguishedName
        
        $results.Add([PSCustomObject]@{
            Email = $contact.Mail
            ImmutableId = $contact.ImmutableId
            Status = "Success"
            OldOU = $oldOU
            NewOU = ($newLocation -split ',', 2)[1]
            ErrorMessage = ""
        })
        
        Write-Host "Moved: $($contact.Mail)" -ForegroundColor Green
    }
    catch {
        $failed++
        $results.Add([PSCustomObject]@{
            Email = $contact.Mail
            ImmutableId = $contact.ImmutableId
            Status = "Failed"
            OldOU = $oldOU
            NewOU = "N/A"
            ErrorMessage = $_.Exception.Message
        })
        
        Write-Warning "Failed to move contact $($contact.Mail): $_"
    }
}

# Generate summary report
$summaryReport = @"
==========================================
Contact Move Operation Summary
Generated on: $(Get-Date)
==========================================

Operation Statistics:
-------------------
Total Contacts Processed: $totalToMove
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
Write-Host "Successfully moved: $successful contacts" -ForegroundColor $(if ($successful -eq $totalToMove) { 'Green' } else { 'Yellow' })
if ($failed -gt 0) {
    Write-Host "Failed to move: $failed contacts" -ForegroundColor Red
}