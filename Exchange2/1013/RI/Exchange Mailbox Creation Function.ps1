#Requires -Modules PSWriteHTML

function Get-DomainInfo {
    [CmdletBinding()]
    param()
    
    try {
        $domainInfo = Get-ADDomain
        
        # Try to find any organizational unit
        $availableOU = Get-ADOrganizationalUnit -Filter * -SearchBase $domainInfo.DistinguishedName | Select-Object -First 1
        
        if (-not $availableOU) {
            # If no OUs found, use the default Users container
            $usersContainer = Get-ADContainer -Filter "name -eq 'Users'" -SearchBase $domainInfo.DistinguishedName
            $containerPath = $usersContainer.DistinguishedName
        } else {
            $containerPath = $availableOU.DistinguishedName
        }
        
        return @{
            DomainDN = $domainInfo.DistinguishedName
            DomainName = $domainInfo.DNSRoot
            ContainerPath = $containerPath
            Success = $true
            Error = $null
        }
    }
    catch {
        return @{
            DomainDN = $null
            DomainName = $null
            ContainerPath = $null
            Success = $false
            Error = $_.Exception.Message
        }
    }
}

function New-CustomMailbox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserAlias,
        
        [Parameter()]
        [string]$ReportPath = "C:\MailboxReports",
        
        [Parameter()]
        [string]$Password = "DefaultP@ssw0rd123!",

        [Parameter()]
        [string]$ArchiveDatabase = "" # Will be auto-detected if not specified
    )

    # Get domain information
    $domainInfo = Get-DomainInfo
    if (-not $domainInfo.Success) {
        Write-Error "Failed to get domain information: $($domainInfo.Error)"
        return
    }

    # Create secure password
    $SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force

    # Get available archive database if not specified
    if (-not $ArchiveDatabase) {
        try {
            $ArchiveDatabase = (Get-MailboxDatabase | Where-Object { $_.Recovery -eq $false } | Select-Object -First 1).Name
            if (-not $ArchiveDatabase) {
                Write-Error "No suitable mailbox database found for archive"
                return
            }
        }
        catch {
            Write-Error "Failed to get mailbox database: $($_.Exception.Message)"
            return
        }
    }

    # Parameters for New-Mailbox
    $MailboxParams = @{
        Name                 = $UserAlias
        Alias               = $UserAlias
        OrganizationalUnit  = $domainInfo.ContainerPath
        Password            = $SecurePassword
        UserPrincipalName   = "$UserAlias@$($domainInfo.DomainName)"
        SamAccountName      = $UserAlias
        FirstName           = $UserAlias
        LastName           = $UserAlias
        DisplayName         = $UserAlias
        ResetPasswordOnNextLogon = $true
        Archive             = $true
        ArchiveDatabase    = $ArchiveDatabase
    }

    try {
        # Create mailbox with archive
        $NewMailbox = New-Mailbox @MailboxParams
        
        # Wait briefly for the mailbox to be created
        Start-Sleep -Seconds 2
        
        # Get detailed mailbox information
        $MailboxInfo = Get-Mailbox -Identity $UserAlias -ErrorAction SilentlyContinue
        $ArchiveInfo = Get-Mailbox -Identity $UserAlias -Archive -ErrorAction SilentlyContinue
        
        # Check if mailbox is on-premises
        $DatabaseInfo = Get-MailboxDatabase -Identity $MailboxInfo.Database -ErrorAction SilentlyContinue
        $IsOnPrem = $null -ne $DatabaseInfo
        
        # Prepare result object
        $Result = [PSCustomObject]@{
            UserAlias = $UserAlias
            Status = "Success"
            EmailAddress = $MailboxInfo.PrimarySmtpAddress
            Created = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Error = $null
            Domain = $domainInfo.DomainName
            Container = $domainInfo.ContainerPath
            ArchiveEnabled = $null -ne $ArchiveInfo
            ArchiveDatabase = $ArchiveDatabase
            MailboxDatabase = $MailboxInfo.Database
            IsOnPremises = $IsOnPrem
        }

        Write-Host "Mailbox created successfully:" -ForegroundColor Green
        Write-Host "User: $UserAlias" -ForegroundColor Green
        Write-Host "Email: $($MailboxInfo.PrimarySmtpAddress)" -ForegroundColor Green
        Write-Host "Container: $($domainInfo.ContainerPath)" -ForegroundColor Green
        Write-Host "Archive Enabled: $($Result.ArchiveEnabled)" -ForegroundColor Green
        Write-Host "Archive Database: $ArchiveDatabase" -ForegroundColor Green
        Write-Host "On-Premises: $IsOnPrem" -ForegroundColor Green
    }
    catch {
        $Result = [PSCustomObject]@{
            UserAlias = $UserAlias
            Status = "Failed"
            EmailAddress = $null
            Created = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Error = $_.Exception.Message
            Domain = $domainInfo.DomainName
            Container = $domainInfo.ContainerPath
            ArchiveEnabled = $false
            ArchiveDatabase = $null
            MailboxDatabase = $null
            IsOnPremises = $null
        }

        Write-Host "Failed to create mailbox:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }

    # Rest of the reporting code remains the same
    # Ensure report directory exists
    if (-not (Test-Path -Path $ReportPath)) {
        New-Item -ItemType Directory -Path $ReportPath | Out-Null
    }

    # Generate timestamp for reports
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    
    # Export to CSV
    $csvPath = Join-Path $ReportPath "MailboxCreation_$timestamp.csv"
    $Result | Export-Csv -Path $csvPath -NoTypeInformation

    # Create HTML report
    $htmlPath = Join-Path $ReportPath "MailboxCreation_$timestamp.html"
    
    $metadata = @{
        GeneratedBy = $env:USERNAME
        GeneratedOn = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        TotalUsers = 1
        SuccessCount = if ($Result.Status -eq "Success") { 1 } else { 0 }
        FailureCount = if ($Result.Status -eq "Failed") { 1 } else { 0 }
        Domain = $domainInfo.DomainName
        Container = $domainInfo.ContainerPath
        ArchiveDatabase = $Result.ArchiveDatabase
        IsOnPremises = $Result.IsOnPremises
    }

    New-HTML -Title "Mailbox Creation Report" -FilePath $htmlPath -ShowHTML {
        New-HTMLSection -HeaderText "Generation Summary" {
            New-HTMLPanel {
                New-HTMLText -Text @"
                <h3>Report Details</h3>
                <ul>
                    <li>Generated By: $($metadata.GeneratedBy)</li>
                    <li>Generated On: $($metadata.GeneratedOn)</li>
                    <li>Total Users Processed: $($metadata.TotalUsers)</li>
                    <li>Successful Creations: $($metadata.SuccessCount)</li>
                    <li>Failed Creations: $($metadata.FailureCount)</li>
                    <li>Domain: $($metadata.Domain)</li>
                    <li>Container: $($metadata.Container)</li>
                    <li>Archive Database: $($metadata.ArchiveDatabase)</li>
                    <li>On-Premises: $($metadata.IsOnPremises)</li>
                </ul>
"@
            }
        }
        
        New-HTMLSection -HeaderText "Mailbox Creation Results" {
            New-HTMLTable -DataTable $Result -ScrollX -Buttons @('copyHtml5', 'excelHtml5', 'csvHtml5') {
                New-TableCondition -Name 'Status' -ComparisonType string -Operator eq -Value 'Failed' -BackgroundColor Salmon -Color Black
                New-TableCondition -Name 'Status' -ComparisonType string -Operator eq -Value 'Success' -BackgroundColor LightGreen -Color Black
            }
        }
    }

    Write-Host "`nReports generated:" -ForegroundColor Green
    Write-Host "CSV Report: $csvPath" -ForegroundColor Green
    Write-Host "HTML Report: $htmlPath" -ForegroundColor Green

    return $Result
}

# Example usage for A0Test002RI
$Params = @{
    UserAlias = "A0Test002RI"
}

New-CustomMailbox @Params