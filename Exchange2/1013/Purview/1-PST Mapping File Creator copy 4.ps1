#Requires -Modules PSWriteHTML

function Get-MailboxFromPSTName {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$PSTName,
        
        [Parameter(Mandatory)]
        [string]$DomainName
    )
    
    # Extract username from PST name pattern (username_archive_timestamp.pst)
    if ($PSTName -match '^([^_]+)_archive_\d{8}_\d{6}\.pst$') {
        $username = $Matches[1]
        # Handle special case for ri-support
        if ($username -eq 'risupport') {
            $username = 'ri-support'
        }
        return "$username@$DomainName"
    }
    
    Write-Warning "Could not extract mailbox from PST name: $PSTName"
    return $null
}


function Show-PreflightChecks {
    [CmdletBinding()]
    param()
    
    $continue = $false
    while (-not $continue) {
        Write-Host "`nPST Import Pre-flight Checks" -ForegroundColor Yellow
        Write-Host "------------------------" -ForegroundColor Yellow
        $response = Read-Host "Have you scanned all PST files using ScanPST.exe? (Y/N)"
        
        if ($response -eq 'Y') {
            $continue = $true
        }
        else {
            Write-Host "`nWARNING: Please scan your PST files before proceeding!" -ForegroundColor Red
            Write-Host "You can use: C:\code\Purview\4-automate-scanPST-interaction copy 11.py" -ForegroundColor Cyan
            Write-Host "Run it using: py 'C:\code\Purview\4-automate-scanPST-interaction copy 11.py'" -ForegroundColor Cyan
            
            $exit = Read-Host "Do you want to exit and scan the files first? (Y/N)"
            if ($exit -eq 'Y') {
                throw "Please scan PST files and retry the operation when ready."
            }
        }
    }
}

function Show-PostMappingInstructions {
    [CmdletBinding()]
    param()
    
    Write-Host "`nNext Steps and Reminders:" -ForegroundColor Yellow
    Write-Host "------------------------" -ForegroundColor Yellow
    
    $instructions = @(
        "1. Install Azure Storage Explorer Desktop app"
        "   - Connect using 'Blob Container' or 'Directory' option"
        "   - Select 'Shared Access Signature URL (SAS)'"
        ""
        "2. FilePath Column in CSV:"
        "   - Must match the Azure Storage folder name"
        "   - Leave empty if PSTs are in root folder"
        ""
        "3. Exchange Online Permissions:"
        "   - Add Import/Export role to Organization Management"
        ""
        "4. Purview Import Settings:"
        "   - Select 'Import everything (recommended)'"
        "   - Don't apply filters during import"
    )
    
    $instructions | ForEach-Object {
        Write-Host $_ -ForegroundColor Cyan
    }
    
    Write-Host "`nDetailed documentation has been created: README.md" -ForegroundColor Green
}





function New-PSTMappingFile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$PSTFolder,
        
        [Parameter(Mandatory)]
        [string]$DomainName,
        
        [Parameter()]
        [bool]$IsArchive = $false,
        
        [Parameter()]
        [string]$OutputPath = ".\PST_Mapping",
        
        [Parameter()]
        [string]$TargetRootFolder = "/"
    )



    # Run pre-flight checks
    Show-PreflightChecks

    # Ensure output directory exists
    $null = New-Item -ItemType Directory -Force -Path $OutputPath
    
    # Generate timestamp for the file names
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = Join-Path $OutputPath "PST_Mapping_$timestamp.csv"
    $htmlPath = Join-Path $OutputPath "PST_Mapping_$timestamp.html"
    
    # Initialize collection for mappings
    $mappings = [System.Collections.Generic.List[object]]::new()

    
    # Get all PST files
    $pstFiles = Get-ChildItem -Path $PSTFolder -Filter "*.pst"
    
    foreach ($pst in $pstFiles) {
        $targetMailbox = Get-MailboxFromPSTName -PSTName $pst.Name -DomainName $DomainName
        
        if (-not $targetMailbox) {
            Write-Warning "Skipping $($pst.Name) - Could not determine target mailbox"
            continue
        }

        $mapping = [PSCustomObject]@{
            'Workload'            = 'Exchange'
            'FilePath'            = ''
            'Name'                = $pst.Name
            'Mailbox'             = $targetMailbox
            'IsArchive'           = $IsArchive.ToString().ToUpper()
            'TargetRootFolder'    = $TargetRootFolder
            'ContentCodePage'     = ''
            'SPFileContainer'     = ''
            'SPManifestContainer' = ''
            'SPSiteUrl'           = ''
            'FullPath'            = $pst.FullName
            'SizeMB'              = [math]::Round($pst.Length / 1MB, 2)
            'LastWriteTime'       = $pst.LastWriteTime
        }
        
        $mappings.Add($mapping)
    }

    # Export to CSV (Microsoft format)
    $mappings | Select-Object -ExcludeProperty FullPath, SizeMB, LastWriteTime | 
    Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    # Create metadata for report
    $metadata = @{
        GeneratedBy   = $env:USERNAME
        GeneratedOn   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        TotalPSTFiles = $mappings.Count
        TotalSizeMB   = [math]::Round(($mappings | Measure-Object -Property SizeMB -Sum).Sum, 2)
        DomainName    = $DomainName
        IsArchive     = $IsArchive
        MappingFile   = $csvPath
    }

    # Generate HTML Report
    New-HTML -TitleText "PST Mapping Report" -FilePath $htmlPath -ShowHTML {
        New-HTMLSection -HeaderText "Generation Summary" {
            New-HTMLPanel {
                New-HTMLText -Text @"
                <h3>Report Details</h3>
                <ul>
                    <li>Generated By: $($metadata.GeneratedBy)</li>
                    <li>Generated On: $($metadata.GeneratedOn)</li>
                    <li>Total PST Files: $($metadata.TotalPSTFiles)</li>
                    <li>Total Size (MB): $($metadata.TotalSizeMB)</li>
                    <li>Domain Name: $($metadata.DomainName)</li>
                    <li>Import to Archive: $($metadata.IsArchive)</li>
                </ul>
"@
            }
        }
        
        New-HTMLSection -HeaderText "PST Mapping Details" {
            New-HTMLTable -DataTable $mappings -ScrollX -Buttons @('copyHtml5', 'excelHtml5', 'csvHtml5') -SearchBuilder {
                New-TableCondition -Name 'SizeMB' -ComparisonType number -Operator gt -Value 1000 -BackgroundColor LightSalmon -Color Black
                New-TableCondition -Name 'SizeMB' -ComparisonType number -Operator lt -Value 100 -BackgroundColor LightGreen -Color Black
            }
        }
    }

    # Console Output
    Write-Host "`nPST Mapping Report Summary:" -ForegroundColor Cyan
    Write-Host "Total PST Files: $($metadata.TotalPSTFiles)" -ForegroundColor Green
    Write-Host "Total Size (MB): $($metadata.TotalSizeMB)" -ForegroundColor Green
    Write-Host "Domain Name: $($metadata.DomainName)" -ForegroundColor Green
    Write-Host "Import to Archive: $($metadata.IsArchive)" -ForegroundColor Green
    Write-Host "`nOutput Files:" -ForegroundColor Cyan
    Write-Host "CSV Mapping: $csvPath" -ForegroundColor Green
    Write-Host "HTML Report: $htmlPath" -ForegroundColor Green

    return $metadata


    # Show post-mapping instructions
    Show-PostMappingInstructions
    

}

# Example usage
$mappingParams = @{
    PSTFolder        = "C:\ExchangeArchives\repaired3"
    DomainName       = "tunngavik.com"
    IsArchive        = $true
    TargetRootFolder = "/"
    OutputPath       = ".\PST_Mapping"
}

New-PSTMappingFile @mappingParams