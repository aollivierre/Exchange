function New-PSTMappingFile {
    param (
        [Parameter(Mandatory)]
        [string]$PSTPath,
        
        [Parameter(Mandatory)]
        [string]$TargetMailbox,
        
        [Parameter()]
        [bool]$IsArchive = $false,
        
        [Parameter()]
        [string]$OutputPath = ".\PST_Mapping",
        
        [Parameter()]
        [string]$TargetRootFolder = "/"
    )
    
    # Ensure output directory exists
    $null = New-Item -ItemType Directory -Force -Path $OutputPath
    
    # Get just the PST filename without path
    $pstFileName = Split-Path $PSTPath -Leaf
    
    # Generate timestamp for the file name
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = Join-Path $OutputPath "PST_Mapping_$timestamp.csv"
    
    # Create mapping object with EXACT required fields per Microsoft documentation
    $mapping = [PSCustomObject]@{
        'Workload' = 'Exchange'
        'FilePath' = ''  # Leave blank as we're uploading to root
        'Name' = $pstFileName
        'Mailbox' = $TargetMailbox
        'IsArchive' = $IsArchive.ToString().ToUpper()  # Must be TRUE or FALSE in uppercase
        'TargetRootFolder' = $TargetRootFolder
        'ContentCodePage' = ''  # Optional, leave blank unless needed for DBCS
        'SPFileContainer' = ''  # Not used for Exchange imports
        'SPManifestContainer' = ''  # Not used for Exchange imports
        'SPSiteUrl' = ''  # Not used for Exchange imports
    }
    
    # Export to CSV with specific encoding and format
    $mapping | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "Mapping file created successfully at: $csvPath" -ForegroundColor Green
    Write-Host "Mapping details:" -ForegroundColor Cyan
    Write-Host "PST File: $pstFileName" -ForegroundColor Cyan
    Write-Host "Target Mailbox: $TargetMailbox" -ForegroundColor Cyan
    Write-Host "Import to Archive: $IsArchive" -ForegroundColor Cyan
    Write-Host "Target Folder: $TargetRootFolder" -ForegroundColor Cyan
    
    return @{
        MappingFile = $csvPath
        MappingDetails = $mapping
    }
}

# Example usage
$mappingParams = @{
    PSTPath = "C:\PST\Administrator_archive_20241205_102100.pst"
    TargetMailbox = "PurviewPSTNetworkUploadtest001@tunngavik.com"
    IsArchive = $true
    TargetRootFolder = "/"  # Root level import
    OutputPath = ".\PST_Mapping"
}

New-PSTMappingFile @mappingParams