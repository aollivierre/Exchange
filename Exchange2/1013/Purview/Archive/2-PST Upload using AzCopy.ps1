function Start-PSTUpload {
    param (
        [Parameter(Mandatory)]
        [string]$PSTPath,
        
        [Parameter(Mandatory)]
        [string]$SasUrl,
        
        [Parameter()]
        [string]$AzCopyPath = "C:\Users\Administrator\AzCopy\azcopy.exe",
        
        [Parameter()]
        [string]$LogPath = ".\PST_Upload_Logs"
    )
    
    # Ensure paths exist
    if (-not (Test-Path $PSTPath)) {
        throw "PST file not found at path: $PSTPath"
    }
    
    if (-not (Test-Path $AzCopyPath)) {
        throw "AzCopy not found at path: $AzCopyPath"
    }
    
    # Create log directory if it doesn't exist
    $null = New-Item -ItemType Directory -Force -Path $LogPath
    
    # Prepare AzCopy parameters
    $azcopyParams = @(
        'copy'
        $PSTPath
        $SasUrl
        '--overwrite=ifSourceNewer'
        '--check-length=true'
        '--recursive=true'
        '--log-level=INFO'
        "--output-type=json"
    )
    
    Write-Host "Starting PST upload process..." -ForegroundColor Yellow
    Write-Host "Source PST: $PSTPath" -ForegroundColor Cyan
    
    try {
        # Start upload
        $result = & $AzCopyPath @azcopyParams | ConvertFrom-Json
        
        # Generate timestamp for reports
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $htmlPath = Join-Path $LogPath "PST_Upload_Report_$timestamp.html"
        
        # Prepare report data
        $reportData = [PSCustomObject]@{
            SourceFile = $PSTPath
            FileSize = "{0:N2} MB" -f ((Get-Item $PSTPath).Length / 1MB)
            UploadStartTime = Get-Date
            Status = if ($LASTEXITCODE -eq 0) { "Success" } else { "Failed" }
            ErrorMessage = if ($LASTEXITCODE -ne 0) { $result.Error } else { "" }
        }
        
        # Generate HTML report
        New-HTML -Title "PST Upload Report" -FilePath $htmlPath -ShowHTML {
            New-HTMLSection -HeaderText "Upload Summary" {
                New-HTMLPanel {
                    New-HTMLText -Text @"
                    <h3>Upload Details</h3>
                    <ul>
                        <li>File: $($reportData.SourceFile)</li>
                        <li>Size: $($reportData.FileSize)</li>
                        <li>Upload Time: $($reportData.UploadStartTime)</li>
                        <li>Status: $($reportData.Status)</li>
                        <li>Error: $($reportData.ErrorMessage)</li>
                    </ul>
"@
                }
            }
            
            New-HTMLSection -HeaderText "Upload Results" {
                New-HTMLTable -DataTable $reportData -ScrollX
            }
        }
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "PST upload completed successfully!" -ForegroundColor Green
            Write-Host "Report generated at: $htmlPath" -ForegroundColor Green
        } else {
            Write-Host "PST upload failed. Check the report for details." -ForegroundColor Red
            Write-Host "Report generated at: $htmlPath" -ForegroundColor Yellow
        }
        
        return $reportData
        
    } catch {
        Write-Error "Error during upload: $_"
        throw
    }
}

# Example usage
$uploadParams = @{
    PSTPath = "C:\PST\Administrator_archive_20241205_102100.pst"
    SasUrl = "https://ebb534ec5c8c49568557303.blob.core.windows.net/ingestiondata?sv=2015-04-05&sr=c&si=IngestionSasForAzCopy202412191529268229&sig=NJiH3tufGFvWV4aTpXbqeRal%2BSKZNRSI2TQJpYRX8tc%3D&se=2025-01-18T15%3A29%3A27Z"
    LogPath = ".\PST_Upload_Logs"
}

Start-PSTUpload @uploadParams