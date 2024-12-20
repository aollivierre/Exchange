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
    
    # Create a unique log file for this upload
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $azcopyLogFile = Join-Path $LogPath "azcopy_$timestamp.log"
    
    # Prepare AzCopy parameters
    $azcopyParams = @(
        'copy'
        "`"$PSTPath`""  # Quoted to handle spaces in path
        "`"$SasUrl`""   # Quoted to handle special characters
        '--overwrite=true'
        '--log-level=DEBUG'
        "--output-type=text"
        "--log-file=`"$azcopyLogFile`""
    )
    
    Write-Host "Starting PST upload process..." -ForegroundColor Yellow
    Write-Host "Source PST: $PSTPath" -ForegroundColor Cyan
    Write-Host "Log file: $azcopyLogFile" -ForegroundColor Cyan
    
    try {
        # Start upload and capture output
        $output = & $AzCopyPath @azcopyParams 2>&1
        $exitCode = $LASTEXITCODE
        
        # Get log content if file exists
        $logContent = if (Test-Path $azcopyLogFile) {
            Get-Content $azcopyLogFile -Raw
        } else {
            "No log file generated"
        }
        
        # Generate timestamp for reports
        $htmlPath = Join-Path $LogPath "PST_Upload_Report_$timestamp.html"
        
        # Prepare report data
        $reportData = [PSCustomObject]@{
            SourceFile = $PSTPath
            FileSize = "{0:N2} MB" -f ((Get-Item $PSTPath).Length / 1MB)
            UploadStartTime = Get-Date
            Status = if ($exitCode -eq 0) { "Success" } else { "Failed" }
            ExitCode = $exitCode
            ErrorMessage = if ($exitCode -ne 0) { 
                "Exit Code: $exitCode`nOutput: $($output | Out-String)" 
            } else { "" }
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
                        <li>Exit Code: $($reportData.ExitCode)</li>
                    </ul>
"@
                }
            }
            
            if ($exitCode -ne 0) {
                New-HTMLSection -HeaderText "Error Details" {
                    New-HTMLPanel {
                        New-HTMLText -Text @"
                        <h3>Error Information</h3>
                        <pre>$($output | Out-String)</pre>
                        <h3>AzCopy Log</h3>
                        <pre>$logContent</pre>
"@
                    }
                }
            }
            
            New-HTMLSection -HeaderText "Upload Results" {
                New-HTMLTable -DataTable $reportData -ScrollX
            }
        }
        
        # Display immediate results in console
        if ($exitCode -eq 0) {
            Write-Host "PST upload completed successfully!" -ForegroundColor Green
        } else {
            Write-Host "PST upload failed with exit code: $exitCode" -ForegroundColor Red
            Write-Host "Error output:" -ForegroundColor Red
            $output | ForEach-Object { Write-Host $_ -ForegroundColor Red }
        }
        Write-Host "Report generated at: $htmlPath" -ForegroundColor Cyan
        
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