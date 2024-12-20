#Requires -Modules PSWriteHTML

function Test-PSTFile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$FilePath,
        
        [Parameter()]
        [int]$MaxSizeGB = 20
    )
    
    $fileInfo = Get-Item $FilePath
    $sizeGB = $fileInfo.Length / 1GB
    
    [PSCustomObject]@{
        File             = $fileInfo.Name
        Path             = $fileInfo.FullName
        SizeGB           = [math]::Round($sizeGB, 2)
        IsValidSize      = $sizeGB -le $MaxSizeGB
        Extension        = $fileInfo.Extension.ToLower()
        IsValidExtension = $fileInfo.Extension.ToLower() -eq '.pst'
        LastWriteTime    = $fileInfo.LastWriteTime
    }
}

function Start-PSTUpload {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ParameterSetName = 'SingleFile')]
        [string]$PSTPath,
        
        [Parameter(Mandatory, ParameterSetName = 'Directory')]
        [string]$PSTDirectory,
        
        [Parameter(Mandatory)]
        [string]$SasUrl,
        
        [Parameter()]
        [string]$AzCopyPath = "C:\Users\aollivierre\AzCopy\azcopy.exe",
        
        [Parameter()]
        [string]$LogPath = ".\PST_Upload_Logs",
        
        [Parameter()]
        [int]$MaxSizeGB = 20
    )
    
    # Create log directory if it doesn't exist
    $null = New-Item -ItemType Directory -Force -Path $LogPath
    
    # Get PST files based on input parameter set
    $pstFiles = switch ($PSCmdlet.ParameterSetName) {
        'SingleFile' {
            if (-not (Test-Path $PSTPath)) {
                throw "PST file not found at path: $PSTPath"
            }
            @(Get-Item $PSTPath)
        }
        'Directory' {
            if (-not (Test-Path $PSTDirectory)) {
                throw "Directory not found: $PSTDirectory"
            }
            Get-ChildItem -Path $PSTDirectory -Filter *.pst
        }
    }
    
    # Validate AzCopy existence
    if (-not (Test-Path $AzCopyPath)) {
        throw "AzCopy not found at path: $AzCopyPath"
    }
    
    # Generate timestamp for reports
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $htmlPath = Join-Path $LogPath "PST_Upload_Report_$timestamp.html"
    
    # Validate files before upload
    Write-Host "Validating PST files..." -ForegroundColor Yellow
    $validationResults = foreach ($file in $pstFiles) {
        $validation = Test-PSTFile -FilePath $file.FullName -MaxSizeGB $MaxSizeGB
        
        if (-not $validation.IsValidExtension) {
            Write-Host "Warning: $($file.Name) is not a PST file." -ForegroundColor Yellow
        }
        if (-not $validation.IsValidSize) {
            Write-Host "Warning: $($file.Name) exceeds maximum size of $MaxSizeGB GB (Size: $($validation.SizeGB) GB)" -ForegroundColor Yellow
        }
        
        $validation
    }
    
    # Filter out invalid files
    $validFiles = $validationResults | Where-Object { $_.IsValidSize -and $_.IsValidExtension }
    $invalidFiles = $validationResults | Where-Object { -not ($_.IsValidSize -and $_.IsValidExtension) }
    
    if (-not $validFiles) {
        throw "No valid PST files found to upload!"
    }
    
    # Prepare source path
    $sourcePath = switch ($PSCmdlet.ParameterSetName) {
        'SingleFile' { Split-Path $PSTPath -Parent }
        'Directory' { $PSTDirectory }
    }
    
    Write-Host "`nStarting PST upload process..." -ForegroundColor Yellow
    Write-Host "Source: $sourcePath" -ForegroundColor Cyan
    Write-Host "Total files to upload: $($validFiles.Count)" -ForegroundColor Cyan
    Write-Host "Total size to upload: $([math]::Round(($validFiles | Measure-Object -Property SizeGB -Sum).Sum, 2)) GB" -ForegroundColor Cyan
    
    try {
        # Prepare AzCopy arguments
        $azcopyArgs = @(
            'copy'
            $sourcePath
            $SasUrl
            '--from-to=LocalBlob'
            '--overwrite=true'
            '--recursive=true'
            '--log-level=INFO'
            '--output-type=json'
            '--put-md5'
            '--check-length=true'
        )
        
        # Start upload using Start-Process
        $processParams = @{
            FilePath               = $AzCopyPath
            ArgumentList           = $azcopyArgs
            Wait                   = $true
            PassThru               = $true
            NoNewWindow            = $true
            RedirectStandardError  = "$LogPath\error_$timestamp.log"
            RedirectStandardOutput = "$LogPath\output_$timestamp.log"
        }
        
        $process = Start-Process @processParams
        $exitCode = $process.ExitCode
        
        # Read output and error logs
        $errorLog = if (Test-Path "$LogPath\error_$timestamp.log") { 
            Get-Content "$LogPath\error_$timestamp.log" -Raw 
        }
        else { "" }
        
        $outputLog = if (Test-Path "$LogPath\output_$timestamp.log") { 
            Get-Content "$LogPath\output_$timestamp.log" -Raw 
        }
        else { "" }
        
        # Prepare report data
        $reportData = [PSCustomObject]@{
            TotalFiles      = $validFiles.Count
            TotalSizeGB     = [math]::Round(($validFiles | Measure-Object -Property SizeGB -Sum).Sum, 2)
            UploadStartTime = Get-Date
            Status          = if ($exitCode -eq 0) { "Success" } else { "Failed" }
            ExitCode        = $exitCode
            ValidFiles      = $validFiles
            InvalidFiles    = $invalidFiles
            ErrorMessage    = $errorLog
            OutputMessage   = $outputLog
        }
        
        # Generate HTML report
        New-HTML -Title "PST Upload Report" -FilePath $htmlPath -ShowHTML {
            New-HTMLSection -HeaderText "Upload Summary" {
                New-HTMLPanel {
                    New-HTMLText -Text @"
                    <h3>Upload Details</h3>
                    <ul>
                        <li>Total Files: $($reportData.TotalFiles)</li>
                        <li>Total Size: $($reportData.TotalSizeGB) GB</li>
                        <li>Upload Time: $($reportData.UploadStartTime)</li>
                        <li>Status: $($reportData.Status)</li>
                        <li>Exit Code: $($reportData.ExitCode)</li>
                    </ul>
"@
                }
            }
            
            if ($invalidFiles) {
                New-HTMLSection -HeaderText "Skipped Files" {
                    New-HTMLTable -DataTable $invalidFiles -ScrollX
                }
            }
            
            New-HTMLSection -HeaderText "Uploaded Files" {
                New-HTMLTable -DataTable $validFiles -ScrollX
            }
            
            if ($exitCode -ne 0) {
                New-HTMLSection -HeaderText "Error Details" {
                    New-HTMLPanel {
                        New-HTMLText -Text @"
                        <h3>Error Log</h3>
                        <pre>$($reportData.ErrorMessage)</pre>
                        <h3>Output Log</h3>
                        <pre>$($reportData.OutputMessage)</pre>
"@
                    }
                }
            }
        }
        
        # Display results
        if ($exitCode -eq 0) {
            Write-Host "`nPST upload completed successfully!" -ForegroundColor Green
        }
        else {
            Write-Host "`nPST upload failed with exit code: $exitCode" -ForegroundColor Red
        }
        Write-Host "Report generated at: $htmlPath" -ForegroundColor Cyan
        
        return $reportData
    }
    catch {
        Write-Error "Error during upload: $_"
        throw
    }
}

# Example usage
$uploadParams = @{
    PSTDirectory = "D:\ExchangeArchives\Import"
    SasUrl       = "https://ebb534ec5c8c49568557303.blob.core.windows.net/ingestiondata?sv=2015-04-05&sr=c&si=IngestionSasForAzCopy202412191529268229&sig=2BJ6DFHeKvRYgvFfa1uxhnVm9sspSgpoxOIIKlH%2BSPQ%3D&se=2025-01-18T16%3A31%3A10Z"
    LogPath      = ".\PST_Upload_Logs"
    MaxSizeGB    = 20
}

Start-PSTUpload @uploadParams