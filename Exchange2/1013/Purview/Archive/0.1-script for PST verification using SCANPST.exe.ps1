function Start-PSTVerification {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$PSTFolder,
        
        [Parameter()]
        [string]$OutputPath = ".\PST_Verification",
        
        [Parameter()]
        [int]$MaxParallelScans = 2,
        
        [Parameter()]
        [switch]$Recursive
    )
    
    # Find SCANPST.EXE
    $outlookPaths = @(
        "${env:ProgramFiles}\Microsoft Office\root\Office16"
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16"
        "${env:ProgramFiles}\Microsoft Office\Office16"
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16"
    )

    
    
    $scanpstPath = $outlookPaths | ForEach-Object { 
        Join-Path $_ "SCANPST.EXE" 
    } | Where-Object { Test-Path $_ } | Select-Object -First 1
    
    if (-not $scanpstPath) {
        Write-Error "SCANPST.EXE not found in common Office locations"
        return
    }
    
    # Ensure output directory exists
    $null = New-Item -ItemType Directory -Force -Path $OutputPath
    
    # Generate timestamp for reports
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = Join-Path $OutputPath "PST_Verification_$timestamp.csv"
    $htmlPath = Join-Path $OutputPath "PST_Verification_$timestamp.html"
    
    # Initialize results collection
    $results = [System.Collections.Generic.List[object]]::new()
    
    # Get all PST files
    $pstFiles = if ($Recursive) {
        Get-ChildItem -Path $PSTFolder -Filter "*.pst" -Recurse
    }
    else {
        Get-ChildItem -Path $PSTFolder -Filter "*.pst"
    }
    
    foreach ($pst in $pstFiles) {
        try {
            Write-Host "Verifying PST: $($pst.Name)" -ForegroundColor Yellow
            $startTime = Get-Date
            
            # Run SCANPST
            $process = Start-Process -FilePath $scanpstPath -ArgumentList """$($pst.FullName)""" -PassThru -Wait
            
            $endTime = Get-Date
            $duration = ($endTime - $startTime).TotalSeconds
            
            $result = [PSCustomObject]@{
                Name = $pst.Name
                FullPath = $pst.FullName
                SizeMB = [math]::Round($pst.Length / 1MB, 2)
                LastWriteTime = $pst.LastWriteTime
                Status = if ($process.ExitCode -eq 0) { "Success" } else { "Failed" }
                ExitCode = $process.ExitCode
                DurationSeconds = [math]::Round($duration, 2)
                VerificationTime = $startTime
            }
            
            $results.Add($result)
            
            # Console output
            $color = if ($result.Status -eq "Success") { "Green" } else { "Red" }
            Write-Host "Completed: $($pst.Name) - Status: $($result.Status)" -ForegroundColor $color
        }
        catch {
            Write-Warning "Failed to process $($pst.Name): $_"
            $results.Add([PSCustomObject]@{
                Name = $pst.Name
                FullPath = $pst.FullName
                SizeMB = [math]::Round($pst.Length / 1MB, 2)
                LastWriteTime = $pst.LastWriteTime
                Status = "Error"
                ExitCode = -1
                DurationSeconds = 0
                VerificationTime = Get-Date
            })
        }
    }
    
    # Export to CSV
    $results | Export-Csv -Path $csvPath -NoTypeInformation
    
    # Create HTML Report
    New-HTML -TitleText "PST Verification Report" -FilePath $htmlPath -ShowHTML {
        New-HTMLSection -HeaderText "Verification Summary" {
            New-HTMLPanel {
                New-HTMLText -Text @"
                <h3>Report Details</h3>
                <ul>
                    <li>Generated On: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</li>
                    <li>Total PST Files: $($results.Count)</li>
                    <li>Successful: $(($results | Where-Object Status -eq "Success").Count)</li>
                    <li>Failed: $(($results | Where-Object Status -eq "Failed").Count)</li>
                    <li>Errors: $(($results | Where-Object Status -eq "Error").Count)</li>
                    <li>Total Size (MB): $([math]::Round(($results | Measure-Object -Property SizeMB -Sum).Sum, 2))</li>
                </ul>
"@
            }
        }
        
        New-HTMLSection -HeaderText "Verification Details" {
            New-HTMLTable -DataTable $results -ScrollX -Buttons @('copyHtml5', 'excelHtml5', 'csvHtml5') -SearchBuilder {
                New-TableCondition -Name 'Status' -ComparisonType string -Operator eq -Value 'Failed' -BackgroundColor Salmon -Color Black
                New-TableCondition -Name 'Status' -ComparisonType string -Operator eq -Value 'Success' -BackgroundColor LightGreen -Color Black
                New-TableCondition -Name 'Status' -ComparisonType string -Operator eq -Value 'Error' -BackgroundColor Red -Color White
            }
        }
    }
    
    # Console Summary
    Write-Host "`nPST Verification Summary:" -ForegroundColor Cyan
    Write-Host "Total PST Files: $($results.Count)" -ForegroundColor Green
    Write-Host "Successful: $(($results | Where-Object Status -eq "Success").Count)" -ForegroundColor Green
    Write-Host "Failed: $(($results | Where-Object Status -eq "Failed").Count)" -ForegroundColor Yellow
    Write-Host "Errors: $(($results | Where-Object Status -eq "Error").Count)" -ForegroundColor Red
    Write-Host "`nReports generated:" -ForegroundColor Cyan
    Write-Host "CSV Report: $csvPath" -ForegroundColor Green
    Write-Host "HTML Report: $htmlPath" -ForegroundColor Green
    
    return $results
}

# Example usage
$verificationParams = @{
    PSTFolder = "D:\ExchangeArchives\Import"
    OutputPath = ".\PST_Verification"
    Recursive = $true  # Include subfolders
    MaxParallelScans = 2  # Limit concurrent scans
}

Start-PSTVerification @verificationParams