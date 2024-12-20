#Requires -Modules PSWriteHTML

function Initialize-DownloadLocation {
    param (
        [string]$Path
    )
    
    if (-not (Test-Path -Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
    else {
        Get-ChildItem -Path $Path -File | Remove-Item -Force
    }
}

function Export-AzCopyReport {
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$InstallResults,
        
        [Parameter(Mandatory)]
        [string]$OutputDir
    )
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $htmlPath = Join-Path $OutputDir "AzCopy_Install_Report_$timestamp.html"
    $csvPath = Join-Path $OutputDir "AzCopy_Install_Report_$timestamp.csv"
    
    # Export to CSV
    $InstallResults | Export-Csv -Path $csvPath -NoTypeInformation
    
    $htmlParams = @{
        Title    = "AzCopy Installation Report"
        FilePath = $htmlPath
        ShowHTML = $true
    }

    New-HTML @htmlParams {
        New-HTMLSection -HeaderText "Installation Summary" {
            New-HTMLPanel {
                New-HTMLList -Type Unordered {
                    New-HTMLListItem -Text "Installation Time: $($InstallResults.InstallTime)"
                    New-HTMLListItem -Text "Status: $($InstallResults.Status)"
                    New-HTMLListItem -Text "Version: $($InstallResults.Version)"
                    New-HTMLListItem -Text "Install Path: $($InstallResults.InstallPath)"
                    New-HTMLListItem -Text "PowerShell Version: $($PSVersionTable.PSVersion.ToString())"
                }
            }
        }
        
        New-HTMLSection -HeaderText "Installation Details" {
            $tableParams = @{
                DataTable = $InstallResults
                ScrollX   = $true
                Buttons   = @('copyHtml5', 'excelHtml5', 'csvHtml5')
            }

            New-HTMLTable @tableParams {
                New-TableCondition -Name 'Status' -ComparisonType string -Operator eq -Value 'Failed' -BackgroundColor Salmon -Color Black
                New-TableCondition -Name 'Status' -ComparisonType string -Operator eq -Value 'Success' -BackgroundColor LightGreen -Color Black
            }
        }
    }

    $reportInfo = @{
        CSVPath  = $csvPath
        HTMLPath = $htmlPath
    }

    Write-Host "`nReports generated:" -ForegroundColor Green
    Write-Host "CSV Report: $csvPath" -ForegroundColor Green
    Write-Host "HTML Report: $htmlPath" -ForegroundColor Green
    
    return $reportInfo
}

function Get-AzCopyDownloadUrl {
    $isPowerShell7 = $PSVersionTable.PSVersion.Major -ge 7
    $webParams = @{
        Uri                = 'https://aka.ms/downloadazcopy-v10-windows'
        MaximumRedirection = 0
        ErrorAction        = 'SilentlyContinue'
    }

    if ($isPowerShell7) {
        $webParams['SkipHttpErrorCheck'] = $true
        try {
            $response = Invoke-WebRequest @webParams
            return $response.Headers.Location
        }
        catch {
            throw "Failed to get download URL in PowerShell 7: $_"
        }
    }
    else {
        try {
            $response = Invoke-WebRequest @webParams
        }
        catch {
            return $_.Exception.Response.Headers.Location
        }

        if (-not $response.Headers.Location) {
            throw "Failed to get download URL in Windows PowerShell"
        }
        return $response.Headers.Location
    }
}

function Install-AzCopy {
    param (
        [string]$InstallPath
    )

    $ErrorActionPreference = 'Stop'
    
    try {
        $downloadParams = @{
            Path        = Join-Path -Path $InstallPath -ChildPath 'azcopy.zip'
            ExtractPath = Join-Path -Path $InstallPath -ChildPath 'extracted'
        }
        
        # Get download URL using version-appropriate method
        $downloadUrl = Get-AzCopyDownloadUrl
        
        Write-Host "Downloading AzCopy from: $downloadUrl"
        $webClient = [System.Net.WebClient]::new()
        $webClient.DownloadFile($downloadUrl, $downloadParams.Path)
        
        Write-Host "Extracting AzCopy to: $($downloadParams.ExtractPath)"
        Expand-Archive -Path $downloadParams.Path -DestinationPath $downloadParams.ExtractPath -Force
        
        $azCopyExe = Get-ChildItem -Path $downloadParams.ExtractPath -Recurse -Filter 'azcopy.exe' | 
        Select-Object -First 1
        
        if ($azCopyExe) {
            $finalPath = Join-Path -Path $InstallPath -ChildPath 'azcopy.exe'
            Move-Item -Path $azCopyExe.FullName -Destination $finalPath -Force
            
            # Clean up
            Remove-Item -Path $downloadParams.Path -Force
            Remove-Item -Path $downloadParams.ExtractPath -Recurse -Force
            
            # Test AzCopy and get version
            $version = & $finalPath --version
            
            $results = [PSCustomObject]@{
                InstallTime       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Status            = "Success"
                Version           = $version
                InstallPath       = $finalPath
                DownloadUrl       = $downloadUrl
                PowerShellVersion = $PSVersionTable.PSVersion.ToString()
                Error             = $null
            }
        }
        else {
            throw "Could not find AzCopy executable in extracted files"
        }
    }
    catch {
        $results = [PSCustomObject]@{
            InstallTime       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Status            = "Failed"
            Version           = $null
            InstallPath       = $InstallPath
            DownloadUrl       = $downloadUrl
            PowerShellVersion = $PSVersionTable.PSVersion.ToString()
            Error             = $_.Exception.Message
        }
        
        Write-Error "Failed to install AzCopy: $($_.Exception.Message)"
    }
    
    # Generate reports
    Export-AzCopyReport -InstallResults $results -OutputDir $InstallPath
    
    return $results
}

# Main execution
$installPath = Join-Path -Path $env:USERPROFILE -ChildPath 'AzCopy'
Initialize-DownloadLocation -Path $installPath
Install-AzCopy -InstallPath $installPath