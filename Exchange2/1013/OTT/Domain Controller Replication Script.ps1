# Import required modules
Import-Module ActiveDirectory
Import-Module PSWriteHTML

function Start-DCReplication {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$OutputPath = "$env:USERPROFILE\Desktop\ReplicationReport"
        # [Parameter()]
        # [switch]$Verbose
    )

    function Get-DomainControllers {
        $params = @{
            Filter     = { PrimaryGroupID -eq 516 }
            Properties = 'Name', 'Site'
        }
        Get-ADComputer @params
    }

    function Invoke-Replication {
        param (
            [Parameter(Mandatory)]
            [object[]]$DomainControllers
        )

        $results = [System.Collections.ArrayList]::new()

        foreach ($dc in $DomainControllers) {
            $status = [PSCustomObject]@{
                ComputerName = $dc.Name
                Site        = $dc.Site
                Status      = 'Unknown'
                StartTime   = Get-Date
                EndTime    = $null
                Duration   = $null
            }

            try {
                $params = @{
                    ComputerName = $dc.Name
                    Command     = 'repadmin /syncall /AdeP'
                }
                $null = Invoke-Command @params
                $status.Status = 'Success'
            }
            catch {
                $status.Status = "Failed: $($_.Exception.Message)"
            }

            $status.EndTime = Get-Date
            $status.Duration = $status.EndTime - $status.StartTime
            $null = $results.Add($status)
        }

        $results
    }

    function Export-ReplicationResults {
        param (
            [Parameter(Mandatory)]
            [object[]]$Results,
            [Parameter(Mandatory)]
            [string]$OutputPath
        )

        # Create output directory if it doesn't exist
        if (-not (Test-Path -Path $OutputPath)) {
            $null = New-Item -Path $OutputPath -ItemType Directory
        }

        # Console output
        Write-Host "`nReplication Summary:"
        $Results | Format-Table -AutoSize

        # CSV Export
        $csvPath = Join-Path -Path $OutputPath -ChildPath "ReplicationReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $Results | Export-Csv -Path $csvPath -NoTypeInformation

        # HTML Report
        $htmlPath = Join-Path -Path $OutputPath -ChildPath "ReplicationReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        
        $params = @{
            FilePath = $htmlPath
            Title    = "Domain Controller Replication Report"
            DefaultSortColumn = "StartTime"
            SearchBuilder = $true
        }
        
        New-HTML @params {
            New-HTMLTable -DataTable $Results -Title "Replication Results" -HideFooter
        }

        # Return paths for reference
        [PSCustomObject]@{
            CSVPath  = $csvPath
            HTMLPath = $htmlPath
        }
    }

    # Main execution
    try {
        Write-Verbose "Getting domain controllers..."
        $dcs = Get-DomainControllers

        Write-Verbose "Starting replication..."
        $results = Invoke-Replication -DomainControllers $dcs

        Write-Verbose "Exporting results..."
        $exportPaths = Export-ReplicationResults -Results $results -OutputPath $OutputPath

        Write-Host "`nReports generated:"
        Write-Host "CSV Report: $($exportPaths.CSVPath)"
        Write-Host "HTML Report: $($exportPaths.HTMLPath)"
    }
    catch {
        Write-Error "Error during replication: $_"
    }
}


# Basic usage
Start-DCReplication

# Specify custom output path
# Start-DCReplication -OutputPath "C:\Reports\DCReplication"

# Run with verbose output
# Start-DCReplication -Verbose