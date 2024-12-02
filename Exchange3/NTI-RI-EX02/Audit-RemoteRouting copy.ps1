# Function to ensure PSWriteHTML module is installed
function Install-RequiredModule {
    if (-not (Get-Module -ListAvailable -Name PSWriteHTML)) {
        Write-Host "Installing PSWriteHTML module..." -ForegroundColor Yellow
        try {
            Install-Module -Name PSWriteHTML -Force -AllowClobber -Scope CurrentUser
            Import-Module PSWriteHTML
            Write-Host "PSWriteHTML module installed successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Error installing PSWriteHTML module: $_" -ForegroundColor Red
            Write-Host "Please run PowerShell as administrator and try again." -ForegroundColor Yellow
            exit
        }
    }
    else {
        Import-Module PSWriteHTML
    }
}

# Install and import required module
Install-RequiredModule

function Get-AllMailboxRoutingAudit {
    [CmdletBinding()]
    param ()

    try {
        Write-Host "Collecting information for all mailboxes..." -ForegroundColor Cyan
        
        $mailboxes = @()
        
        # Get regular mailboxes
        Write-Host "Getting regular mailboxes..." -ForegroundColor Yellow
        Get-Mailbox -ResultSize Unlimited | ForEach-Object {
            # Find onmicrosoft.com routing address
            $routingAddress = $_.EmailAddresses | 
                Where-Object { $_ -match "smtp:.+\.mail\.onmicrosoft\.com$" } | 
                Select-Object -First 1
            
            $mailboxes += [PSCustomObject]@{
                DisplayName = $_.DisplayName
                Type = "Regular"
                PrimarySmtpAddress = $_.PrimarySmtpAddress
                RoutingAddress = if ($routingAddress) { 
                    $routingAddress.ToString().Replace("smtp:", "") 
                } elseif ($_.ExternalEmailAddress) { 
                    $_.ExternalEmailAddress.ToString() 
                } else { 
                    "Not Set" 
                }
                RecipientType = $_.RecipientTypeDetails
                EmailAddressPolicyEnabled = $_.EmailAddressPolicyEnabled
                EmailAddresses = ($_.EmailAddresses | Where-Object { $_ -match "smtp:" } | ForEach-Object { $_.ToString() }) -join ", "
                Database = $_.Database
            }
        }

        # Get remote mailboxes
        Write-Host "Getting remote mailboxes..." -ForegroundColor Yellow
        Get-RemoteMailbox -ResultSize Unlimited | ForEach-Object {
            $mailboxes += [PSCustomObject]@{
                DisplayName = $_.DisplayName
                Type = "Remote"
                PrimarySmtpAddress = $_.PrimarySmtpAddress
                RoutingAddress = if ($_.RemoteRoutingAddress) { $_.RemoteRoutingAddress.ToString() } else { "Not Set" }
                RecipientType = $_.RecipientTypeDetails
                EmailAddressPolicyEnabled = $_.EmailAddressPolicyEnabled
                EmailAddresses = ($_.EmailAddresses | Where-Object { $_ -match "smtp:" } | ForEach-Object { $_.ToString() }) -join ", "
                Database = "Remote"
            }
        }

        # Calculate summary stats
        $totalMailboxes = $mailboxes.Count
        $regularMailboxes = ($mailboxes | Where-Object { $_.Type -eq "Regular" }).Count
        $remoteMailboxes = ($mailboxes | Where-Object { $_.Type -eq "Remote" }).Count
        $policyEnabled = ($mailboxes | Where-Object { $_.EmailAddressPolicyEnabled }).Count
        $noRouting = ($mailboxes | Where-Object { $_.RoutingAddress -eq "Not Set" }).Count

        # Display summary in console
        Write-Host "`nSummary:" -ForegroundColor Green
        Write-Host "Total Mailboxes: $totalMailboxes" -ForegroundColor Yellow
        Write-Host "Regular Mailboxes: $regularMailboxes" -ForegroundColor Yellow
        Write-Host "Remote Mailboxes: $remoteMailboxes" -ForegroundColor Yellow
        Write-Host "Mailboxes with Email Policy Enabled: $policyEnabled" -ForegroundColor Yellow
        Write-Host "Mailboxes without Routing Address: $noRouting" -ForegroundColor Yellow

        # Output to HTML view
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $mailboxes | Sort-Object Type, DisplayName | 
            Out-HtmlView -Title "All Mailboxes Routing Audit - $timestamp"
    }
    catch {
        Write-Error "Error generating audit report: $_"
        Write-Host "Full Error Details:" -ForegroundColor Red
        $_.Exception | Format-List -Force
    }
}

# Run the audit
Get-AllMailboxRoutingAudit