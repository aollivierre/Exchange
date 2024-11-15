function Remove-EnhancedRemoteMailbox {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity,

        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    try {
        # Get mailbox information before removal
        $mailbox = Get-RemoteMailbox -Identity $Identity -ErrorAction Stop
        
        Write-Host "Found mailbox to remove:" -ForegroundColor Cyan
        Write-Host "DisplayName: $($mailbox.DisplayName)" -ForegroundColor Yellow
        Write-Host "Primary SMTP: $($mailbox.PrimarySmtpAddress)" -ForegroundColor Yellow
        Write-Host "Remote Routing: $($mailbox.RemoteRoutingAddress)" -ForegroundColor Yellow
        Write-Host "Recipient Type: $($mailbox.RecipientTypeDetails)" -ForegroundColor Yellow

        if (-not $Force) {
            $confirmation = Read-Host "Are you sure you want to remove this mailbox? (Y/N)"
            if ($confirmation -ne 'Y') {
                Write-Host "Operation cancelled." -ForegroundColor Yellow
                return
            }
        }

        # First remove the remote mailbox
        Write-Host "Removing remote mailbox..." -ForegroundColor Yellow
        Disable-RemoteMailbox -Identity $Identity -Confirm:$false

        # Get the AD user
        $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$($mailbox.UserPrincipalName)'" -ErrorAction SilentlyContinue
        
        if ($adUser) {
            Write-Host "Removing AD user account..." -ForegroundColor Yellow
            Remove-ADUser -Identity $adUser.DistinguishedName -Confirm:$false
        }

        # Force AD replication
        Write-Host "Forcing AD replication..." -ForegroundColor Yellow
        Get-ADDomainController -Filter * | ForEach-Object {
            $dc = $_.Name
            Write-Host "Replicating changes to $dc..." -ForegroundColor Gray
            Invoke-Expression "Repadmin /syncall $dc /AeD"
        }

        # Verify removal
        Start-Sleep -Seconds 5
        $verifyMailbox = Get-RemoteMailbox -Identity $Identity -ErrorAction SilentlyContinue
        $verifyUser = Get-ADUser -Filter "UserPrincipalName -eq '$($mailbox.UserPrincipalName)'" -ErrorAction SilentlyContinue

        if (-not $verifyMailbox -and -not $verifyUser) {
            Write-Host "`nMailbox and associated objects successfully removed!" -ForegroundColor Green
        }
        else {
            Write-Warning "Some objects may still exist:"
            if ($verifyMailbox) { Write-Warning "Remote mailbox still exists" }
            if ($verifyUser) { Write-Warning "AD user still exists" }
        }
    }
    catch {
        Write-Error "Error removing mailbox: $_"
        Write-Host "Full Error Details:" -ForegroundColor Red
        $_.Exception | Format-List -Force
    }
}

# Remove the DMARC mailbox
Write-Host "Preparing to remove DMARC shared mailbox..." -ForegroundColor Cyan
Remove-EnhancedRemoteMailbox -Identity "DMARC" -Force