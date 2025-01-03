$ProgressPreference = 'SilentlyContinue'

$servers = @{
    CB = @{
        IP = '10.5.1.9'
        FQDN = 'NTI-CB-EX01.CB.tunngavik.local'
    }
 }
 
function Get-ExchangeCredential {
   param ([string]$SecretsFilePath)
   if (Test-Path $SecretsFilePath) {
       try {
           return Import-Clixml -Path $SecretsFilePath
       } catch {
           Write-Warning "Invalid credentials. Enter new credentials."
           Remove-Item -Path $SecretsFilePath -Force
           return Get-ExchangeCredential -SecretsFilePath $SecretsFilePath
       }
   } else {
       Write-Host "Enter domain credentials (domain\username)" -ForegroundColor Yellow
       $Cred = Get-Credential
       $Cred | Export-Clixml -Path $SecretsFilePath
       return $Cred
   }
}

# Specify required cmdlets
$requiredCommands = @(
    'Get-Mailbox',
    'New-MailboxExportRequest' ,
    'Get-MailboxExportRequest' ,
    'Remove-MailboxExportRequest' ,
    'Get-MailboxStatistics',
    'Get-MailboxDatabase'
)

$sessionParams = @{
   IdleTimeout = 2147483
   SkipCACheck = $true
   SkipCNCheck = $true
   SkipRevocationCheck = $true
   MaximumReceivedDataSizePerCommand = 1GB
   MaximumRedirection = 0
   NoCompression = $false
}

foreach ($server in $servers.GetEnumerator()) {
   $kerberosParams = @{
       ConfigurationName = 'Microsoft.Exchange'
       ConnectionUri = "http://$($server.Value.FQDN)/PowerShell/"
       Authentication = 'Kerberos'
       AllowRedirection = $true
       SessionOption = (New-PSSessionOption @sessionParams)
       ErrorAction = 'Stop'
   }

   $basicParams = @{
       ConfigurationName = 'Microsoft.Exchange'
       ConnectionUri = "http://$($server.Value.IP)/PowerShell/"
       Authentication = 'Basic'
       AllowRedirection = $true
       SessionOption = (New-PSSessionOption @sessionParams)
       ErrorAction = 'Stop'
   }

   $importParams = @{
       DisableNameChecking = $true
       AllowClobber = $true
       CommandName = $requiredCommands
   }

   try {
       Write-Host "Attempting Kerberos authentication to $($server.Key)..." -ForegroundColor Cyan
       
       if (-not (Test-Connection -ComputerName $server.Value.FQDN -Count 1 -Quiet)) {
           throw "DNS resolution failed for $($server.Value.FQDN)"
       }
       
       $SecretsFile = Join-Path -Path $PSScriptRoot -ChildPath "secrets.exchange.$($server.Value.FQDN).psd1"
       $kerberosParams.Credential = Get-ExchangeCredential -SecretsFilePath $SecretsFile
       
       $servers[$server.Key].Session = New-PSSession @kerberosParams
       Import-PSSession $servers[$server.Key].Session @importParams
       
       Write-Host "Successfully connected to $($server.Key) using Kerberos!" -ForegroundColor Green
       continue
   }
   catch {
       Write-Host "Kerberos authentication failed: $_" -ForegroundColor Red
       Write-Host "Falling back to Basic authentication..." -ForegroundColor Yellow
       
       try {
           Write-Host "Configuring WinRM for Basic auth..." -ForegroundColor Yellow
           winrm quickconfig -force
           Set-Item WSMan:\localhost\Client\TrustedHosts -Value $server.Value.IP -Force
           Set-Item WSMan:\localhost\Client\AllowUnencrypted -Value $true -Force
           Restart-Service WinRM

           $SecretsFile = Join-Path -Path $PSScriptRoot -ChildPath "secrets.exchange.$($server.Value.IP).psd1"
           $basicParams.Credential = Get-ExchangeCredential -SecretsFilePath $SecretsFile

           $servers[$server.Key].Session = New-PSSession @basicParams
           Import-PSSession $servers[$server.Key].Session @importParams
           
           Write-Host "Successfully connected to $($server.Key) using Basic auth!" -ForegroundColor Green
       }
       catch {
           Write-Error "All connection attempts failed for $($server.Key): $_"
       }
   }
}