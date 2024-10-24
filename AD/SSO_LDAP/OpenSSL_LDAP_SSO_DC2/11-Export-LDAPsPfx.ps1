$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

function Export-LdapsPfx {
    param (
        [string]$thumbprint = '27DBDC90B738F29FDB464F81936D4112A78B7059',
        [string]$pfxFileName = 'LDAPS_PRIVATEKEY.pfx',
        [string]$password = 'enter your password'
    )

    $pfxFilePath = Join-Path -Path $scriptDir -ChildPath $pfxFileName
    $pfxPass = (ConvertTo-SecureString -AsPlainText -Force -String $password)
    
    Get-ChildItem "Cert:\LocalMachine\My\$thumbprint" | Export-PfxCertificate -FilePath $pfxFilePath -Password $pfxPass
}

Export-LdapsPfx
