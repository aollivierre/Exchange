$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
function Invoke-LdifdeImport {
    param (
        [string]$fileName = 'enable_ldaps.txt'
    )

    $filePath = Join-Path -Path $scriptDir -ChildPath $fileName

    ldifde -i -f $filePath
}

Invoke-LdifdeImport