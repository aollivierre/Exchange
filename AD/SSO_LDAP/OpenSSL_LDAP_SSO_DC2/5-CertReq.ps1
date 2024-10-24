$currentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
function New-ADCSR {
    $requestInfPath = Join-Path -Path $currentDir -ChildPath "request.inf"
    $adCsrPath = Join-Path -Path $currentDir -ChildPath "ad.csr"

    certreq -new $requestInfPath $adCsrPath
}

New-ADCSR