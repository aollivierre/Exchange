$currentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
function AcceptADCSR {
    $certpath = Join-Path -Path $currentDir -ChildPath "ad_ldaps_cert.crt"

    certreq -accept $certpath
}

AcceptADCSR