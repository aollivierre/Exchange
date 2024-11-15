param (
    [string]$IPAddress = '0.0.0.0',
    [int]$Port = 161
)

# Load necessary .NET types.
Add-Type -TypeDefinition @'
    using System;
    using System.Net;
    using System.Net.Sockets;
'@

try {
    # Create UDP client to listen on port 161.
    $udpClient = New-Object System.Net.Sockets.UdpClient $Port

    # Listen indefinitely.
    while ($true) {
        try {
            # Receive SNMP request.
            $remoteEndPoint = New-Object System.Net.IPEndPoint ([System.Net.IPAddress]::Any, 0)
            $bytesReceived = $udpClient.Receive([ref]$remoteEndPoint)
            Write-Output "Received $($bytesReceived.Length) bytes from $($remoteEndPoint.Address):$($remoteEndPoint.Port)"

            # Send a basic response.
            $bytesSent = $udpClient.Send($bytesReceived, $bytesReceived.Length, $remoteEndPoint)
            Write-Output "Sent $($bytesSent) bytes to $($remoteEndPoint.Address):$($remoteEndPoint.Port)"
        }
        catch {
            Write-Error "An error occurred while processing data: $_"
        }
    }
}
catch {
    Write-Error "An error occurred while creating the UDP client: $_"
}



Get-Process -Id (Get-NetUDPEndpoint -LocalPort 161).OwningProcess