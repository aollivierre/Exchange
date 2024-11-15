param (
    [string]$IPAddress = '127.0.0.1', # Localhost IP Address.
    [int]$Port = 161                   # Default SNMP port.
)

# Load necessary .NET types.
Add-Type -TypeDefinition @'
    using System;
    using System.Net;
    using System.Net.Sockets;
    using System.Text;
'@

try {
    # Create UDP client.
    $udpClient = New-Object System.Net.Sockets.UdpClient

    # Connect to the local machine on the specified port.
    $udpClient.Connect($IPAddress, $Port)

    # Send a test message.
    $message = [Text.Encoding]::ASCII.GetBytes("Test message")
    $bytesSent = $udpClient.Send($message, $message.Length)
    Write-Output "Sent $($bytesSent) bytes to $IPAddress:$Port"

    # Set a timeout to receive the response (5 seconds).
    $udpClient.Client.ReceiveTimeout = 5000

    # Try to receive a response.
    $remoteEndPoint = New-Object System.Net.IPEndPoint ([System.Net.IPAddress]::Any, 0)
    $bytesReceived = $udpClient.Receive([ref]$remoteEndPoint)

    # If a response is received, print it.
    if ($bytesReceived) {
        Write-Output "Received $($bytesReceived.Length) bytes from $($remoteEndPoint.Address):$($remoteEndPoint.Port)"
    }
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    # Clean up by closing the UDP client.
    $udpClient.Close()
}
