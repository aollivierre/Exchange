# On the local machine
Get-Service -Name 'SNMP' | Select-Object Status, DisplayName, StartType

# # On a remote machine
# Invoke-Command -ComputerName 192.168.2.47 -ScriptBlock {
#     Get-Service -Name 'SNMP' | Select-Object Status, DisplayName, StartType
# } -Credential (Get-Credential)



# On local machine
New-NetFirewallRule -DisplayName "SNMP Port 161" -Direction Inbound -LocalPort 161 -Protocol TCP -Action Allow

# # On remote machine
# Invoke-Command -ComputerName 192.168.2.47 -ScriptBlock {
#     New-NetFirewallRule -DisplayName "SNMP Port 161" -Direction Inbound -LocalPort 161 -Protocol TCP -Action Allow
# } -Credential (Get-Credential)
