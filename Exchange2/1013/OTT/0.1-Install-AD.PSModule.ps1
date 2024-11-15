# Windows PowerShell
# Copyright (C) Microsoft Corporation. All rights reserved.


Install-WindowsFeature -Name "RSAT-AD-PowerShell"
#                                                                                                                                                                                              Install the latest PowerShell for new features and improvements! https://aka.ms/PSWindows                                                                                                                                                                                                                                                                                                 PS C:\Users\aollivierre> Install-WindowsFeature -Name "RSAT-AD-PowerShell"                                                                                                                   
# Success Restart Needed Exit Code      Feature Result
# ------- -------------- ---------      --------------
# True    No             Success        {Remote Server Administration Tools, Activ...


# PS C:\Users\aollivierre> 

Get-Module -ListAvailable ActiveDirectory


#     Directory: C:\Windows\system32\WindowsPowerShell\v1.0\Modules


# ModuleType Version    Name                                ExportedCommands
# ---------- -------    ----                                ----------------
# Manifest   1.0.1.0    ActiveDirectory                     {Add-ADCentralAccessPolicyMember, Add-ADComputerServiceAccount, Add-ADDomainControllerPasswordReplicationPolicy, Add-ADFineGrai...


# PS C:\Users\aollivierre>