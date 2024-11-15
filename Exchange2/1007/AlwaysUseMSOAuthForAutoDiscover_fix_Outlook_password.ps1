#create the following as a scheduled task

# C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -Command "Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Exchange' -Name 'AlwaysUseMSOAuthForAutoDiscover' -Value 1 -Type DWord"




C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -Command "if (-Not (Test-Path 'HKCU:\Software\Microsoft\Exchange')) { New-Item -Path 'HKCU:\Software\Microsoft\Exchange' -Force }; if (-Not (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Exchange' -Name 'AlwaysUseMSOAuthForAutoDiscover' -ErrorAction SilentlyContinue)) { Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Exchange' -Name 'AlwaysUseMSOAuthForAutoDiscover' -Value 1 }"



C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -Command "if (Test-Path 'HKCU:\Software\Microsoft\Exchange') { Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Exchange' -Name 'AlwaysUseMSOAuthForAutoDiscover' -ErrorAction SilentlyContinue; if (-Not (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Exchange' -ErrorAction SilentlyContinue)) { Remove-Item -Path 'HKCU:\Software\Microsoft\Exchange' -Recurse -Force } }"



C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File "C:\code\Outlook\AlwaysUseMSOAuthForAutoDiscover.ps1"

