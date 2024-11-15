Connect-ExchangeOnline


Set-Mailbox -PublicFolder "Mailbox1" -IsHoldForMigration $false


Get-Mailbox -PublicFolder 
Get-Mailbox -PublicFolder | select-object * | Format-List


Remove-Mailbox -PublicFolder -Identity "TestPFMailbox09" -Confirm:$false


New-Mailbox -PublicFolder -Name "TestPFMailbox09"
New-Mailbox -PublicFolder -Name "TestPFMailbox09" -HoldForMigration:$false


New-PublicFolder -Name "TestPF099"


