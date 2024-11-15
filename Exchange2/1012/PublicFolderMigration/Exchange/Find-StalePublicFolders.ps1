$date=(Get-Date).addmonths(-24)
Get-publicfolderstatistics -Resultsize unlimited | Where-Object {$_.lastmodificationtime -le $date}



Get-PublicFolderStatistics -ResultSize Unlimited | Sort-Object LastModificationTime | Export-Csv "C:\PublicFolderStats.csv" -NoTypeInformation



Get-OrganizationConfig | Format-List DefaultPublicFolderAgeLimit