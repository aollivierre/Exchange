Get-PublicFolder -Recurse -ResultSize Unlimited | Export-CliXML C:\Code\CB\Exchange\PFScripts\Exports\OnPrem_PFStructure.xml
Get-PublicFolderStatistics -ResultSize Unlimited | Export-CliXML C:\Code\CB\Exchange\PFScripts\Exports\OnPrem_PFStatistics.xml
Get-PublicFolder -Recurse -ResultSize Unlimited | Get-PublicFolderClientPermission | Select-Object Identity,User,AccessRights -ExpandProperty AccessRights | Export-CliXML C:\Code\CB\Exchange\PFScripts\Exports\OnPrem_PFPerms.xml
Get-MailPublicFolder -ResultSize Unlimited | Export-CliXML C:\Code\CB\Exchange\PFScripts\Exports\OnPrem_MEPF.xml