# After the public folders are removed, run the following commands to remove all public folder mailboxes:


# Ensure you can retrieve the hierarchy mailbox GUID correctly:
$hierarchyMailboxGuid = (Get-OrganizationConfig).RootPublicFolderMailbox.HierarchyMailboxGuid
Write-Host $hierarchyMailboxGuid

# This should display the GUID value. If not, the issue may be with accessing the RootPublicFolderMailbox.HierarchyMailboxGuid property.


# Simplify the mailbox removal commands for easier debugging:
$nonHierarchyMailboxes = Get-Mailbox -PublicFolder | Where-Object {$_.ExchangeGuid -ne $hierarchyMailboxGuid}
$nonHierarchyMailboxes | ForEach-Object { Write-Host "Removing mailbox: $($_.DisplayName)" }
$nonHierarchyMailboxes | Remove-Mailbox -PublicFolder -Confirm:$false -Force

#no output for the non-hierarchy mailbox removal commands, so no non-hierarchy mailboxes were found

$hierarchyMailboxGuid = (Get-OrganizationConfig).RootPublicFolderMailbox.HierarchyMailboxGuid
$hierarchyMailbox = Get-Mailbox -PublicFolder | Where-Object {$_.ExchangeGuid -eq $hierarchyMailboxGuid}
$hierarchyMailbox | ForEach-Object { Write-Host "Removing hierarchy mailbox: $($_.DisplayName)" }
$hierarchyMailbox | Remove-Mailbox -PublicFolder -Confirm:$false -Force

# Removing hierarchy mailbox: Mailbox1
# Exception: Unable to index into an object of type
# "System.Collections.Generic.Dictionary`2[System.String,System.Collections.Generic.IEnumerable`1[System.String]]".

# By breaking down like this, you can see which mailbox names are being processed, giving a clearer idea of where the problem might be.


# Handle soft-deleted mailboxes:
$softDeletedMailboxes = Get-Mailbox -PublicFolder -SoftDeletedMailbox
$softDeletedMailboxes | ForEach-Object { Write-Host "Removing soft deleted mailbox: $($_.DisplayName)" }
$softDeletedMailboxes | ForEach-Object { Remove-Mailbox -PublicFolder $_.PrimarySmtpAddress -PermanentlyDelete:$true -Force -Confirm:$false }


# Removing soft deleted mailbox: primarymailboxName
# Removing soft deleted mailbox: Mailbox2
# Removing soft deleted mailbox: TestPFMailbox09
# Done - no errors removing soft deleted mailboxes

#on the second attempt, the soft deleted mailboxes were removed successfully as the following error means that the soft deleted mailboxes were already removed:


# Remove-Mailbox: |Microsoft.Exchange.Configuration.CmdletProxyException|Error on proxy command 'Remove-Mailbox -Confirm:$False -PermanentlyDelete:$True
# -Identity:'primarymailboxName_01a1dac7@glebecentre.onmicrosoft.com' -PublicFolder:$True -Force:$True' to server
# YQXPR01MB4433.CANPRD01.PROD.OUTLOOK.COM: Server version 15.20.6699.0000, Proxy method PSWS:  NotFound: Error executing cmdlet : {         
# "code": "NotFound",   "message": "Error executing cmdlet",   "details": [     {       "code": "Client",       "target": "",      
# "message": "Ex6F9304|Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException|The operation couldn't be performed because 
# object 'primarymailboxName_01a1dac7@glebecentre.onmicrosoft.com' couldn't be found on 'YQBPR01A12DC007.CANPR01A012.PROD.OUTLOOK.COM'."    
# }   ],   "innererror": {     "message": "Error executing cmdlet",     "type":
# "Microsoft.Exchange.Admin.OData.Core.ODataServiceException",     "stacktrace": "   at
# Microsoft.Exchange.AdminApi.CommandInvocation.CommandInvocation.InvokeCommand(QueryContext queryContext, CmdletInvokeInputType
# cmdletInvokeInputType)\r\n   at Microsoft.Exchange.Admin.OData.Core.PathSegmentToExpressionTranslator.Translate(OperationImportSegment    
# segment)\r\n   at Microsoft.Exchange.Admin.OData.Core.QueryContext.ResolveQuery(ODataContext context, Int32 level)\r\n   at
# Microsoft.Exchange.Admin.OData.Core.Handlers.OperationHandler.Process(IODataRequestMessage requestMessage, IODataResponseMessage
# responseMessage)\r\n   at Microsoft.Exchange.Admin.OData.Core.Handlers.RequestHandler.Process(Stream requestStream)",    
# "internalexception": {       "message": "Exception of type
# 'Microsoft.Exchange.Management.PSDirectInvoke.DirectInvokeCmdletExecutionException' was thrown.",       "type":
# "Microsoft.Exchange.Management.PSDirectInvoke.DirectInvokeCmdletExecutionException",       "stacktrace": "   at
# Microsoft.Exchange.Management.PSDirectInvoke.PSDirectInvokeCmdletFactory.InvokeCmdletInternal[TCmdlet,TResult](Func`1 createCmdlet,       
# Action`1 setParameterDelegate, List`1 captureAdditionalIO, List`1 directInvokeExceptions, Boolean
# shouldAllowProactiveProxyToReturnRoutingHint)"     }   },   "adminapi.warnings@odata.type": "#Collection(String)",  
# "@adminapi.warnings": [] } [Server=YQXPR01MB5357,RequestId=a9d8a5ae-6d51-d081-580b-00491acad8ba,TimeStamp=8/18/2023 2:00:52 PM] .
# Remove-Mailbox: Ex6F9304|Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException|The operation couldn't be performed because object
# 'Mailbox2_05306997@glebecentre.onmicrosoft.com' couldn't be found on 'YT3PR01A12DC010.CANPR01A012.PROD.OUTLOOK.COM'.
# Remove-Mailbox: Ex6F9304|Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException|The operation couldn't be performed because object
# 'TestPFMailbox09_d5431507@glebecentre.onmicrosoft.com' couldn't be found on 'YT3PR01A12DC010.CANPR01A012.PROD.OUTLOOK.COM'.

# Now the Get-Mailbox -PublicFolder -SoftDeletedMailbox command returns no results, so the soft deleted mailboxes were removed successfully.

# Handle the CNF mailboxes:
$conflictedMailboxes = Get-Mailbox -PublicFolder -SoftDeletedMailbox | Where-Object { $_.Name -like "*CNF:*" -or $_.Identity -like "*CNF:*" }
$conflictedMailboxes | ForEach-Object { Write-Host "Removing CNF mailbox: $($_.DisplayName)" }
$conflictedMailboxes | ForEach-Object { Remove-Mailbox -PublicFolder $_.ExchangeGUID.GUID -RemoveCNFPublicFolderMailboxPermanently -Force -Confirm:$false }

#no output from the CNF mailbox removal commands, so no CNF mailboxes were found






# PS C:\Code> Get-Recipient  -IncludeSoftDeletedRecipients -RecipientTypeDetails publicfoldermailbox |fl Name, OrganizationalUnit, DistinguishedName, ExchangeGuid

# Name               : Mailbox1
# OrganizationalUnit : canpr01a012.prod.outlook.com/Microsoft Exchange Hosted Organizations/glebecentre.onmicrosoft.com
# DistinguishedName  : CN=Mailbox1,OU=glebecentre.onmicrosoft.com,OU=Microsoft Exchange Hosted
#                      Organizations,DC=CANPR01A012,DC=PROD,DC=OUTLOOK,DC=COM
# ExchangeGuid       : 9e2c0327-a7fc-45c5-85e6-1c5b95e808c6

# Name               : TestPFMailbox09
#                      CNF:a31e6034-6e89-4541-b205-070959a06c8f
# OrganizationalUnit : canpr01a012.prod.outlook.com/Microsoft Exchange Hosted Organizations/glebecentre.onmicrosoft.com
# DistinguishedName  : CN=TestPFMailbox09\0ACNF:a31e6034-6e89-4541-b205-070959a06c8f,OU=glebecentre.onmicrosoft.com,OU=Microsoft Exchange   
#                      Hosted Organizations,DC=CANPR01A012,DC=PROD,DC=OUTLOOK,DC=COM
# ExchangeGuid       : f12aab05-3c7c-4987-9ddf-e0837fd4b6d1




#also follow this article when you have Public Folder Mailboxes that are in a "HoldForMigration" or stuck or CNF state:

# https://learn.microsoft.com/en-ca/exchange/troubleshoot/public-folders/public-folder-migration-errors