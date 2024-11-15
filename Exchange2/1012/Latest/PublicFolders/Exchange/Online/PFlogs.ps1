# PS C:\Code> Remove-Mailbox -PublicFolder -Identity "TestPFMailbox09" -Confirm:$false
# PS C:\Code> Get-Mailbox -PublicFolder 

# Name                      Alias           Database                       ProhibitSendQuota    ExternalDirectoryObjectId
# ----                      -----           --------                       -----------------    -------------------------
# Mailbox1                  Mailbox1_d14a5… CANPR01DG184-db019             99 GB (106,300,440,…
# Mailbox2                  Mailbox2_05306… CANPR01DG038-db080             99 GB (106,300,440,…

# PS C:\Code> New-Mailbox -PublicFolder -Name "TestPFMailbox09" -HoldForMigration:$false

# Name                      Alias           Database                       ProhibitSendQuota    ExternalDirectoryObjectId
# ----                      -----           --------                       -----------------    -------------------------
# TestPFMailbox09           TestPFMailbox0… CANPR01DG635-db106             99 GB (106,300,440,…

# PS C:\Code> Get-Mailbox -PublicFolder

# Name                      Alias           Database                       ProhibitSendQuota    ExternalDirectoryObjectId
# ----                      -----           --------                       -----------------    -------------------------
# Mailbox1                  Mailbox1_d14a5… CANPR01DG184-db019             99 GB (106,300,440,…
# Mailbox2                  Mailbox2_05306… CANPR01DG038-db080             99 GB (106,300,440,…
# TestPFMailbox09           TestPFMailbox0… CANPR01DG635-db106             99 GB (106,300,440,…

# PS C:\Code> Get-Mailbox -PublicFolder | Where-Object { $_.IsRootPublicFolderMailbox -eq $true }

# Name                      Alias           Database                       ProhibitSendQuota    ExternalDirectoryObjectId
# ----                      -----           --------                       -----------------    -------------------------
# Mailbox1                  Mailbox1_d14a5… CANPR01DG184-db019             99 GB (106,300,440,…

# PS C:\Code> New-PublicFolder -Name "TestPF099"
# New-PublicFolder: |Microsoft.Exchange.Data.StoreObjects.ObjectNotFoundException|No active public folder mailboxes were found zation
# glebecentre.onmicrosoft.com. This happens when no public folder mailboxes are provisioned or they are provisioned in
# 'HoldForMigration' mode. If you're not currently performing a migration, create a public folder mailbox.
# PS C:\Code> Get-Mailbox -PublicFolder | Select-Object Name,IsHoldForMigration

# Name            IsHoldForMigration
# ----            ------------------
# Mailbox1
# Mailbox2
# TestPFMailbox09

# PS C:\Code> Get-Mailbox -PublicFolder | Select-Object Name,IsHoldForMigration

# Name            IsHoldForMigration
# ----            ------------------
# Mailbox1
# Mailbox2
# TestPFMailbox09

# PS C:\Code> Set-Mailbox -PublicFolder "Mailbox1" -IsHoldForMigration $false
# Set-Mailbox: A parameter cannot be found that matches parameter name 'IsHoldForMigration'.
# PS C:\Code> Set-Mailbox -PublicFolder "Mailbox1" -HoldForMigration $false  
# Set-Mailbox: A parameter cannot be found that matches parameter name 'HoldForMigration'.
# PS C:\Code> Get-PublicFolder
# Get-PublicFolder: |Microsoft.Exchange.Data.StoreObjects.ObjectNotFoundException|No active public folder mailboxes were found zation
# glebecentre.onmicrosoft.com. This happens when no public folder mailboxes are provisioned or they are provisioned in
# 'HoldForMigration' mode. If you're not currently performing a migration, create a public folder mailbox.
# PS C:\Code> Get-Mailbox -PublicFolder 

# Name                      Alias           Database                       ProhibitSendQuota    ExternalDirectoryObjectId
# ----                      -----           --------                       -----------------    -------------------------
# Mailbox1                  Mailbox1_d14a5… CANPR01DG184-db019             99 GB (106,300,440,…
# Mailbox2                  Mailbox2_05306… CANPR01DG038-db080             99 GB (106,300,440,…
# TestPFMailbox09           TestPFMailbox0… CANPR01DG635-db106             99 GB (106,300,440,…

# PS C:\Code> get-publicFolder
# Get-PublicFolder: |Microsoft.Exchange.Data.StoreObjects.ObjectNotFoundException|No active public folder mailboxes were found zation
# glebecentre.onmicrosoft.com. This happens when no public folder mailboxes are provisioned or they are provisioned in
# 'HoldForMigration' mode. If you're not currently performing a migration, create a public folder mailbox.
# PS C:\Code> (Get-OrganizationConfig).RootPublicFolderMailbox

# IsValid              : True
# CanUpdate            : True
# Type                 : InTransitMailboxGuid
# HierarchyMailboxGuid : 9e2c0327-a7fc-45c5-85e6-1c5b95e808c6
# LockedForMigration   : True


# PS C:\Code> Get-Mailbox -PublicFolder | Where-Object { $_.IsRootPublicFolderMailbox -eq $true }

# Name                      Alias           Database                       ProhibitSendQuota    ExternalDirectoryObjectId
# ----                      -----           --------                       -----------------    -------------------------
# 00,440,…