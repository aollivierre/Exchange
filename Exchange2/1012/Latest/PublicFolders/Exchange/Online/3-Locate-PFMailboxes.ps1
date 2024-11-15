Get-Mailbox -PublicFolder

Get-PublicFolder -Recurse


# 1. Export Public Folder Statistics from Exchange Online:
Get-PublicFolderStatistics -ResultSize Unlimited | Sort-Object LastModificationTime
# Get-PublicFolderStatistics: |Microsoft.Exchange.Data.StoreObjects.ObjectNotFoundException|No active public folder mailboxes were found for organization
# glebecentre.onmicrosoft.com. This happens when no public folder mailboxes are provisioned or they are provisioned in 'HoldForMigration'
# mode. If you're not currently performing a migration, create a public folder mailbox.



# 2. Export Public Folder Mailbox Statistics from Exchange Online:
Get-Mailbox -PublicFolder | Get-MailboxStatistics | Format-List


# Get-Mailbox -PublicFolder | Get-MailboxStatistics

# DisplayName               ItemCount    StorageLimitStatus                                                                   LastLogonTime
# -----------               ---------    ------------------                                                                   ------------- 
# Mailbox1                  2759
# Mailbox2                  2

# PS C:\Code> Get-Mailbox -PublicFolder | Where-Object { $_.IsRootPublicFolderMailbox -eq $true } | select *

# Database                                  : CANPR01DG184-db019
# DatabaseGuid                              : 07b3c28a-06c6-48c4-a459-0e0ec500b5c8
# MailboxProvisioningConstraint             : 
# IsMonitoringMailbox                       : False
# MailboxRegion                             : 
# MailboxRegionLastUpdateTime               : 
# MailboxRegionSuffix                       : None
# MessageRecallProcessingEnabled            : True
# MessageCopyForSentAsEnabled               : False
# MessageCopyForSendOnBehalfEnabled         : False
# MailboxProvisioningPreferences            : {}
# UseDatabaseRetentionDefaults              : False
# RetainDeletedItemsUntilBackup             : False
# DeliverToMailboxAndForward                : False
# IsExcludedFromServingHierarchy            : False
# IsHierarchyReady                          : True
# IsHierarchySyncEnabled                    : True
# IsPublicFolderSystemMailbox               : False
# HasSnackyAppData                          : False
# LitigationHoldEnabled                     : False
# SingleItemRecoveryEnabled                 : True
# RetentionHoldEnabled                      : False
# EndDateForRetentionHold                   : 
# StartDateForRetentionHold                 : 
# RetentionComment                          : 
# RetentionUrl                              : 
# LitigationHoldDate                        : 
# LitigationHoldOwner                       : 
# ElcProcessingDisabled                     : False
# ComplianceTagHoldApplied                  : False
# WasInactiveMailbox                        : False
# DelayHoldApplied                          : False
# DelayReleaseHoldApplied                   : False
# PitrEnabled                               : False
# PitrCopyIntervalInSeconds                 : 0
# PitrPaused                                : False
# PitrPausedTimestamp                       : 
# PitrOffboardedTimestamp                   : 
# InactiveMailboxRetireTime                 : 
# OrphanSoftDeleteTrackingTime              : 
# LitigationHoldDuration                    : Unlimited
# ManagedFolderMailboxPolicy                : 
# RetentionPolicy                           : Default MRM Policy
# AddressBookPolicy                         : 
# CalendarRepairDisabled                    : False
# ExchangeGuid                              : 9e2c0327-a7fc-45c5-85e6-1c5b95e808c6
# MailboxContainerGuid                      : 
# UnifiedMailbox                            : 
# MailboxLocations                          : {1;9e2c0327-a7fc-45c5-85e6-1c5b95e808c6;Primary;CANPRD01.PROD.OUTLOOK.COM;07b3c28a-06c6-48c4- 
#                                             a459-0e0ec500b5c8}
# AggregatedMailboxGuids                    : {}
# ExchangeSecurityDescriptor                : System.Security.AccessControl.RawSecurityDescriptor
# ExchangeUserAccountControl                : AccountDisabled
# AdminDisplayVersion                       : Version 15.20 (Build 6699.20)
# MessageTrackingReadStatusEnabled          : True
# ExternalOofOptions                        : External
# ForwardingAddress                         : 
# ForwardingSmtpAddress                     : 
# RetainDeletedItemsFor                     : 14.00:00:00
# IsMailboxEnabled                          : True
# Languages                                 : {}
# OfflineAddressBook                        : 
# ProhibitSendQuota                         : 99 GB (106,300,440,576 bytes)
# ProhibitSendReceiveQuota                  : 100 GB (107,374,182,400 bytes)
# RecoverableItemsQuota                     : 30 GB (32,212,254,720 bytes)
# RecoverableItemsWarningQuota              : 20 GB (21,474,836,480 bytes)
# CalendarLoggingQuota                      : 6 GB (6,442,450,944 bytes)
# DowngradeHighPriorityMessagesEnabled      : False
# ProtocolSettings                          : {RemotePowerShell§1, MAPI§1§0§§§0§§§§§0, PublicFolderClientAccess§0, IMAP4§1§§§§§§§§§§§§…}    
# RecipientLimits                           : 500
# ImListMigrationCompleted                  : False
# SiloName                                  : 
# IsResource                                : False
# IsLinked                                  : False
# IsShared                                  : True
# IsRootPublicFolderMailbox                 : True
# LinkedMasterAccount                       : NT AUTHORITY\SELF
# ResetPasswordOnNextLogon                  : False
# ResourceCapacity                          : 
# ResourceCustom                            : {}
# ResourceType                              : 
# RoomMailboxAccountEnabled                 : 
# SamAccountName                            : $BKQRE0-285FRRS5MG0N
# SCLDeleteThreshold                        : 
# SCLDeleteEnabled                          : 
# SCLRejectThreshold                        : 
# SCLRejectEnabled                          : 
# SCLQuarantineThreshold                    : 
# SCLQuarantineEnabled                      : 
# SCLJunkThreshold                          : 
# SCLJunkEnabled                            : 
# AntispamBypassEnabled                     : False
# ServerLegacyDN                            : /o=ExchangeLabs/ou=Exchange Administrative Group
#                                             (FYDIBOHF23SPDLT)/cn=Configuration/cn=Servers/cn=YT1PR01MB2940
# ServerName                                : yt1pr01mb2940
# UseDatabaseQuotaDefaults                  : False
# IssueWarningQuota                         : 98 GB (105,226,698,752 bytes)
# RulesQuota                                : 256 KB (262,144 bytes)
# Office                                    : 
# UserPrincipalName                         : Mailbox1_d14a5d5d@glebecentre.ca
# UMEnabled                                 : False
# MaxSafeSenders                            : 
# MaxBlockedSenders                         : 
# NetID                                     : 
# ReconciliationId                          : 
# WindowsLiveID                             : 
# MicrosoftOnlineServicesID                 : 
# ThrottlingPolicy                          : 
# RoleAssignmentPolicy                      : Default Role Assignment Policy
# DefaultPublicFolderMailbox                : 
# EffectivePublicFolderMailbox              : 
# SharingPolicy                             : Default Sharing Policy
# RemoteAccountPolicy                       : 
# MailboxPlan                               : ExchangeOnlineEnterprise-c4427fdf-21e1-4411-92e2-b4b154ee3a3b
# ArchiveDatabase                           : 
# ArchiveDatabaseGuid                       : 00000000-0000-0000-0000-000000000000
# ArchiveGuid                               : 00000000-0000-0000-0000-000000000000
# ArchiveName                               : {}
# JournalArchiveAddress                     : 
# ArchiveQuota                              : 100 GB (107,374,182,400 bytes)
# ArchiveWarningQuota                       : 90 GB (96,636,764,160 bytes)
# ArchiveDomain                             : 
# ArchiveStatus                             : None
# ArchiveState                              : None
# AutoExpandingArchiveEnabled               : False
# DisabledMailboxLocations                  : False
# RemoteRecipientType                       : None
# DisabledArchiveDatabase                   : 
# DisabledArchiveGuid                       : 00000000-0000-0000-0000-000000000000
# QueryBaseDN                               : 
# QueryBaseDNRestrictionEnabled             : False
# MailboxMoveTargetMDB                      : 
# MailboxMoveSourceMDB                      : 
# MailboxMoveFlags                          : None
# MailboxMoveRemoteHostName                 : 
# MailboxMoveBatchName                      : 
# MailboxMoveStatus                         : None
# MailboxRelease                            : 
# ArchiveRelease                            : 
# IsPersonToPersonTextMessagingEnabled      : False
# IsMachineToPersonTextMessagingEnabled     : False
# UserSMimeCertificate                      : {}
# UserCertificate                           : {}
# CalendarVersionStoreDisabled              : False
# ImmutableId                               : 
# PersistedCapabilities                     : {BPOS_S_Enterprise}
# SKUAssigned                               : 
# AuditEnabled                              : False
# AuditLogAgeLimit                          : 90.00:00:00
# AuditAdmin                                : {Update, MoveToDeletedItems, SoftDelete, HardDelete…}
# AuditDelegate                             : {Update, MoveToDeletedItems, SoftDelete, HardDelete…}
# AuditOwner                                : {Update, MoveToDeletedItems, SoftDelete, HardDelete…}
# DefaultAuditSet                           : {Admin, Delegate, Owner}
# WhenMailboxCreated                        : 2022-11-30 3:52:20 PM
# SourceAnchor                              : 
# UsageLocation                             : 
# IsSoftDeletedByRemove                     : False
# IsSoftDeletedByDisable                    : False
# IsInactiveMailbox                         : False
# IncludeInGarbageCollection                : False
# WhenSoftDeleted                           : 
# InPlaceHolds                              : {}
# GeneratedOfflineAddressBooks              : {}
# AccountDisabled                           : True
# StsRefreshTokensValidFrom                 : 
# NonCompliantDevices                       : {}
# EnforcedTimestamps                        : 
# DataEncryptionPolicy                      : 
# MessageCopyForSMTPClientSubmissionEnabled : True
# RecipientThrottlingThreshold              : Standard
# SharedEmailDomainTenant                   : 
# SharedEmailDomainState                    : None
# SharedWithTargetSmtpAddress               : 
# SharedEmailDomainStateLastModified        : 
# EmailAddressDisplayNames                  : {}
# ResourceProvisioningOptions               : {}
# Extensions                                : {}
# HasPicture                                : False
# HasSpokenName                             : False
# IsDirSynced                               : False
# AcceptMessagesOnlyFrom                    : {}
# AcceptMessagesOnlyFromDLMembers           : {}
# AcceptMessagesOnlyFromSendersOrMembers    : {}
# AddressListMembership                     : {\PublicFolderHierarchyMailboxes(VLV), \PublicFolderMailboxes(VLV), \All Recipients(VLV)}     
# AdministrativeUnits                       : {}
# Alias                                     : Mailbox1_d14a5d5d
# ArbitrationMailbox                        : 
# BypassModerationFromSendersOrMembers      : {}
# OrganizationalUnit                        : canpr01a012.prod.outlook.com/Microsoft Exchange Hosted
#                                             Organizations/glebecentre.onmicrosoft.com
# CustomAttribute1                          : 
# CustomAttribute10                         : 
# CustomAttribute11                         : 
# CustomAttribute12                         : 
# CustomAttribute13                         : 
# CustomAttribute14                         : 
# CustomAttribute15                         : 
# CustomAttribute2                          : 
# CustomAttribute3                          : 
# CustomAttribute4                          : 
# CustomAttribute5                          : 
# CustomAttribute6                          : 
# CustomAttribute7                          : 
# CustomAttribute8                          : 
# CustomAttribute9                          : 
# ExtensionCustomAttribute1                 : {}
# ExtensionCustomAttribute2                 : {}
# ExtensionCustomAttribute3                 : {}
# ExtensionCustomAttribute4                 : {}
# ExtensionCustomAttribute5                 : {}
# DisplayName                               : Mailbox1
# EmailAddresses                            : {SMTP:Mailbox1_d14a5d5d@glebecentre.onmicrosoft.com}
# GrantSendOnBehalfTo                       : {}
# ExternalDirectoryObjectId                 : 
# HiddenFromAddressListsEnabled             : True
# LastExchangeChangedTime                   : 
# LegacyExchangeDN                          : /o=ExchangeLabs/ou=Exchange Administrative Group
#                                             (FYDIBOHF23SPDLT)/cn=Recipients/cn=2c60602d9e5446c9a55f42df8f7c8999-Mailbox1
# MaxSendSize                               : 35 MB (36,700,160 bytes)
# MaxReceiveSize                            : 36 MB (37,748,736 bytes)
# ModeratedBy                               : {}
# ModerationEnabled                         : False
# PoliciesIncluded                          : {deab9981-f887-4530-91f5-d54118c394ec, {26491cfc-9e50-4857-861b-0cb8df22b5d7}}
# PoliciesExcluded                          : {}
# EmailAddressPolicyEnabled                 : True
# PrimarySmtpAddress                        : Mailbox1_d14a5d5d@glebecentre.onmicrosoft.com
# RecipientType                             : UserMailbox
# RecipientTypeDetails                      : PublicFolderMailbox
# RejectMessagesFrom                        : {}
# RejectMessagesFromDLMembers               : {}
# RejectMessagesFromSendersOrMembers        : {}
# RequireSenderAuthenticationEnabled        : False
# SimpleDisplayName                         : 
# SendModerationNotifications               : Always
# UMDtmfMap                                 : {emailAddress:6245269131425353, lastNameFirstName:62452691, firstNameLastName:62452691}       
# WindowsEmailAddress                       : Mailbox1_d14a5d5d@glebecentre.onmicrosoft.com
# MailTip                                   : 
# MailTipTranslations                       : {}
# Identity                                  : Mailbox1
# Id                                        : Mailbox1
# IsValid                                   : True
# ExchangeVersion                           : 1.1 (15.0.0.0)
# Name                                      : Mailbox1
# DistinguishedName                         : CN=Mailbox1,OU=glebecentre.onmicrosoft.com,OU=Microsoft Exchange Hosted
#                                             Organizations,DC=CANPR01A012,DC=PROD,DC=OUTLOOK,DC=COM
# ObjectCategory                            : CANPR01A012.PROD.OUTLOOK.COM/Configuration/Schema/Person
# ObjectClass                               : {top, person, organizationalPerson, user}
# WhenChanged                               : 2022-11-30 3:52:26 PM
# WhenCreated                               : 2022-11-30 3:52:20 PM
# WhenChangedUTC                            : 2022-11-30 8:52:26 PM
# WhenCreatedUTC                            : 2022-11-30 8:52:20 PM
# ExchangeObjectId                          : d1448838-1b13-41aa-91f1-69020c41ac92
# OrganizationalUnitRoot                    : glebecentre.onmicrosoft.com
# OrganizationId                            : CANPR01A012.PROD.OUTLOOK.COM/Microsoft Exchange Hosted 
#                                             Organizations/glebecentre.onmicrosoft.com -
#                                             CANPR01A012.PROD.OUTLOOK.COM/ConfigurationUnits/glebecentre.onmicrosoft.com/Configuration     
# Guid                                      : d1448838-1b13-41aa-91f1-69020c41ac92
# OriginatingServer                         : YT3PR01A12DC010.CANPR01A012.PROD.OUTLOOK.COM
# ObjectState                               : Unchanged