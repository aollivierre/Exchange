# Define the backup path
$backupPath = "\\GLB-EX01\publicfolders_export_aug_17_2023"

# Create the directory if it doesn't exist
If(-not (Test-Path $backupPath)) {
    New-Item -Path $backupPath -ItemType Directory
}

# Get all public folder mailboxes
$publicFolderMailboxes = Get-Mailbox -PublicFolder

# Loop through each public folder mailbox and back it up
foreach ($mailbox in $publicFolderMailboxes) {
    $backupFile = "$backupPath\" + $mailbox.Name + "_Backup.pst"
    New-MailboxExportRequest -Mailbox $mailbox.Identity -FilePath $backupFile
}

# Provide feedback to the user
Write-Output "Backup process initiated for all public folder mailboxes."
