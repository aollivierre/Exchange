# PST Import Process Guide

## Prerequisites
- Python installed on your system
- ScanPST utility
- Azure Storage Explorer
- Exchange Online Management permissions
- PSWriteHTML PowerShell module

## Step 1: Scan PST Files
Before uploading PST files to Azure, scan them using ScanPST:

1. Use the automated Python script:
   ```bash
   py 'C:\code\Purview\4-automate-scanPST-interaction copy 11.py'
   ```
2. This script will scan all PST files in the specified folder automatically
3. Wait for the scanning process to complete
4. Fix any reported issues before proceeding

## Step 2: Azure Storage Explorer Setup
1. Download and install [Azure Storage Explorer](https://azure.microsoft.com/features/storage-explorer/)
2. Connect to your storage:
   - Select "Blob Container or Directory" option
   - Choose "Shared Access Signature URL (SAS)"
   - This provides read-only access to verify uploads

## Step 3: CSV Mapping File
The script will generate a CSV mapping file with the following considerations:

1. FilePath Column:
   - Must match the folder name in Azure Storage
   - Leave empty if PSTs are in the root folder
   - Example:
     - If PST is in `archive/2024/`, FilePath should be `archive/2024`
     - If PST is in root, FilePath should be empty

2. Mapping Structure:
   - Each PST file maps to its corresponding mailbox
   - Pattern: `username_archive_timestamp.pst` → `username@domain.com`
   - Special handling for certain mailboxes (e.g., ri-support)

## Step 4: Exchange Online Permissions
1. Add Import/Export permissions:
   - Access Exchange Admin Center
   - Navigate to Roles > Organization Management
   - Add Import Export role
   - This is required for the import process

## Step 5: Purview Import Process
1. After assessment completion:
   - Select "Import everything (recommended)"
   - Avoid applying filters during import
   - This ensures complete data migration

## Additional Resources
- [Using AzCopy v10](https://learn.microsoft.com/azure/storage/common/storage-use-azcopy-v10?tabs=dnf#download-azcopy)
- [Network Upload Guide](https://learn.microsoft.com/purview/use-network-upload-to-import-pst-files#step-2-upload-your-pst-files-to-microsoft-365)
- [Complete Import Guide](https://learn.microsoft.com/purview/importing-pst-files-to-office-365#BKMK_NetworkUpload)

## Troubleshooting
1. File Path Issues:
   - Double-check the FilePath in CSV matches Azure Storage exactly
   - Case sensitivity matters
   - Use forward slashes (/) not backslashes (\)

2. Permission Issues:
   - Verify Exchange Online roles
   - Ensure SAS token hasn't expired
   - Check network upload status in Purview

3. Import Failures:
   - Verify PST file integrity using ScanPST
   - Check CSV mapping accuracy
   - Confirm target mailbox existence and permissions

## Support
For additional assistance:
- Review Microsoft's official documentation
- Contact your Microsoft support representative
- Verify all prerequisites before starting the import process
```

The script now includes:
1. Pre-flight checks to ensure PSTs are scanned
2. Important reminders after mapping generation
3. Reference to detailed documentation
4. Clear error messages and guidance

The README.md provides:
1. Comprehensive step-by-step instructions
2. All necessary links and references
3. Troubleshooting guidance
4. Detailed explanation of each component