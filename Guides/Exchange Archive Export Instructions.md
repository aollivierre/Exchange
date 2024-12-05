# Exchange Archive PST Export Instructions

## Prerequisites
- Exchange Management Shell access
- Exchange Server administrative rights
- Sufficient disk space on local drive
- Network access to destination share

## Step 1: Export Archive Mailboxes to Local Drive

1. Copy script https://raw.githubusercontent.com/aollivierre/Exchange/refs/heads/main/Exchange2/1013/RI/Optimized%20Bulk%20Archive%20Export%20Script.ps1 to Exchange Server (e.g., `C:\Code\ExportArchives.ps1`)
2. Open Exchange Management Shell as Administrator
3. Navigate to script location:
```powershell
cd C:\Code
```
4. Run the export script:
```powershell
.\ExportArchives.ps1
```
5. Monitor the export progress in console window
6. Check local drive for PST files when complete

## Step 2: Transfer PSTs to Network Location

1. Copy robocopy script https://raw.githubusercontent.com/aollivierre/Exchange/refs/heads/main/Exchange2/1013/RI/Robocopy%20PST%20Transfer%20Script.ps1 to same location (e.g., `C:\Code\CopyPSTs.ps1`)
2. Verify network path is accessible
3. Run the robocopy script:
```powershell
.\CopyPSTs.ps1
```
4. Monitor transfer progress
5. Check log file for completion status

## Important Notes
- Run both scripts directly on Exchange Server
- Default concurrent export limit is 2
- Logs are created in destination folder
- Source PSTs remain on local drive after transfer