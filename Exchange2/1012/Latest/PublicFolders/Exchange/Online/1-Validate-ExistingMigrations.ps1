Get-MigrationBatch | Where-Object {$_.MigrationType.ToString() -eq "PublicFolder"}


# Identity              Status  Type         TotalCount
# --------              ------  ----         ----------
# PublicFolderMigration Stopped PublicFolder 2