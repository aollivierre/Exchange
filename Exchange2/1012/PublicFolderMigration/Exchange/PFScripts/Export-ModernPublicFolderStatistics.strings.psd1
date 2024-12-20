ConvertFrom-StringData @'
###PSLOC
ProcessingNonIpmSubtree = Enumerating folders under NON_IPM_SUBTREE...
ProcessingNonIpmSubtreeComplete = Enumerating folders under NON_IPM_SUBTREE completed...{0} folders found.
ProcessingIpmSubtree = Enumerating folders under IPM_SUBTREE...
ProcessingIpmSubtreeComplete = Enumerating folders under IPM_SUBTREE completed...{0} folders found.
ExportToCSV = Exporting folders to a CSV file
RetrievingStatistics = Retrieving statistics...
RetrievingStatisticsComplete = Retrieving statistics complete...{0} folders found.
UniqueFoldersFound = Total unique folders found : {0}.
ProcessedFolders = Folders processed : {0}.
ExportStatistics = Exporting statistics for {0} folders
VersionErrorMessage = This script should be run on Exchange Server 2013 CU15 or later, or Exchange Server 2016 CU4 or later. The following servers are running other versions of Exchange Server:\n\t{0}
ProgressBarActivity = Generating Statistics for Public Folders...
InvalidFolderNames = The following folders have invalid characters in the name ('\\' or '/') and cannot be migrated; please rename them and run the script again:\n\t{0}
InvalidExportFile = Path to the export file '{0}' is invalid. Please provide a valid path.
###PSLOC
'@