#  Copyright (c) Microsoft Corporation.  All rights reserved.
#  
# THIS SAMPLE CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND,
# WHETHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
# IF THIS CODE AND INFORMATION IS MODIFIED, THE ENTIRE RISK OF USE OR RESULTS IN
# CONNECTION WITH THE USE OF THIS CODE AND INFORMATION REMAINS WITH THE USER.



# Get-PSSnapin -Registered | Add-PSSnapin -Passthru

# Get all registered snap-ins
$allSnapins = Get-PSSnapin -Registered

# Filter out only the Exchange related snap-ins
$exchangeSnapins = $allSnapins | Where-Object { $_.Name -like "Microsoft.Exchange.*" }

# Add each Exchange snap-in
foreach ($snapin in $exchangeSnapins) {
    Add-PSSnapin -Name $snapin.Name
}