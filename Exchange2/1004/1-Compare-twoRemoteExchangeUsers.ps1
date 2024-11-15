function Compare-ExchangeUserProperties {
    param (
        [string]$User1,
        [string]$User2
    )

    # Import the Exchange module (if required)
    # This might vary depending on your Exchange setup
    Import-Module ExchangeOnlineManagement

    # Retrieve properties of both users from Exchange
    $properties1 = Get-RemoteMailbox -Identity $User1 -Properties * | Select-Object -Property * -ExcludeProperty PropertyNames, PropertyCount, PropertyLength
    $properties2 = Get-RemoteMailbox -Identity $User2 -Properties * | Select-Object -Property * -ExcludeProperty PropertyNames, PropertyCount, PropertyLength

    # Create an empty array to hold the comparison results
    $comparisonResults = @()

    # Loop through each property in the first user
    foreach ($property in $properties1.PSObject.Properties) {
        # Create a custom object for each property comparison
        $comparisonResult = New-Object PSObject -Property @{
            Property    = $property.Name
            User1_Value = $property.Value
            User2_Value = $properties2.PSObject.Properties | Where-Object { $_.Name -eq $property.Name } | Select-Object -ExpandProperty Value
        }

        # Add the comparison result to the array
        $comparisonResults += $comparisonResult
    }

    # Output the results to a grid view
    $comparisonResults | Out-GridView -Title "Comparison of $User1 and $User2"
}

# Call the function with the specified users
Compare-ExchangeUserProperties -User1 "lcoates" -User2 "vpntest"
