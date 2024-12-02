# Import the Active Directory module
Import-Module ActiveDirectory

# Function to convert hex string to byte array
function ConvertFrom-HexString {
    param([string]$HexString)
    
    # Remove all spaces
    $HexString = $HexString -replace '\s',''
    
    Write-Host "Processing hex string of length: $($HexString.Length)" -ForegroundColor Yellow
    
    $bytes = New-Object byte[] ($HexString.Length / 2)
    
    for ($i = 0; $i -lt $HexString.Length; $i += 2) {
        try {
            $bytes[$i/2] = [Convert]::ToByte($HexString.Substring($i, 2), 16)
        }
        catch {
            Write-Host "Error at position $i with value: $($HexString.Substring($i, 2))" -ForegroundColor Red
            throw
        }
    }
    
    return $bytes
}

# The DN of the user you want to modify
$userDN = "CN=James Sandy,OU=Users,OU=Communications,OU=1-Departmental-Units,DC=ri,DC=nti,DC=local"

# The three hex values
$value1 = "00 02 00 00 20 00 01 47 2B DD 6E CF C3 12 51 91 E3 DC 1D 48 0F 51 7D B5 03 85 F9 FB 14 69 0F DE 27 23 73 39 BC 9C 99 20 00 02 C6 B0 DF 79 07 20 A1 C6 86 07 6B B4 14 CF B4 CA 45 D9 13 6A C1 42 59 4B A6 73 E8 7F 86 5A B7 77 1B 01 03 52 53 41 31 00 08 00 00 03 00 00 00 00 01 00 00 00 00 00 00 00 00 00 00 01 00 01 96 60 ED B4 EF A2 66 D3 50 16 ED 29 0B 02 CA 67 42 28 BC 68 93 AD DF 31 6A 34 B1 E7 04 39 64 5C 42 9D BE 71 7B 6A 52 41 03 C4 98 E3 40 04 7D 94 BF 03 1B CC A4 5C 26 90 CD 33 9D 79 04 9E 29 46 08 4A 4D 17 DF E0 6D 35 D4 9B 46 14 B1 42 C5 4E 0D 17 2D 59 E9 F1 65 87 55 51 96 57 60 C7 BF 5E B3 06 28 9E E6 E9 C3 EC 98 A0 51 CA FF 95 2E 07 5F 03 3C F9 49 E6 D7 68 CE FC A5 EB 76 80 3D AC 3B DC 9A B0 9B 8F B4 75 1C D3 0F A4 B6 F7 78 13 F2 3B E6 C3 0D D3 62 37 B8 74 66 B8 C7 D1 F1 6D 48 53 9A A3 63 60 72 99 91 A8 C5 B5 9B 28 3B 6E EF 72 CF 87 8D 82 3E 9F D8 DB 73 FA A2 AC 34 B8 95 8D 95 CE D2 27 A2 63 CF 17 72 D4 F2 67 97 9D 24 24 BD 35 B6 48 9F 4C C4 85 77 1D 43 01 5C 1F 96 02 03 EA 3F 42 22 FD 0D 9B CF A3 3F 90 13 61 4A B7 53 55 04 B3 94 6E 13 83 CD 38 80 F5 58 37 01 00 04 01 01 00 05 01 10 00 06 D7 32 BE 3A A9 1B 98 4C AD 89 97 08 7E 86 99 E5 0F 00 07 01 00 00 00 00 02 01 00 00 00 00 00 00 00 00 08 00 08 00 00 00 00 00 00 00 00 08 00 09 7F 6D F2 1B B4 ED DA 48"

$value2 = "00 02 00 00 20 00 01 67 6E CA EF 8E 1D BF C3 4C FA E2 B2 EE 30 5B 3E BC 52 AA C8 F3 0B 1F F5 DB FC 3A 73 16 CE 54 E1 20 00 02 B2 85 DA B2 0A 0F C6 5D D4 54 C4 AF 64 4C 97 49 DA 05 6E 83 F4 C4 CC 19 BD 50 10 03 8A 09 48 46 1B 01 03 52 53 41 31 00 08 00 00 03 00 00 00 00 01 00 00 00 00 00 00 00 00 00 00 01 00 01 A2 E0 11 1F 2A D6 28 09 39 8C 94 C8 7F 7A 37 9E EA 9F AE 86 FF 66 03 5E B3 D5 E8 CE 08 46 40 D7 2B 75 07 2B 7C 0F 81 66 D1 DC A1 4D 45 67 84 F7 D6 09 EA 68 6C 4E 59 F1 BA 35 5C EC 1B F6 84 2C FD 5D 34 7A B6 FC EC 8F C6 C0 A3 81 76 5F 91 E2 F1 7E 7B 67 C6 6B 8C DD 4C F3 BF C0 7D 02 FC 2A 9D 34 96 79 9D 3A 07 CA 54 63 C3 89 8D 7F 52 F6 BA D6 01 58 C7 27 8B ED 14 D6 48 7F 8E 31 D1 7E BA AD 75 76 59 1C 5F 2B 53 63 39 71 CE 07 D5 4D 7F 36 8E 4E 4B ED 13 FC 44 C9 73 85 94 E0 8C 1F 47 5C C4 91 DC D9 8D 8A 2D FC D9 37 4D 91 DF 02 FF 12 D5 C6 CF 2D 03 70 E3 B6 A4 24 75 D1 AB 0E D8 12 B9 D5 C0 FF 92 8B DC A6 02 30 E0 37 05 95 01 12 A1 B2 E9 BE EB 9F 86 88 17 A2 EC 24 3D 2D A9 24 A7 E8 43 53 8E FF 00 93 91 91 4A 47 7A 33 26 17 08 A9 37 DC 9C 7F F7 3F F1 94 C0 B1 DC 3D 01 00 04 01 01 00 05 01 10 00 06 5B 7F BE 04 68 C0 4F 4E AE 40 CE EA C4 8F F7 C4 0F 00 07 01 00 00 00 00 02 01 00 00 00 00 00 00 00 00 08 00 08 00 00 00 00 00 00 00 00 08 00 09 96 15 6B 29 E7 BE DB 48"

$value3 = "00 02 00 00 20 00 01 6C 40 43 21 E5 F5 36 1C C3 68 D3 CA 64 8C 29 98 FA 43 FF 44 07 CE 5E DA 41 E9 60 A7 66 50 41 7B 20 00 02 9A 5B 31 0A BD 46 B6 06 C9 42 ED 46 C5 C4 A0 59 0E 81 8F A4 F5 AB 38 A7 0B 2F 68 96 28 3B DF 42 1B 01 03 52 53 41 31 00 08 00 00 03 00 00 00 00 01 00 00 00 00 00 00 00 00 00 00 01 00 01 97 CB 6D F9 94 1C 33 86 6E 7B 21 C1 F6 F6 95 10 20 68 54 14 53 5D 07 D6 EB 29 E7 21 FF 22 0F F0 D3 25 E8 A7 0E 4C AF 6A 88 08 67 E3 31 B7 1E 47 D9 CF 5F 2E 43 0F EB B9 C9 8B 6D 47 A3 89 35 65 45 35 98 F6 28 D2 F7 BF CE 70 AB A8 CB 11 36 86 66 11 C1 86 D5 4D D1 B0 9E B3 57 57 6B AE 73 0A B2 35 96 EB 41 69 08 E4 9B AC 3B D9 95 B2 47 24 28 D4 B7 D4 B7 8B 2E 3C B6 9A 0A C0 08 68 0E BD C6 66 A6 F8 B8 98 8E D4 74 70 B6 24 5A 24 11 E8 35 14 7B D0 F2 4C 4D BD 08 3B FF DA 17 7B 6D E4 4E FA EA 3A 21 31 67 BB 5D 9A 5F 41 80 CD 2A E9 3F 19 BE 03 48 BB BD 35 E0 59 44 93 7B 74 44 73 6C FF 55 3D 13 38 8C 1F 0F 9C 1F D4 3C 62 B6 24 7A D3 05 4F FE 1C A0 B1 14 4A C3 DB 19 30 02 6B 02 94 48 A5 1D AE 4F 41 A6 0D BE 14 59 9F 82 9D 84 F9 59 29 B5 EF 29 3B 8F 09 D6 FB E1 BA C2 61 01 00 04 01 01 00 05 01 10 00 06 05 E7 E1 56 95 B1 BA 46 90 32 3A 3C 71 7C 33 7E 0F 00 07 01 00 00 00 00 02 01 00 00 00 00 00 00 00 00 08 00 08 00 00 00 00 00 00 00 00 08 00 09 4F 7E 9B 0F 18 C6 DC 48"

Write-Host "`nConverting values to byte arrays..." -ForegroundColor Cyan

try {
    # Convert hex strings to byte arrays
    $bytes1 = ConvertFrom-HexString $value1
    $bytes2 = ConvertFrom-HexString $value2
    $bytes3 = ConvertFrom-HexString $value3

    Write-Host "`nAttempting to update AD..." -ForegroundColor Cyan

    # Create the attribute values as DirectoryAttributeModification
    $modification = New-Object System.DirectoryServices.DirectoryAttributeModification
    $modification.Operation = [System.DirectoryServices.DirectoryAttributeOperation]::Replace
    $modification.Name = "msDS-KeyCredentialLink"
    [void]$modification.Add($bytes1)
    [void]$modification.Add($bytes2)
    [void]$modification.Add($bytes3)

    # Get the DirectoryEntry for the user
    $searcher = New-Object System.DirectoryServices.DirectorySearcher
    $searcher.Filter = "(distinguishedName=$userDN)"
    $result = $searcher.FindOne()
    $userEntry = $result.GetDirectoryEntry()

    # Apply the modification
    $userEntry.ModifyAttributes(@($modification))
    $userEntry.CommitChanges()
    
    Write-Host "Successfully updated msDS-KeyCredentialLink attribute" -ForegroundColor Green

    # Verify the update
    $user = Get-ADUser -Identity $userDN -Properties msDS-KeyCredentialLink
    Write-Host "Current number of msDS-KeyCredentialLink values: $($user.'msDS-KeyCredentialLink'.Count)"
}
catch {
    Write-Error "Error occurred: $_"
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
}