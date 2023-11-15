# CreateRemoteSharedMailbox

## Synopsis
Creates a remote shared mailbox in Exchange.

## Description

This PowerShell function, `CreateRemoteSharedMailbox`, is designed to:

1. Prompt the user for input, including the name, first name, last name, user principal name, domain, and organizational unit (OU) name.
2. Automatically determine the OU distinguished name from the provided OU name.
3. Check if the Exchange cmdlets are loaded and if the `New-RemoteMailbox` cmdlet is available.
4. Output initial statistics, including the number of users in the OU and the total number of mailboxes, remote mailboxes, and remote shared mailboxes.
5. Create the remote shared mailbox using the `New-RemoteMailbox` cmdlet.
6. Check if the user and mailbox were created by retrying a few times and waiting 10 seconds between retries.
7. If successful, output the user's GUID and final statistics.
8. If the creation was not successful, output an error message.

## Usage

```powershell
# Call the function
CreateRemoteSharedMailbox