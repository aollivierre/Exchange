#just open AAD Connect and check if the Exchange Mail Public Folders option is NOT selected


# If you're using Microsoft Azure Active Directory Connect (Azure AD Connect) to synchronize your on-premises directories with Azure Active Directory, you need to do the following (if you aren't using Azure AD Connect, you can skip this step):

# On an on-premises computer, open Microsoft Azure Active Directory Connect, and then select Configure.

# On the Additional tasks screen, select Customize synchronization options, and then click Next.

# On the Connect to Azure AD screen, enter the appropriate credentials, and then click Next. Once connected, keep clicking Next until you're on the Optional Features screen.

# Make sure that Exchange Mail Public Folders is not selected. If it isn't selected, you can continue to the next section, Prerequisite steps in Exchange Online. If it is selected, click to clear the check box, and then click Next.

#  Note

# If you don't see Exchange Mail Public Folders as an option on the Optional Features screen, you can exit Microsoft Azure Active Directory Connect and proceed to the next section, Prerequisite steps in Exchange Online.

# After you have cleared the Exchange Mail Public Folders selection, keep clicking Next until you're on the Ready to configure screen, and then click Configure.