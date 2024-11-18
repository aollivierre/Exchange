Contents

[RBAC Roles and Permissions Required 2](#_Toc182559512)

[Best Practices for Administrative Accounts 5](#_Toc182559513)

[Enabling Exchange Hybrid Writeback in Entra ID Connect Sync 7](#_Toc182559514)

[Option 1: Using the Exchange Admin Center 10](#_Toc182559515)

[Offboarding Users with Exchange Admin Center 13](#_Toc182559516)

[Option 2: Using the Exchange Management Shell 21](#_Toc182559517)

**IT Guide: Managing Exchange Attributes and Remote User Mailboxes**

# RBAC Roles and Permissions Required

**RBAC Roles and Permissions Required**

To effectively manage Exchange attributes and remote user mailboxes in a hybrid Exchange environment, specific Role-Based Access Control (RBAC) roles and permissions are required across your on-premises Active Directory (AD), Entra ID (formerly Azure AD), and synchronization tools like Entra ID Connect Sync. This section outlines the necessary roles and best practices for maintaining secure administrative access.

**On-Premises Active Directory and Exchange Server**

**Active Directory Permissions:**

- **Account Operators** or **Delegated User Management Permissions**:
  - Required to create, modify, and delete user accounts in Active Directory.
  - Allows for setting passwords and managing user properties.

**Exchange Server Permissions:**

- **Recipient Management Role Group**:
  - Grants permissions to manage mailboxes, mail users, mail contacts, and distribution groups.
  - Essential for managing recipients in both on-premises and hybrid environments.
- **Organization Management Role Group**:
  - Provides full administrative access to all Exchange features.
  - Should be assigned sparingly due to its broad scope.

**Assigning Roles in Exchange On-Premises:**

- Use the **Exchange Admin Center** or **Exchange Management Shell** to assign administrators to the appropriate role groups.

**Example using Exchange Management Shell:**

Add-RoleGroupMember -Identity "Recipient Management" -Member "AdminUser"

**Entra ID (Formerly Azure Active Directory) and Exchange Online**

**Entra ID Permissions:**

- **User Administrator**:
  - Can manage all aspects of users and groups, including resetting passwords and assigning licenses.
  - Ideal for administrators who handle user account management.
- **Exchange Administrator**:
  - Grants permissions to manage Exchange Online recipients, mailbox settings, and compliance features.
  - Necessary for managing mailboxes and recipient objects in Exchange Online.

**Assigning Roles in Entra ID:**

- Use the **Entra ID Admin Center** or **Microsoft 365 Admin Center** to assign roles.

**Steps:**

- 1. Navigate to the **Entra ID Admin Center** (<https://entra.microsoft.com>).
  2. Select **Roles and administrators**.
  3. Choose the appropriate role (e.g., **User Administrator** or **Exchange Administrator**).
  4. Click **Add assignments** and select the user to assign the role.

**Exchange Online RBAC Roles:**

- **Recipient Management Role Group**:
  - Similar to the on-premises role group but applies to Exchange Online.
  - Allows administrators to manage Exchange Online recipients.

**Assigning Roles in Exchange Online:**

- Use the **Exchange Admin Center** for Exchange Online:

**Steps:**

- 1. Navigate to the **Exchange Admin Center** (<https://admin.exchange.microsoft.com>).
  2. Click on **Permissions** in the left-hand menu.
  3. Under **Admin Roles**, select **Recipient Management**.
  4. Click **Edit** to add members to the role group.

**Entra ID Connect Sync (Formerly Azure AD Connect Sync)**

**Permissions for Entra ID Connect Sync:**

- **On-Premises Permissions:**
  - **Enterprise Admin** permissions are required during the initial setup. (will be used to create MSOL_XXXXXX sync service account in AD.
  - For ongoing synchronization, a service account with **Replicating Directory Changes** permissions is sufficient.
- **Entra ID Permissions:**
  - The account used for Entra ID Connect Sync should have the **Hybrid Identity Administrator** role.
  - This role allows for configuring synchronization settings and managing hybrid identities.

**Best Practices for Entra ID Connect Sync:**

- **Use Dedicated Service Accounts:**
  - Create dedicated service accounts for synchronization processes.
  - Assign only the minimum required permissions to these accounts.
- **Monitor Synchronization:**
  - Regularly check synchronization logs and health reports to ensure that directory synchronization is functioning correctly.

# Best Practices for Administrative Accounts

- **Separate On-Premises and Cloud-Only Admin Accounts:**
  - **Avoid Synchronizing Admin Accounts to Entra ID:**
    - To enhance security, do not synchronize on-premises administrative accounts to Entra ID.
    - Synchronizing admin accounts can expose them to additional risks if cloud credentials are compromised.
  - **Use Cloud-Only Admin Accounts in Entra ID:**
    - Create separate, cloud-only administrative accounts in Entra ID for managing cloud services.
    - These accounts are not linked to on-premises identities and provide an additional layer of security.
- **Implement Multi-Factor Authentication (MFA):**
  - Enable MFA for all administrative accounts to protect against unauthorized access.
    - Option 1 – via Entra ID Conditional Access Policies (Recommended)
    - Option 2 - via Entra ID Legacy Per-user MFA (less recommended)
    - Option 3 – via Entra ID Security Defaults (less recommended)
  - MFA adds a critical layer of security by requiring a second form of verification.
- **Principle of Least Privilege:**
  - Assign the minimal permissions necessary for administrators to perform their tasks.
  - Regularly review and adjust permissions to ensure compliance with security policies.
- **Use Privileged Access Workstations (PAWs):**
  - Administer sensitive systems from dedicated, secure workstations.
  - PAWs reduce the risk of credential theft and unauthorized access.

**Important Considerations**

- **Role Assignments are Environment-Specific:**
  - Roles assigned in on-premises environments do not automatically carry over to Entra ID or Exchange Online.
  - Ensure that administrators have the necessary roles in each environment they manage.
- **Compliance and Audit Trail:**
  - Maintain logs of role assignments and changes to administrative permissions.
  - Regular audits help in detecting unauthorized changes and ensuring compliance with organizational policies.
- **Training and Awareness:**
  - Provide administrators with training on security best practices and the proper use of their assigned roles.
  - Awareness reduces the risk of accidental misconfigurations or security breaches.

**Summary of Required Roles**

| **Task** | **On-Premises Role** | **Entra ID Role** |
| --- | --- | --- |
| Create user accounts | Account Operators | User Administrator |
| Manage mailboxes and recipients | Recipient Management Role Group | Exchange Administrator |
| Assign or remove licenses | N/A | User Administrator |
| Configure Exchange settings | Exchange Administrator Role Group | Exchange Administrator |
| Manage directory synchronization | N/A | Hybrid Identity Administrator |

# Enabling Exchange Hybrid Writeback in Entra ID Connect Sync

In a hybrid Exchange environment, it's recommended to enable **Exchange Hybrid writeback** in **Entra ID Connect Sync** (formerly Azure AD Connect Sync). This feature ensures that your on-premises Active Directory (AD) remains the source of authority by synchronizing certain attributes from Exchange Online back to your on-premises AD. This synchronization maintains consistency of user and mailbox attributes when changes are made to Exchange Online mailboxes.


> 🔔 **Note:** If you have Azure AD Premium P1 (Entra ID P1) licenses, enabling all write-back features in Azure AD Connect Sync (traditional sync produc) towards the very end of the configuratrion wizard is a smart move in most hybrid environments.


**Benefits of Enabling Exchange Hybrid Writeback**

- **Attribute Consistency:** Ensures that changes made to mailboxes in Exchange Online are reflected in your on-premises AD. This includes attributes like proxy addresses and mail routing settings, which are crucial for proper email flow and user management.
- **Source of Authority Maintenance:** Keeps your on-premises AD as the authoritative source for directory information, which is essential for hybrid deployments.
- **Simplified Management:** Reduces administrative overhead by automatically synchronizing necessary attributes, eliminating the need for manual updates in both environments.
- **Enhanced Hybrid Functionality:** Supports features like on-premises distribution group management and allows for seamless coexistence between on-premises and cloud environments.

**How to Enable Exchange Hybrid Writeback**

**Prerequisites:**

- Administrative permissions to modify Entra ID Connect Sync configuration.
- On-premises Exchange schema is updated to support hybrid writeback features.
- Ensure that Entra ID Connect Sync is running on a supported Windows Server version.

**Steps:**

1. **Open Entra ID Connect Sync Configuration:**
    - On the server where Entra ID Connect Sync is installed, open **Azure AD Connect** from the Start menu.
    - If prompted, provide administrative credentials to run the configuration wizard.
2. **Choose to Customize Synchronization Options:**
    - In the **Welcome to Azure AD Connect** screen, select **Configure**.
    - Choose **Customize synchronization options** and click **Next**.
3. **Connect to Entra ID:**
    - Enter your Entra ID (Azure AD) global administrator credentials when prompted.
    - Click **Next** to proceed.
4. **Connect to Active Directory Domain Services:**
    - Ensure your on-premises AD forest is listed and the status is **Verified**.
    - Click **Next**.
5. **Select Directory Sync and Optional Features:**
    - In the **Optional Features** screen, check the box for **Exchange hybrid writeback**.
    - **Note:** If this option is greyed out, ensure that your AD schema supports this feature and that you're running a compatible version of Entra ID Connect Sync.
6. **Complete the Configuration:**
    - Proceed through the remaining steps, accepting defaults unless changes are required.
    - Click **Configure** to apply the new settings.
    - Wait for the configuration to complete and click **Exit**.
7. **Verify the Configuration:**
    - Open **Synchronization Service Manager** from the Start menu.
    - Confirm that synchronization runs are completing successfully.
    - Check the **Connectors** tab to ensure that the writeback configuration is active.

**Important Considerations**

- **Schema Updates:**
  - Enabling Exchange Hybrid writeback may require extending your on-premises AD schema if not already done.
  - Ensure that schema updates are performed carefully and during maintenance windows to minimize impact.
- **Permissions:**
  - The account used by Entra ID Connect Sync must have the necessary permissions to write back attributes to on-premises AD.
  - Typically, the **MSOL** account used by Entra ID Connect Sync is granted these permissions during setup.
- **Attribute Scope:**
  - Exchange Hybrid writeback synchronizes specific attributes such as proxy addresses, delegation settings, and certain mailbox properties.
  - It does **not** write back passwords or other sensitive information.
- **Continuous Synchronization:**
  - Keep Entra ID Connect Sync running continuously to maintain synchronization between your on-premises AD and Entra ID.
  - Regular synchronization ensures that any changes in Exchange Online are promptly reflected on-premises.
- **Monitoring and Maintenance:**
  - Regularly monitor synchronization logs and health metrics.
  - Address any synchronization errors promptly to maintain data consistency.
- **Security and Compliance:**
  - Review organizational policies to ensure that attribute writeback aligns with security and compliance requirements.
  - Ensure that proper auditing is in place to track changes to directory objects.

**Why Exchange Hybrid Writeback is Recommended**

- **Maintains Unified Global Address List (GAL):**
  - Ensures that on-premises users have up-to-date information about cloud mailboxes and vice versa.
- **Simplifies Mail Flow Configuration:**
  - Consistent attributes across environments reduce complexity in mail routing and delivery.
- **Supports Cross-Premises Features:**
  - Enables features like cross-premises mailbox permissions and delegate access.
- **Facilitates Future Migrations:**
  - Keeps the option open for seamless transitions, whether moving more mailboxes to the cloud or back on-premises.

# Option 1: Using the Exchange Admin Center

This section provides step-by-step instructions for onboarding and offboarding users using the **Exchange Admin Center (EAC)** in a hybrid Exchange environment.

**Onboarding Users with Exchange Admin Center**

**Prerequisites:**

- Administrative permissions on both on-premises Exchange and Exchange Online.
- Azure AD Connect sync (traditional sync) is configured and synchronizing accounts properly.
  - The Organizational Unit (OU) under which the underlying AD user object will be created in must be in the sync scope of the Azure AD Connect Sync Scope filters (selected during Azure AD/Entra ID Connect Sync wizard)
- Appropriate licenses available in Microsoft 365.
  - A minimum of Exchange Online P1 license or a SKU that includes Exchange Online P1 license

**Steps:**

1. **Access the Exchange Admin Center:**
    - Open a web browser and navigate to the EAC URL for your on-premises Exchange server (e.g., https://&lt;Your-Exchange-Server&gt;/ecp).
    - Log in with your administrative credentials (like a member of Organization Management role group)
2. **Navigate to Recipients:**
    - In the EAC, click on **Recipients** in the left-hand menu.
    - Select the **Mailboxes** tab.
3. **Create a New Remote Mailbox:**
    - Click the **➕** (Add) button and select **Office 365 Mailbox**.
4. ![image](https://github.com/user-attachments/assets/2282ce96-4340-4eb0-8637-82347a56c32e)



    - - _Note:_ In a hybrid environment, creating a remote mailbox ensures the mailbox is created directly in Exchange Online.
5. **Fill in User Information:**
    - **First name**, **Last name**, **Display name**: Enter the user's details.
    - **User logon name (User Principal Name)**: Specify the user's logon name and select the appropriate domain.
    - **Password**: Assign a temporary password and confirm it.
6. **Specify Organizational Unit (Optional):**
    - If necessary, click **Browse** to select the Organizational Unit (OU) where the user account will be created in Active Directory.
7. **Configure Mailbox Settings:**
    - **Archive mailbox**: Enable if the user requires an online archive mailbox.
    - **Retention policy**: Assign if applicable.
    - **Address book policy**: Assign if applicable.
8. **Save the New User:**
    - Review the information and click **Save** to create the user account and remote mailbox.
9. **Synchronize with Azure AD:**
    - Wait for Azure AD Connect to synchronize the new account to Azure AD.
        - You can force a sync if necessary, using PowerShell via the following command:
            - Start-ADSyncSyncCycle - Policytype Initial
10. **Assign a License in Microsoft 365:**
    - Log in to the **Microsoft 365 Admin Center** (<https://admin.microsoft.com>).
    - Navigate to **Users** > **Active users**.
    - Locate and select the new user.
    - Click **Licenses and Apps**, assign the appropriate license (e.g., Exchange Online Plan), and click **Save changes**.
11. **Verify Mailbox Creation:**
    - Confirm that the mailbox is active in Exchange Online.
    - Send a test email to ensure mail flow is working correctly.

# Offboarding Users with Exchange Admin Center

In a hybrid Exchange environment, offboarding a user requires careful handling to ensure data retention and compliance with company policies. The following steps guide you through the process of offboarding a user while preserving their mailbox data and managing their OneDrive content.

**Prerequisites:**

- Administrative permissions on both on-premises Exchange and Exchange Online.
  - Follow the same RBAC requirements mentioned in this document.
- Understanding of company policies regarding data retention and user account deletion.
- Ensure that a cloud backup solution (e.g., Datto SaaS Protection or another cloud-to-cloud backup service) is in place to back up mailbox and OneDrive data before proceeding.

**Steps:**

1. **Access the Exchange Admin Center (EAC):**
    - Open a web browser and navigate to the **Exchange Admin Center** for Exchange Online: <https://admin.exchange.microsoft.com>.
    - Log in with your administrative credentials.
2. **Navigate to the User's Mailbox:**
    - In the EAC, click on **Recipients** in the left-hand menu.
    - Select the **Mailboxes** tab.
    - Search for and select the mailbox of the user you need to offboard.
3. **Convert the Mailbox to a Shared Mailbox:**
    - With the user's mailbox selected, click on **Convert to shared mailbox** in the toolbar.

![image](https://github.com/user-attachments/assets/a23008c9-74a5-4fda-91ff-0e5b4b4314a9)


    - - **Note:** Converting to a shared mailbox allows you to retain access to the user's email data without requiring an active license.
    - Confirm the conversion when prompted.
        - The mailbox is now a shared mailbox, and permissions can be assigned to other users if necessary.
5. **Remove the License in Microsoft 365:**
    - Navigate to the **Microsoft 365 Admin Center**: <https://admin.microsoft.com>.
    - Go to **Users** > **Active users**.
    - Select the user you are offboarding.
    - Click on **Licenses and Apps**.
    - Uncheck the licenses assigned to the user, especially the **Exchange Online** license.
    - Click **Save changes**.
        - **Note:** Removing the license after converting to a shared mailbox ensures that mailbox data is retained without incurring license costs.
6. **Manage OneDrive Data:**
    - While still in the **Microsoft 365 Admin Center**, select the user and click on **OneDrive**.
    - To retain the user's OneDrive data:
        - **Create an access link**:
            - Click **Create link to files** to generate a direct link to the user's OneDrive files.
        - **Assign a Manager for OneDrive**:
            - Under OneDrive Settings > More Settings > Site Collection Administrators
![image](https://github.com/user-attachments/assets/c636ec37-e81e-49b6-9ea7-edf090031a21)

![image](https://github.com/user-attachments/assets/ac78f593-c1df-4d70-9554-e2911b23ea3c)


- - **Configure OneDrive Retention Settings**:
        - Go to the **SharePoint Admin Center**.
        - Under **Retention**, set the retention period for OneDrive data (up to a maximum of 3,650 days).
        - **Note:** Proper retention settings ensure compliance with data preservation policies.

1. **Backup Data (Recommended):**
    - Before proceeding further, use your cloud backup solution to back up the user's mailbox and OneDrive data.
        - This provides an extra layer of protection against data loss.
        - Solutions like **Datto SaaS Protection** can automate this process.
2. **Disable the User Account in Active Directory:**
    - Open **Active Directory Users and Computers** on your on-premises server.
    - Locate the user account.
    - Right-click on the user and select **Disable Account**.
        - **Note:** Disabling the account prevents the user from accessing on-premises resources.
    - Ensure that **Entra ID Connect Sync** (formerly Azure AD Connect Sync) synchronizes this change to Entra ID.
        - You can force a synchronization using the following PowerShell command on the server running Entra ID Connect Sync:

Start-ADSyncSyncCycle -PolicyType Delta

1. **Block Sign-In (for Cloud-only users)**
    - In the **Microsoft 365 Admin Center**, navigate to **Users** > **Active users**.
    - Select the user account.
    - Toggle **Block sign-in** to **On**.
        - This prevents the user from signing in to Microsoft 365 services.
2. **Revoke authentication tokens:** (for both Hybrid and Cloud-only users)
    - - Go to the **Entra ID Admin Center** (<https://entra.microsoft.com>).
        - Navigate to **Users** > **All users**.
        - Select the user, then click on **Authentication methods**.
        - Click **Require re-register multifactor authentication**.
        - Additionally, you can revoke user sessions by clicking **Revoke sessions** under **Devices**.
3. **(Optional) Remove the User from Entra ID Connect Sync Scope:**
    - If you prefer to remove the user entirely from Entra ID:
        - Exclude the user from the synchronization scope in **Entra ID Connect Sync** configuration.
        - Wait for the synchronization cycle to process the change, or force a sync.
        - The user account will be deleted in Entra ID and moved to the **Deleted users** section.
        - **Restore the User in Entra ID**:
            - Navigate to **Entra ID Admin Center** > **Users** > **Deleted users**.
            - Select the user and click **Restore**.
            - The restored user becomes a cloud-only account.
        - **Run the User Deletion Wizard**:
            - In the **Microsoft 365 Admin Center**, delete the user using the **Delete user** option.
            - This process converts the mailbox to a shared mailbox automatically and allows you to assign OneDrive data to another user.
        - **Note:** This alternative method automates several steps but may not be suitable for all organizations.
4. **Verify Offboarding Completion:**
    - **Mailboxes**:
        - Confirm that the user's mailbox is now a shared mailbox in Exchange Online.
        - Ensure that other users who need access have the appropriate permissions.
    - **Licenses**:
        - Verify that all licenses have been unassigned from the user account.
    - **User Account**:
        - Ensure the user account is disabled in Active Directory and blocked in Entra ID.
    - **Data Access**:
        - Check that OneDrive data has been transferred or is accessible as per your organization's policies.
    - **Security**:
        - Confirm that all authentication sessions have been revoked and that the user cannot access any company resources.

**Important Considerations:**

- **Data Retention Compliance:**
  - Always adhere to your organization's data retention policies.
  - Converting mailboxes to shared mailboxes retains email data without needing a license.
  - OneDrive data retention must be managed separately through SharePoint Online settings.
- **Cloud Backup Solutions:**
  - Utilizing cloud-to-cloud backup solutions like **Datto SaaS Protection** ensures that you have a secure backup of all user data before making any changes.
  - This is especially important if legal holds or compliance requirements mandate data preservation.
- **Disabling vs. Deleting Accounts:**
  - Disabling the user account in Active Directory prevents access but retains the account for any necessary data retrieval.
  - Deleting the account is irreversible and should only be done if you're certain that all data is no longer needed.
- **Disconnecting Mailbox in Exchange On-Premises:**
  - In a hybrid environment, it's generally not necessary to disable or disconnect the mailbox in on-premises Exchange after converting it to a shared mailbox in Exchange Online.
  - The shared mailbox resides in Exchange Online, and the on-premises user account can remain disabled.
- **Security Measures:**
  - Blocking sign-in and revoking authentication tokens prevent the user from accessing any cloud services.
  - Always ensure these steps are taken promptly to secure the environment.
- **Alternative Offboarding Method:**
  - The alternative method of removing the user from the sync scope and restoring them as a cloud-only user may simplify the process but requires careful handling to avoid unintended data loss.
  - Evaluate this option based on your organization's policies and the specific circumstances.

# Option 2: Using the Exchange Management Shell

This section provides step-by-step instructions for onboarding and offboarding users using the **Exchange Management Shell (EMS)** in a hybrid Exchange environment.

**Onboarding Users with Exchange Management Shell**

**Prerequisites:**

- Administrative permissions on the on-premises Exchange server.
  - Follow the RBAC requirements mentioned in the RBAC section.
- Exchange Management Shell is installed and accessible.
- **Entra ID Connect Sync** (formerly Azure AD Connect) is configured and synchronizing accounts properly.
- Appropriate licenses available in Microsoft 365.
- Ensure that PowerShell execution policies allow for script execution if running scripts.

**Steps:**

1. **Open the Exchange Management Shell:**
    - Log in to your on-premises Exchange server or a workstation with the Exchange management tools installed.
    - Open the **Exchange Management Shell** as an administrator.
2. **Create a New Remote Mailbox:**
    - Use the New-RemoteMailbox cmdlet to create a new Active Directory user and associated remote mailbox in Exchange Online.
    - **Example Command:**

$Password = Read-Host "Enter temporary password" -AsSecureString

New-RemoteMailbox -Name "John Doe" -FirstName "John" -LastName "Doe" -UserPrincipalName "<john.doe@yourdomain.com>" -Password $Password

You can instead use this script https://github.com/aollivierre/Exchange/blob/main/Exchange2/1013/Exchange/New-RemoteSharedMailbox-v10%20copy.ps1 to automate the entire process. Please make sure to first update the values inside of the script to meet your needs before running it.

- - - **Note:**
            - Replace "John Doe", "John", "Doe", and "<john.doe@yourdomain.com>" with the user's actual information.
            - The -Password parameter sets a temporary password for the new user account.

1. **Configure Additional User Properties (Optional):**
    - Set additional properties like department, title, or custom attributes if required.
    - **Example:**

Set-User -Identity "<john.doe@yourdomain.com>" -Department "Sales" -Title "Sales Representative"

1. **Verify the Remote Mailbox Creation:**
    - Confirm that the remote mailbox has been created.
    - **Example:**

Get-RemoteMailbox "<john.doe@yourdomain.com>"

1. **Force Entra ID Connect Sync Synchronization:**
    - By default, synchronization occurs every 30 minutes. To synchronize immediately:
    - **On the server running Entra ID Connect Sync:**

Start-ADSyncSyncCycle -PolicyType Delta

- - - Run this command in PowerShell with administrative privileges.

1. **Assign a License in Microsoft 365:**
    - After synchronization, log in to the **Microsoft 365 Admin Center** (<https://admin.microsoft.com>).
    - Navigate to **Users** > **Active users**.
    - Locate and select the new user.
    - Click on **Licenses and Apps**.
    - Assign the appropriate license (e.g., **Office 365 E3**, **Exchange Online Plan 1**) and click **Save changes**.
2. **Configure Exchange Online Mailbox Settings (Optional):**
    - If additional mailbox settings are required, you can configure them using **Exchange Online PowerShell**.
    - **Connect to Exchange Online PowerShell:**

Import-Module ExchangeOnlineManagement

Connect-ExchangeOnline -UserPrincipalName <admin@yourdomain.com>

- - - Replace <admin@yourdomain.com> with your administrator account.
    - **Set Mailbox Settings:**

Set-Mailbox "<john.doe@yourdomain.com>" -RetentionPolicy "Default MRM Policy" -EmailAddresses SMTP:john.doe@yourdomain.com,smtp:j.doe@yourdomain.com

1. **Inform the User:**
    - Provide the user with their account details and temporary password.
    - Instruct them to log in and change their password as soon as possible.
2. **Test the Mailbox:**
    - Verify that the mailbox is functioning correctly.
    - Send a test email to and from the mailbox.

**Notes:**

- **Security Considerations:**
  - Avoid hardcoding passwords in scripts. Use Read-Host or other secure methods to handle passwords.
  - Ensure that all administrative actions comply with your organization's security policies.
- **RBAC Roles:**
  - On-premises Exchange: Membership in the **Recipient Management** role group is required.
  - Entra ID and Exchange Online: Ensure you have the **User Administrator** and **Exchange Administrator** roles.
- **Email Address Policies:**
  - If your organization uses custom email address policies, make sure they are applied after creating the mailbox.
  - Use the Update-EmailAddressPolicy cmdlet if necessary.
- **User Principal Name (UPN):**
  - Ensure that the UPN matches the user's email address for seamless single sign-on experiences.
- **License Assignment:**
  - Users without a license cannot access their mailboxes in Exchange Online.
  - License assignment must be completed promptly after synchronization.



  Yes, you are correct. Using the **Exchange Recipient Management** tools allows IT administrators to manage Exchange recipient objects without the need to maintain an on-premises Exchange mailbox server. This approach enables you to decommission the last on-premises Exchange mailbox server while still retaining the ability to manage mailboxes and recipient attributes in a hybrid environment.

The **Exchange Recipient Management PowerShell module** provides the necessary cmdlets to perform recipient management tasks directly against your on-premises Active Directory (AD). These changes are then synchronized to Entra ID (formerly Azure AD) using Entra ID Connect Sync, ensuring that your Exchange Online environment reflects the updates.

By utilizing this method, organizations can simplify their infrastructure by removing the dependency on on-premises Exchange servers for recipient management, provided that all mailboxes have been migrated to Exchange Online.

---

Let's proceed to **Option 3**, where we'll detail the steps for onboarding and offboarding users using the **Exchange Recipient Management PowerShell module**.

# Option 3: Using the Exchange Recipient Management PowerShell Module

This section provides step-by-step instructions for onboarding and offboarding users using the **Exchange Recipient Management PowerShell module** in a hybrid Exchange environment without an on-premises Exchange mailbox server.

## Overview

- **Purpose:** Allows management of Exchange-related attributes in on-premises AD without maintaining an on-premises Exchange server.
- **Benefits:**
  - Decommission the last on-premises Exchange mailbox server.
  - Reduce infrastructure complexity and costs.
  - Maintain control over recipient management in a hybrid environment.

## Prerequisites

- **Environment Requirements:**
  - All mailboxes have been migrated to Exchange Online.
  - Entra ID Connect Sync is configured and operational.
  - Exchange schema extensions are present in on-premises AD.
- **Administrative Permissions:**
  - Permissions to modify user objects in on-premises AD.
  - Appropriate roles assigned in Entra ID and Exchange Online (as per the RBAC roles section).
- **Tools Required:**
  - **Exchange Recipient Management PowerShell module** installed on a management workstation or server.
  - **PowerShell** with the Active Directory module.
  - **Entra ID Connect Sync** (formerly Azure AD Connect Sync) to synchronize changes to Entra ID.

---

## Onboarding Users with Exchange Recipient Management PowerShell Module/Snap-In

**Steps:**

1. **Install the Exchange Recipient Management Tools:**

   - **Download the Exchange Management Tools:**
     - Obtain the latest version of the **Exchange Server 2019 Setup** files from Microsoft's official site.
    
     - > 🔔 **Note:** You don't need to install Exchange Server Mailbox Server role during Exchange installation; you'll extract and install only the management tools which will install Exchange Management Shell (EMS) and other Exchange related tools but it will not include Exchange Admin Center (EAC) which is the management Web UI as that one is installed if you install the Mailbox server role. EMS won't connect unless there is a mailbox server role enabled Exchange Server available but the EMS will allow you to use the lightweight PowerShell Module/Snap-In

   - **Extract and Install Management Tools:**
     - Run the Exchange Server setup with the `/IAcceptExchangeServerLicenseTerms` and `/InstallManagementTools` switches.
       ```powershell
       Setup.exe /IAcceptExchangeServerLicenseTerms /InstallManagementTools
       ```
     - This installs the necessary PowerShell modules and prerequisites for recipient management.

2. **Open Exchange Management Shell (EMS):**

   - Launch **Exchange Management Shell** or **Windows PowerShell** as an administrator.
   - Import the **Exchange Recipient Management** module if not loaded automatically.
     ```powershell
     Import-Module ExchangeRecipientManagement
     ```

3. **Create a New User in Active Directory:**

   - Use the **Active Directory Users and Computers (ADUC)** console or PowerShell to create a new AD user.
     ```powershell
     New-ADUser -Name "John Doe" -GivenName "John" -Surname "Doe" -SamAccountName "jdoe" -UserPrincipalName "john.doe@yourdomain.com" -AccountPassword (Read-Host -AsSecureString "Enter Password") -Enabled $true
     ```
     - Replace the placeholders with the user's actual information.

4. **Mail-Enable the User:**

   - Use the `Enable-RemoteMailbox` cmdlet to mail-enable the user and associate the mailbox with Exchange Online.
     ```powershell
     Enable-RemoteMailbox -Identity "john.doe@yourdomain.com" -RemoteRoutingAddress "john.doe@yourdomain.mail.onmicrosoft.com"
     ```
     - The `RemoteRoutingAddress` is typically the user's UPN with the `.mail.onmicrosoft.com` domain suffix.

5. **Configure Additional Mailbox Properties (Optional):**

   - Set additional Exchange attributes as needed.
     ```powershell
     Set-RemoteMailbox -Identity "john.doe@yourdomain.com" -Alias "jdoe" -DisplayName "John Doe" -PrimarySmtpAddress "john.doe@yourdomain.com"
     ```

6. **Force Entra ID Connect Sync Synchronization:**

   - On the server running Entra ID Connect Sync, force a synchronization to ensure changes are synced promptly.
     ```powershell
     Start-ADSyncSyncCycle -PolicyType Delta
     ```

7. **Assign a License in Microsoft 365:**

   - Log in to the **Microsoft 365 Admin Center** (`https://admin.microsoft.com`).
   - Navigate to **Users** > **Active users**.
   - Locate and select the new user.
   - Click on **Licenses and Apps**.
   - Assign the appropriate license (e.g., **Office 365 E3**, **Exchange Online Plan 1**) and click **Save changes**.

8. **Inform the User:**

   - Provide the user with their account details and temporary password.
   - Instruct them to log in and change their password as soon as possible.

9. **Test the Mailbox:**

   - Verify that the mailbox is functioning correctly.
   - Send a test email to and from the mailbox.

**Notes:**

- **Exchange Schema Extensions:**
  - The on-premises AD must have the Exchange schema extensions, which are present if Exchange was previously installed.
  - If not present, you need to prepare the AD schema using the Exchange setup with the `/PrepareSchema` switch.
    ```powershell
    Setup.exe /PrepareSchema /IAcceptExchangeServerLicenseTerms
    ```

- **RBAC Roles:**
  - Ensure that your account has sufficient permissions to modify AD user objects and Exchange attributes.
  - Typically, membership in the **Account Operators** or a custom group with delegated permissions is required.

- **PowerShell Modules:**
  - The **Active Directory** module is required to manage AD user accounts.
  - The **Exchange Recipient Management** module provides the necessary cmdlets for Exchange attribute management.

---

## Offboarding Users with Exchange Recipient Management PowerShell Module

**Steps:**

1. **Disable the User Account in Active Directory:**

   - Disable the user's AD account to prevent further access.
     ```powershell
     Disable-ADAccount -Identity "john.doe@yourdomain.com"
     ```

2. **Convert the Mailbox to a Shared Mailbox (Optional):**

   - If you need to retain access to the user's mailbox data without a license, convert it to a shared mailbox.
   - **Connect to Exchange Online PowerShell:**
     ```powershell
     Import-Module ExchangeOnlineManagement
     Connect-ExchangeOnline -UserPrincipalName admin@yourdomain.com
     ```
   - **Convert the Mailbox:**
     ```powershell
     Set-Mailbox -Identity "john.doe@yourdomain.com" -Type Shared
     ```
     - **Note:** Converting to a shared mailbox must be done in Exchange Online.

3. **Remove Licenses in Microsoft 365:**

   - In the **Microsoft 365 Admin Center**, navigate to **Users** > **Active users**.
   - Select the user and go to **Licenses and Apps**.
   - Unassign all licenses from the user and click **Save changes**.

4. **Manage OneDrive Data:**

   - Transfer ownership of the user's OneDrive data if required.
     - In the **Microsoft 365 Admin Center**, select the user and click on **OneDrive**.
     - **Create an access link** or assign a manager to receive the user's OneDrive files.
     - Configure retention settings as per your organization's policies.

5. **Revoke Access and Authentication Tokens:**

   - **Block Sign-In:**
     - In the **Microsoft 365 Admin Center**, toggle **Block sign-in** to **On** for the user.
   - **Revoke Sessions:**
     - Use Entra ID PowerShell or the Entra ID Admin Center to revoke user sessions and reset authentication methods.
     ```powershell
     # Connect to Entra ID
     Connect-MgGraph -Scopes User.ReadWrite.All
     # Revoke sessions
     Revoke-MgUserSignInSession -UserId "john.doe@yourdomain.com"
     ```

6. **Backup Data (Recommended):**

   - Ensure that mailbox and OneDrive data are backed up using a cloud-to-cloud backup solution (e.g., Datto SaaS Protection).

7. **Force Entra ID Connect Sync Synchronization:**

   - On the server running Entra ID Connect Sync, force a synchronization to update the user's status in Entra ID.
     ```powershell
     Start-ADSyncSyncCycle -PolicyType Delta
     ```

8. **Verify Offboarding Completion:**

   - Confirm that the user's mailbox is converted to a shared mailbox and accessible as needed.
   - Ensure the user's account is disabled and that licenses have been removed.
   - Verify that OneDrive data has been managed according to policy.
   - Check that the user can no longer access company resources.

**Notes:**

- **Data Retention Policies:**
  - Always follow your organization's data retention and compliance policies when offboarding users.
  - Converting to a shared mailbox retains email data without requiring a license.

- **License Considerations:**
  - Shared mailboxes with over 50 GB of data or with In-Place Archive enabled require a license.
  - Ensure compliance with Microsoft licensing requirements.

- **Security Measures:**
  - Promptly blocking sign-in and revoking sessions helps prevent unauthorized access.
  - Regularly review access logs for any suspicious activity.

---

## Important Considerations

- **Decommissioning the Last On-Premises Exchange Server:**

  - **Recipient Management Tools Sufficiency:**
    - With the Exchange Recipient Management tools, you can manage recipients without an on-premises Exchange server.
    - This allows you to decommission the last Exchange mailbox server, reducing maintenance overhead.

  - **Schema and Attribute Management:**
    - The Exchange schema extensions must remain in AD.
    - Recipient attributes are still stored in AD and synchronized to Entra ID.

- **Entra ID Connect Sync:**

  - **Role of Entra ID Connect Sync:**
    - Continues to synchronize on-premises AD with Entra ID.
    - Must remain operational to sync changes made via the Recipient Management tools.

  - **Hybrid Writeback:**
    - Ensure that **Exchange Hybrid writeback** is enabled to maintain attribute consistency if required.

- **Administrative Permissions:**

  - **On-Premises AD:**
    - Permissions to modify user objects and Exchange attributes are necessary.
    - Delegate permissions appropriately to adhere to the principle of least privilege.

  - **Exchange Online:**
    - Administrative roles in Entra ID and Exchange Online are required for tasks like converting mailboxes to shared mailboxes.

- **Backup and Recovery:**

  - **Data Protection:**
    - Implement a robust backup solution for cloud data.
    - Regularly test backups to ensure data can be restored when needed.

- **Compliance and Auditing:**

  - **Logging Changes:**
    - Keep records of administrative actions for auditing purposes.
    - Utilize tools like the **Unified Audit Log** in Microsoft 365.

- **Support and Updates:**

  - **Staying Informed:**
    - Keep abreast of updates from Microsoft regarding hybrid environments and management tools.
    - Ensure that management workstations have the latest tools and updates installed.

---
