# Table of Contents

1. [RBAC Roles and Permissions Required](#rbac-roles-and-permissions-required)
   - [On-Premises Active Directory and Exchange Server](#on-premises-active-directory-and-exchange-server)
   - [Entra ID (Formerly Azure Active Directory) and Exchange Online](#entra-id-formerly-azure-active-directory-and-exchange-online)
   - [Entra ID Connect Sync (Formerly Azure AD Connect Sync)](#entra-id-connect-sync-formerly-azure-ad-connect-sync)
   - [Important Considerations](#important-considerations)
2. [Best Practices for Administrative Accounts](#best-practices-for-administrative-accounts)
   - [Separate On-Premises and Cloud-Only Admin Accounts](#separate-on-premises-and-cloud-only-admin-accounts)
   - [Implement Multi-Factor Authentication (MFA)](#implement-multi-factor-authentication-mfa)
   - [Principle of Least Privilege](#principle-of-least-privilege)
   - [Use Privileged Access Workstations (PAWs)](#use-privileged-access-workstations-paws)
3. [Enabling Exchange Hybrid Writeback in Entra ID Connect Sync](#enabling-exchange-hybrid-writeback-in-entra-id-connect-sync)
   - [Benefits of Enabling Exchange Hybrid Writeback](#benefits-of-enabling-exchange-hybrid-writeback)
   - [How to Enable Exchange Hybrid Writeback](#how-to-enable-exchange-hybrid-writeback)
   - [Important Considerations](#important-considerations-1)
   - [Why Exchange Hybrid Writeback is Recommended](#why-exchange-hybrid-writeback-is-recommended)
4. [Option 1: Using the Exchange Admin Center](#option-1-using-the-exchange-admin-center)
   - [Onboarding Users with Exchange Admin Center](#onboarding-users-with-exchange-admin-center)
   - [Offboarding Users with Exchange Admin Center](#offboarding-users-with-exchange-admin-center)
5. [Option 2: Using the Exchange Management Shell](#option-2-using-the-exchange-management-shell)
   - [Onboarding Users with Exchange Management Shell](#onboarding-users-with-exchange-management-shell)
   - [Offboarding Users with Exchange Management Shell](#offboarding-users-with-exchange-management-shell)
6. [Option 3: Using the Exchange Recipient Management PowerShell Module](#option-3-using-the-exchange-recipient-management-powershell-module)
   - [Onboarding Users with Exchange Recipient Management PowerShell Module](#onboarding-users-with-exchange-recipient-management-powershell-module)
   - [Offboarding Users with Exchange Recipient Management PowerShell Module](#offboarding-users-with-exchange-recipient-management-powershell-module)
   - [Important Considerations](#important-considerations-2)
7. [Installing Exchange Server 2019 on Windows Server Core 2019/2022](#installing-exchange-server-2019-on-windows-server-core-20192022)
   - [Overview](#overview)
   - [Operating System Requirements](#operating-system-requirements)
   - [Benefits of Using Server Core](#benefits-of-using-server-core)
   - [Installation Steps](#installation-steps)
   - [Managing Exchange Server on Server Core](#managing-exchange-server-on-server-core)
   - [Recommended Approach Over Server with Desktop Experience](#recommended-approach-over-server-with-desktop-experience)
   - [Important Considerations](#important-considerations-3)
   - [Conclusion](#conclusion)




| Aspect                             | Option 1: Exchange Admin Center (EAC) | Option 2: Exchange Management Shell (EMS) | Option 3: Exchange Recipient Management PowerShell Module |
|------------------------------------|---------------------------------------|-------------------------------------------|----------------------------------------------------------|
| **Description**                    | Web-based GUI for managing Exchange   | PowerShell cmdlets for Exchange management| PowerShell module for recipient management without on-premises Exchange server |
| **Primary Tools Used**             | Exchange Admin Center (Web GUI)       | Exchange Management Shell (PowerShell)    | Exchange Recipient Management PowerShell Module |
| **Prerequisites**                  | - Access to EAC<br>- Admin permissions<br>- Hybrid setup with Exchange Online | - EMS installed<br>- Admin permissions<br>- Hybrid setup with Exchange Online | - Exchange Recipient Management tools installed<br>- Exchange schema extended in AD<br>- No on-premises Exchange server required |
| **Administrative Skill Level**     | Beginner to Intermediate              | Intermediate to Advanced                  | Intermediate to Advanced |
| **Permissions Required**           | - On-premises: Recipient Management Role<br>- Cloud: User Administrator, Exchange Administrator | - On-premises: Recipient Management Role<br>- Cloud: User Administrator, Exchange Administrator | - On-premises AD permissions<br>- Cloud: User Administrator, Exchange Administrator |
| **Ease of Use**                    | User-friendly interface               | Requires knowledge of PowerShell commands | Requires knowledge of PowerShell; no GUI |
| **Capabilities**                   | - Full recipient management<br>- GUI-based management<br>- Onboarding/offboarding | - Advanced management tasks<br>- Scripting and automation<br>- Onboarding/offboarding | - Manage recipients without on-premises Exchange server<br>- Onboarding/offboarding<br>- Decommission on-premises Exchange server |
| **Limitations**                    | - Less suitable for automation<br>- GUI may be slower for bulk operations | - Requires PowerShell expertise<br>- On-premises Exchange server must be maintained | - Requires PowerShell expertise<br>- Limited to recipient management<br>- Some Exchange features unavailable without on-premises server |
| **Use Cases**                      | - Small to medium organizations<br>- Administrators preferring GUI<br>- Regular recipient management | - Automation and scripting<br>- Bulk operations<br>- Advanced configuration | - Organizations migrating fully to Exchange Online<br>- Decommissioning on-premises Exchange servers<br>- Minimizing infrastructure |
| **Best For**                       | - GUI-based management needs<br>- Less frequent changes<br>- Simpler environments | - Complex environments<br>- Need for automation<br>- Experienced admins | - Reducing on-premises footprint<br>- Post-migration management<br>- Maintaining AD as source of authority |
| **Decommission On-Premises Exchange Server** | No                                | No                                        | Yes |
| **Hybrid Exchange Writeback Recommended** | Yes                               | Yes                                       | Yes |
| **Example Commands/Actions**       | - Create mailbox via GUI<br>- Convert mailbox to shared via GUI | - `New-RemoteMailbox`<br>- `Set-Mailbox`<br>- PowerShell scripts | - `Enable-RemoteMailbox`<br>- `Set-RemoteMailbox`<br>- Manage via AD and sync |
| **Management Interface**           | Web browser (EAC URL)                 | Exchange Management Shell (EMS)           | PowerShell on management workstation |
| **Dependency on On-Premises Exchange Server** | Yes                               | Yes                                       | No |
| **Notes**                          | - Simplifies basic tasks<br>- Less flexible for bulk changes | - More control and flexibility<br>- Suitable for automation | - Requires Exchange schema in AD<br>- Ideal for post-migration scenarios |





# Quick Reference Cheat Sheet: Managing Exchange Attributes and Remote User Mailboxes

This cheat sheet provides a concise overview of the three options for onboarding and offboarding users in a hybrid Exchange environment. Use this as a quick reference to streamline your administrative tasks.

---

## Option 1: Using the Exchange Admin Center (EAC)

### Description

- **Web-based GUI** for managing Exchange recipients.
- Ideal for administrators who prefer a graphical interface.

### Onboarding Users

1. **Access EAC**:
   - Navigate to `https://<Your-Exchange-Server>/ecp`.
   - Log in with administrative credentials.

2. **Create Remote Mailbox**:
   - Go to **Recipients** > **Mailboxes**.
   - Click **Add** (➕) > **Office 365 Mailbox**.
   - Enter user details and set a temporary password.

3. **Save and Synchronize**:
   - Save the new user.
   - Force synchronization with Entra ID Connect Sync if necessary:
     ```powershell
     Start-ADSyncSyncCycle -PolicyType Delta
     ```

4. **Assign License**:
   - In the **Microsoft 365 Admin Center**, assign the appropriate license to the user.

### Offboarding Users

1. **Convert to Shared Mailbox**:
   - In EAC, select the user's mailbox.
   - Click **Convert to shared mailbox**.

2. **Remove License**:
   - Unassign licenses in the **Microsoft 365 Admin Center**.

3. **Manage OneDrive Data**:
   - Transfer ownership or set retention policies as needed.

4. **Disable User Account**:
   - Disable the account in **Active Directory Users and Computers**.
   - Block sign-in and revoke sessions in Entra ID.

### Considerations

- **RBAC Roles**: Ensure you have the necessary roles assigned.
- **Data Retention**: Converting to a shared mailbox retains email data without a license.
- **Backup**: Use a cloud-to-cloud backup solution before making changes.

[Detailed Steps for Option 1](#option-1-using-the-exchange-admin-center)

---

## Option 2: Using the Exchange Management Shell (EMS)

### Description

- **PowerShell-based management** for advanced control and automation.
- Suitable for bulk operations and scripting.

### Onboarding Users

1. **Open EMS**:
   - Launch the Exchange Management Shell as an administrator.

2. **Create Remote Mailbox**:
   ```powershell
   $Password = Read-Host "Enter temporary password" -AsSecureString
   New-RemoteMailbox -Name "John Doe" -UserPrincipalName "john.doe@yourdomain.com" -Password $Password
   ```

3. **Configure User Properties (Optional)**:
   ```powershell
   Set-User -Identity "john.doe@yourdomain.com" -Department "Sales"
   ```

4. **Force Synchronization**:
   ```powershell
   Start-ADSyncSyncCycle -PolicyType Delta
   ```

5. **Assign License**:
   - Assign the appropriate license in the **Microsoft 365 Admin Center**.

### Offboarding Users

1. **Disable AD Account**:
   ```powershell
   Disable-ADAccount -Identity "john.doe@yourdomain.com"
   ```

2. **Convert to Shared Mailbox**:
   - Connect to Exchange Online PowerShell:
     ```powershell
     Connect-ExchangeOnline -UserPrincipalName admin@yourdomain.com
     ```
   - Convert mailbox:
     ```powershell
     Set-Mailbox -Identity "john.doe@yourdomain.com" -Type Shared
     ```

3. **Remove License**:
   - Unassign licenses in the **Microsoft 365 Admin Center**.

4. **Revoke Access**:
   - Block sign-in and revoke sessions in Entra ID.

### Considerations

- **Requires PowerShell Expertise**.
- **On-Premises Exchange Server Must Be Maintained**.
- **Ideal for Automation and Advanced Configurations**.

[Detailed Steps for Option 2](#option-2-using-the-exchange-management-shell)

---

## Option 3: Using the Exchange Recipient Management PowerShell Module

### Description

- **Manage recipients without an on-premises Exchange server**.
- Allows decommissioning of the last on-premises Exchange mailbox server.

### Onboarding Users

1. **Install Management Tools**:
   ```powershell
   Setup.exe /IAcceptExchangeServerLicenseTerms /InstallManagementTools
   ```

2. **Create AD User**:
   ```powershell
   New-ADUser -Name "John Doe" -UserPrincipalName "john.doe@yourdomain.com" -AccountPassword (Read-Host -AsSecureString "Enter Password") -Enabled $true
   ```

3. **Enable Remote Mailbox**:
   ```powershell
   Enable-RemoteMailbox -Identity "john.doe@yourdomain.com" -RemoteRoutingAddress "john.doe@yourdomain.mail.onmicrosoft.com"
   ```

4. **Force Synchronization**:
   ```powershell
   Start-ADSyncSyncCycle -PolicyType Delta
   ```

5. **Assign License**:
   - Assign the appropriate license in the **Microsoft 365 Admin Center**.

### Offboarding Users

1. **Disable AD Account**:
   ```powershell
   Disable-ADAccount -Identity "john.doe@yourdomain.com"
   ```

2. **Convert to Shared Mailbox**:
   - Connect to Exchange Online PowerShell:
     ```powershell
     Connect-ExchangeOnline -UserPrincipalName admin@yourdomain.com
     ```
   - Convert mailbox:
     ```powershell
     Set-Mailbox -Identity "john.doe@yourdomain.com" -Type Shared
     ```

3. **Remove License**:
   - Unassign licenses in the **Microsoft 365 Admin Center**.

4. **Revoke Access**:
   - Block sign-in and revoke sessions in Entra ID.

### Considerations

- **No On-Premises Exchange Server Required**.
- **Exchange Schema Extensions Must Be Present in AD**.
- **Ideal for Post-Migration Management**.

[Detailed Steps for Option 3](#option-3-using-the-exchange-recipient-management-powershell-module)

---

## General Best Practices

- **RBAC Roles and Permissions**:
  - Assign minimal permissions necessary.
  - Roles do not sync between on-premises and cloud; assign roles in both environments.

- **Enabling Exchange Hybrid Writeback**:
  - Recommended to maintain attribute consistency.
  - Enable via Entra ID Connect Sync configuration.

- **Administrative Accounts**:
  - Use separate on-premises and cloud-only admin accounts.
  - Avoid synchronizing admin accounts to Entra ID.

- **Security Measures**:
  - Implement Multi-Factor Authentication (MFA) for all admin accounts.
  - Follow the principle of least privilege.
  - Use Privileged Access Workstations (PAWs) for sensitive tasks.

- **Data Backup and Retention**:
  - Use cloud-to-cloud backup solutions before making changes.
  - Configure retention policies for OneDrive and mailboxes.

- **Monitoring and Maintenance**:
  - Regularly monitor synchronization logs.
  - Keep Entra ID Connect Sync updated and operational.
  - Stay informed about updates from Microsoft.

---

**Note**: This cheat sheet provides a high-level overview. For detailed instructions and additional considerations, refer to the full sections linked above.

---

# Navigation Links

- **[Option 1: Using the Exchange Admin Center](#option-1-using-the-exchange-admin-center)**
- **[Option 2: Using the Exchange Management Shell](#option-2-using-the-exchange-management-shell)**
- **[Option 3: Using the Exchange Recipient Management PowerShell Module](#option-3-using-the-exchange-recipient-management-powershell-module)**
- **[RBAC Roles and Permissions Required](#rbac-roles-and-permissions-required)**
- **[Enabling Exchange Hybrid Writeback](#enabling-exchange-hybrid-writeback-in-entra-id-connect-sync)**
- **[Best Practices for Administrative Accounts](#best-practices-for-administrative-accounts)**

---

**Remember**: Always adhere to your organization's policies and compliance requirements when managing user accounts and data.

If you need more detailed guidance, consult the specific sections in the full guide.

---

# End of Cheat Sheet


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
 




# Installing Exchange Server 2019 on Windows Server Core 2019/2022

This section provides guidance on installing **Exchange Server 2019** on **Windows Server Core 2019** or **Windows Server Core 2022**. It highlights the operating system requirements, benefits of using Server Core, and management options available. Deploying Exchange Server on Server Core is recommended over the Desktop Experience version due to its reduced footprint and security advantages.

## Overview

- **Exchange Server 2019** supports installation on **Windows Server Core** editions.
- **Server Core** provides a minimal installation option, reducing the server's surface area for attacks and maintenance requirements.
- Exchange Server is primarily managed via **PowerShell** and remote management tools, making Server Core an ideal platform.

## Operating System Requirements

### Supported Operating Systems

- **Windows Server 2019 Server Core**
- **Windows Server 2022 Server Core**

**Note:** Ensure that the latest cumulative updates for both Windows Server and Exchange Server are applied to support all features and receive security updates.

### Prerequisites

- **Hardware Requirements:**
  - **Processor:** 64-bit processor with at least 2 cores.
  - **Memory:** Minimum of 128 GB RAM (recommended for Mailbox servers).
  - **Disk Space:** At least 30 GB free space on the installation drive, plus additional space for mailbox databases and logs.
- **Active Directory:**
  - Active Directory forest functional level of **Windows Server 2012 R2** or higher.
  - At least one writable Global Catalog server in each Active Directory site where you plan to install Exchange.
- **Networking:**
  - Static IP address assigned to the server.
  - Proper DNS configuration.

### Software Requirements

- **.NET Framework 4.8**
- **Visual C++ Redistributable Package for Visual Studio 2012**
- **Unified Communications Managed API (UCMA) 4.0 Runtime**
- **Windows Features:**
  - **Exchange Server Mailbox role:** Install required Windows features using the provided scripts or manually via PowerShell.

## Benefits of Using Server Core

- **Reduced Attack Surface:**
  - Fewer installed components mean fewer vulnerabilities.
- **Lower Maintenance:**
  - Reduced need for updates and patches compared to the full Desktop Experience.
- **Improved Performance:**
  - Less resource consumption due to the absence of GUI components.
- **Enhanced Security:**
  - Minimalistic design limits the potential for unauthorized access.
- **Efficiency:**
  - Ideal for roles that do not require a GUI, such as Exchange Server, which is primarily managed via PowerShell and remote tools.

## Installation Steps

### Step 1: Prepare the Server

1. **Install Windows Server Core:**
   - Install either **Windows Server 2019 Server Core** or **Windows Server 2022 Server Core**.
   - Configure the server with a static IP address and join it to the domain.

2. **Update the Server:**
   - Install the latest Windows updates.
     ```powershell
     sconfig
     ```
     - Select option **6** to download and install updates.

3. **Configure Time Zone:**
   - Set the correct time zone.
     ```powershell
     tzutil /s "Pacific Standard Time"
     ```

4. **Rename the Server (Optional):**
   - Assign a meaningful name to the server.
     ```powershell
     Rename-Computer -NewName "EXCH2019CORE"
     ```
     - Reboot the server to apply changes.

### Step 2: Install Required Windows Features

- **Install Prerequisites Using PowerShell:**

  For the **Mailbox** role, execute the following command:

  ```powershell
  Install-WindowsFeature Server-Media-Foundation, RSAT-ADDS
  ```

- **Install Other Required Components:**

  - **.NET Framework 4.8:**
    - Download and install the offline installer for .NET Framework 4.8.

  - **Visual C++ Redistributable Package for Visual Studio 2012:**
    - Download and install both **vcredist_x64.exe** and **vcredist_x86.exe**.

  - **Unified Communications Managed API (UCMA) 4.0 Runtime:**
    - Download and install **Ucmaredist.msi**.

### Step 3: Prepare Active Directory

1. **Extend the Schema:**
   - Run the following command from the Exchange installation files directory:
     ```powershell
     Setup.exe /PrepareSchema /IAcceptExchangeServerLicenseTerms
     ```

2. **Prepare Active Directory:**
   - Prepare the Active Directory forest:
     ```powershell
     Setup.exe /PrepareAD /OrganizationName:"YourOrganization" /IAcceptExchangeServerLicenseTerms
     ```
   - Prepare all domains:
     ```powershell
     Setup.exe /PrepareAllDomains /IAcceptExchangeServerLicenseTerms
     ```

### Step 4: Install Exchange Server 2019

1. **Mount the Exchange Server 2019 Installation Media:**
   - Copy the Exchange installation files to a local directory or mount the ISO.

2. **Install Exchange Server Using Setup Command:**
   - Execute the following command to install the Mailbox role:
     ```powershell
     Setup.exe /Mode:Install /Role:Mailbox /IAcceptExchangeServerLicenseTerms
     ```

   - **Note:** If you plan to use the Edge Transport role, it must be installed on a separate server.

3. **Complete the Installation:**
   - The setup will progress through several stages, including copying files, installing the Exchange Server roles, and configuring services.
   - Upon completion, verify that the installation was successful.

### Step 5: Apply the Latest Cumulative Update

- **Download and Install Cumulative Update:**
  - Obtain the latest cumulative update (CU) for Exchange Server 2019 from Microsoft's official website.
  - Run the CU installer:
    ```powershell
    Setup.exe /Mode:Upgrade /IAcceptExchangeServerLicenseTerms
    ```

## Managing Exchange Server on Server Core

With Server Core lacking a graphical user interface (GUI), management is performed using remote tools and command-line interfaces.

### Windows Admin Center (WAC)

- **Overview:**
  - WAC is a browser-based management tool that provides a modern interface for managing servers.
- **Installation:**
  - Install WAC on a separate management server or workstation.
- **Features:**
  - Manage server settings, services, firewall, and performance.
  - Access PowerShell remotely.
- **Usage:**
  - Connect to the Server Core instance via WAC to perform administrative tasks.

### PowerShell

- **Exchange Management Shell (EMS):**
  - Use EMS for Exchange-specific management tasks.
  - Access EMS remotely from a management workstation:
    ```powershell
    Enter-PSSession -ComputerName EXCH2019CORE
    ```
- **Remote PowerShell:**
  - Enable PowerShell remoting on the Server Core server.
    ```powershell
    Enable-PSRemoting -Force
    ```
  - Manage the server using standard PowerShell cmdlets.

### Remote Server Administration Tools (RSAT)

- **Overview:**
  - RSAT allows administrators to manage Windows Servers remotely from a Windows 10/11 workstation.
- **Installation:**
  - Install RSAT components via **Settings** > **Apps** > **Optional features**.
- **Features:**
  - Access tools like Active Directory Users and Computers, DNS management, and more.

### Exchange Administration Center (EAC)

- **Access via Web Browser:**
  - The EAC is a web-based management console accessible from a remote machine.
  - URL format:
    ```
    https://<ExchangeServerFQDN>/ecp
    ```
- **Usage:**
  - Perform recipient management, organization configuration, and other administrative tasks.

### Server Manager

- **Limited Functionality:**
  - Server Manager cannot be run locally on Server Core but can manage Server Core instances remotely.
- **Usage:**
  - Add the Server Core server to Server Manager on a remote workstation or server.
  - Manage roles, features, and basic server settings.

### Other Remote Management Tools

- **Remote Desktop Services (RDS):**
  - Connect via RDS for command-line access.
- **Event Viewer:**
  - Access logs remotely using **Event Viewer** connected to the Server Core instance.
- **Performance Monitor:**
  - Monitor performance counters remotely.

## Recommended Approach Over Server with Desktop Experience

- **Efficiency:**
  - Server Core reduces resource consumption by eliminating unnecessary GUI components.
- **Security:**
  - Fewer installed features reduce the attack surface and potential vulnerabilities.
- **Maintenance:**
  - Decreased need for updates and patches simplifies maintenance schedules.
- **Management Alignment:**
  - Since Exchange Server is primarily managed via PowerShell and remote tools, the absence of a GUI does not hinder administrative tasks.
- **Cost-Effective:**
  - Lower resource requirements can lead to cost savings in hardware and energy consumption.

## Important Considerations

- **Training:**
  - Administrators should be comfortable with command-line tools and remote management.
- **Compatibility:**
  - Verify that all third-party applications and monitoring tools support Server Core.
- **Troubleshooting:**
  - Without a GUI, troubleshooting requires familiarity with command-line diagnostics and remote tools.
- **Backup and Recovery:**
  - Ensure that your backup solutions support Server Core environments.

## Conclusion

Deploying Exchange Server 2019 on Windows Server Core 2019 or 2022 is a strategic choice that enhances security, reduces maintenance overhead, and aligns with modern management practices. By leveraging remote management tools such as Windows Admin Center, PowerShell, and RSAT, administrators can efficiently manage Exchange Server without the need for a local GUI. This approach is recommended over the traditional Desktop Experience installation, particularly in environments where server resources and security are paramount.



# Option 1: Using the Exchange Admin Center

This section provides step-by-step instructions for onboarding and offboarding users using the **Exchange Admin Center (EAC)** in a hybrid Exchange environment.

## Onboarding Users with Exchange Admin Center

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

## Onboarding Users with Exchange Management Shell

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



> 🔔 **Note:** You can instead use this script https://github.com/aollivierre/Exchange/blob/main/Exchange2/1013/Exchange/New-RemoteSharedMailbox-v10%20copy.ps1 to automate the entire process. Please make sure to first update the values inside of the script to meet your needs before running it.



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
   - Import the **Exchange Recipient Management** snap-in
     ```powershell
     Run this https://github.com/aollivierre/Exchange/blob/main/Exchange2/1014/Exchange/1-Copy-Profile-PS5.ps1
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

