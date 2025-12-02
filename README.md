# PS_ACL-SharedMailboxes

PowerShell automation script for managing Microsoft 365 shared mailbox permissions at scale. The script synchronizes mailbox access rights based on an Excel workbook configured as an ACL-style matrix, automatically adding and removing permissions to match the desired state.

## What Does This Script Do?

This script automates the management of shared mailbox permissions in Microsoft 365 by:

1. **Reading Configuration from Excel** - Uses an Excel workbook as a centralized configuration source where you define which users or groups should have access to which shared mailboxes
2. **Querying Active Directory** - Dynamically resolves users based on flexible AD criteria (department, name, etc.) and expands group memberships
3. **Comparing Current vs. Desired State** - Retrieves existing permissions from Microsoft 365 and compares them against your Excel configuration
4. **Synchronizing Permissions** - Automatically adds or removes permissions to ensure actual mailbox access matches your desired configuration
5. **Managing Two Permission Types**:
   - **Read and Manage** (R) - Grants `FullAccess` permission, allowing users to open and manage the shared mailbox
   - **Send As** (S) - Grants `SendAs` permission, allowing users to send emails from the shared mailbox address

## Key Benefits

- **Configuration as Code** - Excel workbook serves as version-controllable documentation of mailbox permissions
- **Dynamic User Selection** - Define permissions based on AD attributes (e.g., "all users in Sales department") rather than maintaining individual user lists
- **Automated Reconciliation** - Script ensures permissions always match configuration, removing permissions no longer needed
- **Bulk Management** - Handle multiple shared mailboxes and permission assignments in a single execution
- **Group Expansion** - Automatically expand AD group memberships with recursive support for nested groups

## Excel Workbook Structure

The included **ACL-SharedMailboxes.xlsx** file uses the following format:

### Configuration Columns (A-E)

Define user/group selection criteria in each row:

- **Column A (Class)**: `User` or `Group`
- **Column B (Field)**: Active Directory property to search (e.g., `Name`, `Department`, `Title`, `Office`)
- **Column C (Search Term)**: Value to match (supports wildcards like `Sales*`)
- **Column D (Recursive)**: `Yes` or `No` - whether to recursively expand group memberships
- **Column E (Active)**: `Yes` or `No` - enable or disable this row

### Mailbox Columns (F onwards)

- **Row 1**: Email addresses of shared mailboxes
- **Rows 2+**: Permission assignments at the intersection of user selection and mailbox:
  - `R` = Read and Manage (FullAccess)
  - `S` = Send As
  - `RS` = Both permissions
  - Empty = No access

### Example Configuration

```
| Class | Field      | Search Term | Recursive | Active | sales@company.com | support@company.com |
|-------|------------|-------------|-----------|--------|-------------------|---------------------|
| Group | Name       | Sales Team  | Yes       | Yes    | RS                |                     |
| User  | Department | Support     | No        | Yes    |                   | R                   |
| Group | Name       | Managers    | No        | Yes    | R                 | S                   |
```

This configuration would:
- Give all Sales Team members (recursively) both Read and Send As rights to sales@company.com
- Give all users in Support department Read access to support@company.com
- Give all Managers Read access to sales@company.com and Send As rights to support@company.com

### Special Features

- **Hidden #CONFIG# worksheet** - Contains reference lists of available user/group properties
- **Skip worksheets** - Prefix any worksheet name with `#` to exclude it from processing
- **Multiple worksheets** - Organize different environments or departments in separate worksheets

## Prerequisites

- **PSExcel PowerShell module** - For reading Excel files
- **Active Directory module** - For querying users and groups
- **Exchange Online Admin role** - Script user needs this Azure AD role
- **Microsoft 365 credentials** - Admin account with mailbox management permissions

## Usage

1. Configure the script variables in `ACL-SharedMailboxes.ps1`:
   - Set `$msol_UserName` to your Microsoft 365 admin account
   - Adjust `$excelSourceFile` path if needed

2. Edit `ACL-SharedMailboxes.xlsx` to define your desired permissions

3. Run the script:
   ```powershell
   .\ACL-SharedMailboxes.ps1
   ```

4. On first run, you'll be prompted for your Microsoft 365 password, which will be encrypted and saved locally

## How It Works

The script follows this workflow:

1. Establishes a PowerShell session to Exchange Online
2. Reads the Excel workbook and processes each non-skipped worksheet
3. For each worksheet:
   - Queries Active Directory based on the criteria in columns A-E
   - Builds lists of users who should have each permission type
   - For each shared mailbox (columns F+):
     - Aggregates all users who should have access based on the matrix
     - Retrieves current permissions from Microsoft 365
     - Calculates differences (users to add, users to remove)
     - Applies changes via Exchange Online cmdlets
4. Closes the connection and reports completion

The script is **idempotent** - running it multiple times with the same configuration is safe and will only make changes when permissions are out of sync.
