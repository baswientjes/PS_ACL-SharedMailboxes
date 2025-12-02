<#
.SYNOPSIS
    Manages Microsoft 365 shared mailbox permissions based on an Excel configuration matrix.

.DESCRIPTION
    ACL-SharedMailboxes automates the management of shared mailbox permissions in Microsoft 365 by:
    - Reading configuration from an Excel workbook in ACL-matrix format
    - Querying Active Directory for users/groups based on flexible criteria
    - Comparing current vs. desired permissions state
    - Synchronizing permissions to match the Excel configuration

    The script supports two permission types:
    - Read and Manage (R) - Grants FullAccess permission
    - Send As (S) - Grants SendAs permission

.PARAMETER ExcelSourceFile
    Path to the Excel workbook containing the permission matrix configuration.

.PARAMETER UserPrincipalName
    Microsoft 365 admin account username (UPN) with Exchange Online Admin role.
    If not specified, you'll be prompted to authenticate interactively.

.PARAMETER Organization
    The Exchange Online organization name (e.g., contoso.onmicrosoft.com).
    Optional parameter for specifying the tenant.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for certificate-based authentication (CBA).
    Use this for unattended/automated scenarios instead of interactive auth.

.PARAMETER AppId
    Application (client) ID for certificate-based authentication.
    Required when using -CertificateThumbprint.

.EXAMPLE
    .\ACL-SharedMailboxes.ps1
    Runs the script with interactive authentication.

.EXAMPLE
    .\ACL-SharedMailboxes.ps1 -UserPrincipalName admin@contoso.com
    Runs the script with the specified user principal name.

.EXAMPLE
    .\ACL-SharedMailboxes.ps1 -AppId "12345678-1234-1234-1234-123456789012" -CertificateThumbprint "A1B2C3..." -Organization "contoso.onmicrosoft.com"
    Runs the script using certificate-based authentication for automation scenarios.

.NOTES
    Version:        3.0.0
    Author:         /u/bwientjes
    Updated:        2025-12-02

    Prerequisites:
    - ExchangeOnlineManagement module (v3.0.0 or later)
    - PSExcel PowerShell module
    - Active Directory module
    - Exchange Online Admin role in Azure AD

    Breaking Changes from v2.0:
    - Now uses ExchangeOnlineManagement module instead of deprecated Remote PowerShell (RPS)
    - Removed Basic authentication (deprecated by Microsoft)
    - Uses modern authentication with MFA support
    - Certificate-based authentication available for automation

.LINK
    https://github.com/baswientjes/PS_ACL-SharedMailboxes
#>

[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param(
    [Parameter()]
    [ValidateScript({Test-Path -Path $_ -PathType Leaf})]
    [string]$ExcelSourceFile = ".\ACL-SharedMailboxes.xlsx",

    [Parameter(ParameterSetName = 'Interactive')]
    [string]$UserPrincipalName,

    [Parameter(ParameterSetName = 'Interactive')]
    [string]$Organization,

    [Parameter(Mandatory, ParameterSetName = 'CertificateAuth')]
    [string]$AppId,

    [Parameter(Mandatory, ParameterSetName = 'CertificateAuth')]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory, ParameterSetName = 'CertificateAuth')]
    [string]$Organization
)

#Requires -Modules ExchangeOnlineManagement, PSExcel, ActiveDirectory

# Script-level variables
$script:WorkingDirectory = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$script:TempExcelFile = Join-Path -Path $script:WorkingDirectory -ChildPath "input.xlsx"

# Color settings for console output
$script:ColorText = "Cyan"
$script:ColorOK = "Green"
$script:ColorError = "Red"

#region Functions

function Connect-M365ExchangeOnline {
    <#
    .SYNOPSIS
        Establishes a connection to Exchange Online using the ExchangeOnlineManagement module.

    .DESCRIPTION
        Connects to Exchange Online using modern authentication. Supports both interactive
        authentication (with MFA) and certificate-based authentication for automation scenarios.

    .PARAMETER UserPrincipalName
        The user principal name for interactive authentication.

    .PARAMETER Organization
        The Exchange Online organization name.

    .PARAMETER AppId
        Application (client) ID for certificate-based authentication.

    .PARAMETER CertificateThumbprint
        Certificate thumbprint for certificate-based authentication.

    .EXAMPLE
        Connect-M365ExchangeOnline -UserPrincipalName "admin@contoso.com"

    .EXAMPLE
        Connect-M365ExchangeOnline -AppId "12345" -CertificateThumbprint "ABC123" -Organization "contoso.onmicrosoft.com"
    #>
    [CmdletBinding(DefaultParameterSetName = 'Interactive')]
    param(
        [Parameter(ParameterSetName = 'Interactive')]
        [string]$UserPrincipalName,

        [Parameter(ParameterSetName = 'Interactive')]
        [string]$Organization,

        [Parameter(Mandatory, ParameterSetName = 'CertificateAuth')]
        [string]$AppId,

        [Parameter(Mandatory, ParameterSetName = 'CertificateAuth')]
        [string]$CertificateThumbprint,

        [Parameter(Mandatory, ParameterSetName = 'CertificateAuth')]
        [string]$Organization
    )

    Write-Verbose "Checking for existing Exchange Online connectivity..."
    Write-Host "Checking for Microsoft 365 Exchange Online connectivity..." -ForegroundColor $script:ColorText

    try {
        # Test if already connected by trying to get a mailbox
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-Host "Already connected to Exchange Online." -ForegroundColor $script:ColorText
        return
    }
    catch {
        Write-Verbose "No active connection found. Establishing new connection..."
        Write-Host "No active connection found." -ForegroundColor $script:ColorText
        Write-Host "Connecting to Exchange Online..." -ForegroundColor $script:ColorText

        try {
            # Build connection parameters based on authentication method
            $connectionParams = @{
                ShowBanner = $false
                ErrorAction = 'Stop'
            }

            if ($PSCmdlet.ParameterSetName -eq 'CertificateAuth') {
                # Certificate-based authentication for automation
                $connectionParams['AppId'] = $AppId
                $connectionParams['CertificateThumbprint'] = $CertificateThumbprint
                $connectionParams['Organization'] = $Organization
                Write-Verbose "Using certificate-based authentication"
            }
            else {
                # Interactive authentication with MFA support
                if ($UserPrincipalName) {
                    $connectionParams['UserPrincipalName'] = $UserPrincipalName
                }
                if ($Organization) {
                    $connectionParams['Organization'] = $Organization
                }
                Write-Verbose "Using interactive authentication"
            }

            # Connect to Exchange Online
            Connect-ExchangeOnline @connectionParams

            # Verify connection
            $null = Get-OrganizationConfig -ErrorAction Stop
            Write-Host "Successfully connected to Exchange Online" -ForegroundColor $script:ColorOK
        }
        catch {
            Write-Host "!- Cannot connect to Exchange Online -!" -ForegroundColor $script:ColorError
            Write-Host "Error: $_" -ForegroundColor $script:ColorError
            Write-Error "Failed to establish Exchange Online connection: $_"
            throw
        }
    }
}

function Disconnect-M365ExchangeOnline {
    <#
    .SYNOPSIS
        Closes the Exchange Online connection.

    .DESCRIPTION
        Disconnects from Exchange Online and cleans up the session.

    .EXAMPLE
        Disconnect-M365ExchangeOnline
    #>
    [CmdletBinding()]
    param()

    try {
        Write-Verbose "Disconnecting from Exchange Online..."
        Write-Host "`nDisconnecting from Exchange Online: " -NoNewline -ForegroundColor $script:ColorText
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Write-Host "OK" -ForegroundColor $script:ColorOK
    }
    catch {
        Write-Warning "Could not disconnect session: $_"
    }
}

function Find-ADUsers {
    <#
    .SYNOPSIS
        Searches Active Directory for users or groups based on criteria.

    .DESCRIPTION
        Queries Active Directory for users or groups matching the specified field and search term.
        Supports recursive group membership expansion.

    .PARAMETER Class
        The object class to search for (User or Group).

    .PARAMETER Field
        The AD property to search in (e.g., Name, Department, Title).

    .PARAMETER SearchTerm
        The value to search for (supports wildcards).

    .PARAMETER Recursive
        Whether to recursively expand group memberships.

    .PARAMETER RowIndex
        The Excel row index being processed (for error messages).

    .OUTPUTS
        System.Array of user objects with UserPrincipalName property.

    .EXAMPLE
        Find-ADUsers -Class 'Group' -Field 'Name' -SearchTerm 'Sales*' -Recursive $true -RowIndex 2
    #>
    [CmdletBinding()]
    [OutputType([System.Array])]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('User', 'Group')]
        [string]$Class,

        [Parameter(Mandatory)]
        [string]$Field,

        [Parameter(Mandatory)]
        [string]$SearchTerm,

        [Parameter(Mandatory)]
        [bool]$Recursive,

        [Parameter()]
        [int]$RowIndex
    )

    $userList = @()

    try {
        switch ($Class) {
            'Group' {
                Write-Verbose "Searching for groups where $Field is like '$SearchTerm'"
                $groups = Get-ADGroup -Filter "$Field -like '$SearchTerm'" -ErrorAction Stop

                if ($groups) {
                    foreach ($group in $groups) {
                        Write-Verbose "Processing group: $($group.Name)"

                        if ($Recursive) {
                            $members = Get-ADGroupMember -Identity $group.Name -Recursive -ErrorAction Stop |
                                       Select-Object -Property UserPrincipalName
                        }
                        else {
                            $members = Get-ADGroupMember -Identity $group.Name -ErrorAction Stop |
                                       Where-Object -Property objectClass -EQ 'user' |
                                       Select-Object -Property UserPrincipalName
                        }

                        $userList += $members
                    }
                }
                else {
                    Write-Host "`nGroups with $Field of '$SearchTerm' do not exist. Skipping row $RowIndex." -ForegroundColor $script:ColorError
                }
            }

            'User' {
                Write-Verbose "Searching for users where $Field is like '$SearchTerm'"
                $users = Get-ADUser -Filter "$Field -like '$SearchTerm'" -ErrorAction Stop |
                         Select-Object -Property UserPrincipalName

                if ($users) {
                    $userList += $users
                }
                else {
                    Write-Host "`nUsers with $Field of '$SearchTerm' do not exist. Skipping row $RowIndex." -ForegroundColor $script:ColorError
                }
            }
        }
    }
    catch {
        Write-Error "Error searching AD for $Class with $Field='$SearchTerm': $_"
    }

    return $userList
}

function Get-UserListsFromWorksheet {
    <#
    .SYNOPSIS
        Generates user lists from Excel worksheet configuration.

    .DESCRIPTION
        Reads the worksheet configuration rows and builds a hashtable of user lists
        indexed by row number, based on the AD search criteria in columns A-E.

    .PARAMETER Worksheet
        The Excel worksheet object to process.

    .OUTPUTS
        System.Collections.Hashtable of user lists indexed by row number.

    .EXAMPLE
        Get-UserListsFromWorksheet -Worksheet $worksheet
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [object]$Worksheet
    )

    $userLists = @{}

    for ($rowIndex = 2; $rowIndex -le $Worksheet.Dimension.Rows; $rowIndex++) {
        # Get the search criteria for this row
        $class = $Worksheet.Cells.Item($rowIndex, 1).Text
        $field = $Worksheet.Cells.Item($rowIndex, 2).Text
        $searchTerm = $Worksheet.Cells.Item($rowIndex, 3).Text
        $recursiveText = $Worksheet.Cells.Item($rowIndex, 4).Text
        $active = $Worksheet.Cells.Item($rowIndex, 5).Text

        # Only proceed if all fields are filled and row is active
        if ($class -and $field -and $searchTerm -and $recursiveText -and ($active -eq 'Yes')) {
            $recursive = $recursiveText -eq 'Yes'

            Write-Verbose "Processing row $rowIndex : Class=$class, Field=$field, SearchTerm=$searchTerm, Recursive=$recursive"

            $userList = Find-ADUsers -Class $class -Field $field -SearchTerm $searchTerm -Recursive $recursive -RowIndex $rowIndex
            $userLists.Add($rowIndex, $userList)
        }
    }

    return $userLists
}

function Sync-MailboxPermissions {
    <#
    .SYNOPSIS
        Synchronizes mailbox permissions to match desired state.

    .DESCRIPTION
        Compares current mailbox permissions with desired permissions and adds/removes
        users as needed to match the configuration.

    .PARAMETER MailboxName
        The email address of the shared mailbox.

    .PARAMETER DesiredReadUsers
        Array of user principal names who should have Read and Manage access.

    .PARAMETER DesiredSendAsUsers
        Array of user principal names who should have Send As access.

    .EXAMPLE
        Sync-MailboxPermissions -MailboxName 'sales@company.com' -DesiredReadUsers @('user1@company.com') -DesiredSendAsUsers @('user2@company.com')
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$MailboxName,

        [Parameter()]
        [string[]]$DesiredReadUsers = @(),

        [Parameter()]
        [string[]]$DesiredSendAsUsers = @()
    )

    Write-Verbose "Syncing permissions for mailbox: $MailboxName"

    try {
        # Get current permissions
        $currentReadPermissions = Get-MailboxPermission -Identity $MailboxName -ErrorAction Stop |
                                  Where-Object {$_.User -ne 'NT AUTHORITY\SELF'} |
                                  Select-Object -ExpandProperty User

        $currentSendAsPermissions = Get-RecipientPermission -Identity $MailboxName -ErrorAction Stop |
                                    Where-Object {$_.Trustee -ne 'NT AUTHORITY\SELF'} |
                                    Select-Object -ExpandProperty Trustee

        # Calculate differences for Read permissions
        $readToRemove = @()
        $readToAdd = @()

        if ($DesiredReadUsers.Count -gt 0 -and $currentReadPermissions.Count -gt 0) {
            $comparison = Compare-Object -ReferenceObject $currentReadPermissions -DifferenceObject $DesiredReadUsers
            $readToRemove = $comparison | Where-Object {$_.SideIndicator -eq '<='} | Select-Object -ExpandProperty InputObject
            $readToAdd = $comparison | Where-Object {$_.SideIndicator -eq '=>'} | Select-Object -ExpandProperty InputObject
        }
        elseif ($DesiredReadUsers.Count -eq 0) {
            $readToRemove = $currentReadPermissions
        }
        elseif ($currentReadPermissions.Count -eq 0) {
            $readToAdd = $DesiredReadUsers
        }

        # Calculate differences for Send As permissions
        $sendAsToRemove = @()
        $sendAsToAdd = @()

        if ($DesiredSendAsUsers.Count -gt 0 -and $currentSendAsPermissions.Count -gt 0) {
            $comparison = Compare-Object -ReferenceObject $currentSendAsPermissions -DifferenceObject $DesiredSendAsUsers
            $sendAsToRemove = $comparison | Where-Object {$_.SideIndicator -eq '<='} | Select-Object -ExpandProperty InputObject
            $sendAsToAdd = $comparison | Where-Object {$_.SideIndicator -eq '=>'} | Select-Object -ExpandProperty InputObject
        }
        elseif ($DesiredSendAsUsers.Count -eq 0) {
            $sendAsToRemove = $currentSendAsPermissions
        }
        elseif ($currentSendAsPermissions.Count -eq 0) {
            $sendAsToAdd = $DesiredSendAsUsers
        }

        # Remove Read and Manage permissions
        Update-PermissionSet -MailboxName $MailboxName -Users $readToRemove -PermissionType 'Read' -Action 'Remove'

        # Remove Send As permissions
        Update-PermissionSet -MailboxName $MailboxName -Users $sendAsToRemove -PermissionType 'SendAs' -Action 'Remove'

        # Add Read and Manage permissions
        Update-PermissionSet -MailboxName $MailboxName -Users $readToAdd -PermissionType 'Read' -Action 'Add'

        # Add Send As permissions
        Update-PermissionSet -MailboxName $MailboxName -Users $sendAsToAdd -PermissionType 'SendAs' -Action 'Add'
    }
    catch {
        Write-Error "Failed to sync permissions for mailbox $MailboxName : $_"
    }
}

function Update-PermissionSet {
    <#
    .SYNOPSIS
        Adds or removes a set of permissions for a mailbox.

    .DESCRIPTION
        Processes a list of users to add or remove specific permission types from a mailbox
        with progress indication.

    .PARAMETER MailboxName
        The email address of the shared mailbox.

    .PARAMETER Users
        Array of user principal names to process.

    .PARAMETER PermissionType
        Type of permission (Read or SendAs).

    .PARAMETER Action
        Action to perform (Add or Remove).

    .EXAMPLE
        Update-PermissionSet -MailboxName 'sales@company.com' -Users @('user@company.com') -PermissionType 'Read' -Action 'Add'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$MailboxName,

        [Parameter()]
        [string[]]$Users = @(),

        [Parameter(Mandatory)]
        [ValidateSet('Read', 'SendAs')]
        [string]$PermissionType,

        [Parameter(Mandatory)]
        [ValidateSet('Add', 'Remove')]
        [string]$Action
    )

    $numberOfUsers = $Users.Count

    if ($numberOfUsers -eq 0) {
        $permissionLabel = if ($PermissionType -eq 'Read') { 'Read and Manage' } else { 'Send As' }
        Write-Host "  - No users to be ${Action}ed from/to '$permissionLabel'" -ForegroundColor $script:ColorOK
        return
    }

    $permissionLabel = if ($PermissionType -eq 'Read') { 'Read and Manage' } else { 'Send As' }
    $activityMessage = "  - $Action $numberOfUsers user(s) to/from '$permissionLabel'"

    try {
        $processedCount = 0
        foreach ($user in $Users) {
            $processedCount++
            $percentComplete = ($processedCount / $numberOfUsers) * 100
            Write-Progress -Activity $activityMessage -Status "$processedCount of $numberOfUsers" -PercentComplete $percentComplete

            if ($PermissionType -eq 'Read') {
                if ($Action -eq 'Add') {
                    $null = Add-MailboxPermission -Identity $MailboxName -User $user -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop
                }
                else {
                    $null = Remove-MailboxPermission -Identity $MailboxName -User $user -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop
                }
            }
            else {
                if ($Action -eq 'Add') {
                    $null = Add-RecipientPermission -Identity $MailboxName -Trustee $user -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                }
                else {
                    $null = Remove-RecipientPermission -Identity $MailboxName -Trustee $user -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                }
            }
        }

        Write-Progress -Activity $activityMessage -Completed
        Write-Host "  - $Action $numberOfUsers user(s) to/from '$permissionLabel': " -NoNewline -ForegroundColor $script:ColorText
        Write-Host "OK" -ForegroundColor $script:ColorOK
    }
    catch {
        Write-Progress -Activity $activityMessage -Completed
        Write-Error "Failed to $Action permission for mailbox $MailboxName : $_"
    }
}

#endregion Functions

#region Main Script

try {
    # Display header
    Clear-Host
    Write-Host "ACL-SharedMailboxes v3.0.0" -ForegroundColor $script:ColorText
    Write-Host "-------------------------------------------------------------------------------------" -ForegroundColor $script:ColorText
    Write-Host ""
    Write-Host "Read Excel workbook: " -ForegroundColor $script:ColorText -NoNewline
    Write-Host $ExcelSourceFile -ForegroundColor White
    Write-Host ""

    # Import required modules
    Write-Verbose "Importing required modules..."
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Import-Module PSExcel -ErrorAction Stop

    # Connect to Exchange Online
    $connectionParams = @{}
    if ($PSCmdlet.ParameterSetName -eq 'CertificateAuth') {
        $connectionParams['AppId'] = $AppId
        $connectionParams['CertificateThumbprint'] = $CertificateThumbprint
        $connectionParams['Organization'] = $Organization
    }
    else {
        if ($UserPrincipalName) {
            $connectionParams['UserPrincipalName'] = $UserPrincipalName
        }
        if ($Organization) {
            $connectionParams['Organization'] = $Organization
        }
    }

    Connect-M365ExchangeOnline @connectionParams
    Write-Host ""

    # Create temporary copy of Excel file to avoid locking issues
    Write-Verbose "Creating temporary copy of Excel workbook..."
    Copy-Item -Path $ExcelSourceFile -Destination $script:TempExcelFile -Force -ErrorAction Stop

    # Open the Excel workbook
    $excelObject = New-Excel -Path $script:TempExcelFile
    $workbook = $excelObject | Get-Workbook

    # Process each worksheet
    foreach ($worksheet in $workbook.Worksheets) {
        $worksheetName = $worksheet.Name

        # Skip worksheets with # prefix
        if ($worksheetName -like '#*') {
            Write-Verbose "Skipping worksheet: $worksheetName"
            continue
        }

        Write-Host "Worksheet: " -ForegroundColor $script:ColorText -NoNewline
        Write-Host $worksheetName -ForegroundColor White
        Write-Host ""

        # Generate user lists based on columns A-E
        Write-Host "Generating user lists based on information entered in columns A through E..." -ForegroundColor $script:ColorText
        $userLists = Get-UserListsFromWorksheet -Worksheet $worksheet

        # Process each shared mailbox (starting from column 6)
        for ($colIndex = 6; $colIndex -le $worksheet.Dimension.Columns; $colIndex++) {
            $sharedMailboxName = $worksheet.Cells.Item(1, $colIndex).Text

            if (-not $sharedMailboxName) {
                continue
            }

            Write-Host "`nShared Mailbox: " -ForegroundColor $script:ColorText -NoNewline
            Write-Host $sharedMailboxName -ForegroundColor White

            # Collect users who should have each permission type
            $desiredReadUsers = @()
            $desiredSendAsUsers = @()

            for ($rowIndex = 2; $rowIndex -le $worksheet.Dimension.Rows; $rowIndex++) {
                $activeRow = $worksheet.Cells.Item($rowIndex, 5).Text

                if ($activeRow -ne 'Yes') {
                    continue
                }

                $rights = $worksheet.Cells.Item($rowIndex, $colIndex).Text

                # Check for Read and Manage rights
                if ($rights -like '*[Rr]*') {
                    $desiredReadUsers += $userLists[$rowIndex]
                }

                # Check for Send As rights
                if ($rights -like '*[Ss]*') {
                    $desiredSendAsUsers += $userLists[$rowIndex]
                }
            }

            # Remove duplicates and extract UPNs
            $desiredReadUsers = $desiredReadUsers |
                                Sort-Object -Property UserPrincipalName -Unique |
                                Select-Object -ExpandProperty UserPrincipalName

            $desiredSendAsUsers = $desiredSendAsUsers |
                                  Sort-Object -Property UserPrincipalName -Unique |
                                  Select-Object -ExpandProperty UserPrincipalName

            Write-Host "  - $($desiredReadUsers.Count) user(s) need 'Read and Manage' access." -ForegroundColor $script:ColorText
            Write-Host "  - $($desiredSendAsUsers.Count) user(s) need 'Send As' access." -ForegroundColor $script:ColorText

            # Synchronize permissions
            Sync-MailboxPermissions -MailboxName $sharedMailboxName -DesiredReadUsers $desiredReadUsers -DesiredSendAsUsers $desiredSendAsUsers
        }
    }

    Write-Host ""
    Write-Host "-------------------------------------------------------------------------------------" -ForegroundColor $script:ColorText
    Write-Host "Script executed successfully." -ForegroundColor $script:ColorOK
    Write-Host "-------------------------------------------------------------------------------------" -ForegroundColor $script:ColorText
    Write-Host ""
}
catch {
    Write-Host ""
    Write-Host "-------------------------------------------------------------------------------------" -ForegroundColor $script:ColorError
    Write-Host "Script execution failed!" -ForegroundColor $script:ColorError
    Write-Host "Error: $_" -ForegroundColor $script:ColorError
    Write-Host "-------------------------------------------------------------------------------------" -ForegroundColor $script:ColorError
    Write-Host ""
    throw
}
finally {
    # Cleanup temporary file
    if (Test-Path -Path $script:TempExcelFile) {
        Write-Verbose "Removing temporary Excel file..."
        Remove-Item -Path $script:TempExcelFile -Force -ErrorAction SilentlyContinue
    }

    # Disconnect from Exchange Online
    Disconnect-M365ExchangeOnline
}

#endregion Main Script
