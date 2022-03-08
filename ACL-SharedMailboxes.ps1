<#======================================================================================
ACL-SharedMailboxes v1.0.1                                                    2022-03-07
----------------------------------------------------------------------------------------
Author: /u/bwientjes
======================================================================================#>

# Where to find the input Excel workbook
$excelSourceFile                        = ".\ACL-SharedMailboxes.xlsx"                          # Path to input Excel workbook

# Layout settings
$colorText                              = "Cyan"
$colorOK                                = "Green"
$colorError                             = "Red"

# Office365 credentials
$msol_URI								= "https://outlook.office365.com/powershell-liveid/"	# FQDN of Exchange Client Access Server
$msol_Auth								= "Basic"											    # "Kerberos" or "Basic"
$msol_UserName							= "msol_user@company.com"                               # Microsoft Online user	(this user needs the Exchange Online Admin role in AzureAD)	   


<#======================================================================================
Do not edit under this section unless you know what you're doing!
======================================================================================#>

# Internal variables
$WorkingDirectory    					    = Split-Path ((Get-Variable MyInvocation).Value).MyCommand.Path
$EncryptedPassword_File					    = "EncryptedPassword.txt"
$EncryptedPassword_LocalFile			    = Join-Path $WorkingDirectory $EncryptedPassword_File

Function ConnectExchange($msol_URI, $msol_UserName, $EncryptedPassword_LocalFile, $msol_Auth) {
    Write-Host "Check for Microsoft 365 connectivity..." -ForegroundColor $colorText

    Try {
        $Session = Get-OutboundConnector -ErrorAction SilentlyContinue | Out-null
        Write-Host "Already connected." -ForegroundColor $colorText
    }
    Catch {
        if (!$Session) {
            Write-Host "No active session found." -ForegroundColor $colorText
            Write-Host "Starting new PowerShell session..." -ForegroundColor $colorText
        
            If (!(Test-Path $EncryptedPassword_LocalFile)) { 
                Write-Host "$EncryptedPassword_File not found" -ForegroundColor $colorError
                Read-Host -Prompt "Enter password for user $msol_UserName" -AsSecureString | ConvertFrom-SecureString | Out-File $EncryptedPassword_LocalFile
                
                $msol_Password    = Get-Content $EncryptedPassword_LocalFile | ConvertTo-SecureString
                $msol_Credentials = New-Object -typename System.Management.Automation.PSCredential `
                                               -argumentlist $msol_UserName, $msol_Password
            
            } Else {
                $msol_Password    = Get-Content $EncryptedPassword_LocalFile | ConvertTo-SecureString
                $msol_Credentials = New-Object -typename System.Management.Automation.PSCredential `
                                               -argumentlist $msol_UserName, $msol_Password
            }
            Try {
                $PSSession = New-PSSession -ConfigurationName Microsoft.Exchange `
                                           -ConnectionUri $msol_URI `
                                           -Credential $msol_Credentials `
                                           -Authentication $msol_Auth `
                                           -AllowRedirection `
                                           -Name "ACLSharedMailboxes"
            
                Import-PSSession $PSSession -AllowClobber -DisableNameChecking | Out-null
                                
                $Session = Get-PSSession -Name "ACLSharedMailboxes" -ErrorAction Stop
                Write-Host "Session started" -ForegroundColor $colorOK
            }
            Catch  {
                Write-Host "!- Canot connect to Microsoft 365 -!" -ForegroundColor $colorError
                Write-Host "Double check your credentials. You can delete the ExcryptedPassword.txt file to have the script prompt for a password again." -ForegroundColor $colorError
                
				Scheduledtaskcode "1" ;# Last Run Result = 0x1
                break
            }
        
        }
    }
}

Function CloseConnectionExchange() {
    Write-Host ""
	Write-Host "Disconnect the Remote PowerShell session: " -NoNewLine -ForegroundColor $colorText
	Remove-PSSession -Name "ACLSharedMailboxes"
	Write-Host "OK" -ForegroundColor $colorOK
}

Function SearchUsers($class, $field, $searchTerm, $recursive) {
    <#
    Takes a user class, a field to search in, a search term, and wether or not searches should be done recursively
    ald returns an object list of users that are found.
    #>
    $sourceList = @()

    switch($class) {

        'Group' {
            # Check is groups are found.
            $groupsFound = Get-ADGroup -Filter "${field} -like '${searchTerm}'" -ErrorAction 'silentlycontinue'
            if($groupsFound) {
                # If the group exists, add all members of the group.
                Get-ADGroup -Filter "${field} -like '${searchTerm}'" | ForEach-Object {
                    if($recursive -eq 'Yes') {
                        Get-ADGroupMember -Identity $_.Name -Recursive | Select-Object UserPrincipalName | ForEach-Object {
                            $sourceList += $_
                        }
                    } else {
                        Get-ADGroupMember -Identity $_.Name | Select-Object UserPrincipalName | Where-Object objectClass -eq "user" | ForEach-Object {
                            $sourceList += $_
                        }
                    }
                }
            } else {
                Write-Host ""
                Write-Host "Groups with ${field} of '${searchTerm}' do not exist. Skipping row ${rowIndex}." -ForegroundColor Red
            }
            Break
        }
    
        'User' {
            # Check is users are found.
            $usersFound = Get-ADUser -Filter "${field} -like '${searchTerm}'" -ErrorAction 'silentlycontinue'
            if($usersFound) {
                # If a user is found, add it to the source list.
                Get-ADUser -Filter "${field} -like '${searchTerm}'" | Select-Object UserPrincipalName | ForEach-Object {
                    $sourceList += $_
                }
            } else {
                Write-Host ""
                Write-Host "Users with ${field} of '${searchTerm}' do not exist. Skipping row ${rowIndex}." -ForegroundColor Red
            }
            Break
        }
    
    }

    Return $sourceList
}

Function GeneratedUserLists() {
    $userLists = @{}
    for($rowIndex=2; $rowIndex -le $worksheet.Dimension.Rows; $rowIndex++) {

        # Get the search term for this rights assignment
        $class          = $worksheet.Cells.Item($rowIndex,1).text
        $field          = $worksheet.Cells.Item($rowIndex,2).text
        $searchTerm     = $worksheet.Cells.Item($rowIndex,3).text
        $recursive      = $worksheet.Cells.Item($rowIndex,4).text
        $active         = $worksheet.Cells.Item($rowIndex,5).text

        # Only proceed if all fields are filled and $active is set to 'Yes', otherwise we have either an inactive row or incomplete information.
        if(($class -ne '') -and ($field -ne '') -and ($searchTerm -ne '') -and ($recursive -ne '') -and ($active -eq 'Yes')) {
            $userList = SearchUsers $class $field $searchTerm $recursive
            $userLists.Add($rowIndex,$userList)
        }

    }

    Return $userLists

}


# First and foremost, we need functionality from the PSExcel module.
Import-Module "PSExcel"

# Put some pretty header on the console
Clear-Host
Write-Host "ACL-SharedMailboxes v1.0.1" -ForegroundColor $colorText
Write-Host "-------------------------------------------------------------------------------------" -ForegroundColor $colorText
Write-Host ""
Write-Host "Read Excel workbook: " -ForegroundColor $colorText -NoNewline
Write-Host $excelSourceFile -ForegroundColor White
Write-Host ""
ConnectExchange $msol_URI $msol_UserName $EncryptedPassword_LocalFile $msol_Auth
Write-Host ""

# Now that we have Excel cmdlets through PSExcel, let's open the workbook. First, make a local copy so that the script won't fail when someone has the workbook open.
Copy-Item $excelSourceFile ".\input.xslx"
$excelObject = New-Excel -Path ".\input.xslx"
$workbook = $excelObject | Get-Workbook

ForEach($worksheet in @($workbook.Worksheets)) {
    $worksheetName = $worksheet.Name
    # Only process worksheets without the # (skip) prefix, i.e. worksheets whose names are preceded with # are skipped.
    if($worksheetName -notlike "#*") {

        Write-Host "Worksheet: " -ForegroundColor $colorText -NoNewLine
        Write-Host $worksheetName -ForegroundColor White
        Write-Host ""

        # Generate a hashtable that contains arrays of user objects as defined in the first few columns in the Excel workbook.
        # The key index corresponds with a row numer (i.e. $userLists[2] is the aggregated user list according to the search on row 2).
        # These lists will later be used to provide access (after concatenating a set of lists and sanitizing them).
        Write-Host "Generating user lists based on information entered in columns A through E..." -ForegroundColor $colorText
        $userLists = GeneratedUserLists

        # Process all columns, staring at column 6 (containing shared mailbox adresses).
        for($colIndex=6; $colIndex -le $worksheet.Dimension.Columns; $colIndex++) {

            # Check if the cell is filled. Skip empty ones.
            $sharedMailBoxName = $worksheet.Cells.Item(1,$colIndex).text
            if($sharedMailBoxName -ne '') {

                Write-Host ""
                Write-Host "Shared Mailbox: " -ForegroundColor $colorText -NoNewline
                Write-Host $sharedMailBoxName -ForegroundColor White

                # Clear the user lists so that we can fill them.
                $newRead = @()
                $newSendAs = @()

                # Now look for filled cells in that column, and construct a list for Reand and Manage access and Sens As access.
                for($rowIndex=2; $rowIndex -le $worksheet.Dimension.Rows; $rowIndex++) {

                    # Process only active rows.
                    $activeRow = $worksheet.Cells.Item($rowIndex,5).text
                    if($activeRow -eq 'Yes') {

                        $rights = $worksheet.Cells.Item($rowIndex,$colIndex).text

                        # Check for Read and Manage rights
                        if($rights.toLower() -like '*r*') {
                            $newRead += $userLists[$rowIndex]
                        }
    
                        # Check for Send As rights
                        if($rights.toLower() -like '*s*') {
                            $newSendAs += $userLists[$rowIndex]
                        }

                    }
    
                }

                # Cleanup duplicates from the "Read and Manage" and "Send As" lists ($userListRead and $userListSendAs respectively), and select only the UPN (we don't need the rest).
                $newRead    = $newRead | Sort-Object -Property UserPrincipalName -Unique
                $newSendAs  = $newSendAs | Sort-Object -Property UserPrincipalName -Unique
                $newRead    = $newRead.UserPrincipalName    #.toLower() (doesn't work if the array is empty, but this is not case sensitive so it won't matter)
                $newSendAs  = $newSendAs.UserPrincipalName  #.toLower() (doesn't work if the array is empty, but this is not case sensitive so it won't matter)
                Write-Host "  -" $newRead.Count "user(s) need 'Read and Manage' access." -ForegroundColor $colorText
                Write-Host "  -" $newSendAs.Count "user(s) need 'Send As' access." -ForegroundColor $colorText

                # Get the currect "Read and Manage" and "Send As" access for that mailbox, so we can compare them later. Omit the NT AUTHORITY\SELF.
                $currentRead    = Get-MailboxPermission -Identity $sharedMailBoxName | Select-Object User | Where-Object {($_.User -ne "NT AUTHORITY\SELF")}
                $currentSendAs  = Get-RecipientPermission -Identity $sharedMailBoxName | Select-Object Trustee | Where-Object {($_.Trustee -ne "NT AUTHORITY\SELF")}

                # Convert any email addresses to lowercase, so that we will be comparing apples to apples later.
                $currentRead    = $currentRead.User         #.toLower() (doesn't work if the array is empty, but this is not case sensitive so it won't matter)
                $currentSendAs  = $currentSendAs.Trustee    #.toLower() (doesn't work if the array is empty, but this is not case sensitive so it won't matter)

                # Compare the current and user Read lists to construct an Add and Remove list
                if(($newRead.Count -gt 0) -and ($currentRead.Count -gt 0)) {
                    $comparison     = Compare-Object -ReferenceObject $currentRead -DifferenceObject $newRead
                    $removeRead     = $comparison.where{$_.SideIndicator -eq '<='}.InputObject
                    $addRead        = $comparison.where{$_.SideIndicator -eq '=>'}.InputObject
                } elseif($newRead.Count -eq 0) {
                    $removeRead     = $currentRead
                } elseif($currentRead.Count -eq 0) {
                    $addRead        = $newRead
                }

                # Compare the current and user Send As lists to construct an Add and Remove list
                if(($newSendAs.Count -gt 0) -and ($currentSendAs.Count -gt 0)) {
                    $comparison     = Compare-Object -ReferenceObject $currentSendAs -DifferenceObject $newSendAs
                    $removeSendAs   = $comparison.where{$_.SideIndicator -eq '<='}.InputObject
                    $addSendAs      = $comparison.where{$_.SideIndicator -eq '=>'}.InputObject
                } elseif($newSendAs.Count -eq 0) {
                    $removeSendAs   = $currentSendAs
                } elseif($currentSendAs.Count -eq 0) {
                    $addSendAs      = $newSendAs
                }

                # Remove users that no longer need access to each access type
                $numberOfUsers = $removeRead.Count
                if($numberOfUsers -gt 0) {
                    $pct = 0
                    $pctstep = 100 / $numberOfUsers
                    ForEach($user in $removeRead) {
                        $pct += $pctstep
                        [int]$step = $pct
                        Write-Progress -Activity "  - Remove $numberOfUsers user(s) from 'Read and Manage'" -Status "$step%" -PercentComplete $pct
                        $dummy = Remove-MailboxPermission -Identity $sharedMailBoxName -User $user -AccessRights FullAccess -Confirm:$false
                    }
                    Write-Progress -Activity "  - Remove $numberOfUsers user(s) from 'Read and Manage'" -Completed
                    Write-Host "  - Remove $numberOfUsers user(s) from 'Read and Manage': " -NoNewLine -ForegroundColor $colorText
                    Write-Host "OK" -ForegroundColor $colorOK
                } else {
                    Write-Host "  - No users to be removed from 'Read and Manage'" -ForegroundColor $colorOK
                }

                $numberOfUsers = $removeSendAs.Count
                if($numberOfUsers -gt 0) {
                    $pct = 0
                    $pctstep = 100 / $numberOfUsers
                    ForEach($user in $removeSendAs) {
                        $pct += $pctstep
                        [int]$step = $pct
                        Write-Progress -Activity "  - Remove $numberOfUsers user(s) from 'Send As'" -Status "$step%" -PercentComplete $pct
                        $dummy = Remove-RecipientPermission -Identity $sharedMailBoxName -Trustee $user -AccessRights SendAs -Confirm:$false
                    }
                    Write-Progress -Activity "  - Remove $numberOfUsers user(s) from 'Send As'" -Completed
                    Write-Host "  - Remove $numberOfUsers user(s) from 'Send As': " -NoNewLine -ForegroundColor $colorText
                    Write-Host "OK" -ForegroundColor $colorOK
                } else {
                    Write-Host "  - No users to be removed from 'Send As'" -ForegroundColor $colorOK
                }

                # Add users that need access to each access type.
                $numberOfUsers = $addRead.Count
                if($numberOfUsers -gt 0) {
                    $pct = 0
                    $pctstep = 100 / $numberOfUsers
                    ForEach($user in $addRead) {
                        $pct += $pctstep
                        [int]$step = $pct
                        Write-Progress -Activity "  - Add $numberOfUsers user(s) to 'Read and Manage'" -Status "$step%" -PercentComplete $pct
                        $dummy = Add-MailboxPermission -Identity $sharedMailBoxName -User $user -AccessRights FullAccess -Confirm:$false
                    }
                    Write-Progress -Activity "  - Add $numberOfUsers user(s) to 'Read and Manage'" -Completed
                    Write-Host "  - Add $numberOfUsers user(s) to 'Read and Manage': " -NoNewLine -ForegroundColor $colorText
                    Write-Host "OK" -ForegroundColor $colorOK
                } else {
                    Write-Host "  - No users to be added to 'Read and Manage'" -ForegroundColor $colorOK
                }

                $numberOfUsers = $addSendAs.Count
                if($numberOfUsers -gt 0) {
                    $pct = 0
                    $pctstep = 100 / $numberOfUsers
                    ForEach($user in $addSendAs) {
                        $pct += $pctstep
                        [int]$step = $pct
                        Write-Progress -Activity "  - Add $numberOfUsers user(s) to 'Send As'" -Status "$step%" -PercentComplete $pct
                        $dummy = Add-RecipientPermission -Identity $sharedMailBoxName -Trustee $user -AccessRights SendAs -Confirm:$false
                    }
                    Write-Progress -Activity "  - Add $numberOfUsers user(s) to 'Send As'" -Completed
                    Write-Host "  - Add $numberOfUsers user(s) to 'Send As': " -NoNewLine -ForegroundColor $colorText
                    Write-Host "OK" -ForegroundColor $colorOK
                } else {
                    Write-Host "  - No users to be added to 'Send As'" -ForegroundColor $colorOK
                }
    
            }

        }

    }

}

# When all is done, remove the local copy of the Excel workbook and close the Microsoft 365 connection.
Remove-Item ".\input.xslx"
CloseConnectionExchange

Write-Host ""
Write-Host "-------------------------------------------------------------------------------------" -ForegroundColor $colorText
Write-Host "Script executed successfuly." -ForegroundColor $colorOK
Write-Host "-------------------------------------------------------------------------------------" -ForegroundColor $colorText
Write-Host ""
Write-Host ""
