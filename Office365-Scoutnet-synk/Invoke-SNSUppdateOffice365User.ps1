
$Script:SNSyncGroupName='scoutnet'
$Script:SNSyncGroupDescription="Säkerhetsgrupp som används vid synkronisering med Scoutnet. Användare i gruppen avaktiveras om de inte är kvar i Scoutnet."

$Script:SNSyncGroupDisabledUsersName='scoutnetDisabledUsers'
$Script:SNSyncGroupDisabledUsersDescription="Säkerhetsgrupp som används vid synkronisering med Scoutnet. Användare i gruppen är avaktiverade och finns inte längre med i Scoutnet."
$Script:SNSAllUsersGroupName=""
$Script:SNSLicenseAssignment=@{""=""}
$Script:SNSPreferredLanguage="sv-SE"
$Script:SNSUsageLocation="SE"
$Script:SNSWaitMailboxCreationMaxTime="1200"
$Script:SNSWaitMailboxCreationPollTime="30"
$Script:SNSSignatureText=""
$Script:SNSSignatureHtml=""

$Script:SNSNewUserEmailSubject=""
$Script:SNSNewUserEmailText=""
$Script:SNSemailSMTPServer = "outlook.office365.com"
$Script:SNSNewUserInfoEmailSubject=""
$Script:SNSNewUserInfoEmailText=""
$Script:SNSemailFromAddress = ""

Export-ModuleMember -Variable SNSyncGroupName,SNSyncGroupDescription,SNSyncGroupDisabledUsersName,SNSyncGroupDisabledUsersDescription
Export-ModuleMember -Variable SNSLicenseAssignment,SNSPreferredLanguage,SNSUsageLocation,SNSWaitMailboxCreationMaxTime,SNSWaitMailboxCreationPollTime
Export-ModuleMember -Variable SNSSignatureText,SNSSignatureHtml,SNSNewUserEmailSubject,SNSNewUserEmailText,SNSemailSMTPServer,SNSNewUserInfoEmailSubject,SNSNewUserInfoEmailText,SNSemailFromAddress

function Invoke-SNSUppdateOffice365User
{
    <#
    .SYNOPSIS
        Main function for syncronisation of office 365 users with scoutnet.

    .INPUTS
        None. You cannot pipe objects to Get-SNSUppdateOffice365User.

    .OUTPUTS
        None.
    #>
    param (
        [Parameter(Mandatory=$False, HelpMessage="Credentials for api/group/customlists.")]
        [ValidateNotNull()]
        [pscredential]$CredentialCustomlists,

        [Parameter(Mandatory=$False, HelpMessage="Credentials for api/group/memberlist.")]
        [ValidateNotNull()]
        [pscredential]$CredentialMemberlist,

        [Parameter(Mandatory=$False, HelpMessage="Credentials for office365.")]
        [ValidateNotNull()]
        [pscredential]$Credential365,

        [Parameter(Mandatory=$False, HelpMessage="Maillist to process.")]
        $MailListId,

        [Parameter(Mandatory=$False, HelpMessage="Domain name for office365 mail addresses.")]
        [ValidateNotNull()]
        [string]$DomainName
    )

    try
    {
        Get-MsolDomain -ErrorAction Stop > $null
    }
    catch
    {
        Write-SNSLog "Connecting to Office 365..."
        try
        {
            Connect-MsolService -Credential $Credential365 -ErrorAction Stop
        }
        catch
        {
            Write-SNSLog -Level "Error" "Could not connect to Office 365. Error $_"
            throw
        }
    }

    Write-SNSLog "Start of user account update"
    $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential365 -Authentication Basic -AllowRedirection
    Import-PSSession $ExchangeSession -AllowClobber -CommandName Set-MailContact,Set-Mailbox,Get-Mailbox,Remove-DistributionGroupMember,Add-DistributionGroupMember,Get-DistributionGroupMember,Get-DistributionGroup,Set-MailboxMessageConfiguration > $null

    [System.Collections.ArrayList]$allOffice365Users = Get-Mailbox -RecipientTypeDetails "UserMailbox"
    [System.Collections.ArrayList]$SecurityGroupScoutnet = Get-SNSUsersInSecurityGroupScoutnet -allOffice365Users $allOffice365Users
    [System.Collections.ArrayList]$UsersInSecurityGroupScoutnetDisabledUsers = Get-SNSUsersInSecurityGroupScoutnetDisabledUsers -allOffice365Users $allOffice365Users
    [System.Collections.ArrayList]$MemberListScoutnet = Get-SNSSoutnetLeaders -CredentialCustomlists $CredentialCustomlists -CredentialMemberlist $CredentialMemberlist -MailListId $MailListId
    $GroupMemberlist = Get-SNSApiGroupMemberlist -Credential $CredentialMemberlist

    if (!$GroupMemberlist)
    {
        throw "No data returned from Scoutnet. Update aborted."
    }

    Write-SNSLog "Total number of office365 users: $($allOffice365Users.Count)"
    Write-SNSLog "Total number of Scoutnet members: $($MemberListScoutnet.Count)"

    Write-SNSLog "Start check"

    $NewMembers = [System.Collections.ArrayList]::new()
    $MembersToUpdate = [System.Collections.ArrayList]::new()
    $MembersToActivate = [System.Collections.ArrayList]::new()
    foreach($Member in $MemberListScoutnet)
    {
        # Check if the member do have an account.
        $account = $allOffice365Users | Where-Object -FilterScript {$_.CustomAttribute1 -like $Member}
        $isInSecurityGroupScoutnet = $SecurityGroupScoutnet | Where-Object -FilterScript {$_.CustomAttribute1 -like $Member}
        $isInUsersInSecurityGroupScoutnetDisabledUsers = $UsersInSecurityGroupScoutnetDisabledUsers | Where-Object -FilterScript {$_.CustomAttribute1 -like $Member}

        if ($account)
        {
            if ($isInSecurityGroupScoutnet)
            {
                # The member is in SecurityGroupScoutnet. Update office 365 account if needed.
                [void]$MembersToUpdate.Add($account)
                $SecurityGroupScoutnet.RemoveAt($SecurityGroupScoutnet.IndexOf($isInSecurityGroupScoutnet))
            }
            else
            {
                Write-SNSLog "Account not to update: '$($account.Name)'"
            }

            # Remove from list of all accounts.
            $allOffice365Users.RemoveAt($allOffice365Users.IndexOf($account))
        }
        elseif ($isInUsersInSecurityGroupScoutnetDisabledUsers)
        {
            # Returning member. Enable and update the office 365 account.
            [void]$MembersToActivate.Add($isInUsersInSecurityGroupScoutnetDisabledUsers)
            $UsersInSecurityGroupScoutnetDisabledUsers.RemoveAt($UsersInSecurityGroupScoutnetDisabledUsers.IndexOf($isInUsersInSecurityGroupScoutnetDisabledUsers))
            # Remove from list of all accounts.
            $allOffice365Users.RemoveAt($allOffice365Users.IndexOf($isInUsersInSecurityGroupScoutnetDisabledUsers))
        }
        else
        {
            # New member. Add account.
            $MemberData = $GroupMemberlist.data[$Member]
            [void]$NewMembers.Add($MemberData)
        }
    }

    Write-SNSLog "Start check done"
    Write-SNSLog "Number of accounts to create: $($NewMembers.Count)"
    if ($NewMembers.Count -gt 0)
    {
        try
        {
            Invoke-SNSCreateUserAndUpdateUserData -memberData $NewMembers -Credential365 $Credential365 -DomainName $DomainName
        }
        catch
        {
            Write-SNSLog -Level "Error" "Error during user account creation. Error $_"
        }
    }

    Write-SNSLog "Number of accounts to activate: $($MembersToActivate.Count)"
    $MembersToActivate | ForEach-Object {
        Invoke-SNSEnableAccount -AccountData $_
        [void]$MembersToUpdate.Add($_)
    }

    Write-SNSLog "Number of accounts to disable: $($SecurityGroupScoutnet.Count)"
    $SecurityGroupScoutnet | ForEach-Object {
        Invoke-SNSDisableAccount -AccountData $_
        [void]$UsersInSecurityGroupScoutnetDisabledUsers.Add($_)
    }

    Write-SNSLog "Number of accounts to check for update: $($MembersToUpdate.Count)"
    $MembersToUpdate | ForEach-Object {
        Invoke-SNSUpdateAccount -AccountData $_ -Credential $CredentialMemberlist
    }

    Write-SNSLog "Number of disabled accounts: $($UsersInSecurityGroupScoutnetDisabledUsers.Count)"
    $UsersInSecurityGroupScoutnetDisabledUsers | ForEach-Object {
        Write-SNSLog "Disabled account '$($_.Name)'"
        $allOffice365Users.RemoveAt($allOffice365Users.IndexOf($_))
    }

    Write-SNSLog "Number of accounts not connected to Scoutnet: $($allOffice365Users.Count)"
    $allOffice365Users | ForEach-Object {Write-SNSLog "Account not connected to Scoutnet '$($_.Name)'"}

    Remove-PSSession $ExchangeSession
}


function Invoke-SNSCreateUserAndUpdateUserData
{
    <#
    .SYNOPSIS
        Creates and updates the new account.
    
    .INPUTS
        None. You cannot pipe objects to Get-SNSUpdateAccount.

    .OUTPUTS
        None.
    #>
    param (
        [Parameter(Mandatory=$false, HelpMessage="List of members to create.")]
        [ValidateNotNull()]
        $memberData,

        [Parameter(Mandatory=$false, HelpMessage="Credentials for office365")]
        [ValidateNotNull()]
        [Alias("Credential")]
        [pscredential]$Credential365,

        [Parameter(Mandatory=$False, HelpMessage="Domain name for office365 mail addresses.")]
        [ValidateNotNull()]
        [string]$DomainName
    )

    $LicenseAssignment = [System.Collections.ArrayList]::new()
    $LicenseOptions = [System.Collections.ArrayList]::new()
    $SecurityGroupScoutnet = Get-SNSSecurityGroupScoutnet

    # Create licensing options.
    foreach($LicenseKey in $Script:SNSLicenseAssignment)
    {
        [void]$LicenseAssignment.Add($LicenseKey)
        $LO = New-MsolLicenseOptions -AccountSkuId $LicenseKey -DisabledPlans $Script:SNSLicenseAssignment[$LicenseKey] -ErrorAction "Stop"
        [void]$LicenseOptions.Add($LO)
    }

    if ($LicenseOptions.Count -eq 0)
    {
        $msg = "The parameter 'Script:SNSLicenseAssignment' did not contain any valid licenses."
        $msg += "Creation of account cannot be executed!"
        throw ($msg)
    }


    $newAccounts = [ordered]@{}
    $LastAccountUserPrincipalName=$null
    foreach($member in $memberData)
    {
        # Create the new account
        $newAccount = Invoke-SNSAddOffice365User -MemberData $member -PreferredLanguage `
            $Script:SNSPreferredLanguage -DomainName $DomainName -UsageLocation $Script:SNSUsageLocation `
            -LicenseAssignment $LicenseAssignment -LicenseOptions $LicenseOptions -Credential365 $Credential365

        if ($newAccount)
        {
            $newAccounts.Add($member.member_no.value, @($member,$newAccount))
            $LastAccountUserPrincipalName = $newAccount.UserPrincipalName
            try
            {
                # Add the user to the group of active users.
                Add-MsolGroupMember -GroupObjectId $SecurityGroupScoutnet.ObjectId -GroupMemberObjectId $newAccount.ObjectId -ErrorAction "Stop"
            }
            Catch
            {
                Write-SNSLog -Level "Warn" "Could not add contact '$($newAccount.DisplayName)' to group $($Script:SNSyncGroupName). Error $_"
            }
        }
    }

    $maxDateTimeout = (Get-Date).AddSeconds($Script:SNSWaitMailboxCreationMaxTime)
    $doLoop = $true
    Write-SNSLog "Wait for the mailbox to be created. It takes som time..."

    try
    {
        while($doLoop)
        {
            Start-Sleep -s $Script:SNSWaitMailboxCreationPollTime

            try
            {
                Get-Mailbox -Identity $LastAccountUserPrincipalName -RecipientTypeDetails "UserMailbox" -ErrorAction "Stop" > $null
                # Mailboxes is created.
                $doLoop=$false
                break
            }
            catch
            {
                Write-SNSLog "Wait for the mailbox to be created. It takes som time..."
            }

            if ($maxDateTimeout -lt (Get-Date))
            {
                # timeout limit reached so exception
                $msg = "The creation of user mailboxes did not"
                $msg += "complete within the timeout limit of "
                $msg += "$($Script:SNSWaitMailboxCreationMaxTime) seconds, so polling "
                $msg += "for mailbox creation was halted."
                throw ($msg)
            }
        }

        foreach($newAccountId in $newAccounts.Keys)
        {
            $member =  $newAccounts[$newAccountId][0]
            $newAccount =  $newAccounts[$newAccountId][1]

            $SignatureHtml = $Script:SNSSignatureHtml -Replace "<DisplayName>", $newAccount.DisplayName
            $SignatureText = $Script:SNSSignatureText -Replace "<DisplayName>", $newAccount.DisplayName

            Set-Mailbox -Identity $newAccount.UserPrincipalName -CustomAttribute1 $newAccountId -ErrorAction "Stop"
            Set-MailboxMessageConfiguration $newAccount.UserPrincipalName -IsReplyAllTheDefaultResponse $false `
                -SignatureHtml $SignatureHtml -SignatureText $SignatureText -AutoAddSignature $true -AutoAddSignatureOnMobile $true `
                -AutoAddSignatureOnReply $true -SignatureTextOnMobile $SignatureText -ErrorAction "Stop"

            $emailFromAddress = $Script:SNSemailFromAddress
            if ([string]::IsNullOrWhiteSpace($emailFromAddress))
            {
                # Use the login credential as from address.
                $emailFromAddress = $Credential365.UserName
            }

            # Send e-mail to the user with the new password and account info. The password must be replaced at first login.
            $NewUserEmailText = $Script:SNSNewUserEmailText -Replace "<DisplayName>", $newAccount.DisplayName -Replace "<Password>", $newAccount.Password -Replace "<UserPrincipalName>", $newAccount.UserPrincipalName
            Send-MailMessage -From $emailFromAddress -to $member.email.value -Body $NewUserEmailText `
                -SmtpServer $Script:SNSemailSMTPServer -Credential $Credential365 -UseSsl -Subject $Script:SNSNewUserEmailSubject -Encoding ([System.Text.Encoding]::UTF8)

            if (![string]::IsNullOrWhiteSpace($Script:SNSNewUserInfoEmailText))
            {
                # Extra info mail requested. Send it to the new account.
                $NewUserEmailText = $Script:SNSNewUserInfoEmailText -Replace "<DisplayName>", $newAccount.DisplayName -Replace "<UserPrincipalName>", $newAccount.UserPrincipalName
                Send-MailMessage -From $emailFromAddress -to  $newAccount.UserPrincipalName -Body $NewUserEmailText `
                    -SmtpServer $Script:SNSemailSMTPServer -Credential $Credential365 -UseSsl -Subject $Script:SNSNewUserInfoEmailSubject -Encoding ([System.Text.Encoding]::UTF8)
            }

            Write-SNSLog "The account for $($newAccount.DisplayName)' is updated and ready for use."

            if ($Script:SNSAllUsersGroupName)
            {
                try
                {
                    # Add the user to the Distribution Group for all users with office 365 account.
                    Add-DistributionGroupMember -Identity $Script:SNSAllUsersGroupName -Member $AccountData.Identity -ErrorAction "Stop"
                }
                Catch
                {
                    if ($_.CategoryInfo.Reason -ne "MemberAlreadyExistsException")
                    {
                        Write-SNSLog -Level "Warn" "Could not add contact $($AccountData.Identity) to group $($Script:SNSAllUsersGroupName). Error $_"
                    }
                }
            }
        }
    }
    catch
    {
        Write-SNSLog -Level "Error" "Error during user account creation. Error $_"
    }
}

function Invoke-SNSUpdateAccount
{
    <#
    .SYNOPSIS
        Updates the account with information from Scoutnet, if the data is changed.
    
    .INPUTS
        None. You cannot pipe objects to Get-SNSUpdateAccount.

    .OUTPUTS
        None.
    #>
    param (
        [Parameter(Mandatory=$false, HelpMessage="User to update.")]
        [ValidateNotNull()]
        $AccountData,

        [Parameter(Mandatory=$False, HelpMessage="Credentials for api/group/memberlist.")]
        [ValidateNotNull()]
        [pscredential]$CredentialMemberlist

    )

    try
    {
        $GroupMemberlist = Get-SNSApiGroupMemberlist -Credential $CredentialMemberlist
        $MemberData = $GroupMemberlist.data[$AccountData.CustomAttribute1]
        $O365MemberData = Get-MsolUser -ObjectId $AccountData.ExternalDirectoryObjectId -ErrorAction "Stop"

        if ($MemberData)
        {
            $StreetAddress = $MemberData.address_1.value
            if ($MemberData.address_2.value)
            {
                $StreetAddress += " " + $MemberData.address_2.value
            }
            if ($MemberData.address_3.value)
            {
                $StreetAddress += " " + $MemberData.address_3.value
            }

            if ([string]::IsNullOrEmpty($StreetAddress))
            {
                $StreetAddress = ""
            }

            $AlternateEmailAddresses = $MemberData.email.value
            if ($O365MemberData.UserPrincipalName -like $AlternateEmailAddresses)
            {
                # Do not use the office 365 email as alternate email.
                # Try to use contact_alt_email.
                $AlternateEmailAddresses = $MemberData.contact_alt_email.value
                if ($O365MemberData.UserPrincipalName -like $AlternateEmailAddresses)
                {
                    # contact_alt_email not usable.
                    $AlternateEmailAddresses = ""
                }
            }

            if ([string]::IsNullOrEmpty($AlternateEmailAddresses))
            {
                $AlternateEmailAddresses = ""
            }

            if (($MemberData.first_name.value -notlike $O365MemberData.FirstName) -or 
                ($MemberData.last_name.value -notlike $O365MemberData.LastName) -or
                ($StreetAddress -notlike $O365MemberData.StreetAddress) -or
                ($MemberData.postcode.value -notlike $O365MemberData.PostalCode) -or
                ($MemberData.town.value -notlike $O365MemberData.City) -or
                ($MemberData.country.value -notlike $O365MemberData.Country) -or
                ($AlternateEmailAddresses -notlike $O365MemberData.AlternateEmailAddresses) -or
                ($MemberData.contact_mobile_phone.value -notlike $O365MemberData.MobilePhone))
            {
                if ([string]::IsNullOrEmpty($AlternateEmailAddresses))
                {
                    # Option -AlternateEmailAddresses expects an array. Create empty array.
                    $AlternateEmailAddresses = @()
                }

                # Update user data.
                Set-MsolUser -ObjectId $AccountData.ExternalDirectoryObjectId `
                    -DisplayName "$($MemberData.first_name.value) $($MemberData.last_name.value)" `
                    -FirstName $MemberData.first_name.value  `
                    -LastName $MemberData.last_name.value `
                    -StreetAddress $StreetAddress `
                    -PostalCode $MemberData.postcode.value `
                    -City $MemberData.town.value `
                    -Country $MemberData.country.value `
                    -AlternateEmailAddresses $AlternateEmailAddresses `
                    -MobilePhone $MemberData.contact_mobile_phone.value `
                    -ErrorAction Stop

                Write-SNSLog "User '$($AccountData.name)' uppdated with new info from Scoutnet."
                if ($Script:SNSAllUsersGroupName)
                {
                    try
                    {
                        # Add the user to the Distribution Group for all users with office 365 account.
                        Add-DistributionGroupMember -Identity $Script:SNSAllUsersGroupName -Member $AccountData.Identity -ErrorAction "Stop"
                    }
                    Catch
                    {
                        if ($_.CategoryInfo.Reason -ne "MemberAlreadyExistsException")
                        {
                            Write-SNSLog -Level "Warn" "Could not add contact $($AccountData.Identity) to group $($Script:SNSAllUsersGroupName). Error $_"
                        }
                    }
                }
            }
        }
        else
        {
            Write-SNSLog -Level "Warn" "Could not update user $($AccountData.name). Error user not found in scoutnet data."
        }
    }
    catch
    {
        Write-SNSLog -Level "Warn" "Could not update user $($AccountData.name). Error $_"    
    }
}

function Invoke-SNSDisableAccount
{
    <#
    .SYNOPSIS
        Disables the account and moves the user to SNSSecurityGroupScoutnetDisabledUsers.

    .DESCRIPTION
        The account is disabled and the user cannot login. The user is also moved
        to SNSSecurityGroupScoutnetDisabledUsers, and removed from SNSAllUsersGroup.
        No data is deleted and the licens is still in use.
    
    .INPUTS
        None. You cannot pipe objects to Get-SNSDisableAccount.

    .OUTPUTS
        None.
    #>
    param (
        [Parameter(Mandatory=$false, HelpMessage="User to disable.")]
        [ValidateNotNull()]
        $AccountData
    )

    try
    {
        Write-SNSLog "Disabling user '$($AccountData.name)' with Id '$($AccountData.ExternalDirectoryObjectId)'" 
        Set-MsolUser -ObjectId $AccountData.ExternalDirectoryObjectId -BlockCredential $true -ErrorAction "Stop"

        # Remove the user from the group of active users.
        $SecurityGroupScoutnet = Get-SNSSecurityGroupScoutnet
        Remove-MsolGroupMember -GroupObjectId $SecurityGroupScoutnet.ObjectId -GroupMemberObjectId $AccountData.ExternalDirectoryObjectId -ErrorAction "Stop"

        # Add the user to the group of disabled users.
        $SNSSecurityGroupScoutnetDisabledUsers = Get-SNSSecurityGroupScoutnetDisabledUsers
        Add-MsolGroupMember -GroupObjectId $SNSSecurityGroupScoutnetDisabledUsers.ObjectId -GroupMemberObjectId $AccountData.ExternalDirectoryObjectId -ErrorAction "Stop"

        if ($Script:SNSAllUsersGroupName)
        {
            try
            {
                # Remove the user from the Distribution Group for all users with office 365 account.
                Remove-DistributionGroupMember -Identity $Script:SNSAllUsersGroupName -Member $AccountData.Identity -Confirm:$Y -ErrorAction "Stop"
            }
            Catch
            {
                if ($_.CategoryInfo.Reason -ne "MemberAlreadyExistsException")
                {
                    Write-SNSLog -Level "Warn" "Could not remove contact $($AccountData.Identity) from group $($Script:SNSAllUsersGroupName). Error $_"
                }
            }
        }
    }
    catch
    {
        Write-SNSLog -Level "Error" "Could not disable user $($AccountData.name). Error $_"    
    }
}

function Invoke-SNSEnableAccount
{
    <#
    .SYNOPSIS
        Enables the account and moves the user to SNSSecurityGroupScoutnet.

    .DESCRIPTION
        The account is enabled and the user login. The user is also moved
        to SNSSecurityGroupScoutnet, and added to SNSAllUsersGroup.
        
        This is only valid for existing accounts.
    
    .INPUTS
        None. You cannot pipe objects to Get-SNSEnableAccount.

    .OUTPUTS
        None.
    #>
    param (
        [Parameter(Mandatory=$false, HelpMessage="User to enable.")]
        [ValidateNotNull()]
        $AccountData
    )

    try
    {
        Write-SNSLog "Enabling user '$($AccountData.name)' with Id '$($AccountData.ExternalDirectoryObjectId)'" 
        Set-MsolUser -ObjectId $AccountData.ExternalDirectoryObjectId -BlockCredential $false -ErrorAction "Stop"

        # Remove the user from the group of disabled users.
        $SNSSecurityGroupScoutnetDisabledUsers = Get-SNSSecurityGroupScoutnetDisabledUsers
        Remove-MsolGroupMember -GroupObjectId $SNSSecurityGroupScoutnetDisabledUsers.ObjectId -GroupMemberObjectId $AccountData.ExternalDirectoryObjectId -ErrorAction "Stop"

        # Add the user to the group of active users.
        $SecurityGroupScoutnet = Get-SNSSecurityGroupScoutnet
        Add-MsolGroupMember -GroupObjectId $SecurityGroupScoutnet.ObjectId -GroupMemberObjectId $AccountData.ExternalDirectoryObjectId -ErrorAction "Stop"

        if ($Script:SNSAllUsersGroupName)
        {
            # Add the user to the Distribution Group for all users with office 365 account.
            Add-DistributionGroupMember -Identity $Script:SNSAllUsersGroupName -Member $AccountData.Identity -ErrorAction "Stop"
        }
    }
    catch
    {
        Write-SNSLog -Level "Error" "Could not enable user $($AccountData.name). Error $_"    
    }
}

function Get-SNSSoutnetLeaders
{
    [OutputType([System.Collections.ArrayList])]
    param (
        [Parameter(Mandatory=$False, HelpMessage="Credentials for api/group/customlists.")]
        [ValidateNotNull()]
        [pscredential]$CredentialCustomlists,

        [Parameter(Mandatory=$False, HelpMessage="Credentials for api/group/memberlist.")]
        [ValidateNotNull()]
        [pscredential]$CredentialMemberlist,

        [Parameter(Mandatory=$False, HelpMessage="Maillist to process.")]
        $MailListId
    )

    $GroupMemberlist = Get-SNSApiGroupMemberlist -Credential $CredentialMemberlist
    $MemberListScoutnet = [System.Collections.ArrayList]::new()
    # Fetch members from scoutnet from the selected maillists.
    foreach($ListId in $MailListId)
    {
        $SelectedList = Get-SNSApiGroupCustomlist -Credential $CredentialCustomlists -listid $ListId
        $SelectedList.data.keys | ForEach-Object {[void]$MemberListScoutnet.Add($_)}
    }

    if ($MemberListScoutnet.Count -eq 0)
    {
        # No members found in list or list not specified.
        # Select users with a role in group or unit and add to MemberListScoutnet
        foreach ($user in $GroupMemberlist.data.values)
        {
            if ( ![string]::IsNullOrWhiteSpace($user.group_role.value) -or ![string]::IsNullOrWhiteSpace($user.unit_role.value))
            {
                Write-SNSLog "Found member '$($user.first_name.value) $($user.last_name.value)' ($($user.member_no.value)) with roles '$($user.group_role.value)', '$($user.unit_role.value)'"
                [void]$MemberListScoutnet.Add($user.member_no.value)
            }
        }
    }

    return $MemberListScoutnet
}

function Get-SNSUsersInSecurityGroupScoutnet
{
    [OutputType([System.Collections.ArrayList])]
    param (
        [Parameter(Mandatory=$False, HelpMessage="List of office 365 users.")]
        [ValidateNotNull()]
        $allOffice365Users
    )
    return Get-SNSUsersInSecurityGroup -allOffice365Users $allOffice365Users -Name $Script:SNSyncGroupName -Description $Script:SNSyncGroupDescription
}

function Get-SNSUsersInSecurityGroupScoutnetDisabledUsers
{
    [OutputType([System.Collections.ArrayList])]
    param (
        [Parameter(Mandatory=$False, HelpMessage="List of office 365 users.")]
        [ValidateNotNull()]
        $allOffice365Users
    )
    return Get-SNSUsersInSecurityGroup -allOffice365Users $allOffice365Users -Name $Script:SNSyncGroupDisabledUsersName -Description $Script:SNSyncGroupDisabledUsersDescription
}

function Get-SNSSecurityGroupScoutnet
{
    return Get-SNSSecurityGroup -Name $Script:SNSyncGroupName -Description $Script:SNSyncGroupDescription
}

function Get-SNSSecurityGroupScoutnetDisabledUsers
{
    return Get-SNSSecurityGroup -Name $Script:SNSyncGroupDisabledUsersName -Description $Script:SNSyncGroupDisabledUsersDescription
}

function Get-SNSUsersInSecurityGroup
{
    [OutputType([System.Collections.ArrayList])]
    param (
        [Parameter(Mandatory=$False, HelpMessage="Group name")]
        [ValidateNotNull()]
        [string]$Name,

        [Parameter(Mandatory=$False, HelpMessage="Group description.")]
        [ValidateNotNull()]
        [string]$Description,

        [Parameter(Mandatory=$False, HelpMessage="List of office 365 users.")]
        [ValidateNotNull()]
        $allOffice365Users
    )

    $securityGroup = Get-SNSSecurityGroup -Name $Name -Description $Description

    $securityGroupMembers = [System.Collections.ArrayList]::new()
    if ($securityGroup)
    {
        $GroupMembers = Get-MsolGroupMember -GroupObjectId $securityGroup.ObjectId
        # Get the mailbox info fore each group member.
        foreach ($GroupMember in $GroupMembers)
        {
            $member = $allOffice365Users | Where-Object -FilterScript {$_.ExternalDirectoryObjectId -eq $GroupMember.ObjectId}
            if ($member)
            {
                [void]$securityGroupMembers.Add($member)
            }
        }
    }

    return $securityGroupMembers
}

function Get-SNSSecurityGroup
{
    [OutputType([System.Collections.ArrayList])]
    param (
        [Parameter(Mandatory=$False, HelpMessage="Group name")]
        [ValidateNotNull()]
        [string]$Name,

        [Parameter(Mandatory=$False, HelpMessage="Group description.")]
        [ValidateNotNull()]
        [string]$Description

    )

    $securityGroups = Get-MsolGroup -SearchString $Name -ErrorAction "Stop"

    $securityGroup=$null
    foreach($group in $securityGroups)
    {
        if ($group.DisplayName -like $Name)
        {
            $securityGroup = $group
            break
        }
    }

    if (!$securityGroup)
    {
        Write-SNSLog -Level "Warn" "Security group $name is not found. Creating the group."
        $securityGroup = New-MsolGroup -DisplayName $Name -Description $Description -ErrorAction "Stop"
    }

    return $securityGroup
}
