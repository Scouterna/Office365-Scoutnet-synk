function Invoke-SNSUppdateOffice365User
{
    <#
    .SYNOPSIS
        Main function for syncronisation of office 365 users with scoutnet.

    .DESCRIPTION
        Fetches the distribution groups members and updates corressponding distribution groups based on the info from scoutnet.
        As all members of an distribution group must be present in exchange as user or contact, contacts will be created for external addresses.

    .INPUTS
        None. You cannot pipe objects to Invoke-SNSUppdateOffice365User.

    .OUTPUTS
        None.
    
    .LINK
        https://github.com/scouternasetjanster/Office365-Scoutnet-synk

    .PARAMETER Configuration
        Configuration to use. If not specified the cached configuration will be used.
    #>
    [CmdletBinding(HelpURI = 'https://github.com/scouternasetjanster/Office365-Scoutnet-synk',
                PositionalBinding = $False)]
    param (
        [Parameter(Mandatory=$False, HelpMessage="Configuration to use. If not specified the cached configuration will be used.")]
        $Configuration
    )

    if ($Configuration)
    {
        $Script:SNSConf = $Configuration
    }

    if (!$Script:SNSConf)
    {
        throw "No configuration specified. Please provide a configuration!"
    }

    try
    {
        Get-MsolDomain -ErrorAction Stop > $null
    }
    catch
    {
        Write-SNSLog "Connecting to Office 365..."
        try
        {
            Connect-MsolService -Credential $Script:SNSConf.Credential365 -ErrorAction Stop
        }
        catch
        {
            Write-SNSLog -Level "Error" "Could not connect to Office 365. Error $_"
            throw
        }
    }

    Write-SNSLog "Start of user account update"
    $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Script:SNSConf.Credential365 -Authentication Basic -AllowRedirection
    Import-PSSession $ExchangeSession -AllowClobber -CommandName Set-MailContact,Set-Mailbox,Get-Mailbox,Remove-DistributionGroupMember,Add-DistributionGroupMember,Get-DistributionGroupMember,Get-DistributionGroup,Set-MailboxMessageConfiguration,Set-MailboxAutoReplyConfiguration > $null

    [System.Collections.ArrayList]$allOffice365Users = Get-Mailbox -RecipientTypeDetails "UserMailbox"
    [System.Collections.ArrayList]$SecurityGroupScoutnet = Get-SNSUsersInSecurityGroupScoutnet -allOffice365Users $allOffice365Users
    [System.Collections.ArrayList]$UsersInSecurityGroupScoutnetDisabledUsers = Get-SNSUsersInSecurityGroupScoutnetDisabledUsers -allOffice365Users $allOffice365Users
    [System.Collections.ArrayList]$MemberListScoutnet = Get-SNSSoutnetLeaders -CredentialCustomlists $Script:SNSConf.CredentialCustomlists -CredentialMemberlist $Script:SNSConf.CredentialMemberlist -MailListId $Script:SNSConf.UserSyncMailListId
    $GroupMemberlist = Get-SNSApiGroupMemberlist -Credential $Script:SNSConf.CredentialMemberlist

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
            elseif ($isInUsersInSecurityGroupScoutnetDisabledUsers)
            {
                # Returning member. Enable and update the office 365 account.
                [void]$MembersToActivate.Add($isInUsersInSecurityGroupScoutnetDisabledUsers)
                $UsersInSecurityGroupScoutnetDisabledUsers.RemoveAt($UsersInSecurityGroupScoutnetDisabledUsers.IndexOf($isInUsersInSecurityGroupScoutnetDisabledUsers))
            }
            else
            {
                Write-SNSLog "Account not to update: '$($account.Name)'"
            }

            # Remove from list of all accounts.
            $allOffice365Users.RemoveAt($allOffice365Users.IndexOf($account))
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
            Invoke-SNSCreateUserAndUpdateUserData -memberData $NewMembers
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

    Write-SNSLog "Number of disabled accounts: $($UsersInSecurityGroupScoutnetDisabledUsers.Count)"
    if ($UsersInSecurityGroupScoutnetDisabledUsers.Count -gt 0)
    {
        $UsersInSecurityGroupScoutnetDisabledUsers | ForEach-Object {
            Write-SNSLog "Disabled account '$($_.Name)'"
            $allOffice365Users.RemoveAt($allOffice365Users.IndexOf($_))
        }
    }

    Write-SNSLog "Number of accounts to disable: $($SecurityGroupScoutnet.Count)"
    $SecurityGroupScoutnet | ForEach-Object {
        Invoke-SNSDisableAccount -AccountData $_
        $allOffice365Users.RemoveAt($allOffice365Users.IndexOf($_))
    }

    Write-SNSLog "Number of accounts to check for update: $($MembersToUpdate.Count)"
    $MembersToUpdate | ForEach-Object {
        Invoke-SNSUpdateAccount -AccountData $_ -Credential $Script:SNSConf.CredentialMemberlist
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
        None. You cannot pipe objects to Invoke-SNSCreateUserAndUpdateUserData.

    .OUTPUTS
        None.
    #>
    param (
        [Parameter(Mandatory=$false, HelpMessage="List of members to create.")]
        [ValidateNotNull()]
        $memberData
    )

    $SecurityGroupScoutnet = Get-SNSSecurityGroupScoutnet

    $newAccounts = [ordered]@{}
    $LastAccountUserPrincipalName=$null
    foreach($MemberData in $memberData)
    {
#region Generate the new UserPrincipalName
        $DisplayName = "$($MemberData.first_name.value) $($MemberData.last_name.value)"
        $UserName = "$($MemberData.first_name.value).$($MemberData.last_name.value)".ToLower()
        # Convert UTF encoded names and create corresponding ASCII version.
        $UserName = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($UserName))
    
        $UserPrincipalName = "$($UserName)@$($Script:SNSConf.DomainName)"
    
        $office365User  = Get-MsolUser -UserPrincipalName $UserPrincipalName -ErrorAction SilentlyContinue
    
        if ($office365User)
        {
            # Mailaddress alredy exists. Try with an extra number.
            For ($cnt=1; $cnt -le 5; $cnt++)
            {
                $UserPrincipalName = "$($UserName).$($cnt)@$($Script:SNSConf.DomainName)"
                $office365User  = Get-MsolUser -UserPrincipalName $UserPrincipalName -ErrorAction SilentlyContinue
                if (!$office365User)
                {
                    # Mailaddress not used. Uset i!
                    break
                }
            }
        }
#endregion

        if (!$office365User)
        {
            try
            {
#region Fetch data for the new user
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
                if ($UserPrincipalName -like $AlternateEmailAddresses)
                {
                    # Do not use the office 365 email as alternate email.
                    # Try to use contact_alt_email.
                    $AlternateEmailAddresses = $MemberData.contact_alt_email.value
                    if ($UserPrincipalName -like $AlternateEmailAddresses)
                    {
                        # contact_alt_email not usable.
                        $AlternateEmailAddresses = ""
                    }
                }
    
                if ([string]::IsNullOrEmpty($AlternateEmailAddresses))
                {
                    # Option -AlternateEmailAddresses expects an array. Create empty array if there is no AlternateEmailAddresses.
                    $AlternateEmailAddresses = @()
                }
#endregion

#region Create the new user.
                $newAccount = New-MsolUser -UserPrincipalName $UserPrincipalName -DisplayName $DisplayName `
                    -FirstName $MemberData.first_name.value  `
                    -LastName $MemberData.last_name.value `
                    -StreetAddress $StreetAddress `
                    -PostalCode $MemberData.postcode.value `
                    -City $MemberData.town.value `
                    -Country $MemberData.country.value `
                    -AlternateEmailAddresses $AlternateEmailAddresses `
                    -MobilePhone $MemberData.contact_mobile_phone.value `
                    -PreferredLanguage $Script:SNSConf.PreferredLanguage `
                    -UsageLocation $Script:SNSConf.UsageLocation `
                    -LicenseAssignment $Script:SNSConf.LicenseAssignment `
                    -LicenseOptions $Script:SNSConf.LicenseOptions `
                    -ErrorAction Stop

                Write-SNSLog "User '$($newAccount.UserPrincipalName)' added for member id '$($MemberData.member_no.value)'."
#endregion

#region Add the user to the group of active users.
                $newAccounts.Add($MemberData.member_no.value, @($MemberData, $newAccount))
                $LastAccountUserPrincipalName = $newAccount.UserPrincipalName
                try
                {
                    Add-MsolGroupMember -GroupObjectId $SecurityGroupScoutnet.ObjectId -GroupMemberObjectId $newAccount.ObjectId -ErrorAction "Stop"
                }
                Catch
                {
                    Write-SNSLog -Level "Warn" "Could not add contact '$($newAccount.DisplayName)' to group $($Script:SNSConf.SyncGroupName). Error $_"
                }    
#endregion
            }
            catch
            {
                Write-SNSLog -Level "Error" "Could not create user '$($UserPrincipalName)' for member '$DisplayName' with id '$($MemberData.member_no.value)'. Error $_"
            }
        }
        else
        {
            Write-SNSLog -Level "Error" "Mailaddress $($UserPrincipalName) is alredy in use. Can not add a user for member '$DisplayName' with id '$($MemberData.member_no.value)'"
        }
    }

#region Wait for mailbox creation.
    $maxDateTimeout = (Get-Date).AddSeconds($Script:SNSConf.WaitMailboxCreationMaxTime)
    if (!$LastAccountUserPrincipalName)
    {
        Write-SNSLog "No user created."
        return
    }
    else
    {
        $doLoop = $true
        Write-SNSLog "Wait for the mailbox for the new users to be created. It can take som time..."
        Start-Sleep -s $Script:SNSConf.WaitMailboxCreationPollTime
    }

    try
    {
        while($doLoop)
        {
            Start-Sleep -s $Script:SNSConf.WaitMailboxCreationPollTime
            try
            {
                Get-Mailbox -Identity $LastAccountUserPrincipalName -RecipientTypeDetails "UserMailbox" -ErrorAction "Stop" > $null
                # Mailboxes is created.
                $doLoop=$false
                break
            }
            catch
            {
                Write-SNSLog "Still waiting..."
            }

            if ($maxDateTimeout -lt (Get-Date))
            {
                # timeout limit reached so exception
                $msg = "The creation of user mailboxes did not"
                $msg += "complete within the timeout limit of "
                $msg += "$($Script:SNSConf.WaitMailboxCreationMaxTime) seconds, so polling "
                $msg += "for mailbox creation was halted."
                throw ($msg)
            }
        }
#endregion

        foreach($newAccountId in $newAccounts.Keys)
        {
#region Update the mailbox configuration for the new account.
            Write-SNSLog "Updating account '$($newAccount.DisplayName)'."
            $member =  $newAccounts[$newAccountId][0]
            $newAccount =  $newAccounts[$newAccountId][1]

            $SignatureHtml = $Script:SNSConf.SignatureHtml -Replace "<DisplayName>", $newAccount.DisplayName
            $SignatureText = $Script:SNSConf.SignatureText -Replace "<DisplayName>", $newAccount.DisplayName

            try
            {
                Set-Mailbox -Identity $newAccount.UserPrincipalName -CustomAttribute1 $newAccountId -ErrorAction "Stop"
                Set-MailboxMessageConfiguration $newAccount.UserPrincipalName -IsReplyAllTheDefaultResponse $false `
                    -SignatureHtml $SignatureHtml -SignatureText $SignatureText -AutoAddSignature $true -AutoAddSignatureOnMobile $true `
                    -AutoAddSignatureOnReply $true -SignatureTextOnMobile $SignatureText -ErrorAction "Stop"                    
            }
            catch
            {
                Write-SNSLog -Level "Error" "Could not update mailbox for user '$($newAccount.UserPrincipalName)'. Error $_"                
            }
#endregion

#region Send e-mail to the user with the new password and account info
            $emailFromAddress = $Script:SNSConf.EmailFromAddress
            if ([string]::IsNullOrWhiteSpace($emailFromAddress))
            {
                # Use the login credential as from address.
                $emailFromAddress = $Script:SNSConf.Credential365.UserName
            }

            if ([string]::IsNullOrWhiteSpace($member.email.value))
            {
                Write-SNSLog "No valid email address in scoutnet for user '$($newAccount.UserPrincipalName)'. Notify the user about the new account."
            }
            else
            {
                try
                {
                    # Send e-mail to the user with the new password and account info. The password must be replaced at first login.
                    $NewUserEmailText = $Script:SNSConf.NewUserEmailText -Replace "<DisplayName>", $newAccount.DisplayName -Replace "<Password>", $newAccount.Password -Replace "<UserPrincipalName>", $newAccount.UserPrincipalName
                    Send-MailMessage -From $emailFromAddress -to $member.email.value -Body $NewUserEmailText `
                        -SmtpServer $Script:SNSConf.EmailSMTPServer -Port $Script:SNSConf.SmtpPort -Credential $Script:SNSConf.Credential365 `
                        -UseSsl -Subject $Script:SNSConf.NewUserEmailSubject -Encoding ([System.Text.Encoding]::UTF8) -ErrorAction "Stop"
                }
                Catch
                {
                    Write-SNSLog -Level "Warn" "Could not send email to $($member.email.value). Error $_"
                }
            }
#endregion

#region Extra info mail requested. Send it to the new account.
            if (![string]::IsNullOrWhiteSpace($Script:SNSConf.NewUserInfoEmailText) -and ![string]::IsNullOrWhiteSpace($Script:SNSConf.NewUserInfoEmailSubject))
            {
                try
                {
                    # Extra info mail requested. Send it to the new account.
                    $NewUserEmailText = $Script:SNSConf.NewUserInfoEmailText -Replace "<DisplayName>", $newAccount.DisplayName -Replace "<UserPrincipalName>", $newAccount.UserPrincipalName
                    Send-MailMessage -From $emailFromAddress -to  $newAccount.UserPrincipalName -Body $NewUserEmailText `
                        -SmtpServer $Script:SNSConf.EmailSMTPServer -Port $Script:SNSConf.SmtpPort -Credential $Script:SNSConf.Credential365 `
                        -UseSsl -Subject $Script:SNSConf.NewUserInfoEmailSubject -Encoding ([System.Text.Encoding]::UTF8) -ErrorAction "Stop"
                }
                Catch
                {
                    Write-SNSLog -Level "Warn" "Could not send email to $($newAccount.UserPrincipalName). Error $_"
                }
            }
#endregion

#region Add the user to the Distribution Group for all users with office 365 account.
            if ($Script:SNSConf.AllUsersGroupName)
            {
                try
                {
                    # Add the user to the Distribution Group for all users with office 365 account.
                    Add-DistributionGroupMember -Identity $Script:SNSConf.AllUsersGroupName -Member $newAccount.UserPrincipalName -ErrorAction "Stop"
                    Write-SNSLog "The account '$($newAccount.DisplayName)' is added to distribution group '$($Script:SNSConf.AllUsersGroupName)'"
                }
                Catch
                {
                    if ($_.CategoryInfo.Reason -ne "MemberAlreadyExistsException")
                    {
                        Write-SNSLog -Level "Warn" "Could not add contact $($newAccount.UserPrincipalName) to group $($Script:SNSConf.AllUsersGroupName). Error $_"
                    }
                }
            }
            Write-SNSLog "The account for '$($newAccount.DisplayName)' is updated and ready for use."
#endregion
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
                if ($Script:SNSConf.AllUsersGroupName)
                {
                    try
                    {
                        # Add the user to the Distribution Group for all users with office 365 account.
                        Add-DistributionGroupMember -Identity $Script:SNSConf.AllUsersGroupName -Member $AccountData.Identity -ErrorAction "Stop"
                    }
                    Catch
                    {
                        if ($_.CategoryInfo.Reason -ne "MemberAlreadyExistsException")
                        {
                            Write-SNSLog -Level "Warn" "Could not add contact $($AccountData.Identity) to group $($Script:SNSConf.AllUsersGroupName). Error $_"
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
        Any mail forward is disabled.
        If the setting DisabledAccountsAutoReplyText contains any message, the message
        is set as autoreply message.

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

        # Mark the account as hidden so the user is not shown in the global address book.
        # Remove any forwarders enabled by the user.
        Set-Mailbox -Identity $AccountData.UserPrincipalName -HiddenFromAddressListsEnabled $True -ForwardingAddress $null -ForwardingSmtpAddress $null -DeliverToMailboxAndForward $false -ErrorAction "Stop"

        # Remove the user from the group of active users.
        $SecurityGroupScoutnet = Get-SNSSecurityGroupScoutnet
        Remove-MsolGroupMember -GroupObjectId $SecurityGroupScoutnet.ObjectId -GroupMemberObjectId $AccountData.ExternalDirectoryObjectId -ErrorAction "Stop"

        # Add the user to the group of disabled users.
        $SNSSecurityGroupScoutnetDisabledUsers = Get-SNSSecurityGroupScoutnetDisabledUsers
        Add-MsolGroupMember -GroupObjectId $SNSSecurityGroupScoutnetDisabledUsers.ObjectId -GroupMemberObjectId $AccountData.ExternalDirectoryObjectId -ErrorAction "Stop"

        if ($Script:SNSConf.AllUsersGroupName)
        {
            try
            {
                # Remove the user from the Distribution Group for all users with office 365 account.
                Remove-DistributionGroupMember -Identity $Script:SNSConf.AllUsersGroupName -Member $AccountData.Identity -Confirm:$Y -ErrorAction "Stop"
            }
            Catch
            {
                if ($_.CategoryInfo.Reason -ne "MemberAlreadyExistsException")
                {
                    Write-SNSLog -Level "Warn" "Could not remove contact $($AccountData.Identity) from group $($Script:SNSConf.AllUsersGroupName). Error $_"
                }
            }
        }

        if (![string]::IsNullOrWhiteSpace($Script:SNSConf.DisabledAccountsAutoReplyText))
        {
            $DisabledAccountsAutoReplyText = $Script:SNSConf.DisabledAccountsAutoReplyText -Replace "<DisplayName>", $AccountData.DisplayName
            try
            {
                Set-MailboxAutoReplyConfiguration -Identity $AccountData.UserPrincipalName -AutoReplyState Enabled -ExternalAudience All -InternalMessage $DisabledAccountsAutoReplyText -ExternalMessage $DisabledAccountsAutoReplyText
            }
            catch
            {
                Write-SNSLog -Level "Error" "Could not set autoreply message for user $($AccountData.name). Error $_"
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
        Any autoreply is disabled.
        
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

        # Add the user to the global address book.
        Set-Mailbox -Identity $AccountData.UserPrincipalName -HiddenFromAddressListsEnabled $false -ErrorAction "Stop"

        # Remove the user from the group of disabled users.
        $SNSSecurityGroupScoutnetDisabledUsers = Get-SNSSecurityGroupScoutnetDisabledUsers
        Remove-MsolGroupMember -GroupObjectId $SNSSecurityGroupScoutnetDisabledUsers.ObjectId -GroupMemberObjectId $AccountData.ExternalDirectoryObjectId -ErrorAction "Stop"

        # Add the user to the group of active users.
        $SecurityGroupScoutnet = Get-SNSSecurityGroupScoutnet
        Add-MsolGroupMember -GroupObjectId $SecurityGroupScoutnet.ObjectId -GroupMemberObjectId $AccountData.ExternalDirectoryObjectId -ErrorAction "Stop"

        if ($Script:SNSConf.AllUsersGroupName)
        {
            # Add the user to the Distribution Group for all users with office 365 account.
            Add-DistributionGroupMember -Identity $Script:SNSConf.AllUsersGroupName -Member $AccountData.Identity -ErrorAction "Stop"
        }

        try
        {
            Set-MailboxAutoReplyConfiguration -Identity $AccountData.UserPrincipalName -AutoReplyState Disabled -InternalMessage $null -ExternalMessage $null
        }
        catch
        {
            Write-SNSLog -Level "Error" "Could not disable the autoreply message for user $($AccountData.name). Error $_"
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
    return Get-SNSUsersInSecurityGroup -allOffice365Users $allOffice365Users -Name $Script:SNSConf.SyncGroupName -Description $Script:SNSConf.SyncGroupDescription
}

function Get-SNSUsersInSecurityGroupScoutnetDisabledUsers
{
    [OutputType([System.Collections.ArrayList])]
    param (
        [Parameter(Mandatory=$False, HelpMessage="List of office 365 users.")]
        [ValidateNotNull()]
        $allOffice365Users
    )
    return Get-SNSUsersInSecurityGroup -allOffice365Users $allOffice365Users -Name $Script:SNSConf.SyncGroupDisabledUsersName -Description $Script:SNSConf.SyncGroupDisabledUsersDescription
}

function Get-SNSSecurityGroupScoutnet
{
    return Get-SNSSecurityGroup -Name $Script:SNSConf.SyncGroupName -Description $Script:SNSConf.SyncGroupDescription
}

function Get-SNSSecurityGroupScoutnetDisabledUsers
{
    return Get-SNSSecurityGroup -Name $Script:SNSConf.SyncGroupDisabledUsersName -Description $Script:SNSConf.SyncGroupDisabledUsersDescription
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
