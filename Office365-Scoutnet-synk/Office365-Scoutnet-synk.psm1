#Requires -Version 5.1

[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"

#region import everything we need

. $PSScriptRoot\SNSConfiguration.ps1
. $PSScriptRoot\Write-SNSLog.ps1
. $PSScriptRoot\ConvertTo-SNSJSONHash.ps1
. $PSScriptRoot\Receive-SNSApiJson.ps1
. $PSScriptRoot\Get-SNSApiGroupCustomlist.ps1
. $PSScriptRoot\Get-SNSMaillistInfo.ps1
. $PSScriptRoot\Get-SNSUserEmail.ps1
. $PSScriptRoot\Get-SNSExchangeMailListMember.ps1
. $PSScriptRoot\Get-SNSApiGroupMemberlist.ps1
. $PSScriptRoot\Invoke-SNSUppdateOffice365User.ps1

#endregion

# Configuration holder.
$script:SNSConf=[SNSConfiguration]::new()
Export-ModuleMember -Variable SNSConf

function Add-Office365User
{
    [OutputType([bool])]
    param (
        [Parameter(Mandatory=$True, HelpMessage="List with all Office365 users.")]
        [ValidateNotNull()]
        $allOffice365Users,

        [Parameter(Mandatory=$True, HelpMessage="Scoutnet Id for member to add.")]
        [ValidateNotNull()]
        $member,

        [Parameter(Mandatory=$True, HelpMessage="Data for the member from Scoutnet.")]
        [ValidateNotNull()]
        $MemberData,

        [Parameter(Mandatory=$True, HelpMessage="Name of group to add member in.")]
        [ValidateNotNull()]
        $distGroupName,

        [Parameter(HelpMessage="Do warn about missing Office 365 user.")]
        [switch] $doWarn
        )

    $memberSearchStr = "*$member"
    $recipient = $allOffice365Users | Where-Object {$_.CustomAttribute1 -like $memberSearchStr}

    $userAdded = $False

    if ($recipient)
    {
        try
        {
            Write-SNSLog "Adding member $($recipient.DisplayName) with id $($recipient.Id) to distribution group $distGroupName"
            Add-DistributionGroupMember -Identity $distGroupName -Member $recipient.Id  -ErrorAction "stop"
            $userAdded = $True
        }
        catch
        {
            if ($_.CategoryInfo.Reason -ne "MemberAlreadyExistsException")
            {
                Write-SNSLog -Level "Warn" "Could not add contact $($recipient.DisplayName) to $distGroupName. Error $_"
            }
        }
    }
    elseif ($doWarn)
    {
        Write-SNSLog -Level "Warn" "Member $($MemberData.first_name) $($MemberData.last_name) not found in office 365. Please make sure that CustomAttribute1 contains Scoutnet Id for the user."
    }
    return $userAdded
}


function Add-MailContactToList
{
    param (
        [ValidateNotNull()]
        $Epost,

        [ValidateNotNull()]
        $DisplayName,

        [ValidateNotNull()]
        $DistGroupName
        )

    $ExistingMailContact = get-recipient $Epost -ErrorAction "SilentlyContinue"
    if ($null -eq $ExistingMailContact)
    {
        Write-SNSLog "Creating Contact $Epost for $DisplayName"
        try
        {
            New-MailContact -Name $Epost -ExternalEmailAddress $Epost -ErrorAction "stop" > $null

            # Set the name of the member in the company field. This is visibel in Office 365 admin console.
            Set-Contact -Identity $Epost -Company "$DisplayName"
            Set-MailContact -Identity $Epost -HiddenFromAddressListsEnabled $true
        }
        Catch
        {
            Write-SNSLog -Level "Warn" "Could not create mail contact with address $Epost. Error $_"
        }
    }

    try
    {
        Add-DistributionGroupMember -Identity $DistGroupName -Member $Epost -ErrorAction "stop"
        Write-SNSLog "Adding contact $Epost for $DisplayName to distribution group $DistGroupName"
    }
    Catch
    {
        if ($_.CategoryInfo.Reason -ne "MemberAlreadyExistsException")
        {
            Write-SNSLog -Level "Warn" "Could not add contact $Epost for $DisplayName. Error $_"
        }
    }
}


function SNSUpdateExchangeDistributionGroups
{
    <#
    .SYNOPSIS
        Updates exchange distribution groups based on scoutnet maillists.

    .DESCRIPTION
        Fetches the distribution groups members and updates corressponding distribution groups based on the info from scoutnet.
        As all members of an distribution group must be present in exchange as user or contact, contacts will be created for external addresses.

    .INPUTS
        None. You cannot pipe objects to SNSUpdateExchangeDistributionGroups.

    .OUTPUTS
        None.
    
    .LINK
        https://github.com/scouternasetjanster/Office365-Scoutnet-synk

    .PARAMETER ValidationHash
        Hash value used to validate if scoutnet is updated.

    .PARAMETER Configuration
        Configuration to use. If not specified the cached configuration will be used.
    #>
    [CmdletBinding(HelpURI = 'https://github.com/scouternasetjanster/Office365-Scoutnet-synk',
                PositionalBinding = $False)]
    [OutputType([string])]
    param (
        [Parameter(Mandatory=$True, HelpMessage="Hash value used to validate if Scoutnet is updated.")]
        [ValidateNotNull()]
        [string]$ValidationHash,

        [Parameter(Mandatory=$False, HelpMessage="Configuratin to use. If not specified the cached configuration will be used.")]
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

    $MailListSettings = $Script:SNSConf.MailListSettings

    # Fetch maillist info from scoutnet.
    $CustomLists, $CustomListsHash = Get-SNSMaillistInfo -CredentialCustomlists $Script:SNSConf.CredentialCustomlists -MaillistsIds $MailListSettings.values

    # Fetch all members and their mailaddresses.
    $allMailAddresses, $allMailAddressesHash = Get-SNSUserEmail -CredentialMemberlist $Script:SNSConf.CredentialMemberlist

    $MailListSettingsHash = ("{0:X8}" -f (($MailListSettings | ConvertTo-Json).GetHashCode()))

    $NewValidationHash = "0x{0}{1}{2}" -f ($CustomListsHash, $allMailAddressesHash, $MailListSettingsHash)

    Write-SNSLog "Saved validation hash: $ValidationHash new value $NewValidationHash"

    if ($ValidationHash -eq $NewValidationHash)
    {
        Write-SNSLog "Scoutnet is not updated. No update will be done."
    }
    else
    {
        Write-SNSLog "Scoutnet is updated. Starting to update the distribution groups."

        $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Script:SNSConf.Credential365 -Authentication Basic -AllowRedirection
        Import-PSSession $ExchangeSession -AllowClobber -CommandName Set-MailContact,Set-Contact,New-MailContact,Remove-MailContact,Remove-DistributionGroupMember,Get-Recipient,Add-DistributionGroupMember,Get-Mailbox > $null

        $otherMailListsMembers, $mailListsToProcessMembers = Get-SNSExchangeMailListMember -Credential365 $Script:SNSConf.Credential365 -ExchangeSession $ExchangeSession -Maillists $MailListSettings.Keys

        # Clean the distribution groups first.
        Write-SNSLog " "
        Write-SNSLog "Removing contacts in distribution groups"
        foreach ($distGroupName in $MailListSettings.keys)
        {
            Write-SNSLog "Remove contacts in distribution group $($distGroupName)"
            $distGroupMembers= Get-DistributionGroupMember -Identity $distGroupName
            foreach ($medlem in $distGroupMembers)
            {
                Remove-DistributionGroupMember -Identity $distGroupName -Member $medlem.Identity -Confirm:$Y
            }
        }
        Write-SNSLog "Removed contacts in distribution groups"

        $allOffice365Users = Get-Mailbox -RecipientTypeDetails "UserMailbox"

        foreach ($distGroupName in $MailListSettings.keys)
        {
            $distGroupId = $MailListSettings[$distGroupName].scoutnet_list_id
            Write-SNSLog " "
            Write-SNSLog "Adding Contacts in distribution group $($distGroupName)"

            $scouter_synk_option = $MailListSettings[$distGroupName].scouter_synk_option
            $scouter_synk_option = $scouter_synk_option.ToLower()

            Write-SNSLog "Scouter synk option is '$scouter_synk_option' for distribution group $distGroupName"

            # Fetch the mail addresses to add for scouter.
            $listData = $CustomLists[$distGroupId].scouter
            foreach ($member in $listData)
            {
                if ($null -eq $member)
                {
                    Write-SNSLog -Level "Warn" "Mail list $distGroupName contained a null value."
                    continue
                }
                if (-not $AllMailAddresses.ContainsKey($member))
                {
                    Write-SNSLog -Level "Warn" "Id '$member' from $distGroupName not found in list of all members."
                    continue
                }

                $MemberData = $AllMailAddresses[$member]
                $displayName = "$($MemberData.first_name) $($MemberData.last_name)"

                if ($MailListSettings[$distGroupName].ignore_user)
                {
                    if ($MailListSettings[$distGroupName].ignore_user.Contains($member))
                    {
                        Write-SNSLog "Ignoring '$($displayName)' with Id '$member' for $distGroupName."
                        continue
                    }
                }

                # Get valid addresses to add based on list setting.
                $mailaddresses = [System.Collections.ArrayList]::new()
                if ($scouter_synk_option.contains("p"))
                {
                    [void]$mailaddresses.Add($MemberData.primary_email)
                }
                elseif ($scouter_synk_option.contains("a"))
                {
                    [void]$mailaddresses.Add($MemberData.alt_email)
                }
                elseif ($scouter_synk_option.contains("f"))
                {
                    $MemberData.contacts_addresses | ForEach-Object {[void]$mailaddresses.Add($_)}
                }
                else
                {
                    $MemberData.mailaddresses | ForEach-Object {[void]$mailaddresses.Add($_)}
                }
                $mailaddresses = $mailaddresses | Sort-Object -Unique

                $AddMemberOffice365Address = $False
                $AddMemberScoutnetAddress = $True
                $AddMemberOffice365AddressTryFirst = $False # First try to add office 365 address. If that fails add the scoutnet version.

                if ($scouter_synk_option.contains("@"))
                {
                    $AddMemberScoutnetAddress = $False
                    $AddMemberOffice365Address = $True
                }
                elseif ($scouter_synk_option.contains("&"))
                {
                    $AddMemberScoutnetAddress = $True
                    $AddMemberOffice365Address = $True
                }

                if ($scouter_synk_option.contains("t"))
                {
                    $AddMemberScoutnetAddress = $False
                    $AddMemberOffice365AddressTryFirst = $True
                    $AddMemberOffice365Address = $True
                }

                if ($mailaddresses.Length -eq 0)
                {
                    Write-SNSLog -Level "Warn" "No email addresses found for $($displayName)."
                    continue
                }

                $AddMemberScoutnetAddress = $AddMemberScoutnetAddress
                if ($AddMemberOffice365Address)
                {
                    $result = Add-Office365user -allOffice365Users $allOffice365Users -Member $member -MemberData $AllMailAddresses[$member] -distGroupName $distGroupName
                    if (!$result)
                    {
                        if ($AddMemberOffice365AddressTryFirst)
                        {
                            $AddMemberScoutnetAddress = $True
                        }
                    }
                }

                if ($AddMemberScoutnetAddress)
                {
                    $mailaddresses | ForEach-Object {
                        if (![string]::IsNullOrWhiteSpace($_))
                        {
                            Add-MailContactToList -Epost $_ -DisplayName $displayName -DistGroupName $distGroupName
                            $mailListsToProcessMembers.Remove($_)
                        }
                    }
                }
            }

            $ledare_synk_option = $MailListSettings[$distGroupName].ledare_synk_option
            $ledare_synk_option = $ledare_synk_option.ToLower()

            Write-SNSLog "Ledare synk option is '$ledare_synk_option' for distribution group $distGroupName"
            # Get the settings for ledare in this list.
            $AddLedareOffice365Address = $True
            $AddLedareScoutnetAddress = $False
            $AddLedareOffice365AddressTryFirst = $False # First try to add office 365 address. If that fails add the scoutnet version.

            if ($ledare_synk_option -like "-")
            {
                $AddLedareScoutnetAddress = $True
                $AddLedareOffice365Address = $False
            }
            elseif ($ledare_synk_option -like "&")
            {
                $AddLedareScoutnetAddress = $True
            }

            if ($ledare_synk_option -like "t")
            {
                $AddLedareScoutnetAddress = $False
                $AddLedareOffice365AddressTryFirst = $True
                $AddLedareOffice365Address = $True
            }

            $listData = $CustomLists[$distGroupId].ledare
            foreach ($member in $listData)
            {
                if ($null -eq $member)
                {
                    Write-SNSLog -Level "Warn" "Maillist $distGroupName contained a null value."
                    continue
                }
                if (-not $AllMailAddresses.ContainsKey($member))
                {
                    Write-SNSLog -Level "Warn" "Id '$member' from $distGroupName not found in list of all members."
                    continue
                }

                $MemberData = $AllMailAddresses[$member]
                $displayName = "$($MemberData.first_name) $($MemberData.last_name)"

                if ($MailListSettings[$distGroupName].ignore_user)
                {
                    if ($MailListSettings[$distGroupName].ignore_user.Contains($member))
                    {
                        Write-SNSLog "Ignoring '$($displayName)' with Id '$member' for $distGroupName."
                        continue
                    }
                }

                $DoAddLedareScoutnetAddress = $AddLedareScoutnetAddress
                if ($AddLedareOffice365Address)
                {
                    $result = Add-Office365user -allOffice365Users $allOffice365Users -Member $member -MemberData $MemberData -distGroupName $distGroupName -doWarn
                    if (!$result)
                    {
                        if ($AddLedareOffice365AddressTryFirst)
                        {
                            $DoAddLedareScoutnetAddress = $True
                        }
                    }
                }

                if ($DoAddLedareScoutnetAddress)
                {
                    if ([string]::IsNullOrWhiteSpace($MemberData.primary_email))
                    {
                        Write-SNSLog -Level "Warn" "Primary mailaddres in scoutnet for '$displayName' is empty. Cannot add member to list."
                    }
                    else
                    {
                        Add-MailContactToList -Epost $MemberData.primary_email -DisplayName $displayName -DistGroupName $distGroupName
                        $mailListsToProcessMembers.Remove($MemberData.primary_email)
                    }
                }
            }

            # Add all mailaddresses listed in email_addresses for the maillist.
            foreach ($email in $MailListSettings[$distGroupName].email_addresses)
            {
                if ([string]::IsNullOrWhiteSpace($email))
                {
                    continue
                }

                Add-MailContactToList -Epost $email -DisplayName $email -DistGroupName $distGroupName
                $mailListsToProcessMembers.Remove($email)
            }
        }

        Write-SNSLog " "
        Write-SNSLog "Removing old contacts"
        Write-SNSLog "Number of contacts to check for removal $($mailListsToProcessMembers.count)"

        # Delete all contacts that still is in the list of contacts to process.
        $mailListsToProcessMembers.values | ForEach-Object {
            $medlem = $_
            if ($medlem.RecipientType -eq "MailContact")
            {
                # Check this contact is member in any other maillist.
                $IsInmaillistMembers = $otherMailListsMembers[$_.Identity]
                if ($null -eq $IsInmaillistMembers)
                {
                    # Not used in any maillists. Remove the contact.
                    Write-SNSLog "Removing MailContact $($medlem.Identity)"
                    Remove-MailContact $medlem.Identity -Confirm:$Y
                }
            }
        }
        
        Remove-PSSession $ExchangeSession
        Write-SNSLog " "
        Write-SNSLog "Update done new hash value is $NewValidationHash"
    }
    return $NewValidationHash
}
