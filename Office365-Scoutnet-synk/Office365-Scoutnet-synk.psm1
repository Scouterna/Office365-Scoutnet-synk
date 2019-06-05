﻿#Requires -Version 5.1

[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"

#region import everything we need

. $PSScriptRoot\Write-SNSLog.ps1
. $PSScriptRoot\ConvertTo-SNSJSONHash.ps1
. $PSScriptRoot\Receive-SNSApiJson.ps1
. $PSScriptRoot\Get-SNSApiGroupCustomlist.ps1
. $PSScriptRoot\Get-SNSMaillistInfo.ps1
. $PSScriptRoot\Get-SNSUserEmail.ps1
. $PSScriptRoot\Get-SNSExchangeMailListMember.ps1

#endregion

function AddOffice365User
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
            Write-SNSLog -Level "Warn" "Could not add contact $($recipient.DisplayName) to $distGroupName. Error $_"
        }
    }
    elseif ($doWarn)
    {
        Write-SNSLog -Level "Warn" "Member $($MemberData.first_name) $($MemberData.last_name) not found in office 365. Please make sure that CustomAttribute1 contains Scoutnet Id for the user."
    }
    return $userAdded
}

function SNSUpdateExchangeDistributionGroups
{
    [CmdletBinding(HelpURI = 'https://github.com/scouternasetjanster/Office365-Scoutnet-synk')]
    <#
    .SYNOPSIS
        Updates exchange distribution groups based on scoutnet maillists.

    .DESCRIPTION
        Fetches the distribution groups members and updates corressponding distribution groups based on the info from scoutnet.
        As all members of an distribution group must be present in exchange as user or contact, contacts will be created for external addresses.

        Function behavior.
        1) Validate that Scoutnet is updated. If Scoutnet is not updated abort.
        2) Remove all members from the specifed distribution groups
        3) Remove corresponding contacts if they are not part of any distribution groups that is not connected to Scoutnet.
        4) For each distribution group add the contacts based on the maillist info from Scoutnet. Contacts is created if needed.
        5) For leaders add their exchange user.
        6) Add admin mail addresses to all distribution groups if specified.
        7) Return the hash values from Scoutnet to be stored an used as reference.

    .INPUTS
        None. You cannot pipe objects to SNSUpdateExchangeDistributionGroups.

    .OUTPUTS
        None.

    .PARAMETER CredentialCustomlists
        Credentials for api/group/customlists

    .PARAMETER CredentialMemberlist
        Credentials for api/group/memberlist

    .PARAMETER Credential365
        Credentials for office365 that can execute needed servlets.

    .PARAMETER MailListSettings
        Maillists to process. A hashtable with maillist info.

    .PARAMETER ValidationHash
        Hash value used to validate if scoutnet is updated.

    .PARAMETER DomainName
        Domain name for office365 mail addresses.
    #>

    [OutputType([string])]
    param (
        [Parameter(Mandatory=$True, HelpMessage="Credentials for api/group/customlists.")]
        [ValidateNotNull()]
        [pscredential]$CredentialCustomlists,

        [Parameter(Mandatory=$True, HelpMessage="Credentials for api/group/memberlist.")]
        [ValidateNotNull()]
        [pscredential]$CredentialMemberlist,

        [Parameter(Mandatory=$True, HelpMessage="Credentials for office365.")]
        [ValidateNotNull()]
        [pscredential]$Credential365,

        [Parameter(Mandatory=$True, HelpMessage="Maillists to process.")]
        [ValidateNotNull()]
        $MailListSettings,

        [Parameter(Mandatory=$True, HelpMessage="Hash value used to validate if Scoutnet is updated.")]
        [ValidateNotNull()]
        [string]$ValidationHash,

        [Parameter(Mandatory=$True, HelpMessage="Domain name for office365 mail addresses.")]
        [ValidateNotNull()]
        [string]$DomainName
        )

    # Fetch maillist info from scoutnet.
    $CustomLists, $CustomListsHash = Get-SNSMaillistInfo -CredentialCustomlists $CredentialCustomlists -MaillistsIds $MailListSettings.values

    # Fetch all members and their mailaddresses.
    $allMailAddresses, $allMailAddressesHash = Get-SNSUserEmail -CredentialMemberlist $CredentialMemberlist

    $NewValidationHash = "0x{0}{1}" -f ($CustomListsHash, $allMailAddressesHash)

    Write-SNSLog "Saved validation hash: $ValidationHash new value $NewValidationHash"

    if ($ValidationHash -eq $NewValidationHash)
    {
        Write-SNSLog "Scoutnet is not updated. No update will be done."
    }
    else
    {
        Write-SNSLog "Scoutnet is updated. Starting to update the distribution groups."

        $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential365 -Authentication Basic -AllowRedirection
        Import-PSSession $ExchangeSession -AllowClobber -CommandName Set-MailContact,Set-Contact,New-MailContact,Remove-MailContact,Remove-DistributionGroupMember,Get-Recipient,Add-DistributionGroupMember,Get-Mailbox > $null

        $otherMailListsMembers, $mailListsToProcessMembers = Get-SNSExchangeMailListMember -Credential365 $Credential365 -ExchangeSession $ExchangeSession -Maillists $MailListSettings.Keys

        # Clean the distribution groups first, so any removed addresses in Scoutnet is removed in o365.
        Write-SNSLog " "
        Write-SNSLog "Removing contacts in distribution groups"

        # First delete the mail contact, as that also deletes the contact from the distribution list.
        $mailListsToProcessMembers | ForEach-Object {
            $medlem = $_
            if ($medlem.RecipientType -eq "MailContact")
            {
                # Check this contact is member in any other maillist.
                $IsInmaillistMembers = $otherMailListsMembers |Where-Object {$_.Identity -eq $medlem.Identity}
                if ($null -eq $IsInmaillistMembers)
                {
                    Write-SNSLog "Removing MailContact $($medlem.Identity)"
                    Remove-MailContact $medlem.Identity -Confirm:$Y
                }
            }
        }

        foreach ($distGroupName in $MailListSettings.keys)
        {
            Write-SNSLog "Remove contacts in distribution group $($distGroupName)"
            $distGroupMembers= Get-DistributionGroupMember -Identity $distGroupName
            foreach ($medlem in $distGroupMembers)
            {
                if ($medlem.RecipientType -eq "MailContact")
                {
                    Write-SNSLog "Removing mail contact $($medlem.Name) for $($medlem.Company) from distribution group $($distGroupName)"
                }
                else
                {
                    Write-SNSLog "Removing user $($medlem.Name) from distribution group $($distGroupName)"
                }
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
                }

                if ($scouter_synk_option.contains("t"))
                {
                    $AddMemberScoutnetAddress = $False
                    $AddMemberOffice365AddressTryFirst = $True
                    $AddMemberOffice365Address = $True
                }

                if ($mailaddresses.Length -eq 0)
                {
                    Write-SNSLog -Level "Warn" "No email addresses found for $($MemberData.first_name) $($MemberData.last_name)."
                    continue
                }

                $AddMemberScoutnetAddress = $AddMemberScoutnetAddress
                if ($AddMemberOffice365Address)
                {
                    $result = AddOffice365user -allOffice365Users $allOffice365Users -Member $member -MemberData $AllMailAddresses[$member] -distGroupName $distGroupName
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
                        $displayName = "$($MemberData.first_name) $($MemberData.last_name)"

                        $epost = $_
                        $ExistingMailContact = get-recipient $epost -ErrorAction "SilentlyContinue"
                        if ($null -eq $ExistingMailContact)
                        {
                            Write-SNSLog "Creating Contact $epost for $displayName"
                            try
                            {
                                New-MailContact -Name $epost -ExternalEmailAddress $epost -ErrorAction "stop" > $null

                                # Set the name of the member in the company field. This is visibel in Office 365 admin console.
                                Set-Contact -Identity $epost -Company "$displayName"
                                Set-MailContact -Identity $epost -HiddenFromAddressListsEnabled $true
                            }
                            Catch
                            {
                                Write-SNSLog -Level "Warn" "Could not create mail contact with address $epost. Error $_"
                            }
                        }
                        Write-SNSLog "Adding contact $epost for $displayName to distribution group $distGroupName"
                        try
                        {
                            Add-DistributionGroupMember -Identity $distGroupName -Member $epost -ErrorAction "stop"
                        }
                        Catch
                        {
                            Write-SNSLog -Level "Warn" "Could not add contact $epost to $distGroupName. Error $_"
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

                $DoAddLedareScoutnetAddress = $AddLedareScoutnetAddress
                if ($AddLedareOffice365Address)
                {
                    $result = AddOffice365user -allOffice365Users $allOffice365Users -Member $member -MemberData $AllMailAddresses[$member] -distGroupName $distGroupName -doWarn
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
                    $MemberData = $AllMailAddresses[$member]

                    $ExistingMailContact = get-recipient $MemberData.primary_email -ErrorAction "SilentlyContinue"
                    $displayName = "$($MemberData.first_name) $($MemberData.last_name)"
                    if ($null -eq $ExistingMailContact)
                    {
                        Write-SNSLog "Creating Contact $($MemberData.primary_email) for $displayName"
                        try
                        {
                            New-MailContact -Name $MemberData.primary_email -ExternalEmailAddress $MemberData.primary_email -ErrorAction "stop" > $null
                            Set-Contact -Identity $MemberData.primary_email -Company "$displayName"
                            Set-MailContact -Identity $MemberData.primary_email -HiddenFromAddressListsEnabled $true
                        }
                        Catch
                        {
                            Write-SNSLog -Level "Warn" "Could not create mail contact with address $($MemberData.primary_email). Error $_"
                        }
                    }

                    Write-SNSLog "Adding contact $($MemberData.primary_email) for $displayName to distribution group $distGroupName"
                    try
                    {
                        Add-DistributionGroupMember -Identity $distGroupName -Member $MemberData.primary_email -ErrorAction "stop"
                    }
                    Catch
                    {
                        Write-SNSLog -Level "Warn" "Could not add contact $($MemberData.primary_email) to $distGroupName. Error $_"
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

                $ExistingMailContact = get-recipient $email -ErrorAction "SilentlyContinue"
                if ($null -eq $ExistingMailContact)
                {
                    try
                    {
                        New-MailContact -Name $email -ExternalEmailAddress $email -ErrorAction "stop" > $null
                        Set-MailContact -Identity $email -HiddenFromAddressListsEnabled $true
                        Write-SNSLog "Creating Contact $($email)"
                    }
                    Catch
                    {
                        Write-SNSLog -Level "Warn" "Could not create mail contact with address $email. Error $_"
                    }
                }

                try
                {
                    Add-DistributionGroupMember -Identity $distGroupName -Member $email  -ErrorAction "stop"
                    Write-SNSLog "Adding contact $email to distribution group $distGroupName"
                }
                catch
                {
                    Write-SNSLog -Level "Warn" "Could not add contact $email to $distGroupName. Error $_"
                }
            }
        }
        Remove-PSSession $ExchangeSession
    }
    return $NewValidationHash
}
