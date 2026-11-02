function Get-SNSExchangeMailListMember
{
    <#
    .SYNOPSIS
        Fetches members of exchange distribution groups.

    .DESCRIPTION
        Fetches the distribution groups members and returns them in two hashtables, plus a third hashtable with the current member addresses.
        The first hashtable is "other" distribution groups members that can be checked if a mailaddress can be reomoved or not.
        The second hashtable is the members of the configured distribution groups.
        Only contacts is returned in these two. Users with mailboxes is not returned.
        The distribution group list is fetched once; configured lists are matched by SMTP address, alias, name or identity.
        A third value is returned with the addresses of all current members (any recipient type) of each configured list,
        keyed by the configured name. It is used to skip updates of lists that already have the wanted members.

    .INPUTS
        None. You cannot pipe objects to Get-SNSExchangeMailListMember.

    .OUTPUTS
        Three parts is returned.
        The first part is the otherMailListsMembers hashtable.
        The second part is the mailListsToProcessMembers hashtable.
        The third part is a hashtable with the current member addresses per configured list.

    .PARAMETER Maillists
        Distribution groups that will be part of mailListsToProcessMembers.
    #>

    [OutputType([System.Collections.Hashtable], [System.Collections.Hashtable], [System.Collections.Hashtable])]
    param (
        [Parameter(Mandatory=$True, HelpMessage="Distribution groups that will be part of mailListsToProcessMembers.")]
        [string[]]$Maillists
        )

    $otherMailListsMembers = @{}
    $mailListsToProcessMembers = @{}

    $mailListGroups = @{}
    $groupMemberKeys = @{}
    $matchedConfiguredNames = @{}
    $configuredGroupNames = @{}
    foreach ($mailList in $Maillists)
    {
        $configuredGroupNames[[string]$mailList] = $true
    }
    try
    {
        $groups = Get-DistributionGroup -Verbose:$false -ErrorAction Stop
    }
    catch
    {
        Write-SNSLog -Level "Error" "Could not fetch data for all distribution groups. Error $_"
        throw
    }

    try
    {
        foreach($group in $groups)
        {
            $matchedGroupName = $null
            $groupNames = @(
                [string]$group.PrimarySmtpAddress
                [string]$group.Alias
                [string]$group.Name
                [string]$group.Identity
            )
            foreach ($groupName in $groupNames)
            {
                if (-not [string]::IsNullOrWhiteSpace($groupName) -and $configuredGroupNames.ContainsKey($groupName))
                {
                    if (-not [string]::IsNullOrWhiteSpace([string]$group.ExchangeObjectId))
                    {
                        $mailListGroups[[string]$group.ExchangeObjectId] = $true
                    }
                    $matchedGroupName = $groupName
                    $matchedConfiguredNames[$groupName] = $true
                    break
                }
            }

            Write-SNSLog "Get distribution list $($group.DisplayName)"
            $data = Get-DistributionGroupMember -Verbose:$false -Identity "$($group.ExchangeObjectId)" -ErrorAction Stop
            if ($null -ne $matchedGroupName)
            {
                # One address set per member (primary, external and proxy addresses), used to skip unchanged groups.
                $memberKeySets = [System.Collections.Generic.List[object]]::new()
                foreach ($member in @($data))
                {
                    $memberKeySets.Add(@(Get-SNSRecipientAddressKey -Recipient $member))
                }
                $groupMemberKeys[$matchedGroupName] = $memberKeySets
            }
            if ($mailListGroups.ContainsKey([string]$group.ExchangeObjectId))
            {
                $data | ForEach-Object {
                    if ($_.RecipientType -eq "MailContact")
                    {
                        $mailListsToProcessMembers[$_.Identity] = $_
                    }
                }
            }
            else
            {
                $data | ForEach-Object {
                    if ($_.RecipientType -eq "MailContact")
                    {
                        $otherMailListsMembers[$_.Identity] = $_
                    }
                }
            }
        }
    }
    catch
    {
        Write-SNSLog -Level "Error" "Fetch of distribution group members failed. Error $_"
        throw
    }
    foreach ($configuredName in $configuredGroupNames.Keys)
    {
        if (-not $matchedConfiguredNames.ContainsKey($configuredName))
        {
            Write-SNSLog -Level "Warn" "Distribution group '$configuredName' was not found by SMTP address, alias, name or identity. Its current members are unknown, so its members are always updated. Use the full e-mail address of the group as key in MailListSettings."
        }
    }
    Write-SNSLog "Done"

    return $otherMailListsMembers, $mailListsToProcessMembers, $groupMemberKeys
}