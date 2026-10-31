function Get-SNSExchangeMailListMember
{
    <#
    .SYNOPSIS
        Fetches members of exchange distribution groups.

    .DESCRIPTION
        Fetches the distribution groups members and returns them in a ArrayList and a hashtable.
        The ArrayList is "other" distribution groups members that can be checked if a mailaddress can be reomoved or not.
        The hashtable is "other" distribution groups members
        Only contacts is returned. Users with mailboxes is not returned.
        The distribution group list is fetched once; configured lists are matched by SMTP address, alias, name or identity.

    .INPUTS
        None. You cannot pipe objects to Get-SNSExchangeMailListMember.

    .OUTPUTS
        Two parts is returned.
        The first part is the otherMailListsMembers ArrayList.
        The second part is the mailListsToProcessMembers ArrayList.

    .PARAMETER Maillists
        Distribution groups that will be part of mailListsToProcessMembers.
    #>

    [OutputType([System.Collections.ArrayList], [System.Collections.ArrayList])]
    param (
        [Parameter(Mandatory=$True, HelpMessage="Distribution groups that will be part of mailListsToProcessMembers.")]
        [string[]]$Maillists
        )

    $otherMailListsMembers = @{}
    $mailListsToProcessMembers = @{}

    $mailListGroups = @{}
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
                    break
                }
            }

            Write-SNSLog "Get distribution list $($group.DisplayName)"
            $data = Get-DistributionGroupMember -Verbose:$false -Identity "$($group.ExchangeObjectId)" -ErrorAction Stop
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
    Write-SNSLog "Done"

    return $otherMailListsMembers, $mailListsToProcessMembers
}