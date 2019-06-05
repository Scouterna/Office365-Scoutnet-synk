function Get-SNSExchangeMailListMember
{
    <#
    .SYNOPSIS
        Fetches members of exchange distribution groups.

    .DESCRIPTION
        Fetches the distribution groups members and returns them in a ArrayList and a hashtable.
        The ArrayList is "other" distribution groups members that can be checked if a mailaddress can be reomoved or not.
        The hashtable is "other" distribution groups members

    .INPUTS
        None. You cannot pipe objects to Get-SNSExchangeMailListMember.

    .OUTPUTS
        Two parts is returned.
        The first part is the otherMailListsMembers ArrayList.
        The second part is the mailListsToProcessMembers ArrayList.

    .PARAMETER Credential365
        Credentials for office365 that can execute Get-DistributionGroupMember and Get-DistributionGroup for the selected DistributionGroup.

    .PARAMETER ExchangeSession
        Exchange session to use for Import-PSSession.

    .PARAMETER Maillists
        Distribution groups that will be part of mailListsToProcessMembers.
    #>

    [OutputType([System.Collections.ArrayList], [System.Collections.ArrayList])]
    param (
        [Parameter(Mandatory=$True, HelpMessage="Credentials for office365")]
        [ValidateNotNull()]
        [Alias("Credential")]
        [pscredential]$Credential365,

        [Parameter(Mandatory=$True, HelpMessage="Exchange session to use.")]
        $ExchangeSession,

        [Parameter(Mandatory=$True, HelpMessage="Distribution groups that will be part of mailListsToProcessMembers.")]
        [string[]]$Maillists
        )

    Import-PSSession $ExchangeSession -AllowClobber -CommandName Get-DistributionGroupMember,Get-DistributionGroup > $null

    $otherMailListsMembers = [System.Collections.ArrayList]::new()
    $mailListsToProcessMembers = [System.Collections.ArrayList]::new()

    $mailListGroups = @()
    foreach($mailList in $Maillists)
    {
        $mailListGroups += (Get-DistributionGroup $mailList).Identity
    }

    $groups = Get-DistributionGroup
    foreach($group in $groups)
    {
        Write-SNSLog "Get distribution list $($group.DisplayName)"
        $data = Get-DistributionGroupMember -Identity $group.Identity
        if ($mailListGroups.Contains($group.Identity))
        {
            $data | ForEach-Object {[void]$mailListsToProcessMembers.Add($_)}
        }
        else
        {
            $data | ForEach-Object {[void]$otherMailListsMembers.Add($_)}
        }
    }
    Write-SNSLog "Done"

    $otherMailListsMembers = $otherMailListsMembers | Sort-Object -Unique -Property ExchangeObjectId
    $mailListsToProcessMembers = $mailListsToProcessMembers | Sort-Object -Unique -Property ExchangeObjectId
    return $otherMailListsMembers, $mailListsToProcessMembers
}