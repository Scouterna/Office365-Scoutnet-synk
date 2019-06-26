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
        
    try
    {
        Import-PSSession $ExchangeSession -AllowClobber -CommandName Get-DistributionGroupMember,Get-DistributionGroup -ErrorAction Stop > $null
    }
    catch
    {
        Write-SNSLog -Level "Error" "Could not import needed functions. Error $_"
        throw
    }

    $otherMailListsMembers = @{}
    $mailListsToProcessMembers = @{}

    $mailListGroups = @()
    foreach($mailList in $Maillists)
    {
        try
        {
            $mailListGroups += (Get-DistributionGroup $mailList -ErrorAction Stop).ExchangeObjectId
        }
        catch
        {
            Write-SNSLog -Level "Error" "Could not fetch data for distribution group '$mailList'. Error $_"
            throw
        }
    }

    try
    {
        $groups = Get-DistributionGroup -ErrorAction Stop
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
            Write-SNSLog "Get distribution list $($group.DisplayName)"
            $data = Get-DistributionGroupMember -Identity "$($group.ExchangeObjectId)" -ErrorAction Stop
            if ($mailListGroups.Contains($group.ExchangeObjectId))
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