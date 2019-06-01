function Get-SNSApiGroupCustomlist
{
    <#
    .SYNOPSIS
        Fetches the maillist info from scoutnet using the provided credentials.

    .DESCRIPTION
        Scoutnet API api/group/customlists returns a JSON with maillist info.
        If the maillist ID is provided all members in the list is returned.
        If the maillist ID is not provided information of the lists that can be fetched is provided.
        The credential for the API is configured in scoutnet.

    .INPUTS
        None. You cannot pipe objects to Get-SNSApiGroupCustomlist.

    .OUTPUTS
        The data from ConvertFrom-Json.

    .PARAMETER ApiUrl
        Url to the scoutnet API for fetching the group memberlist.

    .PARAMETER Credential
        Credentials for api/group/customlists

    .PARAMETER listid
        Id for the list to fetch. If it is empty the list info is fetched.

    .PARAMETER ruleid
        RUle id is used to only fetch rule data for a list. listid must be present.

    .LINK
        https://www.scoutnet.se

    .EXAMPLE
        Get-SNSApiGroupCustomlist -Credential $CredentialCustomLists -Uri "https://www.scoutnet.se/api/group/customlists"
    #>

    [OutputType([System.Collections.Hashtable])]
    param (
        [Parameter(Mandatory=$True, HelpMessage="Url to Scoutnet API for fetching the group customlists")]
        [ValidateNotNull()]
        [string]$Uri,

        [Parameter(Mandatory=$True, HelpMessage="Credentials for api/group/customlists")]
        [ValidateNotNull()]
        [pscredential]$Credential,

        [Parameter(Mandatory=$False, HelpMessage="Id for the list to fetch. If it is empty then list info is fetched.")]
        [string]$listid = "",

        [Parameter(Mandatory=$False, HelpMessage="Rule id. Only usable in combination with listid")]
        [string]$ruleid
        )

    $ApiUri = $Uri + "?list_id=" + $listid
    if (![string]::IsNullOrWhiteSpace($listid) -And ![string]::IsNullOrWhiteSpace($ruleid))
    {
        $ApiUri += "&rule_id=" + $ruleid
    }

    $JsonHashTable = Receive-SNSApiJson -Uri $ApiUri -Credential $Credential
    return $JsonHashTable
}
