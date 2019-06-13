# Cache
$Script:SNSApiGroupMemberlistCache=$null
Export-ModuleMember -Variable SNSApiGroupMemberlistCache

function Get-SNSApiGroupMemberlist
{
    <#
    .SYNOPSIS
        Fetches all users from from scoutnet based on the credential.

    .DESCRIPTION
        Scoutnet API api/group/memberlist returns a JSON with info about all members.
        This function fetches this json list and returns the data.
        The credential for the API is configured in scoutnet.

    .INPUTS
        None. You cannot pipe objects to Get-SNSUserEmail.

    .OUTPUTS
        The data from ConvertFrom-Json.

    .PARAMETER Credential
        Credentials for api/group/memberlist

    .PARAMETER Uri
        Url for the API. Defaults to https://www.scoutnet.se/api/group/memberlist

    .LINK
        https://www.scoutnet.se

    .EXAMPLE
        Get-SNSUserEmail -CredentialMemberlist $CredentialMemberlist
    #>

    [OutputType([System.Collections.Hashtable])]
    param (
        [Parameter(Mandatory=$False, HelpMessage="Credentials for api/group/memberlist")]
        [ValidateNotNull()]
        [pscredential]$Credential,

        [Parameter(Mandatory=$False, HelpMessage="Url for api/group/memberlist.")]
        [ValidateNotNull()]
        [string]$Uri = "https://www.scoutnet.se/api/group/memberlist"
        )

    if ($Script:SNSApiGroupMemberlistCache)
    {
        return $Script:SNSApiGroupMemberlistCache
    }
    else
    {
        $Script:SNSApiGroupMemberlistCache = Receive-SNSApiJson -Uri $Uri -Credential $Credential
        return $Script:SNSApiGroupMemberlistCache
    }
}