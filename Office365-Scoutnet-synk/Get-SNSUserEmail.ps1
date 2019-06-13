function Get-SNSUserEmail
{
    <#
    .SYNOPSIS
        Fetches all e-mail addresses for all users from from scoutnet based on the credential.

    .DESCRIPTION
        Scoutnet API api/group/memberlist returns a JSON with info about all members.
        This function fetches this json list and creates a hashtable with the members mailaddresses.
        This hashtable can be used to generate maillists.

    .INPUTS
        None. You cannot pipe objects to Get-SNSUserEmail.

    .OUTPUTS
        Two parts is returned. The first part is the hastable. The second part is a hashcode that can be used to check if scoutnet is updated.

    .PARAMETER CredentialMemberlist
        Credentials for api/group/memberlist

    .PARAMETER UriApiMemberList
        Url for the API. Defaults to https://www.scoutnet.se/api/group/memberlist

    .LINK
        https://www.scoutnet.se

    .EXAMPLE
        Get-SNSUserEmail -CredentialMemberlist $CredentialMemberlist -UriApiMemberList "https://www.scoutnet.se/api/group/memberlist"
    #>

    [OutputType([System.Collections.Hashtable], [string])]
    param (
        [Parameter(Mandatory=$True, HelpMessage="Credentials for api/group/memberlist")]
        [ValidateNotNull()]
        [Alias("Credential")]
        [pscredential]$CredentialMemberlist,

        [Parameter(Mandatory=$False, HelpMessage="Url for api/group/memberlist.")]
        [ValidateNotNull()]
        [Alias("Uri")]
        [string]$UriApiMemberList = "https://www.scoutnet.se/api/group/memberlist"
        )

    # Fetch all mail addresses from Scoutnet.
    $allUsers = Get-SNSApiGroupMemberlist -Uri $UriApiMemberList -Credential $CredentialMemberlist

    $allMailAddresses = @{}
    foreach ($member in $allUsers.data.values)
    {
        $memberData = @{}
        $memberData.Add("first_name", $member.first_name.value)
        $memberData.Add("last_name", $member.last_name.value)
        $mailaddresses = @()
        $contacts_addresses = @() # Parents...
        if (![string]::IsNullOrWhiteSpace($member.email.value))
        {
            $memberData.Add("primary_email", $member.email.value)
            $mailaddresses += $member.email.value
        }
        if (![string]::IsNullOrWhiteSpace($member.contact_email_mum.value))
        {
            $mailaddresses += $member.contact_email_mum.value
            $contacts_addresses += $member.contact_email_mum.value
        }
        if (![string]::IsNullOrWhiteSpace($member.contact_email_dad.value))
        {
            $mailaddresses += $member.contact_email_dad.value
            $contacts_addresses += $member.contact_email_dad.value
        }
        if (![string]::IsNullOrWhiteSpace($member.contact_alt_email.value))
        {
            $mailaddresses += $member.contact_alt_email.value
            $memberData.Add("alt_email", $member.contact_alt_email.value)
        }
        $mailaddresses = $mailaddresses | Sort-Object -Unique
        $memberData.Add("mailaddresses", $mailaddresses)
        $contacts_addresses = $contacts_addresses | Sort-Object -Unique
        $memberData.Add("contacts_addresses", $contacts_addresses)
        $allMailAddresses.Add($member.member_no.value, $memberData)
    }

    $str = $allMailAddresses | ConvertTo-Json
    return $allMailAddresses, ("{0:X8}" -f ($str.GetHashCode()))
}