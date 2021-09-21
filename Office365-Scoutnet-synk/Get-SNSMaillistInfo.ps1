function Get-SNSMaillistInfo
{
    <#
    .SYNOPSIS
        Fetches the maillist info from scoutnet and creates a hashtable with all lists and all users.

    .DESCRIPTION
        Scoutnet API api/group/customlists returns a JSON with maillist info.
        This function fetches all customlists and creates a hashtable with all lists and all users.
        This hashtable can be used to generate maillists.

        The function supports that one rule is named 'ledare'. This is to be able to handle them separate.

    .INPUTS
        None. You cannot pipe objects to Get-SNSMaillistInfo.

    .OUTPUTS
        Two parts is returned. The first part is the hastable. The second part is a hashcode that can be used to check if scoutnet is updated.

    .PARAMETER CredentialCustomlists
        Credentials for api/group/customlists

    .PARAMETER UriApiCustomList
        Url for the API. Defaults to https://www.scoutnet.se/api/group/customlists

    .PARAMETER Maillists
        Maillists to process. A hashtable there the key kan be used to find the list in office365, and the value is the scoutnet list Id.

    .LINK
        https://www.scoutnet.se

    .EXAMPLE
        Get-SNSMaillistInfo -CredentialCustomlists $CredentialCustomLists -Uri "https://www.scoutnet.se/api/group/customlists"
    #>

    [OutputType([System.Collections.Hashtable], [string])]
    param (
        [Parameter(Mandatory=$True, HelpMessage="Credentials for api/group/customlists")]
        [ValidateNotNull()]
        [Alias("Credential")]
        [pscredential]$CredentialCustomlists,

        [Parameter(Mandatory=$False, HelpMessage="Url for api/group/customlists.")]
        [ValidateNotNull()]
        [Alias("Uri")]
        [string]$UriApiCustomList = "https://www.scoutnet.se/api/group/customlists",

        [Parameter(Mandatory=$True, HelpMessage="Maillists to process. A hash-table there the key kan be used to find the list in office365, and the value is the Scoutnet list Id.")]
        [ValidateNotNull()]
        $MaillistsIds
        )

    # Fetch exixting custom lists.
    $CustomLists = @{}

    Write-SNSLog "Fetch the current maillists from Scoutnet"
    $CustomListInfo = Get-SNSApiGroupCustomlist -Credential $CredentialCustomlists -Uri $UriApiCustomList

    # For each list fetch the user id of the members.
    foreach ($key in $CustomListInfo.Keys)
    {
        if ($MaillistsIds.ContainsValue($key) | Sort-Object -Unique)
        {
            $ledare = @()
            $scouter = @()

            Write-SNSLog ("Fetching data for maillist {0}" -f ($CustomListInfo[$key].title))
            $enum = $MaillistsIds.GetEnumerator().Where({$_.ContainsValue($key)})

            if ($enum[0].statisk_lista)
            {
                # Static list. Fetch users from the list not the rules.
                $ledareData = Get-SNSApiGroupCustomlist -Credential $CredentialCustomlists -Uri $UriApiCustomList -listid $key
                $ledare += $ledareData.data.keys
            }
            else
            {
                # Fetch rule information.
                foreach ($rule in $CustomListInfo[$key].rules.Keys)
                {
                    if ($CustomListInfo[$key].rules[$rule].title -like "*ledare*")
                    {
                        # Fetch the list marked "ledare"
                        $ledareData = Get-SNSApiGroupCustomlist -Credential $CredentialCustomlists -Uri $UriApiCustomList -listid $key -ruleid $CustomListInfo[$key].rules[$rule].id
                        $ledare += $ledareData.data.keys
                    }
                    else
                    {
                        # Fetch the other lists
                        $scouterData = Get-SNSApiGroupCustomlist -Credential $CredentialCustomlists -Uri $UriApiCustomList -listid $key -ruleid $CustomListInfo[$key].rules[$rule].id
                        $scouter += $scouterData.data.keys
                    }
                }
            }
            # Create hashtable with all lists and all users.
            $CustomListData = @{}
            $CustomListData.Add("title",   $CustomListInfo[$key].title)
            $CustomListData.Add("ledare",  $ledare)
            $CustomListData.Add("scouter", $scouter)
            $CustomLists.Add($key, $CustomListData)
        }
    }
    Write-SNSLog "Done"

    $str = $CustomLists | ConvertTo-Json
    return $CustomLists, ("{0:X8}" -f ($str.GetHashCode()))
}