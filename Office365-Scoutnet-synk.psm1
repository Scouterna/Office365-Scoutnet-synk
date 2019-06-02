#Requires -Version 5.1

[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"

function Write-SNSLog
{
<#
.Synopsis
   Write-SNSLog writes a message to a specified log file with the current time stamp.

.DESCRIPTION
   The Write-SNSLog function is designed to add logging capability to other scripts.
   In addition to writing output and/or verbose you can write to a log file for
   later debugging.
   The global variable $Global:LogPath can be used to oweride the default logfile scoutnetSync.log.

.PARAMETER Message
   Message is the content that you wish to add to the log file.

.PARAMETER Level
   Specify the criticality of the log information being written to the log (i.e. Error, Warning, Informational)
#>
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("LogContent")]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [ValidateSet("Error","Warn","Info")]
        [string]$Level="Info"
    )

    Process
    {
        $Path='scoutnetSync.log'
        if (-not [String]::IsNullOrWhiteSpace($Global:LogPath))
        {
            $Path=$Global:LogPath
        }

        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path.
        if (!(Test-Path $Path))
        {
            Write-Verbose "Creating $Path."
            New-Item $Path -Force -ItemType File > $null
        }

        # Format Date for our Log File
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        # Write message to error, warning, or verbose pipeline and specify $LevelText
        switch ($Level) {
            'Error' {
                Write-Error $Message
                $LevelText = 'ERROR:'
                }
            'Warn' {
                Write-Warning $Message
                $LevelText = 'WARNING:'
                }
            'Info' {
                Write-Verbose $Message
                $LevelText = 'INFO:'
                }
            }

        # Write log entry to $Path
        "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append -Encoding UTF8
    }
}

function ConvertTo-SNSJSONHash
{
    <#
    .SYNOPSIS
        Converts the PSCustomObject returned by ConvertFrom-Json to hashtable.

    .DESCRIPTION
        The function handles multi level JSON data and converts all levels to hashtables.
        HTML escape codes is unescaped before adding the data to the hashtable.
        The result is a nested hashtable.

    .INPUTS
        The data from ConvertFrom-Json.

    .OUTPUTS
        A hashtable with the JSON data.

    .PARAMETER root
        The data from ConvertFrom-Json.

    .EXAMPLE
        Invoke-WebRequest -Uri $ApiUrl | ConvertFrom-Json | ConvertTo-SNSJSONHash
    #>
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        $root
    )
    $hash = @{}

    $keys = $root | Get-Member -MemberType NoteProperty | Select-Object -exp Name

    $keys | ForEach-Object{
        $key=$_
        $obj=$root.$($_)
        if ($obj -match "@{")
        {
            $nesthash=ConvertTo-SNSJSONHash $obj
            $hash.add($key,$nesthash)
        }
        else
        {
            # Use unescape so the data is readable.
            $obj = [System.Text.RegularExpressions.Regex]::Unescape($obj)
            $hash.add($key,$obj)
        }
    }
    return $hash
}

function Receive-SNSApiJson
{
    <#
    .SYNOPSIS
        Fetches the API data from scoutnet using the provided credentials.

    .DESCRIPTION
        Scoutnet API returns a JSON with the resulting data.
        This functions uses Invoke-WebRequest to fetch the data,
        and parses the returned JSON and returns a multilevel hashtable with the data.
        The credential for the API is configured in scoutnet.

    .INPUTS
        None. You cannot pipe objects to Receive-SNSApiJson.

    .OUTPUTS
        The multi level hashtable.

    .PARAMETER Uri
        Url to the scoutnet API for fetching the group memberlist.

    .PARAMETER Credential
        Scoutnet Credentials for the selected API.

    .LINK
        https://www.scoutnet.se

    .EXAMPLE
        $PWordMembers = ConvertTo-SecureString -String "Your api key" -AsPlainText -Force
        $User ="0000" # Your API group ID.
        $CredentialMembers = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $PWordMembers
        Receive-SNSApiJson -Credential $CredentialMembers -Uri "https://www.scoutnet.se/api/group/memberlist"
    #>

    param (
        [Parameter(Mandatory=$True, HelpMessage="Url to Scoutnet API for fetching the group memberlist")]
        [ValidateNotNull()]
        [string]$Uri,

        [Parameter(Mandatory=$True, HelpMessage="Credentials for api/group/memberlist")]
        [ValidateNotNull()]
        [pscredential]$Credential
        )

    # Powershell version 5 do not support basic auth.
    # Workaround is to create the Authorization header with the data from the Credential parameter.
    $bytes = [System.Text.Encoding]::UTF8.GetBytes(
        ('{0}:{1}' -f $Credential.UserName, $Credential.GetNetworkCredential().Password)
    )
    $Authorization = 'Basic {0}' -f ([Convert]::ToBase64String($bytes))

    # Create the Authorization header.
    $Headers = @{ Authorization = $Authorization }

    # Recuest the data and convert the result to a hash table. Basic parsing is mandatory on Azure automation.
    $JsonHashTable = Invoke-WebRequest -Uri $Uri -Headers $Headers -UseBasicParsing -ErrorAction "Stop" | ConvertFrom-Json | ConvertTo-SNSJSONHash
    return $JsonHashTable
}

function Get-SNSApiGroupCustomlists
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
        None. You cannot pipe objects to Get-SNSApiGroupCustomlists.

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
        Get-SNSApiGroupCustomlists -Credential $CredentialCustomLists -Uri "https://www.scoutnet.se/api/group/customlists"
    #>

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
    $CustomListInfo = Get-SNSApiGroupCustomlists -Credential $CredentialCustomlists -Uri $UriApiCustomList

    # For each list fetch the user id of the members.
    foreach ($key in $CustomListInfo.Keys)
    {
        if ($MaillistsIds.ContainsValue($key) | Sort-Object -Unique)
        {
            $ledare = @()
            $scouter = @()

            Write-SNSLog ("Fetching data for maillist {0}" -f ($CustomListInfo[$key].title))
            # Fetch rule information.
            foreach ($rule in $CustomListInfo[$key].rules.Keys)
            {
                if ($CustomListInfo[$key].rules[$rule].title -like "*ledare*")
                {
                    # Fetch the list marked "ledare"
                    $ledareData = Get-SNSApiGroupCustomlists -Credential $CredentialCustomlists -Uri $UriApiCustomList -listid $key -ruleid $CustomListInfo[$key].rules[$rule].id
                    $ledare += $ledareData.data.keys
                }
                else
                {
                    # Fetch the other lists
                    $scouterData = Get-SNSApiGroupCustomlists -Credential $CredentialCustomlists -Uri $UriApiCustomList -listid $key -ruleid $CustomListInfo[$key].rules[$rule].id
                    $scouter += $scouterData.data.keys
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

function Get-SNSUserEmails
{
    <#
    .SYNOPSIS
        Fetches all e-mail addresses for all users from from scoutnet based on the credential.

    .DESCRIPTION
        Scoutnet API api/group/memberlist returns a JSON with info about all members.
        This function fetches this json list and creates a hashtable with the members mailaddresses.
        This hashtable can be used to generate maillists.

    .INPUTS
        None. You cannot pipe objects to Get-SNSUserEmails.

    .OUTPUTS
        Two parts is returned. The first part is the hastable. The second part is a hashcode that can be used to check if scoutnet is updated.

    .PARAMETER CredentialMemberlist
        Credentials for api/group/memberlist

    .PARAMETER UriApiMemberList
        Url for the API. Defaults to https://www.scoutnet.se/api/group/memberlist

    .LINK
        https://www.scoutnet.se

    .EXAMPLE
        Get-SNSUserEmails -CredentialMemberlist $CredentialMemberlist -UriApiMemberList "https://www.scoutnet.se/api/group/memberlist"
    #>
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
    $allUsers = Receive-SNSApiJson -Uri $UriApiMemberList -Credential $CredentialMemberlist

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

function Get-SNSExchangeMailListMembers
{
    <#
    .SYNOPSIS
        Fetches members of exchange distribution groups.

    .DESCRIPTION
        Fetches the distribution groups members and returns them in a ArrayList and a hashtable.
        The ArrayList is "other" distribution groups members that can be checked if a mailaddress can be reomoved or not.
        The hashtable is "other" distribution groups members

    .INPUTS
        None. You cannot pipe objects to Get-SNSUserEmails.

    .OUTPUTS
        Two parts is returned.
        The first part is the otherMailListsMembers ArrayList.
        The second part is the mailListsToProcessMembers.

    .PARAMETER Credential365
        Credentials for office365 that can execute Get-DistributionGroupMember and Get-DistributionGroup for the selected DistributionGroup.

    .PARAMETER ExchangeSession
        Exchange session to use for Import-PSSession.

    .PARAMETER Maillists
        Distribution groups that will be part of mailListsToProcessMembers.
    #>
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
    $mailListsToProcessMembers = @{}

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
            $mailListsMembers = [System.Collections.ArrayList]::new()
            $data | ForEach-Object {[void]$mailListsMembers.Add($_)}
            $mailListsToProcessMembers[$group.Identity] = $mailListsMembers
        }
        else
        {
            $data | ForEach-Object {[void]$otherMailListsMembers.Add($_)}
        }
    }
    Write-SNSLog "Done"

    return $otherMailListsMembers, $mailListsToProcessMembers
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
    $allMailAddresses, $allMailAddressesHash = Get-SNSUserEmails -CredentialMemberlist $CredentialMemberlist

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
        Import-PSSession $ExchangeSession -AllowClobber -CommandName Set-MailContact,Set-Contact,New-MailContact,Remove-MailContact,Remove-DistributionGroupMember,Get-Recipient,Add-DistributionGroupMember > $null

        $otherMailListsMembers, $mailListsToProcessMembers = Get-SNSExchangeMailListMembers -Credential365 $Credential365 -ExchangeSession $ExchangeSession -Maillists $MailListSettings.Keys

        # Clean the distribution groups first, so any removed addresses in Scoutnet is removed in o365.
        Write-SNSLog " "
        Write-SNSLog "Removing contacts in distribution groups"
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

                if ($medlem.RecipientType -eq "MailContact")
                {
                    # Check this contact is member in any other maillist.
                    $IsInmaillistMembers = $otherMailListsMembers |Where-Object {$_.Identity -eq $medlem.Identity}
                    if ($null -eq $IsInmaillistMembers)
                    {
                        Write-SNSLog "Remove-MailContact $($medlem.Identity)"
                        Remove-MailContact $medlem.Identity -Confirm:$Y
                    }
                }
            }
        }
        Write-SNSLog "Removed contacts in distribution groups"

        foreach ($distGroupName in $MailListSettings.keys)
        {
            $distGroupId = $MailListSettings[$distGroupName].scoutnet_list_id
            Write-SNSLog " "
            Write-SNSLog "Adding Contacts in distribution group $($distGroupName)"

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
                $MemberData.mailaddresses | ForEach-Object {
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

            $ledare_synk_option = $MailListSettings[$distGroupName].ledare_synk_option

            Write-SNSLog "Ledare synk option is '$ledare_synk_option' for distribution group $distGroupName"
            # Get the settings for ledare in this list.
            $AddLedareOffice365Address = $True
            $AddLedareScoutnetAddress = $False
            if ($ledare_synk_option -like "-")
            {
                $AddLedareScoutnetAddress = $True
                $AddLedareOffice365Address = $False
            }
            elseif ($ledare_synk_option -like "&")
            {
                $AddLedareScoutnetAddress = $True
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

                if ($AddLedareOffice365Address)
                {
                    $memberSearchStr = "*$member"
                    $recipient = Get-Recipient | Where-Object {$_.CustomAttribute1 -like $memberSearchStr}

                    if ($recipient)
                    {
                        try
                        {
                            Write-SNSLog "Adding member $($recipient.DisplayName) with id $($recipient.Id) to distribution group $distGroupName"
                            Add-DistributionGroupMember -Identity $distGroupName -Member $recipient.Id  -ErrorAction "stop"
                        }
                        catch
                        {
                            Write-SNSLog -Level "Warn" "Could not add contact $($recipient.DisplayName) to $distGroupName. Error $_"
                        }
                    }
                    else
                    {
                        $MemberData = $AllMailAddresses[$member]
                        $displayName = "$($MemberData.first_name) $($MemberData.last_name)"
                        Write-SNSLog "Adding $displayName to distribution group $distGroupName"
                        $email = "$($MemberData.first_name).$($MemberData.last_name)@$($DomainName)".ToLower()

                        # Convert UTF encoded names and create corresponding ASCII version.
                        $email = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($email))
                        try
                        {
                            Add-DistributionGroupMember -Identity $distGroupName -Member $email  -ErrorAction "stop"   
                        }
                        catch
                        {
                            Write-SNSLog -Level "Warn" "Could not add contact $email to $distGroupName. Error $_"
                        }
                    }
                }

                if ($AddLedareScoutnetAddress)
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

Export-ModuleMember Write-SNSLog, ConvertTo-SNSJSONHash, Receive-SNSApiJson, Get-SNSApiGroupCustomlists, Get-SNSUserEmails, Get-SNSMaillistInfo, SNSUpdateExchangeDistributionGroups
