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

    [OutputType([System.Collections.Hashtable])]
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