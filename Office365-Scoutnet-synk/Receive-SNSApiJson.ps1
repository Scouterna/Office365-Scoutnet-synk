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

    .PARAMETER FileName
        Optional name of the JSON file to save.

    .PARAMETER SaveFile
        Save the downloaded JSON to the configured directory.

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
        [pscredential]$Credential,

        [Parameter(Mandatory=$False, HelpMessage="Name of the JSON file to save.")]
        [string]$FileName,

        [Parameter(Mandatory=$False, HelpMessage="Save the downloaded JSON file.")]
        [bool]$SaveFile = $false
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
    $Response = Invoke-WebRequest -Uri $Uri -Headers $Headers -UseBasicParsing -ErrorAction "Stop"
    $JsonContent = if ($Response -is [string]) { $Response } else { $Response.Content }

    if ($SaveFile)
    {
        if ([string]::IsNullOrWhiteSpace($FileName) -or
            [System.IO.Path]::GetFileName($FileName) -ne $FileName -or
            [System.IO.Path]::IsPathRooted($FileName) -or
            $FileName -in @('.', '..'))
        {
            throw "FileName must be a file name without a directory path when SaveFile is enabled."
        }

        $Directory = $script:SNSConf.JsonFileDirectory
        if ([string]::IsNullOrWhiteSpace($Directory))
        {
            throw "SNSConfiguration.JsonFileDirectory must be set when saving API JSON files."
        }

        if (-not (Test-Path -LiteralPath $Directory -PathType Container))
        {
            New-Item -Path $Directory -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        $FilePath = Join-Path -Path $Directory -ChildPath $FileName
        [System.IO.File]::WriteAllText(
            $FilePath,
            $JsonContent,
            [System.Text.UTF8Encoding]::new($false)
        )
    }

    $JsonHashTable = $JsonContent | ConvertFrom-Json -ErrorAction "Stop" | ConvertTo-SNSJSONHash
    return $JsonHashTable
}