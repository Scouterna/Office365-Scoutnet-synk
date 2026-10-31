function Receive-SNSApiJson
{
    <#
    .SYNOPSIS
        Fetches the API data from scoutnet using the provided credentials.

    .DESCRIPTION
        Scoutnet API returns a JSON with the resulting data.
        This function uses Invoke-WebRequest to fetch the data and returns the
        JSON as a multilevel hashtable.
        The credential for the API is configured in scoutnet.

    .INPUTS
        None. You cannot pipe objects to Receive-SNSApiJson.

    .OUTPUTS
        A multilevel dictionary.

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

    [OutputType([System.Collections.IDictionary])]
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

    $Response = Invoke-WebRequest `
        -Uri $Uri `
        -Authentication Basic `
        -Credential $Credential `
        -ErrorAction Stop
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

    return $JsonContent | ConvertFrom-Json -AsHashtable -ErrorAction Stop
}