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
    [OutputType([System.Collections.Hashtable])]
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
