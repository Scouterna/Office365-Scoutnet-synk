function ConvertTo-SNSCanonicalValue
{
    param (
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value)
    {
        return $null
    }

    if ($Value -is [System.Collections.IDictionary])
    {
        $keys = [string[]]@($Value.Keys | ForEach-Object { [string]$_ })
        [Array]::Sort($keys, [System.StringComparer]::Ordinal)
        $canonicalValue = [ordered]@{}
        foreach ($key in $keys)
        {
            $canonicalValue[$key] = ConvertTo-SNSCanonicalValue -Value $Value[$key]
        }
        return ,$canonicalValue
    }

    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string])
    {
        $items = @($Value | ForEach-Object { ConvertTo-SNSCanonicalValue -Value $_ })
        [Array]::Sort($items, [System.StringComparer]::Ordinal)
        return ,$items
    }

    return $Value
}

function Get-SNSStableHash
{
    param (
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$InputObject
    )

    $canonicalValue = ConvertTo-SNSCanonicalValue -Value $InputObject
    $json = ConvertTo-Json -InputObject $canonicalValue -Depth 100 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $hashBytes = [System.Security.Cryptography.SHA256]::HashData($bytes)
    return [System.Convert]::ToHexString($hashBytes)
}
