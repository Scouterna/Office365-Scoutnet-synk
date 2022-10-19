function Connect-SnSOffice365
{
    <#
    .SYNOPSIS
        Connects to Office 365 using provided parameters.

    .DESCRIPTION
        Connects to Office 365.

    .INPUTS
        None. You cannot pipe objects to Connect-SnSOffice365.

    .OUTPUTS
        None. You cannot pipe objects from Connect-SnSOffice365.

    .LINK
        https://github.com/scouternasetjanster/Office365-Scoutnet-synk

    .PARAMETER Configuration
        Configuration to use. If not specified the cached configuration will be used.
    #>

    [CmdletBinding(HelpURI = 'https://github.com/scouternasetjanster/Office365-Scoutnet-synk',
                PositionalBinding = $False)]
    param (
        [Parameter(Mandatory=$False, HelpMessage="Configuration to use. If not specified the cached configuration will be used.")]
        $Configuration,
        [Parameter(HelpMessage="Enable use of managed identity when running in Azure runnbooks.")]
        [switch]$ManagedIdentity = $false
    )

    if ($Configuration)
    {
        $Script:SNSConf = $Configuration
    }

    if (!$Script:SNSConf)
    {
        throw "No configuration specified. Please provide a configuration!"
    }

    if ($ManagedIdentity)
    {
        Write-SNSLog "Using managed identity for domain $($Script:SNSConf.DomainName)"
    }

    try
    {
        if ($ManagedIdentity)
        {
            $token = (Get-AzAccessToken -ResourceTypeName MSGraph  -ErrorAction "Stop").token
            Connect-MgGraph -AccessToken $token -Scopes $Script:SNSConf.RequiredScopes -ErrorAction "Stop"
        }
        else
        {
            Connect-MgGraph -ContextScope Process -Scopes $Script:SNSConf.RequiredScopes -ErrorAction "Stop"
        }
    }
    Catch
    {
        Write-SNSLog -Level "Error" "Kunde logga in på MgGraph Error $_"
        throw
    }

    try
    {
        if ($ManagedIdentity)
        {
            Connect-ExchangeOnline -ManagedIdentity -Organization $Script:SNSConf.DomainName -ShowBanner:$false -CommandName $Script:SNSConf.commandNames -Verbose:$false -ErrorAction "Stop"
        }
        else
        {
            Connect-ExchangeOnline -ShowBanner:$false -CommandName $Script:SNSConf.commandNames -Verbose:$false -ErrorAction "Stop"
        }
    }
    Catch
    {
        Write-SNSLog -Level "Error" "Kunde logga in på ExchangeOnline Error $_"
        throw
    }
}
