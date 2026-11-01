function Connect-SnSOffice365
{
    <#
    .SYNOPSIS
        Connects to Microsoft Graph and Exchange Online.

    .DESCRIPTION
        Signs in to Microsoft Graph and Exchange Online. Run it once before the synchronization functions.

        The sign-in uses the domain name, scopes and allowed Exchange commands from the configuration.
        Exchange Online is connected with only the commands the module needs (CommandName in the configuration).

        Interactive (default): uses device code sign-in for both services. Follow the instructions
        written to the console and sign in with an account that is allowed to make the changes.

        With -ManagedIdentity: uses the system-assigned managed identity of the Azure Automation account.
        The identity needs the Graph application permissions, the Exchange.ManageAsApp permission and
        the Exchange Administrator role, which setupAzure.ps1 gives it.

        Throws an error if no configuration is available, or if the sign-in fails.

    .INPUTS
        None. You cannot pipe objects to Connect-SnSOffice365.

    .OUTPUTS
        None. You cannot pipe objects from Connect-SnSOffice365.

    .LINK
        https://github.com/scouternasetjanster/Office365-Scoutnet-synk

    .PARAMETER Configuration
        Configuration to use, created with New-SNSConfiguration. It is cached in the module,
        so it can be left out if a configuration has already been given in this session.

    .PARAMETER ManagedIdentity
        Sign in with the managed identity of the Azure Automation account instead of device code.
        Use this when the synchronization runs in an Azure runbook.

    .EXAMPLE
        $conf = New-SNSConfiguration -LicenseAssignment $licenseAssignment
        $conf.DomainName = 'example.se'
        Connect-SnSOffice365 -Configuration $conf

        Signs in interactively with device code, using the given configuration.

    .EXAMPLE
        Connect-SnSOffice365 -ManagedIdentity -Configuration $conf -ErrorAction Stop

        Signs in with the managed identity in an Azure runbook.
    #>

    [CmdletBinding(HelpURI = 'https://github.com/scouternasetjanster/Office365-Scoutnet-synk',
                PositionalBinding = $False)]
    param (
        [Parameter(Mandatory=$False, HelpMessage="Configuration to use. If not specified the cached configuration will be used.")]
        $Configuration,
        [Parameter(HelpMessage="Sign in with the managed identity of the Azure Automation account instead of device code.")]
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

    if ($ManagedIdentity)
    {
        Connect-AzAccount -Identity -ErrorAction "Stop" |out-null
        $token = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -ErrorAction "Stop").token
        Connect-MgGraph -AccessToken $token -ErrorAction "Stop" |out-null
    }
    else
    {
        Connect-MgGraph -UseDeviceCode -ContextScope Process -Scopes $Script:SNSConf.RequiredScopes -ErrorAction "Stop"
    }

    try
    {
        if ($ManagedIdentity)
        {
            Connect-ExchangeOnline -ManagedIdentity -Organization $Script:SNSConf.DomainName -ShowBanner:$false -CommandName $Script:SNSConf.commandNames -Verbose:$false -ErrorAction "Stop"
        }
        else
        {
            Connect-ExchangeOnline -Device -ShowBanner:$false -CommandName $Script:SNSConf.commandNames -Verbose:$false -ErrorAction "Stop"
        }
    }
    Catch
    {
        Write-SNSLog -Level "Error" "Kunde logga in på ExchangeOnline Error $_"
        throw
    }
}
