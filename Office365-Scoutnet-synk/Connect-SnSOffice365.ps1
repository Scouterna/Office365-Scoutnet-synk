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

    .PARAMETER ConnectionParameters
        Connection parameters

    .EXAMPLE
        $ConnectionParameters = @{
            UserPrincipalName = "username@domain"
        }
        Connect-SnSOffice365 -ConnectionParameters $ConnectionParameters
    #>

    param (
        [Parameter(Mandatory=$True, HelpMessage="Connection parameters for Connect-SnSOffice365")]
        [ValidateNotNull()]
        $ConnectionParameters
        )

    try
    {
        # Logga in på ExchangeOnline med ExchangeOnlineManagement modulen.
        $RequiredScopes = @("Directory.AccessAsUser.All",
                            "Directory.ReadWrite.All",
                            "Directory.Read.All",
                            "GroupMember.Read.All",
                            "GroupMember.ReadWrite.All",
                            "Group.ReadWrite.All",
                            "Group.Read.All"
                            "User.ReadWrite.All",
                            “User.Read.All”,
                            "Mail.Send")
        Connect-MgGraph -Scopes $RequiredScopes -ErrorAction "Stop"
    }
    Catch
    {
        Write-SNSLog -Level "Error" "Kunde logga in på ExchangeOnline Error $_"
        throw
    }

    try
    {
        # Logga in på ExchangeOnline med ExchangeOnlineManagement modulen.
        Connect-ExchangeOnline @ConnectionParameters -ErrorAction "Stop" -Verbose:$false -CommandName Get-EXOMailbox,Get-EXORecipient,Get-DistributionGroupMember,Get-DistributionGroup,Update-DistributionGroupMember,New-MailContact,Set-Contact,Set-MailContact,Remove-MailContact,Set-MailContact,Set-Mailbox,Remove-DistributionGroupMember,Add-DistributionGroupMember,Get-DistributionGroupMember,Get-DistributionGroup,Set-MailboxMessageConfiguration,Set-MailboxAutoReplyConfiguration
    }
    Catch
    {
        Write-SNSLog -Level "Error" "Kunde logga in på ExchangeOnline Error $_"
        throw
    }
}
