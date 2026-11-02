function Disconnect-SnSOffice365
{
    <#
    .SYNOPSIS
        Disconnects from Microsoft Graph and Exchange Online.

    .DESCRIPTION
        Signs out from Exchange Online and Microsoft Graph. Run it when the synchronization is done.
        A failed sign-out from Exchange Online is logged. A failed sign-out from Microsoft Graph
        is logged and then thrown as an error.

    .INPUTS
        None. You cannot pipe objects to Disconnect-SnSOffice365.

    .OUTPUTS
        None. You cannot pipe objects from Disconnect-SnSOffice365.

    .EXAMPLE
        Disconnect-SnSOffice365

        Signs out from both services.
    #>

    try
    {
        # Logga ut ifrån ExchangeOnline.
        Disconnect-ExchangeOnline -ErrorAction "Stop" -Confirm:$false -Verbose:$false
    }
    Catch
    {
        Write-SNSLog -Level "Error" "Utloggning ifrån ExchangeOnline returnerade felet $_"
    }

    try
    {
        # Logga ut från Microsoft Graph
        Disconnect-MgGraph -ErrorAction "Stop" -Verbose:$false |out-null
    }
    Catch
    {
        Write-SNSLog -Level "Error" "Disconnect-MgGraph returnerade felet $_"
        throw
    }
}
