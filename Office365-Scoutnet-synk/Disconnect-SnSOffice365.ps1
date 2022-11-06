function Disconnect-SnSOffice365
{
    <#
    .SYNOPSIS
        Disonnects from Office 365.

    .DESCRIPTION
        Disconnects to Office 365.

    .INPUTS
        None. You cannot pipe objects to Disconnect-SnSOffice365.

    .OUTPUTS
        None. You cannot pipe objects from Disconnect-SnSOffice365.

    .EXAMPLE
        Disconnect-SnSOffice365
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
        # Logga ut ifrån office 365
        Disconnect-MgGraph -ErrorAction "Stop" -Verbose:$false |out-null
    }
    Catch
    {
        Write-SNSLog -Level "Error" "Disconnect-MgGraph returnerade felet $_"
        throw
    }
}
