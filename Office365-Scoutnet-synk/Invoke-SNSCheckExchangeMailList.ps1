function Invoke-SNSCheckExchangeMailList
{
    <#
    .SYNOPSIS
        Checks and if needed creates exchange distribution group.

    .DESCRIPTION
        Validates the settings for an exchange distribution group and creates the group if it is missing.

    .INPUTS
        None. You cannot pipe objects to Get-SNSExchangeMailListMember.

    .OUTPUTS
        None.

    .PARAMETER ExchangeSession
        Exchange session to use for Import-PSSession.

    .PARAMETER Settings
        Settings for the distribution group.
    #>

    param (
        [Parameter(Mandatory=$True, HelpMessage="Exchange session to use.")]
        $ExchangeSession,

        [Parameter(Mandatory=$True, HelpMessage="Settings for the distribution group.")]
        $Settings
        )
        
    try
    {
        Import-PSSession $ExchangeSession -AllowClobber -CommandName New-DistributionGroup,Get-DistributionGroup -ErrorAction Stop > $null
    }
    catch
    {
        Write-SNSLog -Level "Error" "Could not import needed functions. Error $_"
        throw
    }

    try
    {
        $mailListGroup = Get-DistributionGroup $mailList -ErrorAction Stop
    }
    catch
    {
        Write-SNSLoggit  "Could not fetch data for distribution group '$mailList'. Trying to create the list."
        try
        {
            $GroupType = "Distribution"
            if ($Settings.IsSecurityGroup)
            {
                $GroupType = "Security"
            }
            $mailListGroup = New-DistributionGroup -Name $Settings.Name -DisplayName $Settings.DisplayName -Type $GroupType -MemberDepartRestriction "Closed" `
                                -MemberJoinRestriction  "Closed" -ErrorAction Stop
            #$mailListGroup = $NewListGroup.ExchangeObjectId
        }
        catch
        {
            Write-SNSLog -Level "Error" "Could not create distribution group. Error $_"
            throw
        }
    }

    try
    {
        Write-SNSLog "Updating distribution list $($group.DisplayName)"
    }
    catch
    {
        Write-SNSLog -Level "Error" "Could not update group data. Error $_"
        throw
    }
    Write-SNSLog "Done"
}