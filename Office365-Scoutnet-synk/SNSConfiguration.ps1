class SNSConfiguration
{
    [string]$LogFilePath='scoutnetSync.log'
    [string]$SyncGroupName='scoutnet'
    [string]$SyncGroupDescription="Säkerhetsgrupp som används vid synkronisering med Scoutnet. Användare i gruppen avaktiveras om de inte är kvar i Scoutnet."
    [string]$SyncGroupDisabledUsersName='scoutnetDisabledUsers'
    [string]$SyncGroupDisabledUsersDescription="Säkerhetsgrupp som används vid synkronisering med Scoutnet. Användare i gruppen är avaktiverade och finns inte längre med i Scoutnet."
    [string]$AllUsersGroupName=""
    hidden [string[]]$LicenseAssignment =@()
    hidden [System.Array]$LicenseOptions =@()
    [string]$PreferredLanguage="sv-SE"
    [string]$UsageLocation="SE"
    [string]$WaitMailboxCreationMaxTime="1200"
    [string]$WaitMailboxCreationPollTime="30"
    [string]$SignatureText=""
    [string]$SignatureHtml=""
    [string]$NewUserEmailSubject=""
    [string]$NewUserEmailText=""
    [string]$EmailSMTPServer = "outlook.office365.com"
    [string]$SmtpPort = '587'
    [string]$NewUserInfoEmailSubject=""
    [string]$NewUserInfoEmailText=""
    [string]$EmailFromAddress = ""
    [string]$UserSyncMailListId
    [pscredential]$CredentialCustomlists
    [pscredential]$CredentialMemberlist
    [pscredential]$Credential365
    [System.Collections.Hashtable]$MailListSettings
    [string]$DomainName
    [string]$DisabledAccountsAutoReplyText
    [System.Collections.Hashtable]$ApiGroupMemberlistCache
}

function New-SNSConfiguration
{
    <#
    .SYNOPSIS
        Creates a new configuration.

    .INPUTS
        None. You cannot pipe objects to Get-SNSConfiguration.

    .OUTPUTS
        The created configuration

    .PARAMETER CredentialMemberlist
        Credentials for api/group/memberlist

    .PARAMETER CredentialCustomlists
        Credentials for api/group/customlists

    .PARAMETER Credential365
        Credentials for office365 that can execute needed servlets.

    .PARAMETER MailListSettings
        Maillists to process for maillist syncronisation. A hashtable with maillist info.

    .PARAMETER DomainName
        Domain name for office365 mail addresses.

    .PARAMETER LicenseAssignment
        License assignment data to use when creating a new account.

    .PARAMETER UserSyncMailListId
        List Id to use when syncronising user accounts. If it is empty (default value) then all members with a role will be given an account.

    .PARAMETER LogFilePath
        Logfile name and path. Used to override the default value.

    .PARAMETER AllUsersGroupName
        Name of distribution group to add all new accounts to. The distribution group must exist.

    .PARAMETER SignatureText
        Outlook online signature for new users. Text version.

    .PARAMETER SignatureHtml
        Outlook online signature for new users. Html version.

    .PARAMETER NewUserEmailSubject
        Subject for the mail to the user about the new account. Is sent to the member Scoutnet primary e-mail address.

    .PARAMETER NewUserEmailText
        Body for the mail to the user about the new account. Is sent to the member Scoutnet primary e-mail address.

    .PARAMETER NewUserInfoEmailSubject
        Subject for the mail to the users new e-mail address. Can be used to inform the user about the system.

    .PARAMETER NewUserInfoEmailText
        Body for the mail to the users new e-mail address. Can be used to inform the user about the system.

    .PARAMETER EmailFromAddress
        From address for all mails. The authenticaion account must be able to send as this address.
        If not set the name for Credential365 is used as from address.

    .PARAMETER DisabledAccountsAutoReplyText
        Autoreply message for disabled accounts.
        If set this text is set as autoreply message for disabled accounts. Html is supported.
    #>
    [OutputType([SNSConfiguration])]
    param
    (
        [Parameter(HelpMessage="Credentials for api/group/memberlist.")]
        [pscredential]$CredentialMemberlist,

        [Parameter(HelpMessage="Credentials for api/group/customlist.")]
        [pscredential]$CredentialCustomlists,

        [Parameter(HelpMessage="Credentials for office365.")]
        [pscredential]$Credential365,

        [Parameter(HelpMessage="Maillists to process.")]
        $MailListSettings,

        [Parameter(HelpMessage="Domain name for office365 mail addresses.")]
        [string]$DomainName,

        [Parameter(HelpMessage="License data.")]
        $LicenseAssignment,

        [Parameter(HelpMessage="List Id to use when syncronising user accounts.")]
        $UserSyncMailListId,

        [Parameter(HelpMessage="Logfile name and path.")]
        [string]$LogFilePath,

        [Parameter(HelpMessage="Name of distribution group to add all new accounts to.")]
        [string]$AllUsersGroupName,

        [Parameter(HelpMessage="Outlook online signature for new users. Text version.")]
        [string]$SignatureText,

        [Parameter(HelpMessage="Outlook online signature for new users. Html version.")]
        [string]$SignatureHtml,

        [Parameter(HelpMessage="Subject for the mail to the user about the new account.")]
        [string]$NewUserEmailSubject,

        [Parameter(HelpMessage="Body for the mail to the user about the new account.")]
        [string]$NewUserEmailText,

        [Parameter(HelpMessage="Subject for the mail to the users new e-mail address.")]
        [string]$NewUserInfoEmailSubject,

        [Parameter(HelpMessage="Body for the mail to the users new e-mail address.")]
        [string]$NewUserInfoEmailText,

        [Parameter(HelpMessage="Autoreply message for disabled accounts.")]
        [string]$DisabledAccountsAutoReplyText
    )

    $conf = [SNSConfiguration]::new()

    if ($NewUserInfoEmailText)
    {
        $conf.NewUserInfoEmailText = $NewUserInfoEmailText
    }

    if ($NewUserInfoEmailSubject)
    {
        $conf.NewUserInfoEmailSubject = $NewUserInfoEmailSubject
    }

    if ($NewUserEmailText)
    {
        $conf.NewUserEmailText = $NewUserEmailText
    }

    if ($NewUserEmailSubject)
    {
        $conf.NewUserEmailSubject = $NewUserEmailSubject
    }

    if ($SignatureHtml)
    {
        $conf.SignatureHtml = $SignatureHtml
    }

    if ($SignatureText)
    {
        $conf.SignatureText = $SignatureText
    }

    if ($AllUsersGroupName)
    {
        $conf.AllUsersGroupName = $AllUsersGroupName
    }

    if ($LogFilePath)
    {
        $conf.LogFilePath = $LogFilePath
    }

    if ($CredentialMemberlist)
    {
        $conf.CredentialMemberlist = $CredentialMemberlist
    }
    
    if ($CredentialCustomlists)
    {
        $conf.CredentialCustomlists = $CredentialCustomlists
    }

    if ($Credential365)
    {
        $conf.Credential365 = $Credential365
    }

    if ($DomainName)
    {
        $conf.DomainName = $DomainName
    }

    if ($UserSyncMailListId)
    {
        $conf.UserSyncMailListId = $UserSyncMailListId
    }

    if ($DisabledAccountsAutoReplyText)
    {
        $conf.DisabledAccountsAutoReplyText = $DisabledAccountsAutoReplyText
    }

    if ($LicenseAssignment)
    {
        # Create licensing options.
        foreach($LicenseKey in $LicenseAssignment.Keys)
        {
            $conf.LicenseAssignment += "$LicenseKey"
            try
            {
                if (![string]::IsNullOrWhiteSpace($LicenseAssignment[$LicenseKey]))
                {
                    $LO = New-MsolLicenseOptions -AccountSkuId $LicenseKey -DisabledPlans $LicenseAssignment[$LicenseKey] -ErrorAction "Stop"
                    $conf.LicenseOptions += $LO
                }
                else
                {
                    $LO = New-MsolLicenseOptions -AccountSkuId $LicenseKey -DisabledPlans $null -ErrorAction "Stop"
                    $conf.LicenseOptions += $LO
                }
            }
            catch
            {
                throw "Could not create MsolLicenseOptions. Error: $_"
            }
        }
    
        if ($conf.LicenseOptions.Count -eq 0)
        {
            $msg = "The parameter 'SNSLicenseAssignment' did not contain any valid licenses."
            $msg += "Creation of account cannot be executed!"
            throw ($msg)
        }    
    }
    return $conf
}

function Get-SNSConfiguration
{
    <#
    .SYNOPSIS
        Fetch current configuration.

    .INPUTS
        None. You cannot pipe objects to Get-SNSConfiguration.

    .OUTPUTS
        None.

    .PARAMETER new
        Create new empty configuration.
    #>
    [OutputType([SNSConfiguration])]
    param
    (
    )

    return $script:SNSConf
}

function Set-SNSConfiguration
{
    <#
    .SYNOPSIS
        Set a newconfiguration.

    .INPUTS
        None. You cannot pipe objects to Set-SNSConfiguration.

    .OUTPUTS
        None.

    .PARAMETER new
        Create new empty configuration.
    #>
    param
    (
        [Parameter(Mandatory=$True, HelpMessage="The new configuration.")]
        $Configuration
    )

    $script:SNSConf = $Configuration
}