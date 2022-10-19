class SNSConfiguration
{
    [string]$LogFilePath='scoutnetSync.log'
    [string]$SyncGroupName='scoutnet'
    [string]$SyncGroupDescription="Säkerhetsgrupp som används vid synkronisering med Scoutnet. Användare i gruppen avaktiveras om de inte är kvar i Scoutnet."
    [string]$SyncGroupDisabledUsersName='scoutnetDisabledUsers'
    [string]$SyncGroupDisabledUsersDescription="Säkerhetsgrupp som används vid synkronisering med Scoutnet. Användare i gruppen är avaktiverade och finns inte längre med i Scoutnet."
    [string]$AllUsersGroupName=""
    hidden [System.Collections.Hashtable]$LicenseAssignment
    [string]$PreferredLanguage="sv-SE"
    [string]$UsageLocation="SE"
    [string]$WaitMailboxCreationMaxTime="1200"
    [string]$WaitMailboxCreationPollTime="30"
    [string]$SignatureText=""
    [string]$SignatureHtml=""
    [string]$NewUserEmailSubject=""
    [string]$NewUserEmailText=""
    [string]$NewUserEmailContentType="Text"
    [string]$NewUserInfoEmailSubject=""
    [string]$NewUserInfoEmailText=""
    [string]$NewUserInfoEmailContentType="Text"
    [string]$EmailFromAddress = ""
    [string]$UserSyncMailListId
    [pscredential]$CredentialCustomlists
    [pscredential]$CredentialMemberlist
    [System.Collections.Hashtable]$MailListSettings
    [string]$DomainName
    [string]$DisabledAccountsAutoReplyText
    [System.Collections.Hashtable]$ApiGroupMemberlistCache
    hidden [String]$commandNames
    hidden [String[]]$RequiredScopes = @("Directory.AccessAsUser.All",
    "Directory.ReadWrite.All",
    "Directory.Read.All",
    "GroupMember.Read.All",
    "GroupMember.ReadWrite.All",
    "Group.ReadWrite.All",
    "Group.Read.All"
    "User.ReadWrite.All",
    “User.Read.All”,
    "Mail.Send")
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

    .PARAMETER NewUserEmailContentType
        Content type for the email, supported types is HTML and Text. Default is Text.

    .PARAMETER NewUserInfoEmailSubject
        Subject for the mail to the users new e-mail address. Can be used to inform the user about the system.

    .PARAMETER NewUserInfoEmailText
        Body for the mail to the users new e-mail address. Can be used to inform the user about the system.

    .PARAMETER NewUserInfoEmailContentType
        Content type for the email, supported types is HTML and Text. Default is Text.

    .PARAMETER EmailFromAddress
        From address for all mails. The authenticaion account must be able to send as this address.

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

        [Parameter(HelpMessage="Maillists to process.")]
        $MailListSettings,

        [Parameter(HelpMessage="Domain name for office365 mail addresses.")]
        [string]$DomainName,

        [Parameter(Mandatory=$true, HelpMessage="License data.")]
        [ValidateNotNull()]
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

        [Parameter(HelpMessage="Content type for the email, supported types is Html and Text. Default is Text.")]
        [string]$NewUserEmailContentType,

        [Parameter(HelpMessage="Subject for the mail to the users new e-mail address.")]
        [string]$NewUserInfoEmailSubject,

        [Parameter(HelpMessage="Body for the mail to the users new e-mail address.")]
        [string]$NewUserInfoEmailText,

        [Parameter(HelpMessage="Content type for the email, supported types is Html and Text. Default is Text.")]
        [string]$NewUserInfoEmailContentType,

        [Parameter(HelpMessage="Autoreply message for disabled accounts.")]
        [string]$DisabledAccountsAutoReplyText
    )

    $conf = [SNSConfiguration]::new()

    if ($MailListSettings)
    {
        $conf.MailListSettings = $MailListSettings
    }

    if ($NewUserInfoEmailText)
    {
        $conf.NewUserInfoEmailText = $NewUserInfoEmailText
    }

    if ($NewUserInfoEmailContentType)
    {
        $conf.NewUserInfoEmailContentType = $NewUserInfoEmailContentType
    }

    if ($NewUserInfoEmailSubject)
    {
        $conf.NewUserInfoEmailSubject = $NewUserInfoEmailSubject
    }

    if ($NewUserEmailText)
    {
        $conf.NewUserEmailText = $NewUserEmailText
    }

    if ($NewUserEmailContentType)
    {
        $conf.NewUserEmailContentType = $NewUserEmailContentType
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
        $conf.LicenseAssignment = $LicenseAssignment
    }

    # Exchange online commands to use.
    $conf.commandNames = "Get-EXOMailbox,Get-EXORecipient,"
    $conf.commandNames += "Get-DistributionGroupMember,Get-DistributionGroup,"
    $conf.commandNames += "Update-DistributionGroupMember,New-MailContact,"
    $conf.commandNames += "Set-Contact,Set-MailContact,Remove-MailContact,"
    $conf.commandNames += "Set-MailContact,Set-Mailbox,Remove-DistributionGroupMember,"
    $conf.commandNames += "Add-DistributionGroupMember,Get-DistributionGroupMember,"
    $conf.commandNames += "Get-DistributionGroup,Set-MailboxMessageConfiguration,"
    $conf.commandNames += "Set-MailboxAutoReplyConfiguration"

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