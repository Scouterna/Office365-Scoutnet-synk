# Exchange Online can take a while to see a newly created mailbox, which gives an error that
# the object could not be found. Retry only that error.
function Invoke-SNExchangeRetry
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Action,
        [int]$MaxAttempts = 12,
        [int]$DelaySeconds = 5
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++)
    {
        try
        {
            return & $Action
        }
        catch
        {
            $isNotFound = "$($_.Exception.Message)" -match "couldn.t be found|can.t be found|ManagementObjectNotFound|was not found"
            if (-not $isNotFound -or $attempt -ge $MaxAttempts) { throw }
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

function Get-SNSharedMailboxPermissionPlan
{
    <#
    .SYNOPSIS
        Works out which accounts to add to and remove from one kind of shared mailbox permission.

    .DESCRIPTION
        Only accounts that exist in KnownUsers are ever planned for removal. Permissions that
        belong to anything else, for example groups, service accounts or deleted accounts, are
        never touched. Accounts in ProtectedUserIds are kept even if they do not match the rule.

    .PARAMETER ExistingNames
        Identities that currently have the permission, as returned by Exchange.

    .PARAMETER DesiredUserIds
        Hashtable with the ExternalDirectoryObjectId of every account that should have the
        permission as keys.

    .PARAMETER ProtectedUserIds
        Hashtable with the ExternalDirectoryObjectId of accounts that are never removed.

    .PARAMETER KnownUsers
        Hashtable that maps lower case user principal names and primary SMTP addresses to the
        mailbox of the account.

    .PARAMETER UsersById
        Hashtable that maps ExternalDirectoryObjectId to the mailbox of the account.

    .OUTPUTS
        PSCustomObject with Add, Remove and Kept, each a list of account mailboxes.
    #>
    param (
        [Parameter()]
        [string[]]$ExistingNames = @(),

        [Parameter(Mandatory)]
        [hashtable]$DesiredUserIds,

        [Parameter(Mandatory)]
        [hashtable]$ProtectedUserIds,

        [Parameter(Mandatory)]
        [hashtable]$KnownUsers,

        [Parameter(Mandatory)]
        [hashtable]$UsersById
    )

    $existingUserIds = @{}
    foreach ($name in $ExistingNames)
    {
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $user = $KnownUsers[$name.Trim().ToLowerInvariant()]
        if ($null -ne $user)
        {
            $existingUserIds[[string]$user.ExternalDirectoryObjectId] = $user
        }
    }

    $add = @()
    foreach ($userId in $DesiredUserIds.Keys)
    {
        if (-not $existingUserIds.ContainsKey([string]$userId))
        {
            $add += $UsersById[[string]$userId]
        }
    }

    $remove = @()
    $kept = @()
    foreach ($userId in $existingUserIds.Keys)
    {
        if ($DesiredUserIds.ContainsKey([string]$userId)) { continue }
        if ($ProtectedUserIds.ContainsKey([string]$userId))
        {
            $kept += $existingUserIds[$userId]
        }
        else
        {
            $remove += $existingUserIds[$userId]
        }
    }

    return [pscustomobject]@{
        Add = $add
        Remove = $remove
        Kept = $kept
    }
}

function Update-SNSharedMailbox
{
    <#
    .SYNOPSIS
        Creates configured shared mailboxes and synchronizes who can use them with Scoutnet.

    .DESCRIPTION
        Only shared mailboxes explicitly configured in SharedMailboxSettings are managed. Other
        mailboxes are never read for changes or modified.

        SharedMailboxSettings is a hashtable where each key is the display name of the shared
        mailbox and each value is a rule with these entries:

        Required for all rules
          mail_nickname  Alias of the mailbox. It identifies the mailbox and is used for the
                         e-mail address when the mailbox is created. Letters, digits, '.', '_'
                         and '-' are allowed.
        Optional for all rules
          source         'mail_list', 'troop_roles' or 'group_roles'. Selects the members in the
                         same way as for Microsoft365GroupSettings. Omit to manage only
                         keep_members without a Scoutnet connection.
          email_address  Primary e-mail address, used when the mailbox is created. Default is
                         mail_nickname@DomainName, so DomainName must be configured when it is
                         left out.
          full_access    Give the members Full Access (read and manage). Default $true. $false
                         means that Full Access is not managed at all.
          send_as        Give the members Send As (send mail as the mailbox). Default $true.
                         $false means that Send As is not managed at all.
          auto_mapping   Make the mailbox show up automatically in Outlook for the members given
                         Full Access. Default $true. Only used when a member is added.
          hidden_from_address_list  Set to hide the mailbox from the address lists. Only
                         changed when configured.
          copy_sent_items  Set to $true to keep a copy of the mail sent with Send As in the
                         mailbox. Only changed when configured.
          keep_members   Accounts (user principal names or object ids) to give the configured
                         permissions to and never remove. Required when source is omitted.

        source = 'mail_list', 'group_roles' and 'troop_roles' take the same fields as for
        Update-SNMicrosoft365Group.

        The function validates all rules first, fetches Scoutnet data only when a rule has a
        source, fetches the Exchange mailboxes, and only then changes Exchange Online. A member
        is mapped to an account through the mailbox CustomAttribute1, which holds the Scoutnet
        member number.
        A member matching a rule without an account is written to the log and skipped.

        The mailbox is found through its alias. It is created as a shared mailbox if it does not
        exist, and a shared mailbox is never removed. An alias that belongs to something else
        than a shared mailbox causes an error and the mailbox is not changed. Accounts that
        match the rule and lack a permission are given it, and accounts that no longer match
        lose it. Only permissions of accounts with a Microsoft 365 mailbox are removed. Permissions
        of groups and other recipients are never removed. Accounts in keep_members are given the
        configured permissions and are never removed. When source is omitted, membership and
        permissions are managed only from keep_members, and no Scoutnet data is fetched.

        Full Access can take up to an hour before it is visible in Outlook and Send As up to
        15 minutes.

    .INPUTS
        None. You cannot pipe objects to Update-SNSharedMailbox.

    .OUTPUTS
        None.

    .PARAMETER Configuration
        Configuration to use. If omitted, the cached module configuration is used. A supplied
        configuration also becomes the cached configuration.

    .PARAMETER WhatIf
        Shows the mailbox creations and permission changes without applying them.

    .EXAMPLE
        $conf.SharedMailboxSettings = @{
            'Kassoren' = @{
                source        = 'group_roles'
                mail_nickname = 'kassoren'
                roles         = @('Kassor')
                copy_sent_items = $true
            }
        }
        Update-SNSharedMailbox -Configuration $conf

    .EXAMPLE
        $conf.SharedMailboxSettings = @{
            'Info' = @{
                source           = 'mail_list'
                mail_nickname    = 'info'
                scoutnet_list_id = '1234'
                full_access      = $false
                keep_members     = @('admin@example.org')
            }
        }
        Update-SNSharedMailbox -Configuration $conf -WhatIf

    .EXAMPLE
        $conf.SharedMailboxSettings = @{
            'Extern' = @{
                mail_nickname = 'extern'
                keep_members  = @('person@example.org', 'object-id')
            }
        }
        Update-SNSharedMailbox -Configuration $conf
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter()]
        $Configuration
    )

    $snsConfiguration = $Script:SNSConf
    if ($null -ne $Configuration)
    {
        $snsConfiguration = $Configuration
        $Script:SNSConf = $Configuration
    }
    if ($null -eq $snsConfiguration)
    {
        throw 'No configuration specified. Please provide a configuration.'
    }

    $mailboxSettings = $snsConfiguration.SharedMailboxSettings
    if ($null -eq $mailboxSettings -or $mailboxSettings.Count -eq 0)
    {
        Write-SNSLog 'No shared mailbox rules are configured. No shared mailboxes will be changed.'
        return
    }

    # Validate every rule before anything is fetched or changed.
    $usedAliases = @{}
    foreach ($mailboxName in $mailboxSettings.Keys)
    {
        $setting = $mailboxSettings[$mailboxName]
        $source = ([string](Get-SNMicrosoft365Value -Object $setting -Name 'source')).Trim().ToLowerInvariant()
        $alias = [string](Get-SNMicrosoft365Value -Object $setting -Name 'mail_nickname')

        if ([string]::IsNullOrWhiteSpace([string]$mailboxName))
        {
            throw 'Shared mailbox configuration keys must contain a display name.'
        }
        if ([string]::IsNullOrWhiteSpace($alias))
        {
            throw "Shared mailbox '$mailboxName' must configure mail_nickname."
        }
        if ($alias -notmatch '^[A-Za-z0-9._-]+$')
        {
            throw "Shared mailbox '$mailboxName' has an invalid mail_nickname '$alias'. Use letters, digits, '.', '_' and '-'."
        }
        if ($usedAliases.ContainsKey($alias.ToLowerInvariant()))
        {
            throw "Shared mailboxes '$($usedAliases[$alias.ToLowerInvariant()])' and '$mailboxName' use the same mail_nickname '$alias'."
        }
        $usedAliases[$alias.ToLowerInvariant()] = $mailboxName
        if ($source -notin @('', 'mail_list', 'troop_roles', 'group_roles'))
        {
            throw "Shared mailbox '$mailboxName' has unsupported source '$source'. Use mail_list, troop_roles, or group_roles."
        }
        if ([string]::IsNullOrWhiteSpace($source) -and
            @((Get-SNMicrosoft365Value -Object $setting -Name 'keep_members') |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -eq 0)
        {
            throw "Shared mailbox '$mailboxName' must configure keep_members when source is omitted."
        }
        if ([string]::IsNullOrWhiteSpace([string](Get-SNMicrosoft365Value -Object $setting -Name 'email_address')) -and
            [string]::IsNullOrWhiteSpace([string]$snsConfiguration.DomainName))
        {
            throw "Shared mailbox '$mailboxName' must configure email_address, or DomainName must be configured."
        }
        foreach ($flagName in @('full_access', 'send_as', 'auto_mapping', 'hidden_from_address_list', 'copy_sent_items'))
        {
            $flag = Get-SNMicrosoft365Value -Object $setting -Name $flagName
            if ($null -ne $flag -and $flag -isnot [bool])
            {
                throw "Shared mailbox '$mailboxName' has an invalid $flagName. Use `$true or `$false."
            }
        }
        if ((Get-SNMicrosoft365Value -Object $setting -Name 'full_access') -eq $false -and
            (Get-SNMicrosoft365Value -Object $setting -Name 'send_as') -eq $false)
        {
            throw "Shared mailbox '$mailboxName' must give at least one permission. Remove full_access = `$false or send_as = `$false."
        }

        if (-not [string]::IsNullOrWhiteSpace($source))
        {
            Assert-SNMicrosoft365SourceRule -Name ([string]$mailboxName) -Setting $setting -Source $source -Label 'Shared mailbox'
        }
    }

    $usesScoutnet = @($mailboxSettings.Values | Where-Object {
        -not [string]::IsNullOrWhiteSpace(([string](Get-SNMicrosoft365Value -Object $_ -Name 'source')).Trim())
    }).Count -gt 0
    $scoutnetData = if ($usesScoutnet)
    {
        Get-SNMicrosoft365ScoutnetData -Configuration $snsConfiguration -Rules $mailboxSettings -Label 'shared mailboxes'
    }

    # CustomAttribute1 holds the Scoutnet member number of each Microsoft 365 account.
    $allOffice365Users = @(
        Get-EXOMailbox -RecipientTypeDetails 'UserMailbox' `
            -Properties CustomAttribute1, UserPrincipalName, PrimarySmtpAddress `
            -Verbose:$false -ErrorAction Stop
    )
    $office365UsersByScoutnetId = @{}
    $usersById = @{}
    $knownUsers = @{}
    foreach ($office365User in $allOffice365Users)
    {
        if ([string]::IsNullOrWhiteSpace([string]$office365User.ExternalDirectoryObjectId))
        {
            continue
        }
        $usersById[[string]$office365User.ExternalDirectoryObjectId] = $office365User
        foreach ($address in @($office365User.UserPrincipalName, $office365User.PrimarySmtpAddress))
        {
            if (-not [string]::IsNullOrWhiteSpace([string]$address))
            {
                $knownUsers[([string]$address).ToLowerInvariant()] = $office365User
            }
        }
        $scoutnetId = [string]$office365User.CustomAttribute1
        if (-not [string]::IsNullOrWhiteSpace($scoutnetId))
        {
            $office365UsersByScoutnetId[$scoutnetId] = $office365User
        }
    }

    # Resolve desired members before touching Exchange Online.
    $desiredMemberIdsByMailbox = @{}
    $keepUserIdsByMailbox = @{}
    foreach ($mailboxName in $mailboxSettings.Keys)
    {
        $setting = $mailboxSettings[$mailboxName]
        $source = ([string](Get-SNMicrosoft365Value -Object $setting -Name 'source')).Trim().ToLowerInvariant()
        $desiredMemberIdsByMailbox[$mailboxName] = if ([string]::IsNullOrWhiteSpace($source))
        {
            @{}
        }
        else
        {
            Get-SNMicrosoft365DesiredMemberId -Setting $setting -ScoutnetData $scoutnetData
        }

        $keepUserIds = @{}
        foreach ($keepMember in @(Get-SNMicrosoft365Value -Object $setting -Name 'keep_members'))
        {
            if ([string]::IsNullOrWhiteSpace([string]$keepMember)) { continue }
            $keepId = ([string]$keepMember).Trim()
            $keepUser = $knownUsers[$keepId.ToLowerInvariant()]
            if ($null -ne $keepUser)
            {
                $keepId = [string]$keepUser.ExternalDirectoryObjectId
            }
            elseif (-not $usersById.ContainsKey($keepId))
            {
                throw "Shared mailbox '$mailboxName' has keep_members account '$keepMember' that could not be resolved to a Microsoft 365 mailbox."
            }
            $keepUserIds[$keepId] = $true
        }
        $keepUserIdsByMailbox[$mailboxName] = $keepUserIds
    }

    foreach ($mailboxName in $mailboxSettings.Keys)
    {
        $setting = $mailboxSettings[$mailboxName]
        $alias = [string](Get-SNMicrosoft365Value -Object $setting -Name 'mail_nickname')
        $emailAddress = [string](Get-SNMicrosoft365Value -Object $setting -Name 'email_address')
        if ([string]::IsNullOrWhiteSpace($emailAddress))
        {
            $emailAddress = "$alias@$($snsConfiguration.DomainName)"
        }
        $manageFullAccess = (Get-SNMicrosoft365Value -Object $setting -Name 'full_access') -ne $false
        $manageSendAs = (Get-SNMicrosoft365Value -Object $setting -Name 'send_as') -ne $false
        $autoMapping = (Get-SNMicrosoft365Value -Object $setting -Name 'auto_mapping') -ne $false
        $hidden = Get-SNMicrosoft365Value -Object $setting -Name 'hidden_from_address_list'
        $copySentItems = Get-SNMicrosoft365Value -Object $setting -Name 'copy_sent_items'

        $recipients = @(Get-EXORecipient -Filter "Alias -eq '$alias'" -ErrorAction Stop)
        if ($recipients.Count -gt 1)
        {
            throw "More than one recipient has the alias '$alias' for shared mailbox '$mailboxName'."
        }
        $mailbox = $recipients | Select-Object -First 1
        # Never touch a recipient with the same alias that is not a shared mailbox.
        if ($null -ne $mailbox -and [string]$mailbox.RecipientTypeDetails -ne 'SharedMailbox')
        {
            throw "The alias '$alias' for shared mailbox '$mailboxName' belongs to a recipient of type '$($mailbox.RecipientTypeDetails)', not a shared mailbox. It was not changed."
        }

        $changed = $false
        $created = $false
        if ($null -eq $mailbox)
        {
            if (-not $PSCmdlet.ShouldProcess([string]$mailboxName, 'Create shared mailbox'))
            {
                continue
            }
            $mailbox = New-Mailbox -Shared -Name ([string]$mailboxName) -DisplayName ([string]$mailboxName) `
                -Alias $alias -PrimarySmtpAddress $emailAddress -ErrorAction Stop
            $changed = $true
            $created = $true
            Write-SNSLog -Level 'Info' "Created shared mailbox '$mailboxName' ($emailAddress)."
        }

        $identity = [string]$mailbox.ExternalDirectoryObjectId
        if ([string]::IsNullOrWhiteSpace($identity))
        {
            $identity = if ($created) { $emailAddress } else { [string]$mailbox.PrimarySmtpAddress }
        }
        # A newly created mailbox may not be visible everywhere yet.
        $maxAttempts = if ($created) { 12 } else { 1 }

        # Settings are only changed when configured.
        $mailboxParameters = @{}
        if ($null -ne $hidden -or $null -ne $copySentItems)
        {
            $current = if ($created) { $null } else {
                Get-EXOMailbox -Identity $identity `
                    -Properties HiddenFromAddressListsEnabled, MessageCopyForSentAsEnabled `
                    -Verbose:$false -ErrorAction Stop
            }
            if ($null -ne $hidden -and $current.HiddenFromAddressListsEnabled -ne $hidden)
            {
                $mailboxParameters.HiddenFromAddressListsEnabled = $hidden
            }
            if ($null -ne $copySentItems -and $current.MessageCopyForSentAsEnabled -ne $copySentItems)
            {
                $mailboxParameters.MessageCopyForSentAsEnabled = $copySentItems
            }
        }
        if ($mailboxParameters.Count -gt 0 -and
            $PSCmdlet.ShouldProcess([string]$mailboxName, 'Change shared mailbox settings'))
        {
            Invoke-SNExchangeRetry -MaxAttempts $maxAttempts -Action {
                Set-Mailbox -Identity $identity @mailboxParameters -ErrorAction Stop
            }
            $changed = $true
        }

        $desiredUserIds = @{}
        foreach ($memberId in $desiredMemberIdsByMailbox[$mailboxName].Keys)
        {
            if ($office365UsersByScoutnetId.ContainsKey([string]$memberId))
            {
                $userId = [string]$office365UsersByScoutnetId[[string]$memberId].ExternalDirectoryObjectId
                $desiredUserIds[$userId] = $true
            }
            else
            {
                Write-SNSLog -Level 'Info' "Scoutnet member '$memberId' matches shared mailbox '$mailboxName' but has no Microsoft 365 user account. The member is skipped."
            }
        }
        foreach ($userId in $keepUserIdsByMailbox[$mailboxName].Keys)
        {
            $desiredUserIds[[string]$userId] = $true
        }

        $protectedUserIds = @{}
        foreach ($userId in $keepUserIdsByMailbox[$mailboxName].Keys)
        {
            $protectedUserIds[[string]$userId] = $true
        }

        $permissionKinds = @()
        if ($manageFullAccess) { $permissionKinds += 'FullAccess' }
        if ($manageSendAs) { $permissionKinds += 'SendAs' }

        foreach ($kind in $permissionKinds)
        {
            $existingNames = if ($kind -eq 'FullAccess')
            {
                @(Invoke-SNExchangeRetry -MaxAttempts $maxAttempts -Action {
                    Get-MailboxPermission -Identity $identity -ErrorAction Stop
                } | Where-Object {
                    $_.IsInherited -ne $true -and $_.Deny -ne $true -and @($_.AccessRights) -contains 'FullAccess'
                } | ForEach-Object { [string]$_.User })
            }
            else
            {
                @(Invoke-SNExchangeRetry -MaxAttempts $maxAttempts -Action {
                    Get-RecipientPermission -Identity $identity -ErrorAction Stop
                } | Where-Object {
                    $_.IsInherited -ne $true -and [string]$_.AccessControlType -ne 'Deny' -and @($_.AccessRights) -contains 'SendAs'
                } | ForEach-Object { [string]$_.Trustee })
            }

            $plan = Get-SNSharedMailboxPermissionPlan -ExistingNames $existingNames `
                -DesiredUserIds $desiredUserIds -ProtectedUserIds $protectedUserIds `
                -KnownUsers $knownUsers -UsersById $usersById

            foreach ($user in $plan.Add)
            {
                $userPrincipalName = [string]$user.UserPrincipalName
                if (-not $PSCmdlet.ShouldProcess("$userPrincipalName on $mailboxName", "Add $kind permission"))
                {
                    continue
                }
                Invoke-SNExchangeRetry -MaxAttempts $maxAttempts -Action {
                    if ($kind -eq 'FullAccess')
                    {
                        Add-MailboxPermission -Identity $identity -User $userPrincipalName `
                            -AccessRights FullAccess -InheritanceType All -AutoMapping $autoMapping `
                            -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                    else
                    {
                        Add-RecipientPermission -Identity $identity -Trustee $userPrincipalName `
                            -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                }
                Write-SNSLog -Level 'Info' "Gave $userPrincipalName $kind permission on shared mailbox '$mailboxName'."
                $changed = $true
            }
            foreach ($user in $plan.Kept)
            {
                Write-SNSLog -Level 'Info' "$($user.UserPrincipalName) has $kind permission on shared mailbox '$mailboxName' but does not match the rule. The permission is kept because the account is listed in keep_members."
            }
            foreach ($user in $plan.Remove)
            {
                $userPrincipalName = [string]$user.UserPrincipalName
                if (-not $PSCmdlet.ShouldProcess("$userPrincipalName on $mailboxName", "Remove $kind permission"))
                {
                    continue
                }
                if ($kind -eq 'FullAccess')
                {
                    Remove-MailboxPermission -Identity $identity -User $userPrincipalName `
                        -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop
                }
                else
                {
                    Remove-RecipientPermission -Identity $identity -Trustee $userPrincipalName `
                        -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                }
                Write-SNSLog -Level 'Info' "Removed $kind permission for $userPrincipalName on shared mailbox '$mailboxName'."
                $changed = $true
            }
        }

        if (-not $changed)
        {
            Write-SNSLog -Level 'Info' "Shared mailbox '$mailboxName' is up to date. No changes needed."
        }
    }
}
