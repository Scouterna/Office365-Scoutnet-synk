function ConvertTo-SNSStringList
{
    param ($Value)

    @($Value | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ -ne '' })
}


function Set-SNSMailListSetting
{
    <#
    .SYNOPSIS
        Applies delegate, moderation, membership approval and delivery management settings to a distribution group.

    .DESCRIPTION
        Defaults that are always enforced: member_join_restriction and member_depart_restriction = Closed,
        moderation_notification = Never and allow_external_senders = $true. Other settings (send_as, require_moderation,
        moderators, bypass_moderation) are only applied when they are specified in the rule.
        Only values that differ are changed with a single Set-DistributionGroup call. Lists are always written when specified.

    .PARAMETER GroupName
        The key of the rule in MailListSettings.

    .PARAMETER Setting
        The rule from MailListSettings.

    .PARAMETER Created
        The group was just created, so Get-DistributionGroup is retried until the group is visible.
    #>
    [CmdletBinding(SupportsShouldProcess = $True)]
    param (
        [Parameter(Mandatory = $True)]
        [string]$GroupName,

        [Parameter()]
        $Setting,

        [switch]$Created
    )

    try
    {
        $group = $null
        $attempts = if ($Created) { 5 } else { 1 }
        for ($attempt = 1; $attempt -le $attempts -and $null -eq $group; $attempt++)
        {
            $group = Get-DistributionGroup -Identity $GroupName -ErrorAction "SilentlyContinue" -Verbose:$false
            if ($null -eq $group -and $attempt -lt $attempts)
            {
                Start-Sleep -Seconds 3
            }
        }
        if ($null -eq $group)
        {
            Write-SNSLog -Level "Warn" "Distribution group $GroupName was not found. Group settings are not applied."
            return
        }

        $changes = @{}

        $allowExternal = Get-SNMicrosoft365Value -Object $Setting -Name 'allow_external_senders'
        $allowExternal = if ($null -eq $allowExternal) { $true } else { [bool]$allowExternal }
        if ([bool]$group.RequireSenderAuthenticationEnabled -eq $allowExternal)
        {
            $changes.RequireSenderAuthenticationEnabled = -not $allowExternal
        }

        $restrictions = @(
            @{ Key = 'member_join_restriction'; Parameter = 'MemberJoinRestriction'; Allowed = @('Closed', 'Open', 'ApprovalRequired') }
            @{ Key = 'member_depart_restriction'; Parameter = 'MemberDepartRestriction'; Allowed = @('Closed', 'Open') }
            @{ Key = 'moderation_notification'; Parameter = 'SendModerationNotifications'; Allowed = @('Never', 'Internal', 'Always'); Default = 'Never' }
        )
        foreach ($restriction in $restrictions)
        {
            $default = if ($restriction.Default) { $restriction.Default } else { 'Closed' }
            $wanted = ([string](Get-SNMicrosoft365Value -Object $Setting -Name $restriction.Key)).Trim()
            if ($wanted -eq '') { $wanted = $default }
            $match = $restriction.Allowed | Where-Object { $_ -eq $wanted } | Select-Object -First 1
            if ($null -eq $match)
            {
                Write-SNSLog -Level "Warn" "Invalid $($restriction.Key) '$wanted' for $GroupName. Allowed values are $($restriction.Allowed -join ', '). The setting is not changed."
                continue
            }
            if ([string]$group.($restriction.Parameter) -ne $match)
            {
                $changes[$restriction.Parameter] = $match
            }
        }

        $requireModeration = Get-SNMicrosoft365Value -Object $Setting -Name 'require_moderation'
        if ($null -ne $requireModeration -and [bool]$group.ModerationEnabled -ne [bool]$requireModeration)
        {
            $changes.ModerationEnabled = [bool]$requireModeration
        }
        $listSettings = @{ moderators = 'ModeratedBy'; bypass_moderation = 'BypassModerationFromSendersOrMembers' }
        foreach ($key in $listSettings.Keys)
        {
            if ($Setting -is [System.Collections.IDictionary] -and $Setting.Contains($key))
            {
                $changes[$listSettings[$key]] = @(ConvertTo-SNSStringList -Value $Setting[$key])
            }
        }

        if ($changes.Count -gt 0 -and $PSCmdlet.ShouldProcess($GroupName, "Set-DistributionGroup"))
        {
            Write-SNSLog "Updating settings of $GroupName`: $(($changes.Keys | Sort-Object) -join ', ')"
            Set-DistributionGroup -Identity $GroupName @changes -ErrorAction "stop" -Verbose:$false -Confirm:$false
        }

        foreach ($trustee in (ConvertTo-SNSStringList -Value (Get-SNMicrosoft365Value -Object $Setting -Name 'send_as')))
        {
            if ($PSCmdlet.ShouldProcess($GroupName, "Add-RecipientPermission SendAs $trustee"))
            {
                try
                {
                    $permissionWarnings = @()
                    Add-RecipientPermission -Identity $GroupName -Trustee $trustee -AccessRights SendAs -Confirm:$false -ErrorAction "stop" -Verbose:$false -WarningAction "SilentlyContinue" -WarningVariable permissionWarnings | Out-Null
                    $alreadyPresentWarnings = @($permissionWarnings | Where-Object { "$_" -match 'appropriate access control entry is already present' })
                    if ($alreadyPresentWarnings.Count -gt 0)
                    {
                        Write-SNSLog "Send As permission for $trustee already exists on $GroupName"
                    }
                    else
                    {
                        Write-SNSLog "Send As permission request completed for $trustee on $GroupName"
                    }
                    foreach ($warning in @($permissionWarnings | Where-Object { "$_" -notmatch 'appropriate access control entry is already present' }))
                    {
                        Write-Warning "$warning"
                    }
                }
                catch
                {
                    Write-SNSLog -Level "Warn" "Could not give $trustee Send As permission on $GroupName. Error $_"
                }
            }
        }
    }
    catch
    {
        Write-SNSLog -Level "Error" "Update of settings for distribution group $GroupName failed! Error $_"
    }
}
