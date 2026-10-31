function Get-SNSMailListGroupType
{
    <#
    .SYNOPSIS
        Returns the group_type of a distribution group rule in MailListSettings.

    .DESCRIPTION
        Without group_type the group must already exist in Exchange, which is the original behavior.
        With group_type = 'security' a missing group is created as a mail-enabled security group, and with
        'distribution' as a regular distribution list.

    .PARAMETER Setting
        The rule from MailListSettings.

    .OUTPUTS
        System.String. Empty, 'security' or 'distribution' (lower case). Throws for other values.
    #>
    [OutputType([string])]
    param (
        [Parameter()]
        $Setting
    )

    $groupType = ([string](Get-SNMicrosoft365Value -Object $Setting -Name 'group_type')).Trim().ToLowerInvariant()
    if ($groupType -notin @('', 'security', 'distribution'))
    {
        throw "Unsupported group_type '$groupType' for a distribution group rule. Use security, distribution or omit group_type."
    }
    return $groupType
}

function Get-SNSMailListSource
{
    <#
    .SYNOPSIS
        Returns the source of a distribution group rule in MailListSettings.

    .DESCRIPTION
        A rule without 'source' (or with source 'mail_list') is driven by a Scoutnet maillist, which is
        the original behavior. Sources 'troop_roles' and 'group_roles' are driven by the roles
        of the leaders in Scoutnet.

    .PARAMETER Setting
        The rule from MailListSettings.

    .OUTPUTS
        System.String. One of mail_list, troop_roles or group_roles (lower case).
    #>
    [OutputType([string])]
    param (
        [Parameter()]
        $Setting
    )

    $source = ([string](Get-SNMicrosoft365Value -Object $Setting -Name 'source')).Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($source))
    {
        return 'mail_list'
    }
    return $source
}

function Test-SNSRoleListSetting
{
    <#
    .SYNOPSIS
        Validates a role based distribution group rule. Throws when the rule is not valid.

    .PARAMETER Name
        Name of the distribution group.

    .PARAMETER Setting
        The rule from MailListSettings.
    #>
    param (
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter()]
        $Setting
    )

    $source = Get-SNSMailListSource -Setting $Setting
    if ($source -eq 'mail_list')
    {
        return
    }
    if ($source -notin @('troop_roles', 'group_roles'))
    {
        throw "Distribution group '$Name' has unsupported source '$source'. Use mail_list, troop_roles, or group_roles."
    }

    if ($source -eq 'group_roles')
    {
        $roles = @(Get-SNMicrosoft365Value -Object $Setting -Name 'roles')
        if ($roles.Count -eq 0 -or @($roles | Where-Object { [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0)
        {
            throw "Distribution group '$Name' must configure one or more non-empty roles for source group_roles."
        }
        return
    }

    $troopSettings = @(Get-SNMicrosoft365Value -Object $Setting -Name 'troops')
    if ($troopSettings.Count -eq 0)
    {
        throw "Distribution group '$Name' must configure troops for source troop_roles."
    }
    foreach ($troopSetting in $troopSettings)
    {
        $troopId = [string](Get-SNMicrosoft365Value -Object $troopSetting -Name 'id')
        $troopName = [string](Get-SNMicrosoft365Value -Object $troopSetting -Name 'name')
        $roles = @(Get-SNMicrosoft365Value -Object $troopSetting -Name 'roles').Where({ $null -ne $_ })
        if ([string]::IsNullOrWhiteSpace($troopId) -and [string]::IsNullOrWhiteSpace($troopName))
        {
            throw "Each troop rule for distribution group '$Name' must configure name or id."
        }
        $unitMembers = (Get-SNMicrosoft365Value -Object $troopSetting -Name 'unit_members') -eq $true
        if (($roles.Count -eq 0 -and -not $unitMembers) -or
            @($roles | Where-Object { [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0)
        {
            throw "Each troop rule for distribution group '$Name' must configure one or more non-empty roles, or unit_members = `$true."
        }
    }
}

function Get-SNSRoleListMember
{
    <#
    .SYNOPSIS
        Finds the Scoutnet members that match the role based distribution group rules.

    .DESCRIPTION
        Rules with source 'troop_roles' or 'group_roles' are matched against the roles of every member
        in the Scoutnet member list, in the same way as for Microsoft 365 groups. Rules without
        source are ignored here.

    .PARAMETER MailListSettings
        The MailListSettings hashtable, keyed by distribution group name.

    .PARAMETER MemberList
        The member list returned by Get-SNSApiGroupMemberlist.

    .OUTPUTS
        System.Collections.Hashtable. Keyed by distribution group name, for role based rules only. Each
        value has sorted Scoutnet member numbers in 'ledare' (matched by role) and 'scouter' (members of
        a unit selected with unit_members who are not matched by role).
    #>
    [OutputType([System.Collections.Hashtable])]
    param (
        [Parameter(Mandatory)]
        [hashtable]$MailListSettings,

        [Parameter(Mandatory)]
        $MemberList
    )

    $result = @{}
    if ($null -eq $MemberList -or $MemberList.data -isnot [System.Collections.IDictionary])
    {
        throw "The member list does not contain the expected 'data' dictionary."
    }

    $unitNamesById = @{}
    foreach ($member in $MemberList.data.Values)
    {
        $unit = Get-SNMicrosoft365Value -Object $member -Name 'unit'
        $unitId = [string](Get-SNMicrosoft365Value -Object $unit -Name 'raw_value')
        $unitName = [string](Get-SNMicrosoft365Value -Object $unit -Name 'value')
        if (-not [string]::IsNullOrWhiteSpace($unitId) -and -not [string]::IsNullOrWhiteSpace($unitName))
        {
            $unitNamesById[$unitId] = $unitName
        }
    }

    foreach ($listName in $MailListSettings.Keys)
    {
        $setting = $MailListSettings[$listName]
        $source = Get-SNSMailListSource -Setting $setting
        if ($source -eq 'mail_list')
        {
            continue
        }

        $leaderIds = [System.Collections.Generic.List[object]]::new()
        $scoutIds = [System.Collections.Generic.List[object]]::new()
        foreach ($memberEntry in $MemberList.data.GetEnumerator())
        {
            $memberId = Get-SNMicrosoft365Value -Object $memberEntry.Value.member_no -Name 'value'
            if ([string]::IsNullOrWhiteSpace([string]$memberId))
            {
                $memberId = $memberEntry.Key
            }
            if (Test-SNMicrosoft365RoleGroupMatch -Member $memberEntry.Value -Source $source `
                -Setting $setting -UnitNamesById $unitNamesById -RolesOnly)
            {
                $leaderIds.Add($memberId)
            }
            elseif (Test-SNMicrosoft365UnitMember -Member $memberEntry.Value -Setting $setting)
            {
                $scoutIds.Add($memberId)
            }
        }
        $result[$listName] = @{
            ledare = @($leaderIds | Sort-Object)
            scouter = @($scoutIds | Sort-Object)
        }
    }
    return $result
}
