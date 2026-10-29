function Get-SNMicrosoft365Value
{
    <#
    .SYNOPSIS
        Reads a named value from a dictionary or object.

    .DESCRIPTION
        Returns the value for a key when Object is an IDictionary, otherwise the value of the
        property with that name. Scoutnet data and configuration can be either hashtables or
        PSCustomObjects, so this gives uniform, strict-mode safe access.

    .PARAMETER Object
        Dictionary or object to read from. May be null.

    .PARAMETER Name
        Key or property name to read.

    .OUTPUTS
        System.Object. The value, or $null when Object is null or does not contain the name.
    #>
    param (
        [Parameter()]
        $Object,

        [Parameter(Mandatory)]
        [string]$Name
    )

    if ($null -eq $Object)
    {
        return $null
    }

    if ($Object -is [System.Collections.IDictionary])
    {
        if ($Object.Contains($Name))
        {
            return $Object[$Name]
        }
        return $null
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -ne $property)
    {
        return $property.Value
    }
    return $null
}

function Get-SNMicrosoft365RoleName
{
    <#
    .SYNOPSIS
        Collects role names from nested Scoutnet role data.

    .DESCRIPTION
        Recursively walks dictionaries, collections and objects and writes the value of every
        'role_name' entry to the pipeline. Used for both roles.value.group and the entries
        below a single troop id in roles.value.troop.

    .PARAMETER RoleData
        Nested role data for a group or a single troop. May be null.

    .OUTPUTS
        System.String. One item per role name found.
    #>
    param (
        [Parameter()]
        $RoleData
    )

    if ($null -eq $RoleData)
    {
        return
    }

    $roleName = Get-SNMicrosoft365Value -Object $RoleData -Name 'role_name'
    if (-not [string]::IsNullOrWhiteSpace([string]$roleName))
    {
        Write-Output $roleName
        return
    }

    if ($RoleData -is [System.Collections.IDictionary])
    {
        foreach ($value in $RoleData.Values)
        {
            Get-SNMicrosoft365RoleName -RoleData $value
        }
    }
    elseif ($RoleData -is [System.Collections.IEnumerable] -and $RoleData -isnot [string])
    {
        foreach ($value in $RoleData)
        {
            Get-SNMicrosoft365RoleName -RoleData $value
        }
    }
    elseif ($RoleData -isnot [string] -and $RoleData -isnot [ValueType])
    {
        foreach ($property in $RoleData.PSObject.Properties)
        {
            Get-SNMicrosoft365RoleName -RoleData $property.Value
        }
    }
}

function Test-SNMicrosoft365ConfiguredRole
{
    <#
    .SYNOPSIS
        Tests whether a member has at least one of the configured roles.

    .PARAMETER ActualRoles
        Role names the member has.

    .PARAMETER ConfiguredRoles
        Role names from the group rule. The comparison is case-insensitive. The wildcard '*'
        matches any role, so the member matches when it has at least one role.

    .OUTPUTS
        System.Boolean. $true when any configured role is among the actual roles, or when '*'
        is configured and the member has any role.
    #>
    param (
        [Parameter()]
        [string[]]$ActualRoles = @(),

        [Parameter(Mandatory)]
        [string[]]$ConfiguredRoles
    )

    foreach ($role in $ConfiguredRoles)
    {
        if ($role -eq '*' -and @($ActualRoles | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -gt 0)
        {
            return $true
        }
        if ($ActualRoles -contains $role)
        {
            return $true
        }
    }
    return $false
}

function Test-SNMicrosoft365UnitMember
{
    <#
    .SYNOPSIS
        Tests whether a member belongs to a troop selected with unit_members in a troop rule.

    .DESCRIPTION
        Membership is read from the member's 'unit' object. A troop rule with unit_members = $true
        selects the members whose unit has the configured name (unit.value), or else the
        configured id (unit.raw_value).

    .PARAMETER Member
        Member record from Get-SNSApiGroupMemberlist.

    .PARAMETER Setting
        The rule with a 'troops' list.

    .OUTPUTS
        System.Boolean.
    #>
    param (
        [Parameter(Mandatory)]
        $Member,

        [Parameter(Mandatory)]
        $Setting
    )

    $unit = Get-SNMicrosoft365Value -Object $Member -Name 'unit'
    $unitId = [string](Get-SNMicrosoft365Value -Object $unit -Name 'raw_value')
    $unitName = [string](Get-SNMicrosoft365Value -Object $unit -Name 'value')

    foreach ($troopSetting in @(Get-SNMicrosoft365Value -Object $Setting -Name 'troops'))
    {
        if ((Get-SNMicrosoft365Value -Object $troopSetting -Name 'unit_members') -ne $true)
        {
            continue
        }
        $configuredId = [string](Get-SNMicrosoft365Value -Object $troopSetting -Name 'id')
        $configuredName = [string](Get-SNMicrosoft365Value -Object $troopSetting -Name 'name')
        if (-not [string]::IsNullOrWhiteSpace($configuredName) -and $unitName -eq $configuredName)
        {
            return $true
        }
        if (-not [string]::IsNullOrWhiteSpace($configuredId) -and $unitId -eq $configuredId)
        {
            return $true
        }
    }
    return $false
}

function Test-SNMicrosoft365RoleGroupMatch
{
    <#
    .SYNOPSIS
        Tests whether a Scoutnet member matches a role based group rule.

    .DESCRIPTION
        For source 'group_roles' the member matches when roles.value.group contains any of the
        configured roles. For source 'troop_roles' each rule in 'troops' is checked against the
        member's roles.value.troop entries. A troop is selected by its name, resolved through
        UnitNamesById, and if the name does not match by its configured id. The member matches
        when any selected troop holds any of that rule's roles.

    .PARAMETER Member
        Member record from Get-SNSApiGroupMemberlist.

    .PARAMETER Source
        Rule source, either 'group_roles' or 'troop_roles'.

    .PARAMETER Setting
        The group rule from Microsoft365GroupSettings.

    .PARAMETER UnitNamesById
        Troop names keyed by troop id, built from the 'unit' objects (raw_value and value).

    .PARAMETER RolesOnly
        Ignore unit_members and only test roles.

    .OUTPUTS
        System.Boolean.
    #>
    param (
        [Parameter(Mandatory)]
        $Member,

        [Parameter(Mandatory)]
        [string]$Source,

        [Parameter(Mandatory)]
        $Setting,

        [Parameter(Mandatory)]
        [hashtable]$UnitNamesById,

        [Parameter()]
        [switch]$RolesOnly
    )

    $roleData = Get-SNMicrosoft365Value -Object $Member.roles -Name 'value'
    if ($Source -eq 'group_roles')
    {
        $groupRoles = Get-SNMicrosoft365Value -Object $roleData -Name 'group'
        $actualRoles = @(Get-SNMicrosoft365RoleName -RoleData $groupRoles)
        $configuredRoles = @((Get-SNMicrosoft365Value -Object $Setting -Name 'roles') | ForEach-Object { [string]$_ })
        return Test-SNMicrosoft365ConfiguredRole -ActualRoles $actualRoles -ConfiguredRoles $configuredRoles
    }

    if (-not $RolesOnly -and (Test-SNMicrosoft365UnitMember -Member $Member -Setting $Setting))
    {
        return $true
    }

    $troopRoles = Get-SNMicrosoft365Value -Object $roleData -Name 'troop'
    if ($troopRoles -isnot [System.Collections.IDictionary])
    {
        return $false
    }

    foreach ($troopSetting in @(Get-SNMicrosoft365Value -Object $Setting -Name 'troops'))
    {
        $configuredId = [string](Get-SNMicrosoft365Value -Object $troopSetting -Name 'id')
        $configuredName = [string](Get-SNMicrosoft365Value -Object $troopSetting -Name 'name')
        $configuredRoles = @((Get-SNMicrosoft365Value -Object $troopSetting -Name 'roles') | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_ })
        if ($configuredRoles.Count -eq 0)
        {
            continue
        }

        foreach ($troopEntry in $troopRoles.GetEnumerator())
        {
            $troopId = [string]$troopEntry.Key
            $troopMatches = -not [string]::IsNullOrWhiteSpace($configuredName) -and
                $UnitNamesById[$troopId] -eq $configuredName
            if (-not $troopMatches -and -not [string]::IsNullOrWhiteSpace($configuredId))
            {
                $troopMatches = $troopId -eq $configuredId
            }
            if (-not $troopMatches)
            {
                continue
            }

            $actualRoles = @(Get-SNMicrosoft365RoleName -RoleData $troopEntry.Value)
            if (Test-SNMicrosoft365ConfiguredRole -ActualRoles $actualRoles -ConfiguredRoles $configuredRoles)
            {
                return $true
            }
        }
    }
    return $false
}

function Update-SNMicrosoft365Group
{
    <#
    .SYNOPSIS
        Synchronizes configured Microsoft 365 group memberships with Scoutnet.

    .DESCRIPTION
        Only groups explicitly configured in Microsoft365GroupSettings are managed. Other
        groups are never read for changes or modified.

        Microsoft365GroupSettings is a hashtable where each key is the group display name and
        each value is a rule with these entries:

        Required for all rules
          source         'mail_list', 'troop_roles' or 'group_roles'.
          mail_nickname  Unique mail nickname, used when the group is created.
        Optional for all rules
          description    Group description, used when the group is created.
          create_team    Set to $true to also make the group a Microsoft Teams team.
          keep_members   Accounts (user principal names or object ids) that are never removed
                         from the group even if they do not match the rule.
          team_owners    Owners (user principal names or object ids) added to the group when
                         the team is created. Teams requires the group to have an owner.

        source = 'mail_list' (membership from a Scoutnet maillist)
          scoutnet_list_id  Required. Id of the list in Scoutnet.
          member_type       Optional. 'ledare' (default), 'scouter' or 'both'.
          statisk_lista     Optional. Set for static Scoutnet lists.

        source = 'group_roles' (roles on the scout group)
          roles  Required. Role names, for example 'Kårledare'. One match is enough.
                 Use '*' to match every member with any role on the scout group.

        source = 'troop_roles' (roles in one or more troops)
          troops  Required. A list of rules, each with 'roles' and 'name' and/or 'id'. Add
                  unit_members = $true to also include the members of that troop (the member's
                  'unit'); 'roles' is then optional.
                  The troop name is matched first and the troop id is the fallback.
                  Use roles = '*' to match every member with any role in that troop.

        The function validates all rules first, fetches the member list, the maillists and the
        Exchange mailboxes, and only then changes Microsoft 365. A member is mapped to an
        account through the mailbox CustomAttribute1, which holds the Scoutnet member number.
        A member matching a rule without an account, for example a leader too young for an
        own account, is written to the log and skipped.

        For each rule the group is created as a Microsoft 365 (Unified) group if it does not
        exist. Matching accounts missing in the group are added and members that no longer
        match, for example leaders whose account was closed or who left the troop, are removed.
        Owners of the group and accounts listed in keep_members are never removed, which also
        protects an existing team whose owners are not part of the rule.
        Existing groups that are not Microsoft 365 groups, or display names that are not
        unique, cause an error and the group is not changed.

        When create_team is $true the group is made a Teams team after its membership has been
        synchronized, unless it already is a team. A team is never removed or changed if
        create_team is omitted. Creating teams needs the Microsoft.Graph.Teams module (New-MgTeam)
        and the Team.Create permission. If the group is new, creating the team is retried a few
        times because Teams can take a while to see a new group.

    .INPUTS
        None. You cannot pipe objects to Update-SNMicrosoft365Group.

    .OUTPUTS
        None.

    .PARAMETER Configuration
        Configuration to use. If omitted, the cached module configuration is used. A supplied
        configuration also becomes the cached configuration.

    .PARAMETER WhatIf
        Shows the group creations and membership changes without applying them.

    .EXAMPLE
        $conf.Microsoft365GroupSettings = @{
            'Styrelsen' = @{
                source        = 'group_roles'
                mail_nickname = 'styrelsen'
                roles         = @('Ordförande', 'Kassör', 'Utmanarscoutrepresentant')
                create_team   = $true
                team_owners   = @('admin@example.org')
            }
        }
        Update-SNMicrosoft365Group -Configuration $conf

    .EXAMPLE
        $conf.Microsoft365GroupSettings = @{
            'Avdelningsledare Äventyrarna' = @{
                source        = 'troop_roles'
                mail_nickname = 'avdelningsledare-aventyrarna'
                troops        = @(@{ name = 'Äventyrarna'; id = '2'; roles = @('Avdelningsledare', 'Ledare') })
            }
        }
        Update-SNMicrosoft365Group -Configuration $conf -WhatIf

    .EXAMPLE
        $conf.Microsoft365GroupSettings = @{
            'Alla ledare Äventyrarna' = @{
                source        = 'troop_roles'
                mail_nickname = 'alla-ledare-aventyrarna'
                troops        = @(@{ name = 'Äventyrarna'; id = '2'; roles = '*' })
            }
        }
        Update-SNMicrosoft365Group -Configuration $conf
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

    $groupSettings = $snsConfiguration.Microsoft365GroupSettings
    if ($null -eq $groupSettings -or $groupSettings.Count -eq 0)
    {
        Write-SNSLog 'No Microsoft 365 group rules are configured. No groups will be changed.'
        return
    }

    # Validate every rule before anything is fetched or changed.
    foreach ($groupName in $groupSettings.Keys)
    {
        $setting = $groupSettings[$groupName]
        $source = ([string](Get-SNMicrosoft365Value -Object $setting -Name 'source')).ToLowerInvariant()
        $mailNickname =  [string](Get-SNMicrosoft365Value -Object $setting -Name 'mail_nickname')

        if ([string]::IsNullOrWhiteSpace([string]$groupName))
        {
            throw 'Microsoft 365 group configuration keys must contain a display name.'
        }
        if ([string]::IsNullOrWhiteSpace($mailNickname))
        {
            throw "Microsoft 365 group '$groupName' must configure mail_nickname."
        }
        if ($source -notin @('mail_list', 'troop_roles', 'group_roles'))
        {
            throw "Microsoft 365 group '$groupName' has unsupported source '$source'. Use mail_list, troop_roles, or group_roles."
        }
        if ((Get-SNMicrosoft365Value -Object $setting -Name 'create_team') -eq $true)
        {
            if (-not (Get-Command -Name New-MgTeam -ErrorAction SilentlyContinue))
            {
                throw "Microsoft 365 group '$groupName' has create_team set but New-MgTeam is not available. Install the Microsoft.Graph.Teams module."
            }
            if (@(Get-SNMicrosoft365Value -Object $setting -Name 'team_owners').Where({ -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -eq 0)
            {
                throw "Microsoft 365 group '$groupName' has create_team set and must configure team_owners, since a team needs an owner."
            }
        }

        if ($source -eq 'mail_list' -and
            [string]::IsNullOrWhiteSpace([string](Get-SNMicrosoft365Value -Object $setting -Name 'scoutnet_list_id')))
        {
            throw "Microsoft 365 group '$groupName' must configure scoutnet_list_id for source mail_list."
        }
        if ($source -eq 'mail_list')
        {
            $memberType = [string](Get-SNMicrosoft365Value -Object $setting -Name 'member_type')
            if ([string]::IsNullOrWhiteSpace($memberType))
            {
                $memberType = 'ledare'
            }
            if ($memberType.ToLowerInvariant() -notin @('ledare', 'scouter', 'both'))
            {
                throw "Microsoft 365 group '$groupName' has unsupported member_type '$memberType'. Use ledare, scouter, or both."
            }
        }
        elseif ($source -eq 'group_roles')
        {
            $configuredRoles = @(Get-SNMicrosoft365Value -Object $setting -Name 'roles')
            if ($configuredRoles.Count -eq 0 -or
                @($configuredRoles | Where-Object { [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0)
            {
                throw "Microsoft 365 group '$groupName' must configure one or more non-empty roles for source group_roles."
            }
        }
        else
        {
            $troopSettings = @(Get-SNMicrosoft365Value -Object $setting -Name 'troops')
            if ($troopSettings.Count -eq 0)
            {
                throw "Microsoft 365 group '$groupName' must configure troops for source troop_roles."
            }
            foreach ($troopSetting in $troopSettings)
            {
                $troopId = [string](Get-SNMicrosoft365Value -Object $troopSetting -Name 'id')
                $troopName = [string](Get-SNMicrosoft365Value -Object $troopSetting -Name 'name')
                $roles = @(Get-SNMicrosoft365Value -Object $troopSetting -Name 'roles').Where({ $null -ne $_ })
                if ([string]::IsNullOrWhiteSpace($troopId) -and [string]::IsNullOrWhiteSpace($troopName))
                {
                    throw "Each troop rule for Microsoft 365 group '$groupName' must configure name or id."
                }
                $unitMembers = (Get-SNMicrosoft365Value -Object $troopSetting -Name 'unit_members') -eq $true
                if (($roles.Count -eq 0 -and -not $unitMembers) -or
                    @($roles | Where-Object { [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0)
                {
                    throw "Each troop rule for Microsoft 365 group '$groupName' must configure one or more non-empty roles, or unit_members = `$true."
                }
            }
        }
    }

    $memberlist = Get-SNSApiGroupMemberlist -Credential $snsConfiguration.CredentialMemberlist
    if ($null -eq $memberlist -or $memberlist.data -isnot [System.Collections.IDictionary])
    {
        throw "The member list does not contain the expected 'data' dictionary. No Microsoft 365 groups were changed."
    }

    # Troop names are only available on each member's 'unit' object, so build an id to name map.
    $unitNamesById = @{}
    foreach ($member in $memberlist.data.Values)
    {
        $unit = Get-SNMicrosoft365Value -Object $member -Name 'unit'
        $unitId = [string](Get-SNMicrosoft365Value -Object $unit -Name 'raw_value')
        $unitName = [string](Get-SNMicrosoft365Value -Object $unit -Name 'value')
        if (-not [string]::IsNullOrWhiteSpace($unitId) -and -not [string]::IsNullOrWhiteSpace($unitName))
        {
            $unitNamesById[$unitId] = $unitName
        }
    }

    # Only fetch the Scoutnet maillists that the rules refer to.
    $mailListSettings = @()
    foreach ($groupName in $groupSettings.Keys)
    {
        $setting = $groupSettings[$groupName]
        if (([string](Get-SNMicrosoft365Value -Object $setting -Name 'source')).ToLowerInvariant() -eq 'mail_list')
        {
            $listSetting = @{
                scoutnet_list_id = [string](Get-SNMicrosoft365Value -Object $setting -Name 'scoutnet_list_id')
            }
            $staticList = Get-SNMicrosoft365Value -Object $setting -Name 'statisk_lista'
            if ($null -ne $staticList)
            {
                $listSetting.statisk_lista = $staticList
            }
            $mailListSettings += $listSetting
        }
    }

    $customLists = @{}
    if ($mailListSettings.Count -gt 0)
    {
        $customLists, $customListsHash = Get-SNSMaillistInfo `
            -CredentialCustomlists $snsConfiguration.CredentialCustomlists `
            -MaillistsIds $mailListSettings
        foreach ($listSetting in $mailListSettings)
        {
            if (-not $customLists.ContainsKey([string]$listSetting.scoutnet_list_id))
            {
                throw "Scoutnet mailing list '$($listSetting.scoutnet_list_id)' was not found. No Microsoft 365 groups were changed."
            }
        }
    }

    # CustomAttribute1 holds the Scoutnet member number of each Microsoft 365 account.
    $allOffice365Users = @(
        Get-EXOMailbox -RecipientTypeDetails 'UserMailbox' -Properties CustomAttribute1 -Verbose:$false -ErrorAction Stop
    )
    $office365UsersByScoutnetId = @{}
    foreach ($office365User in $allOffice365Users)
    {
        $scoutnetId = [string]$office365User.CustomAttribute1
        if (-not [string]::IsNullOrWhiteSpace($scoutnetId) -and
            -not [string]::IsNullOrWhiteSpace([string]$office365User.ExternalDirectoryObjectId))
        {
            $office365UsersByScoutnetId[$scoutnetId] = $office365User
        }
    }

    # Work out the wanted Scoutnet member numbers per group before touching Microsoft 365.
    $desiredMemberIdsByGroup = @{}
    foreach ($groupName in $groupSettings.Keys)
    {
        $setting = $groupSettings[$groupName]
        $source = ([string](Get-SNMicrosoft365Value -Object $setting -Name 'source')).ToLowerInvariant()
        $desiredMemberIds = @{}

        if ($source -eq 'mail_list')
        {
            $listId = [string](Get-SNMicrosoft365Value -Object $setting -Name 'scoutnet_list_id')
            $memberType = [string](Get-SNMicrosoft365Value -Object $setting -Name 'member_type')
            if ([string]::IsNullOrWhiteSpace($memberType))
            {
                $memberType = 'ledare'
            }
            $memberType = $memberType.ToLowerInvariant()
            if ($memberType -in @('ledare', 'both'))
            {
                foreach ($memberId in @($customLists[$listId].ledare))
                {
                    if (-not [string]::IsNullOrWhiteSpace([string]$memberId))
                    {
                        $desiredMemberIds[[string]$memberId] = $true
                    }
                }
            }
            if ($memberType -in @('scouter', 'both'))
            {
                foreach ($memberId in @($customLists[$listId].scouter))
                {
                    if (-not [string]::IsNullOrWhiteSpace([string]$memberId))
                    {
                        $desiredMemberIds[[string]$memberId] = $true
                    }
                }
            }
        }
        else
        {
            foreach ($memberEntry in $memberlist.data.GetEnumerator())
            {
                $member = $memberEntry.Value
                if (Test-SNMicrosoft365RoleGroupMatch -Member $member -Source $source `
                    -Setting $setting -UnitNamesById $unitNamesById)
                {
                    $memberId = [string](Get-SNMicrosoft365Value -Object $member.member_no -Name 'value')
                    if ([string]::IsNullOrWhiteSpace($memberId))
                    {
                        $memberId = [string]$memberEntry.Key
                    }
                    if (-not [string]::IsNullOrWhiteSpace($memberId))
                    {
                        $desiredMemberIds[$memberId] = $true
                    }
                }
            }
        }

        $desiredMemberIdsByGroup[$groupName] = $desiredMemberIds
    }

    # Find or create each configured group and make its membership match the rule.
    foreach ($groupName in $groupSettings.Keys)
    {
        $setting = $groupSettings[$groupName]
        $escapedName = ([string]$groupName).Replace("'", "''")
        $groups = @(
            Get-MgGroup -Filter "displayName eq '$escapedName'" `
                -Property Id, DisplayName, GroupTypes, MailEnabled, SecurityEnabled, ResourceProvisioningOptions `
                -All -ErrorAction Stop
        )
        $matchingGroups = @($groups | Where-Object { $_.DisplayName -ceq [string]$groupName })
        if ($matchingGroups.Count -gt 1)
        {
            throw "More than one Microsoft 365 group has display name '$groupName'."
        }

        $group = $matchingGroups | Select-Object -First 1
        # Never touch a group with the same name that is not a Microsoft 365 group.
        if ($null -ne $group -and
            ($group.MailEnabled -ne $true -or @($group.GroupTypes) -notcontains 'Unified'))
        {
            throw "The configured display name '$groupName' belongs to a non-Microsoft 365 group. It was not changed."
        }

        $groupChanged = $false
        if ($null -eq $group)
        {
            $description = [string](Get-SNMicrosoft365Value -Object $setting -Name 'description')
            $groupParameters = @{
                Description = $description
                DisplayName = [string]$groupName
                GroupTypes = @('Unified')
                MailEnabled = $true
                MailNickname = [string](Get-SNMicrosoft365Value -Object $setting -Name 'mail_nickname')
                SecurityEnabled = $false
            }
            if (-not $PSCmdlet.ShouldProcess([string]$groupName, 'Create Microsoft 365 group'))
            {
                continue
            }
            $group = New-MgGroup -BodyParameter $groupParameters -ErrorAction Stop
            $groupChanged = $true
            if ($null -eq $group -or [string]::IsNullOrWhiteSpace([string]$group.Id))
            {
                throw "Microsoft Graph did not return an ID after creating group '$groupName'."
            }
        }

        $desiredDirectoryObjectIds = @{}
        foreach ($memberId in $desiredMemberIdsByGroup[$groupName].Keys)
        {
            if ($office365UsersByScoutnetId.ContainsKey([string]$memberId))
            {
                $userId = [string]$office365UsersByScoutnetId[[string]$memberId].ExternalDirectoryObjectId
                $desiredDirectoryObjectIds[$userId] = $true
            }
            else
            {
                Write-SNSLog -Level 'Info' "Scoutnet member '$memberId' matches Microsoft 365 group '$groupName' but has no Microsoft 365 user account. The member is skipped."
            }
        }

        $existingGroupMembers = @(
            Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop
        )
        $existingDirectoryObjectIds = @{}
        foreach ($groupMember in $existingGroupMembers)
        {
            if (-not [string]::IsNullOrWhiteSpace([string]$groupMember.Id))
            {
                $existingDirectoryObjectIds[[string]$groupMember.Id] = $true
            }
        }

        foreach ($userId in $desiredDirectoryObjectIds.Keys)
        {
            if (-not $existingDirectoryObjectIds.ContainsKey([string]$userId) -and
                $PSCmdlet.ShouldProcess("$userId in $groupName", 'Add member to Microsoft 365 group'))
            {
                New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $userId -ErrorAction Stop
                $groupChanged = $true
            }
        }
        $staleDirectoryObjectIds = @(
            $existingDirectoryObjectIds.Keys |
                Where-Object { -not $desiredDirectoryObjectIds.ContainsKey([string]$_) }
        )
        if ($staleDirectoryObjectIds.Count -gt 0)
        {
            # Group owners and the accounts listed in keep_members are never removed.
            $protectedObjectIds = @{}
            foreach ($owner in @(Get-MgGroupOwner -GroupId $group.Id -All -ErrorAction Stop))
            {
                $protectedObjectIds[[string]$owner.Id] = $true
            }
            foreach ($keepMember in @(Get-SNMicrosoft365Value -Object $setting -Name 'keep_members'))
            {
                if ([string]::IsNullOrWhiteSpace([string]$keepMember)) { continue }
                $keepId = [string]$keepMember
                if ($keepId.Contains('@'))
                {
                    $keepId = [string](Get-MgUser -UserId $keepId -Property Id -ErrorAction Stop).Id
                }
                $protectedObjectIds[$keepId] = $true
            }
        }
        foreach ($userId in $staleDirectoryObjectIds)
        {
            if ($protectedObjectIds.ContainsKey([string]$userId))
            {
                Write-SNSLog -Level 'Info' "Member '$userId' of Microsoft 365 group '$groupName' does not match the rule but is an owner or listed in keep_members. The member is kept."
                continue
            }
            if ($PSCmdlet.ShouldProcess("$userId from $groupName", 'Remove member from Microsoft 365 group'))
            {
                Remove-MgGroupMemberDirectoryObjectByRef -GroupId $group.Id -DirectoryObjectId $userId -ErrorAction Stop
                $groupChanged = $true
            }
        }

        $teamWanted = (Get-SNMicrosoft365Value -Object $setting -Name 'create_team') -eq $true
        $teamExists = @($group.ResourceProvisioningOptions) -contains 'Team'

        # Optionally make the group a Teams team. An existing team is left as it is.
        if ($teamWanted -and -not $teamExists -and
            $PSCmdlet.ShouldProcess([string]$groupName, 'Create Microsoft Teams team'))
        {
            $existingOwnerIds = @(Get-MgGroupOwner -GroupId $group.Id -All -ErrorAction Stop).Id
            foreach ($teamOwner in @(Get-SNMicrosoft365Value -Object $setting -Name 'team_owners'))
            {
                if ([string]::IsNullOrWhiteSpace([string]$teamOwner)) { continue }
                $ownerId = [string](Get-MgUser -UserId ([string]$teamOwner) -Property Id -ErrorAction Stop).Id
                if ($existingOwnerIds -notcontains $ownerId)
                {
                    New-MgGroupOwnerByRef -GroupId $group.Id `
                        -OdataId "https://graph.microsoft.com/v1.0/users/$ownerId" -ErrorAction Stop
                }
            }

            $teamParameters = @{
                'template@odata.bind' = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
                'group@odata.bind' = "https://graph.microsoft.com/v1.0/groups('$($group.Id)')"
            }
            $maxAttempts = 5
            for ($attempt = 1; $attempt -le $maxAttempts; $attempt++)
            {
                try
                {
                    New-MgTeam -BodyParameter $teamParameters -ErrorAction Stop | Out-Null
                    Write-SNSLog -Level 'Info' "Created Microsoft Teams team for group '$groupName'."
                    $groupChanged = $true
                    break
                }
                catch
                {
                    if ($attempt -eq $maxAttempts) { throw }
                    Start-Sleep -Seconds 15
                }
            }
        }

        if (-not $groupChanged)
        {
            $teamText = if ($teamWanted -and $teamExists) { ' The Microsoft Teams team already exists.' } else { '' }
            Write-SNSLog -Level 'Info' "Microsoft 365 group '$groupName' is up to date. No changes needed.$teamText"
        }
    }
}
