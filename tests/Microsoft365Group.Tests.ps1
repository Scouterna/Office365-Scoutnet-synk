BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Office365-Scoutnet-synk\Office365-Scoutnet-synk.psm1'
    Import-Module $modulePath -Force -ErrorAction Stop
}

AfterAll {
    Remove-Module Office365-Scoutnet-synk -ErrorAction SilentlyContinue
}

Describe 'Update-SNMicrosoft365Group' {
    It 'accepts Microsoft 365 group rules when creating configuration' {
        InModuleScope Office365-Scoutnet-synk {
            $groupSettings = @{
                'Kårledning' = @{
                    source = 'group_roles'
                    mail_nickname = 'karledning'
                    roles = @('Kårledare')
                }
            }

            $configuration = New-SNSConfiguration `
                -LicenseAssignment @{} `
                -Microsoft365GroupSettings $groupSettings

            $configuration.Microsoft365GroupSettings['Kårledning'].mail_nickname |
                Should -Be 'karledning'
        }
    }

    It 'creates configured groups and synchronizes memberships from troop roles using the troop ID fallback' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXOMailbox {}
            function Get-MgGroup {
                param($Filter, $Property, $All, $ErrorAction)
            }
            function Get-MgGroupMember {
                param($GroupId, $All, $ErrorAction)
            }
            function New-MgGroup {
                param($BodyParameter, $ErrorAction)
            }
            function New-MgGroupMember {
                param($GroupId, $DirectoryObjectId, $ErrorAction)
            }
            function Remove-MgGroupMemberDirectoryObjectByRef {
                param($GroupId, $DirectoryObjectId, $ErrorAction)
            }

            $configuration = [SNSConfiguration]::new()
            $configuration.CredentialMemberlist = [pscredential]::new(
                'test',
                (ConvertTo-SecureString 'test' -AsPlainText -Force)
            )
            $configuration.Microsoft365GroupSettings = @{
                'Avdelningsledare Äventyrarna' = @{
                    source = 'troop_roles'
                    mail_nickname = 'avdelningsledare-aventyrarna'
                    description = 'Avdelningens ledare'
                    troops = @(
                        @{
                            name = 'Gammalt avdelningsnamn'
                            id = '2'
                            roles = @('Ledare')
                        }
                    )
                }
            }
            $memberlist = @{
                data = @{
                    member1 = @{
                        member_no = @{ value = 'member-1' }
                        unit = @{ raw_value = '2'; value = 'Äventyrarna' }
                        roles = @{
                            value = @{
                                troop = @{
                                    '2' = @{
                                        '3' = @{ role_name = 'Ledare' }
                                    }
                                }
                                group = @{}
                            }
                        }
                    }
                }
            }
            $createdGroup = [pscustomobject]@{
                Id = 'created-group'
                DisplayName = 'Avdelningsledare Äventyrarna'
                GroupTypes = @('Unified')
                MailEnabled = $true
            }

            Mock Get-SNSApiGroupMemberlist { return $memberlist }
            Mock Get-EXOMailbox {
                return @(
                    [pscustomobject]@{
                        CustomAttribute1 = 'member-1'
                        ExternalDirectoryObjectId = 'user-1'
                    }
                )
            }
            Mock Get-MgGroup { return @() }
            Mock New-MgGroup { return $createdGroup }
            Mock Get-MgGroupMember { return @() }
            Mock New-MgGroupMember {}
            Mock Remove-MgGroupMemberDirectoryObjectByRef {}
            Mock Write-SNSLog {}

            Update-SNMicrosoft365Group -Configuration $configuration

            Should -Invoke New-MgGroup -Times 1 -Exactly -ParameterFilter {
                $BodyParameter.DisplayName -eq 'Avdelningsledare Äventyrarna' -and
                $BodyParameter.MailNickname -eq 'avdelningsledare-aventyrarna' -and
                $BodyParameter.GroupTypes -contains 'Unified' -and
                $BodyParameter.MailEnabled -eq $true -and
                $BodyParameter.SecurityEnabled -eq $false
            }
            Should -Invoke New-MgGroupMember -Times 1 -Exactly -ParameterFilter {
                $GroupId -eq 'created-group' -and $DirectoryObjectId -eq 'user-1'
            }
        }
    }

    It 'adds matching account holders, removes stale members, and logs matching members without accounts' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXOMailbox {}
            function Get-MgGroup {
                param($Filter, $Property, $All, $ErrorAction)
            }
            function Get-MgGroupMember {
                param($GroupId, $All, $ErrorAction)
            }
            function Get-MgGroupOwner {
                param($GroupId, $All, $ErrorAction)
            }
            function New-MgGroupMember {
                param($GroupId, $DirectoryObjectId, $ErrorAction)
            }
            function Remove-MgGroupMemberDirectoryObjectByRef {
                param($GroupId, $DirectoryObjectId, $ErrorAction)
            }

            $configuration = [SNSConfiguration]::new()
            $configuration.CredentialMemberlist = [pscredential]::new(
                'test',
                (ConvertTo-SecureString 'test' -AsPlainText -Force)
            )
            $configuration.Microsoft365GroupSettings = @{
                'Kårledning' = @{
                    source = 'group_roles'
                    mail_nickname = 'karledning'
                    roles = @('Kårledare')
                }
            }
            $memberlist = @{
                data = @{
                    member1 = @{
                        member_no = @{ value = 'member-1' }
                        roles = @{ value = @{ group = @{ g1 = @{ r1 = @{ role_name = 'Kårledare' } } }; troop = @{} } }
                    }
                    member2 = @{
                        member_no = @{ value = 'member-2' }
                        roles = @{ value = @{ group = @{ g2 = @{ r2 = @{ role_name = 'Kårledare' } } }; troop = @{} } }
                    }
                    member3 = @{
                        member_no = @{ value = 'member-3' }
                        roles = @{ value = @{ group = @{ g3 = @{ r3 = @{ role_name = 'Kårledare' } } }; troop = @{} } }
                    }
                    member4 = @{
                        member_no = @{ value = 'member-4' }
                        roles = @{ value = @{ group = @{ g4 = @{ r4 = @{ role_name = 'Scout' } } }; troop = @{} } }
                    }
                }
            }
            $existingGroup = [pscustomobject]@{
                Id = 'group-1'
                DisplayName = 'Kårledning'
                GroupTypes = @('Unified')
                MailEnabled = $true
            }

            Mock Get-SNSApiGroupMemberlist { return $memberlist }
            Mock Get-EXOMailbox {
                return @(
                    [pscustomobject]@{ CustomAttribute1 = 'member-1'; ExternalDirectoryObjectId = 'user-1' }
                    [pscustomobject]@{ CustomAttribute1 = 'member-2'; ExternalDirectoryObjectId = 'user-2' }
                    [pscustomobject]@{ CustomAttribute1 = 'member-4'; ExternalDirectoryObjectId = 'user-4' }
                )
            }
            Mock Get-MgGroup { return @($existingGroup) }
            Mock Get-MgGroupMember {
                return @(
                    [pscustomobject]@{ Id = 'user-1' }
                    [pscustomobject]@{ Id = 'obsolete-user' }
                )
            }
            Mock Get-MgGroupOwner { return @() }
            Mock New-MgGroupMember {}
            Mock Remove-MgGroupMemberDirectoryObjectByRef {}
            Mock Write-SNSLog {}

            Update-SNMicrosoft365Group -Configuration $configuration

            Should -Invoke New-MgGroupMember -Times 1 -Exactly -ParameterFilter {
                $GroupId -eq 'group-1' -and $DirectoryObjectId -eq 'user-2'
            }
            Should -Invoke Remove-MgGroupMemberDirectoryObjectByRef -Times 1 -Exactly -ParameterFilter {
                $GroupId -eq 'group-1' -and $DirectoryObjectId -eq 'obsolete-user'
            }
            Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
                $Level -eq 'Info' -and $Message -like "*member-3*has no Microsoft 365 user account*"
            }
        }
    }

    It 'matches any role in the selected troop or scout group when roles is *' {
        InModuleScope Office365-Scoutnet-synk {
            $unitNames = @{ '2' = 'Äventyrarna'; '3' = 'Spårarna' }
            $troopMember = @{
                roles = @{ value = @{ group = @{}; troop = @{ '2' = @{ '9' = @{ role_name = 'Vad som helst' } } } } }
            }
            $otherTroopMember = @{
                roles = @{ value = @{ group = @{}; troop = @{ '3' = @{ '9' = @{ role_name = 'Ledare' } } } } }
            }
            $noRoleMember = @{
                roles = @{ value = @{ group = @{}; troop = @{ '2' = @{} } } }
            }
            $troopRule = @{ troops = @(@{ name = 'Äventyrarna'; id = '2'; roles = '*' }) }

            Test-SNMicrosoft365RoleGroupMatch -Member $troopMember -Source 'troop_roles' -Setting $troopRule -UnitNamesById $unitNames |
                Should -BeTrue
            Test-SNMicrosoft365RoleGroupMatch -Member $otherTroopMember -Source 'troop_roles' -Setting $troopRule -UnitNamesById $unitNames |
                Should -BeFalse
            Test-SNMicrosoft365RoleGroupMatch -Member $noRoleMember -Source 'troop_roles' -Setting $troopRule -UnitNamesById $unitNames |
                Should -BeFalse

            $groupMember = @{
                roles = @{ value = @{ group = @{ '1' = @{ '5' = @{ role_name = 'Kassör' } } }; troop = @{} } }
            }
            Test-SNMicrosoft365RoleGroupMatch -Member $groupMember -Source 'group_roles' -Setting @{ roles = '*' } -UnitNamesById $unitNames |
                Should -BeTrue
            Test-SNMicrosoft365RoleGroupMatch -Member $troopMember -Source 'group_roles' -Setting @{ roles = '*' } -UnitNamesById $unitNames |
                Should -BeFalse
        }
    }

    It 'includes unit members when unit_members is set on a troop rule' {
        InModuleScope Office365-Scoutnet-synk {
            $unitNames = @{ '2' = 'Äventyrarna' }
            $scout = @{ unit = @{ raw_value = '2'; value = 'Äventyrarna' }; roles = @{ value = @{ group = @{}; troop = @{} } } }
            $other = @{ unit = @{ raw_value = '3'; value = 'Spårarna' }; roles = @{ value = @{ group = @{}; troop = @{} } } }
            $rule = @{ troops = @(@{ name = 'Fel namn'; id = '2'; unit_members = $true }) }

            Test-SNMicrosoft365RoleGroupMatch -Member $scout -Source 'troop_roles' -Setting $rule -UnitNamesById $unitNames |
                Should -BeTrue
            Test-SNMicrosoft365RoleGroupMatch -Member $other -Source 'troop_roles' -Setting $rule -UnitNamesById $unitNames |
                Should -BeFalse
            Test-SNMicrosoft365RoleGroupMatch -Member $scout -Source 'troop_roles' -Setting $rule -UnitNamesById $unitNames -RolesOnly |
                Should -BeFalse
        }
    }

    It 'never removes group owners or keep_members accounts even if they do not match the rule' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXOMailbox {}
            function Get-MgGroup { param($Filter, $Property, $All, $ErrorAction) }
            function Get-MgGroupMember { param($GroupId, $All, $ErrorAction) }
            function Get-MgGroupOwner { param($GroupId, $All, $ErrorAction) }
            function Get-MgUser { param($UserId, $Property, $ErrorAction) }
            function New-MgGroupMember { param($GroupId, $DirectoryObjectId, $ErrorAction) }
            function Remove-MgGroupMemberDirectoryObjectByRef { param($GroupId, $DirectoryObjectId, $ErrorAction) }

            $configuration = [SNSConfiguration]::new()
            $configuration.CredentialMemberlist = [pscredential]::new(
                'test',
                (ConvertTo-SecureString 'test' -AsPlainText -Force)
            )
            $configuration.Microsoft365GroupSettings = @{
                'Styrelsen' = @{
                    source = 'group_roles'
                    mail_nickname = 'styrelsen'
                    roles = @('Ordförande')
                    keep_members = @('keeper@example.org', 'keeper-id-2')
                }
            }
            $existingGroup = [pscustomobject]@{
                Id = 'group-1'
                DisplayName = 'Styrelsen'
                GroupTypes = @('Unified')
                MailEnabled = $true
            }

            Mock Get-SNSApiGroupMemberlist { return @{ data = @{} } }
            Mock Get-EXOMailbox { return @() }
            Mock Get-MgGroup { return @($existingGroup) }
            Mock Get-MgGroupMember {
                return @(
                    [pscustomobject]@{ Id = 'owner-1' }
                    [pscustomobject]@{ Id = 'keeper-id-1' }
                    [pscustomobject]@{ Id = 'keeper-id-2' }
                    [pscustomobject]@{ Id = 'stale-user' }
                )
            }
            Mock Get-MgGroupOwner { return @([pscustomobject]@{ Id = 'owner-1' }) }
            Mock Get-MgUser { return [pscustomobject]@{ Id = 'keeper-id-1' } }
            Mock Remove-MgGroupMemberDirectoryObjectByRef {}
            Mock Write-SNSLog {}

            Update-SNMicrosoft365Group -Configuration $configuration

            Should -Invoke Remove-MgGroupMemberDirectoryObjectByRef -Times 1 -Exactly
            Should -Invoke Remove-MgGroupMemberDirectoryObjectByRef -Times 1 -Exactly -ParameterFilter {
                $DirectoryObjectId -eq 'stale-user'
            }
        }
    }

    It 'uses the configured Scoutnet mailing list and defaults to its leader members' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXOMailbox {}
            function Get-MgGroup {
                param($Filter, $Property, $All, $ErrorAction)
            }
            function Get-MgGroupMember {
                param($GroupId, $All, $ErrorAction)
            }
            function New-MgGroupMember {
                param($GroupId, $DirectoryObjectId, $ErrorAction)
            }
            function Remove-MgGroupMemberDirectoryObjectByRef {
                param($GroupId, $DirectoryObjectId, $ErrorAction)
            }

            $configuration = [SNSConfiguration]::new()
            $configuration.CredentialCustomlists = [pscredential]::new(
                'test',
                (ConvertTo-SecureString 'test' -AsPlainText -Force)
            )
            $configuration.CredentialMemberlist = [pscredential]::new(
                'test',
                (ConvertTo-SecureString 'test' -AsPlainText -Force)
            )
            $configuration.Microsoft365GroupSettings = @{
                'Scoutnet list group' = @{
                    source = 'mail_list'
                    mail_nickname = 'scoutnet-list-group'
                    scoutnet_list_id = 'list-1'
                }
            }
            $memberlist = @{ data = @{} }
            $existingGroup = [pscustomobject]@{
                Id = 'group-1'
                DisplayName = 'Scoutnet list group'
                GroupTypes = @('Unified')
                MailEnabled = $true
            }
            $mockCustomLists = @{
                'list-1' = @{
                    ledare = @('member-1')
                    scouter = @('member-2')
                }
            }

            Mock Get-SNSApiGroupMemberlist { return $memberlist }
            Mock Get-SNSMaillistInfo { return $mockCustomLists, 'hash' }
            Mock Get-EXOMailbox {
                return @(
                    [pscustomobject]@{ CustomAttribute1 = 'member-1'; ExternalDirectoryObjectId = 'user-1' }
                    [pscustomobject]@{ CustomAttribute1 = 'member-2'; ExternalDirectoryObjectId = 'user-2' }
                )
            }
            Mock Get-MgGroup { return @($existingGroup) }
            Mock Get-MgGroupMember { return @() }
            Mock New-MgGroupMember {}
            Mock Remove-MgGroupMemberDirectoryObjectByRef {}

            Update-SNMicrosoft365Group -Configuration $configuration

            Should -Invoke Get-SNSMaillistInfo -Times 1 -Exactly -ParameterFilter {
                $MaillistsIds.Count -eq 1 -and $MaillistsIds[0].scoutnet_list_id -eq 'list-1'
            }
            Should -Invoke New-MgGroupMember -Times 1 -Exactly -ParameterFilter {
                $GroupId -eq 'group-1' -and $DirectoryObjectId -eq 'user-1'
            }
        }
    }

    It 'leaves Microsoft 365 untouched when no group rules are configured' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-SNSApiGroupMemberlist {}
            $configuration = [SNSConfiguration]::new()
            Mock Get-SNSApiGroupMemberlist { throw 'The member list must not be requested' }
            Mock Write-SNSLog {}

            Update-SNMicrosoft365Group -Configuration $configuration

            Should -Invoke Get-SNSApiGroupMemberlist -Times 0 -Exactly
        }
    }

    It 'makes a group a Teams team once, with owners, when create_team is set' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXOMailbox {}
            function Get-MgGroup { param($Filter, $Property, $All, $ErrorAction) }
            function Get-MgGroupMember { param($GroupId, $All, $ErrorAction) }
            function Get-MgGroupOwner { param($GroupId, $All, $ErrorAction) }
            function Get-MgUser { param($UserId, $Property, $ErrorAction) }
            function New-MgGroupOwnerByRef { param($GroupId, $OdataId, $ErrorAction) }
            function New-MgTeam { param($BodyParameter, $ErrorAction) }

            $configuration = [SNSConfiguration]::new()
            $configuration.CredentialMemberlist = [pscredential]::new(
                'test',
                (ConvertTo-SecureString 'test' -AsPlainText -Force)
            )
            $configuration.Microsoft365GroupSettings = @{
                'Styrelsen' = @{
                    source = 'group_roles'
                    mail_nickname = 'styrelsen'
                    roles = @('Ordförande')
                    create_team = $true
                    team_owners = @('admin@example.org')
                }
            }
            $existingGroup = [pscustomobject]@{
                Id = 'group-1'
                DisplayName = 'Styrelsen'
                GroupTypes = @('Unified')
                MailEnabled = $true
                ResourceProvisioningOptions = @()
            }

            Mock Get-SNSApiGroupMemberlist { return @{ data = @{} } }
            Mock Get-EXOMailbox { return @() }
            Mock Get-MgGroup { return $existingGroup }
            Mock Get-MgGroupMember { return @() }
            Mock Get-MgGroupOwner { return @() }
            Mock Get-MgUser { return [pscustomobject]@{ Id = 'owner-1' } }
            Mock New-MgGroupOwnerByRef {}
            Mock New-MgTeam {}
            Mock Write-SNSLog {}

            Update-SNMicrosoft365Group -Configuration $configuration

            Should -Invoke New-MgGroupOwnerByRef -Times 1 -Exactly -ParameterFilter {
                $GroupId -eq 'group-1' -and $OdataId -like '*/users/owner-1'
            }
            Should -Invoke New-MgTeam -Times 1 -Exactly -ParameterFilter {
                $BodyParameter['group@odata.bind'] -like "*groups('group-1')"
            }

            $existingGroup.ResourceProvisioningOptions = @('Team')
            Update-SNMicrosoft365Group -Configuration $configuration
            Should -Invoke New-MgTeam -Times 1 -Exactly
        }
    }

    It 'rejects create_team without team_owners' {
        InModuleScope Office365-Scoutnet-synk {
            function New-MgTeam { param($BodyParameter, $ErrorAction) }

            $configuration = [SNSConfiguration]::new()
            $configuration.Microsoft365GroupSettings = @{
                'Styrelsen' = @{
                    source = 'group_roles'
                    mail_nickname = 'styrelsen'
                    roles = @('Ordförande')
                    create_team = $true
                }
            }

            { Update-SNMicrosoft365Group -Configuration $configuration } |
                Should -Throw '*team_owners*'
        }
    }
}
