BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Office365-Scoutnet-synk\Office365-Scoutnet-synk.psm1'
    Import-Module $modulePath -Force -ErrorAction Stop
}

AfterAll {
    Remove-Module Office365-Scoutnet-synk -ErrorAction SilentlyContinue
}

Describe 'Invoke-SNSCreateUserAndUpdateUserData' {
    It 'creates a user and deduplicates alternate mail addresses' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-MgUser {}
            function New-MgUser {
                param($BodyParameter, $ErrorAction)
            }
            function Set-MgUserLicense {}
            function New-MgGroupMember {}
            function Get-EXOMailbox {}
            function Set-Mailbox {}
            function Set-MailboxMessageConfiguration {}

            $credential = [pscredential]::new('test', (ConvertTo-SecureString 'test' -AsPlainText -Force))
            $script:SNSConf = [SNSConfiguration]::new()
            $script:SNSConf.DomainName = 'example.org'
            $script:SNSConf.WaitMailboxCreationMaxTime = '1'
            $script:SNSConf.WaitMailboxCreationPollTime = '0'
            $script:SNSConf.SignatureHtml = '<DisplayName>'
            $script:SNSConf.SignatureText = 'Hello <DisplayName>'
            $script:SNSConf.CredentialMemberlist = $credential
            $member = [pscustomobject]@{
                member_no = [pscustomobject]@{ value = 'member-1' }
                first_name = [pscustomobject]@{ value = 'Ada' }
                last_name = [pscustomobject]@{ value = 'Lovelace' }
                email = [pscustomobject]@{ value = 'ada@example.org' }
                contact_alt_email = [pscustomobject]@{ value = 'ada@example.org' }
                postcode = [pscustomobject]@{ value = '' }
                town = [pscustomobject]@{ value = '' }
                country = [pscustomobject]@{ value = '' }
                contact_mobile_phone = [pscustomobject]@{ value = '' }
                address_1 = [pscustomobject]@{ value = '' }
                address_2 = [pscustomobject]@{ value = '' }
                address_3 = [pscustomobject]@{ value = '' }
            }
            Mock Write-SNSLog {}
            Mock Get-SNSSecurityGroupScoutnet { return [pscustomobject]@{ Id = 'active-group' } }
            Mock Get-RandomPassword { return 'TestPassword123' }
            Mock Get-MgUser { return $null }
            Mock New-MgUser {
                return [pscustomobject]@{
                    Id = 'user-id'
                    UserPrincipalName = 'ada.lovelace@example.org'
                    DisplayName = 'Ada Lovelace'
                }
            }
            Mock Set-MgUserLicense {}
            Mock New-MgGroupMember {}
            Mock Start-Sleep {}
            Mock Get-EXOMailbox {}
            Mock Set-Mailbox {}
            Mock Set-MailboxMessageConfiguration {}

            Invoke-SNSCreateUserAndUpdateUserData -memberData @($member) -userLicenseData @()

            Should -Invoke New-MgUser -Times 1 -Exactly -ParameterFilter {
                $BodyParameter.UserPrincipalName -eq 'ada.lovelace@example.org' -and
                $BodyParameter.OtherMails.Count -eq 1 -and
                $BodyParameter.OtherMails[0] -eq 'ada@example.org'
            }
            Should -Invoke Set-MgUserLicense -Times 1 -Exactly
            Should -Invoke New-MgGroupMember -Times 1 -Exactly
            Should -Invoke Set-Mailbox -Times 1 -Exactly
        }
    }

    It 'does not attempt to create a user without a Scoutnet email address' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-MgUser {}
            function New-MgUser {}

            $script:SNSConf = [SNSConfiguration]::new()
            $member = [pscustomobject]@{
                email = [pscustomobject]@{ value = ' ' }
                first_name = [pscustomobject]@{ value = 'Ada' }
                last_name = [pscustomobject]@{ value = 'Lovelace' }
            }

            Mock Write-SNSLog {}
            Mock Get-SNSSecurityGroupScoutnet { return [pscustomobject]@{ Id = 'active-group' } }
            Mock Get-MgUser {}
            Mock New-MgUser {}

            Invoke-SNSCreateUserAndUpdateUserData -memberData @($member) -userLicenseData @()

            Should -Invoke Get-MgUser -Times 0 -Exactly
            Should -Invoke New-MgUser -Times 0 -Exactly
        }
    }
}

Describe 'Invoke-SNSUpdateAccount' {
    It 'updates changed profile fields and keeps duplicate Scoutnet email addresses unique' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-MgUser {}
            function Update-MgUser {
                param($UserId, $BodyParameter, $ErrorAction)
            }

            $credential = [pscredential]::new('test', (ConvertTo-SecureString 'test' -AsPlainText -Force))
            $accountData = [pscustomobject]@{
                name = 'Ada Lovelace'
                CustomAttribute1 = 'member-1'
                ExternalDirectoryObjectId = 'user-id'
                UserPrincipalName = 'ada.lovelace@example.org'
            }
            $member = [pscustomobject]@{
                address_1 = [pscustomobject]@{ value = '1 Main Street' }
                address_2 = [pscustomobject]@{ value = '' }
                address_3 = [pscustomobject]@{ value = '' }
                postcode = [pscustomobject]@{ value = '12345' }
                town = [pscustomobject]@{ value = 'London' }
                country = [pscustomobject]@{ value = 'GB' }
                first_name = [pscustomobject]@{ value = 'Ada' }
                last_name = [pscustomobject]@{ value = 'Lovelace' }
                contact_mobile_phone = [pscustomobject]@{ value = '+44123456789' }
                email = [pscustomobject]@{ value = 'ada@example.org' }
                contact_alt_email = [pscustomobject]@{ value = 'ada@example.org' }
            }
            $existingUser = [pscustomobject]@{
                StreetAddress = ''
                PostalCode = ''
                City = ''
                Country = ''
                DisplayName = ''
                GivenName = ''
                Surname = ''
                MobilePhone = ''
                OtherMails = @()
            }
            Mock Write-SNSLog {}
            Mock Get-SNSApiGroupMemberlist { return @{ data = @{ 'member-1' = $member } } }
            Mock Get-MgUser { return $existingUser }
            Mock Update-MgUser {}

            Invoke-SNSUpdateAccount -AccountData $accountData -CredentialMemberlist $credential

            Should -Invoke Update-MgUser -Times 1 -Exactly -ParameterFilter {
                $BodyParameter.StreetAddress -eq '1 Main Street' -and
                $BodyParameter.DisplayName -eq 'Ada Lovelace' -and
                $BodyParameter.OtherMails.Count -eq 1 -and
                $BodyParameter.OtherMails[0] -eq 'ada@example.org'
            }
        }
    }
}

Describe 'Invoke-SNSDisableAccount' {
    It 'disables the account and moves it from the active group to the disabled group' {
        InModuleScope Office365-Scoutnet-synk {
            function Update-MgUser {
                param($UserId, $AccountEnabled, $ErrorAction)
            }
            function Set-Mailbox {}
            function Remove-MgGroupMemberByRef {
                param($GroupId, $DirectoryObjectId, $ErrorAction)
            }
            function New-MgGroupMember {
                param($GroupId, $DirectoryObjectId, $ErrorAction)
            }
            function Set-MailboxAutoReplyConfiguration {
                param($Identity, $AutoReplyState, $ExternalAudience, $InternalMessage, $ExternalMessage)
            }

            $script:SNSConf = [SNSConfiguration]::new()
            $script:SNSConf.DisabledAccountsAutoReplyText = 'Contact <DisplayName>'
            $accountData = [pscustomobject]@{
                name = 'Ada Lovelace'
                DisplayName = 'Ada Lovelace'
                ExternalDirectoryObjectId = 'user-id'
                UserPrincipalName = 'ada.lovelace@example.org'
                Identity = 'ada.lovelace@example.org'
            }
            Mock Write-SNSLog {}
            Mock Update-MgUser {}
            Mock Set-Mailbox {}
            Mock Get-SNSSecurityGroupScoutnet { return [pscustomobject]@{ Id = 'active-group' } }
            Mock Get-SNSDisabledScoutnetUserGroup { return [pscustomobject]@{ Id = 'disabled-group' } }
            Mock Remove-MgGroupMemberByRef {}
            Mock New-MgGroupMember {}
            Mock Set-MailboxAutoReplyConfiguration {}

            Invoke-SNSDisableAccount -AccountData $accountData

            Should -Invoke Update-MgUser -Times 1 -Exactly -ParameterFilter {
                $UserId -eq 'user-id' -and $AccountEnabled -eq $false
            }
            Should -Invoke Remove-MgGroupMemberByRef -Times 1 -Exactly -ParameterFilter {
                $GroupId -eq 'active-group' -and $DirectoryObjectId -eq 'user-id'
            }
            Should -Invoke New-MgGroupMember -Times 1 -Exactly -ParameterFilter {
                $GroupId -eq 'disabled-group' -and $DirectoryObjectId -eq 'user-id'
            }
            Should -Invoke Set-MailboxAutoReplyConfiguration -Times 1 -Exactly -ParameterFilter {
                $InternalMessage -eq 'Contact Ada Lovelace'
            }
        }
    }
}

Describe 'Invoke-SNSEnableAccount' {
    It 'enables the account, moves it to the active group, and disables autoreply' {
        InModuleScope Office365-Scoutnet-synk {
            function Update-MgUser {
                param($UserId, $AccountEnabled, $ErrorAction)
            }
            function Set-Mailbox {
                param($Identity, $HiddenFromAddressListsEnabled, $ErrorAction)
            }
            function Remove-MgGroupMemberByRef {
                param($GroupId, $DirectoryObjectId, $ErrorAction)
            }
            function New-MgGroupMember {
                param($GroupId, $DirectoryObjectId, $ErrorAction)
            }
            function Set-MailboxAutoReplyConfiguration {
                param($Identity, $AutoReplyState, $InternalMessage, $ExternalMessage)
            }

            $script:SNSConf = [SNSConfiguration]::new()
            $accountData = [pscustomobject]@{
                name = 'Ada Lovelace'
                ExternalDirectoryObjectId = 'user-id'
                UserPrincipalName = 'ada.lovelace@example.org'
                Identity = 'ada.lovelace@example.org'
            }

            Mock Write-SNSLog {}
            Mock Update-MgUser {}
            Mock Set-Mailbox {}
            Mock Get-SNSDisabledScoutnetUserGroup { return [pscustomobject]@{ Id = 'disabled-group' } }
            Mock Get-SNSSecurityGroupScoutnet { return [pscustomobject]@{ Id = 'active-group' } }
            Mock Remove-MgGroupMemberByRef {}
            Mock New-MgGroupMember {}
            Mock Set-MailboxAutoReplyConfiguration {}

            Invoke-SNSEnableAccount -AccountData $accountData

            Should -Invoke Update-MgUser -Times 1 -Exactly -ParameterFilter {
                $UserId -eq 'user-id' -and $AccountEnabled -eq $true
            }
            Should -Invoke Set-Mailbox -Times 1 -Exactly -ParameterFilter {
                $Identity -eq 'ada.lovelace@example.org' -and
                $HiddenFromAddressListsEnabled -eq $false
            }
            Should -Invoke Remove-MgGroupMemberByRef -Times 1 -Exactly -ParameterFilter {
                $GroupId -eq 'disabled-group' -and $DirectoryObjectId -eq 'user-id'
            }
            Should -Invoke New-MgGroupMember -Times 1 -Exactly -ParameterFilter {
                $GroupId -eq 'active-group' -and $DirectoryObjectId -eq 'user-id'
            }
            Should -Invoke Set-MailboxAutoReplyConfiguration -Times 1 -Exactly -ParameterFilter {
                $Identity -eq 'ada.lovelace@example.org' -and $AutoReplyState -eq 'Disabled'
            }
        }
    }
}

Describe 'Get-SNSSecurityGroup' {
    It 'returns an existing group with the matching display name' {
        InModuleScope Office365-Scoutnet-synk {
            $matchingGroup = [pscustomobject]@{
                Id = 'matching-group'
                DisplayName = 'Scoutnet users'
            }
            Mock Get-MgGroup {
                return @(
                    [pscustomobject]@{ Id = 'other-group'; DisplayName = 'Other users' }
                    $matchingGroup
                )
            }
            Mock New-MgGroup {}

            $group = Get-SNSSecurityGroup -Name 'Scoutnet users' -Description 'Synchronized users'

            $group.Id | Should -Be 'matching-group'
            Should -Invoke Get-MgGroup -Times 1 -Exactly -ParameterFilter {
                $Search -eq 'displayName:Scoutnet users' -and $ConsistencyLevel -eq 'eventual'
            }
            Should -Invoke New-MgGroup -Times 0 -Exactly
        }
    }

    It 'creates a security group when no matching group exists' {
        InModuleScope Office365-Scoutnet-synk {
            $createdGroup = [pscustomobject]@{
                Id = 'new-group'
                DisplayName = 'Scoutnet users'
            }
            Mock Get-MgGroup {
                return @([pscustomobject]@{ Id = 'other-group'; DisplayName = 'Other users' })
            }
            Mock New-MgGroup { return $createdGroup }
            Mock Write-SNSLog {}

            $group = Get-SNSSecurityGroup -Name 'Scoutnet users' -Description 'Synchronized users'

            $group.Id | Should -Be 'new-group'
            Should -Invoke New-MgGroup -Times 1 -Exactly -ParameterFilter {
                $BodyParameter.DisplayName -eq 'Scoutnet users' -and
                $BodyParameter.Description -eq 'Synchronized users' -and
                $BodyParameter.MailEnabled -eq $false -and
                $BodyParameter.SecurityEnabled -eq $true -and
                $BodyParameter.MailNickname -eq 'Scoutnet users' -and
                $BodyParameter.GroupTypes.Count -eq 0
            }
            Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
                $Level -eq 'Warn' -and $Message -like '*Creating the group*'
            }
        }
    }
}

Describe 'Get-SNSUsersInSecurityGroup' {
    It 'returns only group members that match Office 365 users, always as a collection' {
        InModuleScope Office365-Scoutnet-synk {
            $expectedUser = [pscustomobject]@{
                ExternalDirectoryObjectId = 'user-1'
                DisplayName = 'Ada Lovelace'
            }
            $allOffice365Users = @(
                $expectedUser
                [pscustomobject]@{
                    ExternalDirectoryObjectId = 'user-2'
                    DisplayName = 'Grace Hopper'
                }
            )
            Mock Get-SNSSecurityGroup {
                return [pscustomobject]@{ Id = 'security-group' }
            }
            Mock Get-MgGroupMember {
                return @(
                    [pscustomobject]@{ Id = 'user-1' }
                    [pscustomobject]@{ Id = 'deleted-user' }
                )
            }

            $members = Get-SNSUsersInSecurityGroup -Name 'Scoutnet users' `
                -Description 'Synchronized users' -allOffice365Users $allOffice365Users

            ($members -is [System.Collections.ArrayList]) | Should -BeTrue
            $members.Count | Should -Be 1
            $members[0].ExternalDirectoryObjectId | Should -Be 'user-1'
            Should -Invoke Get-MgGroupMember -Times 1 -Exactly -ParameterFilter {
                $GroupId -eq 'security-group'
            }
        }
    }
}

Describe 'Invoke-SNSUppdateOffice365User' {
            It 'creates, updates, re-enables, disables, and retains users according to Scoutnet membership' {
                InModuleScope Office365-Scoutnet-synk {
                    function Get-MgSubscribedSku {}
                    function Get-EXOMailbox {}

                    $configuration = [SNSConfiguration]::new()
                    $configuration.CredentialCustomlists = [pscredential]::new(
                        'test',
                        [System.Security.SecureString]::new()
                    )
                    $configuration.CredentialMemberlist = $configuration.CredentialCustomlists
                    $activeAccount = [pscustomobject]@{
                        Name = 'Active Scout'
                        CustomAttribute1 = 'active-member'
                        ExternalDirectoryObjectId = 'active-id'
                    }
                    $returningAccount = [pscustomobject]@{
                        Name = 'Returning Scout'
                        CustomAttribute1 = 'returning-member'
                        ExternalDirectoryObjectId = 'returning-id'
                    }
                    $disabledAccount = [pscustomobject]@{
                        Name = 'Still Disabled'
                        CustomAttribute1 = 'disabled-member'
                        ExternalDirectoryObjectId = 'disabled-id'
                    }
                    $leavingAccount = [pscustomobject]@{
                        Name = 'Leaving Scout'
                        CustomAttribute1 = 'leaving-member'
                        ExternalDirectoryObjectId = 'leaving-id'
                    }
                    $unconnectedAccount = [pscustomobject]@{
                        Name = 'Unconnected User'
                        CustomAttribute1 = 'unconnected-member'
                        ExternalDirectoryObjectId = 'unconnected-id'
                    }
                    $disabledGroupAccounts = [System.Collections.ArrayList]::new()
                    [void]$disabledGroupAccounts.Add($returningAccount)
                    [void]$disabledGroupAccounts.Add($disabledAccount)
                    $activeGroupAccounts = [System.Collections.ArrayList]::new()
                    [void]$activeGroupAccounts.Add($activeAccount)
                    [void]$activeGroupAccounts.Add($leavingAccount)
                    $newMemberData = [pscustomobject]@{ member_no = 'new-member' }
                    $scoutnetData = @{
                        data = @{
                            'new-member' = $newMemberData
                        }
                    }
                    $allAccounts = @(
                        $activeAccount
                        $returningAccount
                        $disabledAccount
                        $leavingAccount
                        $unconnectedAccount
                    )

                    Mock Write-SNSLog {}
                    Mock Get-EXOMailbox { return $allAccounts }
                    Mock Get-SNSUsersInSecurityGroupScoutnet { return $activeGroupAccounts }
                    Mock Get-SNSDisabledScoutnetUserList { return $disabledGroupAccounts }
                    Mock Get-SNSScoutnetLeaderList {
                        return [System.Collections.ArrayList]@(
                            'active-member',
                            'returning-member',
                            'new-member',
                            'missing-data'
                        )
                    }
                    Mock Get-SNSApiGroupMemberlist { return $scoutnetData }
                    Mock Invoke-SNSCreateUserAndUpdateUserData {}
                    Mock Invoke-SNSEnableAccount {}
                    Mock Invoke-SNSDisableAccount {}
                    Mock Invoke-SNSUpdateAccount {}

                    Invoke-SNSUppdateOffice365User -Configuration $configuration

                    Should -Invoke Invoke-SNSCreateUserAndUpdateUserData -Times 1 -Exactly -ParameterFilter {
                        $memberData.Count -eq 1 -and $memberData[0].member_no -eq 'new-member'
                    }
                    Should -Invoke Invoke-SNSEnableAccount -Times 1 -Exactly -ParameterFilter {
                        $AccountData.CustomAttribute1 -eq 'returning-member'
                    }
                    Should -Invoke Invoke-SNSDisableAccount -Times 1 -Exactly -ParameterFilter {
                        $AccountData.CustomAttribute1 -eq 'leaving-member'
                    }
                    Should -Invoke Invoke-SNSUpdateAccount -Times 2 -Exactly
                    Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
                        $Level -eq 'Warn' -and $Message -like '*missing-data*'
                    }
                    Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
                        $Message -eq "Account not connected to Scoutnet 'Unconnected User'"
                    }
                }
            }

            It 'aborts before processing accounts when Scoutnet returns no data' {
                InModuleScope Office365-Scoutnet-synk {
                    function Get-EXOMailbox {}

                    $configuration = [SNSConfiguration]::new()
                    $configuration.CredentialCustomlists = [pscredential]::new(
                        'test',
                        [System.Security.SecureString]::new()
                    )
                    $configuration.CredentialMemberlist = $configuration.CredentialCustomlists
                    Mock Write-SNSLog {}
                    Mock Get-EXOMailbox {
                        return [pscustomobject]@{ CustomAttribute1 = 'existing-member' }
                    }
                    Mock Get-SNSUsersInSecurityGroupScoutnet { return [System.Collections.ArrayList]::new() }
                    Mock Get-SNSDisabledScoutnetUserList { return [System.Collections.ArrayList]::new() }
                    Mock Get-SNSScoutnetLeaderList { return [System.Collections.ArrayList]::new() }
                    Mock Get-SNSApiGroupMemberlist { return $null }
                    Mock Invoke-SNSCreateUserAndUpdateUserData {}

                    { Invoke-SNSUppdateOffice365User -Configuration $configuration } |
                        Should -Throw '*No data returned from Scoutnet*'
                    Should -Invoke Invoke-SNSCreateUserAndUpdateUserData -Times 0 -Exactly
                }
            }

            It 'logs and stops when a configured license SKU cannot be found' {
                InModuleScope Office365-Scoutnet-synk {
                    function Get-MgSubscribedSku {}
                    function Get-EXOMailbox {}

                    $configuration = [SNSConfiguration]::new()
                    $configuration.LicenseAssignment = @{ 'Missing SKU' = @() }
                    Mock Write-SNSLog {}
                    Mock Get-MgSubscribedSku { return @() }
                    Mock Get-EXOMailbox {}

                    Invoke-SNSUppdateOffice365User -Configuration $configuration

                    Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
                        $Level -eq 'Error' -and $Message -like "*Missing SKU*"
                    }
                    Should -Invoke Get-EXOMailbox -Times 0 -Exactly
                }
            }
}

Describe 'Get-SNSUsersInSecurityGroup missing group' {
    It 'returns an empty collection without querying members when the group does not exist' {
        InModuleScope Office365-Scoutnet-synk {
            Mock Get-SNSSecurityGroup { return $null }
            Mock Get-MgGroupMember {}

            $members = Get-SNSUsersInSecurityGroup -Name 'Missing group' `
                -Description 'Group not found' -allOffice365Users @()

            ($members -is [System.Collections.ArrayList]) | Should -BeTrue
            $members.Count | Should -Be 0
            Should -Invoke Get-MgGroupMember -Times 0 -Exactly
        }
    }
}
