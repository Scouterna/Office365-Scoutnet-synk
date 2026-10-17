BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Office365-Scoutnet-synk'

    . (Join-Path $modulePath 'SNSConfiguration.ps1')
    . (Join-Path $modulePath 'Write-SNSLog.ps1')
    . (Join-Path $modulePath 'Receive-SNSApiJson.ps1')
    . (Join-Path $modulePath 'Get-SNSApiGroupMemberlist.ps1')
    . (Join-Path $modulePath 'Get-SNSApiGroupCustomlist.ps1')
    . (Join-Path $modulePath 'Get-SNSMaillistInfo.ps1')
    . (Join-Path $modulePath 'Get-SNSUserEmail.ps1')
    . (Join-Path $modulePath 'Invoke-SNSUppdateOffice365User.ps1')

    $credential = [pscredential]::new('test', (ConvertTo-SecureString 'test' -AsPlainText -Force))
}

Describe 'New-SNSConfiguration' {
    It 'creates a configuration with defaults and supplied values' {
        $configuration = New-SNSConfiguration -LicenseAssignment @{} `
            -DomainName 'example.org' -LogFilePath 'test.log' `
            -AllUsersGroupName 'all-users'

        $configuration | Should -BeOfType [SNSConfiguration]
        $configuration.DomainName | Should -Be 'example.org'
        $configuration.LogFilePath | Should -Be 'test.log'
        $configuration.AllUsersGroupName | Should -Be 'all-users'
        $configuration.PreferredLanguage | Should -Be 'sv-SE'
        $configuration.UsageLocation | Should -Be 'SE'
    }
}

Describe 'Write-SNSLog' {
    BeforeEach {
        $script:SNSConf = [SNSConfiguration]::new()
        $script:SNSConf.LogFilePath = Join-Path $TestDrive 'sns.log'
    }

    It 'creates a log file containing the timestamp, level, and message' {
        Write-SNSLog -Message 'Offline test message'

        $log = Get-Content -LiteralPath $script:SNSConf.LogFilePath -Raw
        $log | Should -Match '\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} INFO: Offline test message'
    }
}

Describe 'Get-SNSApiGroupMemberlist cache' {
    BeforeEach {
        $script:SNSConf = [SNSConfiguration]::new()
        $script:SNSConf.ApiGroupMemberlistCache = $null

        Mock Receive-SNSApiJson {
            return @{ data = @{ id = 'member-list' } }
        }
    }

    It 'fetches the member list once and reuses the cached response' {
        $first = Get-SNSApiGroupMemberlist -Credential $credential -Uri 'https://example.org/first'
        $second = Get-SNSApiGroupMemberlist -Credential $credential -Uri 'https://example.org/second'

        $first.data.id | Should -Be 'member-list'
        $second.data.id | Should -Be 'member-list'
        Should -Invoke Receive-SNSApiJson -Times 1 -Exactly
    }
}

Describe 'Get-SNSUserEmail' {
    BeforeEach {
        Mock Get-SNSApiGroupMemberlist {
            return @{
                data = @{
                    values = @(
                        [pscustomobject]@{
                            member_no = [pscustomobject]@{ value = '1001' }
                            first_name = [pscustomobject]@{ value = 'Ada' }
                            last_name = [pscustomobject]@{ value = 'Lovelace' }
                            email = [pscustomobject]@{ value = 'ada@example.org' }
                            contact_email_mum = [pscustomobject]@{ value = 'parent@example.org' }
                            contact_email_dad = [pscustomobject]@{ value = 'parent@example.org' }
                            contact_alt_email = [pscustomobject]@{ value = 'alt@example.org' }
                        }
                        [pscustomobject]@{
                            member_no = [pscustomobject]@{ value = '1002' }
                            first_name = [pscustomobject]@{ value = 'Grace' }
                            last_name = [pscustomobject]@{ value = 'Hopper' }
                            email = [pscustomobject]@{ value = ' ' }
                            contact_email_mum = [pscustomobject]@{ value = '' }
                            contact_email_dad = [pscustomobject]@{ value = $null }
                            contact_alt_email = [pscustomobject]@{ value = '' }
                        }
                    )
                }
            }
        }
    }

    It 'collects unique member and contact addresses and omits blank addresses' {
        $mailData, $hash = Get-SNSUserEmail -CredentialMemberlist $credential

        $mailData['1001'].primary_email | Should -Be 'ada@example.org'
        $mailData['1001'].alt_email | Should -Be 'alt@example.org'
        @($mailData['1001'].mailaddresses) | Should -HaveCount 3
        @($mailData['1001'].contacts_addresses) | Should -HaveCount 1
        $mailData['1002'].mailaddresses | Should -BeNullOrEmpty
        $mailData['1002'].contacts_addresses | Should -BeNullOrEmpty
        $mailData['1002'].ContainsKey('primary_email') | Should -Be $false
        $mailData['1002'].ContainsKey('alt_email') | Should -Be $false
        $hash | Should -Match '^[0-9A-F]{8}$'
    }

    It 'handles parent-only, alternate-only, and duplicate primary and alternate addresses' {
        Mock Get-SNSApiGroupMemberlist {
            return @{
                data = @{
                    values = @(
                        [pscustomobject]@{
                            member_no = [pscustomobject]@{ value = 'parent-only' }
                            first_name = [pscustomobject]@{ value = 'Parent' }
                            last_name = [pscustomobject]@{ value = 'Only' }
                            email = [pscustomobject]@{ value = '' }
                            contact_email_mum = [pscustomobject]@{ value = 'parent@example.org' }
                            contact_email_dad = [pscustomobject]@{ value = ' ' }
                            contact_alt_email = [pscustomobject]@{ value = '' }
                        }
                        [pscustomobject]@{
                            member_no = [pscustomobject]@{ value = 'alternate-only' }
                            first_name = [pscustomobject]@{ value = 'Alternate' }
                            last_name = [pscustomobject]@{ value = 'Only' }
                            email = [pscustomobject]@{ value = $null }
                            contact_email_mum = [pscustomobject]@{ value = '' }
                            contact_email_dad = [pscustomobject]@{ value = '' }
                            contact_alt_email = [pscustomobject]@{ value = 'alternate@example.org' }
                        }
                        [pscustomobject]@{
                            member_no = [pscustomobject]@{ value = 'duplicate' }
                            first_name = [pscustomobject]@{ value = 'Duplicate' }
                            last_name = [pscustomobject]@{ value = 'Address' }
                            email = [pscustomobject]@{ value = 'same@example.org' }
                            contact_email_mum = [pscustomobject]@{ value = '' }
                            contact_email_dad = [pscustomobject]@{ value = '' }
                            contact_alt_email = [pscustomobject]@{ value = 'same@example.org' }
                        }
                    )
                }
            }
        }

        $mailData, $null = Get-SNSUserEmail -CredentialMemberlist $credential

        $mailData['parent-only'].ContainsKey('primary_email') | Should -Be $false
        @($mailData['parent-only'].mailaddresses) | Should -Be @('parent@example.org')
        @($mailData['parent-only'].contacts_addresses) | Should -Be @('parent@example.org')
        $mailData['alternate-only'].ContainsKey('primary_email') | Should -Be $false
        $mailData['alternate-only'].alt_email | Should -Be 'alternate@example.org'
        @($mailData['alternate-only'].mailaddresses) | Should -Be @('alternate@example.org')
        @($mailData['duplicate'].mailaddresses) | Should -Be @('same@example.org')
    }
}

Describe 'Get-SNSApiGroupCustomlist' {
    It 'builds the list and rule query parameters' {
        Mock Receive-SNSApiJson {
            $script:receivedUri = $Uri
            return @{ success = $true }
        }

        $result = Get-SNSApiGroupCustomlist -Credential $credential `
            -Uri 'https://example.org/customlists' -listid 'list-1' -ruleid 'rule-2'

        $result.success | Should -Be $true
        $script:receivedUri | Should -Be 'https://example.org/customlists?list_id=list-1&rule_id=rule-2'
    }
}

Describe 'Get-SNSMaillistInfo' {
    BeforeEach {
        Mock Write-SNSLog {}
    }

    It 'separates leader and other members for configured dynamic lists' {
        Mock Get-SNSApiGroupCustomlist {
            if ([string]::IsNullOrWhiteSpace($listid)) {
                return @{
                    'list-1' = @{
                        title = 'Scout list'
                        rules = @{
                            'rule-leaders' = @{ title = 'Ledare'; id = 'leaders' }
                            'rule-scouts' = @{ title = 'Scouter'; id = 'scouts' }
                        }
                    }
                }
            }

            if ($ruleid -eq 'leaders') {
                return @{ data = @{ keys = @('leader-1') } }
            }

            return @{ data = @{ keys = @('scout-1', 'scout-2') } }
        }

        $lists, $hash = Get-SNSMaillistInfo -CredentialCustomlists $credential `
            -MaillistsIds @(@{ scoutnet_list_id = 'list-1' })

        $lists['list-1'].title | Should -Be 'Scout list'
        @($lists['list-1'].ledare) | Should -Be @('leader-1')
        @($lists['list-1'].scouter) | Should -Be @('scout-1', 'scout-2')
        $hash | Should -Match '^[0-9A-F]{8}$'
    }

    It 'puts static-list members in the leader collection without fetching rules' {
        Mock Get-SNSApiGroupCustomlist {
            if ([string]::IsNullOrWhiteSpace($listid)) {
                return @{
                    'list-1' = @{
                        title = 'Static list'
                        rules = @{}
                    }
                }
            }

            return @{ data = @{ keys = @('member-1', 'member-2') } }
        }

        $lists, $null = Get-SNSMaillistInfo -CredentialCustomlists $credential `
            -MaillistsIds @(@{ scoutnet_list_id = 'list-1'; statisk_lista = 'Ja' })

        @($lists['list-1'].ledare) | Should -Be @('member-1', 'member-2')
        $lists['list-1'].scouter | Should -BeNullOrEmpty
        Should -Invoke Get-SNSApiGroupCustomlist -Times 2 -Exactly
    }

    It 'ignores unconfigured lists and returns empty member collections for lists without rules' {
        Mock Get-SNSApiGroupCustomlist {
            return @{
                'configured-list' = @{ title = 'Configured'; rules = @{} }
                'unconfigured-list' = @{
                    title = 'Not configured'
                    rules = @{ 'rule-1' = @{ title = 'Scouter'; id = 'scouts' } }
                }
            }
        }

        $lists, $null = Get-SNSMaillistInfo -CredentialCustomlists $credential `
            -MaillistsIds @(@{ scoutnet_list_id = 'configured-list' })

        $lists.Keys | Should -Be @('configured-list')
        $lists['configured-list'].ledare | Should -BeNullOrEmpty
        $lists['configured-list'].scouter | Should -BeNullOrEmpty
        Should -Invoke Get-SNSApiGroupCustomlist -Times 1 -Exactly
    }

    It 'processes a duplicate Scoutnet list ID only once' {
        Mock Get-SNSApiGroupCustomlist {
            if ([string]::IsNullOrWhiteSpace($listid)) {
                return @{
                    'list-1' = @{
                        title = 'Shared list'
                        rules = @{ 'rule-1' = @{ title = 'Scouter'; id = 'scouts' } }
                    }
                }
            }

            return @{ data = @{ keys = @('member-1') } }
        }

        $lists, $null = Get-SNSMaillistInfo -CredentialCustomlists $credential `
            -MaillistsIds @(
                @{ scoutnet_list_id = 'list-1' }
                @{ scoutnet_list_id = 'list-1' }
            )

        @($lists['list-1'].scouter) | Should -Be @('member-1')
        Should -Invoke Get-SNSApiGroupCustomlist -Times 2 -Exactly
    }
}

Describe 'Get-RandomPassword' {
    It 'returns the requested number of alphanumeric characters' {
        $password = Get-RandomPassword -length 24

        $password.Length | Should -Be 24
        $password | Should -Match '^[a-zA-Z0-9]+$'
    }
}
