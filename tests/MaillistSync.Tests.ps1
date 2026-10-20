BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Office365-Scoutnet-synk\Office365-Scoutnet-synk.psm1'
    Import-Module $modulePath -Force -ErrorAction Stop
}

AfterAll {
    Remove-Module Office365-Scoutnet-synk -ErrorAction SilentlyContinue
}

Describe 'SNSUpdateExchangeDistributionGroups' {
    It 'returns computed mail data without updating or removing members when WhatIf is used' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXOMailbox {}
            function Update-DistributionGroupMember {}
            function Remove-MailContact {}

            $configuration = [SNSConfiguration]::new()
            $credential = [pscredential]::new('test', (ConvertTo-SecureString 'test' -AsPlainText -Force))
            $configuration.CredentialCustomlists = $credential
            $configuration.CredentialMemberlist = $credential
            $configuration.MailListSettings = @{
                scouts = @{
                    scoutnet_list_id = 'list-1'
                    scouter_synk_option = 'p'
                    ledare_synk_option = '-'
                    email_addresses = @('extra@example.org')
                }
            }

            $memberData = @{
                'scout-1' = @{
                    first_name = 'Ada'
                    last_name = 'Lovelace'
                    primary_email = 'ada@example.org'
                    mailaddresses = @('ada@example.org')
                    contacts_addresses = @()
                }
                'leader-1' = @{
                    first_name = 'Grace'
                    last_name = 'Hopper'
                    primary_email = 'grace@example.org'
                    mailaddresses = @('grace@example.org')
                    contacts_addresses = @()
                }
            }
            $customLists = @{
                'list-1' = @{
                    scouter = @('scout-1')
                    ledare = @('leader-1')
                }
            }
            $contactsToKeep = @{
                'old@example.org' = [pscustomobject]@{
                    Identity = 'old@example.org'
                    RecipientType = 'MailContact'
                }
            }

            Mock Write-SNSLog {}
            Mock Get-SNSMaillistInfo { return $customLists, 'list-hash' }
            Mock Get-SNSUserEmail { return $memberData, 'member-hash' }
            Mock Get-SNSExchangeMailListMember { return @{}, $contactsToKeep }
            Mock Get-EXOMailbox { return @() }
            Mock New-SNSMailContact { return $true }
            Mock Update-DistributionGroupMember {}
            Mock Remove-MailContact {}

            $validationHash, $mailData = SNSUpdateExchangeDistributionGroups `
                -ValidationHash 'previous-hash' -Configuration $configuration `
                -ReturnMaildata -WhatIf

            $validationHash | Should -Match '^0xlist-hashmember-hash'
            @($mailData.scouts) | Should -Contain 'ada@example.org'
            @($mailData.scouts) | Should -Contain 'grace@example.org'
            @($mailData.scouts) | Should -Contain 'extra@example.org'
            Should -Invoke New-SNSMailContact -Times 3 -Exactly
            Should -Invoke Update-DistributionGroupMember -Times 0 -Exactly
            Should -Invoke Remove-MailContact -Times 0 -Exactly
        }
    }

    It 'skips Exchange synchronization when the validation hash is unchanged' {
        InModuleScope Office365-Scoutnet-synk {
            $credential = [pscredential]::new('test', (ConvertTo-SecureString 'test' -AsPlainText -Force))
            $configuration = [SNSConfiguration]::new()
            $configuration.CredentialCustomlists = $credential
            $configuration.CredentialMemberlist = $credential
            $configuration.MailListSettings = @{
                scouts = @{
                    scoutnet_list_id = 'list-1'
                    scouter_synk_option = 'p'
                    ledare_synk_option = '-'
                    email_addresses = @()
                }
            }

            Mock Write-SNSLog {}
            Mock Get-SNSMaillistInfo { return @{}, 'list-hash' }
            Mock Get-SNSUserEmail { return @{}, 'member-hash' }
            Mock Get-SNSExchangeMailListMember {}

            $mailSettingsHash = Get-SNSStableHash -InputObject $configuration.MailListSettings
            $expectedHash = "0xlist-hashmember-hash$mailSettingsHash"
            $validationHash = SNSUpdateExchangeDistributionGroups `
                -ValidationHash $expectedHash -Configuration $configuration

            $validationHash | Should -Be $expectedHash
            Should -Invoke Get-SNSExchangeMailListMember -Times 0 -Exactly
        }
    }

    It 'omits missing primary and family addresses while retaining alternate addresses' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXOMailbox {}
            function Update-DistributionGroupMember {}
            function Remove-MailContact {}

            $configuration = [SNSConfiguration]::new()
            $credential = [pscredential]::new('test', (ConvertTo-SecureString 'test' -AsPlainText -Force))
            $configuration.CredentialCustomlists = $credential
            $configuration.CredentialMemberlist = $credential
            $configuration.MailListSettings = @{
                primary = @{
                    scoutnet_list_id = 'list-primary'
                    scouter_synk_option = 'p'
                    ledare_synk_option = '-'
                    email_addresses = @()
                }
                alternate = @{
                    scoutnet_list_id = 'list-alternate'
                    scouter_synk_option = 'a'
                    ledare_synk_option = '-'
                    email_addresses = @()
                }
                family = @{
                    scoutnet_list_id = 'list-family'
                    scouter_synk_option = 'f'
                    ledare_synk_option = '-'
                    email_addresses = @()
                }
                default = @{
                    scoutnet_list_id = 'list-default'
                    scouter_synk_option = ''
                    ledare_synk_option = '-'
                    email_addresses = @()
                }
            }
            $memberData = @{
                member = @{
                    first_name = 'Ada'
                    last_name = 'Lovelace'
                    mailaddresses = @('alternate@example.org')
                    contacts_addresses = @()
                    alt_email = 'alternate@example.org'
                }
            }
            $customLists = @{
                'list-primary' = @{ scouter = @('member'); ledare = @() }
                'list-alternate' = @{ scouter = @('member'); ledare = @() }
                'list-family' = @{ scouter = @('member'); ledare = @() }
                'list-default' = @{ scouter = @('member'); ledare = @() }
            }

            Mock Write-SNSLog {}
            Mock Get-SNSMaillistInfo { return $customLists, 'list-hash' }
            Mock Get-SNSUserEmail { return $memberData, 'member-hash' }
            Mock Get-SNSExchangeMailListMember { return @{}, @{} }
            Mock Get-EXOMailbox { return @() }
            Mock New-SNSMailContact { return $true }
            Mock Update-DistributionGroupMember {}
            Mock Remove-MailContact {}

            $null, $mailData = SNSUpdateExchangeDistributionGroups `
                -ValidationHash 'previous-hash' -Configuration $configuration `
                -ReturnMaildata -WhatIf

            @($mailData.primary) | Should -BeNullOrEmpty
            @($mailData.family) | Should -BeNullOrEmpty
            @($mailData.alternate) | Should -Be @('alternate@example.org')
            @($mailData.default) | Should -Be @('alternate@example.org')
            Should -Invoke New-SNSMailContact -Times 2 -Exactly
            Should -Invoke Update-DistributionGroupMember -Times 0 -Exactly
        }
    }

    It 'handles Office 365 address options, ignored and unknown members, failed contacts, and obsolete contacts' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXOMailbox {}
            function Update-DistributionGroupMember {}
            function Remove-MailContact {}

            $configuration = [SNSConfiguration]::new()
            $credential = [pscredential]::new('test', (ConvertTo-SecureString 'test' -AsPlainText -Force))
            $configuration.CredentialCustomlists = $credential
            $configuration.CredentialMemberlist = $credential
            $configuration.MailListSettings = @{
                office = @{
                    scoutnet_list_id = 'list-office'
                    scouter_synk_option = '@'
                    ledare_synk_option = '-'
                    email_addresses = @()
                }
                both = @{
                    scoutnet_list_id = 'list-both'
                    scouter_synk_option = '&'
                    ledare_synk_option = '-'
                    email_addresses = @()
                }
                fallback = @{
                    scoutnet_list_id = 'list-fallback'
                    scouter_synk_option = 't'
                    ledare_synk_option = '-'
                    email_addresses = @()
                }
                ignored = @{
                    scoutnet_list_id = 'list-ignored'
                    scouter_synk_option = ''
                    ledare_synk_option = '-'
                    ignore_user = @('ignored-member')
                    email_addresses = @()
                }
                unknown = @{
                    scoutnet_list_id = 'list-unknown'
                    scouter_synk_option = ''
                    ledare_synk_option = '-'
                    email_addresses = @()
                }
                failed = @{
                    scoutnet_list_id = 'list-failed'
                    scouter_synk_option = 'p'
                    ledare_synk_option = '-'
                    email_addresses = @()
                }
            }

            $memberData = @{
                office = @{
                    first_name = 'Office'
                    last_name = 'Member'
                    primary_email = 'office-scout@example.org'
                    mailaddresses = @('office-scout@example.org')
                    contacts_addresses = @()
                }
                both = @{
                    first_name = 'Both'
                    last_name = 'Member'
                    primary_email = 'both-scout@example.org'
                    mailaddresses = @('both-scout@example.org')
                    contacts_addresses = @()
                }
                fallback = @{
                    first_name = 'Fallback'
                    last_name = 'Member'
                    primary_email = 'fallback@example.org'
                    mailaddresses = @('fallback@example.org')
                    contacts_addresses = @()
                }
                'ignored-member' = @{
                    first_name = 'Ignored'
                    last_name = 'Member'
                    primary_email = 'ignored@example.org'
                    mailaddresses = @('ignored@example.org')
                    contacts_addresses = @()
                }
                failed = @{
                    first_name = 'Failed'
                    last_name = 'Contact'
                    primary_email = 'failed@example.org'
                    mailaddresses = @('failed@example.org')
                    contacts_addresses = @()
                }
            }
            $customLists = @{
                'list-office' = @{ scouter = @('office'); ledare = @() }
                'list-both' = @{ scouter = @('both'); ledare = @() }
                'list-fallback' = @{ scouter = @('fallback'); ledare = @() }
                'list-ignored' = @{ scouter = @('ignored-member'); ledare = @() }
                'list-unknown' = @{ scouter = @('missing-member'); ledare = @() }
                'list-failed' = @{ scouter = @('failed'); ledare = @() }
            }
            $mailContacts = @{
                'obsolete@example.org' = [pscustomobject]@{
                    Identity = 'obsolete@example.org'
                    RecipientType = 'MailContact'
                }
                'shared@example.org' = [pscustomobject]@{
                    Identity = 'shared@example.org'
                    RecipientType = 'MailContact'
                }
                'failed@example.org' = [pscustomobject]@{
                    Identity = 'failed@example.org'
                    RecipientType = 'MailContact'
                }
            }
            $otherListContacts = @{
                'shared@example.org' = [pscustomobject]@{
                    Identity = 'shared@example.org'
                    RecipientType = 'MailContact'
                }
            }

            $global:MailListSyncMessages = @()
            Mock Write-SNSLog { $global:MailListSyncMessages += $Message }
            Mock Get-SNSMaillistInfo { return $customLists, 'list-hash' }
            Mock Get-SNSUserEmail { return $memberData, 'member-hash' }
            Mock Get-SNSExchangeMailListMember { return $otherListContacts, $mailContacts }
            Mock Get-EXOMailbox { return @([pscustomobject]@{ Id = 'unused-user' }) }
            Mock Add-Office365UserToList {
                if ($Member -in @('office', 'both')) {
                    return $true, "$Member@office.org"
                }
                return $false, $null
            }
            Mock New-SNSMailContact {
                return $Epost -ne 'failed@example.org'
            }
            Mock Update-DistributionGroupMember {}
            Mock Remove-MailContact {}

            $null, $mailData = SNSUpdateExchangeDistributionGroups `
                -ValidationHash 'previous-hash' -Configuration $configuration `
                -ReturnMaildata -WhatIf

            @($mailData.office) | Should -Be @('office@office.org')
            @($mailData.both) | Should -Contain 'both@office.org'
            @($mailData.both) | Should -Contain 'both-scout@example.org'
            @($mailData.fallback) | Should -Be @('fallback@example.org')
            @($mailData.ignored) | Should -BeNullOrEmpty
            @($mailData.unknown) | Should -BeNullOrEmpty
            @($mailData.failed) | Should -BeNullOrEmpty
            Should -Invoke Add-Office365UserToList -Times 3 -Exactly
            Should -Invoke New-SNSMailContact -Times 3 -Exactly
            $global:MailListSyncMessages | Should -Contain 'Removing MailContact obsolete@example.org'
            $global:MailListSyncMessages | Should -Not -Contain 'Removing MailContact shared@example.org'
            $global:MailListSyncMessages | Should -Not -Contain 'Removing MailContact failed@example.org'
            Should -Invoke Remove-MailContact -Times 0 -Exactly
            Remove-Variable MailListSyncMessages -Scope Global
        }
    }
}
