BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Office365-Scoutnet-synk\Office365-Scoutnet-synk.psm1'
    Import-Module $modulePath -Force -ErrorAction Stop
}

AfterAll {
    Remove-Module Office365-Scoutnet-synk -ErrorAction SilentlyContinue
}

Describe 'SNSUpdateExchangeDistributionGroups role based rules' {
    BeforeEach {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXOMailbox {}
            function Update-DistributionGroupMember {}
            function Remove-MailContact {}

            $script:credential = [pscredential]::new('test', (ConvertTo-SecureString 'test' -AsPlainText -Force))
            $script:memberlist = @{
                data = @{
                    '1' = @{
                        member_no = @{ value = 1 }
                        unit = @{ raw_value = '10'; value = 'Aventyrarna' }
                        roles = @{ value = @{
                            troop = @{ '10' = @{ '100' = @{ role_name = 'Avdelningsledare' } } }
                            group = @{ '5' = @{ '200' = @{ role_name = 'Ordforande' } } }
                        } }
                    }
                    '2' = @{
                        member_no = @{ value = 2 }
                        unit = @{ raw_value = '10'; value = 'Aventyrarna' }
                        roles = @{ value = @{ troop = @{ '10' = @{ '101' = @{ role_name = 'Assisterande' } } } } }
                    }
                    '3' = @{
                        member_no = @{ value = 3 }
                        unit = @{ raw_value = '11'; value = 'Utmanarna' }
                        roles = @{ value = @{ troop = @{ '11' = @{ '102' = @{ role_name = 'Avdelningsledare' } } } } }
                    }
                    '4' = @{
                        member_no = @{ value = 4 }
                        unit = @{ raw_value = '10'; value = 'Aventyrarna' }
                        roles = @{ value = @{} }
                    }
                }
            }
            $script:addresses = @{}
            foreach ($id in 1..4)
            {
                $script:addresses[$id] = @{
                    first_name = "F$id"; last_name = "L$id"
                    primary_email = "user$id@example.org"
                    mailaddresses = @("user$id@example.org"); contacts_addresses = @()
                }
            }

            Mock Write-SNSLog {}
            Mock Get-SNSMaillistInfo { return @{}, 'list-hash' }
            Mock Get-SNSUserEmail { return $script:addresses, 'member-hash' }
            Mock Get-SNSApiGroupMemberlist { return $script:memberlist }
            Mock Get-SNSExchangeMailListMember { return @{}, @{} }
            Mock Get-EXOMailbox { return @() }
            Mock New-SNSMailContact { return $true }
            Mock Update-DistributionGroupMember {}
            Mock Remove-MailContact {}
        }
    }

    It 'builds lists from troop roles, troop id fallback, wildcard and group roles' {
        InModuleScope Office365-Scoutnet-synk {
            $configuration = [SNSConfiguration]::new()
            $configuration.CredentialCustomlists = $script:credential
            $configuration.CredentialMemberlist = $script:credential
            $configuration.MailListSettings = @{
                leaders = @{
                    source = 'troop_roles'
                    ledare_synk_option = '-'
                    troops = @(@{ name = 'Aventyrarna'; roles = @('Avdelningsledare') })
                }
                byid = @{
                    source = 'troop_roles'
                    ledare_synk_option = '-'
                    troops = @(@{ name = 'Wrong name'; id = '11'; roles = '*' })
                }
                anyrole = @{
                    source = 'troop_roles'
                    ledare_synk_option = '-'
                    troops = @(@{ name = 'Aventyrarna'; roles = '*' })
                }
                board = @{
                    source = 'group_roles'
                    ledare_synk_option = '-'
                    roles = @('Ordforande')
                }
            }

            $hash, $mailData = SNSUpdateExchangeDistributionGroups `
                -ValidationHash 'old' -Configuration $configuration -ReturnMaildata -WhatIf

            @($mailData.leaders) | Should -Be @('user1@example.org')
            @($mailData.byid) | Should -Be @('user3@example.org')
            @($mailData.anyrole | Sort-Object) | Should -Be @('user1@example.org', 'user2@example.org')
            @($mailData.board) | Should -Be @('user1@example.org')
            Should -Invoke Get-SNSMaillistInfo -Times 1 -Exactly -ParameterFilter { @($MaillistsIds).Count -eq 0 }
        }
    }

    It 'combines unit membership with troop roles' {
        InModuleScope Office365-Scoutnet-synk {
            $configuration = [SNSConfiguration]::new()
            $configuration.CredentialCustomlists = $script:credential
            $configuration.CredentialMemberlist = $script:credential
            $configuration.MailListSettings = @{
                everyone = @{
                    source = 'troop_roles'
                    scouter_synk_option = 'p'
                    ledare_synk_option = '-'
                    troops = @(@{ name = 'Aventyrarna'; roles = @('Avdelningsledare'); unit_members = $true })
                }
                byid = @{
                    source = 'troop_roles'
                    scouter_synk_option = 'p'
                    ledare_synk_option = '-'
                    troops = @(@{ name = 'Wrong name'; id = '11'; unit_members = $true })
                }
            }

            $hash, $mailData = SNSUpdateExchangeDistributionGroups `
                -ValidationHash 'old' -Configuration $configuration -ReturnMaildata -WhatIf

            @($mailData.everyone | Sort-Object) | Should -Be @('user1@example.org', 'user2@example.org', 'user4@example.org')
            @($mailData.byid) | Should -Be @('user3@example.org')
        }
    }

    It 'changes the validation hash when roles change and rejects invalid rules' {
        InModuleScope Office365-Scoutnet-synk {
            $configuration = [SNSConfiguration]::new()
            $configuration.CredentialCustomlists = $script:credential
            $configuration.CredentialMemberlist = $script:credential
            $configuration.MailListSettings = @{
                board = @{ source = 'group_roles'; ledare_synk_option = '-'; roles = @('Ordforande') }
            }
            $hash1 = SNSUpdateExchangeDistributionGroups -ValidationHash 'old' -Configuration $configuration -WhatIf
            $script:memberlist.data['2'].roles.value['group'] = @{ '5' = @{ '201' = @{ role_name = 'Ordforande' } } }
            $hash2 = SNSUpdateExchangeDistributionGroups -ValidationHash 'old' -Configuration $configuration -WhatIf
            $hash2 | Should -Not -Be $hash1

            $configuration.MailListSettings = @{ bad = @{ source = 'group_roles' } }
            { SNSUpdateExchangeDistributionGroups -ValidationHash 'old' -Configuration $configuration -WhatIf } |
                Should -Throw '*non-empty roles*'
        }
    }
}
