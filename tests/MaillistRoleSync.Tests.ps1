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

    It 'creates a missing mail-enabled security group and leaves other group types alone' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXORecipient { param($Identity, $ErrorAction, $Verbose) }
            function New-DistributionGroup { param($Name, $Type, $Alias, $DisplayName, $PrimarySmtpAddress, $ErrorAction, $Verbose, $Confirm) }
            function Set-DistributionGroup { param($Identity, $RequireSenderAuthenticationEnabled, $ErrorAction, $Verbose, $Confirm) }
            Mock Set-DistributionGroup {}
            Mock Get-EXORecipient {
                if ($Identity -eq 'wrongtype') { return [pscustomobject]@{ RecipientTypeDetails = 'MailUniversalDistributionGroup' } }
                if ($Identity -eq 'existing') { return [pscustomobject]@{ RecipientTypeDetails = 'MailUniversalSecurityGroup' } }
                return $null
            }
            Mock New-DistributionGroup {}

            $configuration = [SNSConfiguration]::new()
            $configuration.CredentialCustomlists = $script:credential
            $configuration.CredentialMemberlist = $script:credential
            $configuration.MailListSettings = @{
                newgroup = @{ source = 'group_roles'; group_type = 'security'; mail_nickname = 'newgroup-alias'; display_name = 'Ny grupp'; primary_smtp_address = 'newgroup@example.org'; allow_external_senders = $true; ledare_synk_option = '-'; roles = @('Ordforande') }
                wrongtype = @{ source = 'group_roles'; group_type = 'security'; ledare_synk_option = '-'; roles = @('Ordforande') }
                existing = @{ source = 'group_roles'; group_type = 'security'; ledare_synk_option = '-'; roles = @('Ordforande') }
                plain = @{ source = 'group_roles'; ledare_synk_option = '-'; roles = @('Ordforande') }
                maillist = @{ source = 'group_roles'; group_type = 'distribution'; display_name = 'Maillista'; mail_nickname = 'maillista'; primary_smtp_address = 'maillista@example.org'; ledare_synk_option = '-'; roles = @('Ordforande') }
                incomplete = @{ source = 'group_roles'; group_type = 'distribution'; display_name = 'Ofullständig'; ledare_synk_option = '-'; roles = @('Ordforande') }
            }

            SNSUpdateExchangeDistributionGroups -ValidationHash 'old' -Configuration $configuration | Out-Null

            Should -Invoke New-DistributionGroup -Times 1 -Exactly -ParameterFilter {
                $Name -eq 'maillist' -and $Type -eq 'Distribution' -and $DisplayName -eq 'Maillista' -and
                $Alias -eq 'maillista' -and $PrimarySmtpAddress -eq 'maillista@example.org'
            }
            Should -Invoke New-DistributionGroup -Times 0 -Exactly -ParameterFilter { $Name -eq 'incomplete' }
            Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
                $Level -eq 'Warn' -and $Message -like '*incomplete*mail_nickname, primary_smtp_address*'
            }
            Should -Invoke Update-DistributionGroupMember -Times 0 -Exactly -ParameterFilter { $Identity -eq 'incomplete' }
            Should -Invoke New-DistributionGroup -Times 1 -Exactly -ParameterFilter {
                $Name -eq 'newgroup' -and $Type -eq 'Security' -and $Alias -eq 'newgroup-alias' -and $DisplayName -eq 'Ny grupp'
            }
            Should -Invoke Set-DistributionGroup -Times 1 -Exactly -ParameterFilter {
                $Identity -eq 'newgroup' -and $RequireSenderAuthenticationEnabled -eq $false
            }
            Should -Invoke Update-DistributionGroupMember -Times 4 -Exactly
            Should -Invoke Update-DistributionGroupMember -Times 0 -Exactly -ParameterFilter { $Identity -eq 'wrongtype' }

            $configuration.MailListSettings = @{ bad = @{ source = 'group_roles'; group_type = 'other'; roles = @('x') } }
            { SNSUpdateExchangeDistributionGroups -ValidationHash 'old' -Configuration $configuration -WhatIf } |
                Should -Throw '*Unsupported group_type*'
        }
    }

    It 'creates and updates subject prefix mail flow rules only when add_prefix is true' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXORecipient { param($Identity, $ErrorAction, $Verbose) }
            function Get-TransportRule { param($Identity, $ErrorAction, $Verbose) }
            function New-TransportRule { param($Name, $AnyOfToCCHeader, $PrependSubject, $ExceptIfSubjectContainsWords, $AuditSeverity, $SenderAddressLocation, $RuleErrorAction, $Mode, $ErrorAction, $Verbose, $Confirm) }
            function Set-TransportRule { param($Identity, $PrependSubject, $ExceptIfSubjectContainsWords, $ErrorAction, $Verbose, $Confirm) }
            Mock Get-EXORecipient { return [pscustomobject]@{ PrimarySmtpAddress = 'lookup@example.org' } }
            Mock Get-TransportRule {
                if ($Identity -eq 'Subject prefix for old@example.org') {
                    return [pscustomobject]@{ PrependSubject = '[Gammal] '; ExceptIfSubjectContainsWords = @('[Gammal]') }
                }
                if ($Identity -eq 'Subject prefix for same@example.org') {
                    return [pscustomobject]@{ PrependSubject = '[Samma] '; ExceptIfSubjectContainsWords = @('[Samma]') }
                }
                return $null
            }
            Mock New-TransportRule {}
            Mock Set-TransportRule {}

            $configuration = [SNSConfiguration]::new()
            $configuration.CredentialCustomlists = $script:credential
            $configuration.CredentialMemberlist = $script:credential
            $base = @{ source = 'group_roles'; ledare_synk_option = '-'; roles = @('Ordforande') }
            $configuration.MailListSettings = @{
                'derived@example.org' = $base + @{ add_prefix = $true; display_name = 'Krypen' }
                'text@example.org' = $base + @{ add_prefix = $true; display_name = 'Utmanarledare'; prefix_text = ' [Utm] ' }
                'old@example.org' = $base + @{ add_prefix = $true; prefix_text = '[Ny]' }
                'same@example.org' = $base + @{ add_prefix = $true; prefix_text = '[Samma]' }
                'Alias namn' = $base + @{ add_prefix = $true; display_name = 'Alias' }
                'off@example.org' = $base + @{ add_prefix = $false; display_name = 'Av' }
                'unset@example.org' = $base + @{ display_name = 'Osatt' }
                'noprefix@example.org' = $base + @{ add_prefix = $true }
            }

            SNSUpdateExchangeDistributionGroups -ValidationHash 'old' -Configuration $configuration | Out-Null

            Should -Invoke New-TransportRule -Times 3 -Exactly
            Should -Invoke New-TransportRule -Times 1 -Exactly -ParameterFilter {
                $Name -eq 'Subject prefix for derived@example.org' -and $AnyOfToCCHeader -eq 'derived@example.org' -and
                $PrependSubject -eq '[Krypen] ' -and $ExceptIfSubjectContainsWords -eq '[Krypen]' -and
                $AuditSeverity -eq 'DoNotAudit' -and $SenderAddressLocation -eq 'Header' -and $Mode -eq 'Enforce'
            }
            Should -Invoke New-TransportRule -Times 1 -Exactly -ParameterFilter {
                $AnyOfToCCHeader -eq 'text@example.org' -and $PrependSubject -eq '[Utm] ' -and $ExceptIfSubjectContainsWords -eq '[Utm]'
            }
            Should -Invoke New-TransportRule -Times 1 -Exactly -ParameterFilter {
                $AnyOfToCCHeader -eq 'lookup@example.org' -and $PrependSubject -eq '[Alias] '
            }
            Should -Invoke Set-TransportRule -Times 1 -Exactly -ParameterFilter {
                $Identity -eq 'Subject prefix for old@example.org' -and $PrependSubject -eq '[Ny] ' -and $ExceptIfSubjectContainsWords -eq '[Ny]'
            }
            Should -Invoke New-TransportRule -Times 0 -Exactly -ParameterFilter { $AnyOfToCCHeader -in @('off@example.org', 'unset@example.org', 'same@example.org') }
            Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
                $Level -eq 'Warn' -and $Message -like '*noprefix@example.org*neither prefix_text nor display_name*'
            }

            Should -Invoke New-TransportRule -Times 3 -Exactly
            SNSUpdateExchangeDistributionGroups -ValidationHash 'old' -Configuration $configuration -WhatIf | Out-Null
            Should -Invoke New-TransportRule -Times 3 -Exactly
            Should -Invoke Set-TransportRule -Times 1 -Exactly
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
