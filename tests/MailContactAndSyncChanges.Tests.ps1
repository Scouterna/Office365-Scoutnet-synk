BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Office365-Scoutnet-synk\Office365-Scoutnet-synk.psm1'
    Import-Module $modulePath -Force -ErrorAction Stop
}

AfterAll {
    Remove-Module Office365-Scoutnet-synk -ErrorAction SilentlyContinue
}

Describe 'New-SNSMailContact' {
    It 'creates a contact, sets its company, and hides it from address lists' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXORecipient { param($Identity, $ErrorAction, $Verbose) }
            function New-MailContact { param($Name, $ExternalEmailAddress, $ErrorAction, $Verbose) }
            function Get-Contact { param($Identity, $ErrorAction, $Verbose) }
            function Set-Contact { param($Identity, $Company, $ErrorAction, $Verbose) }
            function Set-MailContact { param($Identity, $HiddenFromAddressListsEnabled, $ErrorAction, $Verbose) }
            function Start-Sleep { param($Seconds) }
            function Write-SNSLog { param($Message, $Level) }

            Mock Get-EXORecipient { return $null }
            Mock New-MailContact {}
            Mock Get-Contact { return [pscustomobject]@{ Identity = $Identity } }
            Mock Start-Sleep {}
            Mock Set-Contact {}
            Mock Set-MailContact {}
            Mock Write-SNSLog {}

            $result = New-SNSMailContact -Epost 'ada@example.org' -DisplayName 'Ada Lovelace'

            $result | Should -BeTrue
            Should -Invoke New-MailContact -Times 1 -Exactly -ParameterFilter {
                $Name -eq 'ada@example.org' -and $ExternalEmailAddress -eq 'ada@example.org'
            }
            Should -Invoke Get-Contact -Times 1 -Exactly
            Should -Invoke Set-Contact -Times 1 -Exactly -ParameterFilter {
                $Identity -eq 'ada@example.org' -and $Company -eq 'Ada Lovelace'
            }
            Should -Invoke Set-MailContact -Times 1 -Exactly -ParameterFilter {
                $Identity -eq 'ada@example.org' -and $HiddenFromAddressListsEnabled
            }
        }
    }

    It 'updates an existing contact with a missing company without creating it again' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXORecipient { param($Identity, $ErrorAction, $Verbose) }
            function New-MailContact { param($Name, $ExternalEmailAddress, $ErrorAction, $Verbose) }
            function Get-Contact { param($Identity, $ErrorAction, $Verbose) }
            function Set-Contact { param($Identity, $Company, $ErrorAction, $Verbose) }
            function Set-MailContact { param($Identity, $HiddenFromAddressListsEnabled, $ErrorAction, $Verbose) }
            function Write-SNSLog { param($Message, $Level) }

            Mock Get-EXORecipient { return [pscustomobject]@{ Identity = $Identity } }
            Mock New-MailContact {}
            Mock Get-Contact { return [pscustomobject]@{ Company = '' } }
            Mock Set-Contact {}
            Mock Set-MailContact {}
            Mock Write-SNSLog {}

            $result = New-SNSMailContact -Epost 'ada@example.org' -DisplayName 'Ada Lovelace'

            $result | Should -BeTrue
            Should -Invoke New-MailContact -Times 0 -Exactly
            Should -Invoke Set-Contact -Times 1 -Exactly -ParameterFilter {
                $Identity -eq 'ada@example.org' -and $Company -eq 'Ada Lovelace'
            }
            Should -Invoke Set-MailContact -Times 1 -Exactly -ParameterFilter {
                $Identity -eq 'ada@example.org' -and $HiddenFromAddressListsEnabled
            }
        }
    }

    It 'leaves an existing contact with a company unchanged' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXORecipient { param($Identity, $ErrorAction, $Verbose) }
            function Get-Contact { param($Identity, $ErrorAction, $Verbose) }
            function Set-Contact { param($Identity, $Company, $ErrorAction, $Verbose) }
            function Set-MailContact { param($Identity, $HiddenFromAddressListsEnabled, $ErrorAction, $Verbose) }
            function Write-SNSLog { param($Message, $Level) }

            Mock Get-EXORecipient { return [pscustomobject]@{ Identity = $Identity } }
            Mock Get-Contact { return [pscustomobject]@{ Company = 'Existing company' } }
            Mock Set-Contact {}
            Mock Set-MailContact {}
            Mock Write-SNSLog {}

            $result = New-SNSMailContact -Epost 'ada@example.org' -DisplayName 'Ada Lovelace'

            $result | Should -BeTrue
            Should -Invoke Get-Contact -Times 1 -Exactly
            Should -Invoke Set-Contact -Times 0 -Exactly
            Should -Invoke Set-MailContact -Times 0 -Exactly
        }
    }

    It 'returns false and logs a warning when contact creation fails' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXORecipient { param($Identity, $ErrorAction, $Verbose) }
            function New-MailContact { param($Name, $ExternalEmailAddress, $ErrorAction, $Verbose) }
            function Start-Sleep { param($Seconds) }
            function Get-Contact { param($Identity, $ErrorAction, $Verbose) }
            function Set-Contact { param($Identity, $Company, $ErrorAction, $Verbose) }
            function Set-MailContact { param($Identity, $HiddenFromAddressListsEnabled, $ErrorAction, $Verbose) }
            function Write-SNSLog { param($Message, $Level) }

            Mock Get-EXORecipient { return $null }
            Mock New-MailContact { throw 'Exchange rejected contact' }
            Mock Start-Sleep {}
            Mock Get-Contact { return $null }
            Mock Set-Contact {}
            Mock Set-MailContact {}
            Mock Write-SNSLog {}

            $result = New-SNSMailContact -Epost 'ada@example.org' -DisplayName 'Ada Lovelace'

            $result | Should -BeFalse
            Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
                $Level -eq 'Warn' -and $Message -like '*Exchange rejected contact*'
            }
            Should -Invoke Set-Contact -Times 0 -Exactly
        }
    }

    It 'does not change Exchange when WhatIf is specified' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXORecipient { param($Identity, $ErrorAction, $Verbose) }
            function New-MailContact { param($Name, $ExternalEmailAddress, $ErrorAction, $Verbose) }
            function Start-Sleep { param($Seconds) }
            function Get-Contact { param($Identity, $ErrorAction, $Verbose) }
            function Set-Contact { param($Identity, $Company, $ErrorAction, $Verbose) }
            function Set-MailContact { param($Identity, $HiddenFromAddressListsEnabled, $ErrorAction, $Verbose) }
            function Write-SNSLog { param($Message, $Level) }

            Mock Get-EXORecipient { return $null }
            Mock New-MailContact {}
            Mock Start-Sleep {}
            Mock Get-Contact {}
            Mock Set-Contact {}
            Mock Set-MailContact {}
            Mock Write-SNSLog {}

            $result = New-SNSMailContact -Epost 'ada@example.org' -DisplayName 'Ada Lovelace' -WhatIf

            $result | Should -BeTrue
            Should -Invoke Get-EXORecipient -Times 1 -Exactly
            Should -Invoke New-MailContact -Times 0 -Exactly
            Should -Invoke Set-Contact -Times 0 -Exactly
            Should -Invoke Set-MailContact -Times 0 -Exactly
        }
    }
}

Describe 'SNSUpdateExchangeDistributionGroups changes' {
    It 'updates distribution members and removes obsolete contacts when synchronization runs' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXOMailbox {
                param($RecipientTypeDetails, $Properties, $Verbose)
            }
            function Update-DistributionGroupMember {
                param($Identity, $Members, $ErrorAction, $Verbose, $Confirm)
            }
            function Remove-MailContact {
                param($Identity, $Confirm, $Verbose)
            }

            $configuration = [SNSConfiguration]::new()
            $credential = [pscredential]::new('test', [System.Security.SecureString]::new())
            $configuration.CredentialCustomlists = $credential
            $configuration.CredentialMemberlist = $credential
            $configuration.MailListSettings = @{
                scouts = @{
                    scoutnet_list_id    = 'list-1'
                    scouter_synk_option = 'p'
                    ledare_synk_option  = '-'
                    email_addresses     = @('extra@example.org')
                }
            }
            $memberData = @{
                scout  = @{
                    first_name         = 'Ada'
                    last_name          = 'Lovelace'
                    primary_email      = 'ada@example.org'
                    mailaddresses      = @('ada@example.org')
                    contacts_addresses = @()
                }
                leader = @{
                    first_name         = 'Grace'
                    last_name          = 'Hopper'
                    primary_email      = 'grace@example.org'
                    mailaddresses      = @('grace@example.org')
                    contacts_addresses = @()
                }
            }
            $customLists = @{
                'list-1' = @{
                    scouter = @('scout')
                    ledare  = @('leader')
                }
            }
            $obsoleteContact = [pscustomobject]@{
                Identity      = 'obsolete@example.org'
                RecipientType = 'MailContact'
            }

            Mock Write-SNSLog {}
            Mock Get-SNSMaillistInfo { return $customLists, 'new-list-hash' }
            Mock Get-SNSUserEmail { return $memberData, 'new-member-hash' }
            Mock Get-SNSExchangeMailListMember { return @{}, @{ $obsoleteContact.Identity = $obsoleteContact } }
            Mock Get-EXOMailbox { return @() }
            Mock New-SNSMailContact { return $true }
            Mock Update-DistributionGroupMember {}
            Mock Remove-MailContact {}

            $hash, $mailData = SNSUpdateExchangeDistributionGroups `
                -ValidationHash 'old-hash' -Configuration $configuration -ReturnMaildata

            $hash | Should -Match '^0xnew-list-hashnew-member-hash'
            @($mailData.scouts) | Should -Contain 'ada@example.org'
            @($mailData.scouts) | Should -Contain 'grace@example.org'
            @($mailData.scouts) | Should -Contain 'extra@example.org'
            Should -Invoke Update-DistributionGroupMember -Times 1 -Exactly -ParameterFilter {
                $Identity -eq 'scouts' -and
                'ada@example.org' -in $Members -and
                'grace@example.org' -in $Members -and
                'extra@example.org' -in $Members
            }
            Should -Invoke Remove-MailContact -Times 1 -Exactly -ParameterFilter {
                $Identity -eq 'obsolete@example.org'
            }
        }
    }

    It 'logs distribution group update errors and continues synchronization' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXOMailbox {
                param($RecipientTypeDetails, $Properties, $Verbose)
            }
            function Update-DistributionGroupMember {
                param($Identity, $Members, $ErrorAction, $Verbose, $Confirm)
            }
            function Remove-MailContact {
                param($Identity, $Confirm, $Verbose)
            }

            $configuration = [SNSConfiguration]::new()
            $credential = [pscredential]::new('test', [System.Security.SecureString]::new())
            $configuration.CredentialCustomlists = $credential
            $configuration.CredentialMemberlist = $credential
            $configuration.MailListSettings = @{
                scouts = @{
                    scoutnet_list_id    = 'list-1'
                    scouter_synk_option = 'p'
                    ledare_synk_option  = '-'
                    email_addresses     = @()
                }
            }
            $memberData = @{
                scout = @{
                    first_name         = 'Ada'
                    last_name          = 'Lovelace'
                    primary_email      = 'ada@example.org'
                    mailaddresses      = @('ada@example.org')
                    contacts_addresses = @()
                }
            }
            $customLists = @{
                'list-1' = @{
                    scouter = @('scout')
                    ledare  = @()
                }
            }
            $script:loggedMessages = @()

            Mock Write-SNSLog { $script:loggedMessages += $Message }
            Mock Get-SNSMaillistInfo { return $customLists, 'list-hash' }
            Mock Get-SNSUserEmail { return $memberData, 'member-hash' }
            Mock Get-SNSExchangeMailListMember { return @{}, @{} }
            Mock Get-EXOMailbox { return @() }
            Mock New-SNSMailContact { return $true }
            Mock Update-DistributionGroupMember { throw 'group update failed' }
            Mock Remove-MailContact {}

            $hash = SNSUpdateExchangeDistributionGroups `
                -ValidationHash 'old-hash' -Configuration $configuration

            $hash | Should -Match '^0xlist-hashmember-hash'
            $script:loggedMessages | Should -Contain 'Update of distribution group members in scouts failed! Error group update failed'
            Should -Invoke Update-DistributionGroupMember -Times 1 -Exactly
            Remove-Variable loggedMessages -Scope Script
        }
    }

    It 'handles leader address options, empty Office 365 addresses, and contacts shared with other lists' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-EXOMailbox {
                param($RecipientTypeDetails, $Properties, $Verbose)
            }
            function Update-DistributionGroupMember {
                param($Identity, $Members, $ErrorAction, $Verbose, $Confirm)
            }
            function Remove-MailContact {
                param($Identity, $Confirm, $Verbose)
            }

            $configuration = [SNSConfiguration]::new()
            $credential = [pscredential]::new('test', [System.Security.SecureString]::new())
            $configuration.CredentialCustomlists = $credential
            $configuration.CredentialMemberlist = $credential
            $configuration.MailListSettings = @{
                'empty-office' = @{
                    scoutnet_list_id = 'list-empty-office'
                    scouter_synk_option = '@'
                    ledare_synk_option = '-'
                    email_addresses = @()
                }
                'leader-both' = @{
                    scoutnet_list_id = 'list-leader-both'
                    scouter_synk_option = ''
                    ledare_synk_option = '&'
                    email_addresses = @()
                }
                'leader-fallback' = @{
                    scoutnet_list_id = 'list-leader-fallback'
                    scouter_synk_option = ''
                    ledare_synk_option = 't'
                    email_addresses = @()
                }
                'leader-empty' = @{
                    scoutnet_list_id = 'list-leader-empty'
                    scouter_synk_option = ''
                    ledare_synk_option = '-'
                    email_addresses = @()
                }
            }
            $memberData = @{
                'empty-office-member' = @{
                    first_name = 'Office'
                    last_name = 'Empty'
                    primary_email = 'scout@example.org'
                    mailaddresses = @('scout@example.org')
                    contacts_addresses = @()
                }
                'leader-both-member' = @{
                    first_name = 'Leader'
                    last_name = 'Both'
                    primary_email = 'leader-both@example.org'
                    mailaddresses = @('leader-both@example.org')
                    contacts_addresses = @()
                }
                'leader-fallback-member' = @{
                    first_name = 'Leader'
                    last_name = 'Fallback'
                    primary_email = 'leader-fallback@example.org'
                    mailaddresses = @('leader-fallback@example.org')
                    contacts_addresses = @()
                }
                'leader-empty-member' = @{
                    first_name = 'Leader'
                    last_name = 'Empty'
                    primary_email = ' '
                    mailaddresses = @()
                    contacts_addresses = @()
                }
            }
            $customLists = @{
                'list-empty-office' = @{ scouter = @('empty-office-member'); ledare = @() }
                'list-leader-both' = @{ scouter = @(); ledare = @('leader-both-member') }
                'list-leader-fallback' = @{ scouter = @(); ledare = @('leader-fallback-member') }
                'list-leader-empty' = @{ scouter = @(); ledare = @('leader-empty-member', 'unknown-leader') }
            }
            $obsoleteContact = [pscustomobject]@{
                Identity = 'obsolete@example.org'
                RecipientType = 'MailContact'
            }
            $sharedContact = [pscustomobject]@{
                Identity = 'shared@example.org'
                RecipientType = 'MailContact'
            }
            $script:syncMessages = @()

            Mock Write-SNSLog { $script:syncMessages += $Message }
            Mock Get-SNSMaillistInfo { return $customLists, 'list-hash' }
            Mock Get-SNSUserEmail { return $memberData, 'member-hash' }
            Mock Get-SNSExchangeMailListMember {
                $contactsInOtherLists = @{ $sharedContact.Identity = $sharedContact }
                $contactsToProcess = @{
                    $obsoleteContact.Identity = $obsoleteContact
                    $sharedContact.Identity = $sharedContact
                }
                return $contactsInOtherLists, $contactsToProcess
            }
            Mock Get-EXOMailbox {
                return @([pscustomobject]@{ CustomAttribute1 = 'unrelated-user' })
            }
            Mock Add-Office365UserToList {
                if ($Member -eq 'empty-office-member') {
                    return $true, ''
                }
                if ($Member -eq 'leader-both-member') {
                    return $true, 'leader@office.org'
                }
                return $false, $null
            }
            Mock New-SNSMailContact { return $true }
            Mock Update-DistributionGroupMember {}
            Mock Remove-MailContact {}

            $hash, $mailData = SNSUpdateExchangeDistributionGroups `
                -ValidationHash 'old-hash' -Configuration $configuration -ReturnMaildata

            $hash | Should -Match '^0xlist-hashmember-hash'
            @($mailData.'empty-office') | Should -BeNullOrEmpty
            @($mailData.'leader-both') | Should -Contain 'leader@office.org'
            @($mailData.'leader-both') | Should -Contain 'leader-both@example.org'
            @($mailData.'leader-fallback') | Should -Be @('leader-fallback@example.org')
            @($mailData.'leader-empty') | Should -BeNullOrEmpty
            Should -Invoke Add-Office365UserToList -Times 3 -Exactly
            Should -Invoke New-SNSMailContact -Times 2 -Exactly
            Should -Invoke Remove-MailContact -Times 1 -Exactly -ParameterFilter {
                $Identity -eq 'obsolete@example.org'
            }
            Should -Invoke Remove-MailContact -Times 0 -Exactly -ParameterFilter {
                $Identity -eq 'shared@example.org'
            }
            $script:syncMessages | Should -Contain "Primary mailaddres for 'empty-office-member' is empty. Cannot add member to list."
            $script:syncMessages | Should -Contain "Primary mailaddres in scoutnet for 'Leader Empty' is empty. Cannot add member to list."
            $script:syncMessages | Should -Contain "Id 'unknown-leader' from leader-empty not found in list of all members."
            Remove-Variable syncMessages -Scope Script
        }
    }
}

Describe 'Add-Office365UserToList' {
    It 'returns the matched user SMTP address and logs successful addition' {
        InModuleScope Office365-Scoutnet-synk {
            $user = [pscustomobject]@{
                CustomAttribute1 = 'member-123'
                DisplayName = 'Ada Lovelace'
                PrimarySmtpAddress = 'ada@example.org'
            }
            Mock Write-SNSLog {}

            $result, $smtpAddress = Add-Office365UserToList `
                -allOffice365Users @($user) -member 'member-123' `
                -MemberData ([pscustomobject]@{}) -distGroupName 'scouts'

            $result | Should -BeTrue
            $smtpAddress | Should -Be 'ada@example.org'
            Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
                $Message -like '*Ada Lovelace*ada@example.org*scouts*'
            }
        }
    }

    It 'returns false without logging when a user is missing and warnings are disabled' {
        InModuleScope Office365-Scoutnet-synk {
            Mock Write-SNSLog {}

            $result, $smtpAddress = Add-Office365UserToList `
                -allOffice365Users @() -member 'missing-member' `
                -MemberData ([pscustomobject]@{}) -distGroupName 'scouts'

            $result | Should -BeFalse
            $smtpAddress | Should -Be ''
            Should -Invoke Write-SNSLog -Times 0 -Exactly
        }
    }

    It 'logs a warning when a user is missing and warnings are enabled' {
        InModuleScope Office365-Scoutnet-synk {
            Mock Write-SNSLog {}

            $result, $smtpAddress = Add-Office365UserToList `
                -allOffice365Users @() -member 'missing-member' `
                -MemberData ([pscustomobject]@{
                    first_name = 'Grace'
                    last_name = 'Hopper'
                }) -distGroupName 'scouts' -doWarn

            $result | Should -BeFalse
            $smtpAddress | Should -Be ''
            Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
                $Level -eq 'Warn' -and $Message -like '*Grace Hopper*CustomAttribute1*'
            }
        }
    }
}
