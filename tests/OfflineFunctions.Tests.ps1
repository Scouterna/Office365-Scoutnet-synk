BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Office365-Scoutnet-synk'

    . (Join-Path $modulePath 'SNSConfiguration.ps1')
    . (Join-Path $modulePath 'Write-SNSLog.ps1')
    . (Join-Path $modulePath 'Get-SNSStableHash.ps1')
    . (Join-Path $modulePath 'Receive-SNSApiJson.ps1')
    . (Join-Path $modulePath 'Get-SNSApiGroupMemberlist.ps1')
    . (Join-Path $modulePath 'Get-SNSApiGroupCustomlist.ps1')
    . (Join-Path $modulePath 'Get-SNSMaillistInfo.ps1')
    . (Join-Path $modulePath 'Get-SNSUserEmail.ps1')
    . (Join-Path $modulePath 'Invoke-SNSUppdateOffice365User.ps1')
    . (Join-Path $modulePath 'Send-SNSMandatoryMomentReport.ps1')

    $credential = [pscredential]::new('test', (ConvertTo-SecureString 'test' -AsPlainText -Force))
}

Describe 'Get-SNSStableHash' {
    It 'returns the same hash when hashtable and collection ordering differs' {
        $first = @{
            members = @('member-2', 'member-1')
            settings = @{
                enabled = $true
                name = 'Scoutnet'
            }
        }
        $second = @{
            settings = @{
                name = 'Scoutnet'
                enabled = $true
            }
            members = @('member-1', 'member-2')
        }

        $firstHash = Get-SNSStableHash -InputObject $first
        $secondHash = Get-SNSStableHash -InputObject $second

        $firstHash | Should -Be $secondHash
        $firstHash | Should -Match '^[0-9A-F]{64}$'
    }
}

Describe 'New-SNSConfiguration' {
    It 'creates a configuration with defaults and supplied values' {
        $configuration = New-SNSConfiguration -LicenseAssignment @{} `
            -DomainName 'example.org' -LogFilePath 'test.log' `
            -AllUsersGroupName 'all-users' `
            -JsonFileDirectory 'json-output' -SaveJsonFiles $true

        $configuration.GetType().Name | Should -Be 'SNSConfiguration'
        $configuration.DomainName | Should -Be 'example.org'
        $configuration.LogFilePath | Should -Be 'test.log'
        $configuration.AllUsersGroupName | Should -Be 'all-users'
        $configuration.JsonFileDirectory | Should -Be 'json-output'
        $configuration.SaveJsonFiles | Should -BeTrue
        $configuration.PreferredLanguage | Should -Be 'sv-SE'
        $configuration.UsageLocation | Should -Be 'SE'
        $configuration.MandatoryMomentReportFrequency | Should -Be 'Monthly'
        $configuration.MandatoryMomentExcludedRoles | Should -BeNullOrEmpty
        $configuration.MandatoryMomentCourseDescription |
            Should -Be 'Lista över medlemmar som behöver göra om kursen Trygga möten'
        $configuration.MandatoryMomentPoliceCheckDescription |
            Should -Be 'Lista över medlemmar som behöver förnya registerutdraget'
    }

    It 'stores supplied credentials, mailing lists, account email templates, and settings' {
        $mailListSettings = @{
            scouts = @{
                scoutnet_list_id = 'list-1'
                scouter_synk_option = 'p'
            }
        }
        $configuration = New-SNSConfiguration `
            -CredentialMemberlist $credential `
            -CredentialCustomlists $credential `
            -MailListSettings $mailListSettings `
            -DomainName 'example.org' `
            -LicenseAssignment @{ Basic = @('Teams') } `
            -UserSyncMailListId 'leaders-list' `
            -LogFilePath 'sync.log' `
            -AllUsersGroupName 'all-users' `
            -SignatureText 'Hello <DisplayName>' `
            -SignatureHtml '<p>Hello <DisplayName></p>' `
            -NewUserEmailSubject 'Welcome' `
            -NewUserEmailText 'Use <Password> to sign in as <UserPrincipalName>' `
            -NewUserEmailContentType 'HTML' `
            -NewUserInfoEmailSubject 'More information' `
            -NewUserInfoEmailText 'Welcome <DisplayName>' `
            -NewUserInfoEmailContentType 'HTML' `
            -DisabledAccountsAutoReplyText 'Contact <DisplayName>'

        [object]::ReferenceEquals($configuration.CredentialMemberlist, $credential) | Should -BeTrue
        [object]::ReferenceEquals($configuration.CredentialCustomlists, $credential) | Should -BeTrue
        [object]::ReferenceEquals($configuration.MailListSettings, $mailListSettings) | Should -BeTrue
        $configuration.DomainName | Should -Be 'example.org'
        $configuration.UserSyncMailListId | Should -Be 'leaders-list'
        $configuration.LogFilePath | Should -Be 'sync.log'
        $configuration.AllUsersGroupName | Should -Be 'all-users'
        $configuration.SignatureText | Should -Be 'Hello <DisplayName>'
        $configuration.SignatureHtml | Should -Be '<p>Hello <DisplayName></p>'
        $configuration.NewUserEmailSubject | Should -Be 'Welcome'
        $configuration.NewUserEmailText | Should -Be 'Use <Password> to sign in as <UserPrincipalName>'
        $configuration.NewUserEmailContentType | Should -Be 'HTML'
        $configuration.NewUserInfoEmailSubject | Should -Be 'More information'
        $configuration.NewUserInfoEmailText | Should -Be 'Welcome <DisplayName>'
        $configuration.NewUserInfoEmailContentType | Should -Be 'HTML'
        $configuration.DisabledAccountsAutoReplyText | Should -Be 'Contact <DisplayName>'
    }
}

Describe 'SNSConfiguration accessors' {
    It 'returns the current configuration and replaces it when requested' {
        $configuration = [SNSConfiguration]::new()
        $configuration.DomainName = 'example.org'

        Set-SNSConfiguration -Configuration $configuration

        [object]::ReferenceEquals((Get-SNSConfiguration), $configuration) | Should -BeTrue
        (Get-SNSConfiguration).DomainName | Should -Be 'example.org'
    }

    It 'does not replace the current configuration when WhatIf is specified' {
        $current = [SNSConfiguration]::new()
        $current.DomainName = 'current.example.org'
        $candidate = [SNSConfiguration]::new()
        $candidate.DomainName = 'candidate.example.org'
        Set-SNSConfiguration -Configuration $current

        Set-SNSConfiguration -Configuration $candidate -WhatIf

        [object]::ReferenceEquals((Get-SNSConfiguration), $current) | Should -BeTrue
        (Get-SNSConfiguration).DomainName | Should -Be 'current.example.org'
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
        $script:SNSConf.SaveJsonFiles = $true

        Mock Receive-SNSApiJson {
            $script:receivedMemberlistFileName = $FileName
            $script:receivedMemberlistSaveFile = $SaveFile
            return @{ data = @{ id = 'member-list' } }
        }
    }

    It 'fetches the member list once and reuses the cached response' {
        $first = Get-SNSApiGroupMemberlist -Credential $credential -Uri 'https://example.org/first'
        $second = Get-SNSApiGroupMemberlist -Credential $credential -Uri 'https://example.org/second'

        $first.data.id | Should -Be 'member-list'
        $second.data.id | Should -Be 'member-list'
        $script:receivedMemberlistFileName | Should -Be 'memberlist.json'
        $script:receivedMemberlistSaveFile | Should -BeTrue
        Should -Invoke Receive-SNSApiJson -Times 1 -Exactly
    }
}

Describe 'Mandatory moment report' {
    BeforeEach {
        $script:SNSConf = [SNSConfiguration]::new()
        $script:SNSConf.EmailFromAddress = 'reports@example.org'
        $script:SNSConf.MandatoryMomentReportRecipients = @('responsible@example.org', 'other@example.org')
        $script:SNSConf.MandatoryMomentReportFrequency = 'Monthly'
        $script:SNSConf.MandatoryMomentMinimumAge = 15
        $script:SNSConf.MandatoryMomentValidityYears = 3
        $script:SNSConf.MandatoryMomentCourseDescription = 'Kursbeskrivning för rapporten'
        $script:SNSConf.MandatoryMomentPoliceCheckDescription = 'Registerkontrollbeskrivning för rapporten'
        Mock Write-SNSLog {}
    }

    It 'filters by age, roles, and validity and renders separate filterable tables' {
        $memberlist = @{
            data = @{
                member1 = @{
                    first_name = @{ value = 'Ada' }
                    last_name = @{ value = 'Lovelace' }
                    date_of_birth = @{ value = '2000-01-01' }
                    unit = @{ raw_value = '3'; value = 'Test patrull3' }
                    pc_course_89 = @{ value = '2020-01-01' }
                    latest_police_check = @{ value = '2024-01-01' }
                    roles = @{
                        value = @{
                            group = @{ group1 = @{ role1 = @{ role_name = 'Kårledare' } } }
                            troop = @{ '2' = @{ unit1 = @{ role2 = @{ role_name = 'Avdelningsledare' } } } }
                        }
                    }
                }
                member2 = @{
                    first_name = @{ value = 'Grace' }
                    last_name = @{ value = 'Hopper' }
                    date_of_birth = @{ value = '2000-01-01' }
                    pc_course_89 = @{ value = '' }
                    latest_police_check = @{ value = '2020-01-01' }
                    roles = @{
                        value = @{
                            group = @{}
                            troop = @{ '2' = @{ unit1 = @{ role2 = @{ role_name = 'Ledare' } } } }
                        }
                    }
                }
                member3 = @{
                    first_name = @{ value = 'Too' }
                    last_name = @{ value = 'Young' }
                    date_of_birth = @{ value = '2010-02-01' }
                    pc_course_89 = @{ value = '2020-01-01' }
                    latest_police_check = @{ value = '2020-01-01' }
                    roles = @{ value = @{ group = @{ group1 = @{ role1 = @{ role_name = 'Kårledare' } } }; troop = @{} } }
                }
                member4 = @{
                    first_name = @{ value = 'Boundary' }
                    last_name = @{ value = 'Date' }
                    date_of_birth = @{ value = '1990-01-01' }
                    unit = @{ raw_value = '2'; value = 'Äventyrarna' }
                    pc_course_89 = @{ value = '2022-02-01' }
                    latest_police_check = @{ value = '2022-02-01' }
                    roles = @{ value = @{ group = @{ group1 = @{ role1 = @{ role_name = 'Kårledare' } } }; troop = @{} } }
                }
                member5 = @{
                    first_name = @{ value = 'No' }
                    last_name = @{ value = 'Role' }
                    date_of_birth = @{ value = '1990-01-01' }
                    pc_course_89 = @{ value = '2020-01-01' }
                    latest_police_check = @{ value = '2020-01-01' }
                    roles = @{ value = @{ group = @{}; troop = @{} } }
                }
                member6 = @{
                    first_name = @{ value = 'Excluded' }
                    last_name = @{ value = 'Role' }
                    date_of_birth = @{ value = '1990-01-01' }
                    pc_course_89 = @{ value = '2020-01-01' }
                    latest_police_check = @{ value = '2020-01-01' }
                    roles = @{ value = @{ group = @{}; troop = @{ '2' = @{ unit1 = @{ role1 = @{ role_name = 'Scout' } } } } } }
                }
                member7 = @{
                    first_name = @{ value = 'Mixed' }
                    last_name = @{ value = 'Roles' }
                    date_of_birth = @{ value = '1990-01-01' }
                    pc_course_89 = @{ value = '2020-01-01' }
                    latest_police_check = @{ value = '2020-01-01' }
                    roles = @{
                        value = @{
                            group = @{ group1 = @{ role1 = @{ role_name = 'Kårledare' } } }
                            troop = @{ '2' = @{ unit1 = @{ role2 = @{ role_name = 'Scout' } } } }
                        }
                    }
                }
            }
        }

        $html = ConvertTo-SNSMandatoryMomentReportHtml -Memberlist $memberlist `
            -MinimumAge 15 -ValidityYears 3 `
            -CourseDescription $script:SNSConf.MandatoryMomentCourseDescription `
            -PoliceCheckDescription $script:SNSConf.MandatoryMomentPoliceCheckDescription `
            -ExcludedRoles @('Scout') `
            -AsOfDate ([datetime]'2025-02-01')

        $html | Should -Match '<h2>Trygga möten</h2>'
        $html | Should -Match '<h2>Registerutdrag</h2>'
        $html | Should -Match '<p>Kursbeskrivning f&#246;r rapporten</p>'
        $html | Should -Match '<p>Registerkontrollbeskrivning f&#246;r rapporten</p>'
        $reason = '<p>En medlem &#228;r med i listan om medlemmen &#228;r &#228;ldre &#228;n 15 &#229;r, har minst en k&#229;r- eller avdelningsroll och det var minst 3 &#229;r sedan momentet gjordes.</p>'
        ([regex]::Matches($html, [regex]::Escape($reason))).Count | Should -Be 1
        $html.IndexOf($reason) | Should -BeLessThan $html.IndexOf('<h2>Trygga möten</h2>')
        $html | Should -Not -Match 'Filtrera listan'
        $html | Should -Not -Match '<script|<input|oninput=|filterMomentReport'
        $html | Should -Match 'Efternamn</th><th>Förnamn</th><th>Datum när momentet gjordes</th><th>Roller i kåren</th><th>Roller på avdelning'
        $html | Should -Match '<td>Lovelace</td><td>Ada</td><td>2020-01-01</td><td>K&#229;rledare</td><td>&#196;ventyrarna: Avdelningsledare</td>'
        $html | Should -Match '<td>Hopper</td><td>Grace</td><td>Aldrig utf&#246;rt</td>'
        $html | Should -Match '<td>Date</td><td>Boundary</td><td>2022-02-01</td>'
        $html | Should -Match '<td>Date</td><td>Boundary</td><td>2022-02-01</td><td>K&#229;rledare</td><td></td>'
        $html | Should -Not -Match '2024-01-01'
        $html | Should -Not -Match 'Young'
        $html | Should -Not -Match 'No</td><td>Role'
        $html | Should -Not -Match 'Excluded'
        $html | Should -Match '<td>Roles</td><td>Mixed</td><td>2020-01-01</td><td>K&#229;rledare</td><td>&#196;ventyrarna: Scout</td>'
    }

    It 'sends the HTML report to the configured recipients' {
        $memberlist = @{ data = @{} }
        Mock Get-SNSApiGroupMemberlist { return $memberlist }
        Mock Send-MgUserMail { $script:sentReportParameters = $BodyParameter }

        Send-SNSMandatoryMomentReport -Memberlist $memberlist -AsOfDate ([datetime]'2025-02-01')

        Should -Invoke Send-MgUserMail -Times 1 -Exactly -ParameterFilter {
            $UserId -eq 'reports@example.org' -and
            $BodyParameter.Message.ToRecipients.Count -eq 2 -and
            $BodyParameter.Message.Body.ContentType -eq 'HTML'
        }
        $script:sentReportParameters.Message.Body.Content | Should -Match '<h2>Trygga möten</h2>'
        $script:sentReportParameters.Message.Body.Content | Should -Match '<h2>Registerutdrag</h2>'
    }

    It 'sends only on the configured cadence and only once per period' {
        Test-SNSMandatoryMomentReportDue -Frequency Monthly -AsOfDate ([datetime]'2025-03-01') |
            Should -BeTrue
        Test-SNSMandatoryMomentReportDue -Frequency Monthly -AsOfDate ([datetime]'2025-03-02') |
            Should -BeFalse
        Test-SNSMandatoryMomentReportDue -Frequency Monthly -LastSentDate ([datetime]'2025-03-01') `
            -AsOfDate ([datetime]'2025-03-01') | Should -BeFalse
        Test-SNSMandatoryMomentReportDue -Frequency Weekly -AsOfDate ([datetime]'2025-03-03') |
            Should -BeTrue
        Test-SNSMandatoryMomentReportDue -Frequency Weekly -LastSentDate ([datetime]'2025-03-03') `
            -AsOfDate ([datetime]'2025-03-03') | Should -BeFalse
        Test-SNSMandatoryMomentReportDue -Frequency Weekly -LastSentDate ([datetime]'2025-03-03') `
            -AsOfDate ([datetime]'2025-03-10') | Should -BeTrue
        Test-SNSMandatoryMomentReportDue -Frequency Daily -AsOfDate ([datetime]'2025-03-02') |
            Should -BeTrue
        Test-SNSMandatoryMomentReportDue -Frequency Daily -LastSentDate ([datetime]'2025-03-02T08:00:00') `
            -AsOfDate ([datetime]'2025-03-02') | Should -BeFalse
        Test-SNSMandatoryMomentReportDue -Frequency Daily -LastSentDate ([datetime]'2025-03-02T08:00:00') `
            -AsOfDate ([datetime]'2025-03-03') | Should -BeTrue
    }

    It 'skips member retrieval and mail when the report is not due' {
        $script:SNSConf.MandatoryMomentReportFrequency = 'Monthly'
        Mock Get-SNSApiGroupMemberlist { throw 'Memberlist must not be requested' }
        Mock Send-MgUserMail {}

        $result = Send-SNSMandatoryMomentReport -LastSentValue '2025-03-01T00:00:00.0000000' `
            -AsOfDate ([datetime]'2025-03-02')

        $result | Should -BeFalse
        Should -Invoke Get-SNSApiGroupMemberlist -Times 0 -Exactly
        Should -Invoke Send-MgUserMail -Times 0 -Exactly
    }

    It 'logs a warning and skips the report for an invalid last-sent automation variable value' {
        Mock Get-SNSApiGroupMemberlist { throw 'Memberlist must not be requested' }
        Mock Send-MgUserMail {}

        $result = Send-SNSMandatoryMomentReport -LastSentValue 'not-a-date' `
            -AsOfDate ([datetime]'2025-03-01')

        $result | Should -BeFalse
        Should -Invoke Get-SNSApiGroupMemberlist -Times 0 -Exactly
        Should -Invoke Send-MgUserMail -Times 0 -Exactly
        Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
            $Level -eq 'Warn' -and $Message -like '*MandatoryMomentReportLastSent*valid date*'
        }
    }

    It 'does not fetch members or send when no recipients are configured' {
        $script:SNSConf.MandatoryMomentReportRecipients = @()
        Mock Get-SNSApiGroupMemberlist { throw 'Memberlist must not be requested' }
        Mock Send-MgUserMail {}

        Send-SNSMandatoryMomentReport -AsOfDate ([datetime]'2025-03-01')

        Should -Invoke Get-SNSApiGroupMemberlist -Times 0 -Exactly
        Should -Invoke Send-MgUserMail -Times 0 -Exactly
        Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
            $Message -like '*no recipients are configured*'
        }
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
        $hash | Should -Match '^[0-9A-F]{64}$'
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
        $script:SNSConf = [SNSConfiguration]::new()
        $script:SNSConf.SaveJsonFiles = $true
        Mock Receive-SNSApiJson {
            $script:receivedUri = $Uri
            $script:receivedFileName = $FileName
            $script:receivedSaveFile = $SaveFile
            return @{ success = $true }
        }

        $result = Get-SNSApiGroupCustomlist -Credential $credential `
            -Uri 'https://example.org/customlists' -listid 'list-1' -ruleid 'rule-2'

        $result.success | Should -Be $true
        $script:receivedUri | Should -Be 'https://example.org/customlists?list_id=list-1&rule_id=rule-2'
        $script:receivedFileName | Should -Be 'customlist_list-1_rule-2.json'
        $script:receivedSaveFile | Should -BeTrue
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
        $hash | Should -Match '^[0-9A-F]{64}$'
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
