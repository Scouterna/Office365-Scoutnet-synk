BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Office365-Scoutnet-synk'

    . (Join-Path $modulePath 'SNSConfiguration.ps1')
    . (Join-Path $modulePath 'Write-SNSLog.ps1')
    . (Join-Path $modulePath 'Get-SNSApiGroupMemberlist.ps1')
    . (Join-Path $modulePath 'Send-SNSMandatoryMomentReport.ps1')
    . (Join-Path $modulePath 'Send-SNSMandatoryMomentReminder.ps1')
    . (Join-Path $modulePath 'Invoke-SNSMandatoryMomentReportAndReminder.ps1')
}

Describe 'Mandatory moment reminder' {
    BeforeEach {
        $script:SNSConf = [SNSConfiguration]::new()
        $script:SNSConf.EmailFromAddress = 'reports@example.org'
        $script:SNSConf.MandatoryMomentReminderDaysBefore = 30
        Mock Write-SNSLog {}
        Mock Send-MgUserMail { $script:sent += , $BodyParameter }
        $script:sent = @()

        $role = @{ value = @{ group = @{ g = @{ r = @{ role_name = 'Kårledare' } } }; troop = @{} } }
        $script:memberlist = @{
            data = @{
                '1' = @{
                    member_no = @{ value = '1001' }
                    first_name = @{ value = 'Ada' }
                    last_name = @{ value = 'Lovelace' }
                    email = @{ value = 'ada@example.org' }
                    date_of_birth = @{ value = '1990-01-01' }
                    pc_course_89 = @{ value = '2022-02-01' }
                    latest_police_check = @{ value = '2024-01-01' }
                    roles = $role
                }
                '2' = @{
                    member_no = @{ value = '1002' }
                    first_name = @{ value = 'Grace' }
                    last_name = @{ value = 'Hopper' }
                    email = @{ value = ' ' }
                    contact_alt_email = @{ value = 'grace@example.org' }
                    date_of_birth = @{ value = '1990-01-01' }
                    pc_course_89 = @{ value = '' }
                    latest_police_check = @{ value = '2024-01-01' }
                    roles = $role
                }
                '3' = @{
                    member_no = @{ value = '1003' }
                    first_name = @{ value = 'No' }
                    last_name = @{ value = 'Mail' }
                    email = @{ value = '' }
                    date_of_birth = @{ value = '1990-01-01' }
                    pc_course_89 = @{ value = '2020-01-01' }
                    latest_police_check = @{ value = '2024-01-01' }
                    roles = $role
                }
                '4' = @{
                    member_no = @{ value = '1004' }
                    first_name = @{ value = 'Up' }
                    last_name = @{ value = 'ToDate' }
                    email = @{ value = 'ok@example.org' }
                    date_of_birth = @{ value = '1990-01-01' }
                    pc_course_89 = @{ value = '2024-06-01' }
                    latest_police_check = @{ value = '2024-06-01' }
                    roles = $role
                }
                '5' = @{
                    member_no = @{ value = '1005' }
                    first_name = @{ value = 'Too' }
                    last_name = @{ value = 'Young' }
                    email = @{ value = 'young@example.org' }
                    date_of_birth = @{ value = '2015-01-01' }
                    pc_course_89 = @{ value = '' }
                    latest_police_check = @{ value = '' }
                    roles = $role
                }
                '6' = @{
                    member_no = @{ value = '1006' }
                    first_name = @{ value = 'No' }
                    last_name = @{ value = 'Role' }
                    email = @{ value = 'norole@example.org' }
                    date_of_birth = @{ value = '1990-01-01' }
                    pc_course_89 = @{ value = '' }
                    latest_police_check = @{ value = '' }
                    roles = @{ value = @{ group = @{}; troop = @{} } }
                }
            }
        }
        $script:asOf = [datetime]'2025-02-01'
    }

    It 'reminds members whose moments are missing, expired or about to expire' {
        Send-SNSMandatoryMomentReminder -Memberlist $script:memberlist -AsOfDate $script:asOf | Should -BeTrue

        $script:sent.Count | Should -Be 2
        $ada = $script:sent | Where-Object { $_.Message.ToRecipients[0].EmailAddress.Address -eq 'ada@example.org' }
        $grace = $script:sent | Where-Object { $_.Message.ToRecipients[0].EmailAddress.Address -eq 'grace@example.org' }
        $ada.Message.Body.ContentType | Should -Be 'HTML'
        $ada.Message.Body.Content | Should -Match 'Trygga m&#246;ten'
        $ada.Message.Body.Content | Should -Match 'Giltighetstiden gick ut 2025-02-01'
        $ada.Message.Body.Content | Should -Not -Match 'Registerutdrag'
        $grace.Message.Body.Content | Should -Match 'Momentet &#228;r aldrig utf&#246;rt'
        Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
            $Level -eq 'Warn' -and $Message -like "*member '1003'*no valid e-mail*"
        }
        Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter { $Message -eq 'Mandatory moment reminders sent: 2.' }
    }

    It 'tells when the validity is about to end' {
        Send-SNSMandatoryMomentReminder -Memberlist $script:memberlist -AsOfDate ([datetime]'2025-01-15') | Should -BeTrue

        $ada = $script:sent | Where-Object { $_.Message.ToRecipients[0].EmailAddress.Address -eq 'ada@example.org' }
        $ada.Message.Body.Content | Should -Match 'Giltighetstiden g&#229;r ut 2025-02-01'
    }

    It 'starts reminding only when the moment has expired with zero days before' {
        $script:SNSConf.MandatoryMomentReminderDaysBefore = 0

        Send-SNSMandatoryMomentReminder -Memberlist $script:memberlist -AsOfDate ([datetime]'2025-01-31') | Should -BeTrue

        $script:sent.Count | Should -Be 1
        $script:sent[0].Message.ToRecipients[0].EmailAddress.Address | Should -Be 'grace@example.org'
    }

    It 'skips moments for exempt members, per moment type' {
        $script:SNSConf.MandatoryMomentCourseExemptMembers = @('1001', ' 1002 ')

        Send-SNSMandatoryMomentReminder -Memberlist $script:memberlist -AsOfDate $script:asOf | Should -BeTrue

        $script:sent.Count | Should -Be 0
    }

    It 'does not exempt a member from the other moment type' {
        $script:memberlist.data['1'].latest_police_check = @{ value = '2020-01-01' }
        $script:SNSConf.MandatoryMomentCourseExemptMembers = @('1001')

        Send-SNSMandatoryMomentReminder -Memberlist $script:memberlist -AsOfDate $script:asOf | Should -BeTrue

        $ada = $script:sent | Where-Object { $_.Message.ToRecipients[0].EmailAddress.Address -eq 'ada@example.org' }
        $ada.Message.Body.Content | Should -Match 'Registerutdrag'
        $ada.Message.Body.Content | Should -Not -Match 'Trygga möten'
    }

    It 'uses the configured texts, links and copy recipients' {
        $script:SNSConf.MandatoryMomentReminderSubject = 'Dags för {FirstName}'
        $script:SNSConf.MandatoryMomentReminderHeading = 'Min rubrik'
        $script:SNSConf.MandatoryMomentReminderIntro = 'Hej {Name}!'
        $script:SNSConf.MandatoryMomentCourseReminderText = 'Gör kursen <nu>'
        $script:SNSConf.MandatoryMomentCourseReminderLinks = @('Läs mer|https://example.org/kurs?a=1&b=2', 'https://example.org/andra')
        $script:SNSConf.MandatoryMomentReminderFooter = 'Hälsningar styrelsen'
        $script:SNSConf.MandatoryMomentReminderCc = @('ansvarig@example.org', ' ')

        Send-SNSMandatoryMomentReminder -Memberlist $script:memberlist -AsOfDate $script:asOf | Should -BeTrue

        $ada = $script:sent | Where-Object { $_.Message.ToRecipients[0].EmailAddress.Address -eq 'ada@example.org' }
        $ada.Message.Subject | Should -Be 'Dags för Ada'
        $ada.Message.CcRecipients.Count | Should -Be 1
        $ada.Message.CcRecipients[0].EmailAddress.Address | Should -Be 'ansvarig@example.org'
        $content = $ada.Message.Body.Content
        $content | Should -Match '<h1>Min rubrik</h1>'
        $content | Should -Match '<p>Hej Ada Lovelace!</p>'
        $content | Should -Match 'G&#246;r kursen &lt;nu&gt;'
        $content | Should -Match '<a href="https://example.org/kurs\?a=1&amp;b=2">L&#228;s mer</a>'
        $content | Should -Match '<a href="https://example.org/andra">https://example.org/andra</a>'
        $content | Should -Match 'H&#228;lsningar styrelsen'
    }

    It 'sends all reminders to the test address without Cc in test mode' {
        $script:SNSConf.MandatoryMomentReminderCc = @('ansvarig@example.org')

        Send-SNSMandatoryMomentReminder -Memberlist $script:memberlist -AsOfDate $script:asOf -TestEmailAddress ' test@example.org ' | Should -BeTrue

        $script:sent.Count | Should -Be 2
        foreach ($mail in $script:sent)
        {
            $mail.Message.ToRecipients.Count | Should -Be 1
            $mail.Message.ToRecipients[0].EmailAddress.Address | Should -Be 'test@example.org'
            $mail.Message.ContainsKey('CcRecipients') | Should -BeFalse
        }
    }

    It 'falls back to the report settings when reminder settings are not configured' {
        $script:SNSConf.MandatoryMomentReportFrequency = 'Yearly'
        { Send-SNSMandatoryMomentReminder -Memberlist $script:memberlist } | Should -Throw '*MandatoryMomentReminderFrequency*'

        $script:SNSConf.MandatoryMomentReportFrequency = 'Monthly'
        $script:SNSConf.MandatoryMomentExcludedRoles = @('Kårledare')
        Send-SNSMandatoryMomentReminder -Memberlist $script:memberlist -AsOfDate $script:asOf | Should -BeTrue
        $script:sent.Count | Should -Be 0

        $script:SNSConf.MandatoryMomentMinimumAge = 40
        $script:SNSConf.MandatoryMomentExcludedRoles = @()
        Send-SNSMandatoryMomentReminder -Memberlist $script:memberlist -AsOfDate $script:asOf | Should -BeTrue
        $script:sent.Count | Should -Be 0
    }

    It 'prefers the reminder-specific settings over the report settings' {
        $script:SNSConf.MandatoryMomentExcludedRoles = @('Kårledare')
        $script:SNSConf.MandatoryMomentReminderExcludedRoles = @('Annan roll')
        $script:SNSConf.MandatoryMomentMinimumAge = 40
        $script:SNSConf.MandatoryMomentReminderMinimumAge = 15
        $script:SNSConf.MandatoryMomentValidityYears = 10
        $script:SNSConf.MandatoryMomentReminderValidityYears = 3
        $script:SNSConf.MandatoryMomentReportFrequency = 'Yearly'
        $script:SNSConf.MandatoryMomentReminderFrequency = 'Monthly'

        Send-SNSMandatoryMomentReminder -Memberlist $script:memberlist -AsOfDate $script:asOf | Should -BeTrue

        $script:sent.Count | Should -Be 2
    }

    It 'rejects links that are not http or https before anything is sent' {
        $script:SNSConf.MandatoryMomentPoliceCheckReminderLinks = @('Klicka|javascript:alert(1)')

        { Send-SNSMandatoryMomentReminder -Memberlist $script:memberlist -AsOfDate $script:asOf } |
            Should -Throw '*reminder link*not valid*'

        Should -Invoke Send-MgUserMail -Times 0 -Exactly
    }

    It 'requires a sender address and valid settings' {
        $script:SNSConf.EmailFromAddress = ''
        { Send-SNSMandatoryMomentReminder -Memberlist $script:memberlist } | Should -Throw '*EmailFromAddress*'

        $script:SNSConf.EmailFromAddress = 'reports@example.org'
        $script:SNSConf.MandatoryMomentReminderFrequency = 'Yearly'
        { Send-SNSMandatoryMomentReminder -Memberlist $script:memberlist } | Should -Throw '*MandatoryMomentReminderFrequency*'

        $script:SNSConf.MandatoryMomentReminderFrequency = 'Monthly'
        $script:SNSConf.MandatoryMomentReminderDaysBefore = -1
        { Send-SNSMandatoryMomentReminder -Memberlist $script:memberlist } | Should -Throw '*MandatoryMomentReminderDaysBefore*'
    }

    It 'skips member retrieval and mail when the reminders are not due' {
        Mock Get-SNSApiGroupMemberlist { throw 'Memberlist must not be requested' }

        $result = Send-SNSMandatoryMomentReminder -LastSentValue '2025-02-01T00:00:00.0000000' `
            -AsOfDate ([datetime]'2025-02-02')

        $result | Should -BeFalse
        Should -Invoke Get-SNSApiGroupMemberlist -Times 0 -Exactly
        Should -Invoke Send-MgUserMail -Times 0 -Exactly
    }

    It 'sends regardless of frequency for an invalid last-sent value and warns' {
        $result = Send-SNSMandatoryMomentReminder -Memberlist $script:memberlist -LastSentValue 'not-a-date' `
            -AsOfDate $script:asOf

        $result | Should -BeTrue
        $script:sent.Count | Should -Be 2
        Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
            $Level -eq 'Warn' -and $Message -like '*MandatoryMomentReminderLastSent*valid date*'
        }
    }

    It 'logs an error for a failed mail and continues with the other members' {
        Mock Send-MgUserMail {
            if ($BodyParameter.Message.ToRecipients[0].EmailAddress.Address -eq 'ada@example.org') { throw 'boom' }
            $script:sent += , $BodyParameter
        }

        Send-SNSMandatoryMomentReminder -Memberlist $script:memberlist -AsOfDate $script:asOf | Should -BeTrue

        $script:sent.Count | Should -Be 1
        Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
            $Level -eq 'Error' -and $Message -like "*member '1001'*boom*"
        }
    }

    It 'applies exemptions to the report as well' {
        $html = ConvertTo-SNSMandatoryMomentReportHtml -Memberlist $script:memberlist `
            -MinimumAge 15 -ValidityYears 3 -CourseExemptMembers @('1002') -AsOfDate $script:asOf

        $html | Should -Not -Match 'Hopper'
        $html | Should -Match 'Mail'
    }
}
