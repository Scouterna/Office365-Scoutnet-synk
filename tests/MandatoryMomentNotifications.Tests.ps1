BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Office365-Scoutnet-synk'
    . (Join-Path $modulePath 'SNSConfiguration.ps1')
    . (Join-Path $modulePath 'Write-SNSLog.ps1')
    . (Join-Path $modulePath 'Send-SNSMandatoryMomentReport.ps1')
    . (Join-Path $modulePath 'Send-SNSMandatoryMomentReminder.ps1')
    . (Join-Path $modulePath 'Invoke-SNSMandatoryMomentReportAndReminder.ps1')
}

Describe 'Invoke-SNSMandatoryMomentReportAndReminder' {
    BeforeEach {
        $script:configuration = [SNSConfiguration]::new()
        $script:configuration.MandatoryMomentReportFrequency = 'Monthly'
        $script:configuration.MandatoryMomentReminderFrequency = ''
        Mock Write-SNSLog {}
        Mock Send-SNSMandatoryMomentReport { $true }
        Mock Send-SNSMandatoryMomentReminder { $true }
    }

    It 'runs both operations by default and returns only the report date' {
        $asOfDate = [datetime]'2025-02-01'
        $result = @(Invoke-SNSMandatoryMomentReportAndReminder `
            -Configuration $script:configuration `
            -LastSentValue '2025-01-01T00:00:00' `
            -AsOfDate $asOfDate)

        $result.Count | Should -Be 1
        $result[0] | Should -BeOfType ([datetime])
        $result[0] | Should -Be $asOfDate
        Should -Invoke Send-SNSMandatoryMomentReport -Times 1 -Exactly -ParameterFilter {
            $FrequencyOverride -eq 'Monthly'
        }
        Should -Invoke Send-SNSMandatoryMomentReminder -Times 1 -Exactly -ParameterFilter {
            $FrequencyOverride -eq 'Monthly'
        }
    }

    It 'warns and uses the report frequency when the configured frequencies differ' {
        $script:configuration.MandatoryMomentReminderFrequency = 'Weekly'

        Invoke-SNSMandatoryMomentReportAndReminder -Configuration $script:configuration `
            -LastSentValue '2025-01-01T00:00:00' | Out-Null

        Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
            $Level -eq 'Warn' -and $Message -like '*both will use the report frequency*'
        }
        Should -Invoke Send-SNSMandatoryMomentReminder -Times 1 -Exactly -ParameterFilter {
            $FrequencyOverride -eq 'Monthly'
        }
    }

    It 'can run only the report' {
        Invoke-SNSMandatoryMomentReportAndReminder -Configuration $script:configuration `
            -LastSentValue '2025-01-01T00:00:00' -NoReminder | Out-Null

        Should -Invoke Send-SNSMandatoryMomentReport -Times 1 -Exactly
        Should -Invoke Send-SNSMandatoryMomentReminder -Times 0 -Exactly
    }

    It 'can run only reminders using the reminder frequency' {
        $script:configuration.MandatoryMomentReminderFrequency = 'Weekly'

        Invoke-SNSMandatoryMomentReportAndReminder -Configuration $script:configuration `
            -LastSentValue '2025-01-01T00:00:00' -NoReport | Out-Null

        Should -Invoke Send-SNSMandatoryMomentReport -Times 0 -Exactly
        Should -Invoke Send-SNSMandatoryMomentReminder -Times 1 -Exactly -ParameterFilter {
            $FrequencyOverride -eq 'Weekly'
        }
    }

    It 'returns no date when neither operation is due or both are disabled' {
        Mock Send-SNSMandatoryMomentReport { $false }
        Mock Send-SNSMandatoryMomentReminder { $false }

        @(Invoke-SNSMandatoryMomentReportAndReminder -Configuration $script:configuration `
            -LastSentValue '2025-01-01T00:00:00').Count | Should -Be 0
        @(Invoke-SNSMandatoryMomentReportAndReminder -Configuration $script:configuration `
            -NoReport -NoReminder).Count | Should -Be 0
    }
}
