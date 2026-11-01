BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Office365-Scoutnet-synk\Office365-Scoutnet-synk.psm1'
    Import-Module $modulePath -Force -ErrorAction Stop
}

AfterAll {
    Remove-Module Office365-Scoutnet-synk -ErrorAction SilentlyContinue
}

Describe 'Get-SNSRunSummary' {
    BeforeEach {
        $script:log = Join-Path $TestDrive 'run.log'
    }

    It 'summarizes counts, warnings and errors' {
        $lines = @(
            '2026-01-01 10:00:00 INFO: Adding Contacts in distribution group a@example.se'
            '2026-01-01 10:00:05 INFO: Members of distribution group a@example.se are unchanged. Update skipped.'
            '2026-01-01 10:00:06 INFO: Adding Contacts in distribution group b@example.se'
            '2026-01-01 10:00:07 INFO: Creating Contact x@example.se for X'
            '2026-01-01 10:00:08 INFO: Adding Contact ''x@example.se'' to ''b@example.se'''
            '2026-01-01 10:00:09 WARNING: Id ''5'' not found in list of all members.'
            '2026-01-01 10:00:10 WARNING: Id ''5'' not found in list of all members.'
            '2026-01-01 10:01:30 ERROR: Update failed'
            'continuation line without timestamp'
        )
        Set-Content -Path $script:log -Value $lines -Encoding UTF8

        $text = Get-SNSRunSummary -LogFilePath $script:log

        $text | Should -Match 'Status:  FEL: körningen har 1 fel'
        $text | Should -Match 'Tid:     1 min 30 s'
        $text | Should -Match 'Behandlade grupper:\s+2'
        $text | Should -Match 'Kontakter skapade:\s+1'
        $text | Should -Match 'Varningar \(2\)'
        $text | Should -Match "not found in list of all members\. \(x2\)"
        $text | Should -Match 'Update failed'
    }

    It 'reports OK without errors or warnings' {
        Set-Content -Path $script:log -Value '2026-01-01 10:00:00 INFO: Done' -Encoding UTF8
        Get-SNSRunSummary -LogFilePath $script:log | Should -Match 'OK: inga fel eller varningar'
    }

    It 'handles a missing log file' {
        Get-SNSRunSummary -LogFilePath (Join-Path $TestDrive 'nope.log') | Should -Match 'loggfilen hittades inte'
    }
}
