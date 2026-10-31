BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Office365-Scoutnet-synk\Office365-Scoutnet-synk.psm1'
    Import-Module $modulePath -Force -ErrorAction Stop
}

AfterAll {
    Remove-Module Office365-Scoutnet-synk -ErrorAction SilentlyContinue
}

Describe 'Get-SNSMaillistInfo' {
    It 'does not add null values for rules without members' {
        InModuleScope Office365-Scoutnet-synk {
            Mock Write-SNSLog {}
            Mock Get-SNSApiGroupCustomlist {
                if ([string]::IsNullOrEmpty($listid))
                {
                    return @{
                        '1' = @{
                            title = 'Lista'
                            rules = @{
                                'a' = @{ id = 'a'; title = 'Ledare' }
                                'b' = @{ id = 'b'; title = 'Scouter' }
                                'c' = @{ id = 'c'; title = 'Scouter 2' }
                            }
                        }
                    }
                }
                switch ($ruleid)
                {
                    'a' { return @{ data = $null } }
                    'b' { return @{ data = @() } }
                    'c' { return @{ data = @{ '10' = @{}; '11' = @{} } } }
                }
            }

            $credential = [pscredential]::new('test', (ConvertTo-SecureString 'test' -AsPlainText -Force))
            $lists, $hash = Get-SNSMaillistInfo -CredentialCustomlists $credential `
                -MaillistsIds @(@{ scoutnet_list_id = '1' })

            @($lists['1'].ledare).Count | Should -Be 0
            @($lists['1'].scouter).Count | Should -Be 2
            $lists['1'].scouter | Should -Not -Contain $null
        }
    }
}
