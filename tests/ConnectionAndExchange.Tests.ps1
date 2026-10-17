BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Office365-Scoutnet-synk\Office365-Scoutnet-synk.psm1'
    Import-Module $modulePath -Force -ErrorAction Stop
}

AfterAll {
    Remove-Module Office365-Scoutnet-synk -ErrorAction SilentlyContinue
}

Describe 'Get-SNSExchangeMailListMember' {
    It 'separates mail contacts in synchronized groups from contacts in other groups' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-DistributionGroup {
                param($Verbose, $ErrorAction)
            }
            function Get-DistributionGroupMember {
                param($Verbose, $Identity, $ErrorAction)
            }

            $synchronizedGroup = [pscustomobject]@{
                ExchangeObjectId = 'synchronized-id'
                DisplayName = 'Synchronized list'
            }
            $otherGroup = [pscustomobject]@{
                ExchangeObjectId = 'other-id'
                DisplayName = 'Other list'
            }
            $synchronizedContact = [pscustomobject]@{
                Identity = 'scout@example.org'
                RecipientType = 'MailContact'
            }
            $otherContact = [pscustomobject]@{
                Identity = 'external@example.org'
                RecipientType = 'MailContact'
            }
            Mock Get-DistributionGroup {
                if ($script:distributionGroupCalls -eq 0) {
                    $script:distributionGroupCalls++
                    return $synchronizedGroup
                }
                $script:distributionGroupCalls++
                return @($synchronizedGroup, $otherGroup)
            }
            Mock Get-DistributionGroupMember {
                if ($Identity -eq 'synchronized-id') {
                    return @(
                        $synchronizedContact
                        [pscustomobject]@{
                            Identity = 'mailbox@example.org'
                            RecipientType = 'UserMailbox'
                        }
                    )
                }
                return $otherContact
            }
            Mock Write-SNSLog {}
            $script:distributionGroupCalls = 0

            $otherMembers, $membersToProcess = Get-SNSExchangeMailListMember -Maillists @('Synchronized list')

            $membersToProcess.Keys | Should -Be @('scout@example.org')
            $otherMembers.Keys | Should -Be @('external@example.org')
            Should -Invoke Get-DistributionGroupMember -Times 2 -Exactly
            Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
                $Message -eq 'Done'
            }
        }
    }

    It 'logs and rethrows errors while fetching a requested distribution group' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-DistributionGroup {
                param($Verbose, $ErrorAction)
            }

            Mock Get-DistributionGroup { throw 'requested group unavailable' }
            Mock Write-SNSLog {}

            { Get-SNSExchangeMailListMember -Maillists @('Missing list') } |
                Should -Throw '*requested group unavailable*'

            Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
                $Level -eq 'Error' -and $Message -like "*Missing list*requested group unavailable*"
            }
        }
    }

    It 'logs and rethrows errors while fetching all distribution groups' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-DistributionGroup {
                param($Verbose, $ErrorAction)
            }

            Mock Get-DistributionGroup {
                if ($script:distributionGroupCalls -eq 1) {
                    throw 'all groups unavailable'
                }
                $script:distributionGroupCalls++
                return [pscustomobject]@{ ExchangeObjectId = 'requested-id' }
            }
            Mock Write-SNSLog {}
            $script:distributionGroupCalls = 0

            { Get-SNSExchangeMailListMember -Maillists @('Requested list') } |
                Should -Throw '*all groups unavailable*'

            Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
                $Level -eq 'Error' -and $Message -like '*all distribution groups*all groups unavailable*'
            }
        }
    }

    It 'logs and rethrows errors while fetching group members' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-DistributionGroup {
                param($Verbose, $ErrorAction)
            }
            function Get-DistributionGroupMember {
                param($Verbose, $Identity, $ErrorAction)
            }

            Mock Get-DistributionGroup {
                $script:distributionGroupCalls++
                return [pscustomobject]@{
                    ExchangeObjectId = 'requested-id'
                    DisplayName = 'Requested list'
                }
            }
            Mock Get-DistributionGroupMember { throw 'members unavailable' }
            Mock Write-SNSLog {}
            $script:distributionGroupCalls = 0

            { Get-SNSExchangeMailListMember -Maillists @('Requested list') } |
                Should -Throw '*members unavailable*'

            Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
                $Level -eq 'Error' -and $Message -like '*distribution group members failed*members unavailable*'
            }
        }
    }
}

Describe 'Connect-SnSOffice365' {
    It 'uses device-code authentication and configured scopes without managed identity' {
        InModuleScope Office365-Scoutnet-synk {
            $script:SNSConf = [SNSConfiguration]::new()
            $script:SNSConf.DomainName = 'example.org'
            function Connect-MgGraph {
                param($AccessToken, [switch]$UseDeviceCode, $ContextScope, $Scopes, $ErrorAction)
            }
            function Connect-ExchangeOnline {
                param(
                    [switch]$ManagedIdentity,
                    [switch]$Device,
                    $Organization,
                    $ShowBanner,
                    $CommandName,
                    $Verbose,
                    $ErrorAction
                )
            }
            function Connect-AzAccount {
                param([switch]$Identity, $ErrorAction)
            }

            Mock Connect-MgGraph {}
            Mock Connect-ExchangeOnline {}
            Mock Connect-AzAccount {}
            Mock Write-SNSLog {}

            Connect-SnSOffice365

            Should -Invoke Connect-MgGraph -Times 1 -Exactly -ParameterFilter {
                $UseDeviceCode -and $ContextScope -eq 'Process' -and
                $Scopes.Count -gt 0
            }
            Should -Invoke Connect-ExchangeOnline -Times 1 -Exactly -ParameterFilter {
                $Device -and $Organization -eq $null
            }
            Should -Invoke Connect-AzAccount -Times 0 -Exactly
        }
    }

    It 'uses managed identity and the acquired Graph token when requested' {
        InModuleScope Office365-Scoutnet-synk {
            $script:SNSConf = [SNSConfiguration]::new()
            $script:SNSConf.DomainName = 'example.org'
            function Connect-AzAccount {
                param([switch]$Identity, $ErrorAction)
            }
            function Get-AzAccessToken {
                param($ResourceTypeName, $ErrorAction)
            }
            function Connect-MgGraph {
                param($AccessToken, [switch]$UseDeviceCode, $ContextScope, $Scopes, $ErrorAction)
            }
            function Connect-ExchangeOnline {
                param(
                    [switch]$ManagedIdentity,
                    [switch]$Device,
                    $Organization,
                    $ShowBanner,
                    $CommandName,
                    $Verbose,
                    $ErrorAction
                )
            }

            Mock Connect-AzAccount {}
            Mock Get-AzAccessToken {
                return [pscustomobject]@{ Token = 'graph-token' }
            }
            Mock Connect-MgGraph {}
            Mock Connect-ExchangeOnline {}
            Mock Write-SNSLog {}

            Connect-SnSOffice365 -ManagedIdentity

            Should -Invoke Connect-AzAccount -Times 1 -Exactly -ParameterFilter {
                $Identity
            }
            Should -Invoke Get-AzAccessToken -Times 1 -Exactly -ParameterFilter {
                $ResourceTypeName -eq 'MSGraph'
            }
            Should -Invoke Connect-MgGraph -Times 1 -Exactly -ParameterFilter {
                $AccessToken -eq 'graph-token'
            }
            Should -Invoke Connect-ExchangeOnline -Times 1 -Exactly -ParameterFilter {
                $ManagedIdentity -and $Organization -eq 'example.org'
            }
        }
    }

    It 'throws when no configuration is available' {
        InModuleScope Office365-Scoutnet-synk {
            $script:SNSConf = $null

            { Connect-SnSOffice365 } | Should -Throw '*No configuration specified*'
        }
    }

    It 'logs and rethrows Exchange connection failures' {
        InModuleScope Office365-Scoutnet-synk {
            $script:SNSConf = [SNSConfiguration]::new()
            function Connect-MgGraph {
                param($AccessToken, [switch]$UseDeviceCode, $ContextScope, $Scopes, $ErrorAction)
            }
            function Connect-ExchangeOnline {
                param(
                    [switch]$ManagedIdentity,
                    [switch]$Device,
                    $Organization,
                    $ShowBanner,
                    $CommandName,
                    $Verbose,
                    $ErrorAction
                )
            }

            Mock Connect-MgGraph {}
            Mock Connect-ExchangeOnline { throw 'Exchange unavailable' }
            Mock Write-SNSLog {}

            { Connect-SnSOffice365 } | Should -Throw '*Exchange unavailable*'

            Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
                $Level -eq 'Error' -and $Message -like '*ExchangeOnline*Exchange unavailable*'
            }
        }
    }
}

Describe 'Disconnect-SnSOffice365' {
    It 'continues to disconnect Graph after Exchange disconnection fails' {
        InModuleScope Office365-Scoutnet-synk {
            function script:Disconnect-ExchangeOnline {
                param($ErrorAction, $Confirm, $Verbose)
            }
            function script:Disconnect-MgGraph {
                param($ErrorAction, $Verbose)
            }

            Mock Disconnect-ExchangeOnline { throw 'Exchange disconnect failed' }
            Mock Disconnect-MgGraph {}
            Mock Write-SNSLog {}

            Disconnect-SnSOffice365

            Should -Invoke Disconnect-MgGraph -Times 1 -Exactly
            Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
                $Level -eq 'Error' -and $Message -like '*ExchangeOnline*Exchange disconnect failed*'
            }
        }
    }

    It 'logs and rethrows Graph disconnection failures' {
        InModuleScope Office365-Scoutnet-synk {
            function script:Disconnect-ExchangeOnline {
                param($ErrorAction, $Confirm, $Verbose)
            }
            function script:Disconnect-MgGraph {
                param($ErrorAction, $Verbose)
            }

            Mock Disconnect-ExchangeOnline {}
            Mock Disconnect-MgGraph { throw 'Graph disconnect failed' }
            Mock Write-SNSLog {}

            { Disconnect-SnSOffice365 } | Should -Throw '*Graph disconnect failed*'

            Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter {
                $Level -eq 'Error' -and $Message -like '*Disconnect-MgGraph*Graph disconnect failed*'
            }
        }
    }
}

Describe 'Receive-SNSApiJson' {
    It 'sends Basic authentication and converts the Scoutnet JSON response' {
        InModuleScope Office365-Scoutnet-synk {
            $securePassword = [System.Security.SecureString]::new()
            foreach ($character in 'api-secret'.ToCharArray()) {
                $securePassword.AppendChar($character)
            }
            $securePassword.MakeReadOnly()
            $credential = [pscredential]::new(
                'scout-user',
                $securePassword
            )
            Mock Invoke-WebRequest {
                return '{"data":{"id":"member-list"}}'
            }

            $result = Receive-SNSApiJson -Uri 'https://example.org/api/memberlist' -Credential $credential

            $result.data.id | Should -Be 'member-list'
            Should -Invoke Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
                $Uri -eq 'https://example.org/api/memberlist' -and
                $Headers.Authorization -eq ('Basic ' + [Convert]::ToBase64String(
                    [System.Text.Encoding]::UTF8.GetBytes(
                        'scout-user:' + $credential.GetNetworkCredential().Password
                    )
                )) -and
                $UseBasicParsing -and $ErrorAction -eq 'Stop'
            }
        }
    }

    It 'propagates web request failures' {
        InModuleScope Office365-Scoutnet-synk {
            $securePassword = [System.Security.SecureString]::new()
            foreach ($character in 'api-secret'.ToCharArray()) {
                $securePassword.AppendChar($character)
            }
            $securePassword.MakeReadOnly()
            $credential = [pscredential]::new(
                'scout-user',
                $securePassword
            )
            Mock Invoke-WebRequest { throw 'Scoutnet unavailable' }

            {
                Receive-SNSApiJson -Uri 'https://example.org/api/memberlist' -Credential $credential
            } | Should -Throw '*Scoutnet unavailable*'
        }
    }
}
