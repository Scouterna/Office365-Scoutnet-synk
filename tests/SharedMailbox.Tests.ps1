BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Office365-Scoutnet-synk\Office365-Scoutnet-synk.psm1'
    Import-Module $modulePath -Force -ErrorAction Stop

    # Stubs for the Exchange Online commands, which only exist in a connected session.
    $script:exchangeStubs = @(
        'Get-EXOMailbox', 'Get-EXORecipient', 'New-Mailbox', 'Set-Mailbox', 'Get-MailboxPermission',
        'Add-MailboxPermission', 'Remove-MailboxPermission', 'Get-RecipientPermission',
        'Add-RecipientPermission', 'Remove-RecipientPermission'
    )
    function global:Get-EXOMailbox { [CmdletBinding()] param($RecipientTypeDetails, $Properties, $Identity) }
    function global:Get-EXORecipient { [CmdletBinding()] param($Filter) }
    function global:New-Mailbox { [CmdletBinding(SupportsShouldProcess)] param([switch]$Shared, $Name, $DisplayName, $Alias, $PrimarySmtpAddress) }
    function global:Set-Mailbox { [CmdletBinding(SupportsShouldProcess)] param($Identity, $HiddenFromAddressListsEnabled, $MessageCopyForSentAsEnabled) }
    function global:Get-MailboxPermission { [CmdletBinding()] param($Identity) }
    function global:Add-MailboxPermission { [CmdletBinding(SupportsShouldProcess)] param($Identity, $User, $AccessRights, $InheritanceType, $AutoMapping) }
    function global:Remove-MailboxPermission { [CmdletBinding(SupportsShouldProcess)] param($Identity, $User, $AccessRights, $InheritanceType) }
    function global:Get-RecipientPermission { [CmdletBinding()] param($Identity) }
    function global:Add-RecipientPermission { [CmdletBinding(SupportsShouldProcess)] param($Identity, $Trustee, $AccessRights) }
    function global:Remove-RecipientPermission { [CmdletBinding(SupportsShouldProcess)] param($Identity, $Trustee, $AccessRights) }
}

AfterAll {
    foreach ($stub in $script:exchangeStubs)
    {
        Remove-Item -Path "function:global:$stub" -ErrorAction SilentlyContinue
    }
    Remove-Module Office365-Scoutnet-synk -ErrorAction SilentlyContinue
}

Describe 'Invoke-SNExchangeRetry' {
    It 'retries a not found error and then succeeds' {
        InModuleScope Office365-Scoutnet-synk {
            Mock Start-Sleep {}
            $script:calls = 0
            $result = Invoke-SNExchangeRetry -Action {
                $script:calls++
                if ($script:calls -lt 3) { throw "The operation couldn't be performed because object 'x' couldn't be found." }
                'ok'
            }
            $result | Should -Be 'ok'
            $script:calls | Should -Be 3
            Should -Invoke Start-Sleep -Times 2 -Exactly
        }
    }

    It 'does not retry other errors' {
        InModuleScope Office365-Scoutnet-synk {
            Mock Start-Sleep {}
            $script:calls = 0
            { Invoke-SNExchangeRetry -Action { $script:calls++; throw 'Access denied' } } |
                Should -Throw 'Access denied'
            $script:calls | Should -Be 1
            Should -Invoke Start-Sleep -Times 0 -Exactly
        }
    }

    It 'gives up after the maximum number of attempts' {
        InModuleScope Office365-Scoutnet-synk {
            Mock Start-Sleep {}
            $script:calls = 0
            { Invoke-SNExchangeRetry -MaxAttempts 2 -Action { $script:calls++; throw "Object couldn't be found" } } |
                Should -Throw
            $script:calls | Should -Be 2
        }
    }
}

Describe 'SharedMailboxSettings configuration' {
    It 'is accepted by New-SNSConfiguration' {
        InModuleScope Office365-Scoutnet-synk {
            $settings = @{
                'Kassoren' = @{ source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor') }
            }
            $configuration = New-SNSConfiguration -LicenseAssignment @{} -SharedMailboxSettings $settings
            $configuration.SharedMailboxSettings['Kassoren'].mail_nickname | Should -Be 'kassoren'
        }
    }

    It 'allows the Exchange commands the function uses' {
        InModuleScope Office365-Scoutnet-synk {
            $configuration = New-SNSConfiguration -LicenseAssignment @{}
            $commands = $configuration.commandNames -split ','
            foreach ($command in @('New-Mailbox', 'Set-Mailbox', 'Get-EXORecipient', 'Get-EXOMailbox',
                    'Get-MailboxPermission', 'Add-MailboxPermission', 'Remove-MailboxPermission',
                    'Get-RecipientPermission', 'Add-RecipientPermission', 'Remove-RecipientPermission'))
            {
                $commands | Should -Contain $command
            }
        }
    }

    It 'exports Update-SNSharedMailbox from the module' {
        (Get-Command -Module Office365-Scoutnet-synk -Name Update-SNSharedMailbox) | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-SNSharedMailboxPermissionPlan' {
    BeforeAll {
        InModuleScope Office365-Scoutnet-synk {
            $script:planUsers = @(
                foreach ($number in 1..4)
                {
                    [pscustomobject]@{
                        ExternalDirectoryObjectId = "id-$number"
                        UserPrincipalName = "u$number@example.se"
                        PrimarySmtpAddress = "user$number@example.se"
                    }
                }
            )
            $script:planKnown = @{}
            $script:planById = @{}
            foreach ($user in $script:planUsers)
            {
                $script:planById[$user.ExternalDirectoryObjectId] = $user
                $script:planKnown[$user.UserPrincipalName] = $user
                $script:planKnown[$user.PrimarySmtpAddress] = $user
            }
        }
    }

    It 'adds missing accounts, removes stale accounts and keeps protected accounts' {
        InModuleScope Office365-Scoutnet-synk {
            $plan = Get-SNSharedMailboxPermissionPlan `
                -ExistingNames @('u2@example.se', 'u3@example.se', 'u4@example.se') `
                -DesiredUserIds @{ 'id-1' = $true; 'id-2' = $true } `
                -ProtectedUserIds @{ 'id-4' = $true } `
                -KnownUsers $script:planKnown -UsersById $script:planById

            @($plan.Add).UserPrincipalName | Should -Be @('u1@example.se')
            @($plan.Remove).UserPrincipalName | Should -Be @('u3@example.se')
            @($plan.Kept).UserPrincipalName | Should -Be @('u4@example.se')
        }
    }

    It 'matches existing identities by primary SMTP address regardless of case' {
        InModuleScope Office365-Scoutnet-synk {
            $plan = Get-SNSharedMailboxPermissionPlan `
                -ExistingNames @('USER1@Example.se') `
                -DesiredUserIds @{ 'id-1' = $true } -ProtectedUserIds @{} `
                -KnownUsers $script:planKnown -UsersById $script:planById

            @($plan.Add).Count | Should -Be 0
            @($plan.Remove).Count | Should -Be 0
        }
    }

    It 'ignores identities that are not Microsoft 365 accounts' {
        InModuleScope Office365-Scoutnet-synk {
            $plan = Get-SNSharedMailboxPermissionPlan `
                -ExistingNames @('NT AUTHORITY\SELF', 'Styrelsen', '', 'S-1-5-21-1') `
                -DesiredUserIds @{} -ProtectedUserIds @{} `
                -KnownUsers $script:planKnown -UsersById $script:planById

            @($plan.Add).Count | Should -Be 0
            @($plan.Remove).Count | Should -Be 0
            @($plan.Kept).Count | Should -Be 0
        }
    }
}

Describe 'Update-SNSharedMailbox' {
    BeforeEach {
        InModuleScope Office365-Scoutnet-synk {
            function script:New-TestConfiguration
            {
                $configuration = [SNSConfiguration]::new()
                $configuration.DomainName = 'example.se'
                $configuration.CredentialMemberlist = [pscredential]::new(
                    'test',
                    (ConvertTo-SecureString 'test' -AsPlainText -Force)
                )
                $configuration.CredentialCustomlists = $configuration.CredentialMemberlist
                return $configuration
            }
            $script:users = @(
                foreach ($number in 1..3)
                {
                    [pscustomobject]@{
                        ExternalDirectoryObjectId = "id-$number"
                        UserPrincipalName = "u$number@example.se"
                        PrimarySmtpAddress = "u$number@example.se"
                        CustomAttribute1 = "member-$number"
                    }
                }
            )
            # member-1, member-2 and member-4 have the role. member-4 has no account.
            $script:memberlist = @{
                data = @{
                    '1' = @{ member_no = @{ value = 'member-1' }; roles = @{ value = @{ group = @{ '1' = @{ role_name = 'Kassor' } }; troop = @{} } } }
                    '2' = @{ member_no = @{ value = 'member-2' }; roles = @{ value = @{ group = @{ '1' = @{ role_name = 'Kassor' } }; troop = @{} } } }
                    '3' = @{ member_no = @{ value = 'member-3' }; roles = @{ value = @{ group = @{ '1' = @{ role_name = 'Annan' } }; troop = @{} } } }
                    '4' = @{ member_no = @{ value = 'member-4' }; roles = @{ value = @{ group = @{ '1' = @{ role_name = 'Kassor' } }; troop = @{} } } }
                }
            }
            $script:recipients = @()
            $script:currentMailbox = [pscustomobject]@{ HiddenFromAddressListsEnabled = $false; MessageCopyForSentAsEnabled = $false }
            $script:fullAccess = @()
            $script:sendAs = @()
            $script:sharedMailbox = [pscustomobject]@{
                RecipientTypeDetails = 'SharedMailbox'
                ExternalDirectoryObjectId = 'mbx-1'
                PrimarySmtpAddress = 'kassoren@example.se'
            }

            Mock Get-SNSApiGroupMemberlist { return $script:memberlist }
            Mock Get-EXOMailbox {
                if ($PSBoundParameters.ContainsKey('Identity')) { return $script:currentMailbox }
                return $script:users
            }
            Mock Get-EXORecipient { return $script:recipients }
            Mock New-Mailbox { return $script:sharedMailbox }
            Mock Set-Mailbox {}
            Mock Get-MailboxPermission { return $script:fullAccess }
            Mock Add-MailboxPermission {}
            Mock Remove-MailboxPermission {}
            Mock Get-RecipientPermission { return $script:sendAs }
            Mock Add-RecipientPermission {}
            Mock Remove-RecipientPermission {}
            Mock Write-SNSLog {}
            Mock Start-Sleep {}
        }
    }

    It 'does nothing when no shared mailboxes are configured' {
        InModuleScope Office365-Scoutnet-synk {
            $configuration = New-TestConfiguration

            Update-SNSharedMailbox -Configuration $configuration

            Should -Invoke Get-SNSApiGroupMemberlist -Times 0 -Exactly
            Should -Invoke Get-EXOMailbox -Times 0 -Exactly
        }
    }

    It 'throws without a configuration' {
        InModuleScope Office365-Scoutnet-synk {
            $previous = $Script:SNSConf
            try
            {
                $Script:SNSConf = $null
                { Update-SNSharedMailbox } | Should -Throw 'No configuration specified*'
            }
            finally
            {
                $Script:SNSConf = $previous
            }
        }
    }

    It 'rejects <Name> before anything is fetched' -ForEach @(
        @{ Name = 'a missing mail_nickname'; Rule = @{ source = 'group_roles'; roles = @('Kassor') }; Message = '*must configure mail_nickname*' }
        @{ Name = 'an invalid mail_nickname'; Rule = @{ source = 'group_roles'; mail_nickname = 'a b'; roles = @('Kassor') }; Message = '*invalid mail_nickname*' }
        @{ Name = 'an unsupported source'; Rule = @{ source = 'other'; mail_nickname = 'kassoren' }; Message = '*unsupported source*' }
        @{ Name = 'missing roles'; Rule = @{ source = 'group_roles'; mail_nickname = 'kassoren' }; Message = '*must configure one or more non-empty roles*' }
        @{ Name = 'a mail_list without id'; Rule = @{ source = 'mail_list'; mail_nickname = 'kassoren' }; Message = '*must configure scoutnet_list_id*' }
        @{ Name = 'a non boolean flag'; Rule = @{ source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor'); send_as = 'yes' }; Message = '*invalid send_as*' }
        @{ Name = 'no permission at all'; Rule = @{ source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor'); send_as = $false; full_access = $false }; Message = '*at least one permission*' }
    ) {
        InModuleScope Office365-Scoutnet-synk -Parameters @{ Rule = $Rule; Message = $Message } {
            $configuration = New-TestConfiguration
            $configuration.DomainName = 'example.se'
            $configuration.SharedMailboxSettings = @{ 'Kassoren' = $Rule }

            { Update-SNSharedMailbox -Configuration $configuration } | Should -Throw $Message

            Should -Invoke Get-SNSApiGroupMemberlist -Times 0 -Exactly
            Should -Invoke Get-EXOMailbox -Times 0 -Exactly
            Should -Invoke New-Mailbox -Times 0 -Exactly
        }
    }

    It 'rejects mailboxes that share a mail_nickname' {
        InModuleScope Office365-Scoutnet-synk {
            $configuration = New-TestConfiguration
            $configuration.DomainName = 'example.se'
            $configuration.SharedMailboxSettings = @{
                'A' = @{ source = 'group_roles'; mail_nickname = 'same'; roles = @('Kassor') }
                'B' = @{ source = 'group_roles'; mail_nickname = 'SAME'; roles = @('Kassor') }
            }

            { Update-SNSharedMailbox -Configuration $configuration } | Should -Throw '*use the same mail_nickname*'
            Should -Invoke Get-EXOMailbox -Times 0 -Exactly
        }
    }

    It 'requires email_address or DomainName' {
        InModuleScope Office365-Scoutnet-synk {
            $configuration = New-TestConfiguration
            $configuration.DomainName = ''
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{ source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor') }
            }

            { Update-SNSharedMailbox -Configuration $configuration } | Should -Throw '*email_address*DomainName*'
        }
    }

    It 'adds keep_members permissions without fetching Scoutnet data when source is omitted' {
        InModuleScope Office365-Scoutnet-synk {
            $script:recipients = @($script:sharedMailbox)
            $configuration = New-TestConfiguration
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{
                    mail_nickname = 'kassoren'
                    keep_members = @('u1@example.se', 'id-2')
                }
            }

            Update-SNSharedMailbox -Configuration $configuration

            Should -Invoke Get-SNSApiGroupMemberlist -Times 0 -Exactly
            foreach ($upn in @('u1@example.se', 'u2@example.se'))
            {
                Should -Invoke Add-MailboxPermission -Times 1 -Exactly -ParameterFilter {
                    $Identity -eq 'mbx-1' -and $User -eq $upn -and $AccessRights -eq 'FullAccess'
                }
                Should -Invoke Add-RecipientPermission -Times 1 -Exactly -ParameterFilter {
                    $Identity -eq 'mbx-1' -and $Trustee -eq $upn -and $AccessRights -eq 'SendAs'
                }
            }
            Should -Invoke Add-MailboxPermission -Times 2 -Exactly
            Should -Invoke Add-RecipientPermission -Times 2 -Exactly
        }
    }

    It 'requires keep_members when source is omitted' {
        InModuleScope Office365-Scoutnet-synk {
            $configuration = New-TestConfiguration
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{ mail_nickname = 'kassoren' }
            }

            { Update-SNSharedMailbox -Configuration $configuration } |
                Should -Throw '*must configure keep_members when source is omitted*'
            Should -Invoke Get-SNSApiGroupMemberlist -Times 0 -Exactly
            Should -Invoke Get-EXOMailbox -Times 0 -Exactly
        }
    }

    It 'creates a missing shared mailbox and gives matching accounts Full Access and Send As' {
        InModuleScope Office365-Scoutnet-synk {
            $configuration = New-TestConfiguration
            $configuration.DomainName = 'example.se'
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{ source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor') }
            }

            Update-SNSharedMailbox -Configuration $configuration

            Should -Invoke New-Mailbox -Times 1 -Exactly -ParameterFilter {
                $Shared -and $Name -eq 'Kassoren' -and $DisplayName -eq 'Kassoren' -and
                $Alias -eq 'kassoren' -and $PrimarySmtpAddress -eq 'kassoren@example.se'
            }
            foreach ($upn in @('u1@example.se', 'u2@example.se'))
            {
                Should -Invoke Add-MailboxPermission -Times 1 -Exactly -ParameterFilter {
                    $Identity -eq 'mbx-1' -and $User -eq $upn -and $AccessRights -eq 'FullAccess' -and $AutoMapping -eq $true
                }
                Should -Invoke Add-RecipientPermission -Times 1 -Exactly -ParameterFilter {
                    $Identity -eq 'mbx-1' -and $Trustee -eq $upn -and $AccessRights -eq 'SendAs'
                }
            }
            Should -Invoke Add-MailboxPermission -Times 2 -Exactly
            Should -Invoke Add-RecipientPermission -Times 2 -Exactly
            Should -Invoke Remove-MailboxPermission -Times 0 -Exactly
            Should -Invoke Remove-RecipientPermission -Times 0 -Exactly
            Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter { $Message -like "*member-4*no Microsoft 365 user account*" }
        }
    }

    It 'uses email_address and auto_mapping from the rule' {
        InModuleScope Office365-Scoutnet-synk {
            $configuration = New-TestConfiguration
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{
                    source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor')
                    email_address = 'kassa@example.org'; auto_mapping = $false
                }
            }

            Update-SNSharedMailbox -Configuration $configuration

            Should -Invoke New-Mailbox -Times 1 -Exactly -ParameterFilter { $PrimarySmtpAddress -eq 'kassa@example.org' }
            Should -Invoke Add-MailboxPermission -Times 2 -Exactly -ParameterFilter { $AutoMapping -eq $false }
        }
    }

    It 'adds missing permissions and removes stale ones from an existing shared mailbox' {
        InModuleScope Office365-Scoutnet-synk {
            $script:recipients = @($script:sharedMailbox)
            $script:fullAccess = @(
                [pscustomobject]@{ User = 'u1@example.se'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }
                [pscustomobject]@{ User = 'u3@example.se'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }
                [pscustomobject]@{ User = 'NT AUTHORITY\SELF'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }
                [pscustomobject]@{ User = 'Styrelsen'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }
                [pscustomobject]@{ User = 'u2@example.se'; AccessRights = @('FullAccess'); IsInherited = $true; Deny = $false }
            )
            $script:sendAs = @(
                [pscustomobject]@{ Trustee = 'u2@example.se'; AccessRights = @('SendAs'); IsInherited = $false; AccessControlType = 'Allow' }
                [pscustomobject]@{ Trustee = 'u3@example.se'; AccessRights = @('SendAs'); IsInherited = $false; AccessControlType = 'Allow' }
                [pscustomobject]@{ Trustee = 'NT AUTHORITY\SELF'; AccessRights = @('SendAs'); IsInherited = $false; AccessControlType = 'Allow' }
            )
            $configuration = New-TestConfiguration
            $configuration.DomainName = 'example.se'
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{ source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor') }
            }

            Update-SNSharedMailbox -Configuration $configuration

            Should -Invoke New-Mailbox -Times 0 -Exactly
            # u2 only has an inherited Full Access entry, which does not count.
            Should -Invoke Add-MailboxPermission -Times 1 -Exactly -ParameterFilter { $User -eq 'u2@example.se' }
            Should -Invoke Add-MailboxPermission -Times 1 -Exactly
            Should -Invoke Remove-MailboxPermission -Times 1 -Exactly -ParameterFilter { $User -eq 'u3@example.se' -and $AccessRights -eq 'FullAccess' }
            Should -Invoke Add-RecipientPermission -Times 1 -Exactly -ParameterFilter { $Trustee -eq 'u1@example.se' }
            Should -Invoke Add-RecipientPermission -Times 1 -Exactly
            Should -Invoke Remove-RecipientPermission -Times 1 -Exactly -ParameterFilter { $Trustee -eq 'u3@example.se' -and $AccessRights -eq 'SendAs' }
        }
    }

    It 'keeps accounts listed in keep_members by user principal name or object id' {
        InModuleScope Office365-Scoutnet-synk {
            $script:recipients = @($script:sharedMailbox)
            $script:fullAccess = @(
                [pscustomobject]@{ User = 'u3@example.se'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }
            )
            $script:sendAs = @(
                [pscustomobject]@{ Trustee = 'u3@example.se'; AccessRights = @('SendAs'); IsInherited = $false; AccessControlType = 'Allow' }
            )
            $configuration = New-TestConfiguration
            $configuration.DomainName = 'example.se'
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{
                    source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor')
                    keep_members = @('U3@example.se')
                }
            }

            Update-SNSharedMailbox -Configuration $configuration
            Should -Invoke Remove-MailboxPermission -Times 0 -Exactly
            Should -Invoke Remove-RecipientPermission -Times 0 -Exactly

            $configuration.SharedMailboxSettings['Kassoren'].keep_members = @('id-3')
            Update-SNSharedMailbox -Configuration $configuration
            Should -Invoke Remove-MailboxPermission -Times 0 -Exactly
            Should -Invoke Remove-RecipientPermission -Times 0 -Exactly
        }
    }

    It 'does not manage Full Access when full_access is false' {
        InModuleScope Office365-Scoutnet-synk {
            $script:recipients = @($script:sharedMailbox)
            $script:fullAccess = @(
                [pscustomobject]@{ User = 'u3@example.se'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }
            )
            $configuration = New-TestConfiguration
            $configuration.DomainName = 'example.se'
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{ source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor'); full_access = $false }
            }

            Update-SNSharedMailbox -Configuration $configuration

            Should -Invoke Get-MailboxPermission -Times 0 -Exactly
            Should -Invoke Add-MailboxPermission -Times 0 -Exactly
            Should -Invoke Remove-MailboxPermission -Times 0 -Exactly
            Should -Invoke Add-RecipientPermission -Times 2 -Exactly
        }
    }

    It 'does not manage Send As when send_as is false' {
        InModuleScope Office365-Scoutnet-synk {
            $script:recipients = @($script:sharedMailbox)
            $configuration = New-TestConfiguration
            $configuration.DomainName = 'example.se'
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{ source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor'); send_as = $false }
            }

            Update-SNSharedMailbox -Configuration $configuration

            Should -Invoke Get-RecipientPermission -Times 0 -Exactly
            Should -Invoke Add-RecipientPermission -Times 0 -Exactly
            Should -Invoke Add-MailboxPermission -Times 2 -Exactly
        }
    }

    It 'selects members from a Scoutnet mail list' {
        InModuleScope Office365-Scoutnet-synk {
            $script:recipients = @($script:sharedMailbox)
            Mock Get-SNSMaillistInfo {
                return @{ '0001' = @{ ledare = @('member-3'); scouter = @('member-1') } }, 'hash'
            }
            $configuration = New-TestConfiguration
            $configuration.DomainName = 'example.se'
            $configuration.SharedMailboxSettings = @{
                'Info' = @{ source = 'mail_list'; mail_nickname = 'info'; scoutnet_list_id = '0001'; send_as = $false }
            }

            Update-SNSharedMailbox -Configuration $configuration

            Should -Invoke Add-MailboxPermission -Times 1 -Exactly -ParameterFilter { $User -eq 'u3@example.se' }
            Should -Invoke Add-MailboxPermission -Times 1 -Exactly
        }
    }

    It 'throws when the alias belongs to another kind of recipient' {
        InModuleScope Office365-Scoutnet-synk {
            $script:recipients = @([pscustomobject]@{ RecipientTypeDetails = 'UserMailbox'; ExternalDirectoryObjectId = 'x'; PrimarySmtpAddress = 'kassoren@example.se' })
            $configuration = New-TestConfiguration
            $configuration.DomainName = 'example.se'
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{ source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor') }
            }

            { Update-SNSharedMailbox -Configuration $configuration } | Should -Throw '*not a shared mailbox*'

            Should -Invoke New-Mailbox -Times 0 -Exactly
            Should -Invoke Add-MailboxPermission -Times 0 -Exactly
            Should -Invoke Add-RecipientPermission -Times 0 -Exactly
        }
    }

    It 'throws when several recipients have the alias' {
        InModuleScope Office365-Scoutnet-synk {
            $script:recipients = @($script:sharedMailbox, $script:sharedMailbox)
            $configuration = New-TestConfiguration
            $configuration.DomainName = 'example.se'
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{ source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor') }
            }

            { Update-SNSharedMailbox -Configuration $configuration } | Should -Throw '*More than one recipient*'
        }
    }

    It 'makes no changes when the permissions already match' {
        InModuleScope Office365-Scoutnet-synk {
            $script:recipients = @($script:sharedMailbox)
            $script:fullAccess = @(
                foreach ($upn in @('u1@example.se', 'u2@example.se'))
                {
                    [pscustomobject]@{ User = $upn; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }
                }
            )
            $script:sendAs = @(
                foreach ($upn in @('u1@example.se', 'u2@example.se'))
                {
                    [pscustomobject]@{ Trustee = $upn; AccessRights = @('SendAs'); IsInherited = $false; AccessControlType = 'Allow' }
                }
            )
            $configuration = New-TestConfiguration
            $configuration.DomainName = 'example.se'
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{ source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor') }
            }

            Update-SNSharedMailbox -Configuration $configuration

            Should -Invoke New-Mailbox -Times 0 -Exactly
            Should -Invoke Set-Mailbox -Times 0 -Exactly
            Should -Invoke Add-MailboxPermission -Times 0 -Exactly
            Should -Invoke Remove-MailboxPermission -Times 0 -Exactly
            Should -Invoke Add-RecipientPermission -Times 0 -Exactly
            Should -Invoke Remove-RecipientPermission -Times 0 -Exactly
        }
    }

    It 'makes no changes with -WhatIf' {
        InModuleScope Office365-Scoutnet-synk {
            $script:recipients = @($script:sharedMailbox)
            $script:fullAccess = @(
                [pscustomobject]@{ User = 'u3@example.se'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }
            )
            $configuration = New-TestConfiguration
            $configuration.DomainName = 'example.se'
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{ source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor'); hidden_from_address_list = $true }
            }

            Update-SNSharedMailbox -Configuration $configuration -WhatIf

            Should -Invoke Set-Mailbox -Times 0 -Exactly
            Should -Invoke Add-MailboxPermission -Times 0 -Exactly
            Should -Invoke Remove-MailboxPermission -Times 0 -Exactly
            Should -Invoke Add-RecipientPermission -Times 0 -Exactly
        }
    }

    It 'does not create a mailbox or change permissions with -WhatIf when the mailbox is missing' {
        InModuleScope Office365-Scoutnet-synk {
            $configuration = New-TestConfiguration
            $configuration.DomainName = 'example.se'
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{ source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor') }
            }

            Update-SNSharedMailbox -Configuration $configuration -WhatIf

            Should -Invoke New-Mailbox -Times 0 -Exactly
            Should -Invoke Add-MailboxPermission -Times 0 -Exactly
        }
    }

    It 'changes hidden_from_address_list and copy_sent_items only when they differ' {
        InModuleScope Office365-Scoutnet-synk {
            $script:recipients = @($script:sharedMailbox)
            $script:currentMailbox = [pscustomobject]@{ HiddenFromAddressListsEnabled = $true; MessageCopyForSentAsEnabled = $false }
            $configuration = New-TestConfiguration
            $configuration.DomainName = 'example.se'
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{
                    source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor')
                    hidden_from_address_list = $true; copy_sent_items = $true
                }
            }

            Update-SNSharedMailbox -Configuration $configuration

            Should -Invoke Set-Mailbox -Times 1 -Exactly -ParameterFilter {
                $Identity -eq 'mbx-1' -and $MessageCopyForSentAsEnabled -eq $true -and
                -not $PSBoundParameters.ContainsKey('HiddenFromAddressListsEnabled')
            }
        }
    }

    It 'applies configured settings to a new mailbox' {
        InModuleScope Office365-Scoutnet-synk {
            $configuration = New-TestConfiguration
            $configuration.DomainName = 'example.se'
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{
                    source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor')
                    hidden_from_address_list = $true; copy_sent_items = $true
                }
            }

            Update-SNSharedMailbox -Configuration $configuration

            Should -Invoke Set-Mailbox -Times 1 -Exactly -ParameterFilter {
                $HiddenFromAddressListsEnabled -eq $true -and $MessageCopyForSentAsEnabled -eq $true
            }
        }
    }

    It 'retries permission lookups for a new mailbox that is not visible yet' {
        InModuleScope Office365-Scoutnet-synk {
            $script:attempts = 0
            Mock Get-MailboxPermission {
                $script:attempts++
                if ($script:attempts -lt 3) { throw "The operation couldn't be performed because object 'kassoren' couldn't be found." }
                return @()
            }
            $configuration = New-TestConfiguration
            $configuration.DomainName = 'example.se'
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{ source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor'); send_as = $false }
            }

            Update-SNSharedMailbox -Configuration $configuration

            $script:attempts | Should -Be 3
            Should -Invoke Add-MailboxPermission -Times 2 -Exactly
        }
    }

    It 'does not retry permission lookups for an existing mailbox' {
        InModuleScope Office365-Scoutnet-synk {
            $script:recipients = @($script:sharedMailbox)
            Mock Get-MailboxPermission { throw "The operation couldn't be performed because object 'kassoren' couldn't be found." }
            $configuration = New-TestConfiguration
            $configuration.DomainName = 'example.se'
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{ source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor'); send_as = $false }
            }

            { Update-SNSharedMailbox -Configuration $configuration } | Should -Throw "*couldn't be found*"

            Should -Invoke Get-MailboxPermission -Times 1 -Exactly
            Should -Invoke Start-Sleep -Times 0 -Exactly
        }
    }

    It 'throws when the member list is missing and changes nothing' {
        InModuleScope Office365-Scoutnet-synk {
            Mock Get-SNSApiGroupMemberlist { return $null }
            $configuration = New-TestConfiguration
            $configuration.DomainName = 'example.se'
            $configuration.SharedMailboxSettings = @{
                'Kassoren' = @{ source = 'group_roles'; mail_nickname = 'kassoren'; roles = @('Kassor') }
            }

            { Update-SNSharedMailbox -Configuration $configuration } | Should -Throw '*No shared mailboxes were changed*'

            Should -Invoke New-Mailbox -Times 0 -Exactly
        }
    }
}
