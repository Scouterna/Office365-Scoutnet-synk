BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Office365-Scoutnet-synk\Office365-Scoutnet-synk.psm1'
    Import-Module $modulePath -Force -ErrorAction Stop
}

AfterAll {
    Remove-Module Office365-Scoutnet-synk -ErrorAction SilentlyContinue
}

Describe 'Invoke-SNSCreateUserAndUpdateUserData' {
    It 'creates a user and deduplicates alternate mail addresses' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-MgUser {}
            function New-MgUser {
                param($BodyParameter, $ErrorAction)
            }
            function Set-MgUserLicense {}
            function New-MgGroupMember {}
            function Get-EXOMailbox {}
            function Set-Mailbox {}
            function Set-MailboxMessageConfiguration {}

            $credential = [pscredential]::new('test', (ConvertTo-SecureString 'test' -AsPlainText -Force))
            $script:SNSConf = [SNSConfiguration]::new()
            $script:SNSConf.DomainName = 'example.org'
            $script:SNSConf.WaitMailboxCreationMaxTime = '1'
            $script:SNSConf.WaitMailboxCreationPollTime = '0'
            $script:SNSConf.SignatureHtml = '<DisplayName>'
            $script:SNSConf.SignatureText = 'Hello <DisplayName>'
            $script:SNSConf.CredentialMemberlist = $credential
            $member = [pscustomobject]@{
                member_no = [pscustomobject]@{ value = 'member-1' }
                first_name = [pscustomobject]@{ value = 'Ada' }
                last_name = [pscustomobject]@{ value = 'Lovelace' }
                email = [pscustomobject]@{ value = 'ada@example.org' }
                contact_alt_email = [pscustomobject]@{ value = 'ada@example.org' }
                postcode = [pscustomobject]@{ value = '' }
                town = [pscustomobject]@{ value = '' }
                country = [pscustomobject]@{ value = '' }
                contact_mobile_phone = [pscustomobject]@{ value = '' }
                address_1 = [pscustomobject]@{ value = '' }
                address_2 = [pscustomobject]@{ value = '' }
                address_3 = [pscustomobject]@{ value = '' }
            }
            Mock Write-SNSLog {}
            Mock Get-SNSSecurityGroupScoutnet { return [pscustomobject]@{ Id = 'active-group' } }
            Mock Get-RandomPassword { return 'TestPassword123' }
            Mock Get-MgUser { return $null }
            Mock New-MgUser {
                return [pscustomobject]@{
                    Id = 'user-id'
                    UserPrincipalName = 'ada.lovelace@example.org'
                    DisplayName = 'Ada Lovelace'
                }
            }
            Mock Set-MgUserLicense {}
            Mock New-MgGroupMember {}
            Mock Start-Sleep {}
            Mock Get-EXOMailbox {}
            Mock Set-Mailbox {}
            Mock Set-MailboxMessageConfiguration {}

            Invoke-SNSCreateUserAndUpdateUserData -memberData @($member) -userLicenseData @()

            Should -Invoke New-MgUser -Times 1 -Exactly -ParameterFilter {
                $BodyParameter.UserPrincipalName -eq 'ada.lovelace@example.org' -and
                $BodyParameter.OtherMails.Count -eq 1 -and
                $BodyParameter.OtherMails[0] -eq 'ada@example.org'
            }
            Should -Invoke Set-MgUserLicense -Times 1 -Exactly
            Should -Invoke New-MgGroupMember -Times 1 -Exactly
            Should -Invoke Set-Mailbox -Times 1 -Exactly
        }
    }

    It 'does not attempt to create a user without a Scoutnet email address' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-MgUser {}
            function New-MgUser {}

            $script:SNSConf = [SNSConfiguration]::new()
            $member = [pscustomobject]@{
                email = [pscustomobject]@{ value = ' ' }
                first_name = [pscustomobject]@{ value = 'Ada' }
                last_name = [pscustomobject]@{ value = 'Lovelace' }
            }

            Mock Write-SNSLog {}
            Mock Get-SNSSecurityGroupScoutnet { return [pscustomobject]@{ Id = 'active-group' } }
            Mock Get-MgUser {}
            Mock New-MgUser {}

            Invoke-SNSCreateUserAndUpdateUserData -memberData @($member) -userLicenseData @()

            Should -Invoke Get-MgUser -Times 0 -Exactly
            Should -Invoke New-MgUser -Times 0 -Exactly
        }
    }
}

Describe 'Invoke-SNSUpdateAccount' {
    It 'updates changed profile fields and keeps duplicate Scoutnet email addresses unique' {
        InModuleScope Office365-Scoutnet-synk {
            function Get-MgUser {}
            function Update-MgUser {
                param($UserId, $BodyParameter, $ErrorAction)
            }

            $credential = [pscredential]::new('test', (ConvertTo-SecureString 'test' -AsPlainText -Force))
            $accountData = [pscustomobject]@{
                name = 'Ada Lovelace'
                CustomAttribute1 = 'member-1'
                ExternalDirectoryObjectId = 'user-id'
                UserPrincipalName = 'ada.lovelace@example.org'
            }
            $member = [pscustomobject]@{
                address_1 = [pscustomobject]@{ value = '1 Main Street' }
                address_2 = [pscustomobject]@{ value = '' }
                address_3 = [pscustomobject]@{ value = '' }
                postcode = [pscustomobject]@{ value = '12345' }
                town = [pscustomobject]@{ value = 'London' }
                country = [pscustomobject]@{ value = 'GB' }
                first_name = [pscustomobject]@{ value = 'Ada' }
                last_name = [pscustomobject]@{ value = 'Lovelace' }
                contact_mobile_phone = [pscustomobject]@{ value = '+44123456789' }
                email = [pscustomobject]@{ value = 'ada@example.org' }
                contact_alt_email = [pscustomobject]@{ value = 'ada@example.org' }
            }
            $existingUser = [pscustomobject]@{
                StreetAddress = ''
                PostalCode = ''
                City = ''
                Country = ''
                DisplayName = ''
                GivenName = ''
                Surname = ''
                MobilePhone = ''
                OtherMails = @()
            }
            Mock Write-SNSLog {}
            Mock Get-SNSApiGroupMemberlist { return @{ data = @{ 'member-1' = $member } } }
            Mock Get-MgUser { return $existingUser }
            Mock Update-MgUser {}

            Invoke-SNSUpdateAccount -AccountData $accountData -CredentialMemberlist $credential

            Should -Invoke Update-MgUser -Times 1 -Exactly -ParameterFilter {
                $BodyParameter.StreetAddress -eq '1 Main Street' -and
                $BodyParameter.DisplayName -eq 'Ada Lovelace' -and
                $BodyParameter.OtherMails.Count -eq 1 -and
                $BodyParameter.OtherMails[0] -eq 'ada@example.org'
            }
        }
    }
}

Describe 'Invoke-SNSDisableAccount' {
    It 'disables the account and moves it from the active group to the disabled group' {
        InModuleScope Office365-Scoutnet-synk {
            function Update-MgUser {
                param($UserId, $AccountEnabled, $ErrorAction)
            }
            function Set-Mailbox {}
            function Remove-MgGroupMemberByRef {
                param($GroupId, $DirectoryObjectId, $ErrorAction)
            }
            function New-MgGroupMember {
                param($GroupId, $DirectoryObjectId, $ErrorAction)
            }
            function Set-MailboxAutoReplyConfiguration {
                param($Identity, $AutoReplyState, $ExternalAudience, $InternalMessage, $ExternalMessage)
            }

            $script:SNSConf = [SNSConfiguration]::new()
            $script:SNSConf.DisabledAccountsAutoReplyText = 'Contact <DisplayName>'
            $accountData = [pscustomobject]@{
                name = 'Ada Lovelace'
                DisplayName = 'Ada Lovelace'
                ExternalDirectoryObjectId = 'user-id'
                UserPrincipalName = 'ada.lovelace@example.org'
                Identity = 'ada.lovelace@example.org'
            }
            Mock Write-SNSLog {}
            Mock Update-MgUser {}
            Mock Set-Mailbox {}
            Mock Get-SNSSecurityGroupScoutnet { return [pscustomobject]@{ Id = 'active-group' } }
            Mock Get-SNSDisabledScoutnetUserGroup { return [pscustomobject]@{ Id = 'disabled-group' } }
            Mock Remove-MgGroupMemberByRef {}
            Mock New-MgGroupMember {}
            Mock Set-MailboxAutoReplyConfiguration {}

            Invoke-SNSDisableAccount -AccountData $accountData

            Should -Invoke Update-MgUser -Times 1 -Exactly -ParameterFilter {
                $UserId -eq 'user-id' -and $AccountEnabled -eq $false
            }
            Should -Invoke Remove-MgGroupMemberByRef -Times 1 -Exactly -ParameterFilter {
                $GroupId -eq 'active-group' -and $DirectoryObjectId -eq 'user-id'
            }
            Should -Invoke New-MgGroupMember -Times 1 -Exactly -ParameterFilter {
                $GroupId -eq 'disabled-group' -and $DirectoryObjectId -eq 'user-id'
            }
            Should -Invoke Set-MailboxAutoReplyConfiguration -Times 1 -Exactly -ParameterFilter {
                $InternalMessage -eq 'Contact Ada Lovelace'
            }
        }
    }
}
