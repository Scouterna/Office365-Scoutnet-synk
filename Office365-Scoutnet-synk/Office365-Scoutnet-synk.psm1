#Requires -Version 7.6

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls13

#region import everything we need

. $PSScriptRoot\SNSConfiguration.ps1
. $PSScriptRoot\Write-SNSLog.ps1
. $PSScriptRoot\Get-SNSStableHash.ps1
. $PSScriptRoot\Receive-SNSApiJson.ps1
. $PSScriptRoot\Get-SNSApiGroupCustomlist.ps1
. $PSScriptRoot\Get-SNSMaillistInfo.ps1
. $PSScriptRoot\Get-SNSUserEmail.ps1
. $PSScriptRoot\Get-SNSExchangeMailListMember.ps1
. $PSScriptRoot\Get-SNSApiGroupMemberlist.ps1
. $PSScriptRoot\Send-SNSMandatoryMomentReport.ps1
. $PSScriptRoot\Update-SNMicrosoft365Group.ps1
. $PSScriptRoot\Get-SNSRoleListMember.ps1
. $PSScriptRoot\Invoke-SNSUppdateOffice365User.ps1
. $PSScriptRoot\Connect-SnSOffice365.ps1
. $PSScriptRoot\Disconnect-SnSOffice365.ps1

#endregion

# Configuration holder.
$script:SNSConf=[SNSConfiguration]::new()

function Add-Office365UserToList
{
    [OutputType([bool], [string])]
    param (
        [Parameter(Mandatory=$True, HelpMessage="List with all Office365 users.")]
        [ValidateNotNull()]
        $allOffice365Users,

        [Parameter(Mandatory=$True, HelpMessage="Scoutnet Id for member to add.")]
        [ValidateNotNull()]
        $member,

        [Parameter(Mandatory=$True, HelpMessage="Data for the member from Scoutnet.")]
        [ValidateNotNull()]
        $MemberData,

        [Parameter(Mandatory=$True, HelpMessage="Name of group to add member in.")]
        [ValidateNotNull()]
        $distGroupName,

        [Parameter(HelpMessage="Do warn about missing Office 365 user.")]
        [switch] $doWarn
        )

    $recipient = $allOffice365Users | Where-Object { $_.CustomAttribute1 -eq $member }

    $userAdded = $False
    $PrimarySmtpAddress = ""

    if ($recipient)
    {
        $PrimarySmtpAddress = $recipient.PrimarySmtpAddress
        $userAdded = $True
        Write-SNSLog "Added member $($recipient.DisplayName) with smtp address $($recipient.PrimarySmtpAddress) to distribution group $distGroupName"
    }
    elseif ($doWarn)
    {
        Write-SNSLog -Level "Warn" "Member $($MemberData.first_name) $($MemberData.last_name) not found in office 365. Please make sure that CustomAttribute1 contains Scoutnet Id for the user."
    }
    return $userAdded, $PrimarySmtpAddress
}

function Add-SNSContactTiming
{
    param (
        [Parameter(Mandatory)]
        [hashtable]$TimingData,

        [Parameter(Mandatory)]
        [string]$GroupName,

        [Parameter(Mandatory)]
        [ValidateSet(
            'RecipientLookupSeconds',
            'ContactManagementSeconds',
            'GetContactSeconds',
            'NewMailContactSeconds',
            'ProvisioningWaitSeconds',
            'SetContactSeconds',
            'SetMailContactSeconds',
            'ContactCacheHits',
            'RecipientIndexHits'
        )]
        [string]$Metric,

        [Parameter()]
        [double]$Seconds = 0
    )

    if (-not $TimingData.ContainsKey($GroupName))
    {
        $TimingData[$GroupName] = @{
            RecipientLookupSeconds = 0.0
            RecipientLookupCount = 0
            ContactManagementSeconds = 0.0
            ContactManagementCount = 0
            GetContactSeconds = 0.0
            GetContactCount = 0
            NewMailContactSeconds = 0.0
            NewMailContactCount = 0
            ProvisioningWaitSeconds = 0.0
            ProvisioningWaitCount = 0
            SetContactSeconds = 0.0
            SetContactCount = 0
            SetMailContactSeconds = 0.0
            SetMailContactCount = 0
            ContactCacheHits = 0
            RecipientIndexHits = 0
        }
    }

    $groupTiming = $TimingData[$GroupName]
    if ($Metric -eq 'ContactCacheHits' -or $Metric -eq 'RecipientIndexHits')
    {
        $groupTiming[$Metric]++
    }
    else
    {
        $groupTiming[$Metric] += $Seconds
        $countName = $Metric -replace 'Seconds$', 'Count'
        $groupTiming[$countName]++
    }
}

function ConvertTo-SNSContactEmailKey
{
    param (
        [Parameter()]
        [AllowNull()]
        [object]$Address
    )

    $emailAddress = ([string]$Address).Trim()
    if ($emailAddress -match '^[^:]+:(.+)$')
    {
        $emailAddress = $Matches[1]
    }
    if ($emailAddress -notmatch '^[^@\s]+@[^@\s]+$')
    {
        return $null
    }
    return $emailAddress
}

function Test-SNSGroupMembersUnchanged
{
    param (
        [Parameter()]
        [AllowNull()]
        $CurrentMemberKeys,

        [Parameter()]
        [AllowNull()]
        $DesiredAddresses,

        [Parameter()]
        [System.Collections.Generic.List[string]]$Differences = [System.Collections.Generic.List[string]]::new()
    )

    # Unknown current members (new group or no data) means the group must be updated.
    if ($null -eq $CurrentMemberKeys)
    {
        $Differences.Add('current members unknown')
        return $false
    }

    $unchanged = $true
    $desired = @{}
    foreach ($address in @($DesiredAddresses))
    {
        $key = ConvertTo-SNSContactEmailKey -Address $address
        if ($null -eq $key)
        {
            $Differences.Add("invalid wanted address '$address'")
            $unchanged = $false
            continue
        }
        $desired[$key] = $false
    }

    foreach ($memberKeys in @($CurrentMemberKeys))
    {
        $matched = $false
        foreach ($key in @($memberKeys))
        {
            if ($desired.ContainsKey($key))
            {
                $desired[$key] = $true
                $matched = $true
            }
        }
        if (-not $matched)
        {
            $Differences.Add("unwanted member with addresses '$(@($memberKeys) -join ', ')'")
            $unchanged = $false
        }
    }

    foreach ($key in $desired.Keys)
    {
        if (-not $desired[$key])
        {
            $Differences.Add("missing member '$key'")
            $unchanged = $false
        }
    }

    return $unchanged
}

function Get-SNSRecipientAddressKey
{
    param (
        [Parameter()]
        $Recipient,

        [Parameter()]
        [string[]]$AdditionalAddresses = @()
    )

    $addresses = [System.Collections.Generic.List[object]]::new()
    foreach ($address in $AdditionalAddresses)
    {
        $addresses.Add($address)
    }
    foreach ($propertyName in @('PrimarySmtpAddress', 'ExternalEmailAddress', 'WindowsEmailAddress', 'EmailAddresses'))
    {
        foreach ($address in @(Get-SNMicrosoft365Value -Object $Recipient -Name $propertyName))
        {
            if ($null -ne $address)
            {
                $addresses.Add($address)
            }
        }
    }

    foreach ($address in $addresses)
    {
        $key = ConvertTo-SNSContactEmailKey -Address $address
        if ($null -ne $key)
        {
            $key
        }
    }
}

function Add-SNSRecipientToIndex
{
    param (
        [Parameter(Mandatory)]
        [hashtable]$RecipientIndex,

        [Parameter()]
        $Recipient,

        [Parameter()]
        [string[]]$AdditionalAddresses = @()
    )

    foreach ($key in @(Get-SNSRecipientAddressKey -Recipient $Recipient -AdditionalAddresses $AdditionalAddresses))
    {
        $RecipientIndex[$key] = $Recipient
    }
}

function Add-SNSContactInfoToCache
{
    param (
        [Parameter(Mandatory)]
        [hashtable]$ContactInfoCache,

        [Parameter()]
        $Contact,

        [Parameter()]
        [string[]]$AdditionalAddresses = @(),

        [Parameter()]
        [AllowNull()]
        [string]$Company
    )

    $companyValue = if ($PSBoundParameters.ContainsKey('Company'))
    {
        $Company
    }
    else
    {
        [string](Get-SNMicrosoft365Value -Object $Contact -Name 'Company')
    }
    $contactInfo = @{ Company = $companyValue }
    foreach ($key in @(Get-SNSRecipientAddressKey -Recipient $Contact -AdditionalAddresses $AdditionalAddresses))
    {
        $ContactInfoCache[$key] = $contactInfo
    }
    return ,$contactInfo
}

function New-SNSMailContact
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([bool])]
    param (
        [ValidateNotNull()]
        $Epost,

        [ValidateNotNull()]
        $DisplayName,

        [Parameter()]
        [hashtable]$ContactCache,

        [Parameter()]
        [hashtable]$ContactInfoCache,

        [Parameter()]
        [hashtable]$RecipientIndex,

        [Parameter()]
        [hashtable]$TimingData,

        [Parameter()]
        [string]$GroupName
        )

    if ($ContactCache -and $ContactCache.ContainsKey([string]$Epost))
    {
        if ($TimingData -and $GroupName)
        {
            Add-SNSContactTiming -TimingData $TimingData -GroupName $GroupName -Metric ContactCacheHits
        }
        return $ContactCache[[string]$Epost]
    }

    $contactCreated = $True
    $recipientKey = ConvertTo-SNSContactEmailKey -Address $Epost
    if ($RecipientIndex -and $null -ne $recipientKey -and $RecipientIndex.ContainsKey($recipientKey))
    {
        $ExistingMailContact = $RecipientIndex[$recipientKey]
        if ($TimingData -and $GroupName)
        {
            Add-SNSContactTiming -TimingData $TimingData -GroupName $GroupName -Metric RecipientIndexHits
        }
    }
    else
    {
        $recipientLookupStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $ExistingMailContact = Get-EXORecipient $Epost -ErrorAction "SilentlyContinue" -Verbose:$false
        $recipientLookupStopwatch.Stop()
        if ($TimingData -and $GroupName)
        {
            Add-SNSContactTiming -TimingData $TimingData -GroupName $GroupName `
                -Metric RecipientLookupSeconds -Seconds $recipientLookupStopwatch.Elapsed.TotalSeconds
        }
        if ($RecipientIndex -and $null -ne $ExistingMailContact)
        {
            Add-SNSRecipientToIndex -RecipientIndex $RecipientIndex -Recipient $ExistingMailContact `
                -AdditionalAddresses @($Epost)
        }
    }

    $contactManagementStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    if (-not $PSCmdlet.ShouldProcess($Epost, "Create or update mail contact"))
    {
        $contactManagementStopwatch.Stop()
        if ($TimingData -and $GroupName)
        {
            Add-SNSContactTiming -TimingData $TimingData -GroupName $GroupName `
                -Metric ContactManagementSeconds -Seconds $contactManagementStopwatch.Elapsed.TotalSeconds
        }
        if ($ContactCache)
        {
            $ContactCache[[string]$Epost] = $contactCreated
        }
        return $contactCreated
    }

    if ($null -eq $ExistingMailContact)
    {
        Write-SNSLog "Creating Contact $Epost for $DisplayName"
        try
        {
            $commandStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            try
            {
                $createdMailContact = New-MailContact -Name $Epost -ExternalEmailAddress $Epost -ErrorAction "stop" -Verbose:$false
            }
            finally
            {
                $commandStopwatch.Stop()
                if ($TimingData -and $GroupName)
                {
                    Add-SNSContactTiming -TimingData $TimingData -GroupName $GroupName `
                        -Metric NewMailContactSeconds -Seconds $commandStopwatch.Elapsed.TotalSeconds
                }
            }
            if ($RecipientIndex)
            {
                Add-SNSRecipientToIndex -RecipientIndex $RecipientIndex -Recipient $createdMailContact `
                    -AdditionalAddresses @($Epost)
            }
            if ($ContactInfoCache)
            {
                $mailContactInfo = Add-SNSContactInfoToCache -ContactInfoCache $ContactInfoCache `
                    -Contact $createdMailContact -AdditionalAddresses @($Epost)
            }

            for($i=1; $i -le 10; $i++)
            {
                # It takes some time befor the new contact can be accessed. Wait and try to access it.
                $waitStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                try
                {
                    Start-Sleep -s 1.0
                }
                finally
                {
                    $waitStopwatch.Stop()
                    if ($TimingData -and $GroupName)
                    {
                        Add-SNSContactTiming -TimingData $TimingData -GroupName $GroupName `
                            -Metric ProvisioningWaitSeconds -Seconds $waitStopwatch.Elapsed.TotalSeconds
                    }
                }

                $commandStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                try
                {
                    $ExistingMailContact = Get-Contact $Epost -ErrorAction "SilentlyContinue" -Verbose:$false
                }
                finally
                {
                    $commandStopwatch.Stop()
                    if ($TimingData -and $GroupName)
                    {
                        Add-SNSContactTiming -TimingData $TimingData -GroupName $GroupName `
                            -Metric GetContactSeconds -Seconds $commandStopwatch.Elapsed.TotalSeconds
                    }
                }
                if ($null -ne $ExistingMailContact)
                {
                    $i = 10
                }
            }
            if ($ContactInfoCache -and $null -eq $mailContactInfo)
            {
                $mailContactInfo = Add-SNSContactInfoToCache -ContactInfoCache $ContactInfoCache `
                    -Contact $ExistingMailContact -AdditionalAddresses @($Epost)
            }

            try
            {
                # Set the name of the member in the company field. This is visibel in Office 365 admin console.
                $commandStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                try
                {
                    Set-Contact -Identity $Epost -Company "$DisplayName" -ErrorAction "stop" -Verbose:$false
                    if ($ContactInfoCache -and $mailContactInfo)
                    {
                        $mailContactInfo.Company = [string]$DisplayName
                    }
                }
                finally
                {
                    $commandStopwatch.Stop()
                    if ($TimingData -and $GroupName)
                    {
                        Add-SNSContactTiming -TimingData $TimingData -GroupName $GroupName `
                            -Metric SetContactSeconds -Seconds $commandStopwatch.Elapsed.TotalSeconds
                    }
                }
                try
                {
                    # Hide the contact from address lists.
                    $commandStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                    try
                    {
                        Set-MailContact -Identity $Epost -HiddenFromAddressListsEnabled $true -ErrorAction "stop" -Verbose:$false
                    }
                    finally
                    {
                        $commandStopwatch.Stop()
                        if ($TimingData -and $GroupName)
                        {
                            Add-SNSContactTiming -TimingData $TimingData -GroupName $GroupName `
                                -Metric SetMailContactSeconds -Seconds $commandStopwatch.Elapsed.TotalSeconds
                        }
                    }
                }
                Catch
                {
                    Write-SNSLog -Level "Warn" "Could not set hidden from address lists field for $Epost, will try again next time the sync runs. Error $_"
                }
            }
            Catch
            {
                Write-SNSLog -Level "Warn" "Could not set company field for $Epost, will try again next time the sync runs. Error $_"
            }
        }
        Catch
        {
            Write-SNSLog -Level "Warn" "Could not create mail contact with address $Epost. Error $_"
            $contactCreated = $False
        }
    }
    else
    {
        try
        {
            $contactKey = ConvertTo-SNSContactEmailKey -Address $Epost
            if ($ContactInfoCache -and $null -ne $contactKey -and $ContactInfoCache.ContainsKey($contactKey))
            {
                $MailContactInfo = $ContactInfoCache[$contactKey]
            }
            else
            {
                $commandStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                try
                {
                    $MailContactInfo = Get-Contact $Epost -ErrorAction "SilentlyContinue" -Verbose:$false
                }
                finally
                {
                    $commandStopwatch.Stop()
                    if ($TimingData -and $GroupName)
                    {
                        Add-SNSContactTiming -TimingData $TimingData -GroupName $GroupName `
                            -Metric GetContactSeconds -Seconds $commandStopwatch.Elapsed.TotalSeconds
                    }
                }
                if ($ContactInfoCache -and $null -ne $MailContactInfo)
                {
                    $MailContactInfo = Add-SNSContactInfoToCache -ContactInfoCache $ContactInfoCache `
                        -Contact $MailContactInfo -AdditionalAddresses @($Epost)
                }
            }
            # Set the company field for the existing contact and hide it from address lists if necessary.
            if ($MailContactInfo.Company -eq "")
            {
                try
                {
                    # Set the company field for the existing contact.
                    $commandStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                    try
                    {
                        Set-Contact -Identity $Epost -Company "$DisplayName" -ErrorAction "stop" -Verbose:$false
                        if ($ContactInfoCache -and $MailContactInfo)
                        {
                            $MailContactInfo.Company = [string]$DisplayName
                        }
                    }
                    finally
                    {
                        $commandStopwatch.Stop()
                        if ($TimingData -and $GroupName)
                        {
                            Add-SNSContactTiming -TimingData $TimingData -GroupName $GroupName `
                                -Metric SetContactSeconds -Seconds $commandStopwatch.Elapsed.TotalSeconds
                        }
                    }
                    try
                    {
                        # Hide the contact from address lists.
                        $commandStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                        try
                        {
                            Set-MailContact -Identity $Epost -HiddenFromAddressListsEnabled $true -ErrorAction "stop" -Verbose:$false
                        }
                        finally
                        {
                            $commandStopwatch.Stop()
                            if ($TimingData -and $GroupName)
                            {
                                Add-SNSContactTiming -TimingData $TimingData -GroupName $GroupName `
                                    -Metric SetMailContactSeconds -Seconds $commandStopwatch.Elapsed.TotalSeconds
                            }
                        }
                    }
                    Catch
                    {
                        Write-SNSLog -Level "Warn" "Could not set hidden from address lists field for $Epost, will try again next time the sync runs. Error $_"
                    }
                }
                Catch
                {
                    Write-SNSLog -Level "Warn" "Could not set company field for $Epost, will try again next time the sync runs. Error $_"
                }
            }
        }
        Catch
        {
            Write-SNSLog -Level "Warn" "Could not get mail contact info for $Epost, will try again next time the sync runs. Error $_"
        }
    }

    $contactManagementStopwatch.Stop()
    if ($TimingData -and $GroupName)
    {
        Add-SNSContactTiming -TimingData $TimingData -GroupName $GroupName `
            -Metric ContactManagementSeconds -Seconds $contactManagementStopwatch.Elapsed.TotalSeconds
    }
    if ($contactCreated -and $ContactCache)
    {
        $ContactCache[[string]$Epost] = $contactCreated
    }

    return $contactCreated
}


function Get-SNSMailListSubjectPrefix
{
    <#
    .SYNOPSIS
        Returns the subject prefix for a distribution group rule in MailListSettings.

    .DESCRIPTION
        prefix_text is used as it is (trimmed). Without prefix_text the prefix is the display_name in square
        brackets, for example [Krypen].

    .PARAMETER Setting
        The rule from MailListSettings.

    .OUTPUTS
        System.String. The prefix without trailing space, or an empty string when neither prefix_text nor display_name is specified.
    #>
    [OutputType([string])]
    param (
        [Parameter()]
        $Setting
    )

    $prefixText = ([string](Get-SNMicrosoft365Value -Object $Setting -Name 'prefix_text')).Trim()
    if ($prefixText -ne '')
    {
        return $prefixText
    }

    $displayName = ([string](Get-SNMicrosoft365Value -Object $Setting -Name 'display_name')).Trim()
    if ($displayName -ne '')
    {
        return "[$displayName]"
    }
    return ''
}


function Set-SNSMailListPrefixRule
{
    <#
    .SYNOPSIS
        Creates or updates the Exchange mail flow rule that prefixes the subject of mail sent to a distribution group.

    .DESCRIPTION
        The rule matches the group address in the To or Cc header, which is evaluated before Exchange expands the
        group, and prepends the prefix to the subject. A subject that already contains the prefix is left as it is.
        The rule is named 'Subject prefix for <address>'. An existing rule with that name gets its prefix updated.

    .PARAMETER GroupName
        The key of the rule in MailListSettings (name or e-mail address of the group).

    .PARAMETER Setting
        The rule from MailListSettings.
    #>
    [CmdletBinding(SupportsShouldProcess = $True)]
    param (
        [Parameter(Mandatory = $True)]
        [string]$GroupName,

        [Parameter()]
        $Setting
    )

    $prefix = Get-SNSMailListSubjectPrefix -Setting $Setting
    if ($prefix -eq '')
    {
        Write-SNSLog -Level "Warn" "add_prefix is set for $GroupName but neither prefix_text nor display_name is specified. No subject prefix rule is created."
        return
    }

    $address = ([string](Get-SNMicrosoft365Value -Object $Setting -Name 'primary_smtp_address')).Trim()
    if ($address -eq '' -and $GroupName -like '*@*')
    {
        $address = $GroupName
    }
    if ($address -eq '')
    {
        $address = [string](Get-EXORecipient -Identity $GroupName -ErrorAction "SilentlyContinue" -Verbose:$false).PrimarySmtpAddress
    }
    if ($address -eq '')
    {
        Write-SNSLog -Level "Warn" "The e-mail address of $GroupName is not known. No subject prefix rule is created."
        return
    }

    $ruleName = "Subject prefix for $address"
    $prependText = "$prefix "
    try
    {
        $rule = Get-TransportRule -Identity $ruleName -ErrorAction "SilentlyContinue" -Verbose:$false
        if ($null -eq $rule)
        {
            if ($PSCmdlet.ShouldProcess($ruleName, "New-TransportRule"))
            {
                Write-SNSLog "Creating mail flow rule '$ruleName' with prefix $prefix"
                New-TransportRule -Name $ruleName -AnyOfToCCHeader $address -PrependSubject $prependText `
                    -ExceptIfSubjectContainsWords $prefix -AuditSeverity "DoNotAudit" -SenderAddressLocation "Header" `
                    -RuleErrorAction "Ignore" -Mode "Enforce" -ErrorAction "stop" -Verbose:$false -Confirm:$false | Out-Null
            }
        }
        elseif ($rule.PrependSubject -ne $prependText -or (@($rule.ExceptIfSubjectContainsWords) -join '|') -ne $prefix)
        {
            if ($PSCmdlet.ShouldProcess($ruleName, "Set-TransportRule"))
            {
                Write-SNSLog "Updating mail flow rule '$ruleName' to prefix $prefix"
                Set-TransportRule -Identity $ruleName -PrependSubject $prependText -ExceptIfSubjectContainsWords $prefix `
                    -ErrorAction "stop" -Verbose:$false -Confirm:$false | Out-Null
            }
        }
    }
    catch
    {
        Write-SNSLog -Level "Error" "Creation or update of mail flow rule '$ruleName' failed! Error $_"
    }
}


function Write-SNSTimingLog
{
    param (
        [Parameter()]
        [bool]$Enabled,

        [Parameter(Mandatory)]
        [string]$Phase,

        [Parameter(Mandatory)]
        [System.Diagnostics.Stopwatch]$Stopwatch
    )

    $Stopwatch.Stop()
    if ($Enabled)
    {
        Write-SNSLog ("Timing: {0} took {1:N2} seconds" -f $Phase, $Stopwatch.Elapsed.TotalSeconds)
    }
}

function SNSUpdateExchangeDistributionGroups
{
    <#
    .SYNOPSIS
        Updates exchange distribution groups based on scoutnet maillists.

    .DESCRIPTION
        Fetches the distribution groups members and updates corressponding distribution groups based on the info from scoutnet.
        As all members of an distribution group must be present in exchange as user or contact, contacts will be created for external addresses.
        Logs elapsed time for the main data-fetching and synchronization phases.

        A rule in MailListSettings without 'source' uses a Scoutnet maillist (scoutnet_list_id). With source = 'troop_roles'
        (troops = @(@{ name; id; roles })) or source = 'group_roles' (roles) the leaders are instead selected by their roles in
        Scoutnet, in the same way as for Update-SNMicrosoft365Group, and roles = '*' matches any role. Role based rules
        use ledare_synk_option to decide which address is added. scoutnet_list_id is not used.
        group_type = 'security' or 'distribution' (works for all sources) creates the group as a mail-enabled security group
        or a regular distribution list with
        New-DistributionGroup when it does not exist. display_name (the visible name, since the rule key is
        normally the e-mail address), mail_nickname (alias) and primary_smtp_address must all be specified to create
        a group; otherwise a warning is logged and the group is skipped while the other lists continue.
        allow_external_senders = $true also
        lets senders outside the organization mail the new group (RequireSenderAuthenticationEnabled = $false);
        it is only applied when the group is created. An existing group of another type is not
        changed. Without group_type the
        group must already exist.
        add_prefix = $true creates (or updates) an Exchange mail flow rule that prepends a prefix to the subject
        of mail sent to the group. The prefix is prefix_text, or [display_name] when prefix_text is not specified.
        A troop rule with unit_members = $true also adds the members of that unit (the 'unit' object of the member)
        using scouter_synk_option, so a list can hold the scouts of a troop plus its leaders. Roles are optional then.

    .INPUTS
        None. You cannot pipe objects to SNSUpdateExchangeDistributionGroups.

    .OUTPUTS
        The new validation hash as a string. With -ReturnMaildata, returns the validation hash and a hashtable of mail addresses.

    .LINK
        https://github.com/scouternasetjanster/Office365-Scoutnet-synk

    .PARAMETER ValidationHash
        Hash value used to validate if scoutnet is updated.

    .PARAMETER Configuration
        Configuration to use. If not specified the cached configuration will be used.

    .PARAMETER ReturnMaildata
        Enables return of a Hashtable with mailaddresses for the list. Can be used together with -WhatIf to check settings.

    .PARAMETER WhatIf
        The WhatIf switch simulates the actions of the command.
        You can use this switch to view the changes that would occur without actually applying those changes.
        You don't need to specify a value with this switch.
    #>
    [CmdletBinding(HelpURI = 'https://github.com/scouternasetjanster/Office365-Scoutnet-synk',
                PositionalBinding = $False,
                SupportsShouldProcess = $True)]
    [OutputType([string], [object[]])]
    param (
        [Parameter(Mandatory=$True, HelpMessage="Hash value used to validate if Scoutnet is updated.")]
        [ValidateNotNull()]
        [string]$ValidationHash,

        [Parameter(Mandatory=$False, HelpMessage="Configuratin to use. If not specified the cached configuration will be used.")]
        $Configuration,

        [Parameter(HelpMessage="Enables return of a Hashtable with mailaddresses for the list. Can be used together with -WhatIf to check settings.")]
        [switch] $ReturnMaildata
    )

    $SNSConf = $Script:SNSConf
    if ($Configuration)
    {
        $SNSConf = $Configuration
    }

    if (!$SNSConf)
    {
        throw "No configuration specified. Please provide a configuration!"
    }

    $totalStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $phaseStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $AddressesInMailLists = @{}

    $MailListSettings = $SNSConf.MailListSettings

    # Rules without source are driven by Scoutnet maillists. Rules with source troop_roles or group_roles are driven by roles.
    $roleListNames = @()
    foreach ($distGroupName in $MailListSettings.Keys)
    {
        Test-SNSRoleListSetting -Name $distGroupName -Setting $MailListSettings[$distGroupName]
        [void](Get-SNSMailListGroupType -Setting $MailListSettings[$distGroupName])
        if ((Get-SNSMailListSource -Setting $MailListSettings[$distGroupName]) -ne 'mail_list')
        {
            $roleListNames += $distGroupName
        }
    }
    Write-SNSTimingLog -Enabled $SNSConf.EnableTimingLog -Phase 'Validate distribution group settings' -Stopwatch $phaseStopwatch

    # Fetch maillist info from scoutnet.
    $maillistRules = @($MailListSettings.Keys | Where-Object { $roleListNames -notcontains $_ } | ForEach-Object { $MailListSettings[$_] })
    $phaseStopwatch.Restart()
    $CustomLists, $CustomListsHash = Get-SNSMaillistInfo -CredentialCustomlists $SNSConf.CredentialCustomlists -MaillistsIds $maillistRules
    Write-SNSTimingLog -Enabled $SNSConf.EnableTimingLog -Phase 'Fetch Scoutnet mail lists' -Stopwatch $phaseStopwatch

    # Fetch all members and their mailaddresses.
    $phaseStopwatch.Restart()
    $allMailAddresses, $allMailAddressesHash = Get-SNSUserEmail -CredentialMemberlist $SNSConf.CredentialMemberlist
    Write-SNSTimingLog -Enabled $SNSConf.EnableTimingLog -Phase 'Fetch and prepare member email addresses' -Stopwatch $phaseStopwatch

    $phaseStopwatch.Restart()
    $MailListSettingsHash = Get-SNSStableHash -InputObject $MailListSettings

    $OutValidationHash = $ValidationHash
    $NewValidationHash = "0x{0}{1}{2}" -f ($CustomListsHash, $allMailAddressesHash, $MailListSettingsHash)

    # Roles are not part of the mail address hash, so the matched members of role based rules are added to the validation hash.
    $roleListMembers = @{}
    if ($roleListNames.Count -gt 0)
    {
        $memberlist = Get-SNSApiGroupMemberlist -Credential $SNSConf.CredentialMemberlist
        $roleListMembers = Get-SNSRoleListMember -MailListSettings $MailListSettings -MemberList $memberlist
        $NewValidationHash += Get-SNSStableHash -InputObject $roleListMembers
    }
    Write-SNSTimingLog -Enabled $SNSConf.EnableTimingLog -Phase 'Match role-based list members and calculate validation hash' -Stopwatch $phaseStopwatch

    Write-SNSLog "Saved validation hash: $ValidationHash new value $NewValidationHash"

    if ($ValidationHash -eq $NewValidationHash)
    {
        Write-SNSLog "Scoutnet is not updated. No update will be done."
    }
    else
    {
        Write-SNSLog "Scoutnet is updated. Starting to update the distribution groups."
        $phaseStopwatch.Restart()
        $otherMailListsMembers, $mailListsToProcessMembers, $currentGroupMemberKeys = Get-SNSExchangeMailListMember -Maillists $MailListSettings.Keys
        Write-SNSTimingLog -Enabled $SNSConf.EnableTimingLog -Phase 'Fetch Exchange distribution group members' -Stopwatch $phaseStopwatch

        $phaseStopwatch.Restart()
        $allOffice365Users = Get-EXOMailbox -RecipientTypeDetails "UserMailbox" -Properties CustomAttribute1 -Verbose:$false
        Write-SNSTimingLog -Enabled $SNSConf.EnableTimingLog -Phase 'Fetch Exchange mailboxes' -Stopwatch $phaseStopwatch

        $phaseStopwatch.Restart()
        try
        {
            $allRecipients = @(Get-EXORecipient -ResultSize Unlimited `
                -Properties PrimarySmtpAddress, ExternalEmailAddress, EmailAddresses, Company `
                -ErrorAction Stop -Verbose:$false)
        }
        catch
        {
            Write-SNSLog -Level "Error" "Could not fetch Exchange recipients for synchronization. Error $_"
            throw
        }
        $recipientIndex = @{}
        $mailContactInfoCache = @{}
        foreach ($recipient in $allRecipients)
        {
            Add-SNSRecipientToIndex -RecipientIndex $recipientIndex -Recipient $recipient
            if ((Get-SNMicrosoft365Value -Object $recipient -Name 'RecipientTypeDetails') -eq 'MailContact')
            {
                [void](Add-SNSContactInfoToCache -ContactInfoCache $mailContactInfoCache -Contact $recipient)
            }
        }
        Write-SNSTimingLog -Enabled $SNSConf.EnableTimingLog -Phase 'Fetch and index Exchange recipients' -Stopwatch $phaseStopwatch

        $mailContactCache = @{}
        $contactTimingByGroup = @{}
        $phaseStopwatch.Restart()
        foreach ($distGroupName in $MailListSettings.keys)
        {
            $groupStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $AddressesToAdd = @{}
#region Synch e-mail addresses from rule list 'scouter' using scouter_synk_option.
            $distGroupId = $MailListSettings[$distGroupName].scoutnet_list_id
            Write-SNSLog " "
            Write-SNSLog "Adding Contacts in distribution group $($distGroupName)"

            $isRoleList = $roleListNames -contains $distGroupName

            $scouter_synk_option = [string]$MailListSettings[$distGroupName].scouter_synk_option
            $scouter_synk_option = $scouter_synk_option.ToLower()

            # Fetch the mail addresses to add for scouter. Role based rules only contain leaders.
            Write-SNSLog "Scouter synk option is '$scouter_synk_option' for distribution group $distGroupName"
            if ($isRoleList)
            {
                $listData = $roleListMembers[$distGroupName].scouter
            }
            else
            {
                $listData = $CustomLists[$distGroupId].scouter
            }
            foreach ($member in $listData)
            {
                if ($null -eq $member)
                {
                    Write-SNSLog -Level "Warn" "Mail list $distGroupName type 'scouter' contained a null value."
                    continue
                }
                if (-not $AllMailAddresses.ContainsKey($member))
                {
                    Write-SNSLog -Level "Warn" "Id '$member' from $distGroupName not found in list of all members."
                    continue
                }

                $MemberData = $AllMailAddresses[$member]
                $displayName = "$($MemberData.first_name) $($MemberData.last_name)"

                if ($MailListSettings[$distGroupName].ignore_user)
                {
                    if ($MailListSettings[$distGroupName].ignore_user.Contains($member))
                    {
                        Write-SNSLog "Ignoring '$($displayName)' with Id '$member' for $distGroupName."
                        continue
                    }
                }

                # Get valid addresses to add based on list setting.
                $mailaddresses = [System.Collections.ArrayList]::new()
                if ($scouter_synk_option.contains("p"))
                {
                    [void]$mailaddresses.Add($MemberData.primary_email)
                }
                elseif ($scouter_synk_option.contains("a"))
                {
                    [void]$mailaddresses.Add($MemberData.alt_email)
                }
                elseif ($scouter_synk_option.contains("f"))
                {
                    $MemberData.contacts_addresses | ForEach-Object {[void]$mailaddresses.Add($_)}
                }
                else
                {
                    $MemberData.mailaddresses | ForEach-Object {[void]$mailaddresses.Add($_)}
                }
                $mailaddresses = $mailaddresses | Sort-Object -Unique

                $AddMemberOffice365Address = $False
                $AddMemberScoutnetAddress = $True
                $AddMemberOffice365AddressTryFirst = $False # First try to add office 365 address. If that fails add the scoutnet version.

                if ($scouter_synk_option.contains("@"))
                {
                    $AddMemberScoutnetAddress = $False
                    $AddMemberOffice365Address = $True
                }
                elseif ($scouter_synk_option.contains("&"))
                {
                    $AddMemberScoutnetAddress = $True
                    $AddMemberOffice365Address = $True
                }

                if ($scouter_synk_option.contains("t"))
                {
                    $AddMemberScoutnetAddress = $False
                    $AddMemberOffice365AddressTryFirst = $True
                    $AddMemberOffice365Address = $True
                }

                if ($mailaddresses.Length -eq 0)
                {
                    Write-SNSLog -Level "Warn" "No email addresses found for $($displayName)."
                    continue
                }

                $AddMemberScoutnetAddress = $AddMemberScoutnetAddress
                if ($AddMemberOffice365Address)
                {
                    $result, $PrimarySmtpAddress = Add-Office365userToList -allOffice365Users $allOffice365Users -Member $member -MemberData $AllMailAddresses[$member] -distGroupName $distGroupName
                    if ($result)
                    {
                        if ([string]::IsNullOrWhiteSpace($PrimarySmtpAddress))
                        {
                            Write-SNSLog -Level "Warn" "Primary mailaddres for '$member' is empty. Cannot add member to list."
                        }
                        else
                        {
                            $AddressesToAdd[$PrimarySmtpAddress] = $PrimarySmtpAddress
                        }
                    }
                    else
                    {
                        if ($AddMemberOffice365AddressTryFirst)
                        {
                            $AddMemberScoutnetAddress = $True
                        }
                    }
                }

                if ($AddMemberScoutnetAddress)
                {
                    $mailaddresses | ForEach-Object {
                        if (![string]::IsNullOrWhiteSpace($_))
                        {
                            $ContactCreated = New-SNSMailContact -Epost $_ -DisplayName $displayName `
                                -ContactCache $mailContactCache -ContactInfoCache $mailContactInfoCache -RecipientIndex $recipientIndex `
                                -TimingData $contactTimingByGroup -GroupName $distGroupName

                            if ($ContactCreated)
                            {
                                $AddressesToAdd[$_] = $_
                                Write-SNSLog "Adding Contact '$_' to '$distGroupName'"
                            }
                            $mailListsToProcessMembers.Remove($_)
                        }
                    }
                }
            }
#endregion

#region Synch e-mail addresses from rule list 'ledare' using ledare_synk_option.
            $ledare_synk_option = [string]$MailListSettings[$distGroupName].ledare_synk_option
            $ledare_synk_option = $ledare_synk_option.ToLower()

            Write-SNSLog "Ledare synk option is '$ledare_synk_option' for distribution group $distGroupName"
            # Get the settings for ledare in this list.
            $AddLedareOffice365Address = $True
            $AddLedareScoutnetAddress = $False
            $AddLedareOffice365AddressTryFirst = $False # First try to add office 365 address. If that fails add the scoutnet version.

            if ($ledare_synk_option -like "-")
            {
                $AddLedareScoutnetAddress = $True
                $AddLedareOffice365Address = $False
            }
            elseif ($ledare_synk_option -like "&")
            {
                $AddLedareScoutnetAddress = $True
            }

            if ($ledare_synk_option -like "t")
            {
                $AddLedareScoutnetAddress = $False
                $AddLedareOffice365AddressTryFirst = $True
                $AddLedareOffice365Address = $True
            }

            if ($isRoleList)
            {
                $listData = $roleListMembers[$distGroupName].ledare
            }
            else
            {
                $listData = $CustomLists[$distGroupId].ledare
            }
            foreach ($member in $listData)
            {
                if ($null -eq $member)
                {
                    Write-SNSLog -Level "Warn" "Maillist $distGroupName contained a null value."
                    continue
                }
                if (-not $AllMailAddresses.ContainsKey($member))
                {
                    Write-SNSLog -Level "Warn" "Id '$member' from $distGroupName not found in list of all members."
                    continue
                }

                $MemberData = $AllMailAddresses[$member]
                $displayName = "$($MemberData.first_name) $($MemberData.last_name)"

                if ($MailListSettings[$distGroupName].ignore_user)
                {
                    if ($MailListSettings[$distGroupName].ignore_user.Contains($member))
                    {
                        Write-SNSLog "Ignoring '$($displayName)' with Id '$member' for $distGroupName."
                        continue
                    }
                }

                $DoAddLedareScoutnetAddress = $AddLedareScoutnetAddress
                if ($AddLedareOffice365Address)
                {
                    $result, $PrimarySmtpAddress = Add-Office365userToList -allOffice365Users $allOffice365Users -Member $member -MemberData $MemberData -distGroupName $distGroupName
                    if ($result)
                    {
                        if ([string]::IsNullOrWhiteSpace($PrimarySmtpAddress))
                        {
                            Write-SNSLog -Level "Warn" "Primary mailaddres for '$member' is empty. Cannot add member to list."
                        }
                        else
                        {
                            $AddressesToAdd[$PrimarySmtpAddress] = $PrimarySmtpAddress
                        }
                    }
                    else
                    {
                        if ($AddLedareOffice365AddressTryFirst)
                        {
                            $DoAddLedareScoutnetAddress = $True
                        }
                    }
                }

                if ($DoAddLedareScoutnetAddress)
                {
                    if ([string]::IsNullOrWhiteSpace($MemberData.primary_email))
                    {
                        Write-SNSLog -Level "Warn" "Primary mailaddres in scoutnet for '$displayName' is empty. Cannot add member to list."
                    }
                    else
                    {
                        $ContactCreated = New-SNSMailContact -Epost $MemberData.primary_email -DisplayName $displayName `
                            -ContactCache $mailContactCache -ContactInfoCache $mailContactInfoCache -RecipientIndex $recipientIndex `
                            -TimingData $contactTimingByGroup -GroupName $distGroupName

                        if ($ContactCreated)
                        {
                            Write-SNSLog "Adding Contact '$($MemberData.primary_email)' to '$distGroupName'"
                            $AddressesToAdd[$MemberData.primary_email] = $MemberData.primary_email
                        }
                        $mailListsToProcessMembers.Remove($MemberData.primary_email)
                    }
                }
            }

#endregion


#region Add all mailaddresses listed in email_addresses for the maillist.
            foreach ($email in $MailListSettings[$distGroupName].email_addresses)
            {
                if ([string]::IsNullOrWhiteSpace($email))
                {
                    continue
                }

                $ContactCreated = New-SNSMailContact -Epost $email -DisplayName $email `
                    -ContactCache $mailContactCache -ContactInfoCache $mailContactInfoCache -RecipientIndex $recipientIndex `
                    -TimingData $contactTimingByGroup -GroupName $distGroupName

                if ($ContactCreated)
                {
                    Write-SNSLog "Adding Contact '$email' to '$distGroupName'"
                    $AddressesToAdd[$email] = $email
                }
                $mailListsToProcessMembers.Remove($email)
            }
#endregion

#region Update maillist
            # Create a missing group when the rule asks for it: a mail-enabled security group or a regular distribution list.
            $groupType = Get-SNSMailListGroupType -Setting $MailListSettings[$distGroupName]
            if ($groupType -ne '')
            {
                $exchangeType = if ($groupType -eq 'security') { 'Security' } else { 'Distribution' }
                $exchangeTypeDetails = if ($groupType -eq 'security') { 'MailUniversalSecurityGroup' } else { 'MailUniversalDistributionGroup' }
                $existingGroup = Get-EXORecipient -Identity $distGroupName -ErrorAction "SilentlyContinue" -Verbose:$false
                if ($null -eq $existingGroup)
                {
                    $newGroup = @{
                        Name = $distGroupName
                        Type = $exchangeType
                        ErrorAction = "stop"
                        Verbose = $false
                    }
                    # The rule key is the e-mail address of the group, so the display name is configured separately.
                    $displayName = [string]$MailListSettings[$distGroupName].display_name
                    $alias = [string]$MailListSettings[$distGroupName].mail_nickname
                    $smtpAddress = [string]$MailListSettings[$distGroupName].primary_smtp_address
                    $missingFields = @(
                        if ([string]::IsNullOrWhiteSpace($displayName)) { 'display_name' }
                        if ([string]::IsNullOrWhiteSpace($alias)) { 'mail_nickname' }
                        if ([string]::IsNullOrWhiteSpace($smtpAddress)) { 'primary_smtp_address' }
                    )
                    if ($missingFields.Count -gt 0)
                    {
                        Write-SNSLog -Level "Warn" "Group $distGroupName does not exist and cannot be created because $($missingFields -join ', ') is not specified. The group is skipped."
                        Write-SNSTimingLog -Enabled $SNSConf.EnableTimingLog -Phase "Distribution group '$distGroupName'" -Stopwatch $groupStopwatch
                        continue
                    }
                    $newGroup.DisplayName = $displayName
                    $newGroup.Alias = $alias
                    $newGroup.PrimarySmtpAddress = $smtpAddress
                    try
                    {
                        if ($PSCmdlet.ShouldProcess($distGroupName, "New-DistributionGroup"))
                        {
                            Write-SNSLog "Creating $exchangeType group $distGroupName"
                            New-DistributionGroup @newGroup -Confirm:$false | Out-Null

                            # Set-DistributionGroup is used since New-DistributionGroup has no parameter for this.
                            if ((Get-SNMicrosoft365Value -Object $MailListSettings[$distGroupName] -Name 'allow_external_senders') -eq $true)
                            {
                                for ($attempt = 1; $attempt -le 5; $attempt++)
                                {
                                    try
                                    {
                                        Set-DistributionGroup -Identity $distGroupName -RequireSenderAuthenticationEnabled $false -ErrorAction "stop" -Verbose:$false -Confirm:$false
                                        Write-SNSLog "Allowed external senders for $distGroupName"
                                        break
                                    }
                                    catch
                                    {
                                        if ($attempt -eq 5)
                                        {
                                            Write-SNSLog -Level "Warn" "Could not allow external senders for $distGroupName. Error $_"
                                        }
                                        else
                                        {
                                            # A new group can take a while to be visible.
                                            Start-Sleep -Seconds 3
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        Write-SNSLog -Level "Error" "Creation of $exchangeType group $distGroupName failed! Error $_"
                        Write-SNSTimingLog -Enabled $SNSConf.EnableTimingLog -Phase "Distribution group '$distGroupName'" -Stopwatch $groupStopwatch
                        continue
                    }
                }
                elseif ($existingGroup.RecipientTypeDetails -ne $exchangeTypeDetails)
                {
                    Write-SNSLog -Level "Error" "Group $distGroupName exists but is not of type $exchangeTypeDetails. It was not changed."
                    Write-SNSTimingLog -Enabled $SNSConf.EnableTimingLog -Phase "Distribution group '$distGroupName'" -Stopwatch $groupStopwatch
                    continue
                }
            }
            if ((Get-SNMicrosoft365Value -Object $MailListSettings[$distGroupName] -Name 'add_prefix') -eq $true)
            {
                $prefixRuleStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                Set-SNSMailListPrefixRule -GroupName $distGroupName -Setting $MailListSettings[$distGroupName]
                Write-SNSTimingLog -Enabled $SNSConf.EnableTimingLog -Phase "Update mail flow rule for '$distGroupName'" -Stopwatch $prefixRuleStopwatch
            }
            $UpdateDistGroup = @{
                Identity = $distGroupName
                Members =  $AddressesToAdd.Values
                ErrorAction = "stop"
                Verbose = $false
            }
            $groupMembershipStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $currentMemberKeys = $null
            if ($currentGroupMemberKeys -and $currentGroupMemberKeys.ContainsKey([string]$distGroupName))
            {
                $currentMemberKeys = $currentGroupMemberKeys[[string]$distGroupName]
            }
            try
            {
                $memberDifferences = [System.Collections.Generic.List[string]]::new()
                if (Test-SNSGroupMembersUnchanged -CurrentMemberKeys $currentMemberKeys -DesiredAddresses $AddressesToAdd.Values -Differences $memberDifferences)
                {
                    Write-SNSLog "Members of distribution group $distGroupName are unchanged. Update skipped."
                }
                elseif ($PSCmdlet.ShouldProcess($distGroupName, "Update-DistributionGroupMember"))
                {
                    if ($SNSConf.EnableTimingLog)
                    {
                        Write-SNSLog "Distribution group $distGroupName will be updated: $($memberDifferences -join '; ')"
                    }
                    Update-DistributionGroupMember @UpdateDistGroup -Confirm:$false
                }
            }
            catch
            {
                Write-SNSLog -Level "Error" "Update of distribution group members in $distGroupName failed! Error $_"
            }
            Write-SNSTimingLog -Enabled $SNSConf.EnableTimingLog -Phase "Update Exchange group members for '$distGroupName'" -Stopwatch $groupMembershipStopwatch
            $AddressesInMailLists[$distGroupName] = $AddressesToAdd.Values
            $contactTiming = $contactTimingByGroup[$distGroupName]
            if ($null -eq $contactTiming)
            {
                $contactTiming = @{
                    RecipientLookupSeconds = 0.0
                    RecipientLookupCount = 0
                    ContactManagementSeconds = 0.0
                    ContactManagementCount = 0
                    GetContactSeconds = 0.0
                    GetContactCount = 0
                    NewMailContactSeconds = 0.0
                    NewMailContactCount = 0
                    ProvisioningWaitSeconds = 0.0
                    ProvisioningWaitCount = 0
                    SetContactSeconds = 0.0
                    SetContactCount = 0
                    SetMailContactSeconds = 0.0
                    SetMailContactCount = 0
                    ContactCacheHits = 0
                    RecipientIndexHits = 0
                }
            }
            if ($SNSConf.EnableTimingLog)
            {
                Write-SNSLog ("Timing: Contact work for '{0}': {1} uncached recipient lookups in {2:N2}s; {3} cache hits; {4} contact management operations in {5:N2}s; {6} recipient index hits" -f `
                    $distGroupName, $contactTiming.RecipientLookupCount, $contactTiming.RecipientLookupSeconds, `
                    $contactTiming.ContactCacheHits, $contactTiming.ContactManagementCount, $contactTiming.ContactManagementSeconds, `
                    $contactTiming.RecipientIndexHits)
                Write-SNSLog ("Timing: Contact commands for '{0}': Get-Contact {1} in {2:N2}s; New-MailContact {3} in {4:N2}s; provisioning waits {5} in {6:N2}s; Set-Contact {7} in {8:N2}s; Set-MailContact {9} in {10:N2}s" -f `
                    $distGroupName, $contactTiming.GetContactCount, $contactTiming.GetContactSeconds, `
                    $contactTiming.NewMailContactCount, $contactTiming.NewMailContactSeconds, `
                    $contactTiming.ProvisioningWaitCount, $contactTiming.ProvisioningWaitSeconds, `
                    $contactTiming.SetContactCount, $contactTiming.SetContactSeconds, `
                    $contactTiming.SetMailContactCount, $contactTiming.SetMailContactSeconds)
            }
            Write-SNSTimingLog -Enabled $SNSConf.EnableTimingLog -Phase "Distribution group '$distGroupName'" -Stopwatch $groupStopwatch
#endregion
        }
        Write-SNSTimingLog -Enabled $SNSConf.EnableTimingLog -Phase 'Build and synchronize distribution group addresses' -Stopwatch $phaseStopwatch

        $phaseStopwatch.Restart()
        Write-SNSLog " "
        Write-SNSLog "Removing old contacts"
        Write-SNSLog "Number of contacts to check for removal $($mailListsToProcessMembers.count)"

        # Delete all contacts that still is in the list of contacts to process.
        $mailListsToProcessMembers.values | ForEach-Object {
            $medlem = $_
            if ($medlem.RecipientType -eq "MailContact")
            {
                # Check this contact is member in any other maillist.
                $IsInmaillistMembers = $otherMailListsMembers[$_.Identity]
                if ($null -eq $IsInmaillistMembers)
                {
                    # Not used in any maillists. Remove the contact. Remove-MailContact handles -WhatIf.
                    Write-SNSLog "Removing MailContact $($medlem.Identity)"
                    if ($PSCmdlet.ShouldProcess($medlem.Identity, "Remove-MailContact"))
                    {
                        Remove-MailContact $medlem.Identity -Confirm:$false -Verbose:$false
                    }
                }
            }
        }
        Write-SNSTimingLog -Enabled $SNSConf.EnableTimingLog -Phase 'Check and remove unused contacts' -Stopwatch $phaseStopwatch

        Write-SNSLog " "
        Write-SNSLog "Update done new hash value is $NewValidationHash"
        $OutValidationHash = $NewValidationHash
    }

    Write-SNSTimingLog -Enabled $SNSConf.EnableTimingLog -Phase 'Total distribution group synchronization' -Stopwatch $totalStopwatch
    if ($ReturnMaildata)
    {
        return $OutValidationHash, $AddressesInMailLists
    }
    else
    {
        return $OutValidationHash
    }
}
