<#
.SYNOPSIS
    Gets a named value from a dictionary or object.

.DESCRIPTION
    Reads a key from an IDictionary or a property from another PowerShell object.
    Returns null when the object is null or does not contain the requested name.

.PARAMETER Object
    Dictionary or object to read from.

.PARAMETER Name
    Key or property name to retrieve.

.OUTPUTS
    System.Object
    The value associated with the name, or null when it is not present.
#>
function Get-SNSMandatoryMomentValue
{
    param (
        [Parameter()]
        $Object,

        [Parameter(Mandatory)]
        [string]$Name
    )

    if ($null -eq $Object)
    {
        return $null
    }

    if ($Object -is [System.Collections.IDictionary])
    {
        if ($Object.Contains($Name))
        {
            return $Object[$Name]
        }
        return $null
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -ne $property)
    {
        return $property.Value
    }
    return $null
}

<#
.SYNOPSIS
    Gets role names from nested Scoutnet role data.

.DESCRIPTION
    Recursively searches dictionaries and collections for role_name values and
    writes each discovered role name to the pipeline.

.PARAMETER RoleData
    Nested role information from a Scoutnet member record.

.OUTPUTS
    System.String
    The role names found in the supplied role data.
#>
function Get-SNSMandatoryMomentRoleName
{
    param (
        [Parameter()]
        $RoleData
    )

    if ($null -eq $RoleData)
    {
        return
    }

    $roleName = Get-SNSMandatoryMomentValue -Object $RoleData -Name 'role_name'
    if (-not [string]::IsNullOrWhiteSpace([string]$roleName))
    {
        Write-Output $roleName
        return
    }

    if ($RoleData -is [System.Collections.IDictionary])
    {
        foreach ($value in $RoleData.Values)
        {
            Get-SNSMandatoryMomentRoleName -RoleData $value
        }
    }
    elseif ($RoleData -is [System.Collections.IEnumerable] -and $RoleData -isnot [string])
    {
        foreach ($value in $RoleData)
        {
            Get-SNSMandatoryMomentRoleName -RoleData $value
        }
    }
}

<#
.SYNOPSIS
    Converts a Scoutnet date field to a date-only value.

.DESCRIPTION
    Reads the value member of a Scoutnet field object and parses it using
    invariant culture. Empty values return null. Invalid dates cause an error
    that identifies the member field.

.PARAMETER Value
    Scoutnet field object containing a value property.

.PARAMETER FieldName
    Member field name, used in an error if the date cannot be parsed.

.OUTPUTS
    System.DateTime
    The parsed date with its time component removed, or null for an empty value.

.EXAMPLE
    ConvertTo-SNSMandatoryMomentDate -Value @{ value = '2025-03-01' } -FieldName 'pc_course_89'
#>
function ConvertTo-SNSMandatoryMomentDate
{
    param (
        [Parameter()]
        $Value,

        [Parameter(Mandatory)]
        [string]$FieldName
    )

    $dateValue = Get-SNSMandatoryMomentValue -Object $Value -Name 'value'
    if ([string]::IsNullOrWhiteSpace([string]$dateValue))
    {
        return $null
    }

    $date = [datetime]::MinValue
    $parsed = [datetime]::TryParse(
        [string]$dateValue,
        [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::None,
        [ref]$date
    )
    if (-not $parsed)
    {
        throw "Could not parse '$dateValue' in member field '$FieldName' as a date."
    }
    return $date.Date
}

<#
.SYNOPSIS
    Creates the HTML report for Scoutnet mandatory moments.

.DESCRIPTION
    Builds separate tables for Trygga möten and Registerutdrag. Includes members
    older than the configured minimum age who hold at least one non-excluded
    group or troop role and whose moment date is missing or older than the
    configured validity period. Each table is preceded by an explanation of
    these inclusion criteria.

.PARAMETER Memberlist
    Scoutnet member list returned by Get-SNSApiGroupMemberlist.

.PARAMETER MinimumAge
    Minimum age threshold. Members must be older than this age.

.PARAMETER ValidityYears
    Number of years a completed mandatory moment remains valid.

.PARAMETER CourseDescription
    Configured description displayed above the Trygga möten table.

.PARAMETER PoliceCheckDescription
    Configured description displayed above the Registerutdrag table.

.PARAMETER ExcludedRoles
    Role names that do not qualify a member for inclusion in the report. A
    member with at least one other role is still included.

.PARAMETER AsOfDate
    Date used for age and validity calculations. Defaults to today.

.OUTPUTS
    System.String
    Complete HTML document for the mandatory moment report.

.EXAMPLE
    $html = ConvertTo-SNSMandatoryMomentReportHtml -Memberlist $memberlist `
        -MinimumAge 15 -ValidityYears 3 `
        -CourseDescription $conf.MandatoryMomentCourseDescription `
        -PoliceCheckDescription $conf.MandatoryMomentPoliceCheckDescription
#>
function ConvertTo-SNSMandatoryMomentReportHtml
{
    param (
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Memberlist,

        [Parameter(Mandatory)]
        [int]$MinimumAge,

        [Parameter(Mandatory)]
        [int]$ValidityYears,

        [Parameter()]
        [string]$CourseDescription,

        [Parameter()]
        [string]$PoliceCheckDescription,

        [Parameter()]
        [string[]]$ExcludedRoles = @(),

        [Parameter()]
        [datetime]$AsOfDate = [datetime]::Today
    )

    if ($null -eq $Memberlist.data -or $Memberlist.data -isnot [System.Collections.IDictionary])
    {
        throw "The member list does not contain the expected 'data' dictionary."
    }

    $members = @($Memberlist.data.Values)
    $unitNamesById = @{}
    foreach ($member in $members)
    {
        $unit = Get-SNSMandatoryMomentValue -Object $member -Name 'unit'
        $unitId = Get-SNSMandatoryMomentValue -Object $unit -Name 'raw_value'
        $unitName = Get-SNSMandatoryMomentValue -Object $unit -Name 'value'
        if (-not [string]::IsNullOrWhiteSpace([string]$unitId) -and
            -not [string]::IsNullOrWhiteSpace([string]$unitName))
        {
            $unitNamesById[[string]$unitId] = [string]$unitName
        }
    }

    $cutoffDate = $AsOfDate.Date.AddYears(-$ValidityYears)
    $moments = @(
        @{ Field = 'pc_course_89'; Name = 'Trygga möten'; Description = $CourseDescription }
        @{ Field = 'latest_police_check'; Name = 'Registerutdrag'; Description = $PoliceCheckDescription }
    )

    $html = [System.Text.StringBuilder]::new()
    [void]$html.AppendLine('<!doctype html><html lang="sv"><head><meta charset="utf-8">')
    [void]$html.AppendLine('<meta name="viewport" content="width=device-width, initial-scale=1">')
    [void]$html.AppendLine('<style>body{font-family:Arial,sans-serif;color:#222}table{border-collapse:collapse;width:100%;margin-bottom:2em}th,td{border:1px solid #bbb;padding:6px;text-align:left}th{background:#eee}section{margin-bottom:2em}</style>')
    [void]$html.AppendLine('</head><body>')
    $inclusionReason = "En medlem är med i listan om medlemmen är äldre än $MinimumAge år, har minst en kår- eller avdelningsroll och det var minst $ValidityYears år sedan momentet gjordes."
    $inclusionReason = [System.Net.WebUtility]::HtmlEncode($inclusionReason)
    [void]$html.AppendLine("<p>$inclusionReason</p>")

    foreach ($moment in $moments)
    {
        $description = [System.Net.WebUtility]::HtmlEncode($moment.Description)
        [void]$html.AppendLine("<section><h2>$($moment.Name)</h2><p>$description</p>")
        [void]$html.AppendLine('<table><thead><tr><th>Efternamn</th><th>Förnamn</th><th>Datum när momentet gjordes</th><th>Roller i kåren</th><th>Roller på avdelning</th></tr></thead><tbody>')

        foreach ($member in $members)
        {
            $birthDate = ConvertTo-SNSMandatoryMomentDate -Value $member.date_of_birth -FieldName 'date_of_birth'
            if ($null -eq $birthDate)
            {
                continue
            }

            $age = $AsOfDate.Year - $birthDate.Year
            if ($birthDate.Date -gt $AsOfDate.Date.AddYears(-$age))
            {
                $age--
            }
            if ($age -le $MinimumAge)
            {
                continue
            }

            $roles = Get-SNSMandatoryMomentValue -Object $member.roles -Name 'value'
            $groupRoles = @(Get-SNSMandatoryMomentRoleName -RoleData (Get-SNSMandatoryMomentValue -Object $roles -Name 'group') | Sort-Object -Unique)
            $troopRoles = Get-SNSMandatoryMomentValue -Object $roles -Name 'troop'
            $unitRoleDescriptions = @(
                if ($troopRoles -is [System.Collections.IDictionary])
                {
                    foreach ($troopEntry in $troopRoles.GetEnumerator())
                    {
                        $troopRoleNames = @(Get-SNSMandatoryMomentRoleName -RoleData $troopEntry.Value | Sort-Object -Unique)
                        if ($troopRoleNames.Count -gt 0)
                        {
                            $troopId = [string]$troopEntry.Key
                            $troopName = $unitNamesById[$troopId]
                            if ([string]::IsNullOrWhiteSpace($troopName))
                            {
                                $troopName = "Avdelning $troopId"
                            }
                            "$troopName`: $($troopRoleNames -join ', ')"
                        }
                    }
                }
            )
            $unitRoles = @(Get-SNSMandatoryMomentRoleName -RoleData $troopRoles | Sort-Object -Unique)
            $qualifyingGroupRoles = @($groupRoles | Where-Object { $_ -notin $ExcludedRoles })
            $qualifyingUnitRoles = @($unitRoles | Where-Object { $_ -notin $ExcludedRoles })
            if ($qualifyingGroupRoles.Count -eq 0 -and $qualifyingUnitRoles.Count -eq 0)
            {
                continue
            }

            $completedDate = ConvertTo-SNSMandatoryMomentDate -Value $member[$moment.Field] -FieldName $moment.Field
            if ($null -ne $completedDate -and $completedDate -gt $cutoffDate)
            {
                continue
            }

            $lastName = [System.Net.WebUtility]::HtmlEncode([string](Get-SNSMandatoryMomentValue -Object $member.last_name -Name 'value'))
            $firstName = [System.Net.WebUtility]::HtmlEncode([string](Get-SNSMandatoryMomentValue -Object $member.first_name -Name 'value'))
            $dateText = if ($null -eq $completedDate) { 'Aldrig utfört' } else { $completedDate.ToString('yyyy-MM-dd') }
            $dateText = [System.Net.WebUtility]::HtmlEncode($dateText)
            $groupRoleText = [System.Net.WebUtility]::HtmlEncode(($groupRoles -join ', '))
            $unitRoleText = [System.Net.WebUtility]::HtmlEncode(($unitRoleDescriptions -join '; '))
            [void]$html.AppendLine("<tr><td>$lastName</td><td>$firstName</td><td>$dateText</td><td>$groupRoleText</td><td>$unitRoleText</td></tr>")
        }

        [void]$html.AppendLine('</tbody></table></section>')
    }

    [void]$html.AppendLine('</body></html>')
    return $html.ToString()
}

<#
.SYNOPSIS
    Determines whether the mandatory moment report is due.

.DESCRIPTION
    Checks whether the configured daily, weekly, or monthly reporting period has
    started and whether a report has already been sent in that period.

.PARAMETER Frequency
    Reporting cadence: Daily, Weekly (Monday), or Monthly (the first day).

.PARAMETER LastSentDate
    Date and time of the last successful report. DateTime.MinValue means no
    report has previously been sent.

.PARAMETER AsOfDate
    Date used to determine the current reporting period. Defaults to today.

.OUTPUTS
    System.Boolean
    True when a report is due; otherwise, false.

.EXAMPLE
    Test-SNSMandatoryMomentReportDue -Frequency Monthly -AsOfDate (Get-Date)
#>
function Test-SNSMandatoryMomentReportDue
{
    param (
        [Parameter(Mandatory)]
        [ValidateSet('Monthly', 'Weekly', 'Daily')]
        [string]$Frequency,

        [Parameter()]
        [datetime]$LastSentDate = [datetime]::MinValue,

        [Parameter()]
        [datetime]$AsOfDate = [datetime]::Today
    )

    $today = $AsOfDate.Date
    switch ($Frequency)
    {
        'Monthly' {
            return $today.Day -eq 1 -and
                ($LastSentDate -eq [datetime]::MinValue -or
                    $LastSentDate.Year -ne $today.Year -or $LastSentDate.Month -ne $today.Month)
        }
        'Weekly' {
            $weekStart = $today.AddDays(-(([int]$today.DayOfWeek + 6) % 7))
            $lastSentWeekStart = $LastSentDate.Date.AddDays(-(([int]$LastSentDate.DayOfWeek + 6) % 7))
            return $today.DayOfWeek -eq [DayOfWeek]::Monday -and
                ($LastSentDate -eq [datetime]::MinValue -or $lastSentWeekStart -ne $weekStart)
        }
        'Daily' {
            return $LastSentDate.Date -lt $today
        }
    }
}

<#
.SYNOPSIS
    Creates and emails the mandatory moment report.

.DESCRIPTION
    Validates the report configuration and saved last-sent value, checks whether
    the configured reporting period is due, generates the HTML report, and sends
    it to the configured recipients using Microsoft Graph. Returns false when
    the report is not due or there are no recipients; returns true after a
    successful send. If Memberlist is not supplied, it is fetched from Scoutnet.

.PARAMETER Configuration
    Report configuration. Defaults to the current SNS configuration.

.PARAMETER Memberlist
    Optional Scoutnet member list. When omitted, the member list is fetched using
    the configured member-list credential.

.PARAMETER LastSentValue
    Raw value of the MandatoryMomentReportLastSent Automation variable. Empty
    means that no prior report date is available.

.PARAMETER AsOfDate
    Date used for reporting cadence, member ages, and moment validity. Defaults
    to today.

.OUTPUTS
    System.Boolean
    True when the report is sent; false when no send is due or no recipients are
    configured.

.EXAMPLE
    $lastSent = Get-AutomationVariable -Name 'MandatoryMomentReportLastSent'
    $sent = Send-SNSMandatoryMomentReport -Configuration $conf -LastSentValue $lastSent
    if ($sent) {
        Set-AutomationVariable -Name 'MandatoryMomentReportLastSent' -Value (Get-Date -Format 'o')
    }
#>
function Send-SNSMandatoryMomentReport
{
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter()]
        [SNSConfiguration]$Configuration = (Get-SNSConfiguration),

        [Parameter()]
        [System.Collections.IDictionary]$Memberlist,

        [Parameter()]
        [object]$LastSentValue,

        [Parameter()]
        [datetime]$AsOfDate = [datetime]::Today
    )

    $lastSentDate = [datetime]::MinValue
    if ($null -ne $LastSentValue -and
        -not [string]::IsNullOrWhiteSpace([string]$LastSentValue))
    {
        $parsedLastSentDate = [datetime]::MinValue
        if (-not [datetime]::TryParse(
            [string]$LastSentValue,
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::RoundtripKind,
            [ref]$parsedLastSentDate
        ))
        {
            Write-SNSLog -Level "Warn" "Automation variable 'MandatoryMomentReportLastSent' does not contain a valid date; using today's date instead."
            $lastSentDate = $AsOfDate.Date
        }
        else
        {
            $lastSentDate = $parsedLastSentDate
        }
    }

    if ($Configuration.MandatoryMomentReportFrequency -notin @('Monthly', 'Weekly', 'Daily'))
    {
        throw "MandatoryMomentReportFrequency must be 'Monthly', 'Weekly', or 'Daily'."
    }
    if ($Configuration.MandatoryMomentMinimumAge -lt 0)
    {
        throw 'MandatoryMomentMinimumAge must be zero or greater.'
    }
    if ($Configuration.MandatoryMomentValidityYears -lt 1)
    {
        throw 'MandatoryMomentValidityYears must be at least one.'
    }

    if (-not (Test-SNSMandatoryMomentReportDue `
        -Frequency $Configuration.MandatoryMomentReportFrequency `
        -LastSentDate $lastSentDate `
        -AsOfDate $AsOfDate))
    {
        Write-SNSLog "Mandatory moment report is not due ($($Configuration.MandatoryMomentReportFrequency))."
        return $false
    }

    $recipients = @($Configuration.MandatoryMomentReportRecipients |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($recipients.Count -eq 0)
    {
        Write-SNSLog 'Mandatory moment report not sent because no recipients are configured.'
        return $false
    }
    if ([string]::IsNullOrWhiteSpace($Configuration.EmailFromAddress))
    {
        throw 'EmailFromAddress must be configured to send the mandatory moment report.'
    }

    if ($null -eq $Memberlist)
    {
        $Memberlist = Get-SNSApiGroupMemberlist -Credential $Configuration.CredentialMemberlist
    }

    $html = ConvertTo-SNSMandatoryMomentReportHtml `
        -Memberlist $Memberlist `
        -MinimumAge $Configuration.MandatoryMomentMinimumAge `
        -ValidityYears $Configuration.MandatoryMomentValidityYears `
        -CourseDescription $Configuration.MandatoryMomentCourseDescription `
        -PoliceCheckDescription $Configuration.MandatoryMomentPoliceCheckDescription `
        -ExcludedRoles $Configuration.MandatoryMomentExcludedRoles `
        -AsOfDate $AsOfDate

    $params = @{
        Message = @{
            Subject = 'Rapport: obligatoriska moment'
            Body = @{
                ContentType = 'HTML'
                Content = $html
            }
            ToRecipients = @(
                foreach ($recipient in $recipients)
                {
                    @{
                        EmailAddress = @{
                            Address = $recipient
                        }
                    }
                }
            )
        }
        SaveToSentItems = 'false'
    }

    Send-MgUserMail -UserId $Configuration.EmailFromAddress -BodyParameter $params -ErrorAction Stop | Out-Null
    Write-SNSLog "Mandatory moment report sent to $($recipients -join ', ')."
    return $true
}
