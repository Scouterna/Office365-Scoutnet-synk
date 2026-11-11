<#
.SYNOPSIS
    Converts a configured reminder link to a text and a URL.

.DESCRIPTION
    A link is either a plain URL or 'Text|URL'. Only http and https URLs are accepted.
    A plain URL is used as its own text.

.PARAMETER Entry
    The configured link.

.OUTPUTS
    System.Collections.Hashtable
    A hashtable with the keys Text and Url.
#>
function ConvertTo-SNSMandatoryMomentReminderLink
{
    param (
        [Parameter(Mandatory)]
        [string]$Entry
    )

    $text = $null
    $url = $Entry.Trim()
    $separator = $url.IndexOf('|')
    if ($separator -ge 0)
    {
        $text = $url.Substring(0, $separator).Trim()
        $url = $url.Substring($separator + 1).Trim()
    }

    $uri = $null
    if (-not [Uri]::TryCreate($url, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -notin @('http', 'https'))
    {
        throw "The reminder link '$Entry' is not valid. Use 'https://...' or 'Text|https://...'."
    }
    if ([string]::IsNullOrWhiteSpace($text))
    {
        $text = $url
    }
    return @{ Text = $text; Url = $url }
}

<#
.SYNOPSIS
    Finds the members who should be reminded about mandatory moments.

.DESCRIPTION
    Returns one object per member that is older than the minimum age, has at least one
    non-excluded group or troop role, has a valid e-mail address in Scoutnet and has at least one
    moment that is missing or expires within DaysBefore days. Members listed as exempt for a moment
    are not reminded about that moment. Members that would be reminded but have no e-mail address
    are logged as a warning.

.PARAMETER Memberlist
    Scoutnet member list returned by Get-SNSApiGroupMemberlist.

.PARAMETER MinimumAge
    Members must be older than this age.

.PARAMETER ValidityYears
    Number of years a completed moment remains valid.

.PARAMETER DaysBefore
    Number of days before the validity ends that the reminders start.

.PARAMETER ExcludedRoles
    Role names that do not qualify a member by themselves.

.PARAMETER CourseExemptMembers
    Member numbers that are never reminded about Trygga moten.

.PARAMETER PoliceCheckExemptMembers
    Member numbers that are never reminded about Registerutdrag.

.PARAMETER AsOfDate
    Date used for age and validity calculations. Defaults to today.

.OUTPUTS
    System.Management.Automation.PSCustomObject
    Objects with MemberNumber, FirstName, LastName, Email and Moments. Each moment has
    Key, Name, Completed (or null) and Expires (or null).
#>
function Get-SNSMandatoryMomentReminderMember
{
    param (
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Memberlist,

        [Parameter(Mandatory)]
        [int]$MinimumAge,

        [Parameter(Mandatory)]
        [int]$ValidityYears,

        [Parameter()]
        [int]$DaysBefore = 0,

        [Parameter()]
        [string[]]$ExcludedRoles = @(),

        [Parameter()]
        [string[]]$CourseExemptMembers = @(),

        [Parameter()]
        [string[]]$PoliceCheckExemptMembers = @(),

        [Parameter()]
        [datetime]$AsOfDate = [datetime]::Today
    )

    if ($null -eq $Memberlist.data -or $Memberlist.data -isnot [System.Collections.IDictionary])
    {
        throw "The member list does not contain the expected 'data' dictionary."
    }

    $clean = { param($List) @($List | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { "$_".Trim() }) }
    $moments = @(
        @{ Field = 'pc_course_89'; Name = 'Trygga möten'; Exempt = (& $clean $CourseExemptMembers) }
        @{ Field = 'latest_police_check'; Name = 'Registerutdrag'; Exempt = (& $clean $PoliceCheckExemptMembers) }
    )
    $remindUntil = $AsOfDate.Date.AddDays($DaysBefore)

    foreach ($memberEntry in $Memberlist.data.GetEnumerator())
    {
        $member = $memberEntry.Value
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
        $groupRoles = @(Get-SNSMandatoryMomentRoleName -RoleData (Get-SNSMandatoryMomentValue -Object $roles -Name 'group'))
        $unitRoles = @(Get-SNSMandatoryMomentRoleName -RoleData (Get-SNSMandatoryMomentValue -Object $roles -Name 'troop'))
        if (@(($groupRoles + $unitRoles) | Where-Object { $_ -notin $ExcludedRoles }).Count -eq 0)
        {
            continue
        }

        $memberNumber = Get-SNSMandatoryMomentMemberNumber -Member $member -Key $memberEntry.Key
        $dueMoments = @(
            foreach ($moment in $moments)
            {
                if ($memberNumber -in $moment.Exempt)
                {
                    continue
                }

                $completed = ConvertTo-SNSMandatoryMomentDate -Value $member[$moment.Field] -FieldName $moment.Field
                $expires = if ($null -eq $completed) { $null } else { $completed.AddYears($ValidityYears) }
                if ($null -eq $expires -or $expires -le $remindUntil)
                {
                    [pscustomobject]@{ Key = $moment.Field; Name = $moment.Name; Completed = $completed; Expires = $expires }
                }
            }
        )
        if ($dueMoments.Count -eq 0)
        {
            continue
        }

        $firstName = [string](Get-SNSMandatoryMomentValue -Object $member.first_name -Name 'value')
        $lastName = [string](Get-SNSMandatoryMomentValue -Object $member.last_name -Name 'value')
        $email = [string](Get-SNSMandatoryMomentValue -Object $member.email -Name 'value')
        if ($email -notmatch '^[^@\s]+@[^@\s]+$')
        {
            $email = [string](Get-SNSMandatoryMomentValue -Object $member.contact_alt_email -Name 'value')
        }
        $email = $email.Trim()
        if ($email -notmatch '^[^@\s]+@[^@\s]+$')
        {
            Write-SNSLog -Level 'Warn' "Mandatory moment reminder not sent to member '$memberNumber' ($firstName $lastName) because the member has no valid e-mail address in Scoutnet."
            continue
        }

        [pscustomobject]@{
            MemberNumber = $memberNumber
            FirstName = $firstName
            LastName = $lastName
            Email = $email
            Moments = $dueMoments
        }
    }
}

<#
.SYNOPSIS
    Creates the HTML body of a mandatory moment reminder.

.DESCRIPTION
    Builds the mail from the texts in the configuration. {FirstName}, {LastName} and {Name} in the
    texts are replaced with the member's name. All text is HTML encoded.

.PARAMETER Reminder
    A member returned by Get-SNSMandatoryMomentReminderMember.

.PARAMETER Configuration
    The configuration with the reminder texts and links.

.PARAMETER AsOfDate
    Date used to tell if a moment has expired. Defaults to today.

.OUTPUTS
    System.String
    The HTML document.
#>
function ConvertTo-SNSMandatoryMomentReminderHtml
{
    param (
        [Parameter(Mandatory)]
        $Reminder,

        [Parameter(Mandatory)]
        $Configuration,

        [Parameter()]
        [datetime]$AsOfDate = [datetime]::Today
    )

    $encode = {
        param([string]$Text)
        if ([string]::IsNullOrEmpty($Text)) { return '' }
        $expanded = $Text.Replace('{FirstName}', $Reminder.FirstName).Replace('{LastName}', $Reminder.LastName).
            Replace('{Name}', "$($Reminder.FirstName) $($Reminder.LastName)".Trim())
        ([System.Net.WebUtility]::HtmlEncode($expanded) -replace "\r?\n", '<br>')
    }

    $html = [System.Text.StringBuilder]::new()
    [void]$html.AppendLine('<!doctype html><html lang="sv"><head><meta charset="utf-8">')
    [void]$html.AppendLine('<style>body{font-family:Arial,sans-serif;color:#222}section{margin-bottom:1.5em}</style></head><body>')
    if (-not [string]::IsNullOrWhiteSpace($Configuration.MandatoryMomentReminderHeading))
    {
        [void]$html.AppendLine("<h1>$(& $encode $Configuration.MandatoryMomentReminderHeading)</h1>")
    }
    if (-not [string]::IsNullOrWhiteSpace($Configuration.MandatoryMomentReminderIntro))
    {
        [void]$html.AppendLine("<p>$(& $encode $Configuration.MandatoryMomentReminderIntro)</p>")
    }

    foreach ($moment in $Reminder.Moments)
    {
        if ($moment.Key -eq 'pc_course_89')
        {
            $text = $Configuration.MandatoryMomentCourseReminderText
            $links = @($Configuration.MandatoryMomentCourseReminderLinks)
        }
        else
        {
            $text = $Configuration.MandatoryMomentPoliceCheckReminderText
            $links = @($Configuration.MandatoryMomentPoliceCheckReminderLinks)
        }

        if ($null -eq $moment.Completed)
        {
            $status = 'Momentet är aldrig utfört.'
        }
        elseif ($moment.Expires -le $AsOfDate.Date)
        {
            $status = "Senast utfört $($moment.Completed.ToString('yyyy-MM-dd')). Giltighetstiden gick ut $($moment.Expires.ToString('yyyy-MM-dd'))."
        }
        else
        {
            $status = "Senast utfört $($moment.Completed.ToString('yyyy-MM-dd')). Giltighetstiden går ut $($moment.Expires.ToString('yyyy-MM-dd'))."
        }

        [void]$html.AppendLine("<section><h2>$(& $encode $moment.Name)</h2>")
        if (-not [string]::IsNullOrWhiteSpace($text))
        {
            [void]$html.AppendLine("<p>$(& $encode $text)</p>")
        }
        [void]$html.AppendLine("<p>$(& $encode $status)</p>")
        $validLinks = @($links | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($validLinks.Count -gt 0)
        {
            [void]$html.AppendLine('<ul>')
            foreach ($entry in $validLinks)
            {
                $link = ConvertTo-SNSMandatoryMomentReminderLink -Entry $entry
                $url = [System.Net.WebUtility]::HtmlEncode($link.Url)
                $linkText = [System.Net.WebUtility]::HtmlEncode($link.Text)
                [void]$html.AppendLine("<li><a href=""$url"">$linkText</a></li>")
            }
            [void]$html.AppendLine('</ul>')
        }
        [void]$html.AppendLine('</section>')
    }

    if (-not [string]::IsNullOrWhiteSpace($Configuration.MandatoryMomentReminderFooter))
    {
        [void]$html.AppendLine("<p>$(& $encode $Configuration.MandatoryMomentReminderFooter)</p>")
    }
    [void]$html.AppendLine('</body></html>')
    return $html.ToString()
}

<#
.SYNOPSIS
    Sends reminders to members who need to do or renew a mandatory moment.

.DESCRIPTION
    Sends one mail to each member who is older than MandatoryMomentMinimumAge, has at least one
    non-excluded role and has a moment (Trygga moten or Registerutdrag) that is missing or expires
    within MandatoryMomentReminderDaysBefore days. The mail goes to the member's e-mail address in
    Scoutnet, with contact_alt_email as fallback, and to MandatoryMomentReminderCc as a copy.
    Members listed in MandatoryMomentCourseExemptMembers or MandatoryMomentPoliceCheckExemptMembers
    are not reminded about that moment.

    Reminders are sent once per
    period, set by MandatoryMomentReminderFrequency. The texts and links in the mail are taken from
    the MandatoryMomentReminder* and Mandatory*ReminderText/Links settings. A failure for one
    member is logged as an error and does not stop the others.

.PARAMETER Configuration
    The configuration. Defaults to the current SNS configuration.

.PARAMETER Memberlist
    Optional Scoutnet member list. When omitted, it is fetched using the member list credential.

.PARAMETER LastSentValue
    Raw value of the MandatoryMomentReminderLastSent Automation variable. An empty, missing or
    invalid value makes the reminders run regardless of the configured frequency.

.PARAMETER AsOfDate
    Date used for frequency, ages and validity. Defaults to today.

.PARAMETER TestEmailAddress
    Test mode. All reminders are sent to this address instead of the members and
    MandatoryMomentReminderCc is not used.

.OUTPUTS
    System.Boolean
    True when the reminders were handled for the period, even if no member needed one. False when
    reminders are disabled or not due.

.EXAMPLE
    $lastSent = Get-AutomationVariable -Name 'MandatoryMomentReminderLastSent'
    if (Send-SNSMandatoryMomentReminder -Configuration $conf -LastSentValue $lastSent) {
        Set-AutomationVariable -Name 'MandatoryMomentReminderLastSent' -Value (Get-Date -Format 'o')
    }
#>
function Send-SNSMandatoryMomentReminder
{
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter()]
        $Configuration = (Get-SNSConfiguration),

        [Parameter()]
        [System.Collections.IDictionary]$Memberlist,

        [Parameter()]
        [object]$LastSentValue,

        [Parameter()]
        [datetime]$AsOfDate = [datetime]::Today,

        [Parameter(HelpMessage = 'Test mode. All reminders are sent to this address instead of the members, and Cc is not used.')]
        [string]$TestEmailAddress,

        [Parameter()]
        [string]$FrequencyOverride
    )

    # Reminder-specific settings fall back to the report/base settings when not configured.
    if (-not [string]::IsNullOrWhiteSpace($FrequencyOverride))
    {
        $frequency = $FrequencyOverride
    }
    else
    {
        $frequency = if ([string]::IsNullOrWhiteSpace($Configuration.MandatoryMomentReminderFrequency)) { $Configuration.MandatoryMomentReportFrequency } else { $Configuration.MandatoryMomentReminderFrequency }
    }
    $minimumAge = if ($Configuration.MandatoryMomentReminderMinimumAge -gt 0) { $Configuration.MandatoryMomentReminderMinimumAge } else { $Configuration.MandatoryMomentMinimumAge }
    $validityYears = if ($Configuration.MandatoryMomentReminderValidityYears -gt 0) { $Configuration.MandatoryMomentReminderValidityYears } else { $Configuration.MandatoryMomentValidityYears }
    $reminderRoles = @($Configuration.MandatoryMomentReminderExcludedRoles | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $excludedRoles = if ($reminderRoles.Count -gt 0) { $reminderRoles } else { $Configuration.MandatoryMomentExcludedRoles }
    $daysBefore = if ($Configuration.MandatoryMomentReminderDaysBefore -ne 0) { $Configuration.MandatoryMomentReminderDaysBefore } else { $Configuration.MandatoryMomentDaysBefore }

    if ($frequency -notin @('Monthly', 'Weekly', 'Daily'))
    {
        throw "MandatoryMomentReminderFrequency (or MandatoryMomentReportFrequency) must be 'Monthly', 'Weekly', or 'Daily'."
    }
    if ($minimumAge -lt 0)
    {
        throw 'MandatoryMomentMinimumAge must be zero or greater.'
    }
    if ($validityYears -lt 1)
    {
        throw 'MandatoryMomentValidityYears must be at least one.'
    }
    if ($daysBefore -lt 0)
    {
        throw 'MandatoryMomentDaysBefore or MandatoryMomentReminderDaysBefore must be zero or greater.'
    }
    if ([string]::IsNullOrWhiteSpace($Configuration.EmailFromAddress))
    {
        throw 'EmailFromAddress must be configured to send the mandatory moment reminders.'
    }
    foreach ($entry in @($Configuration.MandatoryMomentCourseReminderLinks) + @($Configuration.MandatoryMomentPoliceCheckReminderLinks))
    {
        if (-not [string]::IsNullOrWhiteSpace($entry))
        {
            [void](ConvertTo-SNSMandatoryMomentReminderLink -Entry $entry)
        }
    }

    $lastSentDate = [datetime]::MinValue
    $forceSend = $false
    if ($null -eq $LastSentValue -or [string]::IsNullOrWhiteSpace([string]$LastSentValue))
    {
        Write-SNSLog 'No previous reminder date available; the reminders are sent regardless of the configured frequency.'
        $forceSend = $true
    }
    elseif (-not [datetime]::TryParse(
        [string]$LastSentValue,
        [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::RoundtripKind,
        [ref]$lastSentDate
    ))
    {
        Write-SNSLog -Level 'Warn' "Automation variable 'MandatoryMomentReminderLastSent' does not contain a valid date; the reminders are sent regardless of the configured frequency."
        $lastSentDate = [datetime]::MinValue
        $forceSend = $true
    }

    if (-not $forceSend -and -not (Test-SNSMandatoryMomentReportDue `
        -Frequency $frequency `
        -LastSentDate $lastSentDate `
        -AsOfDate $AsOfDate))
    {
        Write-SNSLog "Mandatory moment reminders are not due ($frequency)."
        return $false
    }

    if ($null -eq $Memberlist)
    {
        $Memberlist = Get-SNSApiGroupMemberlist -Credential $Configuration.CredentialMemberlist
    }

    $reminders = @(Get-SNSMandatoryMomentReminderMember `
        -Memberlist $Memberlist `
        -MinimumAge $minimumAge `
        -ValidityYears $validityYears `
        -DaysBefore $daysBefore `
        -ExcludedRoles $excludedRoles `
        -CourseExemptMembers $Configuration.MandatoryMomentCourseExemptMembers `
        -PoliceCheckExemptMembers $Configuration.MandatoryMomentPoliceCheckExemptMembers `
        -AsOfDate $AsOfDate)

    if (-not [string]::IsNullOrWhiteSpace($TestEmailAddress))
    {
        Write-SNSLog -Level 'Warn' "Test mode: all mandatory moment reminders are sent to '$TestEmailAddress'."
    }

    $ccRecipients = @(
        foreach ($address in @($Configuration.MandatoryMomentReminderCc))
        {
            if (-not [string]::IsNullOrWhiteSpace($TestEmailAddress)) { break }
            if (-not [string]::IsNullOrWhiteSpace($address))
            {
                @{ EmailAddress = @{ Address = $address.Trim() } }
            }
        }
    )

    $subject = if ([string]::IsNullOrWhiteSpace($Configuration.MandatoryMomentReminderSubject)) { 'Påminnelse: obligatoriska moment' } else { $Configuration.MandatoryMomentReminderSubject }
    $sent = 0
    foreach ($reminder in $reminders)
    {
        try
        {
            $message = @{
                Subject = $subject.Replace('{FirstName}', $reminder.FirstName).Replace('{LastName}', $reminder.LastName).
                    Replace('{Name}', "$($reminder.FirstName) $($reminder.LastName)".Trim())
                Body = @{
                    ContentType = 'HTML'
                    Content = ConvertTo-SNSMandatoryMomentReminderHtml -Reminder $reminder -Configuration $Configuration -AsOfDate $AsOfDate
                }
                ToRecipients = @(@{ EmailAddress = @{ Address = $(if ([string]::IsNullOrWhiteSpace($TestEmailAddress)) { $reminder.Email } else { $TestEmailAddress.Trim() }) } })
            }
            if ($ccRecipients.Count -gt 0)
            {
                $message.CcRecipients = $ccRecipients
            }

            Send-MgUserMail -UserId $Configuration.EmailFromAddress -BodyParameter @{ Message = $message; SaveToSentItems = 'false' } -ErrorAction Stop | Out-Null
            $sent++
        }
        catch
        {
            Write-SNSLog -Level 'Error' "Could not send mandatory moment reminder to member '$($reminder.MemberNumber)'. Error: $_"
        }
    }

    Write-SNSLog "Mandatory moment reminders sent: $sent."
    return $true
}
