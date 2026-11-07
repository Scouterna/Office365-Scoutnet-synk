<#
.SYNOPSIS
    Runs the mandatory moment report and reminders.

.DESCRIPTION
    Runs both the report and reminders by default. The report frequency is the
    shared cadence when both are enabled. A warning is logged when the reminder
    frequency differs. The returned date is the shared last-run date to save as
    MandatoryMomentReportLastSent.

.PARAMETER Configuration
    Configuration used by both operations.

.PARAMETER Memberlist
    Optional Scoutnet member list shared by both operations.

.PARAMETER LastSentValue
    Shared last-run date, normally from MandatoryMomentReportLastSent.

.PARAMETER AsOfDate
    Date used for scheduling and moment calculations. Defaults to today.

.PARAMETER TestEmailAddress
    Redirects reminders to this address. The report is still sent to its configured
    recipients.

.PARAMETER NoReport
    Do not send the report.

.PARAMETER NoReminder
    Do not send reminders.

.OUTPUTS
    System.DateTime
    The shared last-run date when either operation was handled; no output when
    neither operation was due.
#>
function Invoke-SNSMandatoryMomentReportAndReminder
{
    [CmdletBinding()]
    [OutputType([datetime])]
    param (
        [Parameter()]
        $Configuration = (Get-SNSConfiguration),

        [Parameter()]
        [System.Collections.IDictionary]$Memberlist,

        [Parameter()]
        [object]$LastSentValue,

        [Parameter()]
        [datetime]$AsOfDate = [datetime]::Today,

        [Parameter(HelpMessage = 'Test mode for reminders.')]
        [string]$TestEmailAddress,

        [Parameter()]
        [switch]$NoReport,

        [Parameter()]
        [switch]$NoReminder
    )

    if ($NoReport -and $NoReminder)
    {
        return
    }

    $reportFrequency = $Configuration.MandatoryMomentReportFrequency
    $reminderFrequency = if ([string]::IsNullOrWhiteSpace($Configuration.MandatoryMomentReminderFrequency)) {
        $reportFrequency
    }
    else {
        $Configuration.MandatoryMomentReminderFrequency
    }

    if (-not $NoReport -and -not $NoReminder)
    {
        if ($reportFrequency -ne $reminderFrequency)
        {
            Write-SNSLog -Level 'Warn' "Mandatory moment report frequency ('$reportFrequency') differs from reminder frequency ('$reminderFrequency'); both will use the report frequency because they share one last-run date."
        }
        $frequency = $reportFrequency
    }
    elseif ($NoReport)
    {
        $frequency = $reminderFrequency
    }
    else
    {
        $frequency = $reportFrequency
    }

    $handled = $false
    if (-not $NoReport)
    {
        $reportHandled = Send-SNSMandatoryMomentReport `
            -Configuration $Configuration `
            -Memberlist $Memberlist `
            -LastSentValue $LastSentValue `
            -AsOfDate $AsOfDate `
            -FrequencyOverride $frequency
        $handled = $handled -or $reportHandled
    }

    if (-not $NoReminder)
    {
        $reminderHandled = Send-SNSMandatoryMomentReminder `
            -Configuration $Configuration `
            -Memberlist $Memberlist `
            -LastSentValue $LastSentValue `
            -AsOfDate $AsOfDate `
            -TestEmailAddress $TestEmailAddress `
            -FrequencyOverride $frequency
        $handled = $handled -or $reminderHandled
    }

    if ($handled)
    {
        return $AsOfDate
    }
}
