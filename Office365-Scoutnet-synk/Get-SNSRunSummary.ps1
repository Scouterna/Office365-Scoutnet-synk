function Get-SNSRunSummary
{
    <#
    .SYNOPSIS
        Summarizes a synchronization run from the log file.

    .DESCRIPTION
        Reads the log file written by Write-SNSLog and returns a short text summary: status,
        run time, number of errors and warnings, and what the run did for distribution groups,
        contacts, user accounts, Microsoft 365 groups and shared mailboxes. Errors and warnings are listed,
        up to MaxListedMessages of each; the rest are counted.
        The text is meant to be placed first in the mail with the log.
        Call it after the last synchronization step, before the log is sent.

    .PARAMETER Configuration
        The configuration. The log file is taken from its LogFilePath.
        If omitted, the current module configuration is used.

    .PARAMETER LogFilePath
        Path to a log file to summarize. Overrides the path in the configuration.

    .PARAMETER MaxListedMessages
        Maximum number of errors and of warnings to list. The rest are counted. Default is 20.

    .INPUTS
        None. You cannot pipe objects to Get-SNSRunSummary.

    .OUTPUTS
        System.String. The summary as plain text.

    .EXAMPLE
        $summary = Get-SNSRunSummary -Configuration $conf
        $bodyData = $summary + [Environment]::NewLine + (Get-Content -Path $conf.LogFilePath -Raw -Encoding UTF8)

        Puts the summary first in the mail body, followed by the full log.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param
    (
        [Parameter(HelpMessage="The configuration. Default is the current configuration.")]
        $Configuration,

        [Parameter(HelpMessage="Path to the log file to summarize.")]
        [string]$LogFilePath,

        [Parameter(HelpMessage="Maximum number of errors and warnings to list.")]
        [ValidateRange(0, 1000)]
        [int]$MaxListedMessages = 20
    )

    if ([string]::IsNullOrWhiteSpace($LogFilePath))
    {
        if ($null -eq $Configuration)
        {
            $Configuration = $script:SNSConf
        }
        $LogFilePath = $Configuration.LogFilePath
    }

    if ([string]::IsNullOrWhiteSpace($LogFilePath) -or -not (Test-Path -LiteralPath $LogFilePath))
    {
        return "Sammanfattning saknas: loggfilen hittades inte."
    }

    $lineRegex = '^(?<time>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}) (?<level>INFO|WARNING|ERROR): (?<text>.*)$'
    $errors = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()
    $messages = [System.Collections.Generic.List[string]]::new()
    $first = $null
    $last = $null

    foreach ($line in (Get-Content -LiteralPath $LogFilePath -Encoding UTF8))
    {
        $m = [regex]::Match($line, $lineRegex)
        if (-not $m.Success)
        {
            continue
        }

        $time = [datetime]::ParseExact($m.Groups['time'].Value, 'yyyy-MM-dd HH:mm:ss', [cultureinfo]::InvariantCulture)
        if ($null -eq $first) { $first = $time }
        $last = $time

        $text = $m.Groups['text'].Value
        switch ($m.Groups['level'].Value)
        {
            'ERROR'   { $errors.Add($text) }
            'WARNING' { $warnings.Add($text) }
            default   { $messages.Add($text) }
        }
    }

    if ($null -eq $first)
    {
        return "Sammanfattning saknas: loggfilen innehåller inga loggrader."
    }

    $count = {
        param([string]$Pattern)
        @($messages | Where-Object { $_ -match $Pattern }).Count
    }

    $lastNumber = {
        param([string]$Pattern)
        $found = @($messages | ForEach-Object { [regex]::Match($_, $Pattern) } | Where-Object { $_.Success })
        if ($found.Count -eq 0) { return $null }
        [int]$found[-1].Groups[1].Value
    }

    $groupsProcessed = & $count '^Adding Contacts in distribution group '
    $groupsUnchanged = & $count '^Members of distribution group .* are unchanged\. Update skipped\.$'
    $groupsCreated = & $count '^Creating (distribution|security|\w+) group '
    $contactsCreated = & $count '^Creating Contact '
    $contactsRemoved = & $count '^Removing MailContact '
    $membersAdded = & $count "^Adding Contact '"
    $usersCreated = & $count "^User '.*' added for member id "
    $usersUpdated = & $count "^User '.*' uppdated with new info from Scoutnet\.$"
    $usersDisabled = & $count "^Disabling user "
    $usersEnabled = & $count "^Enabling user "
    $m365UpToDate = & $count "^Microsoft 365 group '.*' is up to date"
    $teamsCreated = & $count "^Created Microsoft Teams team for group "
    $mailboxesCreated = & $count "^Created shared mailbox '"
    $mailboxesUpToDate = & $count "^Shared mailbox '.*' is up to date"
    $mailboxPermissionsGiven = & $count "^Gave .* permission on shared mailbox '"
    $mailboxPermissionsRemoved = & $count "^Removed .* permission for .* on shared mailbox '"
    $reportSent = & $count '^Mandatory moment report sent to '
    $scoutnetUnchanged = & $count '^Scoutnet is not updated\. No update will be done\.$'

    if ($errors.Count -gt 0)
    {
        $status = "FEL: körningen har $($errors.Count) fel"
    }
    elseif ($warnings.Count -gt 0)
    {
        $status = "VARNINGAR: körningen har $($warnings.Count) varningar"
    }
    else
    {
        $status = "OK: inga fel eller varningar"
    }

    $duration = $last - $first
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("Sammanfattning av körningen")
    $lines.Add("===========================")
    $lines.Add("Status:  $status")
    $lines.Add("Start:   $($first.ToString('yyyy-MM-dd HH:mm:ss'))")
    $lines.Add("Slut:    $($last.ToString('yyyy-MM-dd HH:mm:ss'))")
    $lines.Add("Tid:     $([int][math]::Floor($duration.TotalMinutes)) min $($duration.Seconds) s")
    $lines.Add("")

    $lines.Add("Distributionsgrupper")
    if ($scoutnetUnchanged -gt 0)
    {
        $lines.Add("  Scoutnet var oförändrat sedan förra körningen. Inga grupper uppdaterades.")
    }
    else
    {
        $lines.Add("  Behandlade grupper:      $groupsProcessed")
        $lines.Add("  Oförändrade (hoppade över): $groupsUnchanged")
        $lines.Add("  Skapade grupper:         $groupsCreated")
        $lines.Add("  Kontakter skapade:       $contactsCreated")
        $lines.Add("  Kontakter borttagna:     $contactsRemoved")
        $lines.Add("  Medlemmar i listor:      $membersAdded")
    }
    $lines.Add("")

    $lines.Add("Användarkonton")
    $toCreate = & $lastNumber '^Number of accounts to create: (\d+)'
    $toActivate = & $lastNumber '^Number of accounts to activate: (\d+)'
    $toDisable = & $lastNumber '^Number of accounts to disable: (\d+)'
    if ($null -ne $toCreate) { $lines.Add("  Att skapa:    $toCreate") }
    if ($null -ne $toActivate) { $lines.Add("  Att aktivera: $toActivate") }
    if ($null -ne $toDisable) { $lines.Add("  Att inaktivera: $toDisable") }
    $lines.Add("  Skapade: $usersCreated, uppdaterade: $usersUpdated, inaktiverade: $usersDisabled, aktiverade: $usersEnabled")
    $lines.Add("")

    $lines.Add("Microsoft 365-grupper")
    $lines.Add("  Oförändrade: $m365UpToDate, Teams skapade: $teamsCreated")
    $lines.Add("")

    $lines.Add("Delade mailboxar")
    $lines.Add("  Skapade: $mailboxesCreated, oförändrade: $mailboxesUpToDate")
    $lines.Add("  Rättigheter tillagda: $mailboxPermissionsGiven, borttagna: $mailboxPermissionsRemoved")
    $lines.Add("")
    $lines.Add("Rapport över obligatoriska moment: " + $(if ($reportSent -gt 0) { "skickad" } else { "inte skickad" }))

    $addList = {
        param([string]$Title, $Items, [int]$Limit)
        if ($Items.Count -eq 0) { return }
        $lines.Add("")
        $lines.Add("$Title ($($Items.Count))")
        $unique = @($Items | Group-Object | Sort-Object { $_.Group.Count } -Descending)
        foreach ($item in ($unique | Select-Object -First $Limit))
        {
            $suffix = if ($item.Count -gt 1) { " (x$($item.Count))" } else { "" }
            $lines.Add("  - $($item.Name)$suffix")
        }
        if ($unique.Count -gt $Limit)
        {
            $lines.Add("  ... och $($unique.Count - $Limit) till. Se hela loggen nedan.")
        }
    }

    & $addList "Fel" $errors $MaxListedMessages
    & $addList "Varningar" $warnings $MaxListedMessages

    $lines.Add("")
    $lines.Add("Hela loggen följer nedan.")
    $lines.Add("---------------------------")
    $lines.Add("")

    return ($lines -join [Environment]::NewLine)
}
