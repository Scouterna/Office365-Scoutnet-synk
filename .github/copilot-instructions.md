#### Project Overview

Office365-Scoutnet-synk (v3.2.0) is a PowerShell 7.6+ module that synchronizes Scoutnet members, leader roles, and mailing lists with Microsoft 365 user accounts, distribution lists, Microsoft 365 groups, Teams, and shared mailboxes, and sends compliance reports/reminders for mandatory requirements (*Trygga möten* and *registerutdrag*). Scripts run both locally and as Azure Automation Runbooks (using Managed Identity). 

#### Runtime & Exact Module Dependencies

* **PowerShell:** Minimum 7.6 (PowerShellVersion = '7.6'). Never introduce Windows PowerShell 5.1 syntax or legacy MSOnline / AzureAD cmdlets.
* **Required Modules (Exact Versions):** 

  * Microsoft.Graph.Authentication 2.41.2
  * Microsoft.Graph.Users 2.41.2
  * Microsoft.Graph.Users.Actions 2.41.2
  * Microsoft.Graph.Groups 2.41.2
  * Microsoft.Graph.Teams 2.41.2
  * Microsoft.Graph.Identity.DirectoryManagement 2.41.2
  * ExchangeOnlineManagement 3.10.1
* **Exchange Command Allowlist:** Connect-SnSOffice365 connects to Exchange Online with an explicit list of permitted commands ($conf.commandNames). Whenever a new Exchange Online cmdlet is used in module code, add it to the commandNames list so Managed Identity sessions in Azure Automation can invoke it.

#### Module Architecture & Public Surface

* **Exported Functions (13 total — do not export internal helpers or $SNSConf):** 

  * Configuration & Auth: New-SNSConfiguration, Set-SNSConfiguration, Connect-SnSOffice365, Disconnect-SnSOffice365
  * Sync Operations: Invoke-SNSUppdateOffice365User, SNSUpdateExchangeDistributionGroups, Update-SNMicrosoft365Group, Update-SNSharedMailbox
  * Mandatory Moments: Send-SNSMandatoryMomentReport, Send-SNSMandatoryMomentReminder, Invoke-SNSMandatoryMomentReportAndReminder
  * Logging & Summary: Write-SNSLog, Get-SNSRunSummary
* **Configuration Validation:** Validate all rule hashtables (MailListSettings, Microsoft365GroupSettings, SharedMailboxSettings) up front before fetching external data or mutating Microsoft 365 state, so invalid settings fail fast with zero side effects.
* **Member Linking:** Scoutnet members are linked to Microsoft 365 accounts via their Scoutnet member ID stored in CustomAttribute1.

#### Mandatory PSScriptAnalyzer Rules (PSScriptAnalyzerSettings.psd1)

CI treats any lint warning as an error. All .ps1, .psm1, and .psd1 files must strictly satisfy these 10 enforced rules: 

1. **PSAvoidDefaultValueSwitchParameter**: Never assign a default value (= $true or = $false) to [switch] parameters.
2. **PSMisleadingBacktick**: Never leave trailing spaces after a line-continuation backtick (`).
3. **PSMissingModuleManifestField**: Keep all required fields intact when editing Office365-Scoutnet-synk.psd1.
4. **PSReservedCmdletChar**: Do not use reserved characters in function names.
5. **PSReservedParams**: Never define custom parameters that collide with PowerShell common parameters (Verbose, Debug, ErrorAction, WhatIf, Confirm, etc.).
6. **PSShouldProcess**: Any function declaring SupportsShouldProcess must call $PSCmdlet.ShouldProcess(...) before making changes, and functions calling ShouldProcess must declare SupportsShouldProcess.
7. **PSUseApprovedVerbs**: Use standard PowerShell approved verbs (Get-Verb) for all new functions.
8. **PSAvoidUsingCmdletAliases**: Always write full cmdlet and parameter names (e.g., ForEach-Object, Where-Object, Select-Object, Get-ChildItem)—never use aliases (%, ?, select, gci, cd, rm).
9. **PSUseDeclaredVarsMoreThanAssignments**: Never leave unused local variables in scripts, functions, or tests.
10. **PSAvoidTrailingWhitespace**: Strip all trailing spaces and tabs from every line.

#### Additional PowerShell Coding Standards

* **Dual-Environment Execution:** Core code in Office365-Scoutnet-synk/ must run unattended in Azure Automation Runbooks (-ManagedIdentity) as well as locally. Never use Read-Host or assume interactive sessions inside the module.
* **Array Preservation:** Avoid accidental array unwrapping when returning collections from functions (Write-Output -NoEnumerate or unary comma , $array).
* **Logging & Error Handling:** Use Write-SNSLog for operational logging and wrap external API calls (Scoutnet REST API, Microsoft Graph, Exchange Online) in try/catch blocks with -ErrorAction Stop. Never log secrets, access tokens, or sensitive personal data.

#### Testing, Privacy & Documentation

* **Automated Tests:** Run Pester tests in tests/ (Invoke-Pester) for all new features and bug fixes, mocking all Scoutnet, Graph, and Exchange Online calls.
* **GDPR / PII Protection:** Never commit real Scoutnet member data, personal numbers, email addresses, or tenant IDs. Always sanitize test fixtures using anonymize_json.py.
* **Language & Docs:** Write user-facing documentation, wiki manuals, ReleaseNotes in Office365-Scoutnet-synk.psd1, and default email templates in **Swedish**. Follow .markdownlint.json for all Markdown edits.

#### Documentation & Wiki (Office365-Scoutnet-synk.wiki)

* **Location:** The Swedish user documentation lives in the sibling workspace folder Office365-Scoutnet-synk.wiki (cloned from https://github.com/Scouterna/Office365-Scoutnet-synk.wiki.git).
* **Sync Docs with Code Changes:** Whenever you add, rename, or change a configuration setting, exported function, parameter, or behavior in Office365-Scoutnet-synk, update the corresponding Swedish Markdown file(s) in Office365-Scoutnet-synk.wiki (and follow .markdownlint.json):
  * Home.md, _Sidebar.md & Manual.md: Wiki entrypoints, navigation sidebar, and manual overview
  * Manual-Konfiguration.md: New-SNSConfiguration, Set-SNSConfiguration, Connect-SnSOffice365, Disconnect-SnSOffice365, and global $conf settings
  * Manual-Användarkonton.md: User account sync (Invoke-SNSUppdateOffice365User)
  * Manual-Distributionsgrupper.md: Distribution lists (SNSUpdateExchangeDistributionGroups, $conf.MailListSettings)
  * Manual-Microsoft-365-grupper.md: M365 groups, security groups, and Teams (Update-SNMicrosoft365Group, $conf.Microsoft365GroupSettings)
  * Manual-Delade-mailboxar.md: Shared mailboxes (Update-SNSharedMailbox, $conf.SharedMailboxSettings)
  * Manual-Obligatoriska-moment.md: Reports & reminders (Send-SNSMandatoryMomentReport, Send-SNSMandatoryMomentReminder, Invoke-SNSMandatoryMomentReportAndReminder)
  * Manual-Körningssammanfattning.md: Run summary output (Get-SNSRunSummary)
  * Microsoft-Azure-Runbook.md & Körning-på-dator.md: Azure Automation Runbook and local execution guides
  * Ny-version.md: Upgrade instructions and version requirements
  * Test.md, Automatiserade-tester.md & Anonymisera-nedladdade-JSON-filer.md: -WhatIf testing, Pester tests, and JSON anonymization
  * Teknik.md & Support.md: Technical clarifications and support info
* **GitHub Issues:** When creating or updating GitHub issues using `gh issue create` or `gh issue edit`, model issue descriptions on existing well-structured issues (such as #78 for feature requests) using Swedish headings like "Är din idé relaterad till ett problem?", "Beskriv hur idén kan implementeras", code examples, expected capabilities, and technical references.
* **Token Efficiency:** Only read the specific wiki .md file(s) relevant to the feature being modified—never read all 18 wiki files at once. (Note that Funktionsreferens is auto-generated and published to GitHub Pages, not stored as a file in the wiki repo.)
