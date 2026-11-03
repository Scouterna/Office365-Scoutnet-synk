#Requires -Version 7.6
<#
.SYNOPSIS
    Generates Markdown documentation for the module's public functions.

.DESCRIPTION
    Reads the comment-based help for all exported functions in Office365-Scoutnet-synk
    and writes a Markdown file for each function, an index page, and a Jekyll
    configuration. The pages.yml workflow publishes the generated site to GitHub Pages.

.PARAMETER OutputPath
    Directory where the documentation is written. The directory is not cleared.

.EXAMPLE
    .\tools\New-ModuleDocumentation.ps1 -OutputPath .\docs-source
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

$manifestPath = Join-Path $PSScriptRoot '..\Office365-Scoutnet-synk\Office365-Scoutnet-synk.psd1'
$manifest = Import-PowerShellDataFile -Path $manifestPath
$module = Import-Module $manifestPath -Force -PassThru
$ignoredParameters = 'WhatIf', 'Confirm'

New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
$utf8 = [System.Text.UTF8Encoding]::new($false)

function ConvertTo-HelpText
{
    param($Node)

    if ($null -eq $Node)
    {
        return ''
    }
    return ((@($Node) | ForEach-Object { $_.Text }) -join "`n").Trim()
}

$commands = @(Get-Command -Module $module.Name | Sort-Object Name)
$index = [System.Text.StringBuilder]::new()
[void]$index.AppendLine("---`nlayout: default`ntitle: $($module.Name)`n---`n")
[void]$index.AppendLine("# $($module.Name) $($manifest.ModuleVersion)`n")
[void]$index.AppendLine("$($manifest.Description)`n")
[void]$index.AppendLine("For the user guide, see the [wiki](https://github.com/Scouterna/Office365-Scoutnet-synk/wiki). This site is generated from the module's built-in help.`n")
[void]$index.AppendLine("## Functions`n")

foreach ($command in $commands)
{
    $help = Get-Help $command.Name -Full
    $page = [System.Text.StringBuilder]::new()
    [void]$page.AppendLine("---`nlayout: default`ntitle: $($command.Name)`n---`n")
    [void]$page.AppendLine("# $($command.Name)`n")

    $synopsis = ([string]$help.Synopsis).Trim()
    if ($synopsis)
    {
        [void]$page.AppendLine("$synopsis`n")
    }
    $description = ConvertTo-HelpText $help.description
    if ($description)
    {
        [void]$page.AppendLine("## Description`n`n$description`n")
    }

    [void]$page.AppendLine("## Syntax`n")
    foreach ($syntax in (Get-Command $command.Name -Syntax))
    {
        [void]$page.AppendLine("``````powershell`n$($syntax.Trim())`n```````n")
    }

    $parameters = @($help.parameters.parameter | Where-Object { $_.name -notin $ignoredParameters })
    if ($parameters.Count -gt 0)
    {
        [void]$page.AppendLine("## Parameters`n")
        foreach ($parameter in $parameters)
        {
            $typeName = $parameter.type.name
            $required = if ($parameter.required -eq 'true') { 'required' } else { 'optional' }
            [void]$page.AppendLine("### -$($parameter.name)`n")
            [void]$page.AppendLine("Type: ``$typeName``; $required.`n")
            $parameterDescription = ConvertTo-HelpText $parameter.description
            if ($parameterDescription)
            {
                [void]$page.AppendLine("$parameterDescription`n")
            }
        }
    }

    $examples = @($help.examples.example)
    if ($examples.Count -gt 0 -and $null -ne $examples[0])
    {
        [void]$page.AppendLine("## Examples`n")
        $number = 1
        foreach ($example in $examples)
        {
            [void]$page.AppendLine("### Example $number`n")
            $code = ([string]$example.code).Trim()
            [void]$page.AppendLine("``````powershell`n$code`n```````n")
            $remarks = ConvertTo-HelpText $example.remarks
            if ($remarks)
            {
                [void]$page.AppendLine("$remarks`n")
            }
            $number++
        }
    }

    [System.IO.File]::WriteAllText((Join-Path $OutputPath "$($command.Name).md"), $page.ToString(), $utf8)
    $line = "- [$($command.Name)]($($command.Name).html)"
    if ($synopsis)
    {
        $line += " - $synopsis"
    }
    [void]$index.AppendLine($line)
}

[System.IO.File]::WriteAllText((Join-Path $OutputPath 'index.md'), $index.ToString(), $utf8)
[System.IO.File]::WriteAllText(
    (Join-Path $OutputPath '_config.yml'),
    "title: $($module.Name)`ntheme: jekyll-theme-midnight`n",
    $utf8)

Write-Output "Generated documentation for $($commands.Count) functions in $OutputPath"
