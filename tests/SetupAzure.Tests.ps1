BeforeAll {
    $repositoryRoot = Split-Path -Path $PSScriptRoot -Parent
    $setupScriptPath = Join-Path -Path $repositoryRoot -ChildPath 'setupAzure.ps1'
    $manifestPath = Join-Path -Path $repositoryRoot -ChildPath 'Office365-Scoutnet-synk\Office365-Scoutnet-synk.psd1'

    $tokens = $null
    $parseErrors = $null
    $setupAst = [System.Management.Automation.Language.Parser]::ParseFile(
        $setupScriptPath,
        [ref]$tokens,
        [ref]$parseErrors
    )
    if ($parseErrors.Count -gt 0) {
        throw "Could not parse setupAzure.ps1: $($parseErrors -join '; ')"
    }

    $manifest = Import-PowerShellDataFile -Path $manifestPath -ErrorAction Stop
    $scriptVariables = @{}
    foreach ($variableName in @('GraphDesiredVersion', 'ExchangeOnlineManagementVersion', 'ScoutnetModuleVersion')) {
        $assignment = $setupAst.Find(
            {
                param($node)
                $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
                $node.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
                $node.Left.VariablePath.UserPath -eq $variableName
            },
            $true
        )
        if ($null -eq $assignment) {
            throw "Could not find `$$variableName assignment in setupAzure.ps1."
        }

        if ($assignment.Right -is [System.Management.Automation.Language.CommandExpressionAst]) {
            $expression = $assignment.Right.Expression
        }
        else {
            $expression = $assignment.Right.PipelineElements[0].Expression
        }
        if ($expression -is [System.Management.Automation.Language.StringConstantExpressionAst] -or
            $expression -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) {
            $scriptVariables[$variableName] = $expression.Value
        }
        else {
            $scriptVariables[$variableName] = $expression.SafeGetValue()
        }
    }

    $runtimePackagesAssignment = $setupAst.Find(
        {
            param($node)
            $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $node.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
            $node.Left.VariablePath.UserPath -eq 'runtimePackages'
        },
        $true
    )
    if ($null -eq $runtimePackagesAssignment) {
        throw 'Could not find $runtimePackages assignment in setupAzure.ps1.'
    }

    $runtimePackages = @(
        $runtimePackagesAssignment.Right.FindAll(
            { param($node) $node -is [System.Management.Automation.Language.HashtableAst] },
            $true
        ) | ForEach-Object {
            $package = @{}
            foreach ($pair in $_.KeyValuePairs) {
                $key = $pair.Item1.SafeGetValue()
                $expression = $pair.Item2.PipelineElements[0].Expression
                if ($expression -is [System.Management.Automation.Language.VariableExpressionAst]) {
                    $package[$key] = $scriptVariables[$expression.VariablePath.UserPath]
                }
                else {
                    $package[$key] = $expression.SafeGetValue()
                }
            }

            [pscustomobject]$package
        }
    )
}

Describe 'setupAzure runtime package versions' {
    It 'installs the versions declared by the Office365-Scoutnet-synk manifest' {
        $expectedPackages = @(
            $manifest.RequiredModules | ForEach-Object {
                [pscustomobject]@{
                    Name = $_.ModuleName
                    Version = [string]$_.RequiredVersion
                }
            }
        ) + @(
            [pscustomobject]@{
                Name = 'Office365-Scoutnet-synk'
                Version = [string]$manifest.ModuleVersion
            }
        )

        $runtimePackages.Count | Should -Be $expectedPackages.Count
        foreach ($expectedPackage in $expectedPackages) {
            $matchingPackages = @($runtimePackages | Where-Object Name -eq $expectedPackage.Name)
            $matchingPackages.Count | Should -Be 1
            $matchingPackages[0].Version | Should -Be $expectedPackage.Version
        }
    }
}
