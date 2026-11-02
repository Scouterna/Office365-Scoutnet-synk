BeforeAll {
    $repositoryRoot = Split-Path -Path $PSScriptRoot -Parent
    $manifestPath = Join-Path -Path $repositoryRoot -ChildPath 'Office365-Scoutnet-synk\Office365-Scoutnet-synk.psd1'
    $manifest = Import-PowerShellDataFile -Path $manifestPath -ErrorAction Stop
    $moduleVersion = [version]$manifest.ModuleVersion
    $moduleMajorMinor = '{0}.{1}' -f $moduleVersion.Major, $moduleVersion.Minor

    function Get-RequiresStatement {
        param([string]$Path)

        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
        if ($errors.Count -gt 0) {
            throw "Could not parse ${Path}: $($errors -join '; ')"
        }
        $ast.ScriptRequirements
    }

    function Get-RepoPath {
        param([string]$RelativePath)
        Join-Path -Path $repositoryRoot -ChildPath $RelativePath
    }
}

Describe 'Version references in scripts match the module manifest' {
    BeforeDiscovery {
        $syncScripts = @(
            @{ File = 'MaillistSynk.ps1' }
            @{ File = 'MaillistSynk_local.ps1' }
        )
        $workflows = @(
            @{ File = '.github\workflows\lint.yml' }
            @{ File = '.github\workflows\pages.yml' }
            @{ File = '.github\workflows\release.yml' }
        )
    }

    Context '#Requires in <File>' -ForEach $syncScripts {
        It 'requires a module version that is not newer than the manifest' {
            $requirements = Get-RequiresStatement -Path (Get-RepoPath $File)
            $required = $requirements.RequiredModules |
                Where-Object Name -eq 'Office365-Scoutnet-synk'
            $required | Should -Not -BeNullOrEmpty -Because "$File must have: #Requires -Modules @{ ModuleName=`"Office365-Scoutnet-synk`"; ModuleVersion=`"$moduleMajorMinor`" }"

            $requiredVersion = [version]$required.Version
            $requiredVersion | Should -BeLessOrEqual $moduleVersion -Because "$File requires module version $requiredVersion, but the manifest is $moduleVersion. Update ModuleVersion in the #Requires line to `"$moduleMajorMinor`""
        }

        It 'requires the module major.minor version of the manifest' {
            $requirements = Get-RequiresStatement -Path (Get-RepoPath $File)
            $required = $requirements.RequiredModules |
                Where-Object Name -eq 'Office365-Scoutnet-synk'
            $requiredVersion = [version]$required.Version
            $actual = '{0}.{1}' -f $requiredVersion.Major, $requiredVersion.Minor
            $actual | Should -Be $moduleMajorMinor -Because "$File has ModuleVersion `"$($required.Version)`" in #Requires, but the manifest version is $moduleVersion. Change it to ModuleVersion=`"$moduleMajorMinor`""
        }

        It 'requires the PowerShell version of the manifest' {
            $requirements = Get-RequiresStatement -Path (Get-RepoPath $File)
            $requirements.RequiredPSVersion | Should -Be ([version]$manifest.PowerShellVersion) -Because "$File must have: #Requires -Version $($manifest.PowerShellVersion) (PowerShellVersion in the manifest)"
        }
    }

    Context 'setupAzure.ps1' {
        It 'installs the manifest version of the module' {
            $content = Get-Content -Path (Get-RepoPath 'setupAzure.ps1') -Raw
            $found = [regex]::Match($content, '\$ScoutnetModuleVersion\s*=\s*"(?<v>[^"]+)"')
            $found.Success | Should -BeTrue -Because 'setupAzure.ps1 must assign $ScoutnetModuleVersion'
            $found.Groups['v'].Value | Should -Be $manifest.ModuleVersion -Because "`$ScoutnetModuleVersion in setupAzure.ps1 must be $($manifest.ModuleVersion)"
        }
    }

    Context 'updateraAzureOffice365ScoutnetSync.ps1' {
        BeforeAll {
            $updateContent = Get-Content -Path (Get-RepoPath 'updateraAzureOffice365ScoutnetSync.ps1') -Raw
        }

        It 'has the manifest version as default ModuleVersion' {
            $found = [regex]::Match($updateContent, "\[string\]\`$ModuleVersion\s*=\s*'(?<v>[^']+)'")
            $found.Success | Should -BeTrue -Because 'updateraAzureOffice365ScoutnetSync.ps1 must have a [string]$ModuleVersion default'
            $found.Groups['v'].Value | Should -Be $manifest.ModuleVersion -Because "The default of -ModuleVersion in updateraAzureOffice365ScoutnetSync.ps1 must be $($manifest.ModuleVersion)"
        }

        It 'uses the manifest version in the help examples' {
            $exampleVersions = [regex]::Matches($updateContent, '(?:-ModuleVersion |till version )(\d+\.\d+\.\d+)') |
                ForEach-Object { $_.Groups[1].Value } |
                Sort-Object -Unique
            $exampleVersions | Should -Not -BeNullOrEmpty
            foreach ($version in $exampleVersions) {
                $version | Should -Be $manifest.ModuleVersion -Because "A help example in updateraAzureOffice365ScoutnetSync.ps1 mentions version $version, but the manifest is $($manifest.ModuleVersion)"
            }

            $packageVersions = [regex]::Matches($updateContent, 'Office365-Scoutnet-synk\.(\d+\.\d+\.\d+)\.zip') |
                ForEach-Object { $_.Groups[1].Value } |
                Sort-Object -Unique
            foreach ($version in $packageVersions) {
                $version | Should -Be $manifest.ModuleVersion -Because "A zip file name in updateraAzureOffice365ScoutnetSync.ps1 mentions version $version, but the manifest is $($manifest.ModuleVersion)"
            }
        }
    }

    Context 'Workflow <File>' -ForEach $workflows {
        It 'installs the dependency versions declared in the manifest' {
            $content = Get-Content -Path (Get-RepoPath $File) -Raw
            foreach ($dependency in $manifest.RequiredModules) {
                $pattern = 'Install-Module\s+{0}\s+-RequiredVersion\s+(\S+)' -f [regex]::Escape($dependency.ModuleName)
                $found = [regex]::Match($content, $pattern)
                $found.Success | Should -BeTrue -Because "$File must install $($dependency.ModuleName) -RequiredVersion $($dependency.RequiredVersion)"
                $found.Groups[1].Value | Should -Be $dependency.RequiredVersion -Because "$File installs $($dependency.ModuleName) $($found.Groups[1].Value), but the manifest requires $($dependency.RequiredVersion)"
            }
        }
    }
}
