#Requires -Modules Az.Accounts, Az.Automation

param(
    [string]$AutomationAccountName = 'Scoutnet-sync',
    [string]$ResourceGroupName = 'Scoutnet-sync',
    [string]$ModuleVersion = '3.0.0-beta2'
)

$ErrorActionPreference = 'Stop'

$ModuleName = 'Office365-Scoutnet-synk'
$RuntimeEnvironmentName = 'PowerShell-7.6-Office365-Scoutnet-synk'
$RuntimeVersion = '7.6'
$PowerShellGalleryPackageUri = 'https://www.powershellgallery.com/api/v2/package'

function Invoke-SNSAutomationRestMethod
{
    param(
        [Parameter(Mandatory)]
        [string]$ResourceId,
        [Parameter(Mandatory)]
        [ValidateSet('GET', 'PUT')]
        [string]$Method,
        [string]$Payload
    )

    if ($ResourceId -notmatch '^https://management\.azure\.com/')
    {
        $ResourceId = 'https://management.azure.com/' + $ResourceId.TrimStart('/')
    }

    $requestParameters = @{
        Uri = [uri]"${ResourceId}?api-version=2024-10-23"
        Method = $Method
        ErrorAction = 'Stop'
    }
    if ($PSBoundParameters.ContainsKey('Payload'))
    {
        $requestParameters['Payload'] = $Payload
    }

    Invoke-AzRestMethod @requestParameters
}

function Get-SNSRuntimeEnvironmentPackages
{
    param(
        [Parameter(Mandatory)]
        [string]$PackagesResourceId
    )

    $response = Invoke-SNSAutomationRestMethod -ResourceId $PackagesResourceId -Method GET
    if ($response.StatusCode -ne 200)
    {
        throw "Kunde inte läsa paketen i Runtime Environment. HTTP $($response.StatusCode): $($response.Content)"
    }

    $packageList = $response.Content | ConvertFrom-Json -ErrorAction Stop
    if ($null -eq $packageList.value)
    {
        throw "Azure returnerade en ogiltig paketlista för Runtime Environment. HTTP $($response.StatusCode): $($response.Content)"
    }

    @($packageList.value)
}

function Get-SNSRequiredPackages
{
    $temporaryDirectory = Join-Path `
        -Path ([System.IO.Path]::GetTempPath()) `
        -ChildPath "Office365-Scoutnet-synk-$([guid]::NewGuid().ToString('N'))"
    $packagePath = Join-Path -Path $temporaryDirectory -ChildPath "$ModuleName.nupkg"
    $manifestPath = Join-Path -Path $temporaryDirectory -ChildPath "$ModuleName.psd1"
    New-Item -ItemType Directory -Path $temporaryDirectory -ErrorAction Stop | Out-Null

    try
    {
        Invoke-WebRequest `
            -Uri "$PowerShellGalleryPackageUri/$ModuleName/$ModuleVersion" `
            -OutFile $packagePath `
            -ErrorAction Stop

        $archive = [System.IO.Compression.ZipFile]::OpenRead($packagePath)
        try
        {
            $manifestEntries = @(
                $archive.Entries | Where-Object {
                    $_.FullName -match "(?i)(^|/)$([regex]::Escape($ModuleName))\.psd1$"
                }
            )
            if ($manifestEntries.Count -ne 1)
            {
                throw "Paketet $ModuleName $ModuleVersion innehåller inte exakt ett modulmanifest med namnet $ModuleName.psd1."
            }

            $entryStream = $manifestEntries[0].Open()
            try
            {
                $reader = [System.IO.StreamReader]::new($entryStream)
                try
                {
                    $manifestContent = $reader.ReadToEnd()
                }
                finally
                {
                    $reader.Dispose()
                }
            }
            finally
            {
                $entryStream.Dispose()
            }
        }
        finally
        {
            $archive.Dispose()
        }

        Set-Content -LiteralPath $manifestPath -Value $manifestContent -NoNewline -Encoding utf8
        $manifest = Import-PowerShellDataFile -Path $manifestPath -ErrorAction Stop

        $requiredPackages = [System.Collections.Generic.List[hashtable]]::new()
        $declaredDependencies = @()
        if ($null -ne $manifest.RequiredModules)
        {
            $declaredDependencies = @($manifest.RequiredModules)
        }
        foreach ($dependency in $declaredDependencies)
        {
            if ($dependency -is [string])
            {
                throw "Beroendet '$dependency' i $ModuleName $ModuleVersion saknar ett exakt RequiredVersion i modulmanifestet."
            }

            $dependencyName = $dependency.ModuleName
            $dependencyVersion = $dependency.RequiredVersion
            if ([string]::IsNullOrWhiteSpace($dependencyName) -or
                [string]::IsNullOrWhiteSpace([string]$dependencyVersion))
            {
                throw "Ett RequiredModules-beroende i $ModuleName $ModuleVersion saknar ModuleName eller exakt RequiredVersion."
            }

            if ($dependencyName -eq $ModuleName)
            {
                throw "$ModuleName $ModuleVersion kan inte ange sig själv som beroende."
            }

            $requiredPackages.Add(@{
                Name = $dependencyName
                Version = [string]$dependencyVersion
            })
        }

        $duplicateDependencies = @(
            $requiredPackages |
                Group-Object -Property Name |
                Where-Object { $_.Count -gt 1 }
        )
        if ($duplicateDependencies.Count -gt 0)
        {
            throw "Modulmanifestet för $ModuleName $ModuleVersion anger beroenden flera gånger: $($duplicateDependencies.Name -join ', ')."
        }

        $requiredPackages.Add(@{
            Name = $ModuleName
            Version = $ModuleVersion
        })
        return $requiredPackages.ToArray()
    }
    finally
    {
        Remove-Item -LiteralPath $temporaryDirectory -Recurse -Force -ErrorAction Stop
    }
}

function Install-SNSRuntimeEnvironmentPackage
{
    param(
        [Parameter(Mandatory)]
        [hashtable]$Package,
        [Parameter(Mandatory)]
        [string]$PackageResourceId,
        [Parameter(Mandatory)]
        [string]$PackagesResourceId
    )

    $packagePayload = @{
        properties = @{
            contentLink = @{
                uri = "$PowerShellGalleryPackageUri/$($Package.Name)/$($Package.Version)"
                version = $Package.Version
            }
        }
    } | ConvertTo-Json -Depth 10

    $response = Invoke-SNSAutomationRestMethod `
        -ResourceId $PackageResourceId `
        -Method PUT `
        -Payload $packagePayload
    if ($response.StatusCode -notin @(200, 201, 202))
    {
        throw "Azure accepterade inte paketet $($Package.Name) version $($Package.Version). HTTP $($response.StatusCode): $($response.Content)"
    }

    Write-Host "Azure accepterade importen av $($Package.Name) version $($Package.Version)."

    $timeout = [TimeSpan]::FromMinutes(20)
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $previousState = $null
    do
    {
        Start-Sleep -Seconds 10
        $installedPackage = Get-SNSRuntimeEnvironmentPackages -PackagesResourceId $PackagesResourceId |
            Where-Object { $_.name -eq $Package.Name } |
            Select-Object -First 1

        if ($null -eq $installedPackage)
        {
            if ($stopwatch.Elapsed -ge $timeout)
            {
                throw "Paketet $($Package.Name) version $($Package.Version) saknas i Runtime Environment efter 20 minuter."
            }
            continue
        }

        $provisioningState = $installedPackage.properties.provisioningState
        $installedVersion = $installedPackage.properties.contentLink.version
        if ($provisioningState -ne $previousState)
        {
            Write-Host "Importstatus för $($Package.Name) $installedVersion`: $provisioningState"
            $previousState = $provisioningState
        }

        if ($provisioningState -eq 'Failed')
        {
            $importError = $installedPackage.properties.error
            throw "Importen av $($Package.Name) version $($Package.Version) misslyckades. $($importError.code): $($importError.message)"
        }

        if ($stopwatch.Elapsed -ge $timeout)
        {
            throw "Importen av $($Package.Name) version $($Package.Version) slutfördes inte inom 20 minuter."
        }
    } while ($provisioningState -ne 'Succeeded' -or $installedVersion -ne $Package.Version)
}

Write-Host "Kontrollerar Azure-inloggning och Automation-kontot '$AutomationAccountName' i resursgruppen '$ResourceGroupName'..."
$azContext = Get-AzContext
if ($null -eq $azContext)
{
    throw 'Ingen aktiv Azure-inloggning hittades. Kör Connect-AzAccount och försök igen.'
}

Write-Host "Använder Azure-prenumerationen '$($azContext.Subscription.Name)'."
Get-AzAutomationAccount `
    -Name $AutomationAccountName `
    -ResourceGroupName $ResourceGroupName `
    -ErrorAction Stop | Out-Null

$automationResourcePath = "/subscriptions/$($azContext.Subscription.Id)/resourceGroups/$([uri]::EscapeDataString($ResourceGroupName))/providers/Microsoft.Automation/automationAccounts/$([uri]::EscapeDataString($AutomationAccountName))"
$runtimeEnvironmentResourceId = "$automationResourcePath/runtimeEnvironments/$([uri]::EscapeDataString($RuntimeEnvironmentName))"
$packagesResourceId = "$runtimeEnvironmentResourceId/packages"
$runtimeEnvironmentResponse = Invoke-SNSAutomationRestMethod -ResourceId $runtimeEnvironmentResourceId -Method GET
if ($runtimeEnvironmentResponse.StatusCode -ne 200)
{
    throw "Runtime Environment '$RuntimeEnvironmentName' kunde inte hittas eller läsas. HTTP $($runtimeEnvironmentResponse.StatusCode): $($runtimeEnvironmentResponse.Content)"
}

$runtimeEnvironment = $runtimeEnvironmentResponse.Content | ConvertFrom-Json -ErrorAction Stop
$runtime = $runtimeEnvironment.properties.runtime
if ($runtime.language -ne 'PowerShell' -or $runtime.version -ne $RuntimeVersion)
{
    throw "Runtime Environment '$RuntimeEnvironmentName' har runtime $($runtime.language) $($runtime.version), förväntade PowerShell $RuntimeVersion."
}

$RequiredPackages = Get-SNSRequiredPackages
Write-Host "Uppgraderar paket i Runtime Environment '$RuntimeEnvironmentName' enligt beroendena i $ModuleName $ModuleVersion..."
foreach ($package in $RequiredPackages)
{
    Write-Host "Importerar $($package.Name) version $($package.Version)..."
    $packageResourceId = "$packagesResourceId/$([uri]::EscapeDataString($package.Name))"
    Install-SNSRuntimeEnvironmentPackage `
        -Package $package `
        -PackageResourceId $packageResourceId `
        -PackagesResourceId $packagesResourceId
}

$installedPackages = Get-SNSRuntimeEnvironmentPackages -PackagesResourceId $packagesResourceId
foreach ($package in $RequiredPackages)
{
    $installedPackage = $installedPackages |
        Where-Object { $_.name -eq $package.Name } |
        Select-Object -First 1
    if ($null -eq $installedPackage)
    {
        throw "Slutkontrollen misslyckades: $($package.Name) saknas i Runtime Environment '$RuntimeEnvironmentName'."
    }

    $installedVersion = $installedPackage.properties.contentLink.version
    $provisioningState = $installedPackage.properties.provisioningState
    if ($installedVersion -ne $package.Version -or $provisioningState -ne 'Succeeded')
    {
        throw "Slutkontrollen misslyckades för $($package.Name): version '$installedVersion', status '$provisioningState'. Förväntade version '$($package.Version)' med status 'Succeeded'."
    }
}

Write-Host "Alla beroenden och $ModuleName $ModuleVersion är verifierade i Runtime Environment '$RuntimeEnvironmentName' under Automation-kontot '$AutomationAccountName'."
