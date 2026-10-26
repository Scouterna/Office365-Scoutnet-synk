param(
    [string]$accountName = 'Scoutnet-sync',
    [string]$rgName = 'Scoutnet-sync',
    [string]$location = 'swedencentral'
)

$GraphDesiredVersion = "2.41.2"
$ExchangeOnlineManagementVersion = "3.10.1"
$ScoutnetModuleVersion = "3.1.0"
$AzRuntimePackageVersion = '15.1.0'
$AzureCliRuntimePackageVersion = '2.77.0'
$AutomationRuntimeVersion = '7.6'
$RuntimeEnvironmentName = 'PowerShell-7.6-Office365-Scoutnet-synk'

$RequiredLocalModules = @(
    'Az.Accounts',
    'Az.Resources',
    'Az.Automation',
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Applications',
    'Microsoft.Graph.Identity.DirectoryManagement'
)

$MissingModules = @(
    $RequiredLocalModules | Where-Object {
        -not (Get-Module -ListAvailable -Name $_)
    }
)

if ($MissingModules.Count -gt 0)
{
    $MissingModuleNames = $MissingModules -join ', '
    Write-Host "Följande PowerShell-moduler saknas lokalt: $MissingModuleNames"
    Write-Host "Installera Az-modulerna genom att kopiera och köra:"
    Write-Host "Install-Module -Name Az.Accounts, Az.Resources, Az.Automation -Repository PSGallery -Scope CurrentUser"
    Write-Host "Installera Graph-modulerna genom att kopiera och köra:"
    Write-Host "Install-Module -Name Microsoft.Graph -Repository PSGallery -Scope CurrentUser"
    return
}

Import-Module Az.Accounts -ErrorAction Stop
if (-not (Get-Command Invoke-AzRestMethod -ErrorAction SilentlyContinue))
{
    throw "Az.Accounts saknar cmdleten Invoke-AzRestMethod. Uppdatera Az.Accounts och starta om PowerShell innan du kör skriptet igen."
}

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

function Assert-SNSAutomationRestSuccess
{
    param(
        [Parameter(Mandatory)]
        $Response,
        [Parameter(Mandatory)]
        [string]$Operation
    )

    if ($null -eq $Response -or $Response.StatusCode -notin @(200, 201, 202))
    {
        $statusCode = if ($null -eq $Response) { 'no response' } else { $Response.StatusCode }
        $content = if ($null -eq $Response) { '' } else { $Response.Content }
        throw "$Operation misslyckades. HTTP $statusCode. Azure-svar: $content"
    }
}

function New-SNSRuntimeEnvironmentPackage
{
    param(
        [Parameter(Mandatory)]
        [string]$Name,
        [Parameter(Mandatory)]
        [string]$Version,
        [Parameter(Mandatory)]
        [string]$ContentUri,
        [Parameter(Mandatory)]
        [string]$RuntimeEnvironmentResourcePath
    )

    $packageCollectionResourceId = "$RuntimeEnvironmentResourcePath/packages"
    $packageResourceId = "$RuntimeEnvironmentResourcePath/packages/$([uri]::EscapeDataString($Name))"
    $packagePayload = @{
        properties = @{
            contentLink = @{
                uri = $ContentUri
                version = $Version
            }
        }
    } | ConvertTo-Json -Depth 10

    $createResponse = Invoke-SNSAutomationRestMethod `
        -ResourceId $packageResourceId `
        -Method PUT `
        -Payload $packagePayload
    Assert-SNSAutomationRestSuccess `
        -Response $createResponse `
        -Operation "Skapande av paketet $Name version $Version"
    Write-Host "Azure accepterade paketimporten för $Name (HTTP $($createResponse.StatusCode))."

    $timeout = [TimeSpan]::FromMinutes(20)
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $previousState = $null
    do
    {
        Start-Sleep -Seconds 10
        $response = Invoke-SNSAutomationRestMethod -ResourceId $packageCollectionResourceId -Method GET
        $packageList = $response.Content | ConvertFrom-Json -ErrorAction Stop
        if ($null -eq $packageList.value)
        {
            throw "Azure returnerade ett oväntat svar vid hämtning av paketlistan för Runtime Environment. HTTP $($response.StatusCode): $($response.Content)"
        }

        $package = @($packageList.value | Where-Object { $_.name -eq $Name }) | Select-Object -First 1
        if ($null -eq $package)
        {
            if ($stopwatch.Elapsed -ge $timeout)
            {
                throw "Paketet $Name version $Version saknas i Runtime Environment efter 20 minuter. Senaste svar från Azure: $($response.Content)"
            }
            continue
        }

        $provisioningState = $package.properties.provisioningState
        if ([string]::IsNullOrWhiteSpace($provisioningState))
        {
            throw "Azure returnerade ingen importstatus för paketet $Name version $Version. Paketdetaljer: $($package | ConvertTo-Json -Depth 10 -Compress)"
        }

        if ($provisioningState -ne $previousState)
        {
            Write-Host "Importstatus för $Name`: $provisioningState"
            $previousState = $provisioningState
        }

        if ($provisioningState -eq 'Failed')
        {
            $importError = $package.properties.error
            throw "Importen av paketet $Name version $Version misslyckades i Runtime Environment. $($importError.code): $($importError.message)"
        }

        if ($stopwatch.Elapsed -ge $timeout)
        {
            throw "Importen av paketet $Name version $Version slutfördes inte inom 20 minuter."
        }
    } while ($provisioningState -ne 'Succeeded')
}

function Wait-SNSRuntimeEnvironment
{
    param(
        [Parameter(Mandatory)]
        [string]$ResourceId
    )

    $timeout = [TimeSpan]::FromMinutes(20)
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    do
    {
        $response = Invoke-SNSAutomationRestMethod -ResourceId $ResourceId -Method GET
        if ($response.StatusCode -eq 200)
        {
            $runtimeEnvironment = $response.Content | ConvertFrom-Json -ErrorAction Stop
            $runtime = $runtimeEnvironment.properties.runtime
            if ($runtime.language -ne 'PowerShell' -or $runtime.version -ne $AutomationRuntimeVersion)
            {
                throw "Runtime Environment '$RuntimeEnvironmentName' har oväntad runtime: $($runtime.language) $($runtime.version). Förväntade PowerShell $AutomationRuntimeVersion."
            }

            Write-Host "Runtime Environment '$RuntimeEnvironmentName' är tillgänglig (PowerShell $($runtime.version))."
            return
        }

        if ($response.StatusCode -ne 404)
        {
            throw "Kunde inte verifiera Runtime Environment '$RuntimeEnvironmentName'. HTTP $($response.StatusCode): $($response.Content)"
        }

        if ($stopwatch.Elapsed -ge $timeout)
        {
            throw "Runtime Environment '$RuntimeEnvironmentName' blev inte tillgänglig inom 20 minuter. Senaste svar från Azure: $($response.Content)"
        }

        Write-Host "Väntar på att Runtime Environment '$RuntimeEnvironmentName' ska bli tillgänglig..."
        Start-Sleep -Seconds 10
    } while ($true)
}

Connect-AzAccount

$RequiredScopes = @("Directory.AccessAsUser.All",
    "Directory.ReadWrite.All",
    "Directory.Read.All",
    "GroupMember.Read.All",
    "GroupMember.ReadWrite.All",
    "Group.ReadWrite.All",
    "Group.Read.All",
    "Team.Create",
    "User.ReadWrite.All",
    "User.Read.All",
    "Mail.Send")

Connect-MgGraph -Scopes $RequiredScopes

# Fetch API keys and Scoutnet user name.
$userId = Read-Host "Användarnamn för Scoutnets API. Användarnamnet är Kår-ID för webbtjänster som står på sidan Webbkoppling"
$pwd_secure_string_customlists = Read-Host "Ange API key för Scoutnets API api/group/customlists" -AsSecureString
$pwd_secure_string_memberlist = Read-Host "Ange API key för Scoutnets API api/group/memberlist" -AsSecureString


# Create the resourse group used to store the data,
New-AzResourceGroup -Name $rgName -Location $location

# Create the automation account in the resourse group.
New-AzAutomationAccount -Name $accountName -ResourceGroupName $rgName -Location $location

# Enable managed identity for the automation account.
Set-AzAutomationAccount -Name $accountName -ResourceGroupName $rgName -AssignSystemIdentity

$subscriptionId = (Get-AzContext -ErrorAction Stop).Subscription.Id
$automationResourcePath = "/subscriptions/$subscriptionId/resourceGroups/$([uri]::EscapeDataString($rgName))/providers/Microsoft.Automation/automationAccounts/$([uri]::EscapeDataString($accountName))"
$runtimeEnvironmentResourcePath = "$automationResourcePath/runtimeEnvironments/$([uri]::EscapeDataString($RuntimeEnvironmentName))"
$runtimeEnvironmentPayload = @{
    location = $location
    properties = @{
        description = 'PowerShell 7.6 runtime for Office365-Scoutnet-synk'
        defaultPackages = @{
            Az = $AzRuntimePackageVersion
            'Azure CLI' = $AzureCliRuntimePackageVersion
        }
        runtime = @{
            language = 'PowerShell'
            version = $AutomationRuntimeVersion
        }
    }
} | ConvertTo-Json -Depth 10

$runtimeEnvironmentResponse = Invoke-SNSAutomationRestMethod `
    -ResourceId $runtimeEnvironmentResourcePath `
    -Method PUT `
    -Payload $runtimeEnvironmentPayload
Assert-SNSAutomationRestSuccess `
    -Response $runtimeEnvironmentResponse `
    -Operation "Skapande av Runtime Environment '$RuntimeEnvironmentName'"
Write-Host "Azure accepterade skapandet av Runtime Environment '$RuntimeEnvironmentName' (HTTP $($runtimeEnvironmentResponse.StatusCode))."

Wait-SNSRuntimeEnvironment -ResourceId $runtimeEnvironmentResourcePath

$powershellgallery = 'https://www.powershellgallery.com/api/v2/package'
$runtimePackages = @(
    @{ Name = 'Microsoft.Graph.Authentication'; Version = $GraphDesiredVersion },
    @{ Name = 'Microsoft.Graph.Users'; Version = $GraphDesiredVersion },
    @{ Name = 'Microsoft.Graph.Users.Actions'; Version = $GraphDesiredVersion },
    @{ Name = 'Microsoft.Graph.Groups'; Version = $GraphDesiredVersion },
    @{ Name = 'Microsoft.Graph.Teams'; Version = $GraphDesiredVersion },
    @{ Name = 'Microsoft.Graph.Identity.DirectoryManagement'; Version = $GraphDesiredVersion },
    @{ Name = 'ExchangeOnlineManagement'; Version = $ExchangeOnlineManagementVersion },
    @{ Name = 'Office365-Scoutnet-synk'; Version = $ScoutnetModuleVersion }
)

foreach ($runtimePackage in $runtimePackages)
{
    $packageUri = "$powershellgallery/$($runtimePackage.Name)/$($runtimePackage.Version)"
    Write-Host "Installerar $($runtimePackage.Name) version $($runtimePackage.Version) i $RuntimeEnvironmentName..."
    New-SNSRuntimeEnvironmentPackage `
        -Name $runtimePackage.Name `
        -Version $runtimePackage.Version `
        -ContentUri $packageUri `
        -RuntimeEnvironmentResourcePath $runtimeEnvironmentResourcePath
}

# Fetch the new managed identity
$ManagedIdentity = Get-AzADServicePrincipal -DisplayName $accountName

# Create new permissions for the new identity.
$params = @{
    ServicePrincipalId = $ManagedIdentity.Id # managed identity object id
    PrincipalId = $ManagedIdentity.Id # managed identity object id
    ResourceId = (Get-MgServicePrincipal -Filter "AppId eq '00000002-0000-0ff1-ce00-000000000000'").id # Exchange online
    AppRoleId = "dc50a0fb-09a3-484d-be87-e023b12c6440" # Exchange.ManageAsApp
}
New-MgServicePrincipalAppRoleAssignedTo @params

# Give the identity Exchange rights.
$roleId = (Get-MgRoleManagementDirectoryRoleDefinition -Filter "DisplayName eq 'Exchange Administrator'").id
New-MgRoleManagementDirectoryRoleAssignment -PrincipalId $ManagedIdentity.Id -RoleDefinitionId $roleId -DirectoryScopeId "/"


# Give the identity all needed graph rights.
$GraphApp = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'" # Microsoft Graph

foreach ($scope in $RequiredScopes)
{
    $Role = $GraphApp.AppRoles | Where-Object {$_.Value -eq $scope}
    if ($Role)
    {
        $AppRoleAssignment = @{
            "PrincipalId" = $ManagedIdentity.Id
            "ResourceId" = $GraphApp.Id
            "AppRoleId" = $Role.Id }
        # Assign the Graph permission
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ManagedIdentity.Id -BodyParameter $AppRoleAssignment
    }
}


$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $userId, $pwd_secure_string_customlists
New-AzAutomationCredential -AutomationAccountName $accountName -Name "ScoutnetApiCustomLists" -Value $Credential -ResourceGroupName $rgName

$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $userId, $pwd_secure_string_memberlist
New-AzAutomationCredential -AutomationAccountName $accountName -Name "ScoutnetApiGroupMemberList" -Value $Credential -ResourceGroupName $rgName

New-AzAutomationVariable -AutomationAccountName $accountName -Name "ScoutnetApiUserId" -ResourceGroupName $rgName -Value $userId -Encrypted $False

New-AzAutomationVariable -AutomationAccountName $accountName -Name "ScoutnetMailListsHash" -ResourceGroupName $rgName -Value "tom" -Encrypted $False
