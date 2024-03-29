
# Set the new accout name and resource group name.
$accountName = 'Scoutnet-sync'
$rgName = 'Scoutnet-sync'
$location = 'swedencentral'

$GraphDesiredVersion = "1.28.0"

Connect-AzAccount

$RequiredScopes = @("Directory.AccessAsUser.All",
    "Directory.ReadWrite.All",
    "Directory.Read.All",
    "GroupMember.Read.All",
    "GroupMember.ReadWrite.All",
    "Group.ReadWrite.All",
    "Group.Read.All",
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


$powershellgallery = "https://www.powershellgallery.com/api/v2/package"

$moduleName = 'PowerShellGet'
New-AzAutomationModule -AutomationAccountName $accountName -ResourceGroupName $rgName -Name $moduleName -ContentLinkUri "$powershellgallery/$moduleName" -ErrorAction "Stop"

"Wait for $moduleName to be installed."
do
{
    sleep 1
    $result = Get-AzAutomationModule -AutomationAccountName $accountName -ResourceGroupName $rgName $moduleName -ErrorAction "Stop"
} while (($result.ProvisioningState -eq "Creating") -Or ($result.ProvisioningState -eq "ConnectionTypeImported"))

if ($result.ProvisioningState -eq "Failed")
{
    throw "Could not install $moduleName"
}


# Install ExchangeOnlineManagement
$moduleName = 'ExchangeOnlineManagement'
New-AzAutomationModule -AutomationAccountName $accountName -ResourceGroupName $rgName -Name $moduleName -ContentLinkUri "$powershellgallery/$moduleName" -ErrorAction "Stop"

$moduleName = 'Microsoft.Graph.Authentication'
New-AzAutomationModule -AutomationAccountName $accountName -ResourceGroupName $rgName -Name $moduleName -ContentLinkUri "$powershellgallery/$moduleName/$GraphDesiredVersion" -ErrorAction "Stop"

"Wait for $moduleName to be installed."
do
{
    sleep 1
    $result = Get-AzAutomationModule -AutomationAccountName $accountName -ResourceGroupName $rgName $moduleName -ErrorAction "Stop"
} while (($result.ProvisioningState -eq "Creating") -Or ($result.ProvisioningState -eq "ConnectionTypeImported"))

if ($result.ProvisioningState -eq "Failed")
{
    throw "Could not install $moduleName"
}

$moduleName = 'Microsoft.Graph.Users'
New-AzAutomationModule -AutomationAccountName $accountName -ResourceGroupName $rgName -Name $moduleName -ContentLinkUri "$powershellgallery/$moduleName/$GraphDesiredVersion" -ErrorAction "Stop"

"Wait for $moduleName to be installed."
do
{
    sleep 1
    $result = Get-AzAutomationModule -AutomationAccountName $accountName -ResourceGroupName $rgName $moduleName -ErrorAction "Stop"
} while (($result.ProvisioningState -eq "Creating") -Or ($result.ProvisioningState -eq "ConnectionTypeImported"))


if ($result.ProvisioningState -eq "Failed")
{
    throw "Could not install $moduleName"
}

$moduleName = 'Microsoft.Graph.Identity.DirectoryManagement'
New-AzAutomationModule -AutomationAccountName $accountName -ResourceGroupName $rgName -Name $moduleName -ContentLinkUri "$powershellgallery/$moduleName/$GraphDesiredVersion" -ErrorAction "Stop"

$moduleName = 'Microsoft.Graph.Users.Actions'
New-AzAutomationModule -AutomationAccountName $accountName -ResourceGroupName $rgName -Name $moduleName -ContentLinkUri "$powershellgallery/$moduleName/$GraphDesiredVersion" -ErrorAction "Stop"

$moduleName = 'Microsoft.Graph.Groups'
New-AzAutomationModule -AutomationAccountName $accountName -ResourceGroupName $rgName -Name $moduleName -ContentLinkUri "$powershellgallery/$moduleName/$GraphDesiredVersion" -ErrorAction "Stop"

"Wait for $moduleName to be installed."
do
{
    sleep 1
    $result = Get-AzAutomationModule -AutomationAccountName $accountName -ResourceGroupName $rgName $moduleName -ErrorAction "Stop"
} while (($result.ProvisioningState -eq "Creating") -Or ($result.ProvisioningState -eq "ConnectionTypeImported"))


if ($result.ProvisioningState -eq "Failed")
{
    throw "Could not install $moduleName"
}

$moduleName = 'Microsoft.Graph.Users.Actions'
"Wait for $moduleName to be installed."
do
{
    sleep 1
    $result = Get-AzAutomationModule -AutomationAccountName $accountName -ResourceGroupName $rgName $moduleName -ErrorAction "Stop"
} while (($result.ProvisioningState -eq "Creating") -Or ($result.ProvisioningState -eq "ConnectionTypeImported"))

if ($result.ProvisioningState -eq "Failed")
{
    throw "Could not install $moduleName"
}


# Install Office365-Scoutnet-synk
$moduleName = 'Office365-Scoutnet-synk'
New-AzAutomationModule -AutomationAccountName $accountName -ResourceGroupName $rgName -Name $moduleName -ContentLinkUri "$powershellgallery/$moduleName" -ErrorAction "Stop"

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
New-AzAutomationCredential -AutomationAccountName $accountName -Name "ScoutnetApiCustomLists-Credentials" -Value $Credential -ResourceGroupName $rgName

$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $userId, $pwd_secure_string_memberlist
New-AzAutomationCredential -AutomationAccountName $accountName -Name "ScoutnetApiGroupMemberList-Credentials" -Value $Credential -ResourceGroupName $rgName

New-AzAutomationVariable -AutomationAccountName $accountName -Name "ScoutnetApiUserId" -ResourceGroupName $rgName -Value $userId -Encrypted $False

New-AzAutomationVariable -AutomationAccountName $accountName -Name "ScoutnetMailListsHash" -ResourceGroupName $rgName -Value "tom" -Encrypted $False
