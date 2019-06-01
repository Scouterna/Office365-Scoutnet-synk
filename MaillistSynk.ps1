#Requires -Version 5.1
#Requires -Modules Office365-Scoutnet-synk

# Lämplig inställning i Axure aoutomation.
$ProgressPreference = "silentlyContinue"

# Server att skicka mail via.
$emailSMTPServer = "outlook.office365.com"

# Aktiverar Verbose logg. Standardvärde är silentlyContinue
#$VerbosePreference = "Continue"

# Vem ska mailet med loggen skickas ifrån.
$emailFromAddress = "info@landvetterscout.se"

# Vem ska mailet med loggen skickas till.
$emailToAddress = "karl.thoren@landvetterscout.se"

# Rubrik på mailet.
$emailSubject = "Maillist sync log"

# Domännam för scoutkårens office 365.
$DomainName = "landvetterscout.se"

# Kommaseparerad lista med e-postadresser som läggs in i alla distributions grupper.
$AdminMails = "karl.thoren@landvetterscout.se"

# Hashtable med Office 365 distributions gruppen som nyckel och ID numret på Scoutnets maillista som värde.
# Distributions grupper som är med här kommer att synkroniseras.
$lists = @{"krypen" = "4900"; "ravarna" = "4904"; "rovdjuren" = "4923"; "upptackare" = "4922"; "utmanarna" = "4924"}

# Här börjar själva skriptet.

$moduleInfo = Get-Module Office365-Scoutnet-synk

Write-SNSLog "Använder version $($moduleInfo.Version) av Office365-Scoutnet-synk"

# Hämtar credentials för Scoutnet API och för Office 365.
try
{
    # Credentials för access till Office 365 och för att kunna skicka mail.
    $Credential365 = Get-AutomationPSCredential -Name "MSOnline-Credentials" -ErrorAction "Stop"

    # Credentials för Scoutnets API api/group/customlists
    $CredentialCustomLists = Get-AutomationPSCredential -Name 'ScoutnetApiCustomLists-Credentials' -ErrorAction "Stop"

    # Credentials för Scoutnets API api/group/memberlist
    $CredentialMembers = Get-AutomationPSCredential -Name 'ScoutnetApiGroupMemberList-credentials' -ErrorAction "Stop"
}
Catch
{
    Write-SNSLog -Level "Error" "Kunde inte hämta nödvändiga credentials. Error $_"
    throw
}

try
{
    # Hämtar senaste körningens hash.
    $ValidationHash = Get-AutomationVariable -Name 'ScoutnetMailListsHash' -ErrorAction "Stop"
}
Catch
{
    Write-SNSLog -Level "Error" "Kunde inte hämta variabeln ScoutnetMailListsHash. Error $_"
    throw
}

# Kör updateringsfunktionen.
$NewValidationHash = SNSUpdateExchangeDistributionGroups -CredentialCustomlists $CredentialCustomLists `
    -CredentialMemberlist $CredentialMembers -Credential365 $Credential365 -Maillists $lists `
    -ValidationHash $ValidationHash -AdminMails $AdminMails -DomainName $DomainName

try
{
    # Spara hashen till nästa körning.
    Set-AutomationVariable -Name 'ScoutnetMailListsHash' -Value $NewValidationHash -ErrorAction "Continue"
}
Catch
{
    Write-SNSLog -Level "Error" "Kunde inte spara variabeln ScoutnetMailListsHash. Error $_"
    throw
}

# Skapa ett mail med loggen och skicka till admin.
$bodyData = Get-Content -Path 'PowerShellLog.log' -Raw -Encoding UTF8 -ErrorAction "Continue"
Send-MailMessage -Credential $Credential365 -From $emailFromAddress -To $emailToAddress -Subject $emailSubject -Body $bodyData -SmtpServer $emailSMTPServer -UseSSL -Encoding UTF8 -ErrorAction "Continue"
