#Requires -Version 5.1
#Requires -Modules Office365-Scoutnet-synk

# Lämplig inställning i Axure automation.
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

# Hashtable med id på Office 365 distributionsgruppen som nyckel. 
# Distributions grupper som är med här kommer att synkroniseras.
$mailListSettings = @{
    "utmanarna" = @{ # Namet på distributions gruppen i office 365. Används som grupp ID till Get-DistributionGroupMember.
        "scoutnet_list_id"= "4924"; # Listans Id i Scoutnet.
        "scouter_synk_option" = ""; # Synkoption för scouter. Giltiga värden är m,f,e eller tomt. 
        "ledare_synk_option" = "@"; # Synkoption för ledare. Giltiga väerden är @,- eller &.
        "email_addresses" = "","";  # Lista med e-postadresser. 
    };
    "rovdjuren" = @{
        "scoutnet_list_id"= "4923";
        "scouter_synk_option" = ""; # Alla adresser
        "ledare_synk_option" = "@"; # Bara office 365 adresser
        "email_addresses" = "";
    };
    "upptackare" = @{
        "scoutnet_list_id"= "4922";
        "scouter_synk_option" = ""; # Alla adresser
        "ledare_synk_option" = "@"; # Bara office 365 adresser
        "email_addresses" = "karl.thoren@landvetterscout.se";
    };
    "krypen" = @{
        "scoutnet_list_id"= "4900";
        "scouter_synk_option" = ""; # Alla adresser
        "ledare_synk_option" = "@"; # Bara office 365 adresser
        "email_addresses" = "karl.thoren@landvetterscout.se";
    };
    "ravarna" = @{
        "scoutnet_list_id"= "4904";
        "scouter_synk_option" = ""; # Alla adresser
        "ledare_synk_option" = "@"; # Bara office 365 adresser
        "email_addresses" = "karl.thoren@landvetterscout.se";
    }
}

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
    -CredentialMemberlist $CredentialMembers -Credential365 $Credential365 -MailListSettings $mailListSettings `
    -ValidationHash $ValidationHash -DomainName $DomainName

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
