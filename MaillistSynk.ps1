#Requires -Version 5.1
#Requires -Modules Office365-Scoutnet-synk

# Lämplig inställning i Axure automation.
$ProgressPreference = "silentlyContinue"

# Licenser för nya användare. Byt ut <office 365 licensnamn> med det namnet du har.
# Exemplet nedan lägger in STANDARDPACK och FLOW_FREE. På STANDARDPACK applikationerna "YAMMER_ENTERPRISE", "SWAY","Deskless","POWERAPPS_O365_P1" avstängda.
# För att hitta vad dina licenser heter använd Get-MsolAccountSku ifrån MSonline paketet.
$LicenseAssignment=@{
    "<office 365 licensnamn>:STANDARDPACK" = @(
        "YAMMER_ENTERPRISE", "SWAY","Deskless","POWERAPPS_O365_P1");
        "<office 365 licensnamn>:FLOW_FREE"=""
}

# Hämta ett objekt av konfigurations klassen.
$conf = New-SNSConfiguration -LicenseAssignment $LicenseAssignment

# Gruppnamn för alla ledare. Gruppen måste skapas i office 365 innan den kan användas här.
$conf.AllUsersGroupName='ledare'

# Standardsignatur för nya användare. Textvariant.
$conf.SignatureText=@"
Med vänliga hälsningar

<DisplayName>
"@

# Standardsignatur för nya användare. Html variant.
$conf.SignatureHtml=@"
<html>
    <head>
        <style type="text/css" style="display:none">
<!--
p   {margin-top:0; margin-bottom:0}
-->
        </style>
    </head>
    <body dir="ltr">
        <strong style="">
            <span class="ng-binding" style="color:rgb(00,00,00); font-size:12pt;">Med vänliga hälsningar</span>
        </strong>
        <br style="">
        <br style="">
        <div id="divtagdefaultwrapper" dir="ltr" style="font-size:12pt; color:#005496; font-family:Verdana">
            <table cellpadding="0" cellspacing="0" style="border-collapse:collapse; border-spacing:0px; background-color:transparent; font-family:Verdana,Helvetica,sans-serif">
                <tbody style="">
                    <tr style="">
                        <td valign="top" style="padding:0px 0px 6px; font-family:Verdana; vertical-align:top">
                            <strong style="">
                                <span class="ng-binding" style="color:rgb(00,54,96); font-size:14pt; font-style:italic"><DisplayName></span>
                            </strong>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>
    </body>
</html>
"@

# Rubrik för mailet till ny användare. Skickas till användarens primära e-postadress i Scoutnet.
$conf.NewUserEmailSubject="Ditt office 365 konto är skapat"

# Texten i mailet till ny användare. Skickas till användarens primära e-postadress i Scoutnet.
# Delarna <DisplayName>, <UserPrincipalName> och <Password> byts ut innan mailet skickas. 
$conf.NewUserEmailText=@"
Hej <DisplayName>!

Som ledare i scoutkåren så får du ett mailkonto i scoutkårens Office 365.
Kontot är bland annat till för att komma åt scoutkårens gemensamma dokumentarkiv .
Du får även en e-post adress <UserPrincipalName> som du kan använda för att skicka mail i kårens namn.

Ditt användarnamn är: <UserPrincipalName>
Ditt temporära lösenord är: <Password>
Lösenordet måste bytas första gången du loggar in.

Du kan logga in på Office 365 på https://portal.office.com för att komma åt din nya mailbox.

Mvh
Scoutkåren
"@

# Aktiverar Verbose logg. Standardvärde är silentlyContinue
#$VerbosePreference = "Continue"

# Vem ska mailet med loggen skickas till.
$logEmailToAddress = "karl.thoren@landvetterscout.se"

# Rubrik på mailet.
$logEmailSubject = "Maillist sync log"

# Domännam för scoutkårens office 365.
$conf.DomainName = "landvetterscout.se"

# Hashtable med id på Office 365 distributionsgruppen som nyckel.
# Distributions grupper som är med här kommer att synkroniseras.
$conf.MailListSettings = @{
    "utmanarna" = @{ # Namet på distributions gruppen i office 365. Används som grupp ID till Get-DistributionGroupMember.
        "scoutnet_list_id"= "4924"; # Listans Id i Scoutnet.
        "scouter_synk_option" = ""; # Synkoption för scouter. Giltiga värden är p,f,a eller tomt.
        "ledare_synk_option" = "@"; # Synkoption för ledare. Giltiga värden är @,-,t eller &.
        "email_addresses" = "","";  # Kommaseparerad lista med e-postadresser.
        "ignore_user" = "";         # Kommaseparerad lista med ScoutnetId som inte ska med i listan.
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

# Hämtar credentials för Scoutnet API och för Office 365.
try
{
    # Credentials för access till Office 365 och för att kunna skicka mail.
    $conf.Credential365 = Get-AutomationPSCredential -Name "MSOnline-Credentials" -ErrorAction "Stop"

    # Credentials för Scoutnets API api/group/customlists
    $conf.CredentialCustomLists = Get-AutomationPSCredential -Name 'ScoutnetApiCustomLists-Credentials' -ErrorAction "Stop"

    # Credentials för Scoutnets API api/group/memberlist
    $conf.CredentialMemberlist = Get-AutomationPSCredential -Name 'ScoutnetApiGroupMemberList-credentials' -ErrorAction "Stop"
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

try
{
    # Kör updateringsfunktionen.
    # Först uppdatera användare.
    Invoke-SNSUppdateOffice365User -Configuration $conf

    # Sen uppdatera maillistor.
    $NewValidationHash = SNSUpdateExchangeDistributionGroups -Configuration $conf -ValidationHash $ValidationHash
}
Catch
{
    Write-SNSLog -Level "Error" "Kunde inte köra uppdateringen. Error $_"
}
        
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
$bodyData = Get-Content -Path $conf.LogFilePath -Raw -Encoding UTF8 -ErrorAction "Continue"
Send-MailMessage -Credential $conf.Credential365 -From $conf.emailFromAddress -To $logEmailToAddress -Subject $logEmailSubject `
    -Body $bodyData -SmtpServer $conf.EmailSMTPServer -Port $conf.SmtpPort -UseSSL -Encoding UTF8 -ErrorAction "Continue"
