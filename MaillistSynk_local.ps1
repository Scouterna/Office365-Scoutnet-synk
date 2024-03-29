﻿#Requires -Version 5.1
#Requires -Modules @{ ModuleName="Office365-Scoutnet-synk"; ModuleVersion="2.0" }

# Lämplig inställning i Azure automation.
$ProgressPreference = "silentlyContinue"

# Aktiverar Verbose logg. Standardvärde är silentlyContinue
$VerbosePreference = "Continue"

# Vem ska mailet med loggen skickas ifrån. Byt ut till en adress som du har i din domän.
# Måste vara samma användare som loggar in.
$LogEmailFromAddress = "admin@scoutkåren"

# Vem ska mailet med loggen skickas till. Byt ut till en adminadress eller grupp.
$LogEmailToAddress = "admin@scoutkåren"

# Rubrik på mailet.
$LogEmailSubject = "Maillist sync log"

# Konfiguration av modulen.

# Licenser för nya användare.
# Exemplet nedan lägger in licenserna STANDARDPACK och FLOW_FREE. På STANDARDPACK applikationerna "YAMMER_ENTERPRISE", "SWAY","Deskless","POWERAPPS_O365_P1" avstängda.
# För att lista lisenser kör `Get-MgSubscribedSku -All | Format-List`
$LicenseAssignment=@{
    "STANDARDPACK" = @(
        "YAMMER_ENTERPRISE", "SWAY","Deskless","POWERAPPS_O365_P1");
        "FLOW_FREE"= @()
}

# Skapa ett konfigurationsobjekt och koppla licenshantering och vilken scoutnet maillist som hanterar användarnas konton.
# Byt ut maillist id till ID som matchar ledarna. Ta bort parametern -UserSyncMailListId om du vill att
# alla medlemmar med roller ska få ett konto.
$conf = New-SNSConfiguration -LicenseAssignment $LicenseAssignment -UserSyncMailListId "0000"

# Vem ska mailet till nya användare skickas ifrån. Byt ut till en adress som du har i din domän.
# Måste vara samma användare som loggar in.
$conf.EmailFromAddress = "admin@scoutkåren"

# Domännam för scoutkårens office 365.
$conf.DomainName = "scoutkåren.se"

# Hashtable med id på Office 365 distributionsgruppen som nyckel.
# Distributions grupper som är med här kommer att synkroniseras.
$conf.MailListSettings = @{
    "utmanarna" = @{ # Namet på distributions gruppen i office 365. Används som grupp ID till Get-DistributionGroupMember.
        "scoutnet_list_id"= "0001"; # Listans Id i Scoutnet.
        "scouter_synk_option" = ""; # Synkoption för scouter. Giltiga värden är p,f,a eller tomt.
        "ledare_synk_option" = "@"; # Synkoption för ledare. Giltiga värden är @,-,t eller &.
        "email_addresses" = "","";  # Kommaseparerad lista med e-postadresser.
        "ignore_user" = "";         # Kommaseparerad lista med ScoutnetId som inte ska med i listan.
    };
    "aventyrarna" = @{
        "scoutnet_list_id"= "0002";
        "scouter_synk_option" = ""; # Alla adresser
        "ledare_synk_option" = "@"; # Bara office 365 adresser
        "email_addresses" = "";
    };
    "upptackare" = @{
        "scoutnet_list_id"= "0004";
        "scouter_synk_option" = ""; # Alla adresser
        "ledare_synk_option" = "@"; # Bara office 365 adresser
        "email_addresses" = "";
    };
    "sparare" = @{
        "scoutnet_list_id"= "0005";
        "scouter_synk_option" = ""; # Alla adresser
        "ledare_synk_option" = "@"; # Bara office 365 adresser
        "email_addresses" = "";
    }
}

# Gruppnamn för alla ledare. Gruppen måste skapas i office 365 innan den kan användas här.
$conf.AllUsersGroupName='ledare'

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

# Rubrik för e-brevet som skickas till användarens nya e-postadress.
$conf.NewUserInfoEmailSubject="Välkommen till scoutkårens Office 365"

# Texten för e-brevet som skickas till användarens nya e-postadress.
$conf.NewUserInfoEmailText=@"
Hej <DisplayName>!

Som ledare i scoutkåren har du nu fått ett konto i scoutkårens Office 365.
Kontot är bland annat till för att komma åt scoutkårens gemensamma dokumentarkiv som finns i sharepoint.
Du har en e-post adress <UserPrincipalName> som du kan använda för att skicka mail i kårens namn.

Länkar som är bra att hålla koll på:
Scoutnet: https://www.scoutnet.se

Mvh
Scoutkåren
"@

# Texten i det automatiska svaret, om man skickar brev till medlem som slutat.
# Ta bort hela parametern om du inte vill ha ett automatiskt svar för medlemmar som slutat.
$conf.DisabledAccountsAutoReplyText=@"
<html><body>
<DisplayName> är inte längre medlem i scoutkåren.<br>
Mvh<br>
Scoutkåren
</body></html>
"@

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

# Här börjar själva skriptet.
Remove-Item -Path $conf.LogFilePath

# Skapa credentials för Scoutnet API och för Office 365.
try
{
    # Användarnamn för Scoutnets API. Användarnamnet är Kår-ID för webbtjänster som står på sidan Webbkoppling.
    $UserName = "000000"

    # Credentials för Scoutnets API api/group/customlists
    $apiNyckel = ConvertTo-SecureString -String "ApiNyckel" -AsPlainText -Force
    $conf.CredentialCustomlists = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($UserName, $apiNyckel)  -ErrorAction "Stop"

    # Credentials för Scoutnets API api/group/memberlist
    $apiNyckel = ConvertTo-SecureString -String "ApiNyckel" -AsPlainText -Force
    $conf.CredentialMemberlist = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($UserName, $apiNyckel)  -ErrorAction "Stop"
}
Catch
{
    Write-SNSLog -Level "Error" "Kunde inte hämta skapa credentials. Error $_"
    throw
}

try
{
    # Logga in på office 365
    Connect-SnSOffice365 -Configuration $conf -ErrorAction "Stop"
}
Catch
{
    Write-SNSLog -Level "Error" "Kunde logga in på office 365 Error $_"
    throw
}

# Kör updateringsfunktionen.
try
{
    # Först uppdatera användare.
    Invoke-SNSUppdateOffice365User -Configuration $conf
}
Catch
{
    Write-SNSLog -Level "Error" "Kunde inte köra uppdateringen av användare. Fel: $_"
}

try
{
    # Sen uppdatera maillistor.
    $NewValidationHash, $mailListData = SNSUpdateExchangeDistributionGroups -Configuration $conf -ValidationHash "Tom" -ReturnMaildata
}
Catch
{
    Write-SNSLog -Level "Error" "Kunde inte köra uppdateringen av distributionsgrupper. Fel: $_"
}

$bodyData = Get-Content -Path $conf.LogFilePath -Raw -Encoding UTF8 -ErrorAction "Continue"
$params = @{
    Message = @{
        Subject = $LogEmailSubject
        Body = @{
            ContentType = "Text"
            Content = $bodyData
        }
        ToRecipients = @(
            @{
                EmailAddress = @{
                    Address = $LogEmailToAddress
                }
            }
        )
    }
    SaveToSentItems = "false"
}

Send-MgUserMail -UserId $LogEmailFromAddress -BodyParameter $params


try
{
    # Logga ut ifrån ExchangeOnline.
    Disconnect-SnSOffice365
}
Catch
{
    Write-SNSLog -Level "Error" "Utloggning ifrån ExchangeOnline returnerade felet $_"
    throw
}
