# Office365-Scoutnet-synk

Synkronisering av Scoutnets e-postlistor till Office 365.

Du kan med de här funktionerna synkronisera användarkonton med personer i Scoutnet
som har en funktionärsroll samt synkronisera distributions listor med
e-postlistor i Scoutnet.

Modulen `Office365-Scoutnet-synk` är tänkt att användas i Azure automation,
och köras ifrån en runbook. Microsoft Azure Sponsorship ingår när man får
Microsoft Office 365 non-profit.

Azure gratiskonto kan troligtvis också användas, då det ingår 500 minuter Azure automation.

Modulen går även att köra på en dator som har minst Powershell 5.1 installerad.

Vid problem, fel, frågor eller tips på förbättringar eller fler funktioner
som du saknar; lägg ett ärende under `Issues` eller mejla karl.thoren@scouterna.se

I bland kommer det ny funktionalitet, så håll utkik på en ny version genom
att trycka på knappen **Watch** uppe till höger på sidan
för att du kunna bli notifierad vid en ny version.

Du kan ladda ner den senaste versionen via
<https://github.com/scouternasetjanster/Office365-Scoutnet-synk/releases/latest>
och där kan du också ser vilken funktionalitet som är ny i respektive version.

Eller via https://www.powershellgallery.com/packages/Office365-Scoutnet-synk där senaste
versionen alltid är publicerad.

Läs filen [README.md](README.md) för instruktion om installation och funktionalitet.

## Inställningar

### Generella inställningar

I non-profit portalen aktivera scoutkårens `Microsoft Azure Sponsorship Subscription`
<https://nonprofit.microsoft.com/offers/azure> och sen kan du skapa ett
"Azure Automation Account" som kommer att köra dina skript.
Hur du gör är beskrivet här <https://blog.kloud.com.au/2016/08/24/schedule-office-365-powershell-tasks-using-azure-automation>

1. Skapa ett `Azure Automation Account` och koppla det till
    "Microsoft AzureSponsorship Subscription".
    1. Bra namn är `Scoutnet-synk` på kontot och resursgruppen.
    1. Välj `North Europe` som Location.

1. Lägg till `MSOnline` modulen. Behövs för att kunna skapa användare.

1. Lägg till `Office365-Scoutnet-synk` som en modul.
   1. Gå in på https://www.powershellgallery.com/packages/Office365-Scoutnet-synk
   1. Välj Azure Automation och tryck på "Deploy to Azure Automation"
      Modulen kommer nu installeras på Azure Automation.

1. I Scoutnet aktivera APIet under Webbkoppling (API-nycklar och endpoints).
    Modulen behöver ha tillgång till:
    - *Get a detailed csv/xls/json list of all members*. (api/group/memberlist)
    - *Get a csv/xls/json list of members, based on mailing lists you have set up*.
        (api/group/customlists)

1. I Azure resursgruppen skapa `Credential Asset` för varje API nyckel.
    Användarnamnet är Kår-ID för webbtjänster som står på sidan Webbkoppling.
    Lösenordet är API-nyckeln.
    1. Credential Asset: `ScoutnetApiCustomLists`, API-nyckel för api/group/customlists
    1. Credential Asset: `ScoutnetApiGroupMemberList`, API-nyckel för api/group/memberlist

1. Skapa även en `Credential Asset` med en användare som har adminrättigheter
    på scoutkårens office 356.
    1. Credential Asset: `MSOnline-Credentials`, konto som har adminrättigheter
        på scoutkårens office 356. Rekommendationen är att det är ett separat
        konto som är kopplat till scoutkårens onmicrosoft.com domän.
        T.ex administrator@scoutkaren.onmicrosoft.com

### Synkronisera grupper

1. Logga in på office 365 adminkonsollen och skapa de distributions
    listor du vill använda.
    1. Typen ska vara `Distribution list` för att synkroniseringen ska fungera.
        :warning: **Office365 grupper stöds ej**.
    1. Namnet kan vara beskrivande, men skriv ett alias som är kort och bara har
        **ASCII**  tecken i sig.

1. I Scoutnet skapa `Fördefinierade listor` för distributions listorna.
    T.ex en lista för Spårare som är avsedd för att skicka brev till scouternas föräldrar.
    För att synkroniseringen ska fungera smidig skapa följande regler på varje lista,
    där regeln *ledare* hanteras av [ledare_synk_option](#ledare_synk_option).
    Övriga regler styrs av [scouter_synk_option](#scouter_synk_option).
    - **Ledare:** Regel som matchar ledarna på avdelningen. Döp den till *ledare*.
        Hur adresser synkroniseras styrs av [ledare_synk_option](#ledare_synk_option).
    - **Assistenter:** Regel som matchar assistenterna på avdelningen.
        Döp den till *assistenter*.
        Hur adresser synkroniseras styrs av [scouter_synk_option](#scouter_synk_option).
    - **Scouter:** Regel som matchar scouterna på avdelningen. Döp den till *scouter*.
        Hur adresser synkroniseras styrs av [scouter_synk_option](#scouter_synk_option).

    Regelnamnen behöver inte användas om du bara vill styra med [scouter_synk_option](#scouter_synk_option).

1. I Azure automation skapa runbooken `MaillistSynk` för synkroniseringen.
    Typen ska vara `PowerShell Runbook`

1. I Azure automation under `Shared Resources`, skapa variabeln
    `ScoutnetMailListsHash` av typen `string`.

1. Kopiera koden ifrån exemplet [MaillistSynk.ps1](MaillistSynk.ps1).

1. Ändra inställningarna så att de matchar scoutkårens Scoutnetprofil och Office 365.
    T.ex vilka listor som ska uppdateras.

1. Prova att köra `MaillistSynk`.

1. När `MaillistSynk` fungerar publicera runbooken.

1. Azure automation under `Shared Resources`, skapa en `schedule` för att
    regelbundet köra MaillistSynk.
    1. Rekommendationen är att köra nattetid (kl 3 eller 4), då det kan ta en
        stund att köra MaillistSynk.

## Manual

### Office 365 användarkonton - synkronisering med Scoutnet

Synkroniserar personer som som är med på e-postlistor från Scoutnet genom att
man anger id-numret för e-postlistan i workbooken. Det går bra att ange
flera e-postlistor med kommatecken. Se exemplen i filen.

Om inget id-nummer för en e-postlista anges så tolkar programmet det som alla
personer i kåren som har en avdelningsroll eller roll på kårnivå och skapar
användarkonton på kårens Office 365 åt dem.

Om ett användarkonto vid nästkommande synkronisering ej matchar någon person
som synkroniseras så inaktiveras konto.
Om personen senare matchas aktiveras kontot igen. Det är bara konton som finns
i säkerhetsgruppen `Scoutnet` i Office 365 som berörs vid en synkronisering.
Användarkonton skapas på formen fornamn.efternamn@domännamn.se

Om det finns personer som har samma namn (för- och efternamn) angivet i Scoutnet
kommer de som skapas som nr2 osv skapas på formen fornamn.efternamn.x@domännamn.se
där X motsvarar en siffra från 1-5.

Exchange fältet `CustomAttribute1` innehåller Scoutnet ID.
Lägger du till användare manuellt och de ska kunna automatisk komma med i
distributionsgrupper så behöver du fylla i `CustomAttribute1`
med personens Scoutnet ID.

Funktionen för användarsynkronisering heter `Invoke-SNSUppdateOffice365User`.

#### Inställningar

I exempelfilen så är det några inställningar att ändra:

- Ändra kårens domän namn på variabeln `$conf.DomainName`.
- Ändra mottagaradress i variabeln `LogEmailToAddress`.
- Ändra licensinställningen i variabeln `LicenseAssignment`.
- Ändra eller ta bort grupp för alla ledare i variabeln `$conf.AllUsersGroupName`.
- Ändra standardsignatur i `$conf.SignatureText` och `$conf.SignatureHtml`.
- Ändra rubrik och test på välkomstbrev i `$conf.NewUserEmailSubject` och `$conf.NewUserEmailText`.
    Välkomstbrevet skickas till medlemmes primära e-postadress i scouten.
- Ändra rubrik och test på informationsbrev i `$conf.NewUserInfoEmailSubject` och `$conf.NewUserInfoEmailText`.
    Informationsbrevet skickas till användarens nya e-postadress.
    Kan t.ex innehålla information om var du kan hitta sharepointsiten 

### Office 365 distributionsgrupper - synkronisering med Scoutnet

Synkronisering av Office 365 distributionsgrupper med e-postlistor i Scoutnet.
Där en distributionsgrupp är kopplad till en e-postlista i Scoutnet.

#### mailListSettings

I variabel `$conf.MailListSettings` ställ in de Office 365 distributionsgrupper
som ska synkroniseras med Scoutnet.
Namnet måste gå att matcha mot en distributionsgrupp i Office365.

Exempel:

```powershell
$conf.MailListSettings = @{
    "utmanare" = @{ # Namnet på distributions gruppen i office 365.
        "scoutnet_list_id"= "0001"; # Listans Id i Scoutnet.
        "scouter_synk_option" = ""; # Alla adresser i scoutnet.
        "ledare_synk_option" = "@"; # Bara office 365 adresser
        "email_addresses" = "test1@domain.se","test2@domain.se";  # Lista med e-postadresser.
        "ignore_user" = "12345", "54321"; # Lista med ScoutnetId som ignoreras.
    };
    "aventyrare" = @{ Namnet på distributions gruppen i office 365.
        "scoutnet_list_id"= "0002";
        "scouter_synk_option" = ""; # Alla adresser i scoutnet.
        "ledare_synk_option" = "@"; # Bara office 365 adresser
        "email_addresses" = "";
    };
    "lager" = @{ Namnet på distributions gruppen i office 365.
        "statisk_lista" = "Ja"; # Listat är en statisk Scoutnetlista.
        "scoutnet_list_id"= "0003";
        "ledare_synk_option" = "t"; # Office 365 adresser eller scoutnet adress.
        "scouter_synk_option" = ""; # Används inte för statiska listor
        "email_addresses" = "";
    };
```

##### scouter_synk_option

I fältet `scouter_synk_option` kan du för respektive
distributionsgrupp ange följande:

- "@" Lägg till personens Office365-konto om den har något, annars hoppa över personen.
- "t" Lägg till personens Office365-konto om den har något, annars personens
    e-postadress som listad i Scoutnet.
- "-" Lägg endast till personens e-postadress som listad i Scoutnet.
- "&" Lägg till både personens e-postadress som listad i Scoutnet samt
    Office365-konto om den har något.
- Standard för `scouter_synk_option` är "-" Lägg endast till personens
e-postadress som listad i Scoutnet.

Det går också att ställa in i detta fält vilka e-postadressfält från scoutnet
som ska läggas till:

- "p" Lägg till en medlems primära e-postadress
- "f" Lägg till de e-postadresser som är angivet i fälten Anhörig 1,2.
    Alltså vanligtvis föräldrarna.
- "a" Lägg till en medlems alternativ e-postadress.
- Om man inte anger något används fälten primär e-postadress, anhörig 1,
    anhörig 2 och alternativ e-postadress.

##### ledare_synk_option

I fältet `ledare_synk_option` kan du för respektive
distributionsgrupp ange följande:

- "@" Lägg till personens Office365-konto om den har något, annars hoppa över personen.
- "t" Lägg till personens Office365-konto om den har något, annars personens
    e-postadress som listad i Scoutnet.
- "-" Lägg endast till personens primära e-postadress som listad i Scoutnet.
- "&" Lägg till både personens primära e-postadress som listad i Scoutnet samt
    Office365-konto om den har något.
- Standard för `ledare_synk_option` är "@" Lägg till personens
    Office365-konto om den har något, annars hoppa över personen.

##### email_addresses

Fältet `email_addresses` används för att lägga till extra e-postadresser
till en distributionsgrupp.
Lägg adresserna kommaseparerade.

Exempel:

```powershell
"email_addresses" = "test1@domain.se","test2@domain.se";
```

##### ignore_user

Fältet `ignore_user` används för att inte lägga till specifika användare
till en distributionsgrupp. Används om en e-postlista i Scoutnet matchar lite
mer än vad du vill ha med.

Lägg ScoutnetId kommaseparerade.

Exempel:

```powershell
"ignore_user" = "12345", "54321";
```

##### statisk_lista

Fältet `statisk_lista` används för att markera att listan är en statiskt konfigurerad lista i Scoutnet.
Statiska listor i Scoutnet hanteras bara med regelerna ifrån [ledare_synk_option](#ledare_synk_option).

Exempel:

```powershell
"statisk_lista" = "Ja";
```

## Ny version

- Uppdatering av modulen sker genom att ladda ner en ny version och installera modulen.
- Du hittar senaste versionen av modulen på
    https://www.powershellgallery.com/packages/Office365-Scoutnet-synk eller
    <https://github.com/scouternasetjanster/Office365-Scoutnet-synk/releases/latest>
    och där kan du också ser vilken funktionalitet som är ny i respektive version
    och om du behöver göra något för att uppdatera förutom att uppdatera modulen.
- Du kan hålla dig uppdaterad med nya versioner genom att om du är inloggad
    på Github trycka på knappen **Watch** uppe till höger på sidan för att då
    kunna bli notifierad vid ny version.

## Hjälp

1. Kolla i loggfilen så kanske du lyckas se vad problemet är.
1. Det kan hända att det finns en bugg i den versionen av programmet som du kör vilket
   givet vissa specifika omständigheter yttrar sig för just dig. Se till att du
   har den senaste versionen av programmet.
1. Lägg ett ärende under `Issues` eller mejla.

## Tekniska förtydliganden

### Uppdatering av distributions listor med hjälp av SNSUpdateExchangeDistributionGroups

SNSUpdateExchangeDistributionGroups hämtar maillist medlemmar ifrån Scoutnet och
uppdaterar motsvarande distributions grupp i office 365.
Då alla medlemmar i en distributions grupp måste finnas i Exchange som användare
eller kontakt så kommer en kontakt för varje extern e-postadress att skapas.

**Städning:** Funktionen städar och tar bort kontakter till Scouter som slutat.
Det här är för att kunna följa GDPR. Kontakter till Scouter som slutat ska tas bort.

#### SNSUpdateExchangeDistributionGroups beteende

1. Validera att Scoutnet är uppdaterat. Om Scoutnet inte är uppdaterat avbryts synkroniseringen.
1. Ta bort alla medlemmar ifrån de grupper som ska synkroniseras.
    1. Om kontakten inte används längre, radera den ifrån Exchange. Det här är
    för att kunna följa GDPR. Kontakter till Scouter som slutat ska tas bort.
1. För varje distributions grupp lägg till de medlemmar som e-postlistan i scoutnet
    innehåller. För e-postadresser som är externa skapas det en kontakt i Exchange.
    1. Om office 365 adress är begärd i inställningarna så letas användaren med
    matchande Scoutnet ID upp.
    Exchange fältet `CustomAttribute1` ska innehålla Scoutnet ID.
    1. Lägg in eventuella extra e-postadresser.

### E-postadresser genereras på följande sätt

1. För och efternamn görs om till gemener och tar bort alla mellanslag och
   andra tomrum som är skrivet i de fälten i Scoutnet.
1. Därefter görs bokstäverna om till bokstäver som lämpar sig för e-postadresser.
   T.ex åäö blir aao.
1. Om det därefter finns några fler konstiga tecken kvar som inte lämpar sig för
   e-postadress så tas de bort.
1. Om det finns personer som har samma namn (för- och efternamn) angivet i Scoutnet
kommer de som skapas som nr2 osv skapas på formen fornamn.efternamn.X@domännamn.se
där X motsvarar en siffra från 1-5.
