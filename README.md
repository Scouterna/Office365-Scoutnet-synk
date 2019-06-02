# Office365-Scoutnet-synk
Synkronisering av Scoutnets e-postlistor till Office 365.

Du kan med de här funktionerna synkronisera användarkonton med personer i Scoutnet som
har en funktionärsroll samt synkronisera distributions listor med e-postlistor i Scoutnet.

Modulen Office365-Scoutnet-synk är tänkt att användas i Azure automation, och köras ifrån en runbook.
Microsoft Azure Sponsorship ingår när man får Microsoft Office 365 non-profit.

Azure gratiskonto kan troligtvis också användas, då det ingår 500 minuter Azure automation.

Modulen går även att köra på en dator som har minst Powershell 5.1 installerad.

Vid problem, fel, frågor eller tips på förbättringar eller fler funktioner som du saknar;
lägg ett ärende under "Issues" eller mejla karl.thoren@landvetterscout.se

I bland kommer det ny funktionalitet, så håll utkik på en ny version genom att trycka på knappen
**Watch** uppe till höger på sidan för att du kunna bli notifierad vid en ny version.

Du kan ladda ner den senaste versionen via
https://github.com/scouternasetjanster/Office365-Scoutnet-synk/releases/latest och där kan
du också ser vilken funktionalitet som är ny i respektive version.

Läs filen README.md för instruktion om installation och funktionalitet.

## Inställningar

### Generella inställningar
I non-profit portalen aktivera scoutkårens "Microsoft Azure Sponsorship Subscription" https://nonprofit.microsoft.com/offers/azure
och sen kan du skapa ett "Azure Automation Account" som kommer att köra dina skript.
Hur du gör är beskrivet här https://blog.kloud.com.au/2016/08/24/schedule-office-365-powershell-tasks-using-azure-automation

1. Skapa ett "Azure Automation Account" och koppla det till "Microsoft Azure Sponsorship Subscription".
    1. Bra namn är "Scoutnet-synk" på kontot och resursgruppen.
    1. Välj "North Europe" som Location.

1. Lägg till MSOnline modulen. Behövs för att kunna skapa användare.

1. Lägg till Office365-Scoutnet-synk som en modul.

1. I Scoutnet aktivera APIet under Webkoppling (API-nycklar och endpoints).
    Modulen behöver ha tillgång till:
    - *Get a detailed csv/xls/json list of all members*. (api/group/memberlist)
    - *Get a csv/xls/json list of members, based on mailing lists you have set up*. (api/group/customlists)
 
1. I Azure resursgruppen skapa "Credential Asset" för varje API nyckel.
    Användarnamnet är Kår-ID för webbtjänster som står på sidan Webkoppling.
    Lösenordet är API-nyckeln.
    1. Credential Asset: "ScoutnetApiCustomLists", API-nyckel för api/group/customlists
    1. Credential Asset: "ScoutnetApiGroupMemberList", API-nyckel för api/group/memberlist

1. Skapa även en "Credential Asset" med en användare som har adminrättigheter på scoutkårens office 356.
    1. Credential Asset: "MSOnline-Credentials", konto som har adminrättigheter på scoutkårens office 356.

### Synkronisera grupper
1. Logga in på office 365 adminkonsollen och skapa de distributions listor du vill använda.
    1. Typen ska vara "Distribution list" för att synkroniseringen ska fungera. Office365 grupper stöds ej.
    1. Namnet kan vara beskrivande, men skriv ett alias som är kort och bara har ASCII tecken i sig.

1. I Scoutnet skapa "Fördefinierade listor" för distributions listorna.
    T.ex en lista för Spårare som är avsedd för att skicka brev till scouternas föräldrar.
    För att synkroniseringen ska fungera smidig skapa följande regler på varje lista.
    * Ledare: Regel som matchar ledarna på avdelningen. Döp den till "ledare".
    * Assistenter: Regel som matchar assistenterna på avdelningen. Döp den till "assistenter".
    * Scouter: Regel som matchar scouterna på avdelningen. Döp den till "scouter".

1. I Azure automation skapa runbooken "MaillistSynk" för synkroniseringen. Typen ska vara "PowerShell Runbook"

1. I Axure automation under "Shared Resources", skapa variabeln "ScoutnetMailListsHash" av typen string.

1. Kopiera koden ifrån exemplet MaillistSynk.ps1.

1. Ändra inställningarna så att de matchar. T.ex vilka listor som ska uppdateras.

1. Prova att köra MaillistSynk.

1. När MaillistSynk fungerar publicera runbooken.

1. Axure automation under "Shared Resources", skapa en "schedule" för att regelbundet köra MaillistSynk.
    1. Rekommendationen är att köra nattetid (kl 3 eller 4), då det kan ta en stund att köra MaillistSynk.
    
## Ny version
- Uppdatering av modulen sker genom att ladda ner en ny version och installera modulen.
- Du hittar senaste versionen av modulen på 
  https://github.com/scouternasetjanster/Office365-Scoutnet-synk/releases/latest och där kan
  du också ser vilken funktionalitet som är ny i respektive version och om du behöver göra
  något för att uppdatera förutom att uppdatera modulen.
- Du kan hålla dig uppdaterad med nya versioner genom att om du är inloggad på Github trycka
  på knappen **Watch** uppe till höger på sidan för att då kunna bli notifierad vid ny version.

## Hjälp
1. Om problem uppstått när du kört programmet tidsinställt, testa då att köra
   programmet manuellt en gång och se om felet uppstår då också.
1. Kolla i loggfilen så kanske du lyckas se vad problemet är.
1. Det kan hända att det finns en bugg i den versionen av programmet som du kör vilket
   givet vissa specifika omständigheter yttrar sig för just dig. Se till att du har den
   senaste versionen av programmet.
1. Lägg ett ärende under "Issues" eller mejla.


## Tekniska förtydliganden
### Uppdatering av distributions listor med hjälp av SNSUpdateExchangeDistributionGroups
SNSUpdateExchangeDistributionGroups hämtar maillist medlemmar ifrån Scoutnet och uppdaterar
motsvarande distributions grupp i office 365. 
Då alla medlemmar i en distributions grupp måste finnas i Exchange som användare eller
kontakt så kommer en kontakt för varje extern epostadress att skapas.

**Städning:** Funktionen städar och tar bort kontakter till Scouter som slutat.
Det här är för att kunna följa GDPR. Kontakter till Scouter som slutat ska tas bort.

#### Funktionens beteende.
1. Validera att Scoutnet är uppdaterat. Om Scoutnet inte är uppdaterat avbryts synkroniseringen.
1. Ta bort alla medlemmar ifrån de grupper som ska synkroniseras.
    1. Om kontakten inte används längre, radera den ifrån Exchange.
       Det här är för att kunna följa GDPR. Kontakter till Scouter som slutat ska tas bort.
1. För varje distributions grupp lägg till de medlemmar som maillistan i scoutnet innehåller.
    För e-postadresser som är externa skapas det en kontakt i Exchange.
    1. För ledare (matchade på regeln ledare) så läggs deras office 365 konto in som medlem i distributions gruppen
    1. Lägg in eventuella admin mail adresser.

### E-postadresser genereras på följande sätt.
1. För och efternamn görs om till gemener och tar bort alla mellanslag och
   andra tomrum som är skrivet i de fälten i Scoutnet.
1. Därefter görs bokstäverna om till bokstäver som lämpar sig för e-postadresser.
   T.ex åäö blir aao.
1. Om det därefter finns några fler konstiga tecken kvar som inte lämpar sig för
   e-postadress så tas de bort



