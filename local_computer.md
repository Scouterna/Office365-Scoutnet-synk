# Köra Office365-Scoutnet-synk på lokal dator

Synkronisering av Scoutnets e-postlistor till Office 365 kan även köras lokalt på en dator med Windows 10. Äldre windows versioner borde även de fungera, men det är inte testat.
Powershell 7 eller nyare på Linux funkar och är testat.

För att göra det så behöver du förbereda så din dator har de moduler som behövs.
Alla kommandon kös i powershell

## Steg som adminstratör
Behöver köras en gång på varje dator som ska användas för att köra import skriptet.
1. Starta en powershell instans som adminstratör.
1. Kör `Set-ExecutionPolicy RemoteSigned`
    1. Inställning för att kunna ladda in exchange online modulerna.


## Som användare
1. Starta en powershell instans. (Som din användare)
1. Kör `Install-Module -Name Office365-Scoutnet-synk -Scope CurrentUser` för att installera modulen.


Kopiera exemplet [MaillistSynk_local.ps1](MaillistSynk_local.ps1) till din dator och ändra inställningarna.

I delen "Skapa credentials för Scoutnet API." lägg in dina API nycklar.

## Köra synkningen
1. Starta en powershell instans. (Som din användare)
1. Kör `$VerbosePreference = "Continue"` för att få utskrifter på consollen.
1. Kör `MaillistSynk_local.ps1` för att köra synkningen.

## Uppdateringar
För att installera senaste versionen av Office365-Scoutnet-synk kör
```powershell
Update-Module -Name Office365-Scoutnet-synk -Scope CurrentUser
```