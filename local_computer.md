# Köra Office365-Scoutnet-synk på lokal dator

Synkronisering av Scoutnets e-postlistor till Office 365 kan även köras lokalt på en dator med Windows 10. Äldre windows versioner borde även de fungera, men det är inte testat.
Powershell 6 på Linux funkar inte, då `MSOnline` modulen kräver en windows maskin.

För att göra det så behöver du förbereda så din dator har de moduler som behövs.
Alla kommandon kös i powershell

## Steg som adminstratör
Behöver köras en gång på varje dator som ska användas för att köra import skriptet.
1. Starta en powershell instans som adminstratör.
1. Kör `Set-ExecutionPolicy RemoteSigned`
    1. Inställning för att kunna ladda in exchange online modulerna.


## Som användare
1. Starta en powershell instans. (Som din användare)
1. Kör `Install-Module -Name MSOnline -Scope CurrentUser` för att installera modulen.
1. Kör `Install-Module -Name Office365-Scoutnet-synk -Scope CurrentUser` för att installera modulen.


Kopiera exemplet [MaillistSynk_local.ps1](MaillistSynk_local.ps1) till din dator och ändra inställningarna.

I delen "Skapa credentials för Scoutnet API och för Office 365." lägg in dina API nycklar.

I powershell kör MaillistSynk_local.ps1 för att köra synkningen.
