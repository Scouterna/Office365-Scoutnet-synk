# Office365-Scoutnet-synk

Du kan med dessa program göra följande

1. Synkronisera [Microsoft365][Microsoft365]-användarkonton hos kåren
   med ledare från Scoutnet. Se [manualen för användarkonton][user-accounts].
1. Synkronisera [Microsoft365][Microsoft365]-distributionsgrupper med e-postlistor
   i Scoutnet för att använda som e-postlistor i [Microsoft365][Microsoft365].
   Se [manualen för distributionsgrupper][distribution-groups].
1. Synkronisera Microsoft 365-grupper och Teams från konfigurerade
   Scoutnet-maillistor och roller i Scoutnet. Se
   [manualen för Microsoft 365-grupper][m365-groups].
1. Synkronisera delade mailboxar från konfigurerade Scoutnet-maillistor
   och roller i Scoutnet. Se [manualen för delade mailboxar][shared-mailboxes].
1. Få automatiska rapporter om vilka medlemmar som behöver förnya
   kursen Trygga Möten och behöver förnya registerutdraget.
1. Skicka påminnelser till medlemmar som behöver förnya
   kursen Trygga Möten och behöver förnya registerutdraget. Läs mer om
   [rapporter och påminnelser om obligatoriska moment][mandatory-moments].

Synkroniseringen styrs via en [konfigurationen][configuration]
som sätts upp i en run-book som går att köra i [Microsoft Azure][azure-runbook]
eller [på en lokal dator][local-run].

- Vid problem, fel, frågor eller tips på förbättringar eller fler funktioner
  som du saknar
    - Lägg ett ärende under [Issues][issues]
    - eller skicka ett mail till <karl.thoren@scouterna.se>.
    - Titta även under [wiki/Support][support] sidan.

- I bland kommer det ny funktionalitet, så håll utkik på en ny version genom att
  trycka på knappen **Watch** uppe till höger på sidan för att kunna bli
  notifierad vid en ny version.

- [Ny version][new-version]
  Hur du uppdaterar till en ny version hittar du på [ny version][new-version].

## [Dokumentation - läs mer här][wiki]

[Microsoft365]: # "Tidigare namn Office 356"
[new-version]: https://github.com/Scouterna/Office365-Scoutnet-synk/wiki/Ny-version
[support]: https://github.com/Scouterna/Office365-Scoutnet-synk/wiki/Support
[issues]: https://github.com/Scouterna/Office365-Scoutnet-synk/issues
[wiki]: https://github.com/Scouterna/Office365-Scoutnet-synk/wiki
[user-accounts]: https://github.com/Scouterna/Office365-Scoutnet-synk/wiki/Manual-Anv%C3%A4ndarkonton
[distribution-groups]: https://github.com/Scouterna/Office365-Scoutnet-synk/wiki/Manual-Distributionsgrupper
[m365-groups]: https://github.com/Scouterna/Office365-Scoutnet-synk/wiki/Manual-Microsoft-365-grupper
[shared-mailboxes]: https://github.com/Scouterna/Office365-Scoutnet-synk/wiki/Manual-Delade-mailboxar
[mandatory-moments]: https://github.com/Scouterna/Office365-Scoutnet-synk/wiki/Manual-Obligatoriska-moment
[configuration]: https://github.com/Scouterna/Office365-Scoutnet-synk/wiki/Manual-Konfiguration
[azure-runbook]: https://github.com/Scouterna/Office365-Scoutnet-synk/wiki/Microsoft-Azure-Runbook
[local-run]: https://github.com/Scouterna/Office365-Scoutnet-synk/wiki/K%C3%B6rning-p%C3%A5-dator
