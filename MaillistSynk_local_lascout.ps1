#Requires -Version 5.1
#Requires -Modules @{ ModuleName="Office365-Scoutnet-synk"; ModuleVersion="2.0" }

# Lämplig inställning i Azure automation.
$ProgressPreference = "silentlyContinue"

# Aktiverar Verbose logg. Standardvärde är silentlyContinue
$VerbosePreference = "Continue"

# Vem ska mailet med loggen skickas ifrån.
$LogEmailFromAddress = "info@landvetterscout.se"

# Vem ska mailet med loggen skickas till.
$LogEmailToAddress = "karl.thoren@landvetterscout.se"

# Rubrik på mailet.
$LogEmailSubject = "Maillist sync log"

# Konfiguration av modulen.

# Licenseer för standardkonto.
$LicenseAssignment=@{
    "O365_BUSINESS_ESSENTIALS" = @(
        "YAMMER_ENTERPRISE", "SWAY");
        "FLOW_FREE" = @()
}

# Skapa ett konfigurationsobjekt och koppla licenshantering och vilken scoutnet maillist som hanterar användarnas konton.
$conf = New-SNSConfiguration -LicenseAssignment $LicenseAssignment -UserSyncMailListId "5004"

# Vem ska mailet till nya användare skickas ifrån.
$conf.EmailFromAddress = "info@landvetterscout.se"

# Domännam för scoutkårens office 365.
$conf.DomainName = "landvetterscout.se"

# Hashtable med id på Office 365 distributionsgruppen som nyckel. 
# Distributions grupper som är med här kommer att synkroniseras.
$conf.MailListSettings = @{
    "rover@landvetterscout.se" = @{
        "scoutnet_list_id"= "13491";
        "scouter_synk_option" = "p&"; # Office 365 adresser och primär i scoutnet.
        "ledare_synk_option" = "@"; # Office 365 adresser.
        "email_addresses" = "";  # Lista med e-postadresser.
    };
    "utmanarna" = @{ # Namet på distributions gruppen i office 365. Används som grupp ID till Get-DistributionGroupMember.
        "scoutnet_list_id"= "4924"; # Listans Id i Scoutnet.
        "scouter_synk_option" = ""; # Synkoption för scouter. Giltiga värden är p,f,a,t eller tomt.
        "ledare_synk_option" = "@"; # Synkoption för ledare. Giltiga värden är @,-,t eller &.
        "email_addresses" = "";  # Lista med e-postadresser.
    };
    "lainfo" = @{
        "scoutnet_list_id"= "4989";
        "scouter_synk_option" = "p&"; # Office 365 adresser och primär i scoutnet.
        "ledare_synk_option" = "&"; # Office 365 adresser och primär i scoutnet.
        "email_addresses" = "";  # Lista med e-postadresser.
        "ignore_user" = "3109304", "3149407"; # Användare som inte vill ha mail ifrån lainfo.
    };
    "over18@landvetterscout.se" = @{
        "scoutnet_list_id"= "16218";
        "scouter_synk_option" = "p&"; # Office 365 adresser och primär i scoutnet.
        "ledare_synk_option" = "&"; # Office 365 adresser och primär i scoutnet.
        "email_addresses" = "";  # Lista med e-postadresser.
        "ignore_user" = "3109304", "3149407"; # Användare som inte vill ha mail ifrån lainfo.
    };
    "rovdjuren" = @{
        "scoutnet_list_id"= "4923";
        "scouter_synk_option" = ""; # Alla adresser.
        "ledare_synk_option" = "@"; # Bara office 365 adresser.
        "email_addresses" = "";
    };
    "upptackare" = @{
        "scoutnet_list_id"= "4922";
        "scouter_synk_option" = ""; # Alla adresser.
        "ledare_synk_option" = "@"; # Bara office 365 adresser.
        "email_addresses" = "";
    };
    "krypen" = @{
        "scoutnet_list_id"= "4900";
        "scouter_synk_option" = ""; # Alla adresser.
        "ledare_synk_option" = "@"; # Bara office 365 adresser.
        "email_addresses" = "";
    };
    "ravarna" = @{
        "scoutnet_list_id"= "4904";
        "scouter_synk_option" = ""; # Alla adresser.
        "ledare_synk_option" = "@"; # Bara office 365 adresser.
        "email_addresses" = "";
    };
    "spararledare@landvetterscout.se" = @{
        "scoutnet_list_id"= "5012";
        "scouter_synk_option" = "tp"; # Föredra office 365 adresser för primär i scoutnet.
        "ledare_synk_option" = "@"; # Bara office 365 adresse.
        "email_addresses" = "";
    };
    "ledare-krypen@landvetterscout.se" = @{
        "scoutnet_list_id"= "10923";
        "scouter_synk_option" = "tp"; # Föredra office 365 adresser för primär i scoutnet.
        "ledare_synk_option" = "@"; # Bara office 365 adresse.
        "email_addresses" = "";
    };
    "ledare-ravarna@landvetterscout.se" = @{
        "scoutnet_list_id"= "10924";
        "scouter_synk_option" = "tp"; # Föredra office 365 adresser för primär i scoutnet.
        "ledare_synk_option" = "@"; # Bara office 365 adresse.
        "email_addresses" = "";
    };
    "upptackarledare@landvetterscout.se" = @{
        "scoutnet_list_id"= "5013";
        "scouter_synk_option" = "tp"; # Föredra office 365 adresser för primär i scoutnet.
        "ledare_synk_option" = "@"; # Bara office 365 adresser.
        "email_addresses" = "";
    };
    "aventyrarledare@landvetterscout.se" = @{
        "scoutnet_list_id"= "5014";
        "scouter_synk_option" = "tp"; # Föredra office 365 adresser för primär i scoutnet.
        "ledare_synk_option" = "@"; # Bara office 365 adresser.
        "email_addresses" = "";
    };
    "utmanarledare@landvetterscout.se" = @{
        "scoutnet_list_id"= "5015";
        "scouter_synk_option" = "tp"; # Föredra office 365 adresser för primär i scoutnet.
        "ledare_synk_option" = "@"; # Bara office 365 adresser.
        "email_addresses" = "";
    };
    "avdelningsledare@landvetterscout.se" = @{
        "scoutnet_list_id"= "15030";
        "scouter_synk_option" = "tp"; # Föredra office 365 adresser för primär i scoutnet.
        "ledare_synk_option" = "@"; # Bara office 365 adresser.
        "email_addresses" = "";
    };
}

# Gruppnamn för alla ledare. Gruppen måste skapas i office 365 innan den kan användas här.
$conf.AllUsersGroupName='ledare'

# Rubrik för mailet till ny användare.
$conf.NewUserEmailSubject="Ditt konto i Landvetter scoutkårs Office 365 är skapat"

# Texten i mailet till ny användare.
$conf.NewUserEmailText=@"
Hej <DisplayName>!

Som ledare i Landvetter scoutkår så får du ett mailkonto i Landvetter scoutkårs Office 365.
Kontot är bland annat till för att komma åt scoutkårens gemensamma dokumentarkiv .
Du får en e-post adress <UserPrincipalName> som du kan använda för att skicka mail i kårens namn.

Ditt användarnamn är: <UserPrincipalName>
Ditt temporära lösenord är: <Password>
Lösenordet måste bytas första gången du loggar in.

Du kan logga in på Office 365 på https://www.microsoft365.com för att komma åt din nya mailbox.
Outlook hittar du på https://outlook.office.com

Har du frågor så skicka de till it@landvetterscout.se

Mvh
Landvetter Scoutkår
"@

# Rubrik för e-brevet som skickas till användarens nya e-postadress.
$conf.NewUserInfoEmailSubject="Välkommen till Landvetter scoutkårs Office 365"

# Texten för e-brevet som skickas till användarens nya e-postadress.
$conf.NewUserInfoEmailText=@"
Hej <DisplayName>!

Som ledare i Landvetter scoutkår har du nu fått ett konto i Landvetter scoutkårs Office 365.
Kontot är bland annat till för att komma åt scoutkårens gemensamma dokumentarkiv som finns i sharepoint.
Du har en e-post adress <UserPrincipalName> som du kan använda för att skicka mail i kårens namn.

Länkar som är bra att hålla koll på:
Landvetter scoutkårs sharepoint: https://landvetterscout.sharepoint.com
Praktisk info: https://landvetterscout.sharepoint.com/Delade%20dokument/K%C3%A5rdokument/Praktisk%20info
Kårhandboken: https://landvetterscout.sharepoint.com/Delade%20dokument/K%C3%A5rdokument/PM/K%C3%A5rhandboken.pdf
Hemsidan: https://landvetterscout.se
Scoutnet: https://www.scoutnet.se

Har du frågor så skicka de till it@landvetterscout.se

Mvh
Landvetter Scoutkår
"@

# Texten i det automatiska svaret, om man skickar brev till medlem som slutat.
$conf.DisabledAccountsAutoReplyText=@"
<html><body>
<DisplayName> är inte längre medlem i Landvetter scoutkår.<br>
Skicka din fråga till info@landvetterscout.se så kommer vi försöka att svara på den.<br>
Mvh<br>
Landvetter Scoutkår
</body></html>
"@

# Standardsignatur för nya användare. Textvariant.
$conf.SignatureText=@"
Med vänliga hälsningar

<DisplayName>
Landvetter scoutkår
www.landvetterscout.se
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
                    <tr class="ng-scope" style="">
                        <td valign="top" style="padding:0px 0px 6px; font-family:Verdana; line-height:18px; vertical-align:top">
                            <span class="ng-binding ng-scope" style="color:rgb(00,54,96); font-size:10pt">Landvetter scoutkår<br style=""></span>
                            <span class="ng-binding ng-scope" style="color:rgb(00,54,96); font-size:10pt">www.landvetterscout.se</span>
                        </td>
                    </tr>
                    <tr class="ng-scope" style="">
                        <td valign="top" style="padding:0px 0px 6px; font-family:Verdana; line-height:18px; vertical-align:top">
                            <img alt="Logo" src="" style=""><span id="dataURI" style="display:none">data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAH4AAAB&#43;CAYAAADiI6WIAAAgAElEQVR4nNR9d3yURf7/e&#43;Z59tmSzWbTSSUQCKEESYBQBStYUbCgd4KKep5n4WyH5U7B80SOU892du4UORQLggUpovRepIYEQnrvyfbnmfn9sfvsPrvZDQE9/f6G17JPZuY9z8znM58yn5nnWcIYAwAQQgAAnHNoEyEEnPNw5ZQQwjjn6re/rrY9NfW23bPEUc458&#43;VTACwSbu/evbS9vZ0CEAFIvm8AkAG4AcgxMTFs1KhRTIvT3o9z7r/Hme7XG5yWdipNfXWYDwMtjnPO1PZ&#43;Ko6oYC0TtAPRAjRlQcz23ZiF4tT8UKaG3iu03bPFaesvW7ZMPHr0aB8A6QD6AMgCYHn//feNtbW1BniZbvJ9A16m2wG4U1JSnLfeeqsDQAeAMgB1AKqGDh1ad8stt8jh7hdCMz8DwnUyFOcbn8qobrgwE1zL8J&#43;E0zKe&#43;jrDIhC3x3QmZoZOnh7KVOJExHk8HupwOCgAcf369aZFixYNAHARgLySkpKs9vZ2K6XUJAiCSRAEC6VUIoRQQghEUZT1er1bkiQnADidTsntdhsURRE55&#43;CcM8aYW1GUDkVR7Iqi2K1Wa9vAgQPLABwGsHHevHknL730UjsA2Wg0MlEUoSVsqOSFakLtBOlJg0IjUGrqjebtDY74Bkt9BUGqK1TawzG3J6kMVXWh0hFCLDVfVVHdcEePHjUUFxdn7tq1q/&#43;iRYtyAUwGMF4UxbiUlBSWlpYmWa1WJCYmOvv27eseOHAgzcvLM8XHx1Oz2Yzo6GhZp9OF9o0RQuDxeNDZ2Um7urpoS0sLPXLkiLOoqMhdXl4uNjY2mtrb21FXV8eqq6uZx&#43;NpAbAdwKZ58&#43;YVjRkzpjQnJ6diyJAhTpUZKmPD0LPbdyQGhTOHPZnks8ERxlhoB0IZ7ZdAX7mWMWfVqTD11JnZo5/w6quvmk6cOHHJjh07Lty/f38hgCFGo9Gcl5fHxo8fLw0cONCZk5PDsrOzTYmJiSwqKopG6lekfoczderEsNlsaGpqomVlZfLx48dZcXGxtGXLFvuRI0dEl8tlB3CsoKBg97hx474fNGjQhvvvv9&#43;JYK0ViU5BtNakngQlaFKdK46EMknbWC9VPvUNJsgmR5oI4fwIaOwVANhsNrS3t5umTZuWZbfbp1VUVFwjy3L/mJgYS15envO2224zn3feeUhISKBWqxUGgwGEkDP2N5LGOhscYwxOpxOtra2sqamJHTp0iH7wwQf2H3/80dDR0dEmCEJpZmbmKpPJtHr16tVlMTExdrPZ3E3yf/XkU/VgjIFzTjXflDFGOedUNQe&#43;sqAPYwy&#43;eurfVPu3Fh&#43;KDVe2YsWKhFmzZs0AsIwQYuvTp49nypQpfOHChZ0nTpxQZFnm2sQY44wx/7U2L/Tzv8LJssxPnDihLFq0yDFlyhQlKSnJRQixAVg2a9asGStWrEjQ0iWEZjSkLIh2obgeys4KpzLKT3xfp/z5IUwNnSjdyjXtBTFXUz&#43;0I2CM0S1btphnzZo1zWg0rgTQmJSU5Pjzn/&#43;sfPvtt67a2tqwTAjHLG1&#43;KKN&#43;CVxtba2yZs0ax5NPPulJSEiwAWg0Go0rZ82aNW3Lli1m7fjD0TYcbcIJ10/FBTFDOwF835GYHjTTQphNeYh2CJmV/uuOjg567Ngx69ixY6&#43;NjY3dotfrWwcNGuRZsGCBraSkRHE6nT0yJhxT/i&#43;UMca40&#43;nkJSUlnmeffdaTm5urGAyG9tjY2C1jx4699tixY9aOjo5QDRiW7qGfEMk&#43;Z1y4Bnr8hEpuOC0Qeq3NUzuwYsUK8a677roAwFJBEGxjx451LViwwFZRUaGEI3A4wodjwtmW/xK4iooK5dlnn3WNGTPGIQiCDcDSu&#43;6664IVK1aIPJhhoao7iCchqvsn4SLODLWBMzCTsmDVFe7G3WzQ0qVL46xW62MATlksFsdTTz3lKSoqcnk8nrDS01sC9xbX00T6X&#43;E8Hg8vKipyPfXUUx6z2WwDcMpqtT62dOnSOB7QklqNiJA8yoJNBP0puCAVH2ovItnjkPxQZy6s3WGM4fjx4&#43;bo6OgbJUnaZ7FYHDNnznQdPHjQ43a7IxLxTNeRyntiwq&#43;Jc7vd/MCBA56ZM2e6YmJiXJIk7YuOjr7x&#43;PHj5hBzGUmLdnPYzgUXjrER1XukT6RyNb&#43;6upo&#43;99xzmdnZ2X&#43;jlNpGjRrleuONN2wOh6NHYp1t2bnU&#43;7VwDoeDv/XWW57Ro0d7BEGwZWdn/&#43;25557LrK6uDueJ98aMnhVO64n71UJI5bCzqLcfxhgmTZo0DMDXABxz5szpPHbsmEer1s8kyb0p&#43;6n4X6PM4/Hw48ePe&#43;bMmdMJwAHg60mTJg3jZ8n4cynr5u1pJkJEE6C9Dncjxhi6urro&#43;vXrrRMmTLiFENLYr18/z7vvvqswxhR14KEMD/1EImI4XKR62vJwZb8WLqT/yttvv&#43;3KyspyEUIaJ0yYcMv69eutXV1dQSumUPqHfkL40iOuR9uhqRxRjUSaAA8//HAcgEUAmi&#43;&#43;&#43;GLbd99953K73Wd0kCIxtSdcpHQumF8Spx2b2&#43;3mGzZscF188cU2AM0AFj388MNxPLKd7uZY8xA73xMunCoPCuREutZMBv&#43;sZIzB7XbTJ554wqLT6d7T6XSdt912m&#43;P06dOuSNISjrHhpL0nKT8TUUPzfm4cY4w7XW5e19SiNLd18ua2DqW1o8vT0WXnNQ3NisPp6lUbjDF&#43;&#43;vRpz&#43;233&#43;7R6XQ2nU733hNPPGFxu909rc/DCuKZPmElmgUvBYLCjGGk3j8B7HY7feCBB3IBrLRYLK4FCxZ42tvbIxK3t3/3lH82dX9OXGtHF2eMcUVR&#43;Hc7f&#43;T3/u1N18sfrvJ02uz8j4vecVzxh6cdLe2dyqfrtym/f&#43;Z1Ze22/UqX3cEVxniX3cHdYZauavvt7e38mWee8VgsFheAlQ888ECu3W7XhtMjaeUevXptXjdHLtKMiXCzoAkxb968HABrjUajY9GiRa7Ozs4zErMnSe4tLhR7NlJ7tjiny82//GG38v3uQ1xhjK/6fqcy6Kq7lTueetnR2NqufL15j5J20WyuL7iWv7b8K5fD5eIL3ljuybjkVv7mijXc4XTxfcdKlP9&#43;84PH4XQpkbRZZ2cnX7RokUev19sArJ03b14O1/AohFfdnHIeRjC1uJ5sQUTHIfRjt9vpvHnzciVJ2pKYmOh49dVXHTabLSKjzobQvcWdKfVkSnqDY4zxitpG5Q9//Zcy6bbHmivrGl0Hi0p54c0P8pE3zlVOV9UpTpeb51//gIcOv5rT4VfzvOn38l2HTvD65jblynvnuxLO/42y89AJpb3Txmc/8aLt&#43;ocWun48cZp7Qjae1GSz2fjLL79sS0hIsEmStGXevHm5drs93KZXpE2xiBtjkdRBr3fn3G43feCBB4YAWJ&#43;YmGhbtmxZZ&#43;gOWm&#43;I29raymtqariidIvYnhEbbkKE&#43;5wrTlYUvvvwCU/hzQ8pdPjV/KnXPlS67E6&#43;8J1PFOG8afytFWs8jDH&#43;xcYdisp0mnc1F86bxp9&#43;fZnicLr4Z&#43;u3ueIn3qzc/Ke/exhjfMnKdQ79yOk856q7&#43;cffblEiTUpZlvmyZcs6k5KSHADWP/DAA0N8Nv8n7c75GRtyCFDdP2ac&#43;/fMu53rAoD58&#43;dbXnnllYUGg2Hs008/jZkzZ5opDdpeD6qvdkCbPB4Pli9fzu684w5WVVkZ9j7hcOHqAYE9f&#43;2&#43;d6Q98DPhAGDPkRJ57vNvY//xk1TSifLAvqnMI8s4frqSxsdE26ddOIZ6ZAUfr90KqM0Rb9trt&#43;1DS0cXxo8YLI4aOhBHSsrhdHvYhBFDJJ0g4GRFDf704hK6dvv&#43;sOOjlGLmzJnm&#43;fPnw2QyjX/llVcWzp8/3wLvgQuGAF&#43;0dNPyT3soRT2owajmAIMf6Mtj6rfaIS1B7HY7feSRRxIWL178ksVimfLss89Kc&#43;bMMQmC0O1IkZZ54Q5iyLKMI0eOYMeaNXTRwoWyy&#43;VivcFp&#43;trtIEY4Jp4Lrrqhhf3&#43;mdex&#43;0iJyBiHThSYUS8xh9OFk5W1bHz&#43;YMlqiaLHSytx6ERpYMZz72ffsVP0ZEUtS4630gGZKW6700VrG1sQZ432n5SprG3CI/9Ygk17j3RjPiEEgiDgtttuMzz77LOS1Wq9bPHixS898sgjCTabjfoENnSy0gj88/M49Ihw0JElH0A9hBlkM3bs2GF94YUXngRw42OPPSbec889ktFo9ONCiaklcOi3KIpITk6mV4Ng/4oV4oqPP3a73e4z4npK4aTnbHEdNjve&#43;Phrt8kgyZNGDmW5/dLtHlmhp6vrqSgISEuMo8MG9BVFKqCitlFubbcFpM53K4UxlJRXU0II0pLiu8wmo5wUF4O9R0&#43;KAFjf1CT7hPzBclxMtHvpV987axpbWLi&#43;G41G3H333eK8efMogBtfeOGFJ3fs2GGFRivzkONxIeOm2vGJCE5aQNABTM015Zxj/vz51wO47c477zTcf//91GQyBd0olPk9JUIIKKXIJgSXtXXQfzz5Z7GxqYndf//9VD0c2VMKpxF6q&#43;Z7wikKwzUXjhVnX30RkyQdWju66PHSSiYKVNRLOuRkpbGM5ARKKUFDSxu67A6AcPj0PEAIAI7y2gYAQJ&#43;EWPfg/uksymigjDG8/NjvWG6/dJoYawFAaEt7JyihMiFE0vZFvTaZTLjvvvvE8vJy9uabb942f/7845dccsnbCJy1Yxoc1VyrJsGvkYJmgYYA6iFArTZgAFBTU4MLLrigYMeOHYunTZtmePLJJxEVFQWNOvG3s3rVKjw2bx5OnToFRVEiMo4xBqfTKRsBTAPBS9U14trHHqcL//Y32O32iDg1hTI2kqk5W1ysxYxhA/uKHBAbWtrgdHukwrwc8dJxBdRsMuLqyYU0Kd4KxjlaO7pEu9NFvUyHj&#43;kAQFBe0wgA0Eti0sO3TpcIIZg0ciguGJ0nCpRKja0dVFYUMScrTeqTECv1NHGjoqLw5JNP0quvvlrasWPH4smTJ4&#43;qra1Vx0k1dYN8AI19BwAmhrOjISc4/Q85cM7Z&#43;&#43;&#43;/33/z5s2LcnJypD/96U9ITU2loQRU2&#43;zq6pLfe&#43;MNumfzZvb7Bx&#43;kV151FTUajd00gizLrLa6uisLxCoRgvEAXpQZ7n3tdfTPzmY33Xwz1foOvZHgSIQLhw&#43;H45yjrKYB//5iA1v&#43;zSZWXd9EGed0cP8MuTBvEP4w83KaPySbJdY1gRBCJZ3oFgQqMo9MQXwS79P30VFGAMCYvEFienICyqrr2furN9K12/ezIyXlcLhcNCs1iV04ejj7/Y1XiAVDsrv1S70mhCAtLU187LHH5OLiYsPmzZsXvf/&#43;&#43;3c//vjjJ7V8CoPzP6gCAML8&#43;fMjEseXRwghnBDCT5w4YXn00UcfaW1tnfHaa6/RqVOn6imlEW16R2envHnDBnbL4aO6dzZvJjZJwrhx4/zlKq6zs5O/&#43;&#43;KLtsm1daZBhAAgSCCAyeHAW&#43;VlyoQLL0RCQgIJxYUjTE8T4kxM16aymgY8tPhdtvTLjdRmd7LLJ47id994Ob1yciFNSbCSmsZWpCbGkfTkeEIJQUVtI/9u54/E4XKrM8jLdwL85opJGDs8F9ZoM1xuN/th72FqkCRMHZ&#43;vTCwYKvRNSWI/7Dks7zt&#43;SrfzUBEmjRyKxLiYsBpJ/U5NTRWTk5PJp59&#43;ml5WVsanTJmyMyEhwQWAR8IhsOZAUIBGXfOFCdxQzjmio6P/IIqiY&#43;7cue2MMSV0LRy6Zj569KgyMj9faSIiP0AEPgbgn336qSd0rbpx40bHJEKVIiJyDxW5h&#43;q4h&#43;p4GxH5o6D8jltu4YqiBN2rsrKSHz58uMddL22/Qst6wnlkmc/5yz8V4bxpyqTb5rlOV9dzt8ejVNY18j1HSpQdPxZ5jpSUKe2dNn8bu48U80FX383963jNZ/32Ax71fnanixedruJ7jhQruw&#43;fUEqr6rjD6eLV9c2e3z/zmidq9HX8kjufdFXVN3XbxQzX30cffdSj0&#43;lc0dHRf1B5qOWjln/aj7ZCj7tzS5cuzZIkqeSiiy5ynD592hOJyNq87zdudE3IynK1Ei8zPycCn3z&#43;&#43;Up1dbW/XmtrqzL9sss6FxLKHX6mB5i/j4h8hE5STp486SfEkSNHPDOmT/fcMWfOWW/ChKsbitt24Jgi5V&#43;rTJj1qOewj8Fvf/qtcvk9T/OMS29zJU36rWPY9Hs9v1vwqvL97h8VRVF4TUMzH/ObhzjNC2a6dfxMfry00tvvk&#43;X8kX&#43;8p4y&#43;6UGefsltntSLZnsm3/YYf/LlDzytHV1KXVOr8ps/LXZEj7leeW35Vx63Rz7jmMrKyjwXX3yxS5KkkqVLl2bx7oddwx6FC1rH8/DLOXz66afS/fffPzcuLi7rgQceEDMzM/2&#43;gapK1L&#43;1uKrKSvRpbKTq0mEqISAnTmDXrl2Mcw6n04nXXnnF7fhhk2kOKASEqlyOIQQY5vFgzZo1AICGhgZ2zx134OvVq8XxEyb0aMu136HXPeG&#43;2LjTqdfr2G&#43;vusA9MDOVvrb8K/bYS/&#43;ha7ftR3V9s9TU1mE4dqpCXPL5enr7X17GjyfKkBRnxeihA1noECbmD0GCNRrFZdWY9&#43;K/8fKy1XTfsZOoaWgW6xpbxC37j&#43;KF91eKj/zjPWbUS3j8zhtEp8sj7z1aItqdrrCmSZsyMjLE&#43;&#43;67D7GxsZn333//3E8&#43;&#43;URC8NM0TOWt5m&#43;vy&#43;9r0P&#43;ojZrns7/i&#43;vXrr2hra7vppptuYldffXU3Zy4cIRVFQXVtrZhld4iCL18CkNfUTIuOHYOiKPho&#43;XK25p//FJ93uWkcAYga9fBTj0AEwRWE0jVffcU6Ozux4KmnELN7tziCUDamsFBGhBTOs&#43;/l8pIdPVUh9ktNFi8YnWfotDnk59/9hLZ1dqkNBSoyjoraRnrPX1&#43;Xm9o62O3TLw0K4Eg6kV02caRbknR4/r1P2Zqt&#43;6AoTGNpvSsAt8eDLzftFtfvPIh&#43;6clsUFZa09GTFU6b3dHtqdtQISOEYNq0aeINN9zgbmtru2n9&#43;vVXdHZ2hi7T1UCOfxVHNQ6DGgjQrtsxb968uHfeeeeujIwMy9y5c0VKqVZL&#43;DugdkLN7&#43;rqYkcPH3YPA4FAfJQAQRLjcLhcOH78OPv3woX0D20dYo7PoQswXDurOcaBoK6xUdy4cSPbu2ED5oDCFBPjNBiNEbkX2p/exhQUxuB0uUWTUY/4GAuta26DzekK9E3bju&#43;yuLxaXLN1H8sb2JelJMZ1qWWZfRIxYlA/sbymAat&#43;2NUN5x8yIeiyOVh9cxsTBCoN7Jtmrm9uoy6P3C32rY5LOz5KKX344YcNGRkZlnfeeeeuefPmxUXA&#43;df5Qbofwc&#43;w0dbWVrpy5cob9Xr9BfPmzRMzMjKotiFtCp0IjQ0N8unvNroHgvjG6Q1mdAEQBIG&#43;uHAhJp08hWt8Uo1u0g7/3ykEcDqd&#43;ObLL&#43;X&#43;tXVIBgE3GeVwewLh&#43;hPa39C&#43;axOlFHqdTpZlBU6XG8nxMZB0IgMAUaDISk1yThgxuC1vYFaHJcoEQoD2LjvWbd9PO7rsdOZl5/s7lZOVxobn9MPyNZvR3tFFwTnMJgPLyUrtmlgwpGNw/3SnQfIGqKJMBiRaLeAcrL2zS46LMTOdKHR78UKk1UlGRoY4b948Ua/XX7By5cobW1tbKbwanWpwaoiXauO8WkmnANjcuXOT6urq7r/kkkvY9OnTIQiC/0ZaXyAc0Tf98IPYt67OkuvNgcrEVkFA6cmT7Oiu3fRuQqEnxKfiSVA9gIP7vj0g6OzsZAcOHaKXOF00CoDS3CK5nE6m9Sl6YmyoKYqEo4TQwdkZcm1jCw4VlzGzyShOGDEEAHDjZec7ty1dLG36z/PWPctfNC99/mFnfm52FwDsO3oSVfXNmDIuX6KUyAKlyB&#43;c7RQEStdu3c9ACPpn9MHrT97jPPDJK6Yfliy0bPtgsfTg7GucZpOBpSclIC8nS&#43;yyOdjeoyelodmZktlkpFptqvYxnPailGLGjBni1KlTaV1d3f1z585NUqXch9MG5Rgl3ldlhD6oz7Zu3Uq3bdt2o8ViSb/uuutoSkqKpL2RlnDhCPnV8uXsfEJgIcEyXGU2yQ0NDc7ktjbWB8Fl3Zjl&#43;64DR319vdxy7JgzHwQpAJLsdsOBgwe7SUS4754mRrj6V04abWhoaWdfbdrNnC43/nDTFXRw/wz5odnTpViLmX6zZS/7z&#43;rvkDegr&#43;G1J35P42LMzvLaBlpV38QsZpMcY45yCwLF4H7phpLyajS2ttOkuBj25F0z3VMnFBg&#43;W7&#43;dLlm5gRn0Eu667jIxp28a7phxKfqn92FLv/re7XR7zHkDs5jJoO/W7540WZ8&#43;feh1111HLRZL5rZt227ctm2bFhfsL4Sc4qDMe24OCxcuHABgy7hx4xydnZ3&#43;TfJI60ptfnV1tTK6Tx9&#43;jASWZeonPzPTccWUKc2/k/SK27d020FEfhWo//MMEXiTD&#43;v2LQMB8FtAeAMRuZuK/FVC&#43;dTx45vdbvfZbeD3Irk9Mr/l8X&#43;4osfcoPxz6Sqltb1TOXD8FO&#43;0Ofi6bfuVvlNuV2LG3sjnPv82b&#43;vo4s&#43;9u8IhDL&#43;av/jBF57jpRU8/4YHlKjR1/E9R4r5Z&#43;u38cRJv&#43;GzHn/B1dFlV15a&#43;gWPn3gzNxdez9/7fB13uT1816Eipb3TpqxYu9XWd8ocz5Tf/YWXVdef04HPzs5OZdy4cQ4AWxYuXDggwnk9GhTbBbzS29DQID7&#43;&#43;ONTKKUFDz30kGg2mwNPWUaQFBXLOUdpaSnVuVzoS7S221vPTSmVZRmUKUy1/Z&#43;A4SvNZylXUAKvsm/hHP/mDFYAlxAKKwEICK4ERefx43E7d&#43;5EuBSur71JnHPoRAGP33EDPS8nS37ilfex4I3l7paOLqfCFBiNeuhEkY49bxC76/qpMgfYleePoknxVvvpqjrREmVCamIcpQJFVlqyXFpd76SEsmsvGgfGGb3psknsivNHyQBYYqwFCmNMpxPlRUs&#43;lf&#43;46G2DKFL6xJ03sL6pST06pJF4YTab6R//&#43;EdKCBnx&#43;OOPT2loaFB9Du2&#43;DIu0O2cB8NtRo0ZJkyZN6rbfq94o0lq4qakJ0QpjEjhVFbbXcnPk1DdIm9vbzUkKoyACmsHxPQ/WQmUAtoMhHxRF4DgCjjmgmAridxXTCcH5HV34aPlyVlhYSPX6gFrkYULIkYgXKY6fk5Um/usv98qvLf8S//r4G&#43;mL73eyQVnpSE&#43;Op1dfUIg50y&#43;lkk7EO5&#43;txY1TJkqFeTnU7ZERY47CtReNw6CsNMRZzAzgiI4yYNx5uXTJyg3yhYXDxWfvn4V&#43;acn4dus&#43;vPv5OpyurqfFZdV08shheOT2GZg4cqh2Z63XS1FV8C644AJxxIgR7gMHDvwWwEcAWhCyO&#43;cP72m/CwoKpkmS5Hn77bc96jGqswmLrlq1il9qsShqBM5NA6HYEiLyWSD8OhDFQ3X8YyLw2GC1wAHwMSC8lYi8iYh8MxG5zafitVG9IiLyScOH8/3790fsW6S&#43;R0pq3R9PlPLymgZud7p4ZV2jZ&#43;22/fzz9dt4cVm1YnM4&#43;f7jpxyX3/O0a&#43;SNcz2dNgdv77Qp9c2tvKahmZeUV/Oi01W8pqGZ1zW18sbWdo/d6eK3//mfypBr73Fs3nvE02mzKxW1jcrq73fyVRt38tKqOo/d4eSNLe18z5Fixe5w9ti/nspkWeZvvfWWS5IkV0FBwTQeHLalnHN0251bt24draysvDk3NxejR49m2iUTDyNJoXkAEBMTAxsVqOKbYtrlXBYB/kkEOAHqAMdmztERZvYeAMchAOMJwTifmeB&#43;b9&#43;rPeIAxNTUsNqaGuTn5wdJyU/dnWtu62TPv/spnTqhoOuC0XnmMcMHgRKwoycr8OWm3fLSLzeyQ8WnDQvn3srcHg8&#43;Xb&#43;N7jlSgtrGFjS1dUBWGJJiY5AQa0Fuv3R6&#43;7WX4PKJI9lnG7YZ5jz1svuGKRPZ1AkF0oT8IRBFkXV02cSNuw/JqzbupCOHDKAjhwyIuDsXKalllFIUFhYiNzeXVlZW3rxu3bqvpkyZErQ7J4Yyb/Xq1VlNTU2FU6dOlXNzcw2hTA53s1BiWq1WKAKFDYBFszxTY3MWcFhAUMw59oEj3E69G8AXYBjvmzpc87/algAOU2cXbF1dP/vu3Hm5/aler5MfeWGJ1DclEUaDHoSA1jW1obKuCW6PxwQQNLV1Up0oICM5gZXXNND2LjtIeycDwEwGvRgXY0ZGnwREmQxoaG2HrDCUVtVJi//9GT76djNSEmIhUEqdbg9qGpppUnwsW/jH28JGR3ubCCHIzc2VRowY4T58&#43;HDh6tWrs6ZMmVIKzfmKoJh7Q0MDraqqujY6OrrPzJkzqV6v7xYl0hJMeyPt36mpqQwmEy1pacdILQYB6Wfg&#43;BEchxHZ8fqaM9wBihwSwGlcGnAAsk5Hfe&#43;ZC8Jqma3tZ7i&#43;h8PFRkfhr/fdIlbXN&#43;O7XT&#43;GDoCCA1Sg4JwzURBodkYK7Z&#43;RgoGZKeiyO6nbI9N4azSqG5pZfXMbNRsNEAhh6t0Y5yivaUB5TYP/3v3SktmbT/0B8dbobnQ&#43;08QNLdPr9bjpppvEL774ok9VVdW1DQ0N/0xKSgo6c&#43;e/3rZtW8KqVavOT0pKopMmTZIirRm1EyH0poQQJCQk0JGXXspWgsGpOXZK/CobkDnwLRhsiJzqAGzXrAgCYR7vvWwAmlNTWEpqalD/1H6o3z3ta/eES09OwHNzZ2Pq&#43;HyIghC0AUMoMHJItjzr6otYU1uH/MDzbzsXvLGcddodeOvTb/Hw4nehKMz9weqN7lse&#43;4dcWlXPfnPVBdKU8fkIJ8yFwwa5X3j0TnpeTr9wcfag796WTZgwgSYnJ4urVq06f9u2bQm&#43;bAogsDvHOYfD4RgAoOCaa66h0dHRZwfbggkAACAASURBVAx0aG8emn/nvffSrSl92DbOwdQyDeW6AGw8wzKrE8B3nKE5qFpgEmwGR9yQIRiUm9tt8OH6GUnSe8KNHDIAb8&#43;/ny2cOxvDc7JkS5QJfeJj2Z9uu07&#43;4G8P0dx&#43;6eI3W/Zi874jUqwlirZ12PDt1n1YuXEnGlrbRYNeR2sbW&#43;nLy1bLhBD8c97v8PeH5rhz&#43;6XLMWYTBmSk4IHfXI3lf3&#43;EXjVpNNVLujNqpEh/h&#43;IsFguuueYaAChwOBwDOA&#43;8U9fvybe3t4sA7tPpdMqePXuCDvi3trbyTZt&#43;4OqbKyIdYpDlwP6x2&#43;3m7//nP8qY5GTlWyLwDr9H7vXKl/iCMmf6pAF8U8gBDTfV8eNE5OOsVmX1qlXdAjgHDx7kbW1tP/npmVAcY4wrmvzaxhbl/VXfKebC6z2WsTfyL3/Y5Tlw/JQy4Iq7XKZRM/gPuw8ph0vKeL/L7lCix9ygPPKP91wnK2rO6sGPSP0Lhw&#43;Xdu/erYiiqAC4r729XeS&#43;1ZvWiRABjBswYIDc0NCAtWvX&#43;g9Ibt&#43;&#43;HXPuuAMnT570z6YjR47gxRdfQGNjo38p&#43;Nbbb2H79u3exkQR111/Pf39woVYkJwoz&#43;cMR30zUgbHOt5txzFsqgHwnd8n4eDgKOYc88GQP2MGpkydGrRT09TUhLvvvputW7eOadV8b50lHuLPaHFuj4wdB4/j68173P94f2Xb3c&#43;8Jj&#43;w8C1qd7rEC0YPkwuGDBBb2rtYXWMrZIXh4InTXUP6Z&#43;DqyYWwOZz0nx&#43;uku6a/ypb8MZ/Wz5Zt7Vj876jaO3oCjKbkZzo0P7xEHMbCZecnIwBAwbIAMbB59MRQgLRnKamJgnA&#43;OzsbOmp&#43;X&#43;mixY/j5MnT8LlcmHHju2oqivHJ59&#43;As45PB4PFvz1aTzz/Hz8643XAQC7du3CP15cjPeWvAe73Q5CCNxuNzZt3kynzZ0ruu6Y4/5tlEl&#43;hTMs4xzbenDqggYKYAXnKAJwmANvALgrzor0e//AnnjmGVpdXY0TJ054H/YHsHnzZvnQkUN0586dcLvdfmJxzmG32wPvgNEQsicia5MoCqCUsnc&#43;XSv&#43;fcln5q827ZG6HE7k52bjqd/fLCbHWbHq&#43;52wu1ySrChYv&#43;OAqbWjiz16&#43;ww6dUIBEyjFpr1H6EtLV1n&#43;9s7Hks3h9B/E1PbzTBtgofk94axWKx0/frwEYLyPxzQINHr06BEAFGucVek32cqzJsby119/nZ86dYoXTi7guTdYeeH4kbyxsZF//vnnPHGogeffE8ezhqbx2tpaft999/KsCy18&#43;NjBfPv27VxRFP7ZZ5/yvoNT&#43;KgxBXzdunX8wgsvVNLS0vi5fOLj4nlKSooy88Yblf/&#43;97/cZrPxrq4ufvvttysjRozg1dXVvKuri19/w/WOPiONrsuvuoxXVlb6VeLhw4f5Qw89pJSUlPQY2&#43;9N4Keto0t5&#43;5NvPVfdu0C5&#43;U9/9xw7VcEZY3zP0RKPcdQMh/fZuau4dfxM16v//dLmcrt5WU298sg/3vNMvfsv/Pn3PvGU1zQoZ2N&#43;zjUpisJfffVVjyiKfPTo0SO47yiWqJkhkwCAJtiRdVUc2k978OW3K0EpRbWjGIMnRKP8gwasXr0ayz9ajqxLopGQJ6Fxhw3Lly/H4eOHkDbRgPo9tdi6dStycnLwzbffwJzvRGtNGe655x5YrVY6/drpfq/8bFJ1TTVq62rdV1x5pfjRRx/RPn36yEOHDaW7D29lpRWl9PPVn7LBg4biUPFeQ7/LzTi18ahcW1dD09LSaFtbG56e/zQ2bl6PgoJ8DBgwoFvYmYeoeM67B0/Ua4vZRO&#43;YMYVee/FYSKJILWYTWju68O&#43;V6yEKlLoBcEIgywpd9f0u8arJhaxvSiJdcO9vaWtHF1IS40Tay3hCaD/DlfeEo5QiOzsbvqXcJAAHAa/Op42NjXA6nRcb4wU6cJoFUcnedfGBTTtR&#43;lopEidIMMQJEPt3YeHzz0E2d2Fovh5UACyDCF559WUgpQODUoxQ8jhWrlmBzMxMrN&#43;zCv1n6tHIXbCXKZgwbgKSkpJCeohue7PhBqPX63H02FHpoUfnymKKCwsWtsi/m3OP2OCoErOvisa/3nqdTh5/IVhGC6Izzag2tIgnik6wIYOHsL&#43;/8Ly8s3S9FDOc04aGRsYYA6UUFRUV7ODBg7jiiiuoKIqw2&#43;344IMPWEtLC33iiSeC&#43;hI6MQCOBKvF/3eM2YRn758l/vl3MxljXObgVKBUFAQBsRYzCCEwGfQw6oN2t/0pdFkcaWKeC27o0KFiVlYWOjs7L2xoaHgtKSmJUUIIe&#43;mll8yHDx/OtGTqkDDUAEIAY7wAfSJHg70aSed585JH6tHMqmApcIMK3hV14nkS2kgtEgsF6KIoYvrpUKUcx&#43;NPzUP0EAbZwdC0nWLy5AuQmJgYZsThiRCaoqOjkZaaRp2KXcy73Yoq8UfDHx66U07Ik5AxKQoNjnIs&#43;2IJ6zPSACmKwpwu4qv1q/HWu2&#43;6P/zyPZo93QRLpg7HThxlNlsX2trasHDhQjz44IP0VOkpBgB79&#43;5lC//xLFv84t9lp9PZLT7BQ2xrS0sL6urqwDkHpRSxFjNSEuNoalKcmJ6cQFMS45AUFwNRCPifPQXCQmnQ00bT2eBSU1MRHx&#43;Pw4cPZ7300ktmAP4DsKkAzLooCir6QlMikDExCu4RDEQA6vY7EJWsw/A7YmGI9Q2EA&#43;ZUHUb8Lg6mJBGEAFK0gAFXRqPxaDviB&#43;lRsqoT&#43;YPPR7&#43;sfr32rMMlQRDQN7MvDhXvpzozRfZV0UjK9xgsGToIeoKhv7XCY2fUkikBBIjJkvD9x&#43;vott2bpJTLCbVk6CA7GHb88IOzqaXJsHvnbrp25xfUqXdi46YNcv9&#43;/cW//f0Zt9C/XeLFLnr06BGMHDkKiqJg/fr1TJIkTJo0iYqi6HcUX3/9dRQVFclLliwRDQZ/dDtiwCiS&#43;QiXzqYuACiKgu3btyM5ORk5OTlBuNOnT7vr6uooADO8vC5WGd&#43;fUJhNiSLAOTgHXO0MnAHNR92o/IJDT0xw0Hb0v04Pc4oINZBCBSA6TX2w0bv8jkoRYUyIQtl6G2Ls6SicMibo2Na5pszMTFij4tFwsA2pY0zQWwRv4JYDliy1DwSEAPE5eoizCAQjpeY&#43;IgglMCaIONVVY/ruu&#43;/k515egPgLZepqJ1i7Zp20b//erqONu8zDbrOg5MtOLPtkqbugYKR0qvSUPPvW2Rg/fjwdPnw4EhISQAjBhu82yK8sWUztHU7xiZNPsKFDh9JQraBlXHt7O3Q6HUwmU0SvPRKjw2mI0PL29nbcettsuWBkAX3j9TdpYmIiCCGwO&#43;x479/v0pLGQyIozAAyARSrByyzCIXFkqGDIgPV2&#43;049p8uNHwehb7ySFx98QzM/u1tKMydhIrVChoOOQF4A&#43;ghXfcRnqCjyoP2AzqMH30&#43;gp545cEDCJfClnFAp9Mhd&#43;Bg1B9wQvF4JxmBuo6F/5tzDkIJrNl6RKfqQKiXIQarAHOaQB95cq5syO0SE4fpkTTcgO/3fctWfLXMNODaKEhmAYl5BuzffVCsq69jz/99IXhiBz10aq&#43;7rb1NBoCW1hYs/udCmnqhjiYMlbBm3df&#43;352pqqrCu&#43;&#43;&#43;y5qamvxdt9lsePbZZ&#43;WvvvrKv&#43;wkhKCyshIff/wx2tragsautqWlA&#43;ccbrcbK1euxKlTp7rR6YdNG&#43;VmpVrcdHgN1q9f5w&#43;S/Pjjj/j4mw9oSqEBVIAJQDrnnFHufQo2hQjEYMmSUL/PiZq1wKSBV&#43;Kay6fjgkkXoG/fvtBLeowsGIlRAyei7HMPumo8XiZzlZuBD5M5yr9xIid1GNLT04PVlO&#43;yJ9UVtsyXNWjQINhrFdjqvPfnvn/&#43;a8795zYDZd6/qQ7of1k08n4Xbcq6xAzBQKC3CsiZHk2H3WKl0Wk6gABxAyWUVpW473/49y3rtn1Jh/4mljr0rYai40WUMYY33/6Xu7zzOE0bZ0L8YD3Wb1wHRVEYIQRff/01HnroIbphwwb/c&#43;sHDx7Em2&#43;9ST/5bAVsdpufaW&#43;9/Rbm/nEuDh06JGuZWFlZicOHD/sniUqT4uJiPPfcc/jzX/4s22w27Y8csWUrPmRp46KQdr6Rvvb2K3D6DqK&#43;8PLfmZhpo31GGUEoMQFIJYRQYeLEiaa77757uiCR0Yl5BlR8JePKSdMxaNAgGI1GaPfjBUFAamoq9IIJu78&#43;DmMyhSHW&#43;xSrV9QIFBko/bYT&#43;ppUXD7lCkg6qecTlZFSGG8fAHSiDjXlDbCLrYhR1TtB8BJR&#43;zf3aQJfkWiiMMaLIKJvMggE5lQRpkQRoF4cEYG2uk5x38ajpoE3mkhcjgRXu4IfN59w6QTJ/eIbi6R&#43;N4nEECfAGCei9Lt2OrJgNDGbo9lNd8zwGPq5hQO7DitzZt1JOzo72G2/u8WFfs26k6dO4oYrfoOEhAScPFki3/fY3VRIsaPudDNumHEjUYNeDz78IHvxxRf5lEunkIQE796K0&#43;XE66&#43;/Ju9uWkNPl5dSxUaVMYVjKKUUX36zyrPkw7fFrGt1JLa/Hqe21ZN4Qwo/XVZK3vvkdZIzMwrGOAFl67ropu&#43;2FE&#43;cOHEjhfcBFzMAlH/twnlZhcjIyAjhQbDdOi/vPOT3H4eKlQrKv7NBcXtnJuccLUUutB2gmDR&#43;MiRJCmyf&#43;5igtqXOcI4IKj8CThAEZGcNQFuJDI&#43;De1W8f9PWp/ZD/lYbDCo7Ay7r0mgMnxOH&#43;EHerenk8ww4dPyg9LfFzyAq30ajkkQQEEjRApT4Nnz77Rr20j9fBE2yS4NmWFDfWSF&#43;veYrecmS91BjP2XIvtICKZ5jw8Z1zOl04pVXX6aWQRzZV0Rj485vUVtXywCgvLwca7eups2knL77n7dlleanT59mX6xfQVPHG5F5uYSPV38o7t27h1XXVLO33nmTmvLs1BgnQpAI4kcTvPXOW&#43;4XXnzBHV8I6C0UhBLooil8vJaE2bNnxy1dunQmhTBw9JDxGFM4BpJeCpKg0IALIQRpqWlIjE5B6Z46VB5sgSJzKC6O0k/dKMydiEE5uQFt4VO7KnHVa7XtiB5sGJx6/&#43;JDpYgZyiCZ1XsQdYZ4lYXX2PvLOLjPLPnqaq/D4HRGAnOKDkQHEA7ooiiMyZQIqQ5dyigTBD3146ge2PTRQb5t2zYMnhVFzKkSPC4Zq9/9ge09uoOkXcmJJV0HV7uC6sNtSkqfFLzy9ku07wwB5hQdWk45iauV2y44/0Lxmeeedp32HBT7XxaNPauL6OQJFyE5KRmLXnjOc7htq5g&#43;MQqmBBFN9S048MNxj63TLn697yPa73ITRKPX9OrMFD9&#43;U0bbWQMdMC2aiAYCToCqbXY4W5Ti2bNnrxVmz56duHTp0ptSUlMyb7juRuh0uiBGc58DFXotCAJiY2MxNDcPkjsajT&#43;6Ub/Tg8Lh4zGyYFQ3Lz7sRNKo80hebCiOEAKD3oDa6nq0uhsQmyN5a2gmi9qGv3niJQj32X4CLyHOFheVpIM5TYQgkiCcFC3A6bYTy0BK&#43;hR4X/wQlSyg6nATTRgDknSeERAAyURx9Jtaevx4EXFlViB5uAFUBAyxIg58US6BcmXZR0ulvtMpLFk6dDR34fiWCnd8Qryy8OVnpH7TJBjjRVABiErW4dDmEmHT1u9Z9jVmEp2u8/adAIJEYc3WkeQCA9GasNrddtgblLLZs2d/LcL3ZEWUydxd4gAQP4W6Sz7g9bQL8gswbOgwdHR0wGq19n7pdgamR0o6nQ5ZfbOw7XAx&#43;FSACIBGhP08UVcdfiZCNRm&#43;XMJ/FpxoJOh/ebQX5yOXaKLI/308qERARe/6w5QkIq7QicrWg&#43;g/JhqEeluJydKhJq0GT8z/ExLPkxCVEgMCgrQJJuz7YJv4h3t/lA0DXYhOD5zM0cdS5F4XA7eN0dhs1Y/yrXIEwJqt1&#43;hIbxINFAAMAEQR3u1Y0aA32AGYwsZ9VUnXSKhXdRJw4i3T6XRQHRE/LjTwoMWryxZf29r1aW9wA7IHYNvezWgrdSNuoD5MkINrJmqggYBp8f71c&#43;F8Lx0LtEEIdFFa8wWAEmRNMYMrHIKe&#43;CcVKJBzTTT6XhglSmbq1ybGOAEpU2XadLRJyro0GoRq/BV4g1TBfAr0gGjqqUk0&#43;hkvqYw3SJLkhvfHdoOT1kMOcpxJUFmkOHJoW6Fl4fC9wUVFRSE7YyBq9x9CbH8JRAgwRiVAYDIhuIyEm2i/DE7QEXBR5a1PIjkgGAhMRsE/XA5vLCJhiB4JQ/SaNkkQTksfv/8UJEheV4cQAtFAAJ/EB56yIOjdyYj/Qyl30GC0n/bA2eY7p0u0hzF9OpdornnQHAqkXwHn1RqB6gG4xv8JUtTancNzw3m1ku/MHby/n&#43;50uwKHFkJTT1E2723CL83OhAt09Ozup&#43;ISExIRw/ugvczjlQb4HECNotNeqwQH4N/MCMR7flmcmhdAaRmn0YgqLmRJfS442ckBwAlA9jPe5XKFezWmPwCibTxSCrfsO2MKU6W3OL1ej35pA9BS5ALzwLvPoKkSFLnz5XiXeN79CK8a/nVwAPxhWf8/dQXhdyQRwPmb5OeMk70v2AhhvNsV/i0DJMJ1UHYvGPU/SJRSpKelw3laD9nJfP3TxLg1zrc/YKRKi9Yz/5Vw/qRdJPi46LPQfpyqU4NIfZa4UIl3A7C7XW7RN4xuKvv/aiKEIKVPCiw0ES0nfK8r4SQ4eKOpyznxO2HBpvCXx3kdrjBOrd9uB94I5OUb8SkPcs44n8TbAbgDNt7t9kv8zyXBkXbZIpadA85kMiGrbxYa9nkCnjT3KjntoYTAUgz&#43;SKDqIXvzfnlcqP0P5CDYZ/EtD9QAzbniZEcYiXe5XCLw86rtnnbZznV3LlxZTk4O7NWAvVHpcXdOtYuAtuzMu3r/C5wf3&#43;2fdrzdS0nY3N7h3DaNxMfFxdn79u1b53K5aFdXV0RmnHM6V6txFrjEhESkWDNQv98RsZ1u0&#43;VM7f8iuGCvnPhMBwBvxNQ7m7x/E&#43;LL4&#43;eEc7Ur8NgY69u3b11cXJydFhQU4Oabb65lnDlbWlvC2vefZPMD2s3r5KoScI67c5FwQ3KGofWkGx4bC15eaeKwZ7M797/G&#43;YfIeRCOkOCZE8CplpucE87eKIMz2G&#43;&#43;&#43;ebqgoICqLH6FsaYs76&#43;3pSZkQntZoy2o&#43;ecfGrw59idC434eUFAWmoasDcKZeu7EJUsRlyBBFGnp6xeDLlbld5gwtXpRUNnbPoM7baccIEr3A7vw0n&#43;x6SLGWNttbW1cd76AYR2EoROiLNJYSdS0ELzJ&#43;AIYLFYMGroOPyw4Qd0eNwg2ecBRnPkhnulxH6h1U1vAlY/sQ1eUQSwzi4AFZxzJvokqAqA3WaznfXu3E9KYRy2cw36iKKI/BH5MBgMWLNpG2zZI0Cn3AoSFRMQM98ayu9xqxNItZ1qpkrD/yUu4IZrxkWCJrXqAqpNk9C2zwInv/UIeFOVHUANIYRSAPTBBx9sysvLa&#43;rs7ER7e3s3ovrtKQ/JC2N7g3Chebx7mRbvzzsLnLY/Op0Ow4YOw/SplyBx/1fgH8wHd3QGiOS3h/CSQ0swjW32LZQDjf8vcNqh&#43;XDc66kF5o9amRBf9DewN3A2ON7ZCrQ1IC8vr&#43;HBBx9sAsAoAJaUlMQMBsPm5uZmqD9zEei8RtK1Y/IWaJyXMHYpNI90L9PitXHm3uJUn0f9m1KKgQMG4PqrrsLAtlIor88Fry4BeOCJWy&#43;VuEZqtN88UNatzs&#43;J4xpl4S1TvXYCX54Wx73XxAc/GxyvPcV4Uw0zGAybfE8yUe27z763O&#43;zo6Ozo3SbJ/6UUOk8IQUpKCq684krksjawf/8ZvOwYgtRxoLIvTy3rLrH/G5xa3kuvMki7nAUOHKgvZ2hrAIDNUF9pyn3vqM/MzDzIOe&#43;ora31P17sh/4f3Z1TVVzYe4MjNjYW0668CuOSoqB7&#43;yGwozvAZY&#43;3gkYtB6JhpHuZNyNA558T5w&#43;1qteBMXCijk&#43;jzqFtq5c4txO8skiEx9WWmZl5UK2nvqMeS5YscQLYXlpaCodDEwjhGtV7Bn780rtzPS05ic/oRUdH4&#43;KLLsYFwwdDv&#43;wZ8E2fgLscPm3ri2rxgO3lGlWtdXT9dX4uHNDdNGh8JsK1WN9pJ/9mUO9xsHcydvKgG8D2JUuWOLnvLeXarVgZwJa29rbgn/zS0jOiFuoFo/6XKdKE9Jo&#43;SJKEMYVjcEnBMER99x&#43;wb94JCKlaifjsZcQVBg8W7p&#43;K8xvqkC5rx6LBgWjeIHQ2OHsHRVUxBbADXh4D0LzfNDo6mi1btmw3Y6xl127NS/X/f0g9zTtfmSAIGDVyFK6fejGS930JtuRJ8PYm&#43;LfJOLyqUUNAvwHTLMv8jMRPxRGfpIb6Bto2tTjNOM8Cxw5vAdzOjmXLlu2Ojg78rClF4M1XzGg0FgM4WHq6tJudP5f0S&#43;zOnU2ilKJ/v/649sorkFl5APjwr&#43;BNVd5CEtgq5T776OVT8GpDtd1au35uOPXD/Thvvo/R6k6jr3Iwo3uJk2WwXd/IAA4ajcZi4ntfvfbZOQDAhAkTaq655prdNptNLi4uPmcC&#43;/sXzlaTHsp&#43;Iq63femT3AdDB&#43;dCOroFyn&#43;e8nr8nMPPHv/SS/PhARvqL1OXU&#43;eC830TTgI4HrIK4MSPI1pPvpc4XlkE1Jfjmmuu2TlhwoQa&#43;DQ8ISToVShISkqS09PTP3O73bOLThSl5ubmQhS7vW/v7JLP&#43;fH/GWIPe4s71/upiTGG1tZWlFeU48CRvXKHUE/7zzCBKUdR8f6tcPS/nJKLZgF9&#43;oFQ9bkA9dSqJqkm2m/jgXBLqzPi1EsE/lS9d6ItQwDnD5n3AscVBWzXN4CjsyE9PX1lUlKS376rIdugX5SeNm3akRUrVhSVV5SnHj56GHlD8yCIwrk7cL5xOpwOlJSUoLmlGYmJiUhKTEJ8fDyoQMO3rdLH962e3w89jx8Wp0mMMbS0tqCoqAgnSo/DHdvCEibrMHBEPNVbBHDOkTBUZtXbv0bFh8dkJX&#43;aSMdeAWJJ8BNV2y4H0QTrNGoW6tM5pNc4Dvh8jIDTp12naHHeaoG8M&#43;HQWgdeegiJ8XHHpk2bdsR7Kw6opl1jM9WfHWNDhgz5XVFR0evm6Cg6IHsgJk6YSBMTEkEpPStVyzmHoigoLS3Fzr070K40uM39uWivFGTeIUpRpmik9UlD38wspKWmQa/Xg1IadJ9ebRKF0SqMMXR1dWHHrh04VV4CMdXB0i/S0&#43;g0HaRoGsIEDtnF0VLsZqe&#43;Y7TV3pdhxqOU9MsDRF3ARnP41Tr3SZ7PrHq9blUyVWWg2toIuOCNFa0m8Bt3f55/iOos6AHHwcH3b4Dy3hPIzUi5&#43;9ixY&#43;/C&#43;1t0qpBT/wvvfN&#43;Mc04nTJiwuqioaJ4&#43;Xc7qSD7FPl5zCtlJQ&#43;jI/FFISkrq1SNSDocD5eXlOHjkAKuXS2nqRKOcPcYsiXoKEC65Ohlai9tRU9LkLjq8R5K3U1ilBJYan0lTU1NhjbEiJiYGUVFRfnPT0&#43;6cynCn04mq6iocP3EMpQ1FiB4IeeDdJmrJsFD1zb3&#43;LV4N/UQ9QeIwPY0fBNTuLaenPrkbtsTzgcJpwIB8zWZPiIT5&#43;hJY3qkM8trkgPoNlUyvhfAfrvCZjMDQeDCOBxh8RpzTBrZ1JWBrL50w4fqviOZHon28ZiTUS&#43;ac05qaGik9Pf1xqiOPjXk0QRJ0hNXucVD38Wg5p99gMX9EPqxWa1iGK4qCyspK7N2/Fw2ucmYdwVnqGKNoiBM0gSDtGpZDdnLYG2TY6hXWVeeBrUqhvMUoG5QYMTYmFklJSUhPT0diQiLUN2qHJrfbjeLiEyg6cQINrgoWPUShicP0sGZLmkePVK3hU6EIlZxAna5qGXUHHKg9amAd5tGUFlwMct6FIAajxgRpRBsaBytknL/07hwr3gdl8e2A2/l8VVXVgtTUVDcJ/BiR18Hj2hfb&#43;pLH46EvvPBCweOPP/7vPqOMA0beF2/gMkdHhYdVfOeEp0qP8SMn0ZyBOVB/FlyWZTQ3N2P7rm2oaipjCeM4TR1rZPoYgRKBBDYVAnpOpYLmG2AKoDi9Z8CdbQprPemmnaWcOaop1REJiXFJyMroj7S0NFiiLeCco6y8DHsO7oKdtiBxtICkAgMMMYL3RU5B91Jp1tOLhQJ1mczhaldQd8DFKvdQahP6MT55NiWDRgFRVpBev9dHy03tpA/pi2Y6klBcYBr4zEt4HBQZyr8eBN&#43;7tmLhwoW3P/zwwz/odLruv1gZshWq/jYZtm7dKt16660LqhrLHiq4J15MzNNTgIAp3pcfnF5rZ1HORDoidyRirDEoOlGEk42HYR1KWMakKGpMUCU8mLH&#43;IXj1lXcwqkT46dIdJzs466j00LZSN9pOe&#43;BsUWBg0SBcgBzVgZSxBiQXGKEzaonjGyQJ0NtbdqZn4LrjPHaOun0OVr3LhXZlAFVGTgcZMh5ITA9MgGDliSAxD5JQlRA&#43;GnBNAdHUD7LnIY5DBBw7shXKm4&#43;w/onWN99///0HJ06c6IbmlymgkfiwSyzOOb311ltzl3649LvMyVFJQ39rpaLeb8DgbFfQdMSJ&#43;u0yWLsO5mEe1me0gcZkSd534Klqzm/XQtVtwDP1nj9XZ7TWa&#43;&#43;OAydeSWxT4GhWwGQOS4YOkoVqCB2w49prlaYIZbyGaz3hOAdcbQwtJ92serdMG9qzgdzxoOOmgSRlBjOccz&#43;vghgYWhYkFCEmIBQHTVkYHHfawP67EGzzpy2zbvnt1Pfff38/Cf29ORXBGFOlnCL4F6VZa2srHTJkyOLmzvr7zrszDn1GGoN&#43;vIAzDtnhfdmRaKSgOrV/IZ3SiJJ/yaMdmJYYWvV/JpyGICEatNfpXHCce8fdUeFB5U4ZTWV6ePrkg4&#43;&#43;9BLotAAAEMdJREFUGqTfcCAqRrMa6L4SiXyETavkA5OvZ4yvlMPryS95En3Mhv8cO3bsjtjYWBZOqAFAmD9/PvcVeuef5lkfo9HIS0tLS3dt33Oxo1lJSh1tooIEf&#43;cIIRAk7&#43;O31G/HVbWjdkldh2oOTfo745N4BJ9O7S3OHyMnqrB7w5gBqxhMPK2XxH0O1rnhCAQdgSleQHK&#43;hKQcBQbHaZD9a&#43;Dcv4&#43;xxgbvrLXEgQjaZ9h5wCPX2hRtDd8E90ZgA/cPOmwaLmxt7wBb&#43;leg7nTFrFmzfn/99dc3&#43;nik/h4Uh3fJTgghRJg/f75KdDWT&#43;6SeACCTJk3qbGpqIts27JlkiBNIbLa&#43;u0fDAx0OSC3xmSRvR7vtXmlEzS/J3sJe4&#43;CbHEHuj8ZmEwSONwcCP946PydOMlPEDpSQOESP2LgWLtXvJ7atO5l87CCBxwkkpoH4XgQVdJpJbUxzTVTToE6CACl9pPEJSQiOb10J9v1H7K677nxl8eLFX&#43;n1eubtPvexUu2ud5zC008/TTWZ6rvZVKnner1ecTqdJ9evXz&#43;kuaxraOxAPdFbaUC6fQ5HwH762eQnTEBDB4jpJwKHV9p8I1Sjn73BqTZdW0dljuovAAEJCz4m9vPiCCEQDARRySJJGKxH30KFR6GCyPt/gLLxU8gNTeCCDhBEQKcHoUJAg2l9DwS3Db88BAQiCAeAlx2F8tHzsIpY/8c//vHp/Pz8Fv&#43;08JpxdZowAF6R9zl3FN7IDiKl6dOnF36x&#43;ouvMyaaLENutoqSmVINZzQSG2A8uk0GNflVRNClryTMJIqEU1VygGFqIKXbhPERL7CO/2VwipujrdSN&#43;h&#43;daK0U0E6HgvcdAZI1DGTACBBzbECDacftt2G&#43;MSOEMKrz3NUGZflC8G2r7NdeM&#43;2azz//fENIGL7bch0ACGNMq07DVgKAQ4cO0bFjxy72wPnQsFlWZJwfhQDjwyFCVaXWOfF58D&#43;xzJsT4o2r9EJ37/xXK/OtRJxtCrpqZDQccaOuJIo5hRRKckaBFFwC2n&#43;4hnQkwHi/yEMz6QNlbOvnUD78G4xEeXPnzp33Dh8&#43;HCE8DDcJKAl9baYmrOfXAmre7t27&#43;4wZM&#43;YNKuKyCX9JEmP66cVuS27/hOSh7fnVU9gNF66q9t7j1PsEET6UEX5J5Jr4Efn1cYyw1pMuWrvXidaTLthJGvMMvJCSnEKQ5L5ATAJgjNaEAkKcWw7wsiOQn70JkD0bdu3adXthYWGVzz9j4finJs45hPnz51NCCFcrqBMgJI8CIJRS25EjRypOlpyabG&#43;UrXGD9ESK0rjDGskP&#43;AABe63dUvTbTV&#43;e365p/&#43;8VDprba660voIqJJq/f20cISDGeBGJwwzok2&#43;AJd5OTJ2HIRxdB/eBvZBPHAYaq7wNmaJBRElDBwLeVAXlw2eB&#43;vKyKVOmPDx79uwj0dHRXLs686l8HmrCCSHUv47Xhm5ViVedPdVmcM5RW1sr3nzzzTO2bt&#43;8LOvSaDFnugU6o&#43;o7BBkp&#43;NWyr8PBEoGgGawuUIh2dp8B121ty9WWNFugmt5Aoz20sF8DBwQvIcEBpnB4bAzONobOKjcajoM1Vf6/8q4uNorjjv9mdm995w8wxhzEMSExUdTQmlamakmOOq5CQEIVEQ9IpCiReECpWpfUjauQh9a2CiRWFUU8RkhNRUMg4gEFJaQFQg2xcUgaiIzlNJDYBoyxD2Pf&#43;ey7897eTB92Z3d2b882zoejdqW9252Z3378v&#43;c/O7vFMNSFlFetAamuBVmxCgABe&#43;c1sNNvsNpH1z57&#43;PDhv1VUVIjxdj93nVNGJI3O99Yru04IAwAQQl4hFL96eFupVrWxWDVfDQ4XAXKj1tw6p43jy&#43;xAaTY4p8LWBpu4kot0Jlx6DjqvOO6DF0Q0C4wphvHrGQxdmkqP9&#43;vB5EQwnSSVGu/vAVj2Tc7505jDQmVNF&#43;lbeeGcM259elpYBs456uvr93OGM18cH2dDn6R0s55bF2yuRGgrgeu4sqY6bUy7zV1l0&#43;PMEuehRgj/6dSY10K88jP/OHE3nDs4ERoCtvxAKSBY9KCGh7eWBNf8pgz3fN8w0H&#43;ZgWXb6uvrWyz&#43;uXgq/VOfMgCgSnNzM7F6&#43;bY/kBI43PL3RPIdhBDCH3vsscTU1NTFC&#43;c/XjPWm15euEQ1ipaqitufWIkQKcixa&#43;x2xF1&#43;FzhBIalL6wRPjjt09CrH/84vjkAkjJzbEbkE23hIBuB2z5TR988JjaeVzoaGhl/v3bv3qqZpEIk32BSxGSzPGzMvxbLaSnNzs2A2hzk6x83uvflqBWvbroPpHkggEMDGjRtvp9PpzjP/OFsX69PDJZUBoyisKnIYI67cY7FnNzo3A85Vbic4pDN7CO78k&#43;8Mjts/zi05Q1YObuSzKaP7YIwmh43u3bt3P9Pa2vq5pmnimGKsRWa&#43;Y1qkOkuhzZStnQ8n8jMpIJD8PrG5ARE5ckIIamtrR&#43;Px&#43;JftZzp/mriZWbxopcYKFipU3LnDIiKtDgGEuDttYAVus8O5CCoTWKggcbaJTfD5x5ldPDiW34a7LRkAJAYyrPvgGJkYMK7s2rXr9/v27ftYURSRkhXWWu6r28PtxPWeNXFkcKWpqUloMoVjCuRugZBYl/MXAZaiKLyuru6aYRhffnDqw0fuXE2VFYVVHipXiaMNxDknt8w44RL/iE0AR75mxolUr1tQZDEirm1Z&#43;74TOEIgnr83XRuRjCQBZ0LTx0hqkPQ2NjY27Nmz53RBQUFW8Is4kuXSdkIIk/kqsY4C4Epzc7MMcGm9t/8nL3J0raoqW79&#43;/dXJycmr779zbs3YVT1s&#43;Xzpg4rWuLfNXMFs4rpqcdOyLk2Hc1rJ/7nXzSHciTgbmVecEGZZtZ1dDs6A4UtpvftgjEzcND5vbGz83csvv/ye9TSNnYGDY5m9pl4mjdi3cUpzczOkQpdWe8uFNQDckgzL/0cikf5EInGl/UxndaxXDwcXKUZxRUARWmxrBLfic4&#43;pJJ4ycc3T4bzkFricKECOHXIYNR84xzXIQkEIwBlw66OU3vNWjCaj2e5du3Y17tmz533pESqZwX7M9rtIF847Ose93TlRJmIAT0pX&#43;BcA4Jbm92Wz2TMfvH9&#43;zXBX8h6qgpVUBhSqOgyzNVgwU2ivrQRSYocjLw6eNnJg6MqdS6Nscpt5xYFYKWh3bJdJcfSfnjA&#43;eyvO2aRyYffu3U/v27fv35J592qwl&#43;GizDUO78UJU2/ndb2rxXxZEGyBkHoDwgpQRVHY448/ficej3e2nzv/4EjP1ApuIFC8LAC10BzOJcQhjpB8SfAhpybd53bj7NE513EsIbFdhsQQ4pyLzzfObisEmSA1kkXviQSuHh9PZ6f46YaGht&#43;2trZ&#43;riiKzSMPo3OYzjmXn5/Pi1OamppsJiK/ufddJEsgtN/GRiKR0fHx8fMfXfgoEO/Tf5wYMFByb4AHFylCclwpWyco4q66nLSsVOeOBIhDCouwXgY4Fnb&#43;cbKAAEC8X2fdb8Qyty4kOc/itfr6&#43;j/t3bu3V9M02SB4NZ3KORg4FljOwvri7mp0DkBOZk8u9&#43;JEeV1d3S/PnTv3R6qS&#43;1fvKA0urQmxQJEzN///cXROlGUmmDF0KYWuv44ZPIv&#43;2traP589e/bN2Y6yQXq5haD/bHCu5&#43;rl0Tnvwa1yZkmZyNeL&#43;fXMK2mygNy6dUvdsWPH2pMnT/5BDZFfLK0J6Q&#43;sL9ZKV2rU8XVCg0xqiVy3rPEywRx34A5guYU3ZUXC5tzb/OIAYOyLKdZ3akIfvpjSjDR/Z8OGDX95/fXXP5QGXMSSbyxlpgcu8uLuenROZrDcJg/O3h8cHKQDAwPL6urqnk&#43;lU88ULlHKHlhfbFSuK9ICxdTx75Yz/F8dnbO0HAPtk0bfqQmaHMmOhoKhg21tba9UVlYOVVRUzIWRM7X5&#43;kbnZnNBUnu7vqurizY1NdW0tbU9H4vFNpVWaQtWbioxllQH1UBImHv4Exq5rkCqcEeIAsvF7vyPzmWSDNHLab333YQW69PHS0tLT9TV1b3S0tJycfXq1d/qN4Hk5&#43;pnMvXCX8g&#43;XH4W33XhsqmX21pF9OjRo6WnTp3aduDAgZ1qIVm1rCZE7/lJIcLVQdX8TpvjG53uHJeieg77V7YGkjsAcSyIe7jz28UxA4h2pY1bHycx9EmKGSnes3PnzgNPPPHEka1bt47CmqnsoaGcpIG07fqfK45Yw6yMECL&#43;haYKUA7TfLZnfLjPb0kkEtoLL7xQeezYsWeHhoeeCRTRskUrNbZyU0lw4f0a1JCny&#43;Q17eA2fZ3QV4qu7VYSQ0TQhW8WB25&#43;GCDer&#43;PLdxPpsV6dZibZyLKly97YsmXLa62trQMlJSW6bBXzMBF&#43;9PyquLxz58Q4r&#43;TDXScSddKJvHXixH7CIV8gGxsbo88999xDHR0d23t7ezdTFd9bvCqoLf9ZkV76gKaGyhVKFcEIW608zHBG&#43;examVGSgNga&#43;g3gOANSI1nE&#43;nTjxgeTdKQnbXAD/6mqqjoeiUQO7d&#43;//4o1w0V2g2Lxo5Of5uKr4qadOycfyC9q92I8ODtolKXR2xWTtmlHRwdtb29f/eKLL24C8LRSQKoW3BegC&#43;7T9OXrCoMLV2igAcENAF5GmEcy912aKLqM&#43;MZw3ADi/TputCfT49f14PiNjJFN814Af3/ppZdOrFu3risSiTBvzOOhSw6DfJTMS&#43;854fLNnXNNpM&#43;3TMNEP9x07sDez2QyNBqNqk1NTWUdHR2b7ty5s/X2yO0fBApJuHCJinvXFqrhH4YQKKJQCwlVNPeoV6478L9eu2wOOM45sjqHkeTIJBmGP00ZNzuTRmrEoJkkjy4pX9K9ePHio5FI5ERLS8toOBw2/KYq3w1dvu66nBcjzHaRo2ux&#43;ASF3u6ft50cV9jt5bzAyZMntePHj68aGBjY8vbbb68F8COqoWxBZYCVrizQFiwPGEVLVRRXBFTzO&#43;lATlcAgrGizqe74Fvm4DgHpuIMEzcz&#43;mTUwPiNjBbr1Y3EjQyyOh8F8OmTTz75YWVl5bHNmzf3bNiwQe6L&#43;7k4uS6fWfZi/GKrOeFc78CZCZivOzdNN09uM631yHd&#43;gYtGo2pHR8eyVCr10Pbt22sAPALgUaoiXFCqIFSu0oIF1AiVqUbJclUrWhqgRWEVwVKFgdgBEPIzHTATRzAIB03HsnQyaiA5nGHj1zPJ1Fg2OBXPqqmRrJGOZyk3EAVwHkDnoUOHLoZCoSuRSGRIfrvUd3khjDGvT3Y38EnZerp19kuT8uDyRf7ieHbSSNR5MoEuHOcciUSCAlBHRkaC27ZtWwVgXTqd/vnly5crQFBIKIrNlRQqGmGBYsoChZQGQgRKAWVqiDDry8owUgyZJGeZSWZkJhkz0owaaR5kGa5yxpOcYYIzTIAjWV1dPRgMBv8FoP3IkSM95eXlaQCG9cbI2QRfftuAv9Dn0&#43;qvBTfruXP5Fo&#43;p9w0AfbTdZeK9wYmnLi/OOicAIBqN0ldffbUUQAWA&#43;wBUHj58&#43;N5r165VAFgAQIP5JWV5BczvsMmrDmB8xYoVg0899dRNmF/vuA5gsKGhIRYOh/MFqq4ekd&#43;2Bzdt0swnDT6rrN1scbOeO&#43;e3SDduM0Sqy8nayReIPIkfuCV21jgp9yD22aVLl&#43;jo6GghTKZrMD&#43;ZLq&#43;A&#43;WJfedUB6GVlZcmamhrXfcr36Hc&#43;T3wyLU6inegGz8QwX189V9x/AeiDzxrlHYfKAAAAAElFTkSuQmCC</span>
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
    $UserName = "1143"

    # Credentials för Scoutnets API api/group/customlists
    $apiNyckel = ConvertTo-SecureString -String "aee1a01d1b1849e8c80ed9c98719e1c810e61006" -AsPlainText -Force
    $conf.CredentialCustomlists = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($UserName, $apiNyckel)  -ErrorAction "Stop"

    # Credentials för Scoutnets API api/group/memberlist
    $apiNyckel = ConvertTo-SecureString -String "d1dfa92a669814b43d86ded15a9f6ab1b039c2ad" -AsPlainText -Force
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
    #$NewValidationHash, $mailListData = SNSUpdateExchangeDistributionGroups -Configuration $conf -ValidationHash "Tom" -ReturnMaildata
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
