function Invoke-SNSAddOffice365User
{
    <#
    .SYNOPSIS
        Creates a new Office 365 user or updates data for exixting user.

    .DESCRIPTION
        Creates a new Office 365 user or updates data for exixting user, based on the user data in scoutnet.

    .INPUTS
        None. You cannot pipe objects to Get-SNSAddOffice365User.

    .OUTPUTS
        None.
    #>

    param (
        [Parameter(Mandatory=$False, HelpMessage="Data for the member from Scoutnet.")]
        [ValidateNotNull()]
        $MemberData,

        [Parameter(Mandatory=$False, HelpMessage="User language")]
        [ValidateNotNull()]
        $PreferredLanguage,

        [Parameter(Mandatory=$False, HelpMessage="Domain name for office365 mail addresses.")]
        [ValidateNotNull()]
        [string]$DomainName,

        [Parameter(Mandatory=$False, HelpMessage="UsageLocation.")]
        [ValidateNotNull()]
        $UsageLocation,

        [Parameter(Mandatory=$False, HelpMessage="License assignment.")]
        [ValidateNotNull()]
        $LicenseAssignment,

        [Parameter(Mandatory=$False, HelpMessage="License options from New-MsolLicenseOptions.")]
        [ValidateNotNull()]
        $LicenseOptions,

        [Parameter(Mandatory=$false, HelpMessage="Credentials for office365")]
        [ValidateNotNull()]
        [Alias("Credential")]
        [pscredential]$Credential365
        )

    try
    {
        Get-MsolDomain -ErrorAction Stop > $null
    }
    catch
    {
        Write-SNSLog "Connecting to Office 365..."
        try
        {
            Connect-MsolService -Credential $Credential365 -ErrorAction Stop         
        }
        catch
        {
            Write-SNSLog -Level "Error" "Could not connect to Office 365. Error $_"
            throw
        }
    }

    $DisplayName = "$($MemberData.first_name.value) $($MemberData.last_name.value)"
    $UserName = "$($MemberData.first_name.value).$($MemberData.last_name.value)".ToLower()
    # Convert UTF encoded names and create corresponding ASCII version.
    $UserName = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($UserName))

    $UserPrincipalName = "$($UserName)@$($DomainName)"

    $office365User  = Get-MsolUser -UserPrincipalName $UserPrincipalName -ErrorAction SilentlyContinue

    if ($office365User)
    {
        # Mailaddress alredy exists. Try with an extra number.
        For ($cnt=1; $cnt -le 5; $cnt++)
        {
            $UserPrincipalName = "$($UserName).$($cnt)@$($DomainName)"
            $office365User  = Get-MsolUser -UserPrincipalName $UserPrincipalName -ErrorAction SilentlyContinue
            if (!$office365User)
            {
                # Mailaddress not used. Uset i!
                break
            }
        }
    }

    $newOffice365User = $null

    if (!$office365User)
    {
        try
        {
            $StreetAddress = $MemberData.address_1.value
            if ($MemberData.address_2.value)
            {
                $StreetAddress += " " + $MemberData.address_2.value
            }
            if ($MemberData.address_3.value)
            {
                $StreetAddress += " " + $MemberData.address_3.value
            }

            if ([string]::IsNullOrEmpty($StreetAddress))
            {
                $StreetAddress = ""
            }

            $AlternateEmailAddresses = $MemberData.email.value
            if ($UserPrincipalName -like $AlternateEmailAddresses)
            {
                # Do not use the office 365 email as alternate email.
                # Try to use contact_alt_email.
                $AlternateEmailAddresses = $MemberData.contact_alt_email.value
                if ($UserPrincipalName -like $AlternateEmailAddresses)
                {
                    # contact_alt_email not usable.
                    $AlternateEmailAddresses = ""
                }
            }

            if ([string]::IsNullOrEmpty($AlternateEmailAddresses))
            {
                # Option -AlternateEmailAddresses expects an array. Create empty array.
                $AlternateEmailAddresses = @()
            }

            # Create the user.
            $newOffice365User = New-MsolUser -UserPrincipalName $UserPrincipalName -DisplayName $DisplayName `
                -FirstName $MemberData.first_name.value  `
                -LastName $MemberData.last_name.value `
                -StreetAddress $StreetAddress `
                -PostalCode $MemberData.postcode.value `
                -City $MemberData.town.value `
                -Country $MemberData.country.value `
                -AlternateEmailAddresses $AlternateEmailAddresses `
                -MobilePhone $MemberData.contact_mobile_phone.value `
                -PreferredLanguage $PreferredLanguage `
                -UsageLocation $UsageLocation `
                -PreferredLanguage $PreferredLanguage `
                -LicenseAssignment $LicenseAssignment -LicenseOptions $LicenseOptions -ErrorAction Stop

            Write-SNSLog "User $($UserPrincipalName) added for member id '$($MemberData.member_no.value)'"
        }
        catch
        {
            Write-SNSLog -Level "Error" "Could not create user '$($UserPrincipalName)' for member '$DisplayName' with id '$($MemberData.member_no.value)'. Error $_"
        }
    }
    else
    {
        Write-SNSLog -Level "Error" "Mailaddress $($UserPrincipalName) is alredy in use. Can not add a user for member '$DisplayName' with id '$($MemberData.member_no.value)'"
    }

    return $newOffice365User
}