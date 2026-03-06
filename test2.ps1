$identity = "chrisr@spaar.ca" 
$other = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($ctx, $identity)

if (-not $other) { throw "User not found: $identity" }

$uadOther = $other.GetUnderlyingObject()

[pscustomobject]@{
    SamAccountName = $other.SamAccountName
    UPN            = $other.UserPrincipalName
    DisplayName    = $other.DisplayName
    Enabled        = $other.Enabled
    GivenName      = $other.GivenName
    Surname        = $other.Surname
    EmailAddress   = $other.EmailAddress
    VoiceTel_UP    = $other.VoiceTelephoneNumber
    Mobile_UP      = $other.MobilePhone
    Title_LDAP     = $uadOther.Properties["title"].Value
    Mobile_LDAP    = $uadOther.Properties["mobile"].Value
    Tel_LDAP       = $uadOther.Properties["telephoneNumber"].Value
    IPPhone_LDAP   = $uadOther.Properties["ipPhone"].Value
}