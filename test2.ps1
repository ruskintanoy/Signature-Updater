Add-Type -AssemblyName System.DirectoryServices.AccountManagement
$ctx  = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Domain)
$user = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($ctx, $env:USERNAME)

$uad = $user.GetUnderlyingObject() # DirectoryEntry

[pscustomobject]@{
    SamAccountName = $user.SamAccountName
    UPN            = $user.UserPrincipalName
    DisplayName    = $user.DisplayName
    GivenName      = $user.GivenName
    Surname        = $user.Surname
    EmailAddress   = $user.EmailAddress
    VoiceTel_UP    = $user.VoiceTelephoneNumber      
    Mobile_UP      = $user.MobilePhone               
    Title_LDAP     = $uad.Properties["title"].Value
    Mobile_LDAP    = $uad.Properties["mobile"].Value
    Tel_LDAP       = $uad.Properties["telephoneNumber"].Value
    IPPhone_LDAP   = $uad.Properties["ipPhone"].Value
}