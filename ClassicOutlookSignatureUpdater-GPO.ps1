# Classic Outlook Signature Updater - GPO Script
# Uses System.DirectoryServices to fetch user details and updates Outlook signature based on a template

# Signature name
$signatureName = "SPAARSignature"

# Get current user info from AD using System.DirectoryServices
Add-Type -AssemblyName System.DirectoryServices.AccountManagement
$ctx = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Domain)
$user = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($ctx, $env:USERNAME)

if (-not $user) {
    Write-Error "User '$env:USERNAME' not found in Active Directory."
    exit 1
}

# Map properties from AD to user object
$userProps = [PSCustomObject]@{
    GivenName       = $user.GivenName
    Surname         = $user.Surname
    Title           = $user.GetUnderlyingObject().Properties["title"].Value
    MobilePhone     = $user.VoiceTelephoneNumber
    EmailAddress    = $user.EmailAddress
    SamAccountName  = $user.SamAccountName
}

# Disable Classic Outlook roaming signatures
$setupRegPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup"

if (-not (Test-Path $setupRegPath)) {
    New-Item -Path $setupRegPath -Force | Out-Null
}

New-ItemProperty -Path $setupRegPath `
    -Name "DisableRoamingSignaturesTemporaryToggle" `
    -PropertyType DWord `
    -Value 1 `
    -Force | Out-Null

Write-Host "Disabled roaming signatures for $($userProps.SamAccountName)"

# Load signature template
$templatePath = "\\prodc03\SYSVOL\prophit.local\scripts\Signature Updater\SignatureTemplate.html"
if (-Not (Test-Path $templatePath)) {
    Write-Error "Signature template not found at $templatePath"
    exit 1
}
$templateHtml = Get-Content $templatePath -Raw

# Replace placeholders
$html = $templateHtml
$html = $html -replace "%%FirstName%%", $userProps.GivenName
$html = $html -replace "%%LastName%%", $userProps.Surname
$html = $html -replace "%%Title%%", $userProps.Title
$html = $html -replace "%%Mobile%%", $userProps.MobilePhone
$html = $html -replace "%%Email%%", $userProps.EmailAddress

# Signature folder
$signatureFolder = Join-Path $env:APPDATA "Microsoft\Signatures"
if (!(Test-Path $signatureFolder)) {
    New-Item -Path $signatureFolder -ItemType Directory -Force | Out-Null
}

# Required signature file types
$htmlPath = Join-Path $signatureFolder "$signatureName.htm"
$txtPath  = Join-Path $signatureFolder "$signatureName.txt"
$rtfPath  = Join-Path $signatureFolder "$signatureName.rtf"

# Save signature files
$html | Set-Content -Path $htmlPath -Encoding UTF8
"$($userProps.GivenName) $($userProps.Surname) - $($userProps.Title)" | Set-Content -Path $txtPath
"{\rtf1\ansi\ansicpg1252 {\fonttbl\f0\fswiss Calibri;}\f0\fs22 $($userProps.GivenName) $($userProps.Surname) \line $($userProps.Title)}" | Set-Content -Path $rtfPath


# Optional: Set signature as default for new emails and replies (Prevents user from editing signature if enabled)
# Set signature as default
# $regPath = "HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings"
# if (Test-Path $regPath) {
#     Set-ItemProperty -Path $regPath -Name NewSignature   -Value $signatureName
#     Set-ItemProperty -Path $regPath -Name ReplySignature -Value $signatureName
#     Write-Host "Signature set as default for $($userProps.SamAccountName)"
# } else {
#     Write-Warning "Registry path for Outlook signatures not found."
# }