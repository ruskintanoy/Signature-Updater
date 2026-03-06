# Classic Outlook Signature Updater - GPO Backup Script
# Uses ActiveDirectory module to fetch user details and updates Outlook signature based on a template

# Signature name
$signatureName = "SPAARSignature"

# Load AD module
Import-Module ActiveDirectory -ErrorAction Stop

# Get current user
try {
    $user = Get-ADUser -Identity $env:USERNAME -Properties GivenName, Surname, Title, MobilePhone, EmailAddress
} catch {
    Write-Error "User '$env:USERNAME' not found in AD"
    exit 1
}

# Load signature template
$templatePath = "\\prodc03\SYSVOL\prophit.local\scripts\Signature Updater\SignatureTemplate.html"
if (-Not (Test-Path $templatePath)) {
    Write-Error "Signature template not found at $templatePath"
    exit 1
}
$templateHtml = Get-Content $templatePath -Raw

# Replace placeholders
$html = $templateHtml
$html = $html -replace "%%FirstName%%", $user.GivenName
$html = $html -replace "%%LastName%%", $user.Surname
$html = $html -replace "%%Title%%", $user.Title
$html = $html -replace "%%Mobile%%", $user.MobilePhone
$html = $html -replace "%%Email%%", $user.EmailAddress

# Save to signatures folder
$signatureFolder = Join-Path $env:APPDATA "Microsoft\Signatures"
if (!(Test-Path $signatureFolder)) {
    New-Item -Path $signatureFolder -ItemType Directory -Force | Out-Null
}

# Required file types
$htmlPath = Join-Path $signatureFolder "$signatureName.htm"
$txtPath  = Join-Path $signatureFolder "$signatureName.txt"
$rtfPath  = Join-Path $signatureFolder "$signatureName.rtf"

$html     | Set-Content -Path $htmlPath -Encoding UTF8
"$($user.GivenName) $($user.Surname) - $($user.Title)" | Set-Content -Path $txtPath
"{\rtf1\ansi\ansicpg1252 {\fonttbl\f0\fswiss Calibri;}\f0\fs22 $($user.GivenName) $($user.Surname) \line $($user.Title)}" | Set-Content -Path $rtfPath

# Set signature as default in Outlook
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings"
if (Test-Path $regPath) {
    Set-ItemProperty -Path $regPath -Name NewSignature -Value $signatureName
    Set-ItemProperty -Path $regPath -Name ReplySignature -Value $signatureName
    Write-Host "Signature set as default for $($user.SamAccountName)"
}