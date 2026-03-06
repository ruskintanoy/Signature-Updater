Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "SPAAR Inc. - Signature Updater (New Outlook)"
$form.Size = New-Object System.Drawing.Size(700, 500)
$form.StartPosition = "CenterScreen"

# Radio Buttons
$radioUser = New-Object System.Windows.Forms.RadioButton
$radioUser.Text = "Update specific user"
$radioUser.Location = New-Object System.Drawing.Point(20, 20)
$radioUser.Size = New-Object System.Drawing.Size(200, 20)
$radioUser.Checked = $true
$form.Controls.Add($radioUser)

$radioGroup = New-Object System.Windows.Forms.RadioButton
$radioGroup.Text = "Update distribution group"
$radioGroup.Location = New-Object System.Drawing.Point(20, 80)
$radioGroup.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($radioGroup)

# Textboxes
$userLabel = New-Object System.Windows.Forms.Label
$userLabel.Text = "User Email:"
$userLabel.Location = New-Object System.Drawing.Point(40, 45)
$userLabel.Size = New-Object System.Drawing.Size(80, 20)
$form.Controls.Add($userLabel)

$userBox = New-Object System.Windows.Forms.TextBox
$userBox.Size = New-Object System.Drawing.Size(300, 20)
$userBox.Location = New-Object System.Drawing.Point(130, 43)
$form.Controls.Add($userBox)

$groupLabel = New-Object System.Windows.Forms.Label
$groupLabel.Text = "Group Email:"
$groupLabel.Location = New-Object System.Drawing.Point(40, 105)
$groupLabel.Size = New-Object System.Drawing.Size(80, 20)
$form.Controls.Add($groupLabel)

$groupBox = New-Object System.Windows.Forms.TextBox
$groupBox.Size = New-Object System.Drawing.Size(300, 20)
$groupBox.Location = New-Object System.Drawing.Point(130, 103)
$groupBox.Enabled = $false
$form.Controls.Add($groupBox)

# Enable/Disable boxes based on radio selection
$radioUser.Add_CheckedChanged({ $userBox.Enabled = $true; $groupBox.Enabled = $false })
$radioGroup.Add_CheckedChanged({ $userBox.Enabled = $false; $groupBox.Enabled = $true })

# Log box
$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Size = New-Object System.Drawing.Size(650, 250)
$logBox.Location = New-Object System.Drawing.Point(20, 180)
$logBox.ReadOnly = $true
$logBox.BackColor = "White"
$form.Controls.Add($logBox)

function Write-Log {
    param([string]$message)
    $logBox.AppendText("$message`n")
    $logBox.ScrollToCaret()
}

# Update button
$updateBtn = New-Object System.Windows.Forms.Button
$updateBtn.Text = "Update Signatures"
$updateBtn.Size = New-Object System.Drawing.Size(150, 30)
$updateBtn.Location = New-Object System.Drawing.Point(20, 140)
$form.Controls.Add($updateBtn)

# Signature logic
$updateBtn.Add_Click({

    Write-Log "Connecting to Exchange Online..."
    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Log "Connected to Exchange."
    } catch {
        Write-Log "Unable to connect to Exchange Online."
        return
    }

    # AD context
    try {
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement -ErrorAction Stop
        $ctx = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Domain)
    } catch {
        Write-Log "Unable to initialize AD context: $($_.Exception.Message)"
        Disconnect-ExchangeOnline -Confirm:$false
        return
    }

    $orgConfig = Get-OrganizationConfig
    if (-not $orgConfig.PostponeRoamingSignaturesUntilLater) {
        Write-Log "Enabling roaming signature support..."
        Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater $true
        Write-Log "Roaming signature support enabled."
    } else {
        Write-Log "Roaming signature already enabled."
    }

    $users = @()
    if ($radioUser.Checked) {
        $email = $userBox.Text.Trim()
        if (-not $email) {
            Write-Log "Please enter a user email."
            return
        }
        try {
            $users = @(Get-User -Identity $email -ErrorAction Stop)
            Write-Log "User found: $email"
        } catch {
            Write-Log "Could not find user '$email'."
            return
        }
    } elseif ($radioGroup.Checked) {
        $group = $groupBox.Text.Trim()
        if (-not $group) {
            Write-Log "Please enter a group email."
            return
        }
        try {
            $members = Get-DistributionGroupMember -Identity $group -ResultSize Unlimited -ErrorAction Stop
            foreach ($member in $members) {
                if ($member.RecipientType -eq "UserMailbox") {
                    try {
                        $userObj = Get-User -Identity $member.PrimarySmtpAddress -ErrorAction Stop
                        $users += $userObj
                    } catch {
                        Write-Log "Skipped $($member.Name) — info not available."
                    }
                } else {
                    Write-Log "Skipping non-mailbox user: $($member.Name)"
                }
            }
            Write-Log "Loaded $($users.Count) user(s) from group."
        } catch {
            Write-Log "Could not find group '$group'."
            return
        }
    }

    if (-not $users -or $users.Count -eq 0) {
        Write-Log "No users to found."
        return
    }

    $templatePath = Join-Path $PSScriptRoot "SignatureTemplate.html"
    if (-not (Test-Path $templatePath)) {
        Write-Log "Template file not found."
        return
    }

    $signatureTemplate = Get-Content -Path $templatePath -Raw

    foreach ($user in $users) {
        try {
            # Pull user info from AD
            $adUser = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($ctx, $user.UserPrincipalName)

            if (-not $adUser) {
                Write-Log "AD user not found for: $($user.UserPrincipalName)"
                continue
            }

            if ($adUser.Enabled -ne $true) {
                Write-Log "Skipping disabled account: $($user.UserPrincipalName)"
                continue
            }

            $firstName = $adUser.GivenName
            $lastName  = $adUser.Surname
            $title     = $adUser.GetUnderlyingObject().Properties["title"].Value
            $mobile    = $adUser.VoiceTelephoneNumber
            $email     = $adUser.EmailAddress

            if (-not $firstName) { $firstName = $adUser.DisplayName }
            if (-not $lastName)  { $lastName = "" }
            if (-not $title)     { $title = "Team Member" }
            if (-not $mobile)    { $mobile = "(n/a)" }

            $signatureHtml = $signatureTemplate
            $signatureHtml = $signatureHtml -replace "%%FirstName%%", $firstName
            $signatureHtml = $signatureHtml -replace "%%LastName%%", $lastName
            $signatureHtml = $signatureHtml -replace "%%Title%%", $title
            $signatureHtml = $signatureHtml -replace "%%Mobile%%", $mobile
            $signatureHtml = $signatureHtml -replace "%%Email%%", $email

            Set-MailboxMessageConfiguration -Identity $user.UserPrincipalName `
                -SignatureHTML $signatureHtml `
                -AutoAddSignature $true `
                -AutoAddSignatureOnMobile $true `
                -AutoAddSignatureOnReply $true `
                -ErrorAction Stop

            Write-Log "Signature updated: $($user.UserPrincipalName)"
        } catch {
            Write-Log "Failed for $($user.UserPrincipalName): $($_.Exception.Message)"
        }
    }

    Disconnect-ExchangeOnline -Confirm:$false
    Write-Log "Signature update complete. Disconnected."
})

# Show Form
[void]$form.ShowDialog()