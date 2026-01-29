<#
Developed by: Rhys Saddul

Overview:
• Register an Azure AD app called AutopilotEnrollment
• Assign Microsoft Graph application permissions
• Disable default "user_impersonation"
• Create a 15-year certificate (instead of client secret)
• Upload certificate to app registration
• GUI prompt to set School/Trust + Code → prefix used in cert name (e.g., TCS_)
• Output AppId, Thumbprint, TenantId in popup
• Direct admin-consent link for permissions
#>

# ==============================================================
#   PARAMETERS
# ==============================================================
param (
    [string]$AppName = "AutopilotEnrollment"
)

# ==============================================================
#   CONNECT TO GRAPH
# ==============================================================
Write-Host "🔹 Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes "Application.ReadWrite.All","Directory.ReadWrite.All" | Out-Null
$tenantId = (Get-MgContext).TenantId

# ==============================================================
#   GUI PROMPT – SCHOOL/TRUST + CODE
# ==============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
# Load icon for the window (must be .ico)
$iconPath = "$PSScriptRoot\autopilot.ico"
if (Test-Path $iconPath) {
    $form.Icon = New-Object System.Drawing.Icon($iconPath)
}
$form.Text = "Certificate Naming"
$form.Size = New-Object System.Drawing.Size(350,230)
$form.StartPosition = "CenterScreen"
$form.Topmost = $true

$labelType = New-Object System.Windows.Forms.Label
$labelType.Text = "Select type:"
$labelType.Location = New-Object System.Drawing.Point(10,10)
$labelType.AutoSize = $true
$form.Controls.Add($labelType)

$radioSchool = New-Object System.Windows.Forms.RadioButton
$radioSchool.Text = "School"
$radioSchool.Location = New-Object System.Drawing.Point(20,40)
$radioSchool.AutoSize = $true
$radioSchool.Checked = $true
$form.Controls.Add($radioSchool)

$radioTrust = New-Object System.Windows.Forms.RadioButton
$radioTrust.Text = "Trust"
$radioTrust.Location = New-Object System.Drawing.Point(120,40)
$radioTrust.AutoSize = $true
$form.Controls.Add($radioTrust)

$labelCode = New-Object System.Windows.Forms.Label
$labelCode.Text = "Enter code:"
$labelCode.Location = New-Object System.Drawing.Point(10,80)
$labelCode.AutoSize = $true
$form.Controls.Add($labelCode)

$textCode = New-Object System.Windows.Forms.TextBox
$textCode.Location = New-Object System.Drawing.Point(100,78)
$textCode.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textCode)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(130,120)
$okButton.Add_Click({
    if ([string]::IsNullOrWhiteSpace($textCode.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Enter a valid code.")
    } else {
        $form.Tag = $textCode.Text.Trim()
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    }
})
$form.Controls.Add($okButton)

$form.ShowDialog()
if ($form.DialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "❌ Cancelled by user."
    exit
}
$code = $form.Tag
$prefix = "${code}_"

Write-Host "✔ Using certificate prefix: $prefix"

# ==============================================================
#   REGISTER APPLICATION
# ==============================================================

Write-Host "🆕 Registering application: $AppName"
$app = New-MgApplication -DisplayName $AppName -SignInAudience "AzureADMyOrg" `
    -Web @{ RedirectUris = @("https://login.microsoftonline.com/common/oauth2/nativeclient") }

# Create Service Principal
Write-Host "👤 Creating Service Principal..."
$sp = New-MgServicePrincipal -AppId $app.AppId

# ==============================================================
#   ASSIGN GRAPH API PERMISSIONS
# ==============================================================

Write-Host "🔑 Assigning API Permissions..."

$permissions = @(
    "Device.ReadWrite.All",
    "DeviceManagementManagedDevices.ReadWrite.All",
    "DeviceManagementServiceConfig.ReadWrite.All",
    "Group.ReadWrite.All",
    "GroupMember.ReadWrite.All",
    "User.Read"
)

$graphSp = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"
$resourceAccess = @()

foreach ($perm in $permissions) {
    $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $perm }
    if ($role) {
        $resourceAccess += @{ Id = $role.Id; Type = "Role" }
    } else {
        $scope = $graphSp.Oauth2Permissions | Where-Object { $_.Value -eq $perm }
        if ($scope) {
            $resourceAccess += @{ Id = $scope.Id; Type = "Scope" }
        }
    }
}

Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess @(
    @{
        ResourceAppId = $graphSp.AppId
        ResourceAccess = $resourceAccess
    }
)

# ==============================================================
#   DISABLE DEFAULT USER_IMPERSONATION
# ==============================================================

$appOauth = Get-MgApplication -ApplicationId $app.Id
if ($appOauth.Api.Oauth2PermissionScopes) {
    $modified = $appOauth.Api.Oauth2PermissionScopes
    foreach ($s in $modified) {
        if ($s.Value -eq "user_impersonation") {
            $s.IsEnabled = $false
        }
    }
    Update-MgApplication -ApplicationId $app.Id -Api @{ Oauth2PermissionScopes = $modified }
    Write-Host "⚠️ Disabled default user_impersonation scope."
}

# ==============================================================
#   CERTIFICATE CREATION (REPLACES CLIENT SECRET)
# ==============================================================

Write-Host "📄 Creating self-signed certificate..."

$CertSubject = "CN=${prefix}${AppName}Cert"
$CertPfxPath = "$env:USERPROFILE\Desktop\${prefix}${AppName}.pfx"
$CertCerPath = "$env:USERPROFILE\Desktop\${prefix}${AppName}.cer"

$CertPassword = Read-Host "Enter password to protect private key (.pfx)" -AsSecureString

$expiry = (Get-Date).AddYears(15)
$cert = New-SelfSignedCertificate -Subject $CertSubject `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -NotAfter $expiry `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider"

Export-Certificate -Cert $cert -FilePath $CertCerPath | Out-Null
Export-PfxCertificate -Cert $cert -FilePath $CertPfxPath -Password $CertPassword | Out-Null

$thumb = $cert.Thumbprint

Write-Host "✔ Certificate created: $CertSubject"
Write-Host "📂 $CertPfxPath"
Write-Host "📂 $CertCerPath"

# Upload cert to App Registration
$keyCredential = [Microsoft.Graph.PowerShell.Models.MicrosoftGraphKeyCredential]@{
    Type = "AsymmetricX509Cert"
    Usage = "Verify"
    Key = $cert.RawData
    DisplayName = "${prefix}${AppName}Cert"
    StartDateTime = $cert.NotBefore.ToUniversalTime()
    EndDateTime   = $cert.NotAfter.ToUniversalTime()
}

Update-MgApplication -ApplicationId $app.Id -KeyCredentials @($keyCredential)

Write-Host "✔ Certificate uploaded to application."

# ==============================================================
#   RESULTS POPUP
# ==============================================================

$appId = $app.AppId 

$resultMsg = @"
✅ Autopilot Enrollment App Registration (Certificate Auth)

AppId:
$appId

Thumbprint:
$thumb

Tenant ID:
$tenantId

Certificate Prefix:
$prefix

Saved:
$CertPfxPath
$CertCerPath

⚠️ IMPORTANT:
• Save the certs to IT Glue
• Needed later for Autopilot automation
• Approve API permissions via browser pop up after this
"@

Write-Host $resultMsg

Add-Type -AssemblyName System.Windows.Forms

$form2 = New-Object System.Windows.Forms.Form
if (Test-Path $iconPath) {
    $form2.Icon = New-Object System.Drawing.Icon($iconPath)
}
$form2.Text = "AutopilotEnrollment Registration Complete"
$form2.Size = New-Object System.Drawing.Size(600,350)
$form2.StartPosition = "CenterScreen"
$form2.Topmost = $true
$form2.Add_Shown({ $form2.Activate() })   # 👈 ensures popup is front-most

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.ReadOnly = $true
$textBox.ScrollBars = "Vertical"
$textBox.Size = New-Object System.Drawing.Size(560,230)
$textBox.Location = New-Object System.Drawing.Point(10,10)
$textBox.Text = $resultMsg
$form2.Controls.Add($textBox)


# --- COPY BUTTON ---
$copyButton = New-Object System.Windows.Forms.Button
$copyButton.Text = "Copy to Clipboard"
$copyButton.Size = New-Object System.Drawing.Size(150,30)
$copyButton.Location = New-Object System.Drawing.Point(120,250)
$copyButton.Add_Click({
    [System.Windows.Forms.Clipboard]::SetText($resultMsg)
    [System.Windows.Forms.MessageBox]::Show("Copied to Clipboard!")
})
$form2.Controls.Add($copyButton)

# --- OK BUTTON ---
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Size = New-Object System.Drawing.Size(100,30)
$okButton.Location = New-Object System.Drawing.Point(320,250)
$okButton.Add_Click({
    $form2.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form2.Close()
})
$form2.Controls.Add($okButton)
$form2.ShowDialog()

# ==============================================================
#   ADMIN CONSENT INSTRUCTIONS
# ==============================================================

Write-Host "⚠️ Grant Admin Consent manually in Azure Portal:"
$consentUrl = "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($app.AppId)/isMSAApp~/false"
Write-Host "👉 $consentUrl"
Start-Process msedge.exe -ArgumentList "-inprivate", $consentUrl

Disconnect-MgGraph | Out-Null
