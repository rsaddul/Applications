<#
Developed by: Rhys Saddul
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# ==============================================================
#   MAIN GUI LAYOUT (WPF)
# ==============================================================

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Exchange Online Automation App Registration"
        Height="550" Width="800"
        WindowStartupLocation="CenterScreen"
        Topmost="True"
        FontFamily="Segoe UI" FontSize="12">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Text="Exchange Online Automation App Registration"
                   FontSize="16" FontWeight="Bold" Grid.Row="0" Margin="0,0,0,10"/>

        <Button Name="StartButton" Content="Start Registration"
                Grid.Row="1" Height="35" Width="200"
                HorizontalAlignment="Left" Margin="0,0,0,10"/>

        <TextBox Name="OutputBox" Grid.Row="2"
                 VerticalScrollBarVisibility="Auto"
                 TextWrapping="Wrap"
                 IsReadOnly="True" AcceptsReturn="True"
                 FontFamily="Consolas" Margin="0,0,0,10"/>

        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button Name="CopyButton" Content="Copy Output" Width="140" Margin="0,0,10,0" IsEnabled="False"/>
            <Button Name="ExitButton" Content="Exit" Width="100"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Load WPF GUI
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$startButton = $window.FindName("StartButton")
$exitButton  = $window.FindName("ExitButton")
$copyButton  = $window.FindName("CopyButton")
$outputBox   = $window.FindName("OutputBox")

# ==============================================================
#   Logging Helpers
# ==============================================================

function Write-Log($text) {
    $timestamp = Get-Date -Format "HH:mm:ss"
    $outputBox.AppendText("[$timestamp] $text`r`n")
    $outputBox.ScrollToEnd()
}

function Separator { 
    $outputBox.AppendText("------------------------------------------------------------`r`n")
}

# ==============================================================
#   WPF PASSWORD PROMPT (Password + Confirm)
# ==============================================================

function Show-PasswordPrompt {

    [xml]$xamlPwd = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Enter PFX Password"
        Height="260" Width="420"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        Topmost="True">

    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Text="Enter password to protect the certificate (.pfx):" FontSize="14"/>
        <PasswordBox Name="Pwd1" Grid.Row="1" Margin="0,10,0,10" FontSize="14"/>

        <TextBlock Text="Confirm password:" Grid.Row="2" FontSize="14"/>
        <PasswordBox Name="Pwd2" Grid.Row="3" Margin="0,10,0,10" FontSize="14"/>

        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,10,0,0">
            <Button Name="OK" Width="90" Margin="0,0,10,0">OK</Button>
            <Button Name="Cancel" Width="90">Cancel</Button>
        </StackPanel>

    </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xamlPwd
    $pwdWindow = [Windows.Markup.XamlReader]::Load($reader)

    $pwd1 = $pwdWindow.FindName("Pwd1")
    $pwd2 = $pwdWindow.FindName("Pwd2")
    $btnOK = $pwdWindow.FindName("OK")
    $btnCancel = $pwdWindow.FindName("Cancel")

    $script:PasswordResult = $null

    $btnOK.Add_Click({
        if ($pwd1.Password -eq "") {
            [System.Windows.MessageBox]::Show("Password cannot be empty.","Error")
            return
        }
        if ($pwd1.Password -ne $pwd2.Password) {
            [System.Windows.MessageBox]::Show("Passwords do not match.","Error")
            return
        }

        $script:PasswordResult = $pwd1.SecurePassword
        $pwdWindow.DialogResult = $true
        $pwdWindow.Close()
    })

    $btnCancel.Add_Click({
        $script:PasswordResult = $null
        $pwdWindow.DialogResult = $false
        $pwdWindow.Close()
    })

    $null = $pwdWindow.ShowDialog()
    return $script:PasswordResult
}

# ==============================================================
#   WPF REQUIRED PREFIX PROMPT
# ==============================================================

function Show-PrefixPrompt {

[xml]$xamlPrefix = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Certificate Prefix"
        Height="170" Width="420"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        Topmost="True">

    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Text="Enter prefix (e.g., SchoolCode, TrustCode, KST):" FontSize="14"/>
        <TextBox Name="PrefixBox" Grid.Row="1" Margin="0,10,0,10" FontSize="14"/>

        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Center">
            <Button Name="OK" Width="80" Margin="0,0,10,0">OK</Button>
            <Button Name="Cancel" Width="80">Cancel</Button>
        </StackPanel>

    </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xamlPrefix
    $prefixWindow = [Windows.Markup.XamlReader]::Load($reader)

    $box = $prefixWindow.FindName("PrefixBox")
    $btnOK = $prefixWindow.FindName("OK")
    $btnCancel = $prefixWindow.FindName("Cancel")

    $script:PrefixResult = $null

    $btnOK.Add_Click({
        if ($box.Text.Trim() -eq "") {
            [System.Windows.MessageBox]::Show("Prefix cannot be empty.","Error")
            return
        }
        $script:PrefixResult = $box.Text.Trim()
        $prefixWindow.DialogResult = $true
        $prefixWindow.Close()
    })

    $btnCancel.Add_Click({
        $script:PrefixResult = $null
        $prefixWindow.DialogResult = $false
        $prefixWindow.Close()
    })

    $null = $prefixWindow.ShowDialog()
    return $script:PrefixResult
}

# ==============================================================
#   BUTTON: Start Registration
# ==============================================================

$startButton.Add_Click({

    $startButton.IsEnabled = $false
    $copyButton.IsEnabled = $false
    $outputBox.Clear()

    try {
        # ------------------------------------------------------
        # CONNECT TO GRAPH
        # ------------------------------------------------------
        Write-Log "🔐 Connecting to Microsoft Graph..."
        $window.WindowState = 'Minimized'

        Connect-MgGraph -Scopes `
            "Application.ReadWrite.All",
            "AppRoleAssignment.ReadWrite.All",
            "RoleManagement.ReadWrite.Directory",
            "Directory.ReadWrite.All"

        $window.WindowState = 'Normal'
        Write-Log "✅ Connected to Microsoft Graph"
        Separator

        # ------------------------------------------------------
        # GET CERTIFICATE PASSWORD (WPF)
        # ------------------------------------------------------
        $CertPassword = Show-PasswordPrompt

        if (-not $CertPassword) {
            Write-Log "❌ Password entry cancelled."
            $startButton.IsEnabled = $true
            return
        }

        # ------------------------------------------------------
        # GET PREFIX (WPF)
        # ------------------------------------------------------
        $prefix = Show-PrefixPrompt

        if (-not $prefix) {
            Write-Log "❌ Prefix entry cancelled."
            $startButton.IsEnabled = $true
            return
        }

        $prefix = "${prefix}_"

        # ------------------------------------------------------
        # Certificate creation
        # ------------------------------------------------------
        Write-Log "📄 Creating self-signed certificate..."

        $CertSubject = "CN=${prefix}ExchangeOnlineAutomationCert"
        $CertExportPfx = "$env:USERPROFILE\Desktop\${prefix}ExchangeOnlineAutomation.pfx"
        $CertExportPublic = "$env:USERPROFILE\Desktop\${prefix}ExchangeOnlineAutomation.cer"

        $ExpiryDate = (Get-Date).AddYears(15)

        $cert = New-SelfSignedCertificate -Subject $CertSubject `
            -CertStoreLocation "Cert:\CurrentUser\My" `
            -NotAfter $ExpiryDate `
            -KeyExportPolicy Exportable `
            -KeySpec Signature `
            -KeyLength 2048 `
            -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider"

        Export-Certificate -Cert $cert -FilePath $CertExportPublic | Out-Null
        Export-PfxCertificate -Cert $cert -FilePath $CertExportPfx -Password $CertPassword | Out-Null

        $Thumbprint = $cert.Thumbprint

        Write-Log "✅ Certificate created (Thumbprint: $Thumbprint)"
        Write-Log "📄 Subject: $CertSubject"
        Write-Log "📂 Exported: $CertExportPfx, $CertExportPublic"
        Separator

        # ------------------------------------------------------
        # CREATE APP REGISTRATION
        # ------------------------------------------------------
        $AppDisplayName = "ExchangeOnlineAutomationApp"
        Write-Log "🧩 Creating App Registration: $AppDisplayName"

        $app = New-MgApplication -DisplayName $AppDisplayName
        $AppId = $app.AppId
        $ObjectId = $app.Id

        Write-Log "✅ App created. AppId: $AppId"

        # Upload certificate
        $keyCredential = [Microsoft.Graph.PowerShell.Models.MicrosoftGraphKeyCredential]@{
            Type = "AsymmetricX509Cert"
            Usage = "Verify"
            Key = $cert.RawData
            DisplayName = "AutomationCert"
            StartDateTime = $cert.NotBefore.ToUniversalTime()
            EndDateTime = $cert.NotAfter.ToUniversalTime()
        }

        Update-MgApplication -ApplicationId $ObjectId -KeyCredentials @($keyCredential)

        Write-Log "✅ Certificate uploaded to App Registration"

        # ------------------------------------------------------
        # SERVICE PRINCIPAL
        # ------------------------------------------------------
        Write-Log "🔧 Creating Service Principal..."
        $sp = Get-MgServicePrincipal -Filter "AppId eq '$AppId'"

        if (-not $sp) {
            $sp = New-MgServicePrincipal -AppId $AppId
            Start-Sleep 10
        }
        Write-Log "✅ Service Principal Ready"

        # ------------------------------------------------------
        # ASSIGN Exchange.ManageAsApp
        # ------------------------------------------------------
        Write-Log "🔐 Assigning Exchange.ManageAsApp..."

        $exchangeSP = Get-MgServicePrincipal -Filter "AppId eq '00000002-0000-0ff1-ce00-000000000000'"
        $exchangeRole = $exchangeSP.AppRoles | Where-Object { $_.Value -eq "Exchange.ManageAsApp" }

        $assignment = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id |
                      Where-Object { $_.AppRoleId -eq $exchangeRole.Id }

        if (-not $assignment) {
            New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $sp.Id `
                -PrincipalId $sp.Id `
                -ResourceId $exchangeSP.Id `
                -AppRoleId $exchangeRole.Id | Out-Null

            Write-Log "✅ Exchange.ManageAsApp assigned."
        }
        else {
            Write-Log "ℹ️ Exchange.ManageAsApp already assigned."
        }

        # ------------------------------------------------------
        # ASSIGN DIRECTORY ROLE: Exchange Administrator
        # ------------------------------------------------------
        Write-Log "🛠 Assigning Exchange Administrator role..."

        $role = Get-MgDirectoryRole -All | Where-Object { $_.DisplayName -eq "Exchange Administrator" }

        if (-not $role) {
            Write-Log "⚙️ Enabling directory role..."

            $template = Get-MgDirectoryRoleTemplate -All |
                        Where-Object { $_.DisplayName -eq "Exchange Administrator" }

            $body = @{ roleTemplateId = $template.Id }

            Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/directoryRoles" `
                -Body ($body | ConvertTo-Json)

            Start-Sleep 5
            $role = Get-MgDirectoryRole -All |
                     Where-Object { $_.DisplayName -eq "Exchange Administrator" }
        }

        $isMember = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All |
                    Where-Object { $_.Id -eq $sp.Id }

        if (-not $isMember) {
            New-MgDirectoryRoleMemberByRef `
                -DirectoryRoleId $role.Id `
                -OdataId "https://graph.microsoft.com/v1.0/directoryObjects/$($sp.Id)" | Out-Null

            Write-Log "✅ Exchange Administrator role assigned."
        }
        else {
            Write-Log "ℹ️ Role already assigned."
        }

        Separator

        # ------------------------------------------------------
        # OUTPUT CONNECTION VALUES
        # ------------------------------------------------------
        $Organization = (Get-MgOrganization).VerifiedDomains |
                        Where-Object { $_.IsDefault -eq $true } |
                        Select-Object -ExpandProperty Name

        Write-Log "=== CONNECTION VALUES ==="
        Write-Log "Tenant:        $Organization"
        Write-Log "App ID:        $AppId"
        Write-Log "Thumbprint:    $Thumbprint"
        Write-Log "Organization:  $Organization"

        Separator
        Write-Log "Connect-ExchangeOnline -AppId `"$AppId`" -CertificateThumbprint `"$Thumbprint`" -Organization `"$Organization`""
        Separator

        Write-Log "--------- IMPORTANT - SAVE THESE DETAILS ---------"
        Write-Log "Store the certificates & connection details in IT Glue"
        Write-Log "✅ Completed successfully!"

        $copyButton.IsEnabled = $true
    }
	catch {
	    Write-Log "❌ ERROR: $($_.Exception.Message)"
	}
	finally {

	    Write-Log "🔌 Disconnecting from Microsoft Graph and Exchange..."
	    try {
	        Disconnect-MgGraph | Out-Null
		Disconnect-ExchangeOnline -Confirm:$false | Out-Null
	        Write-Log "✅ Disconnected from Microsoft Graph and Exchange"
	    }
	    catch {
	        Write-Log "⚠️ Could not disconnect Graph: $($_.Exception.Message)"
	    }
    $startButton.IsEnabled = $true
	}
})

# ==============================================================
#   COPY BUTTON — RELIABLE CLIPBOARD SET
# ==============================================================

$copyButton.Add_Click({
    $text = $outputBox.Text -split "`r`n"

    $filtered = $text | Where-Object {
        $_ -match "Tenant:" -or
        $_ -match "App ID:" -or
        $_ -match "Thumbprint:" -or
        $_ -match "Organization:"
    }

    $finalText = ($filtered -join "`r`n")

    $maxRetries = 10
    for ($i = 1; $i -le $maxRetries; $i++) {
        try {
            [System.Windows.Clipboard]::SetText($finalText)
            break
        }
        catch {
            Start-Sleep -Milliseconds 80
        }
    }

    [System.Windows.MessageBox]::Show("Connection details copied.", "Copied")
})

# ==============================================================
#   EXIT BUTTON
# ==============================================================

$exitButton.Add_Click({ $window.Close() })

# ==============================================================
#   SHOW MAIN WINDOW
# ==============================================================

$window.ShowDialog() | Out-Null
