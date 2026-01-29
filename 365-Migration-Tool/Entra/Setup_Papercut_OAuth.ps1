<#
Developed by: Rhys Saddul
#>

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Application.ReadWrite.All","Directory.ReadWrite.All" | Out-Null

Add-Type -AssemblyName PresentationFramework
# GUI Layout (XAML)
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="PaperCut OAuth App Registration"
        Height="550" Width="700"
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
<TextBlock Text="PaperCut OAuth App Registration Tool"
                   FontSize="16" FontWeight="Bold" Grid.Row="0" Margin="0,0,0,10"/>
<Button Name="StartButton" Content="Start Registration" Grid.Row="1"
                Height="30" Width="160" HorizontalAlignment="Left" Margin="0,0,0,10"/>
<TextBox Name="OutputBox" Grid.Row="2" VerticalScrollBarVisibility="Auto"
                 TextWrapping="Wrap" IsReadOnly="True" AcceptsReturn="True"
                 FontFamily="Consolas" Margin="0,0,0,10"/>
<StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
<Button Name="CopyButton" Content="Copy Output" Width="120" Margin="0,0,10,0" IsEnabled="False"/>
<Button Name="ExitButton" Content="Exit" Width="100"/>
</StackPanel>
</Grid>
</Window>
"@
# Load GUI
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
$startButton  = $window.FindName("StartButton")
$exitButton   = $window.FindName("ExitButton")
$copyButton   = $window.FindName("CopyButton")
$outputBox    = $window.FindName("OutputBox")
# Use [ref] to hold final values
$appId = [ref]"" 
$clientSecretVal = [ref]""
$tenantId = [ref]""
# Logging helpers
function Write-Log {
    param ([string]$text)
    $outputBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] $text`r`n")
    $outputBox.ScrollToEnd()
}
function Write-Separator {
    $outputBox.AppendText("--------------------------------------------------`r`n")
    $outputBox.ScrollToEnd()
}
# Start Registration
$startButton.Add_Click({
    $startButton.IsEnabled = $false
    $copyButton.IsEnabled = $false
    $outputBox.Clear()
    $appName = "PaperCut OAuth"
    $redirectUris = @("http://localhost:9191/azure-oauth2-callback")
    try {
        Write-Log "Starting app registration process..."
        Write-Separator
        Write-Log "Registering application: '$appName'"
        $app = New-MgApplication -DisplayName $appName -SignInAudience "AzureADMyOrg" `
            -Web @{ RedirectUris = $redirectUris }
        $appId.Value = $app.AppId
        Write-Log "App created with App ID: $($app.AppId)"
        Write-Log "Creating service principal..."
        $sp = New-MgServicePrincipal -AppId $app.AppId
        Write-Log "Service principal created."
        Write-Log "Assigning Microsoft Graph permissions..."
        $graphSp = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"
        $permissions = @(
            "Mail.Read.All", 
            "Mail.Send", 
            "Mail.Send.All", 
            "Mail.Send.Shared"
        )
        $resourceAccess = @()
        foreach ($perm in $permissions) {
            $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $perm }
            if ($role) {
                $resourceAccess += @{ Id = $role.Id; Type = "Role" }
                Write-Log " → Added: $perm [Role]"
            } else {
                $scope = $graphSp.Oauth2Permissions | Where-Object { $_.Value -eq $perm }
                if ($scope) {
                    $resourceAccess += @{ Id = $scope.Id; Type = "Scope" }
                    Write-Log " → Added: $perm [Scope]"
                }
            }
        }
        Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess @(@{
            ResourceAppId = $graphSp.AppId
            ResourceAccess = $resourceAccess
        })
        Write-Log "Permissions assigned."
        Write-Log "Waiting 10 seconds for propagation..."
        Start-Sleep -Seconds 10
        $tenantId.Value = (Get-MgContext).TenantId
        $consentUrl = "https://login.microsoftonline.com/$($tenantId.Value)/adminconsent?client_id=$($appId.Value)"
        Write-Log "Opening admin consent URL:"
        Write-Log " → $consentUrl"
        Start-Process "msedge.exe" -ArgumentList "--inprivate", $consentUrl
        Write-Log "Creating client secret..."
        $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential @{
            DisplayName = "PaperCut Sync Secret"
            EndDateTime = (Get-Date).AddYears(2)
        }
        $clientSecretVal.Value = $secret.SecretText
        Write-Log "Client secret created."
        Write-Separator
        Write-Log "FINAL DETAILS:"
        Write-Log "App ID:        $($appId.Value)"
        Write-Log "Client Secret: $($clientSecretVal.Value)"
        Write-Log "Tenant ID:     $($tenantId.Value)"
        Write-Separator
        Write-Log "✅ Registration complete. You can now copy the credentials."
        $copyButton.IsEnabled = $true
    }
    catch {
        Write-Log "❌ ERROR: $_"
    }
    finally {
        $startButton.IsEnabled = $true
    }
})
# Copy final credentials to clipboard
$copyButton.Add_Click({
    $finalOutput = "App ID:        $($appId.Value)`r`nClient Secret: $($clientSecretVal.Value)`r`nTenant ID:     $($tenantId.Value)"
    [System.Windows.Clipboard]::SetText($finalOutput)
    [System.Windows.MessageBox]::Show("Final details copied to clipboard.", "Copied")
})
# Exit
$exitButton.Add_Click({
    $window.Close()
})
# Show GUI
$window.ShowDialog() | Out-Null

# Disconnect from Graph when the GUI is closed
Disconnect-MgGraph | Out-Null