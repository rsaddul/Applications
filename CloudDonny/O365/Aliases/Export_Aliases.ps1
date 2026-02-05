<#
Developed by: Rhys Saddul
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ==============================================================
#   INITIALISE LOGGING
# ==============================================================

$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs\Export_Aliases"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$timestamp = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
$successLog = Join-Path $logDir "Export_Aliases_Success_$timestamp.log"
$failLog    = Join-Path $logDir "Export_Aliases_Failed_$timestamp.log"

function Add-ToLog {
    param(
        [System.Windows.Forms.TextBox]$Box,
        [string]$Message,
        [switch]$Failed
    )

    $time  = (Get-Date).ToString("HH:mm:ss")
    $entry = "[$time] $Message"

    if ($Box -ne $null) {
        $Box.AppendText("$entry`r`n")
        $Box.ScrollToCaret()
    }

    try {
        if ($Failed) {
            Add-Content -Path $failLog -Value $entry -ErrorAction SilentlyContinue
        } else {
            Add-Content -Path $successLog -Value $entry -ErrorAction SilentlyContinue
        }
    } catch {}
}

# ==============================================================
#   MAIN SCRIPT FUNCTION
# ==============================================================

function Run-AliasExport {
    param(
        [System.Windows.Forms.TextBox]$OutputBox
    )

    try {
        # Ask for App (Client) ID
        $AppId = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Enter the App (Client) ID:",
            "Exchange Online App Connection",
            ""
        )
        if ([string]::IsNullOrWhiteSpace($AppId)) { throw "App ID is required." }

        # Ask for Certificate Thumbprint
        $Thumbprint = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Enter the Certificate Thumbprint:",
            "Exchange Online App Connection",
            ""
        )
        if ([string]::IsNullOrWhiteSpace($Thumbprint)) { throw "Certificate Thumbprint is required. Run Setup ExchangeOnlineAutomation in Entra area if you don't have this setup." }

        # Ask for Tenant Domain (e.g., thecloudschool.onmicrosoft.com)
        $Organization = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Enter the Tenant Domain (example: thecloudschool.onmicrosoft.com):",
            "Exchange Online App Connection",
            ""
        )
        if ([string]::IsNullOrWhiteSpace($Organization)) { throw "Tenant domain is required." }

        Add-ToLog $OutputBox "Connecting to Exchange Online..."
        Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $Thumbprint -Organization $Organization -ShowBanner:$false
        Add-ToLog $OutputBox "Connected to Exchange Online."
    }
    catch {
        Add-ToLog $OutputBox "Failed to connect to Exchange Online: $($_.Exception.Message)" -Failed
        [System.Windows.Forms.MessageBox]::Show(
            "Connection failed:`n$($_.Exception.Message)",
            "Exchange Online Connection Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    # --- Ask user for CSV input/output paths ---
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = "Select Input CSV (List of UPNs)"
    $ofd.Filter = "CSV Files (*.csv)|*.csv"
    if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Add-ToLog $OutputBox "Operation cancelled by user." -Failed
        return
    }

    $inputCsvPath = $ofd.FileName
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Title = "Save Output CSV As"
    $sfd.Filter = "CSV Files (*.csv)|*.csv"
    $sfd.FileName = "Alias_Export_Output.csv"
    if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Add-ToLog $OutputBox "Output file not selected." -Failed
        return
    }

    $outputCsvPath = $sfd.FileName

    Add-ToLog $OutputBox "Input: $inputCsvPath"
    Add-ToLog $OutputBox "Output: $outputCsvPath"

    # --- Read CSV ---
    $userList = Import-Csv -Path $inputCsvPath -Header "UserPrincipalName"
    if (-not $userList) {
        Add-ToLog $OutputBox "No data found in input CSV." -Failed
        return
    }

    $results = @()
    $total = $userList.Count
    $counter = 0

    foreach ($user in $userList) {
        $counter++
        $upn = $user.UserPrincipalName

        try {
            Add-ToLog $OutputBox "[$counter/$total] Processing: $upn"

            $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop

            $capitalSMTPs = ($mailbox.EmailAddresses | Where-Object { $_ -cmatch '^SMTP:' }) -join ";"
            $lowerSMTPs   = ($mailbox.EmailAddresses | Where-Object { $_ -cmatch '^smtp:' }) -join ";"

            $results += [PSCustomObject]@{
                DisplayName           = $mailbox.DisplayName
                UserPrincipalName     = $mailbox.UserPrincipalName
                CapitalSMTPAliases    = $capitalSMTPs
                LowercaseSMTPAliases  = $lowerSMTPs
            }

            Add-ToLog $OutputBox "Retrieved aliases for $upn"
        }
        catch {
            Add-ToLog -Box $OutputBox -Message "Failed to retrieve ${upn}: $(${_.Exception.Message})" -Failed
        }
    }

    $results | Export-Csv -Path $outputCsvPath -NoTypeInformation
    Add-ToLog $OutputBox "Export completed successfully to:`n$outputCsvPath"

    }
    catch {
        Add-ToLog $OutputBox "Fatal error: $($_.Exception.Message)" -Failed
    }
    finally {
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
        Add-ToLog $OutputBox "Disconnected from Exchange Online."
        Add-ToLog $OutputBox "Logs saved to:`n   $successLog`n   $failLog"
    }
}

# ==============================================================
#   GUI INTERFACE
# ==============================================================

$form = New-Object System.Windows.Forms.Form
$form.Text = "Export User Aliases from Exchange Online"
$form.Size = New-Object System.Drawing.Size(800,500)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::WhiteSmoke

$layout = New-Object System.Windows.Forms.TableLayoutPanel
$layout.Dock = "Fill"
$layout.RowCount = 2
$layout.ColumnCount = 1
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,85)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$form.Controls.Add($layout)

$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.ReadOnly = $true
$outputBox.Dock = "Fill"
$outputBox.BackColor = [System.Drawing.Color]::Black
$outputBox.ForeColor = [System.Drawing.Color]::Lime
$outputBox.Font = New-Object System.Drawing.Font("Consolas",10)
$layout.Controls.Add($outputBox,0,0)

$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.Dock = "Bottom"
$buttonPanel.FlowDirection = "LeftToRight"
$buttonPanel.AutoSize = $true

function New-BlueBtn($text,[ScriptBlock]$onClick) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text
    $b.Width = 200
    $b.Height = 40
    $b.BackColor = [System.Drawing.Color]::CornflowerBlue
    $b.ForeColor = [System.Drawing.Color]::White
    $b.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
    $b.FlatStyle = 'Flat'
    $b.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::RoyalBlue })
    $b.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::CornflowerBlue })
    $b.Add_Click($onClick)
    return $b
}

$btnRun = New-BlueBtn "Export Aliases" { Run-AliasExport -OutputBox $outputBox }
$btnSave = New-BlueBtn "Save Log" {
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "Text Files (*.txt)|*.txt"
    $dlg.FileName = "Alias_Export_Log.txt"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $outputBox.Lines | Out-File -FilePath $dlg.FileName -Encoding UTF8
        Add-ToLog $outputBox "Log saved to $($dlg.FileName)"
    }
}
$btnClose = New-BlueBtn "Close" { $form.Close() }

$buttonPanel.Controls.AddRange(@($btnRun,$btnSave,$btnClose))
$layout.Controls.Add($buttonPanel,0,1)

[void]$form.ShowDialog()
