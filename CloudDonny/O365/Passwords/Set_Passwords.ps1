<#
Developed by: Rhys Saddul
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# ==============================================================
#   INITIALISE LOGGING
# ==============================================================

$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs\Set_Passwords"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$timestamp   = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
$successLog  = Join-Path $logDir "Password_Reset_Success_$timestamp.log"
$failLog     = Join-Path $logDir "Password_Reset_Failed_$timestamp.log"

function Add-ToLog {
    param([System.Windows.Forms.TextBox]$Box, [string]$Message, [switch]$Failed)
    $time  = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
    $entry = "[$time] $Message"

    if ($Box) {
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
#   GUI SETUP
# ==============================================================

$form               = New-Object System.Windows.Forms.Form
$form.Text          = "Reset User Passwords from CSV"
$form.Size          = New-Object System.Drawing.Size(800,520)
$form.StartPosition = "CenterScreen"
$form.BackColor     = [System.Drawing.Color]::WhiteSmoke

$layout = New-Object System.Windows.Forms.TableLayoutPanel
$layout.Dock = "Fill"
$layout.RowCount = 3
$layout.ColumnCount = 1
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)))
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

# --- Blue Button Helper ---
function New-BlueBtn($text,[ScriptBlock]$onClick) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text
    $b.Width = 210
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

# ==============================================================
#   GLOBALS
# ==============================================================

$selectedCSV  = $null
$forceChange  = $true

# ==============================================================
#   BUTTON PANEL
# ==============================================================

$btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$btnPanel.FlowDirection = "LeftToRight"
$btnPanel.Dock = "Top"
$btnPanel.AutoSize = $true

# ==============================================================
#   BUTTONS
# ==============================================================

# --- Generate Template ---
$btnTemplate = New-BlueBtn "📄 Generate Template" {
    try {
        $templatePath = [System.IO.Path]::Combine(
            [Environment]::GetFolderPath("Desktop"),
            "PasswordResetTemplate.csv"
        )

        @"
UserPrincipalName,Password
rsaddul@thecloudschool.co.uk,P@ssword123
jskett@thecloudschool.co.uk,P@ssword456
rahemd@thecloudschool.co.uk,P@ssword789
"@ | Out-File -FilePath $templatePath -Encoding UTF8 -Force

        [System.Windows.Forms.MessageBox]::Show(
            "Template created successfully:`n$templatePath",
            "Template Created",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        Add-ToLog $outputBox "✅ Generated CSV template at: $templatePath"
# ✅ Show popup confirming where the template was saved
[System.Windows.Forms.MessageBox]::Show(
    "Template generated successfully:`n$templatePath",
    "Template Created",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)
    }
    catch {
        Add-ToLog $outputBox "⚠️ Failed to create template: $($_.Exception.Message)" -Failed
    }
}

# --- Select CSV ---
$btnUpload = New-BlueBtn "📂 Select CSV" {
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "CSV Files (*.csv)|*.csv"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $global:selectedCSV = $ofd.FileName
        Add-ToLog $outputBox "✔ Selected CSV: $selectedCSV"
    }
}

# --- Force Change Toggle ---
$btnToggle = New-BlueBtn "🔄 Force Change: ON" {
    $global:forceChange = -not $global:forceChange
    if ($global:forceChange) {
        $this.Text = "🔄 Force Change: ON"
        Add-ToLog $outputBox "✔ Users will be forced to change password on next login."
    } else {
        $this.Text = "🔄 Force Change: OFF"
        Add-ToLog $outputBox "✔ Users will NOT be forced to change password on next login."
    }
}

# --- Run Password Reset ---
$btnRun = New-BlueBtn "▶ Run Password Reset" {
    if (-not $selectedCSV) { Add-ToLog $outputBox "❌ Please select a CSV first." -Failed; return }

    try {
        Add-ToLog $outputBox "🌐 Connecting to Microsoft Graph..."
        Connect-MgGraph -Scopes "User.ReadWrite.All" -ErrorAction Stop | Out-Null
        Add-ToLog $outputBox "✅ Connected to Microsoft Graph."
    }
    catch {
        Add-ToLog $outputBox "❌ Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Failed
        return
    }

    try {
        $users = Import-Csv -Path $selectedCSV
        if (-not $users) { Add-ToLog $outputBox "⚠️ No users found in CSV." -Failed; return }

        $successCount = 0
        $failCount = 0
        $total = $users.Count
        $counter = 0

        foreach ($user in $users) {
            $counter++
            $upn = $user.UserPrincipalName
            $pwd = $user.Password

            if ([string]::IsNullOrWhiteSpace($upn) -or [string]::IsNullOrWhiteSpace($pwd)) {
                Add-ToLog $outputBox "⚠️ [$counter/$total] Skipping invalid row (missing UPN or Password)." -Failed
                continue
            }

            try {
                Add-ToLog $outputBox "🔧 [$counter/$total] Resetting password for $upn..."
                Update-MgUser -UserId $upn -PasswordProfile @{
                    ForceChangePasswordNextSignIn = $forceChange
                    Password = $pwd
                } -ErrorAction Stop

                Add-ToLog $outputBox "✅ Password reset for $upn"
                $successCount++
            }
            catch {
                Add-ToLog $outputBox "❌ Failed to reset $upn: $($_.Exception.Message)" -Failed
                $failCount++
            }
        }

        Add-ToLog $outputBox "🔹 Summary: $successCount success(es), $failCount failure(s)."
        Add-ToLog $outputBox "📁 Logs saved to:`n   $successLog`n   $failLog"
    }
    finally {
        Disconnect-MgGraph | Out-Null
        Add-ToLog $outputBox "🔒 Disconnected from Microsoft Graph."
    }
}

# --- Close Button ---
$btnClose = New-BlueBtn "❌ Close" { $form.Close() }
$btnClose.Dock = "Fill"

# ==============================================================
#   ADD TO LAYOUT
# ==============================================================

$btnPanel.Controls.AddRange(@($btnTemplate,$btnUpload,$btnToggle,$btnRun))
$layout.Controls.Add($btnPanel,0,0)
$layout.Controls.Add($outputBox,0,1)
$layout.Controls.Add($btnClose,0,2)

# ==============================================================
#   SHOW GUI
# ==============================================================

[void]$form.ShowDialog()
