<#
Developed by: Rhys Saddul
#>

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

# ==============================================================
#   INITIALISE LOGGING
# ==============================================================

$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs\Disable_Accounts"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$timestamp  = (Get-Date).ToString("dd-MM-yyyy_HH-mm-ss")
$successLog = Join-Path $logDir "Disable_Accounts_Success_$timestamp.log"
$failLog    = Join-Path $logDir "Disable_Accounts_Failed_$timestamp.log"

# --- Inline log helper ---
function Add-ToLog {
    param([string]$Message, [switch]$Failed)

    $time  = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
    $entry = "[$time] $Message"

    # Output to GUI launcher
    $entry

    # Save to logs
    try {
        if ($Failed) {
            Add-Content -Path $failLog -Value $entry
        } else {
            Add-Content -Path $successLog -Value $entry
        }
    } catch {}
}

# ==============================================================
#   ASK TO GENERATE TEMPLATE
# ==============================================================

$answer = [System.Windows.Forms.MessageBox]::Show(
    "Would you like to generate a CSV template with example data?",
    "Generate Template",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
    try {
        $templatePath = Join-Path ([Environment]::GetFolderPath("Desktop")) "Disable_Accounts_Template.csv"

@"
Email
rsaddul@thecloudschool.co.uk
jskett@thecloudschool.co.uk
rahemd@thecloudschool.co.uk
"@ | Out-File $templatePath -Encoding UTF8 -Force

        [System.Windows.Forms.MessageBox]::Show(
            "Template created successfully:`n$templatePath",
            "Template Created",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        Add-ToLog "✅ Generated CSV template at: $templatePath"
        exit 0
    }
    catch {
        Add-ToLog "⚠️ Failed to create template: $($_.Exception.Message)" -Failed
        exit 1
    }
}

# ==============================================================
#   FILE PICKER
# ==============================================================

$ofd = New-Object System.Windows.Forms.OpenFileDialog
$ofd.Title = "Select your CSV file"
$ofd.Filter = "CSV files (*.csv)|*.csv"
$ofd.InitialDirectory = [Environment]::GetFolderPath("Desktop")

if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Add-ToLog "❌ CSV selection cancelled." -Failed
    exit 1
}

$CSVPath = $ofd.FileName
Add-ToLog "📄 Using CSV file: $CSVPath"

# ==============================================================
#   REVOKE SESSIONS CHOICE
# ==============================================================

$revokePrompt = [System.Windows.Forms.MessageBox]::Show(
    "Do you want to revoke user sessions after disabling accounts?",
    "Revoke Sessions?",
    "YesNo",
    "Question"
)

$RevokeSessions = ($revokePrompt -eq [System.Windows.Forms.DialogResult]::Yes)

if ($RevokeSessions) {
    Add-ToLog "✔ Will revoke sessions after disabling accounts."
} else {
    Add-ToLog "ℹ️ Will NOT revoke sessions."
}

# ==============================================================
#   CONNECT TO GRAPH (using your working method)
# ==============================================================

try {
    Add-ToLog "🌐 Connecting to Microsoft Graph..."

    $scopes = @("User.ReadWrite.All")
    if ($RevokeSessions) { $scopes += "User.RevokeSessions.All" }

    Connect-MgGraph -Scopes $scopes -NoWelcome | Out-Null

    Add-ToLog "✅ Connected to Microsoft Graph."
}
catch {
    Add-ToLog "❌ Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Failed
    exit 1
}

# ==============================================================
#   PROCESS CSV USING WORKING METHOD EXACTLY
# ==============================================================

try {
    $rows = Import-Csv $CSVPath
    if (-not $rows) {
        Add-ToLog "⚠️ CSV contains no rows." -Failed
        exit 1
    }

    $success = 0
    $fails   = 0
    $total   = $rows.Count
    $count   = 0

    foreach ($row in $rows) {
        $count++
        $email = $row.Email.Trim()

        if ([string]::IsNullOrWhiteSpace($email)) {
            Add-ToLog "⚠️ [$count/$total] Skipping empty email row." -Failed
            continue
        }

        Add-ToLog "🔎 [$count/$total] Processing: $email..."

        try {
            # SAME METHOD YOU SAID WORKS
            $user = Get-MgUser -UserId $email -ErrorAction Stop

            if ($user.AccountEnabled -eq $false) {
                Add-ToLog "⚠️ $email is already blocked from signing in."
            }
            else {
                Update-MgUser -UserId $user.Id -AccountEnabled:$false -ErrorAction Stop
                Add-ToLog "✅ Blocked sign-in for: $email"
                $success++
            }

            if ($RevokeSessions) {
                Revoke-MgUserSignInSession -UserId $user.Id -ErrorAction Stop
                Add-ToLog "↪ Revoked sessions for $email"
            }
        }
        catch {
            Add-ToLog "❌ Failed for $email : $($_.Exception.Message)" -Failed
            $fails++
        }
    }

    Add-ToLog "🔹 Summary: $success success(es), $fails failure(s)."
    Add-ToLog "📁 Logs saved to:`n   $successLog`n   $failLog"
}
finally {
    Disconnect-MgGraph | Out-Null
    Add-ToLog "🔒 Disconnected from Microsoft Graph."
}

Add-ToLog "🎉 Completed Disable Accounts process."