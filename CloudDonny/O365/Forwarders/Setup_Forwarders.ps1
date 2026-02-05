<#
Developed by: Rhys Saddul
#>

# ==============================================================
#   ENSURE SCRIPT RUNS IN STA MODE (for GUI + ExchangeOnline)
# ==============================================================

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Write-Host "[STA Check] Relaunching script in STA mode..."
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell.exe"
    $psi.Arguments = "-STA -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Definition)`""
    $psi.UseShellExecute = $false
    [System.Diagnostics.Process]::Start($psi) | Out-Null
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

# ==============================================================
#   INITIALISE LOGGING
# ==============================================================

$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$timestamp  = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
$logFile = Join-Path $logDir "Export_Contacts_$timestamp.log"

function Add-ToLog {
    param([string]$Message, [switch]$Failed)

    $time = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
    $entry = "[$time] $Message"

    # ✅ Send to GUI
    $entry

    # Optional: log to file
    try {
        Add-Content -Path $logFile -Value $entry -ErrorAction SilentlyContinue
    } catch {}
}


# ==============================================================
#   ASK TO GENERATE CSV TEMPLATE
# ==============================================================

$answer = [System.Windows.Forms.MessageBox]::Show(
    "Would you like to generate a CSV template with example data?",
    "Generate Template",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
    try {
        $templatePath = [System.IO.Path]::Combine(
            [Environment]::GetFolderPath("Desktop"),
            "Setup_Forwarde_Template.csv"
        )

        @"
UserPrincipalName,ForwardingSmtpAddress
rsaddul@thecloudschool.co.uk,rsaddul@eduthing-azure.co.uk
jskett@thecloudschool.co.uk,jskett@eduthing-azure.co.uk
rahemd@thecloudschool.co.uk,rahemd@eduthing-azure.co.uk
"@ | Out-File -FilePath $templatePath -Encoding UTF8 -Force

        [System.Windows.Forms.MessageBox]::Show(
            "Template generated successfully:`n$templatePath",
            "Template Created",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        Add-ToLog "✅ Generated CSV template at: $templatePath"
        Add-ToLog "💡 You can now edit the CSV to update forwarding addresses."
        exit 0
    }
    catch {
        Add-ToLog "⚠️ Failed to create template: $($_.Exception.Message)" -Failed
        exit 1
    }
}

# ==============================================================
#   SELECT CSV FILE
# ==============================================================

$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Select your CSV file"
$OpenFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
$OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $CSVPath = $OpenFileDialog.FileName
    Add-ToLog "📄 Using CSV file: $CSVPath"
} else {
    Add-ToLog "❌ Cancelled: CSV file is required." -Failed
    exit 1
}

if (-not (Test-Path $CSVPath)) {
    Add-ToLog "❌ CSV file not found: $CSVPath" -Failed
    exit 1
}

# ==============================================================
#   CONNECT TO EXCHANGE ONLINE
# ==============================================================

try {
    Add-ToLog "🌐 Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowProgress $true -ShowBanner:$false
    Add-ToLog "✅ Connected to Exchange Online."
}
catch {
    Add-ToLog "❌ Failed to connect to Exchange Online: $($_.Exception.Message)" -Failed
    [System.Windows.Forms.MessageBox]::Show(
        "Connection failed:`n$($_.Exception.Message)",
        "Exchange Online Connection Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

# ==============================================================
#   PROCESS MAILBOX FORWARDING UPDATES
# ==============================================================

try {
    $rows = Import-Csv -Path $CSVPath
    if (-not $rows) {
        Add-ToLog "⚠️ No rows found in CSV." -Failed
        exit 1
    }

    $successCount = 0
    $failCount = 0
    $total = $rows.Count
    $counter = 0

    foreach ($row in $rows) {
        $counter++
        $user = $row.UserPrincipalName
        $forward = $row.ForwardingSmtpAddress

        if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($forward)) {
            Add-ToLog "⚠️ Skipping invalid row (missing UserPrincipalName or ForwardingSmtpAddress)." -Failed
            continue
        }

        try {
            Add-ToLog "🔧 [$counter/$total] Setting forwarding for $user → $forward..."
            Set-Mailbox -Identity $user -ForwardingSmtpAddress $forward -DeliverToMailboxAndForward $false -ErrorAction Stop
            Add-ToLog "✅ Forwarding updated for $user → $forward"
            $successCount++
        }
        catch {
            if ($_.Exception.Message -like "*not found*") {
                Add-ToLog "❌ Mailbox not found: $user" -Failed
            } else {
                Add-ToLog "❌ Failed to update ${user}: $($_.Exception.Message)" -Failed
            }
            $failCount++
        }
    }

    Add-ToLog "🔹 Summary: $successCount success(es), $failCount failure(s)."
}
catch {
    Add-ToLog "❌ Fatal error: $($_.Exception.Message)" -Failed
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    Add-ToLog "🔒 Disconnected from Exchange Online."
    Add-ToLog "📁 Logs saved to:`n   $successLog`n   $failLog"
}

Add-ToLog "🎉 Completed Mailbox Forwarding Update process."