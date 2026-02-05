<#
Developed by: Rhys Saddul
#>

param (
   [string]$ExportPath
)

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

# ==============================================================
#   INITIALISE LOGGING
# ==============================================================

$logsRoot = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs"
$logsDir  = Join-Path $logsRoot "Export_Contacts"
$auditsDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Audits"

foreach ($dir in @($logsRoot, $logsDir, $auditsDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
    }
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile   = Join-Path $logsDir "Export_Contacts_$timestamp.log"

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
#   CONFIRM EXPORT PATH
# ==============================================================

if (-not $ExportPath) {
    $ExportPath = Join-Path $auditsDir "Contacts_Export.csv"
}

Add-ToLog "Starting Contacts export..."
Add-ToLog "Export path: $ExportPath"

# ==============================================================
#   CONNECT TO EXCHANGE ONLINE
# ==============================================================

try {
    Add-ToLog "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowProgress $true -ShowBanner:$false
    Add-ToLog "Connected to Exchange Online."
}
catch {
    Add-ToLog "Failed to connect to Exchange Online: $($_.Exception.Message)" -Failed
    [System.Windows.Forms.MessageBox]::Show(
        "Connection failed:`n$($_.Exception.Message)",
        "Exchange Online Connection Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

# ==============================================================
#   EXPORT CONTACTS
# ==============================================================

try {
    Add-ToLog "Exporting Mail Contacts..."

    Get-Contact -RecipientTypeDetails MailContact |
        Select-Object DisplayName, FirstName, LastName, WindowsEmailAddress, RecipientType, DistinguishedName, IsDirSynced |
        Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8 -Force

    Add-ToLog "Contacts exported successfully to: $ExportPath"
# ✅ Show popup confirming where the export was saved
[System.Windows.Forms.MessageBox]::Show(
    "Contacts exported successfully to:`n$ExportPath",
    "Export Complete",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)
}
catch {
    Add-ToLog "Error exporting contacts: $($_.Exception.Message)" -Failed
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    Add-ToLog "Disconnected from Exchange Online."
    Add-ToLog "Log saved to: $logFile"
}

Add-ToLog "Completed Contacts Export process."