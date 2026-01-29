<#
Developed by: Rhys Saddul
#>

param (
    [string]$ExportPath = $(Join-Path (Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Audits") "DL_Export.csv")
)

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

# ==============================================================
#   INITIALISE LOGGING
# ==============================================================

$logsRoot = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs"
$logDir   = Join-Path $logsRoot "Export_DistributionLists"

foreach ($dir in @($logsRoot, $logDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
    }
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile   = Join-Path $logDir "Export_DistributionLists_$timestamp.log"

function Add-ToLog {
    param([string]$Message, [switch]$Failed)

    $time  = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
    $entry = if ($Failed) { "[$time] ❌ $Message" } else { "[$time] $Message" }

    # ✅ Output to GUI (standard output)
    $entry

    # Optional: still save to log file
    try {
        Add-Content -Path $logFile -Value $entry -ErrorAction SilentlyContinue
    } catch {}
}


# ==============================================================
#   CONFIRM EXPORT PATH
# ==============================================================

if (-not $ExportPath) {
    $ExportPath = Join-Path $auditsDir "DL_Export.csv"
}

try {
    $null = New-Item -Path $ExportPath -ItemType File -Force -ErrorAction Stop
    Remove-Item $ExportPath -Force
}
catch {
    Add-ToLog "❌ Cannot write to export path: $ExportPath" -Failed
    exit 1
}

Add-ToLog "▶ Starting Distribution List export..."
Add-ToLog "📁 Export path: $ExportPath"

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
#   EXPORT DISTRIBUTION LISTS AND MEMBERS
# ==============================================================

try {
    Add-ToLog "⏳ Retrieving distribution lists..."

    $membersArray = @()
    $distributionLists = Get-DistributionGroup -ResultSize Unlimited |
        Where-Object { $_.RecipientTypeDetails -eq "MailUniversalDistributionGroup" }

    foreach ($dl in $distributionLists) {
        $dlName  = $dl.DisplayName
        $dlEmail = $dl.PrimarySmtpAddress
        Add-ToLog "📋 Processing DL: $dlName <$dlEmail>"

        try {
            $members = Get-DistributionGroupMember -Identity $dlEmail -ResultSize Unlimited -ErrorAction Stop
            foreach ($member in $members) {
                $membersArray += [PSCustomObject]@{
                    "Distribution List Name"         = $dlName
                    "Distribution List EmailAddress" = $dlEmail
                    "Member Name"                    = $member.DisplayName
                    "Member Email Address"           = $member.PrimarySmtpAddress
                }
            }
        }
        catch {
            Add-ToLog "⚠️ Failed to get members for $dlName : $($_.Exception.Message)" -Failed
        }
    }

    # ==============================================================
    #   SAVE TO CSV
    # ==============================================================

    $membersArray | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8 -Force
    Add-ToLog "✅ Export completed successfully. File saved to: $ExportPath"
# ✅ Show popup confirming where the export was saved
[System.Windows.Forms.MessageBox]::Show(
    "Distribution Lists exported successfully to:`n$ExportPath",
    "Export Complete",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)

}
catch {
    Add-ToLog "❌ Fatal script error: $($_.Exception.Message)" -Failed
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    Add-ToLog "🔒 Disconnected from Exchange Online."
    Add-ToLog "📁 Log saved to: $logFile"
}

Add-ToLog "🎉 Completed Distribution List Export process."