<#
Developed by: Rhys Saddul
#>

param (
    [string]$ExportPath
)

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

# ==============================================================
#   INITIALISE LOGGING & PATHS
# ==============================================================

$logsRoot  = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs"
$logsDir   = Join-Path $logsRoot "Export_M365"
$auditsDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Audits"

foreach ($dir in @($logsRoot, $logsDir, $auditsDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
    }
}

if (-not $ExportPath) {
    $ExportPath = Join-Path $auditsDir "M365_Export.csv"
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile   = Join-Path $logsDir "Export_M365_$timestamp.log"

function Add-ToLog {
    param([string]$Message, [switch]$Failed)

    $time  = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $entry = if ($Failed) { "[$time] ❌ $Message" } else { "[$time] $Message" }

    $entry
    try {
        Add-Content -Path $logFile -Value $entry -ErrorAction SilentlyContinue
    } catch {}
}

Add-ToLog "▶ Starting Microsoft 365 Groups export..."
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
#   EXPORT MICROSOFT 365 GROUPS
# ==============================================================

try {
    Add-ToLog "⏳ Retrieving Microsoft 365 groups..."
    $membersArray = @()

    $groups = Get-UnifiedGroup -ResultSize Unlimited
    foreach ($group in $groups) {
        $groupName = $group.DisplayName
        $groupEmail = $group.PrimarySmtpAddress
        Add-ToLog "📋 Processing group: $groupName <$groupEmail>"

        try {
            $members = Get-UnifiedGroupLinks -Identity $groupEmail -LinkType Members -ResultSize Unlimited -ErrorAction Stop
            foreach ($member in $members) {
                $membersArray += [PSCustomObject]@{
                    "Group Name"           = $groupName
                    "Group Email Address"  = $groupEmail
                    "Member Name"          = $member.DisplayName
                    "Member Email Address" = $member.PrimarySmtpAddress
                }
            }
        }
        catch {
            Add-ToLog "⚠️ Error retrieving members for group ${groupName}: $($_.Exception.Message)" -Failed
        }
    }

    # ==============================================================
    #   SAVE TO CSV
    # ==============================================================

    if ($membersArray.Count -eq 0) {
        Add-ToLog "⚠️ No members found in Microsoft 365 groups." -Failed
    }
    else {
        $membersArray | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8 -Force
        Add-ToLog "✅ Export completed successfully. File saved to: $ExportPath"
# ✅ Show popup confirming where the export was saved
[System.Windows.Forms.MessageBox]::Show(
    "Microsoft 365 Groups exported successfully to:`n$ExportPath",
    "Export Complete",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)
    }
}
catch {
    Add-ToLog "❌ Fatal script error: $($_.Exception.Message)" -Failed
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    Add-ToLog "🔒 Disconnected from Exchange Online."
    Add-ToLog "📁 Log saved to: $logFile"
}

Add-ToLog "🎉 Completed Microsoft 365 Groups Export process."
