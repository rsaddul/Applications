<#
Developed by: Rhys Saddul
#>

param (
    [string]$ExportPath
)

$logsRoot  = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs"
$logsDir   = Join-Path $logsRoot "Export_MailEnabledSecurityGroups"
$auditsDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Audits"

foreach ($dir in @($logsRoot, $logsDir, $auditsDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
    }
}

if (-not $ExportPath) {
    $ExportPath = Join-Path $auditsDir "Mail-Enabled_Security_Groups_Export.csv"
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile   = Join-Path $logsDir "Export_MailEnabledSecurityGroups_$timestamp.log"

function Add-ToLog {
    param([string]$Message, [switch]$Failed)

    $time  = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $entry = if ($Failed) { "[$time] ❌ $Message" } else { "[$time] $Message" }

    $entry
    try {
        Add-Content -Path $logFile -Value $entry -ErrorAction SilentlyContinue
    } catch {}
}

# --- Run Export ---
try {
    Add-ToLog "▶ Starting Mail-Enabled Security Groups export..."
    Add-ToLog "📁 Export path : $ExportPath"

    # Connect to Graph
    Add-ToLog "🔑 Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes "Group.Read.All","User.Read.All" -NoWelcome | Out-Null
    Add-ToLog "✔ Connected to Microsoft Graph."

    # Collect MESGs
    Add-ToLog "⏳ Retrieving Mail-Enabled Security Groups..."
    $membersArray = @()
    $mesgs = Get-MgGroup -All -Filter "securityEnabled eq true and mailEnabled eq true"

    foreach ($group in $mesgs) {
        $groupName         = $group.DisplayName
        $groupEmailAddress = $group.Mail

        try {
            $members = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop
            foreach ($member in $members) {
                $membersArray += [PSCustomObject]@{
                    "Group Name"           = $groupName
                    "Group Email Address"  = $groupEmailAddress
                    "Member Name"          = $member.AdditionalProperties.displayName
                    "Member Email Address" = $member.AdditionalProperties.mail
                }
            }
            Add-ToLog "✔ Processed group: $groupName ($groupEmailAddress)"
        }
        catch {
            Add-ToLog "⚠️ Error retrieving members for group: $groupName - $($_.Exception.Message)" -Failed
        }
    }

    if ($membersArray.Count -eq 0) {
        Add-ToLog "⚠️ No members found in any Mail-Enabled Security Groups."
    }
    else {
        $membersArray |
            Export-Csv -Path $ExportPath -Encoding UTF8 -NoTypeInformation -Force

        Add-ToLog "✅ Export complete. File saved to: $ExportPath"

        # ✅ Show popup confirming where the export was saved
        [System.Windows.Forms.MessageBox]::Show(
            "Mail-Enabled Security Groups exported successfully to:`n$ExportPath",
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
    Disconnect-MgGraph | Out-Null
    Add-ToLog "✔ Disconnected from Microsoft Graph."
}
