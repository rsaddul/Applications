<#
Developed by: Rhys Saddul
#>

param (
    [string]$ExportPath = $(Join-Path (Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Audits") "Teams_Export.csv")
)

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

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
#   CONFIRM EXPORT PATH
# ==============================================================

# --- Confirm Export Path ---
if (-not $ExportPath) {
    $ExportPath = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Enter export path for Teams CSV:", 
        "Teams Export", 
        (Join-Path ([Environment]::GetFolderPath("Desktop")) "Teams_Export.csv")
    )
    if ([string]::IsNullOrWhiteSpace($ExportPath)) {
        "❌ Cancelled: ExportPath is required."
        exit 1
    }
}

# --- Run Export ---
try {
    "▶ Starting Microsoft Teams export..."
    "   Export path : $ExportPath"

    # Connect to Teams
    "🔑 Connecting to Microsoft Teams..."
    Connect-MicrosoftTeams | Out-Null
    "✔ Connected to Microsoft Teams."

    # Collect Teams
    "⏳ Retrieving Teams..."
    $teams = Get-Team -ErrorAction Stop
    $teamsInfo = @()

    foreach ($team in $teams) {
        $teamPrivacy = $team.Visibility
        $groupId     = $team.GroupId
        $teamName    = $team.DisplayName

        "📌 Processing Team: $teamName"

        # Team members
        $teamMembers = @()
        try {
            $teamMembers = Get-TeamUser -GroupId ${groupId} | Select-Object -ExpandProperty User
        }
        catch {
            "⚠️ Failed to get members for team ${teamName}: $($_.Exception.Message)"
        }

        # Channels
        $channels = @()
        try {
            $channels = Get-TeamChannel -GroupId $groupId
        }
        catch {
            "⚠️ Failed to get channels for team ${teamName}: $($_.Exception.Message)"
        }

        foreach ($channel in $channels) {
            $channelName = $channel.DisplayName
            $channelMembers = @()

            try {
                $channelMembers = Get-TeamChannelUser -GroupId $groupId -DisplayName $channel.DisplayName | Select-Object -ExpandProperty User
            }
            catch {
                "⚠️ Could not fetch members for channel $channelName in $teamName"
            }

            $teamsInfo += [PSCustomObject]@{
                "Team Name"       = $teamName
                "Privacy"         = $teamPrivacy
                "GroupId"         = $groupId
                "Team Members"    = $teamMembers -join ", "
                "Channel Name"    = $channelName
                "Channel Members" = $channelMembers -join ", "
            }
        }
    }

    if ($teamsInfo.Count -eq 0) {
        "⚠️ No Teams data collected."
    }
    else {
        $teamsInfo | Export-Csv -Path $ExportPath -Encoding UTF8 -NoTypeInformation -Force
        "✅ Export complete. File saved to: $ExportPath"
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
    "❌ Fatal script error: $($_.Exception.Message)"
}
finally {
    Disconnect-MicrosoftTeams | Out-Null
    "✔ Disconnected from Microsoft Teams."
}
