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
$logsDir   = Join-Path $logsRoot "Export_Guests"
$auditsDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Audits"

foreach ($dir in @($logsRoot, $logsDir, $auditsDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
    }
}

if (-not $ExportPath) {
    $ExportPath = Join-Path $auditsDir "Guests_Export.csv"
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile   = Join-Path $logsDir "Export_Guests_$timestamp.log"

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
    Add-ToLog "▶ Starting Guest Users export..."
    Add-ToLog "Export path : $ExportPath"

    # Connect to Microsoft Graph
    Add-ToLog "🔑 Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes "User.Read.All" -NoWelcome | Out-Null
    Add-ToLog "✔ Connected to Microsoft Graph."

    # Retrieve guest users
    Add-ToLog "⏳ Retrieving guest users..."
    $GuestUsers = @()

    $guests = Get-MgUser -All -Filter "UserType eq 'Guest'" -ErrorAction SilentlyContinue
    foreach ($g in $guests) {
        $GuestUsers += [PSCustomObject]@{
            GivenName         = $g.GivenName
            Surname           = $g.Surname
            DisplayName       = $g.DisplayName
            Mail              = $g.Mail
            UserPrincipalName = $g.UserPrincipalName
            Id                = $g.Id
        }
    }

    if ($GuestUsers.Count -eq 0) {
        Add-ToLog "⚠️ No guest users found."
    }
    else {
        $GuestUsers |
            Export-Csv -Path $ExportPath -Encoding UTF8 -NoTypeInformation -Force

        Add-ToLog "✅ Export completed. File saved to: $ExportPath"

        [System.Windows.Forms.MessageBox]::Show(
            "Guest users exported successfully to:`n$ExportPath",
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