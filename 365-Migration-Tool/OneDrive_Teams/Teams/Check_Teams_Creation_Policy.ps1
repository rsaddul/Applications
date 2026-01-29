<#
Developed by: Rhys Saddul
Purpose: Check which security group is allowed to create Microsoft 365 Groups.
#>

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

# --- Initialise Logs ---
$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs\Check_Teams_Creation_Policy"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$timestamp  = (Get-Date).ToString("dd-MM-yyyy")
$logFile = Join-Path $logDir "Check_GroupCreation_Policy_$timestamp.log"

function Add-ToLog {
    param([string]$Message)
    $time = (Get-Date).ToString("HH:mm:ss")
    $entry = "[$time] $Message"
    $entry
    Add-Content -Path $logFile -Value $entry -ErrorAction SilentlyContinue
}

try {
    Add-ToLog "🌐 Connecting to Microsoft Graph..."
    #[System.Environment]::SetEnvironmentVariable("MSAL_DISABLE_WAM", "1", "Process")
    Connect-MgGraph -Scopes "Directory.ReadWrite.All","Group.Read.All","Group.ReadWrite.All" -NoWelcome | Out-Null
    Add-ToLog "✅ Connected to Microsoft Graph."

    $setting = Get-MgBetaDirectorySetting | Where-Object DisplayName -EQ "Group.Unified"

    if (-not $setting) {
        Add-ToLog "❌ No 'Group.Unified' policy found. Group creation is unrestricted."
        $summary = "No policy found — all users can create Microsoft 365 Groups."
    } else {
        $groupId = $setting.Values | Where-Object { $_.Name -eq "GroupCreationAllowedGroupId" } | Select-Object -ExpandProperty Value

        if (![string]::IsNullOrEmpty($groupId)) {
            $group = Get-MgBetaGroup -GroupId $groupId
            Add-ToLog "✅ Group creation restricted to: $($group.DisplayName)"
            $summary = "Group creation restricted to:`n$($group.DisplayName)"
        } else {
            Add-ToLog "⚠️ GroupCreationAllowedGroupId is not set — no restriction applied."
            $summary = "GroupCreationAllowedGroupId not set — group creation unrestricted."
        }
    }

    # --- Success popup (TopMost) ---
    $topMostWindow = New-Object System.Windows.Forms.Form
    $topMostWindow.TopMost = $true
    $topMostWindow.StartPosition = "CenterScreen"
    $topMostWindow.Size = New-Object System.Drawing.Size(0,0)
    $topMostWindow.Show()
    $topMostWindow.Hide()

    [System.Windows.Forms.MessageBox]::Show(
        $topMostWindow,
        "$summary`n`nLogs saved to:`n$logDir",
        "Group Creation Policy Check",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )

    $topMostWindow.Dispose()

}
catch {
    Add-ToLog "❌ Error checking group creation policy: $($_.Exception.Message)"

    # --- Error popup (also TopMost) ---
    $topMostWindow = New-Object System.Windows.Forms.Form
    $topMostWindow.TopMost = $true
    $topMostWindow.StartPosition = "CenterScreen"
    $topMostWindow.Size = New-Object System.Drawing.Size(0,0)
    $topMostWindow.Show()
    $topMostWindow.Hide()

    [System.Windows.Forms.MessageBox]::Show(
        $topMostWindow,
        "Error occurred:`n$($_.Exception.Message)`n`nCheck logs at:`n$logDir",
        "Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )

    $topMostWindow.Dispose()
}
finally {
    Disconnect-MgGraph | Out-Null
    Add-ToLog "🔒 Disconnected from Microsoft Graph."
}

Add-ToLog "🎉 Completed Group Creation Policy Check."