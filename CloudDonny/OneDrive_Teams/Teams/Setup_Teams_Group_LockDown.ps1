<#
Developed by: Rhys Saddul
Purpose: Restrict Microsoft 365 Group creation to a specific security group.
#>

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

# ==============================================================
#   INITIALISE LOGGING
# ==============================================================

$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs\Setup_Teams_Group_LockDown"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$timestamp  = (Get-Date).ToString("dd-MM-yyyy")
$successLog = Join-Path $logDir "Restrict_Group_Creation_Success_$timestamp.log"
$failLog    = Join-Path $logDir "Restrict_Group_Creation_Failed_$timestamp.log"

function Add-ToLog {
    param([string]$Message, [switch]$Failed)
    $time = (Get-Date).ToString("HH:mm:ss")
    $entry = "[$time] $Message"

    # Output to GUI
    $entry

    # Write to file
    try {
        if ($Failed) { Add-Content -Path $failLog -Value $entry }
        else { Add-Content -Path $successLog -Value $entry }
    } catch {}
}

# ==============================================================
#   PROMPT FOR GROUP NAME
# ==============================================================

$GroupName = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter the name of the Security Group that will be allowed to create M365 Groups:",
    "Allowed Group Name",
    "SG TCS Teams Group Creation"
)

if ([string]::IsNullOrWhiteSpace($GroupName)) {
    Add-ToLog "❌ Cancelled: Group name is required." -Failed
    exit 1
}

$AllowGroupCreation = "False"

# ==============================================================
#   CONNECT TO MICROSOFT GRAPH
# ==============================================================

try {
    Add-ToLog "🌐 Connecting to Microsoft Graph..."
    [System.Environment]::SetEnvironmentVariable("MSAL_DISABLE_WAM", "1", "Process")
    Connect-MgGraph -Scopes "Directory.ReadWrite.All","Group.Read.All" -NoWelcome | Out-Null
    Add-ToLog "🌐 Connected to Microsoft Graph."
}
catch {
    Add-ToLog "❌ Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Failed
    exit 1
}

# ==============================================================
#   CONFIGURE GROUP CREATION SETTINGS
# ==============================================================

# ==============================================================
#   CONFIGURE GROUP CREATION SETTINGS
# ==============================================================

try {
    Add-ToLog "🌐 Checking for existing Group.Unified settings..."
    $settingsObjectID = (Get-MgBetaDirectorySetting | Where-Object DisplayName -EQ "Group.Unified").Id

    if (-not $settingsObjectID) {
        Add-ToLog "⚠️ No existing Group.Unified settings found — creating new configuration..."
        $params = @{
            templateId = "62375ab9-6b52-47ed-826b-58e47e0e304b"
            values = @(
                @{
                    name  = "EnableMSStandardBlockedWords"
                    value = $true
                }
            )
        }
        New-MgBetaDirectorySetting -BodyParameter $params | Out-Null
        $settingsObjectID = (Get-MgBetaDirectorySetting | Where-Object DisplayName -EQ "Group.Unified").Id
        Add-ToLog "✅ Created new Group.Unified settings."
    }

   Add-ToLog "🌐 Locating specified group: $GroupName"
$group = Get-MgBetaGroup -Filter "displayName eq '$GroupName'" -ErrorAction SilentlyContinue

# --- AUTO-CREATE GROUP IF NOT FOUND ---
if (-not $group) {
    Add-ToLog "⚠️ Group not found — creating it automatically: $GroupName"

    $params = @{
        displayName     = $GroupName
        securityEnabled = $true
        mailEnabled     = $false
        mailNickname    = $GroupName.Replace(" ", "")
    }

    try {
        $group = New-MgBetaGroup -BodyParameter $params
        Add-ToLog "✅ Created new security group: $GroupName"
    }
    catch {
        Add-ToLog "❌ Failed to create security group: $($_.Exception.Message)" -Failed
        exit 1
    }
}

$groupId = $group.Id


    Add-ToLog "🌐 Applying restriction to only allow $GroupName to create M365 Groups..."

    $params = @{
        templateId = "62375ab9-6b52-47ed-826b-58e47e0e304b"
        values = @(
            @{
                name  = "EnableGroupCreation"
                value = $AllowGroupCreation
            },
            @{
                name  = "GroupCreationAllowedGroupId"
                value = $groupId
            }
        )
    }

    Update-MgBetaDirectorySetting -DirectorySettingId $settingsObjectID -BodyParameter $params | Out-Null
    Add-ToLog "✅ Restriction applied successfully. Group creation now limited to: $GroupName"

    $updatedSettings = Get-MgBetaDirectorySetting -DirectorySettingId $settingsObjectID
    $settingsList = ($updatedSettings.Values |
    ForEach-Object { "$($_.Name)=$($_.Value)" }) -join ", "

    Add-ToLog "🌐 Current settings applied: $settingsList"

}   # ← THIS IS THE MISSING CLOSING BRACE
catch {
    Add-ToLog "❌ Failed to update group creation settings: $($_.Exception.Message)" -Failed
}
finally {
    Disconnect-MgGraph | Out-Null
    Add-ToLog "🔒 Disconnected from Microsoft Graph."
}


# ==============================================================
#   COMPLETION POPUP (TopMost)
# ==============================================================

$topMostWindow = New-Object System.Windows.Forms.Form
$topMostWindow.TopMost = $true
$topMostWindow.StartPosition = "CenterScreen"
$topMostWindow.Size = New-Object System.Drawing.Size(0,0)
$topMostWindow.Show()
$topMostWindow.Hide()

[System.Windows.Forms.MessageBox]::Show(
    $topMostWindow,
    "Group creation restriction process completed.`n`nAllowed group:`n$GroupName`n`nLogs saved to:`n$logDir",
    "Operation Complete",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)

$topMostWindow.Dispose()

Add-ToLog "✅ Completed Restrict Group Creation process."