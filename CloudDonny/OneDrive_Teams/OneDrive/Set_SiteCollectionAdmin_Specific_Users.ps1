<#
Developed by: Rhys Saddul
#>

# -------------------------------------------------------------
#   LOAD ASSEMBLIES
# -------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

# -------------------------------------------------------------
#   INITIALISE LOGGING
# -------------------------------------------------------------
$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs"

if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}

# Create subfolder
$subLogDir = Join-Path $logDir "OneDrive_Owner_Updates"
if (-not (Test-Path $subLogDir)) {
    New-Item -ItemType Directory -Path $subLogDir | Out-Null
}

$timestamp  = (Get-Date).ToString("dd-MM-yyyy")
$successLog = Join-Path $subLogDir "Set_SiteCollectionAdmin_Specific_Users_$timestamp.log"
$failLog    = Join-Path $subLogDir "Set_SiteCollectionAdmin_Specific_Users_Failed_$timestamp.log"

function Add-ToLog {
    param(
        [string]$Message,
        [switch]$Failed
    )

    $time  = (Get-Date).ToString("HH:mm:ss")
    $entry = "[$time] $Message"

    # Show in console (GUI readable)
    $entry

    try {
        if ($Failed) {
            Add-Content -Path $failLog -Value $entry -ErrorAction SilentlyContinue
        } else {
            Add-Content -Path $successLog -Value $entry -ErrorAction SilentlyContinue
        }
    } catch {}
}

# -------------------------------------------------------------
#   PROMPT FOR SHAREPOINT ADMIN URL
# -------------------------------------------------------------
$sharePointAdminUrl = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter your SharePoint Admin URL:`n(e.g., https://thecloudschool-admin.sharepoint.com/)",
    "SharePoint Admin URL",
    "https://thecloudschool-admin.sharepoint.com/"
)

if ([string]::IsNullOrWhiteSpace($sharePointAdminUrl)) {
    Add-ToLog "❌ No SharePoint Admin URL entered. Exiting..." -Failed
    exit
}

if ($sharePointAdminUrl -notmatch "-admin\.sharepoint\.com/?$") {
    Add-ToLog "❌ Invalid SharePoint Admin URL format." -Failed
    exit
}

Add-ToLog "🔗 Using SharePoint Admin URL: $sharePointAdminUrl"

# --------------------------------
#   PROMPT FOR PnP Application ID
# --------------------------------
$PnPApplicationID = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter your PnP Application ID:`n(e.g., 04c1c1b5-f732-472a-8156-f2b12f06b535 If you dont have this then run the setup PnP Application )",
    "PnP Application ID",
    "04c1c1b5-f732-472a-8156-f2b12f06b535"
)

if ([string]::IsNullOrWhiteSpace($sharePointAdminUrl)) {
    Add-ToLog "❌ No PnP Application ID entered. Exiting..." -Failed
    exit
}

if ($PnPApplicationID -notmatch '^[0-9a-fA-F-]{36}$') {
    Add-ToLog "❌ Invalid PnP Application ID format." -Failed
    exit
}

Add-ToLog "🔗 Using PnP Application ID $PnPApplicationID"

# -------------------------------------------------------------
#   PROMPT TO GENERATE SAMPLE CSV
# -------------------------------------------------------------
$answer = [System.Windows.Forms.MessageBox]::Show(
    "Would you like to generate a sample CSV template on your Desktop?",
    "Generate Template?",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
    try {
        $templatePath = Join-Path ([Environment]::GetFolderPath("Desktop")) "OneDrive_Owner_Template.csv"
@"
OneDriveSiteUrl,GlobalAdmin,SiteCollAdmin
https://eduthingazurelab-my.sharepoint.com/personal/rsaddul_thecloudschool_co_uk,rsaddul@thecloudschool.co.uk,globaladmin@thecloudschool.co.uk
https://eduthingazurelab-my.sharepoint.com/personal/jskett_thecloudschool_co_uk,rsaddul@thecloudschool.co.uk,globaladmin@thecloudschool.co.uk
"@ | Out-File -FilePath $templatePath -Encoding UTF8 -Force

        [System.Windows.Forms.MessageBox]::Show(
            "Template created at:`n$templatePath",
            "Template Created",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        Add-ToLog "✅ Created sample CSV at $templatePath"
        exit
    }
    catch {
        Add-ToLog "❌ Failed to create template: $($_.Exception.Message)" -Failed
        exit
    }
}

# -------------------------------------------------------------
#   PROMPT FOR CSV FILE
# -------------------------------------------------------------
$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
$fileDialog.Title = "Select CSV File"
$fileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
$fileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"

if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $csvPath = $fileDialog.FileName
    Add-ToLog "📄 CSV Selected: $csvPath"
}
else {
    Add-ToLog "❌ No CSV selected. Exiting..." -Failed
    exit
}

if (-not (Test-Path $csvPath)) {
    Add-ToLog "❌ CSV file not found: $csvPath" -Failed
    exit
}

# -------------------------------------------------------------
#   CONNECT TO SHAREPOINT ADMIN
# -------------------------------------------------------------
Add-ToLog "🔐 Connecting to PnP Online..."

try {
    Connect-PnPOnline -Url $sharePointAdminUrl -Interactive -ClientId $PnPApplicationID -ErrorAction Stop
    Add-ToLog "✅ Connected to: $sharePointAdminUrl"
}
catch {
    Add-ToLog "❌ Failed to connect: $($_.Exception.Message)" -Failed
    exit
}

# -------------------------------------------------------------
#   IMPORT CSV
# -------------------------------------------------------------
try {
    $Records = Import-Csv -Path $csvPath -ErrorAction Stop
    Add-ToLog "📥 Loaded $($Records.Count) record(s) from CSV."
}
catch {
    Add-ToLog "❌ Failed to read CSV: $($_.Exception.Message)" -Failed
    exit
}

# -------------------------------------------------------------
#   PROCESS RECORDS
# -------------------------------------------------------------
$successCount = 0
$failCount    = 0

foreach ($Row in $Records) {

    $OneDriveSiteUrl = $Row.OneDriveSiteUrl.Trim()
    $GlobalAdmin     = $Row.GlobalAdmin.Trim()
    $SiteCollAdmin   = $Row.SiteCollAdmin.Trim()

    Add-ToLog "📁 Processing: $OneDriveSiteUrl"

    $Owners = @()
    if ($GlobalAdmin)   { $Owners += $GlobalAdmin }
    if ($SiteCollAdmin) { $Owners += $SiteCollAdmin }

    Add-ToLog "👥 Setting Owners: $($Owners -join ', ')"

    try {
        Set-PnPTenantSite -Url $OneDriveSiteUrl -Owners $Owners -ErrorAction Stop
        Add-ToLog "✅ Updated owners for $OneDriveSiteUrl"
        $successCount++
    }
    catch {
        Add-ToLog "❌ Failed to update $OneDriveSiteUrl — $($_.Exception.Message)" -Failed
        $failCount++
    }
}

# -------------------------------------------------------------
#   DISCONNECT & SUMMARY
# -------------------------------------------------------------
Disconnect-PnPOnline
Add-ToLog "🔒 Disconnected from PnP Online."

Add-ToLog "---------------------------------------------------"
Add-ToLog "SUMMARY: $successCount successful, $failCount failed."
Add-ToLog "Logs saved to:"
Add-ToLog "  $successLog"
Add-ToLog "  $failLog"

# Show summary popup
[System.Windows.Forms.MessageBox]::Show(
    "Operation Completed.`n`nSummary:`n$successCount Successful`n$failCount Failed`n`nLogs saved to:`n$subLogDir",
    "Completed",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)