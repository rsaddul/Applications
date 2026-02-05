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

$subLogDir = Join-Path $logDir "Remove_SiteCollectionAdmin_Specific_Users"
if (-not (Test-Path $subLogDir)) {
    New-Item -ItemType Directory -Path $subLogDir | Out-Null
}

$timestamp  = (Get-Date).ToString("dd-MM-yyyy")
$successLog = Join-Path $subLogDir "Remove_SiteCollectionAdmin_Specific_Users_Success_$timestamp.log"
$failLog    = Join-Path $subLogDir "Remove_SiteCollectionAdmin_Specific_Users_Fail_$timestamp.log"

function Add-ToLog {
    param(
        [string]$Message,
        [switch]$Failed
    )

    $time  = (Get-Date).ToString("HH:mm:ss")
    $entry = "[$time] $Message"

    Write-Host $entry

    try {
        if ($Failed) {
            Add-Content -Path $failLog -Value $entry
        }
        else {
            Add-Content -Path $successLog -Value $entry
        }
    } catch {}
}

# -------------------------------------------------------------
#   PROMPT FOR PnP Application ID
# -------------------------------------------------------------
$clientId = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter your PnP Application ID:",
    "PnP Application ID",
    "0a2e4081-bb3a-45a2-b63d-a5abe8657965"
)

if ([string]::IsNullOrWhiteSpace($clientId)) {
    Add-ToLog "❌ No PnP Application ID entered. Exiting..." -Failed
    exit
}

if ($clientId -notmatch '^[0-9a-fA-F-]{36}$') {
    Add-ToLog "❌ Invalid PnP Application ID format." -Failed
    exit
}

Add-ToLog "🔗 Using PnP Application ID: $clientId"

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
        $templatePath = Join-Path ([Environment]::GetFolderPath("Desktop")) "OneDrive_Remove_Admin_Template.csv"
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
        Add-ToLog "❌ Failed to create CSV template: $($_.Exception.Message)" -Failed
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

# -------------------------------------------------------------
#   IMPORT CSV
# -------------------------------------------------------------
try {
    $Records = Import-Csv -Path $csvPath -ErrorAction Stop
    Add-ToLog "📥 Loaded $($Records.Count) record(s) from CSV."
}
catch {
    Add-ToLog "❌ Failed to load CSV: $($_.Exception.Message)" -Failed
    exit
}

# -------------------------------------------------------------
#   PROCESS REMOVALS
# -------------------------------------------------------------
$successCount = 0
$failCount    = 0

foreach ($Row in $Records) {

    $OneDriveSiteUrl = $Row.OneDriveSiteUrl.Trim()
    $AdminToRemove   = $Row.SiteCollAdmin.Trim()

    Add-ToLog "📁 Processing: $OneDriveSiteUrl"
    Add-ToLog "👤 Removing Site Collection Admin: $AdminToRemove"

    try {
        Connect-PnPOnline -Url $OneDriveSiteUrl -Interactive -ClientId $clientId -ErrorAction Stop

        Remove-PnPSiteCollectionAdmin -Owners $AdminToRemove -ErrorAction Stop

        Add-ToLog "✅ Successfully removed $AdminToRemove from $OneDriveSiteUrl"
        $successCount++
    }
    catch {
        Add-ToLog "❌ Failed to remove admin from $OneDriveSiteUrl — $($_.Exception.Message)" -Failed
        $failCount++
    }
}

Disconnect-PnPOnline
Add-ToLog "🔒 Disconnected from PnP Online."

# -------------------------------------------------------------
#   SUMMARY & POPUP
# -------------------------------------------------------------
Add-ToLog "---------------------------------------------------"
Add-ToLog "SUMMARY: $successCount successful, $failCount failed."
Add-ToLog "Logs saved to:"
Add-ToLog "  $successLog"
Add-ToLog "  $failLog"

[System.Windows.Forms.MessageBox]::Show(
    "Operation Completed.`n`nSummary:`n$successCount Successful`n$failCount Failed`n`nLogs saved to:`n$subLogDir",
    "Completed",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)
