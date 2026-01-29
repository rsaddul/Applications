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
$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Audits"

if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}

$subLogDir = Join-Path $logDir "Get_Specific_OneDrive_URLs"
if (-not (Test-Path $subLogDir)) {
    New-Item -ItemType Directory -Path $subLogDir | Out-Null
}

$timestamp  = (Get-Date).ToString("dd-MM-yyyy")
$successLog = Join-Path $subLogDir "Get_Specific_OneDrive_URLs_Success_$timestamp.log"
$failLog    = Join-Path $subLogDir "Get_Specific_OneDrive_URLs_Fail_$timestamp.log"

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
        } else {
            Add-Content -Path $successLog -Value $entry
        }
    } catch {}
}

# -------------------------------------------------------------
#   PROMPT TENANT ADMIN URL
# -------------------------------------------------------------
$TenantUrl = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter your SharePoint Admin URL:`ne.g. https://thecloudschool-admin.sharepoint.com/",
    "SharePoint Admin URL"
)

# Client ID
if (-not $clientId) {
    $clientId = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Azure App Client ID: Run the Setup PnP Automation script if you dont have these details ", "Client ID", "04c1c1b5-f732-472a-8156-f2b12f06b535")
    if ([string]::IsNullOrWhiteSpace($clientId)) { "❌ Client ID is required."; exit 1 }
}

if ([string]::IsNullOrWhiteSpace($TenantUrl)) {
    Add-ToLog "❌ No SharePoint Admin URL entered. Exiting..." -Failed
    exit
}

if ($TenantUrl -notmatch "-admin\.sharepoint\.com/?$") {
    Add-ToLog "❌ Invalid SharePoint Admin URL format." -Failed
    exit
}

Add-ToLog "🔗 Using Admin URL: $TenantUrl"

# -------------------------------------------------------------
#   OFFER CSV TEMPLATE
# -------------------------------------------------------------
$TemplateResponse = [System.Windows.Forms.MessageBox]::Show(
    "Generate a sample UPN CSV template on Desktop?",
    "Create Template?",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if ($TemplateResponse -eq [System.Windows.Forms.DialogResult]::Yes) {
    try {
        $templatePath = Join-Path ([Environment]::GetFolderPath("Desktop")) "OneDrive_Audit_Template.csv"
@"
UserPrincipalName
rsaddul@thecloudschool.co.uk
jskett@thecloudschool.co.uk
"@ | Out-File -FilePath $templatePath -Encoding UTF8 -Force

        Add-ToLog "✅ Template created: $templatePath"

        [System.Windows.Forms.MessageBox]::Show(
            "Template created at:`n$templatePath",
            "Template Created",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        exit
    }
    catch {
        Add-ToLog "❌ Failed to create template: $($_.Exception.Message)" -Failed
        exit
    }
}

# -------------------------------------------------------------
#   SELECT INPUT CSV (UPN LIST)
# -------------------------------------------------------------
$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
$fileDialog.Title = "Select UPN CSV File"
$fileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
$fileDialog.Filter = "CSV Files (*.csv)|*.csv"

if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $CSVPath = $fileDialog.FileName
    Add-ToLog "📄 Selected UPN CSV: $CSVPath"
} else {
    Add-ToLog "❌ No CSV selected. Exiting..." -Failed
    exit
}

# -------------------------------------------------------------
#   SELECT OUTPUT CSV
# -------------------------------------------------------------
$saveDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveDialog.Title = "Save Audit Output CSV"
$saveDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
$saveDialog.Filter = "CSV Files (*.csv)|*.csv"
$saveDialog.FileName = "OneDrive_Audit_Results.csv"

if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $OutputCSV = $saveDialog.FileName
    Add-ToLog "📁 Output will be saved to: $OutputCSV"
} else {
    Add-ToLog "❌ Output path not selected. Exiting..." -Failed
    exit
}

# -------------------------------------------------------------
#   LOAD CSV
# -------------------------------------------------------------
try {
    $UPNs = Import-Csv -Path $CSVPath -ErrorAction Stop
    Add-ToLog "📥 Loaded $($UPNs.Count) UPN(s) from CSV."
}
catch {
    Add-ToLog "❌ Failed to read CSV: $($_.Exception.Message)" -Failed
    exit
}

# -------------------------------------------------------------
#   CONNECT TO SHAREPOINT ONLINE
# -------------------------------------------------------------
try {
    Connect-PnPOnline -Url $TenantUrl -Interactive -ClientId $clientId
    Add-ToLog "🔐 Connected to SharePoint Admin Center."
}
catch {
    Add-ToLog "❌ Failed to connect: $($_.Exception.Message)" -Failed
    exit
}

# -------------------------------------------------------------
#   FETCH ALL PERSONAL SITES
# -------------------------------------------------------------
Add-ToLog "🔍 Fetching all personal OneDrive sites..."

try {
    $AllSites = Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'"
    Add-ToLog "📊 Retrieved $($AllSites.Count) personal OneDrive sites."
}
catch {
    Add-ToLog "❌ Failed retrieving sites: $($_.Exception.Message)" -Failed
    exit
}

# -------------------------------------------------------------
#   MATCH UPN → ONEDRIVE URL
# -------------------------------------------------------------
$Results = @()
$successCount = 0
$failCount = 0

foreach ($user in $UPNs) {
    $UPN = $user.UserPrincipalName.Trim()

    Add-ToLog "👤 Checking OneDrive for: $UPN"

    $Site = $AllSites | Where-Object { $_.Owner.ToLower() -eq $UPN.ToLower() }

    if ($Site) {
        Add-ToLog "✅ Found: $($Site.Url)"
        $successCount++
    } else {
        Add-ToLog "❌ No OneDrive found for $UPN" -Failed
        $failCount++
    }

    $Results += [PSCustomObject]@{
        UserPrincipalName = $UPN
        OneDriveURL       = if ($Site) { $Site.Url } else { "Not Found" }
    }
}

# -------------------------------------------------------------
#   EXPORT RESULTS
# -------------------------------------------------------------
try {
    $Results | Export-Csv -Path $OutputCSV -NoTypeInformation -Force
    Add-ToLog "📤 Results exported to $OutputCSV"
}
catch {
    Add-ToLog "❌ Failed to export results: $($_.Exception.Message)" -Failed
}

Disconnect-PnPOnline
Add-ToLog "🔒 Disconnected from SharePoint."

# -------------------------------------------------------------
#   SUMMARY POPUP
# -------------------------------------------------------------
[System.Windows.Forms.MessageBox]::Show(
    "OneDrive Audit Completed.`n`nSummary:`n$successCount Found`n$failCount Not Found`n`nLogs saved to:`n$subLogDir",
    "Audit Complete",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)
