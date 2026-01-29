<#
Developed by: Rhys Saddul
#>

param (
    [string]$ExportPath,
    [string]$SharePointAdminUrl,
    [string[]]$AcceptedDomains,
    [string]$ClientId
)

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

# ==============================================================
# INITIALISE LOGGING
# ==============================================================

$logsRoot  = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs"
$logsDir   = Join-Path $logsRoot "Export_365Users"
$auditsDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Audits"

foreach ($dir in @($logsRoot, $logsDir, $auditsDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
    }
}

if (-not $ExportPath) {
    $ExportPath = Join-Path $auditsDir "365UsersExport.csv"
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile   = Join-Path $logsDir "Export_365Users_$timestamp.log"
New-Item -ItemType File -Path $logFile -Force | Out-Null

function Add-ToLog {
    param([string]$Message, [switch]$Failed)
    $time  = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
    $entry = if ($Failed) { "[$time] ❌ $Message" } else { "[$time] $Message" }
    $entry
    Add-Content -Path $logFile -Value $entry -ErrorAction Stop
}

# ==============================================================
# PROMPTS
# ==============================================================

if (-not $SharePointAdminUrl) {
    $SharePointAdminUrl = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Enter SharePoint Admin URL:",
        "SharePoint Admin URL",
        "https://eduthingazurelab-admin.sharepoint.com/"
    )
}

if (-not $AcceptedDomains) {
    $AcceptedDomains = (
        [Microsoft.VisualBasic.Interaction]::InputBox(
            "Enter accepted domains (comma separated):",
            "Accepted Domains",
            "@thecloudschool.co.uk"
        )
    ).Split(",") | ForEach-Object { $_.Trim() }
}

Add-ToLog "▶ Starting Microsoft 365 Users export..."
Add-ToLog "📁 Export path: $ExportPath"

# Client ID (REQUIRED for PnP in PowerShell 7)
if (-not $ClientId) {
    $clientId = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Enter Azure App Client ID: Run the Setup PnP Automation script if you dont have these details",
        "Client ID",
        "04c1c1b5-f732-472a-8156-f2b12f06b535"
    )
    if ([string]::IsNullOrWhiteSpace($clientId)) {
        Add-ToLog "❌ Client ID is required." -Failed
        exit 1
    }
}

# ==============================================================
# CONNECT: MICROSOFT GRAPH (FIRST — CRITICAL)
# ==============================================================

try {
    Add-ToLog "🌐 Connecting to Microsoft Graph..."
    Connect-MgGraph -NoWelcome -ErrorAction Stop
    Add-ToLog "✅ Connected to Microsoft Graph."
}
catch {
    Add-ToLog "❌ Graph connection failed: $($_.Exception.Message)" -Failed
    exit 1
}

# ==============================================================
# GRAPH PHASE — USERS + LICENSES
# ==============================================================

Add-ToLog "👤 Retrieving users..."
$Users = Get-MgUser -All -Property Id,DisplayName,GivenName,Surname,UserPrincipalName

Add-ToLog "📜 Caching license details from Graph..."

$LicenseLookup = @{}
$licCount = $Users.Count
$licIndex = 0

foreach ($u in $Users) {

    $licIndex++

    Write-Progress `
        -Activity "Caching license details from Microsoft Graph" `
        -Status "Processing $licIndex of $licCount ($($u.UserPrincipalName))" `
        -PercentComplete (($licIndex / $licCount) * 100)

    try {
        $lic = Get-MgUserLicenseDetail -UserId $u.Id -ErrorAction Stop
        if ($lic) {
            $LicenseLookup[$u.Id] = ($lic | Select-Object -ExpandProperty SkuPartNumber) -join ", "
        }
        else {
            $LicenseLookup[$u.Id] = "No License Assigned"
        }
    }
    catch {
        $LicenseLookup[$u.Id] = "No License Assigned"
    }
}

Write-Progress -Activity "Caching license details from Microsoft Graph" -Completed
Add-ToLog "✅ Finished caching license details."


# ==============================================================
# SWITCH FROM GRAPH → EXCHANGE (FIX)
# ==============================================================

Disconnect-MgGraph | Out-Null

try {
    Add-ToLog "🌐 Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Add-ToLog "✅ Connected to Exchange Online."
}
catch {
    Add-ToLog "❌ Exchange connection failed: $($_.Exception.Message)" -Failed
    exit 1
}

$ExportData = @()
$TotalUsers = $Users.Count
$i = 0

foreach ($User in $Users) {

    $i++
    Add-ToLog "Processing $i of $TotalUsers - $($User.UserPrincipalName)"

    if ($User.UserPrincipalName -like "*EXT*") { continue }

    if (-not ($AcceptedDomains | ForEach-Object {
            $User.UserPrincipalName.ToLower().EndsWith($_.ToLower().Trim())
        } | Where-Object { $_ })) {
        continue
    }

    $PrimarySMTP = $User.UserPrincipalName
    $MailboxSize = "N/A"
    $Aliases     = @()
    $LicenseType = $LicenseLookup[$User.Id]
    $OneDriveURL = "N/A"
    $OneDriveGB  = "N/A"

# -----------------------------
# EXCHANGE (EXO V3 – SUPPORTED)
# -----------------------------
try {
    $Mailbox = Get-EXOMailbox -Identity $PrimarySMTP -PropertySets All -ErrorAction Stop

    $Aliases = $Mailbox.EmailAddresses |
        Where-Object { $_ -cmatch '^smtp:' } |
        ForEach-Object { $_ -replace '^smtp:', '' }
}
catch {
    Add-ToLog "⚠️ Alias lookup failed for $PrimarySMTP"
}

try {
    $Stats = Get-EXOMailboxStatistics -Identity $PrimarySMTP -ErrorAction Stop

    if ($Stats) {
        $MailboxSize = [math]::Round(
            $Stats.TotalItemSize.Value.ToBytes() / 1GB, 2
        )
    }
}
catch {
    Add-ToLog "⚠️ Mailbox size lookup failed for $PrimarySMTP"
}

    $ExportData += [PSCustomObject]@{
        DisplayName     = $User.DisplayName
        FirstName       = $User.GivenName
        LastName        = $User.Surname
        PrimarySMTP     = $PrimarySMTP
        Aliases         = ($Aliases -join ";")
        MailboxSize     = $MailboxSize
        LicenseType     = $LicenseType
        OneDriveURL     = $OneDriveURL
        OneDriveUsageGB = $OneDriveGB
    }
}

# ==============================================================
# PNP PHASE — ONEDRIVE (LAST)
# ==============================================================

try {
    Add-ToLog "🌐 Connecting to SharePoint Online..."
    Connect-PnPOnline -Url $SharePointAdminUrl -Interactive -ClientId $ClientId -ErrorAction Stop
    Add-ToLog "✅ Connected to SharePoint Online."

    Add-ToLog "📂 Retrieving OneDrive sites..."
    $OneDrives = Get-PnPTenantSite -IncludeOneDriveSites -Detailed |
        Where-Object { $_.Url -like "*-my.sharepoint.com/personal/*" }

    $odLookup = @{}
    foreach ($od in $OneDrives) {
        if ($od.Owner) {
            $odLookup[$od.Owner] = $od
        }
    }

    foreach ($row in $ExportData) {
        if ($odLookup.ContainsKey($row.PrimarySMTP)) {
            $row.OneDriveURL     = $odLookup[$row.PrimarySMTP].Url
            $row.OneDriveUsageGB = [math]::Round($odLookup[$row.PrimarySMTP].StorageUsageCurrent / 1024, 2)
        }
    }
}
catch {
    Add-ToLog "⚠️ OneDrive lookup failed: $($_.Exception.Message)" -Failed
}

# ==============================================================
# EXPORT
# ==============================================================

if ($ExportData.Count -gt 0) {
    $ExportData | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8 -Force
    Add-ToLog "✅ Export complete: $ExportPath"

    [System.Windows.Forms.MessageBox]::Show(
        "Microsoft 365 Users exported to:`n$ExportPath",
        "Export Complete",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}
else {
    Add-ToLog "⚠️ No users matched export criteria." -Failed
}

# ==============================================================
# CLEANUP
# ==============================================================

Disconnect-PnPOnline  | Out-Null
Add-ToLog "🔒 Disconnected from all services."
Add-ToLog "📁 Log saved to: $logFile"
Add-ToLog "🎉 Completed Microsoft 365 Users Export process."