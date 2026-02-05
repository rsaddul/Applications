<#
Cloud 365 Tenant Audit
Developed by: Rhys Saddul

This script acts as an ORCHESTRATOR only.
Each export script:
- Handles its own connections
- Can be run standalone
- Is invoked in-process (GUI launcher controls PS version)

Exports to:
📁 Desktop\CloudDonny_Audits
Logs to:
🗂️ Desktop\CloudDonny_Logs
#>

# ==============================================================
#   STA ENFORCEMENT (GUI COMPATIBILITY)
# ==============================================================

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Write-Warning "Restarting in STA mode for GUI compatibility..."
    Start-Process powershell.exe -ArgumentList "-STA -File `"$PSCommandPath`""
    exit
}

# ==============================================================
#   PATHS & LOGGING
# ==============================================================

$Desktop    = [Environment]::GetFolderPath("Desktop")
$ExportRoot = Join-Path $Desktop "CloudDonny_Audits"
$LogRoot    = Join-Path $Desktop "CloudDonny_Logs"

foreach ($dir in @($ExportRoot, $LogRoot)) {
    if (-not (Test-Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile   = Join-Path $LogRoot "Tenant_Audit_$timestamp.log"

function Add-ToLog {
    param (
        [string]$Message,
        [switch]$Failed
    )

    $time  = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $entry = if ($Failed) { "[$time] ❌ $Message" } else { "[$time] $Message" }

    Write-Host $entry
    try { Add-Content -Path $logFile -Value $entry -ErrorAction SilentlyContinue } catch {}
}

Add-ToLog "▶ Cloud 365 Tenant Audit started."
Add-ToLog "📁 Export path: $ExportRoot"
Add-ToLog "🗂️ Log file: $logFile"

# ==============================================================
#   GUI PROMPTS
# ==============================================================

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

# ==============================================================
#   SECTION SELECTION
# ==============================================================

function Ask-YesNo {
    param ($Message, $Title)
    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    ) -eq [System.Windows.Forms.DialogResult]::Yes
}

$res = [System.Windows.Forms.MessageBox]::Show(
    "Run ALL sections? (Yes = All, No = Pick, Cancel = Abort)",
    "Sections",
    [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if ($res -eq [System.Windows.Forms.DialogResult]::Cancel) {
    Add-ToLog "Audit cancelled by user." -Failed
    exit
}

$Sections = [ordered]@{
    DL       = $true
    MESG     = $true
    M365     = $true
    SG       = $true
    SHARED   = $true
    CONTACTS = $true
    GUESTS   = $true
    USERS    = $true
    TEAMS    = $true
}

if ($res -eq [System.Windows.Forms.DialogResult]::No) {
    $Sections.DL       = Ask-YesNo "Run Distribution Lists?" "DL"
    $Sections.MESG     = Ask-YesNo "Run Mail-Enabled Security Groups?" "MESG"
    $Sections.M365     = Ask-YesNo "Run Microsoft 365 Groups?" "M365"
    $Sections.SG       = Ask-YesNo "Run Security Groups?" "Security Groups"
    $Sections.SHARED   = Ask-YesNo "Run Shared Mailboxes?" "Shared Mailboxes"
    $Sections.CONTACTS = Ask-YesNo "Run Contacts?" "Contacts"
    $Sections.GUESTS   = Ask-YesNo "Run Guests?" "Guests"
    $Sections.USERS    = Ask-YesNo "Run Users?" "Users"
    $Sections.TEAMS    = Ask-YesNo "Run Teams?" "Teams"
}

Add-ToLog "✅ Sections configured."

# ==============================================================
#   SCRIPT INVOKER
# ==============================================================

$ScriptRoot = Split-Path -Parent $PSCommandPath

function Invoke-AuditScript {
    param (
        [string]$ScriptName,
        [string]$SectionName
    )

    $scriptPath = Join-Path $ScriptRoot $ScriptName

    if (-not (Test-Path $scriptPath)) {
        Add-ToLog "❌ Missing script: $ScriptName" -Failed
        return
    }

    Add-ToLog "▶ Running $SectionName..."

    try {
        pwsh -NoProfile -ExecutionPolicy Bypass -File $scriptPath
        Add-ToLog "✔ Completed $SectionName."
    }
    catch {
        Add-ToLog "❌ $SectionName failed: $($_.Exception.Message)" -Failed
    }
}

# ==============================================================
#   EXECUTION
# ==============================================================

if ($Sections.DL)       { Invoke-AuditScript "Distribution_List_Export.ps1"        "Distribution Lists" }
if ($Sections.MESG)     { Invoke-AuditScript "Mail-Enabled_Security_Group_Export.ps1" "Mail-Enabled Security Groups" }
if ($Sections.M365)     { Invoke-AuditScript "M365_Export.ps1"                      "Microsoft 365 Groups" }
if ($Sections.SG)       { Invoke-AuditScript "Security_Group_Export.ps1"            "Security Groups" }
if ($Sections.SHARED)   { Invoke-AuditScript "Shared_Mailbox_Export.ps1"            "Shared Mailboxes" }
if ($Sections.CONTACTS) { Invoke-AuditScript "Contacts_Export.ps1"                  "Contacts" }
if ($Sections.GUESTS)   { Invoke-AuditScript "Guests_Export.ps1"                    "Guests" }
if ($Sections.USERS)    { Invoke-AuditScript "Users_Export.ps1"                     "Users" }
if ($Sections.TEAMS)    { Invoke-AuditScript "Teams_Export.ps1"                     "Teams" }

# ==============================================================
#   ZIP PACKAGING
# ==============================================================

try {
    Add-ToLog "📦 Creating ZIP package..."

    $zipStamp = Get-Date -Format "yyyyMMdd_HHmm"
    $zipPath  = Join-Path $ExportRoot "Tenant_Audit_$zipStamp.zip"

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($zipPath, 'Create')

    Get-ChildItem $ExportRoot -File -Include *.csv | ForEach-Object {
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
            $zip, $_.FullName, $_.Name
        ) | Out-Null
    }

    if (Test-Path $logFile) {
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
            $zip, $logFile, (Split-Path $logFile -Leaf)
        ) | Out-Null
    }

    $zip.Dispose()

    Add-ToLog "✔ ZIP created -> $zipPath"

    [System.Windows.Forms.MessageBox]::Show(
        "Tenant Audit completed successfully.`n`nZIP saved to:`n$zipPath",
        "Audit Complete",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )

}
catch {
    Add-ToLog "❌ ZIP packaging failed: $($_.Exception.Message)" -Failed
}

Add-ToLog "🎉 Tenant Audit finished."