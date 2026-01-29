<#
Developed by: Rhys Saddul
Based on Microsoft’s documentation for Request-SPOPersonalSite.
#>

# ============================================
#   LOAD ASSEMBLIES
# ============================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

# ============================================
#   INITIALISE LOGGING
# ============================================
$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}

$timestamp  = (Get-Date).ToString("dd-MM-yyyy")
$successLog = Join-Path $logDir "Provision_Specific_Users_OneDrive\OneDrive_Provisioning_Success_$timestamp.log"
$failLog    = Join-Path $logDir "Provision_Specific_Users_OneDrive\OneDrive_Provisioning_Failed_$timestamp.log"

function Add-ToLog {
    param(
        [string]$Message,
        [switch]$Failed
    )

    $time  = (Get-Date).ToString("HH:mm:ss")
    $entry = "[$time] $Message"

    # Output to pipeline (GUI-readable)
    $entry

    # Write to log files
    try {
        if ($Failed) {
            Add-Content -Path $failLog -Value $entry -ErrorAction SilentlyContinue
        }
        else {
            Add-Content -Path $successLog -Value $entry -ErrorAction SilentlyContinue
        }
    } catch {}
}

# ============================================
#   PROMPT FOR SHAREPOINT ADMIN URL
# ============================================
$sharePointUrl = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter your SharePoint Admin URL (e.g., https://contoso-admin.sharepoint.com/):",
    "SharePoint Admin URL",
    "https://thecloudschool-admin.sharepoint.com/"
)

if ([string]::IsNullOrWhiteSpace($sharePointUrl)) {
    Add-ToLog "❌ No SharePoint URL entered. Exiting..." -Failed
    exit 1
}

# ============================================
#   PROMPT TO GENERATE USER LIST TEMPLATE
# ============================================
$answer = [System.Windows.Forms.MessageBox]::Show(
    "Would you like to generate a sample user list (TXT file) on your Desktop?",
    "Generate Template",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {

    try {
        $templatePath = Join-Path ([Environment]::GetFolderPath("Desktop")) "UserList_OneDrive_Template.txt"

        @"
rsaddul@thecloudschool.co.uk
jskett@thecloudschool.co.uk
rahemd@thecloudschool.co.uk
"@ | Out-File -FilePath $templatePath -Encoding UTF8 -Force

        [System.Windows.Forms.MessageBox]::Show(
            "Template created successfully:`n$templatePath",
            "Template Created",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        Add-ToLog "✅ Generated example user list at: $templatePath"
        exit 0
    }
    catch {
        Add-ToLog "⚠️ Failed to create user list template: $($_.Exception.Message)" -Failed
        exit 1
    }
}

# ============================================
#   PROMPT FOR USER LIST FILE
# ============================================
$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
$fileDialog.Title = "Select User List File"
$fileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
$fileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"

if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $userListPath = $fileDialog.FileName
    Add-ToLog "📄 Using User List: $userListPath"
}
else {
    Add-ToLog "❌ No user list file selected. Exiting..." -Failed
    exit 1
}

if (-not (Test-Path -Path $userListPath)) {
    Add-ToLog "❌ User list file not found: $userListPath" -Failed
    exit 1
}

# ============================================
#   CONNECT TO SHAREPOINT ONLINE
# ============================================
try {
    Connect-SPOService -Url $sharePointUrl -ErrorAction Stop
    Add-ToLog "✅ Connected to SharePoint Online: $sharePointUrl"
}
catch {
    Add-ToLog "❌ Failed to connect to SharePoint Online: $($_.Exception.Message)" -Failed
    exit 1
}

# ============================================
#   PROCESS USERS IN BATCHES
# ============================================
try {

    $users = Get-Content -Path $userListPath | Where-Object { $_ -and $_.Trim() -ne "" }
    if (-not $users -or $users.Count -eq 0) {
        Add-ToLog "⚠️ No users found in list." -Failed
        exit 1
    }

    $batchSize      = 199
    $currentBatch   = @()
    $processed      = 0
    $total          = $users.Count
    $successCount   = 0
    $failCount      = 0

    Add-ToLog "🚀 Starting OneDrive provisioning for $total user(s)..."

    foreach ($user in $users) {

        $processed++
   	$cleanUser = $user.Trim()
    	$currentBatch += $cleanUser

	Add-ToLog "🔄 Processing user: $user" 

        if ($currentBatch.Count -ge $batchSize) {

            try {
                Request-SPOPersonalSite -UserEmails $currentBatch -NoWait -ErrorAction Stop
                Add-ToLog "✅ Submitted batch for $($currentBatch.Count) users ($processed/$total)."
                $successCount += $currentBatch.Count
            }
            catch {
                Add-ToLog "❌ Failed batch request: $($_.Exception.Message)" -Failed
                $failCount += $currentBatch.Count
            }

            Start-Sleep -Milliseconds 700
            $currentBatch = @()
        }
    }

    if ($currentBatch.Count -gt 0) {
        try {
            Request-SPOPersonalSite -UserEmails $currentBatch -NoWait -ErrorAction Stop
            Add-ToLog "✅ Submitted final batch for $($currentBatch.Count) users."
            $successCount += $currentBatch.Count
        }
        catch {
            Add-ToLog "❌ Failed final batch: $($_.Exception.Message)" -Failed
            $failCount += $currentBatch.Count
        }
    }

    Add-ToLog "🔹 Summary: $successCount submitted successfully, $failCount failed."
    Add-ToLog "📁 Logs saved to:`n   $successLog`n   $failLog"

    # Show summary popup
    [System.Windows.Forms.MessageBox]::Show(
        "OneDrive provisioning completed.`n`nSummary:`n$successCount successful`n$failCount failed`n`nLogs saved to:`n$logDir",
        "Operation Complete",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )

}
catch {
    Add-ToLog "❌ Fatal error during provisioning: $($_.Exception.Message)" -Failed
}
finally {

    try {
        Disconnect-SPOService | Out-Null
        Add-ToLog "🔒 Disconnected from SharePoint Online."
    }
    catch {
        Add-ToLog "⚠️ Could not disconnect from SharePoint Online (session may already be closed)." -Failed
    }
}

Add-ToLog "🎉 Completed OneDrive provisioning process."
