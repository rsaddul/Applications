<#
Developed by: Rhys Saddul
Based on Microsoft’s official documentation for Request-SPOPersonalSite.
#>

Add-Type -AssemblyName Microsoft.VisualBasic

# ==============================================================
#   INITIALISE LOGGING
# ==============================================================

$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$timestamp  = (Get-Date).ToString("dd-MM-yyyy")
$successLog = Join-Path $logDir "Provision_All_Licensed_Users_OneDrive\OneDrive_AllLicensed_Success_$timestamp.log"
$failLog    = Join-Path $logDir "Provision_All_Licensed_Users_OneDrive\OneDrive_AllLicensed_Failed_$timestamp.log"

function Add-ToLog {
    param([string]$Message, [switch]$Failed)
    $time = (Get-Date).ToString("HH:mm:ss")
    $entry = "[$time] $Message"
    Write-Host $entry
    if ($Failed) { Add-Content -Path $failLog -Value $entry } else { Add-Content -Path $successLog -Value $entry }
}

# ==============================================================
#   PROMPT FOR SHAREPOINT ADMIN URL
# ==============================================================

$sharepointURL = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter your SharePoint Admin URL (e.g., https://thecloudschool-admin.sharepoint.com):",
    "SharePoint Admin URL",
    "https://thecloudschool-admin.sharepoint.com"
)

if ([string]::IsNullOrWhiteSpace($sharepointURL)) {
    Add-ToLog "❌ No SharePoint URL entered. Exiting..." -Failed
    exit 1
}

# ==============================================================
#   PROMPT FOR TENANT ID
# ==============================================================

$tenantID = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter your Azure Tenant ID (GUID):",
    "Tenant ID",
    "adbb7476-9ad1-3b86-bc3a-8c83b8d0a3aa"
)

if ([string]::IsNullOrWhiteSpace($tenantID)) {
    Add-ToLog "❌ No Tenant ID entered. Exiting..." -Failed
    exit 1
}

# ==============================================================
#   CONNECT TO GRAPH AND SHAREPOINT
# ==============================================================

try {
    [System.Environment]::SetEnvironmentVariable("MSAL_DISABLE_WAM", "1", "Process")

    Add-ToLog "🔑 Connecting to Microsoft Graph (MFA supported)..."
    Connect-MgGraph -TenantId $tenantID -Scopes "User.Read.All" -NoWelcome | Out-Null
    Add-ToLog "✅ Connected to Microsoft Graph."

    Add-ToLog "🔑 Connecting to SharePoint Online..."
    Connect-SPOService -Url $sharepointURL -ErrorAction Stop
    Add-ToLog "✅ Connected to SharePoint Online."
}
catch {
    Add-ToLog "❌ Failed to connect to Microsoft 365 services: $($_.Exception.Message)" -Failed
    exit 1
}

# ==============================================================
#   GET LICENSED USERS AND REQUEST PERSONAL SITES
# ==============================================================

try {
    Add-ToLog "🔍 Retrieving licensed users..."
    $users = Get-MgUser -Filter "assignedLicenses/$count ne 0" -ConsistencyLevel eventual -CountVariable licensedUserCount -All -Select UserPrincipalName

    if (-not $users -or $users.Count -eq 0) {
        Add-ToLog "⚠️ No licensed users found." -Failed
        exit 1
    }

    $list = @()
    $processed = 0
    $batchSize = 199
    $successCount = 0
    $failCount = 0
    $total = $users.Count

    Add-ToLog "🚀 Starting OneDrive provisioning for $total licensed users..."

    foreach ($u in $users) {
        $processed++
        $list += $u.UserPrincipalName

        if ($list.Count -ge $batchSize) {
            try {
                Request-SPOPersonalSite -UserEmails $list -NoWait -ErrorAction Stop
                Add-ToLog "✅ Submitted batch of $($list.Count) users ($processed/$total)."
                $successCount += $list.Count
            }
            catch {
                Add-ToLog "❌ Failed to submit batch ($processed/$total): $($_.Exception.Message)" -Failed
                $failCount += $list.Count
            }
            Start-Sleep -Milliseconds 700
            $list = @()
        }
    }

    # Final batch
    if ($list.Count -gt 0) {
        try {
            Request-SPOPersonalSite -UserEmails $list -NoWait -ErrorAction Stop
            Add-ToLog "✅ Submitted final batch for $($list.Count) users."
            $successCount += $list.Count
        }
        catch {
            Add-ToLog "❌ Failed to submit final batch: $($_.Exception.Message)" -Failed
            $failCount += $list.Count
        }
    }

    Add-ToLog "🔹 Summary: $successCount success(es), $failCount failure(s)."
    Add-ToLog "📁 Logs saved to:`n   $successLog`n   $failLog"

# ✅ Show popup summary to user
[System.Windows.Forms.MessageBox]::Show(
    "OneDrive provisioning process completed.`n`nSummary:`n$successCount success(es)`n$failCount failure(s)`n`nLogs saved to:`n$logDir",
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
        Add-ToLog "⚠️ Could not disconnect from SharePoint (session may already be closed)." -Failed
    }

    try {
        Disconnect-MgGraph | Out-Null
        Add-ToLog "🔒 Disconnected from Microsoft Graph."
    }
    catch {
        Add-ToLog "⚠️ Could not disconnect from Graph (session may already be closed)." -Failed
    }
}

Add-ToLog "🎉 Completed OneDrive provisioning for all licensed users."
