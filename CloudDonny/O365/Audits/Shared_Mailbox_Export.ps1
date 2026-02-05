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
$logsDir   = Join-Path $logsRoot "Export_SharedMailboxPermissions"
$auditsDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Audits"

foreach ($dir in @($logsRoot, $logsDir, $auditsDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
    }
}

if (-not $ExportPath) {
    $ExportPath = Join-Path $auditsDir "Shared_Mailboxes_Export.csv"
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile   = Join-Path $logsDir "Export_SharedMailboxPermissions_$timestamp.log"

function Add-ToLog {
    param([string]$Message, [switch]$Failed)

    $time  = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $entry = if ($Failed) { "[$time] ❌ $Message" } else { "[$time] $Message" }

    $entry
    try {
        Add-Content -Path $logFile -Value $entry -ErrorAction SilentlyContinue
    } catch {}
}

Add-ToLog "▶ Starting Shared Mailbox Permissions export..."
Add-ToLog "📁 Export path: $ExportPath"

# ==============================================================
#   CONNECT TO EXCHANGE ONLINE
# ==============================================================

try {
    Add-ToLog "🌐 Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowProgress $true -ShowBanner:$false
    Add-ToLog "✅ Connected to Exchange Online."
}
catch {
    Add-ToLog "❌ Failed to connect to Exchange Online: $($_.Exception.Message)" -Failed
    [System.Windows.Forms.MessageBox]::Show(
        "Connection failed:`n$($_.Exception.Message)",
        "Exchange Online Connection Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

# ==============================================================
#   EXPORT SHARED MAILBOX PERMISSIONS
# ==============================================================

try {
    Add-ToLog "⏳ Retrieving shared mailboxes..."
    $permissionsArray = @()

    $mailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
    foreach ($mailbox in $mailboxes) {
        $mailboxName  = $mailbox.DisplayName
        $mailboxEmail = $mailbox.PrimarySmtpAddress
        Add-ToLog "📬 Processing mailbox: $mailboxName <$mailboxEmail>"

        # --- Full Access (Read and Manage) ---
        try {
            $readManage = Get-MailboxPermission -Identity $mailboxName | Where-Object { $_.User -like "*@*" -and $_.AccessRights -contains "FullAccess" }
            foreach ($perm in $readManage) {
                $permissionsArray += [PSCustomObject]@{
                    "Mailbox Name"           = $mailboxName
                    "Shared Mailbox Address" = $mailboxEmail
                    "Permission Type"        = "Read and manage"
                    "User"                   = $perm.User
                }
            }
        } catch {
            Add-ToLog ("⚠️ Failed to read Full Access for {0}: {1}" -f $mailboxName, $_.Exception.Message) -Failed
        }

        # --- Send As ---
        try {
            $sendAs = Get-RecipientPermission -Identity $mailboxName | Where-Object { $_.Trustee -like "*@*" -and $_.AccessRights -contains "SendAs" }
            foreach ($perm in $sendAs) {
                $permissionsArray += [PSCustomObject]@{
                    "Mailbox Name"           = $mailboxName
                    "Shared Mailbox Address" = $mailboxEmail
                    "Permission Type"        = "Send As"
                    "User"                   = $perm.Trustee
                }
            }
        } catch {
            Add-ToLog ("⚠️ Failed to read Send As for {0}: {1}" -f $mailboxName, $_.Exception.Message) -Failed
        }

        # --- Send on Behalf ---
        try {
            if ($mailbox.GrantSendOnBehalfTo) {
                foreach ($delegate in $mailbox.GrantSendOnBehalfTo) {
                    try {
                        $delegateRecipient = Get-Recipient -Identity $delegate -ErrorAction SilentlyContinue
                        if ($delegateRecipient -and $delegateRecipient.PrimarySmtpAddress) {
                            $delegateUpn = $delegateRecipient.PrimarySmtpAddress.ToString()
                        }
                        elseif ($delegateRecipient -and $delegateRecipient.WindowsLiveID) {
                            $delegateUpn = $delegateRecipient.WindowsLiveID.ToString()
                        }
                        else {
                            $delegateUpn = $delegate.ToString()
                        }

                        $permissionsArray += [PSCustomObject]@{
                            "Mailbox Name"           = $mailboxName
                            "Shared Mailbox Address" = $mailboxEmail
                            "Permission Type"        = "Send on behalf of"
                            "User"                   = $delegateUpn
                        }
                    } catch {
                        Add-ToLog ("⚠️ Could not resolve delegate {0} for {1}: {2}" -f $delegate, $mailboxName, $_.Exception.Message) -Failed
                    }
                }
            }
        } catch {
            Add-ToLog ("⚠️ Failed to read Send on Behalf for {0}: {1}" -f $mailboxName, $_.Exception.Message) -Failed
        }

        Add-ToLog "✔ Processed mailbox: $mailboxName"
    } # end foreach mailbox

    # ==============================================================
    #   SAVE TO CSV (AFTER LOOP)
    # ==============================================================

    if ($permissionsArray.Count -eq 0) {
        Add-ToLog "⚠️ No permissions found for any shared mailboxes." -Failed
    }
    else {
        $permissionsArray | Export-Csv -Path $ExportPath -Encoding UTF8 -NoTypeInformation -Force
        Add-ToLog "✅ Export complete. File saved to: $ExportPath"
# ✅ Show popup confirming where the export was saved
[System.Windows.Forms.MessageBox]::Show(
    "Microsoft 365 Groups exported successfully to:`n$ExportPath",
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
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    Add-ToLog "🔒 Disconnected from Exchange Online."
    Add-ToLog "📁 Log saved to: $logFile"
}

Add-ToLog "============================================================="
Add-ToLog "🎉 Completed Shared Mailbox Permissions Export process."
Add-ToLog "============================================================="