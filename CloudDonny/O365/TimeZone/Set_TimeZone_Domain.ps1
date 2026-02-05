<#
Developed by: Rhys Saddul

Overview:
This script sets all mailboxes in a specified domain (entered via prompt)
to the desired TimeZone, DateFormat, and Language.

Notes:
- Uses the simpler and stable VB InputBox method (no STA or Win32 calls).
- Fully compatible with CloudDonny GUI launcher.
- Includes logging, progress feedback, and throttling to avoid session freezes.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

# ==============================================================
#   INITIALISE LOGGING
# ==============================================================

$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs\Set_TimeZone"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$timestamp = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
$successLog = Join-Path $logDir "Set_Mailbox_Regional_Success_$timestamp.log"
$failLog    = Join-Path $logDir "Set_Mailbox_Regional_Failed_$timestamp.log"

function Add-ToLog {
    param(
        [string]$Message,
        [switch]$Failed
    )

    $time  = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
    $entry = "[$time] $Message"

    # Output for GUI (CloudDonny launcher)
    $entry

    # Log to file
    try {
        if ($Failed) {
            Add-Content -Path $failLog -Value $entry -ErrorAction SilentlyContinue
        } else {
            Add-Content -Path $successLog -Value $entry -ErrorAction SilentlyContinue
        }
    } catch {}
}

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
#   PROMPT FOR TARGET DOMAIN
# ==============================================================

$Domain = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter the domain to target (e.g., thecloudschool.co.uk):",
    "Mailbox Domain Filter",
    ""
)
if ([string]::IsNullOrWhiteSpace($Domain)) {
    Add-ToLog "❌ Cancelled: No domain provided." -Failed
    exit 1
}

Add-ToLog "🎯 Targeting mailboxes in domain: $Domain"

# ==============================================================
#   SET REGIONAL SETTINGS FOR MAILBOXES
# ==============================================================

$Region = "GMT Standard Time"
$DateFormat = "dd/MM/yyyy"
$Language = "en-GB"

Add-ToLog "🕒 Applying calendar and regional configurations for all mailboxes in $Domain..."

try {
    $mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object { $_.PrimarySmtpAddress -like "*@$Domain" }

    if (-not $mailboxes) {
        Add-ToLog "⚠️ No mailboxes found in domain $Domain." -Failed
        exit 1
    }

    $count = 0
    $failCount = 0
    $total = $mailboxes.Count

    foreach ($mailbox in $mailboxes) {
        $count++
        $UPN = $mailbox.PrimarySmtpAddress

        Write-Progress -Activity "Updating mailbox regional settings" -Status "$count of $total" -PercentComplete (($count / $total) * 100)

        try {
            # --- Set Calendar Configuration ---
            Set-MailboxCalendarConfiguration -Identity $UPN -WorkingHoursTimeZone $Region -ErrorAction Stop
            Add-ToLog "✅ Set calendar configuration for mailbox: $UPN"

            # --- Set Regional Configuration ---
            Set-MailboxRegionalConfiguration -Identity $UPN -TimeZone $Region -Language $Language -DateFormat $DateFormat -Confirm:$false -ErrorAction Stop
            Add-ToLog "✅ Set regional configuration for mailbox: $UPN"
        }
        catch {
            Add-ToLog ("❌ Failed to update one or more settings for {0}: {1}" -f $UPN, $_.Exception.Message) -Failed
            $failCount++
        }

        Start-Sleep -Milliseconds 250  # Prevent throttling
    }

    Add-ToLog "🔹 Summary: $count processed, $failCount failure(s)."
# ✅ Show popup summary to user
[System.Windows.Forms.MessageBox]::Show(
    "Mailbox timezone and regional settings updated for domain:`n$Domain`n`nSummary:`n$count processed`n$failCount failure(s)`n`nLogs saved to:`n$logDir",
    "Operation Complete",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)

}
catch {
    Add-ToLog "❌ Fatal error while processing mailboxes: $($_.Exception.Message)" -Failed
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    Add-ToLog "🔒 Disconnected from Exchange Online."
    Add-ToLog "📁 Logs saved to:`n   $successLog`n   $failLog"
}

Add-ToLog "🎉 Completed Set Mailbox Regional Configuration process."
