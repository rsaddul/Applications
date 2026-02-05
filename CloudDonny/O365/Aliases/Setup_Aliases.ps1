<#
Developed by: Rhys Saddul
#>

# ==============================================================
#   ENSURE SCRIPT RUNS IN STA MODE (for GUI + ExchangeOnline)
# ==============================================================

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Write-Host "[STA Check] Relaunching script in STA mode..."
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell.exe"
    $psi.Arguments = "-STA -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Definition)`""
    $psi.UseShellExecute = $false
    [System.Diagnostics.Process]::Start($psi) | Out-Null
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

# ==============================================================
#   INITIALISE LOGGING
# ==============================================================

$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$timestamp = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
$successLog = Join-Path $logDir "Setup_Aliases_Success_$timestamp.log"
$failLog    = Join-Path $logDir "Setup_Aliases_Failed_$timestamp.log"

function Add-ToLog {
    param(
        [string]$Message,
        [switch]$Failed
    )
    $time = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
    $entry = "[$time] $Message"
    $entry
    try {
        if ($Failed) { Add-Content $failLog $entry -ErrorAction SilentlyContinue }
        else { Add-Content $successLog $entry -ErrorAction SilentlyContinue }
    } catch {}
}

# ==============================================================
#   ASK TO GENERATE TEMPLATE
# ==============================================================

$answer = [System.Windows.Forms.MessageBox]::Show(
    "Would you like to generate a CSV template with example data?",
    "Generate Template",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
    try {
        $templatePath = [System.IO.Path]::Combine(
            [Environment]::GetFolderPath("Desktop"),
            "Setup_Aliases_Template.csv"
        )

@"
CurrentEmail,PrimarySMTP,Alias1,Alias2,Alias3
rsaddul@thecloudschool.co.uk,rsaddul@eduthing-azure.co.uk,rsaddul@thecloudtrust.co.uk,rhys.saddul@thecloudschool.co.uk,
jskett@thecloudschool.co.uk,jskett@eduthing-azure.co.uk,jskett@thecloudtrust.co.uk,james.skett@thecloudschool.co.uk
"@ | Out-File -FilePath $templatePath -Encoding UTF8 -Force

# Show popup confirming where the template was saved
[System.Windows.Forms.MessageBox]::Show(
    "Template generated successfully:`n$templatePath",
    "Template Created",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)
        Add-ToLog "Generated CSV template at: $templatePath"
        Add-ToLog "You can now open and edit the CSV to add real user aliases."
        exit 0
    }
    catch {
        Add-ToLog "Failed to create template: $($_.Exception.Message)" -Failed
        exit 1
    }
}

# ==============================================================
#   FILE PICKER FOR EXISTING CSV FILE
# ==============================================================

$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Select your Aliases CSV file"
$OpenFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
$OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $CSVPath = $OpenFileDialog.FileName
    Add-ToLog "Using CSV file: $CSVPath"
} else {
    Add-ToLog "Cancelled: CSV file is required." -Failed
    exit 1
}

if (-not (Test-Path $CSVPath)) {
    Add-ToLog "CSV file not found: $CSVPath" -Failed
    exit 1
}

# ==============================================================
#   CONNECT TO EXCHANGE ONLINE AND GRAPH
# ==============================================================

try {
    Add-ToLog "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowProgress $true -ShowBanner:$false
    Add-ToLog "Connected to Exchange Online."
}
catch {
    Add-ToLog "Failed to connect to Exchange Online: $($_.Exception.Message)" -Failed
    [System.Windows.Forms.MessageBox]::Show(
        "Connection failed:`n$($_.Exception.Message)",
        "Exchange Online Connection Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}


try {
    Connect-MgGraph -Scopes "User.ReadWrite.All" | Out-Null
    Add-ToLog "Connected to Microsoft Graph."
} catch {
    Add-ToLog "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Failed
}

# ==============================================================
#   IMPORT AND PROCESS CSV DATA
# ==============================================================

try {
    $rows = Import-Csv -Path $CSVPath
    if (-not $rows) {
        Add-ToLog "No rows found in CSV." -Failed
        exit 1
    }

    $successCount = 0
    $failCount = 0

    foreach ($row in $rows) {
        try {
            $identity   = $row.CurrentEmail
            $newPrimary = $row.PrimarySMTP

            if (-not $identity -or -not $newPrimary) {
                Add-ToLog "Skipping invalid row (missing CurrentEmail or PrimarySMTP)." -Failed
                continue
            }

        $mbx = Get-Mailbox -Identity $identity -ErrorAction Stop
    
    # Gather existing aliases (ignore current primary address)
    $existing = @()
        foreach ($addr in $mbx.EmailAddresses) {
            $s = $addr.ToString().ToLower()
            if ($s -notmatch '^smtp:' -and $s -notmatch "^smtp:$($identity.ToLower())") {
                $existing += $s
            }
        }

    # Gather aliases from CSV columns (Alias1, Alias2, etc.)
    $csvAliases = @()
        foreach ($p in $row.PSObject.Properties.Name) {
            if ($p -match '^Alias\d+$') {
                $val = $row.$p
            if ($val -and $val.Trim() -ne '') {
                $csvAliases += ('smtp:' + $val.Trim().ToLower())
            }
        }
    }

        # Always ensure the old address becomes an alias
        $oldAsAlias = "smtp:$($identity.ToLower())"

        # Combine new primary, existing aliases, new CSV aliases, and the old primary
        $newList = @("SMTP:$newPrimary") + $existing + $csvAliases + $oldAsAlias
    
        # Remove duplicates safely
        $newList = $newList | Select-Object -Unique

            Set-Mailbox -Identity $identity -EmailAddresses $newList -WindowsEmailAddress $newPrimary
            Update-MgUser -UserId $identity -UserPrincipalName $newPrimary
            Add-ToLog "Updated $identity → $newPrimary; Aliases: $((($newList | Where-Object { $_ -match '^smtp:' }) -join ', '))"
            $successCount++
        }
        catch {
            Add-ToLog "Failed for $($row.CurrentEmail): $($_.Exception.Message)" -Failed
            $failCount++
        }
    }

    Add-ToLog "Summary: $successCount success(es), $failCount failure(s)."
}
catch {
    Add-ToLog "Fatal script error: $($_.Exception.Message)" -Failed
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    Disconnect-MgGraph | Out-Null
    Add-ToLog "Disconnected from Exchange Online and Graph."
    Add-ToLog "Logs saved to:`n   $successLog`n   $failLog"
}

Add-ToLog "Completed Setup Aliases process."
