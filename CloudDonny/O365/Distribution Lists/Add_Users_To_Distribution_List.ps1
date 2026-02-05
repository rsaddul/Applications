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

# --- Initialise dual log files ---
$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs\Add_Users_To_Distribution_List"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$timestamp = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
$successLog = Join-Path $logDir "Add_Users_To_Distribution_List_Success_$timestamp.log"
$failLog    = Join-Path $logDir "Add_Users_To_Distribution_List_Failed_$timestamp.log"

# --- Inline log helper ---
function Add-ToLog {
    param([string]$Message, [switch]$Failed)
    $time  = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
    $entry = "[$time] $Message"

    # --- Output to GUI launcher ---
    $entry 

    # --- Also write to log file ---
    try {
        if ($Failed) {
            Add-Content -Path $failLog -Value $entry -ErrorAction SilentlyContinue
        } else {
            Add-Content -Path $successLog -Value $entry -ErrorAction SilentlyContinue
        }
    } catch {}
}

# --- Ask if user wants to generate a template ---
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
            "Add_Users_To_Distribution_List_Template.csv"
        )

        @"
DistributionList,UserPrincipalName
DL_TCS_All_Staff@thecloudschool.co.uk,rsaddul@thecloudschool.co.uk
DL_TCS_All_Staff@thecloudschool.co.uk,jskett@thecloudschool.co.uk
DL_TCS_Office_All_Staff@thecloudschool.co.uk,rahemd@thecloudschool.co.uk
"@ | Out-File -FilePath $templatePath -Encoding UTF8 -Force
# ✅ Show popup confirming where the template was saved
[System.Windows.Forms.MessageBox]::Show(
    "Template generated successfully:`n$templatePath",
    "Template Created",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)
        Add-ToLog "âœ… Generated CSV template at: $templatePath"
        Add-ToLog "ðŸ’¡ You can now open and edit the CSV to add real users and distribution lists."
        exit 0
    }
    catch {
        Add-ToLog "âš ï¸ Failed to create template: $($_.Exception.Message)" -Failed
        exit 1
    }
}

# --- File picker for existing CSV file ---
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Select your CSV file"
$OpenFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
$OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $CSVPath = $OpenFileDialog.FileName
    Add-ToLog "ðŸ“„ Using CSV file: $CSVPath"
} else {
    Add-ToLog "âŒ Cancelled: CSV file is required." -Failed
    exit 1
}

if (-not (Test-Path $CSVPath)) {
    Add-ToLog "âŒ CSV file not found: $CSVPath" -Failed
    exit 1
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

# --- Import and Process CSV ---
try {
    $rows = Import-Csv -Path $CSVPath
    if (-not $rows) {
        Add-ToLog "⚠️ No rows found in CSV." -Failed
        exit 1
    }

    $successCount = 0
    $failCount = 0

    foreach ($row in $rows) {
        try {
            $distributionList = $row.DistributionList
            $user = $row.UserPrincipalName

            if ([string]::IsNullOrWhiteSpace($distributionList) -or [string]::IsNullOrWhiteSpace($user)) {
                Add-ToLog "⚠️ Skipping invalid row (missing DistributionList or UserPrincipalName)." -Failed
                continue
            }

            # Check if DL exists, create if not
            $dl = Get-DistributionGroup -Identity $distributionList -ErrorAction SilentlyContinue
            if (-not $dl) {
                try {
                    Add-ToLog "📦 Creating new distribution list: $distributionList..."
                    New-DistributionGroup -Name $distributionList -PrimarySmtpAddress $distributionList -Type Distribution -ErrorAction Stop | Out-Null
                    Add-ToLog "✅ Created new distribution list: $distributionList."
                    Start-Sleep -Seconds 3
                }
                catch {
                    Add-ToLog "❌ Failed to create distribution list ${distributionList}: $($_.Exception.Message)" -Failed
                    $failCount++
                    continue
                }
            }

            # Add member
            Add-DistributionGroupMember -Identity $distributionList -Member $user -ErrorAction Stop
            Add-ToLog "✅ Added $user to distribution list $distributionList."
            $successCount++
        }
        catch {
            $msg = $_.Exception.Message
            if ($msg -like "*is already a member*") {
                Add-ToLog "⚠️ $user is already a member of $distributionList." -Failed
            }
            elseif ($msg -like "*Couldn't find object*" -or $msg -like "*not found*") {
                Add-ToLog "❌ User or list not found: $user / $distributionList." -Failed
            }
            else {
                Add-ToLog "❌ Failed to add ${user} to ${distributionList}: ${msg}" -Failed
            }
            $failCount++
        }
    }

    Add-ToLog "🔹 Summary: $successCount success(es), $failCount failure(s)."
} # 
catch {
    Add-ToLog "❌ Fatal error: $($_.Exception.Message)" -Failed
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    Add-ToLog "🔒 Disconnected from Exchange Online."
    Add-ToLog "📁 Logs saved to:`n   $successLog`n   $failLog"
}

Add-ToLog "🎉 Completed Add Users to Distribution List process."
