<#
Developed by: Rhys Saddul
#>

# ==============================================================
#   ENSURE SCRIPT RUNS IN STA MODE (for GUI + ExchangeOnline)
# ==============================================================

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
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

$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs\Create_Distribution_Lists"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$timestamp = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
$successLog = Join-Path $logDir "Create_Distribution_List_Success_$timestamp.log"
$failLog    = Join-Path $logDir "Create_Distribution_List_Failed_$timestamp.log"

function Add-ToLog {
    param(
        [string]$Message,
        [switch]$Failed
    )

    $time  = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
    $entry = "[$time] $Message"

    # --- Output to GUI launcher (plain string output works with CloudDonny GUI) ---
    $entry

    # --- Also log to file ---
    try {
        if ($Failed) {
            Add-Content -Path $failLog -Value $entry -ErrorAction SilentlyContinue
        } else {
            Add-Content -Path $successLog -Value $entry -ErrorAction SilentlyContinue
        }
    } catch {}
}

# ==============================================================
#   ASK TO GENERATE CSV TEMPLATE
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
            "Create_Distribution_List_Template.csv"
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
        Add-ToLog "✅ Generated CSV template at: $templatePath"
        Add-ToLog "💡 You can now open and edit the CSV to add real users and distribution lists."
        exit 0
    }
    catch {
        Add-ToLog "⚠️ Failed to create template: $($_.Exception.Message)" -Failed
        exit 1
    }
}

# ==============================================================
#   FILE PICKER FOR EXISTING CSV FILE
# ==============================================================

$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Select your CSV file"
$OpenFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
$OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $CSVPath = $OpenFileDialog.FileName
    Add-ToLog "📄 Using CSV file: $CSVPath"
} else {
    Add-ToLog "❌ Cancelled: CSV file is required." -Failed
    exit 1
}

if (-not (Test-Path $CSVPath)) {
    Add-ToLog "❌ CSV file not found: $CSVPath" -Failed
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

# ==============================================================
#   IMPORT AND PROCESS CSV DATA (Create & Add Members)
# ==============================================================

try {
    $rows = Import-Csv -Path $CSVPath

    if (-not $rows) {
        Add-ToLog "⚠️ No rows found in CSV." -Failed
        exit 1
    }

    $successCount = 0
    $failCount = 0

    foreach ($row in $rows) {
        $distributionList = $row.DistributionList
        $user = $row.UserPrincipalName

        if ([string]::IsNullOrWhiteSpace($distributionList) -or [string]::IsNullOrWhiteSpace($user)) {
            Add-ToLog "⚠️ Skipping invalid row (missing DistributionList or UserPrincipalName)." -Failed
            continue
        }

        # --- Ensure DL exists ---
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

        # --- Add member to DL ---
        try {
            Add-DistributionGroupMember -Identity $distributionList -Member $user -ErrorAction Stop
            Add-ToLog "✅ Added $user to $distributionList."
            $successCount++
        }
        catch {
            $errorMessage = $_.Exception.Message

            switch -Wildcard ($errorMessage) {
                "*is already a member*" {
                    Add-ToLog "⚠️ $user is already a member of $distributionList." -Failed
                }
                "*Couldn't find object*" { 
                    Add-ToLog "❌ User or list not found: $user / $distributionList." -Failed
                }
                "*not found*" {
                    Add-ToLog "❌ User or list not found: $user / $distributionList." -Failed
                }
                default {
                    Add-ToLog "❌ Failed to add ${user} to ${distributionList}: ${errorMessage}" -Failed
                }
            }
            $failCount++
        }
    }

    Add-ToLog "🔹 Summary: $successCount success(es), $failCount failure(s)."
}
catch {
    Add-ToLog "❌ Fatal error: $($_.Exception.Message)" -Failed
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    Add-ToLog "🔒 Disconnected from Exchange Online."
    Add-ToLog "📁 Logs saved to:`n   $successLog`n   $failLog"
}

Add-ToLog "🎉 Completed Add Users to Distribution List process."
