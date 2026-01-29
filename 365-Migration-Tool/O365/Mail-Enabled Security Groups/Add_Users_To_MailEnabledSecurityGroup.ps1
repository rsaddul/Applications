<#
Developed by: Rhys Saddul
#>

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

# ==============================================================
#   INITIALISE LOGGING
# ==============================================================
$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs\Add_Users_To_MailEnabledSecurityGroup"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$timestamp = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
$successLog = Join-Path $logDir "Add_Users_To_Groups_Success_$timestamp.log"
$failLog    = Join-Path $logDir "Add_Users_To_Groups_Failed_$timestamp.log"

function Add-ToLog {
    param([string]$Message, [switch]$Failed)

    $time = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
    $entry = "[$time] $Message"

    # Output to GUI launcher (stdout)
    $entry

    try {
        if ($Failed) {
            Add-Content -Path $failLog -Value $entry -ErrorAction SilentlyContinue
        } else {
            Add-Content -Path $successLog -Value $entry -ErrorAction SilentlyContinue
        }
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
            "Add_Users_To_Groups_Template.csv"
        )

@"
GroupEmail,UserPrincipalName
MESG_TCS_All_Staff@thecloudschool.co.uk,rsaddul@thecloudschool.co.uk
MESG_TCS_All_Staff@thecloudschool.co.uk,jskett@thecloudschool.co.uk
MESG_TCS_Office_All_Staff@thecloudschool.co.uk,rahemd@thecloudschool.co.uk
"@ | Out-File -FilePath $templatePath -Encoding UTF8 -Force

        [System.Windows.Forms.MessageBox]::Show(
            "Template generated successfully:`n$templatePath",
            "Template Created",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        Add-ToLog "✅ Generated CSV template at: $templatePath"
        Add-ToLog "💡 You can now edit the CSV to add real users and group emails."
        exit 0
    }
    catch {
        Add-ToLog "⚠️ Failed to create template: $($_.Exception.Message)" -Failed
        exit 1
    }
}

# ==============================================================
#   FILE PICKER
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
    exit 1
}

# ==============================================================
#   IMPORT & PROCESS CSV
# ==============================================================
try {
    $rows = Import-Csv -Path $CSVPath
    if (-not $rows) {
        Add-ToLog "⚠️ No rows found in CSV." -Failed
        exit 1
    }

    $successCount = 0
    $failCount = 0
    $total = $rows.Count
    $counter = 0

    foreach ($row in $rows) {
        $counter++

        try {
            $groupEmail = $row.GroupEmail
            $userEmail  = $row.UserPrincipalName

            if ([string]::IsNullOrWhiteSpace($groupEmail) -or [string]::IsNullOrWhiteSpace($userEmail)) {
                Add-ToLog "⚠️ [$counter/$total] Skipping invalid row (missing GroupEmail or UserPrincipalName)." -Failed
                continue
            }

            Add-ToLog "🔎 [$counter/$total] Processing: Group=$groupEmail, User=$userEmail..."

            # Validate group
            $group = Get-DistributionGroup -Identity $groupEmail -ErrorAction SilentlyContinue
            if (-not $group) {
                Add-ToLog "❌ Group not found: $groupEmail" -Failed
                $failCount++
                continue
            }

            # Validate user
            $user = Get-Recipient -Identity $userEmail -ErrorAction SilentlyContinue
            if (-not $user) {
                Add-ToLog "❌ User not found: $userEmail" -Failed
                $failCount++
                continue
            }

            # Check if already member
            $members = Get-DistributionGroupMember -Identity $groupEmail -ErrorAction SilentlyContinue
            if ($members.PrimarySmtpAddress -contains $userEmail) {
                Add-ToLog "⚠️ $userEmail is already a member of $groupEmail."
                continue
            }

            # Add to group
            Add-DistributionGroupMember -Identity $groupEmail -Member $userEmail -ErrorAction Stop
            Add-ToLog "✅ Added $userEmail to group $groupEmail."
            $successCount++
        }
        catch {
            Add-ToLog "❌ Failed to add $userEmail to $($groupEmail): $($_.Exception.Message)" -Failed
            $failCount++
        }
    }

    Add-ToLog "🔹 Summary: $successCount success(es), $failCount failure(s)."
    Add-ToLog "📁 Logs saved to:`n   $successLog`n   $failLog"
}
catch {
    Add-ToLog "❌ Fatal error: $($_.Exception.Message)" -Failed
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    Add-ToLog "🔒 Disconnected from Exchange Online."
}

Add-ToLog "🎉 Completed Add Users to Groups process."