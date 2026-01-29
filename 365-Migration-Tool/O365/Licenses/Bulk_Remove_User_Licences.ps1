<#
Developed by: Rhys Saddul
#>

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

# ==============================================================
#   INITIALISE LOGGING
# ==============================================================

$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs\Bulk_Remove_User_Licences"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}

$timestamp  = Get-Date -Format "dd-MM-yyyy_HH-mm-ss"
$successLog = Join-Path $logDir "Remove_Licenses_Success_$timestamp.log"
$failLog    = Join-Path $logDir "Remove_Licenses_Failed_$timestamp.log"

function Add-ToLog {
    param (
        [string]$Message,
        [switch]$Failed
    )

    $time  = Get-Date -Format "dd/MM/yyyy - HH:mm:ss"
    $entry = "[$time] $Message"
    $entry

    try {
        if ($Failed) {
            Add-Content -Path $failLog -Value $entry
        } else {
            Add-Content -Path $successLog -Value $entry
        }
    } catch {}
}

# ==============================================================
#   ASK TO GENERATE TEMPLATE
# ==============================================================

$answer = [System.Windows.Forms.MessageBox]::Show(
    "Would you like to generate a CSV template?",
    "Generate Template",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
    try {
        $templatePath = Join-Path ([Environment]::GetFolderPath("Desktop")) "Remove_Licenses_Template.csv"

@"
UserPrincipalName
rsaddul@thecloudschool.co.uk
jskett@thecloudschool.co.uk
"@ | Out-File -Path $templatePath -Encoding UTF8 -Force

        [System.Windows.Forms.MessageBox]::Show(
            "Template created:`n$templatePath",
            "Success",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        Add-ToLog "CSV template generated at $templatePath"
        exit 0
    }
    catch {
        Add-ToLog "Failed to create template: $($_.Exception.Message)" -Failed
        exit 1
    }
}

# ==============================================================
#   FILE PICKER
# ==============================================================

$ofd = New-Object System.Windows.Forms.OpenFileDialog
$ofd.Title = "Select CSV file"
$ofd.Filter = "CSV files (*.csv)|*.csv"
$ofd.InitialDirectory = [Environment]::GetFolderPath("Desktop")

if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Add-ToLog "CSV selection cancelled." -Failed
    exit 1
}

$CSVPath = $ofd.FileName
Add-ToLog "Using CSV: $CSVPath"

# ==============================================================
#   CONNECT TO MICROSOFT GRAPH
# ==============================================================

try {
    Add-ToLog "Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.Read.All" -NoWelcome | Out-Null
    Add-ToLog "Connected to Microsoft Graph"
}
catch {
    Add-ToLog "Graph connection failed: $($_.Exception.Message)" -Failed
    exit 1
}

# ==============================================================
#   PROCESS CSV
# ==============================================================

try {
    $users = Import-Csv $CSVPath
    if (-not $users) {
        Add-ToLog "CSV contains no data." -Failed
        exit 1
    }

$sku = Get-MgSubscribedSku -All |
    Select-Object `
        SkuPartNumber,
        @{ Name = 'Assigned'; Expression = { $_.ConsumedUnits } },
        @{ Name = 'Available'; Expression = { $_.PrepaidUnits.Enabled - $_.ConsumedUnits } },
        SkuId |
    Sort-Object SkuPartNumber |
    Out-GridView -Title "Select license to REMOVE (Assigned / Available shown)" -PassThru

    if (-not $sku) {
        Add-ToLog "No license selected." -Failed
        exit 1
    }

    Add-ToLog "Selected license: $($sku.SkuPartNumber)"

    $total   = $users.Count
    $count   = 0
    $success = 0
    $fails   = 0

    foreach ($user in $users) {
        $count++
        $upn = $user.UserPrincipalName

        if ([string]::IsNullOrWhiteSpace($upn)) {
            Add-ToLog "[$count/$total] Empty UPN row skipped." -Failed
            continue
        }

        try {
            $licenseDetails = Get-MgUserLicenseDetail -UserId $upn -ErrorAction Stop
            $targetLicense  = $licenseDetails | Where-Object { $_.SkuId -eq $sku.SkuId }

            if (-not $targetLicense) {
                Add-ToLog "[$count/$total] $upn does not have $($sku.SkuPartNumber). Skipping."
                continue
            }

            if ($targetLicense.GroupsAssigningLicense.Count -gt 0) {
                Add-ToLog "[$count/$total] $upn has GROUP-assigned license. Cannot remove." -Failed
                continue
            }

            Add-ToLog "[$count/$total] Removing license from $upn..."
            Set-MgUserLicense -UserId $upn -AddLicenses @() -RemoveLicenses @($sku.SkuId) -ErrorAction Stop

            Add-ToLog "Removed $($sku.SkuPartNumber) from $upn"
            $success++
        }
        catch {
            Add-ToLog "[$count/$total] Failed for ${upn}: $($_.Exception.Message)" -Failed
            $fails++
        }
    }

    Add-ToLog "SUMMARY: $success succeeded | $fails failed"
    Add-ToLog "Logs:`n$successLog`n$failLog"
}
finally {
    Disconnect-MgGraph | Out-Null
    Add-ToLog "Disconnected from Microsoft Graph"
}

Add-ToLog "Bulk license removal completed"
