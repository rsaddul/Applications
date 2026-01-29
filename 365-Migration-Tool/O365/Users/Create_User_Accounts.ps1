<#
Developed by: Rhys Saddul
#>

param (
    [string]$CsvPath,
    [switch]$ForcePasswordChange
)

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

# ==============================================================
#   INITIALISE LOGGING
# ==============================================================

$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs\Create_User_Accounts"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$timestamp  = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
$successLog = Join-Path $logDir "Create_Users_Success_$timestamp.log"
$failLog    = Join-Path $logDir "Create_Users_Failed_$timestamp.log"

function Add-ToLog {
    param([string]$Message, [switch]$Failed)

    $time = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
    $entry = "[$time] $Message"
    [Console]::Error.WriteLine($entry)

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
        $templatePath = Join-Path ([Environment]::GetFolderPath("Desktop")) "UserCreation_Template.csv"

        @"
Name [displayName] Required,User name [userPrincipalName] Required,Initial password [passwordProfile] Required,FirstName,LastName,Department,City,Country,State,PostalCode,StreetAddress,MobilePhone,mailNickName
Rhys Saddul,rsaddul@thecloudschool.co.uk,P@ssword123!,Rhys,Saddul,IT,London,GB,England,E14 3AA,123 Cloud Road,07123456789,rsaddul
James Skett,jskett@thecloudschool.co.uk,P@ssword123!,James,Skett,HR,London,GB,England,E14 3AA,124 Cloud Road,07987654321,jskett
"@ | Out-File -FilePath $templatePath -Encoding UTF8 -Force

        [System.Windows.Forms.MessageBox]::Show(
            "Template generated successfully:`n$templatePath",
            "Template Created",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        Add-ToLog "✅ Generated CSV template at: $templatePath"
        Add-ToLog "💡 You can now edit this file with real user data."
        exit 0
    }
    catch {
        Add-ToLog "⚠️ Failed to create template: $($_.Exception.Message)" -Failed
        exit 1
    }
}

# ==============================================================
#   PROMPT FOR CSV PATH
# ==============================================================

if (-not $CsvPath) {
    $CsvPath = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Enter the full path to your CSV file (e.g. C:\Users\You\Desktop\UserCreation.csv):",
        "CSV Path Input",
        (Join-Path ([Environment]::GetFolderPath("Desktop")) "UserCreation.csv")
    )
}

if ([string]::IsNullOrWhiteSpace($CsvPath) -or -not (Test-Path $CsvPath)) {
    Add-ToLog "❌ Cancelled: CSV file is required." -Failed
    exit 1
}

Add-ToLog "📄 Using CSV file: $CsvPath"

# ==============================================================
#   PROMPT FOR FORCE PASSWORD CHANGE
# ==============================================================

if (-not $PSBoundParameters.ContainsKey("ForcePasswordChange")) {
    $choice = [System.Windows.Forms.MessageBox]::Show(
        "Force password change at next sign-in?",
        "Password Options",
        [System.Windows.Forms.MessageBoxButtons]::YesNo
    )
    $ForcePasswordChange = $choice -eq [System.Windows.Forms.DialogResult]::Yes
}

# ==============================================================
#   CONNECT TO MICROSOFT GRAPH
# ==============================================================

try {
    Add-ToLog "🌐 Connecting to Microsoft Graph..."
    [System.Environment]::SetEnvironmentVariable("MSAL_DISABLE_WAM", "1", "Process")
    Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.ReadWrite.All" -NoWelcome | Out-Null
    Add-ToLog "✅ Connected to Microsoft Graph."
}
catch {
    Add-ToLog "❌ Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Failed
    exit 1
}

# ==============================================================
#   IMPORT AND PROCESS CSV DATA
# ==============================================================

try {
    $usersData = Import-Csv -Path $CsvPath
    if (-not $usersData) { Add-ToLog "⚠️ No users found in CSV." -Failed; exit 1 }

    $total = $usersData.Count
    $successCount = 0
    $failCount = 0
    $i = 0

    foreach ($userRow in $usersData) {
        $i++
        $upn = $userRow.'User name [userPrincipalName] Required'
        $pwd = $userRow.'Initial password [passwordProfile] Required'
        $display = $userRow.'Name [displayName] Required'

        if (-not $upn -or -not $pwd -or -not $display) {
            Add-ToLog "⚠️ [$i/$total] Missing required fields — skipping." -Failed
            continue
        }

        # Check if user already exists
        if (Get-MgUser -UserId $upn -ErrorAction SilentlyContinue) {
            Add-ToLog "⚠️ [$i/$total] $upn already exists — skipping." -Failed
            continue
        }

        $userParams = @{
            DisplayName       = $display
            UserPrincipalName = $upn
            PasswordProfile   = @{
                Password                      = $pwd
                ForceChangePasswordNextSignIn = $ForcePasswordChange
            }
            AccountEnabled = $true
            MailNickName   = if ($userRow.mailNickName) { $userRow.mailNickName } else { ($upn -split "@")[0] }
        }

        if ($userRow.FirstName)    { $userParams.GivenName    = $userRow.FirstName }
        if ($userRow.LastName)     { $userParams.Surname      = $userRow.LastName }
        if ($userRow.Department)   { $userParams.Department   = $userRow.Department }
        if ($userRow.City)         { $userParams.City         = $userRow.City }
        if ($userRow.Country)      { $userParams.Country      = if ($userRow.Country -eq "UK") { "GB" } else { $userRow.Country } }
        if ($userRow.State)        { $userParams.State        = $userRow.State }
        if ($userRow.PostalCode)   { $userParams.PostalCode   = $userRow.PostalCode }
        if ($userRow.StreetAddress){ $userParams.StreetAddress= $userRow.StreetAddress }
        if ($userRow.MobilePhone)  { $userParams.MobilePhone  = $userRow.MobilePhone }

        try {
            Add-ToLog "👤 [$i/$total] Creating user: $upn..."
            New-MgUser @userParams | Out-Null
            Add-ToLog "✅ Created user $upn"
            $successCount++
        }
        catch {
            Add-ToLog "❌ [$i/$total] Error creating $upn : $($_.Exception.Message)" -Failed
            $failCount++
        }
    }

    Add-ToLog "🔹 Summary: $successCount success(es), $failCount failure(s)."

# ✅ Show popup summary to user
[System.Windows.Forms.MessageBox]::Show(
    "User creation completed successfully.`n`nSummary:`n$successCount created`n$failCount failed`n`nLogs saved to:`n$logDir",
    "Operation Complete",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)

    Add-ToLog "📁 Logs saved to:`n   $successLog`n   $failLog"
}
catch {
    Add-ToLog "❌ Fatal error: $($_.Exception.Message)" -Failed
}
finally {
    Disconnect-MgGraph | Out-Null
    Add-ToLog "🔒 Disconnected from Microsoft Graph."
}

Add-ToLog "🎉 Completed user creation process."
