<#
Developed by: Rhys Saddul
#>

param (
    [string]$SchoolCode,
    [string]$TrustCode = "",
    [ValidateSet("Trust","School")]
    [string]$SetupType
)

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

# --- Prompt for SchoolCode if not supplied ---
if (-not $SchoolCode) {
    $SchoolCode = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Enter School Code (Required):", "School Code", ""
    )
    if ([string]::IsNullOrWhiteSpace($SchoolCode)) {
        "❌ Cancelled: SchoolCode is required."
        exit 1
    }
}

# --- Prompt for SetupType if not supplied ---
if (-not $SetupType) {
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Is this a Trust setup? (Yes = Trust, No = School, Cancel = Abort)",
        "Setup Type",
        [System.Windows.Forms.MessageBoxButtons]::YesNoCancel
    )
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) { exit }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $SetupType = "Trust"
        if (-not $TrustCode) {
            $TrustCode = [Microsoft.VisualBasic.Interaction]::InputBox(
                "Enter Trust Code (Optional):", "Trust Code", ""
            )
        }
    } else {
        $SetupType = "School"
    }
}

# --- Validation ---
if ([string]::IsNullOrWhiteSpace($SchoolCode)) {
    "❌ SchoolCode is required."
    exit 1
}
if ([string]::IsNullOrWhiteSpace($SetupType)) {
    "❌ SetupType must be Trust or School."
    exit 1
}

"🔹 Running Mac Groups setup..."
"   SchoolCode : $SchoolCode"
if ($TrustCode) { "   TrustCode  : $TrustCode" }
"   SetupType  : $SetupType"

# --- Connect to Microsoft Graph ---
Connect-MgGraph -Scopes "Group.ReadWrite.All","Directory.ReadWrite.All" | Out-Null

# --- Helper Functions ---
function Get-GroupByName($displayName) {
    try {
        return Get-MgGroup -Filter "displayName eq '$displayName'" `
            -ConsistencyLevel eventual -ErrorAction Stop
    } catch { return $null }
}
function GroupExists($displayName) { return (Get-GroupByName $displayName) -ne $null }
function MemberExists($groupId, $memberId) {
    try { (Get-MgGroupMember -GroupId $groupId -All).Id -contains $memberId } catch { $false }
}
function AddMemberIfNotExists($groupId, $memberId, $relation) {
    if (-not $groupId -or -not $memberId) { return }
    if (MemberExists $groupId $memberId) {
        "⚠️ Nesting already exists: $relation"
    } else {
        New-MgGroupMember -GroupId $groupId -DirectoryObjectId $memberId
        "✅ Added nesting: $relation"
    }
}

# --- Static Groups ---
$staticGroups = @()
if ($SetupType -in @("Trust","School")) {
    $staticGroups += @(
        "[EDU v2 Security] - $SchoolCode All Corporate Mac Devices",
        "[EDU v2 Security] - $SchoolCode All Corporate Shared Mac Devices",
        "[EDU v2 Security] - $SchoolCode All Corporate Personal Mac Devices",
        "[EDU v2 Security] - $SchoolCode Student Shared Mac Enrolment",
        "[EDU v2 Security] - $SchoolCode Student Personal Mac Enrolment",
        "[EDU v2 Security] - $SchoolCode Staff Shared Mac Enrolment",
        "[EDU v2 Security] - $SchoolCode Staff Personal Mac Enrolment"
    )
}
if ($SetupType -eq "Trust" -and $TrustCode) {
    $staticGroups += @(
        "[EDU v2 Security] - $TrustCode All Corporate Mac Devices",
        "[EDU v2 Security] - $TrustCode All Corporate Shared Mac Devices",
        "[EDU v2 Security] - $TrustCode All Corporate Shared Staff Mac Devices",
        "[EDU v2 Security] - $TrustCode All Corporate Shared Student Mac Devices",
        "[EDU v2 Security] - $TrustCode All Corporate Personal Mac Devices",
        "[EDU v2 Security] - $TrustCode All Corporate Personal Staff Mac Devices",
        "[EDU v2 Security] - $TrustCode All Corporate Personal Student Mac Devices"
    )
}

foreach ($groupName in $staticGroups) {
    if (GroupExists $groupName) {
        "⚠️ Group '$groupName' already exists. Skipping."
        continue
    }
    $newGroup = New-MgGroup -DisplayName $groupName -SecurityEnabled:$true -MailEnabled:$false `
        -MailNickname ("grp" + (Get-Random)) -GroupTypes @()
    "✅ Static Group '$groupName' created."
}

# --- Dynamic Groups ---
$dynamicGroups = @()
if ($SetupType -in @("Trust","School")) {
    $dynamicGroups += @(
        "[EDU v2 Security] - $SchoolCode All Staff Corporate Personal Mac Devices",
        "[EDU v2 Security] - $SchoolCode All Student Corporate Personal Mac Devices",
        "[EDU v2 Security] - $SchoolCode All Staff Corporate Shared Mac Devices",
        "[EDU v2 Security] - $SchoolCode All Student Corporate Shared Mac Devices"
    )
}

foreach ($dynamicGroupName in $dynamicGroups) {
    if (GroupExists $dynamicGroupName) {
        "⚠️ Dynamic Group '$dynamicGroupName' already exists. Skipping."
        continue
    }
    $membershipRule = switch ($dynamicGroupName) {
        "[EDU v2 Security] - $SchoolCode All Staff Corporate Shared Mac Devices"   { "(device.enrollmentProfileName -eq ""[EDU v2 Security] - $SchoolCode Staff Shared Mac Enrolment"")" }
        "[EDU v2 Security] - $SchoolCode All Student Corporate Shared Mac Devices" { "(device.enrollmentProfileName -eq ""[EDU v2 Security] - $SchoolCode Student Shared Mac Enrolment"")" }
        "[EDU v2 Security] - $SchoolCode All Staff Corporate Personal Mac Devices" { "(device.enrollmentProfileName -eq ""[EDU v2 Security] - $SchoolCode Staff Personal Mac Enrolment"")" }
        "[EDU v2 Security] - $SchoolCode All Student Corporate Personal Mac Devices" { "(device.enrollmentProfileName -eq ""[EDU v2 Security] - $SchoolCode Student Personal Mac Enrolment"")" }
    }
    if ($membershipRule) {
        $null = New-MgGroup -DisplayName $dynamicGroupName -SecurityEnabled:$true -MailEnabled:$false `
            -MailNickname ("dyn" + (Get-Random)) -GroupTypes "DynamicMembership" `
            -MembershipRule $membershipRule -MembershipRuleProcessingState "On"
        "✅ Dynamic Group '$dynamicGroupName' created with rule: $membershipRule"
    }
}

"⏳ Waiting for newly created groups to be available in Graph..."
Start-Sleep -Seconds 20

# --- Resolve Group IDs for Nesting ---
if ($SetupType -eq "Trust" -and $TrustCode) {
    $trustAllCorp         = (Get-GroupByName "[EDU v2 Security] - $TrustCode All Corporate Mac Devices").Id
    $trustAllPersonal     = (Get-GroupByName "[EDU v2 Security] - $TrustCode All Corporate Personal Mac Devices").Id
    $trustAllShared       = (Get-GroupByName "[EDU v2 Security] - $TrustCode All Corporate Shared Mac Devices").Id
    $trustStaffPersonal   = (Get-GroupByName "[EDU v2 Security] - $TrustCode All Corporate Personal Staff Mac Devices").Id
    $trustStaffShared     = (Get-GroupByName "[EDU v2 Security] - $TrustCode All Corporate Shared Staff Mac Devices").Id
    $trustStudentPersonal = (Get-GroupByName "[EDU v2 Security] - $TrustCode All Corporate Personal Student Mac Devices").Id
    $trustStudentShared   = (Get-GroupByName "[EDU v2 Security] - $TrustCode All Corporate Shared Student Mac Devices").Id
}
if ($SetupType -in @("Trust","School")) {
    $schoolAllCorp         = (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Corporate Mac Devices").Id
    $schoolAllPersonal     = (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Corporate Personal Mac Devices").Id
    $schoolAllShared       = (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Corporate Shared Mac Devices").Id
    $schoolStaffPersonal   = (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Staff Corporate Personal Mac Devices").Id
    $schoolStaffShared     = (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Staff Corporate Shared Mac Devices").Id
    $schoolStudentPersonal = (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Student Corporate Personal Mac Devices").Id
    $schoolStudentShared   = (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Student Corporate Shared Mac Devices").Id
}

# --- Nesting Logic ---
"🔄 Applying nesting rules..."

if ($SetupType -eq "Trust") {
    AddMemberIfNotExists $trustAllShared $trustStaffShared "Trust Shared -> Trust Staff Shared"
    AddMemberIfNotExists $trustAllShared $trustStudentShared "Trust Shared -> Trust Student Shared"

    AddMemberIfNotExists $trustStaffShared $schoolStaffShared "Trust Staff Shared -> School Staff Shared"
    AddMemberIfNotExists $trustStudentShared $schoolStudentShared "Trust Student Shared -> School Student Shared"

    AddMemberIfNotExists $trustAllCorp $trustAllShared "Trust All -> Trust Shared"
    AddMemberIfNotExists $trustAllCorp $trustAllPersonal "Trust All -> Trust Personal"

    AddMemberIfNotExists $trustAllPersonal $trustStaffPersonal "Trust Personal -> Trust Staff Personal"
    AddMemberIfNotExists $trustAllPersonal $trustStudentPersonal "Trust Personal -> Trust Student Personal"

    AddMemberIfNotExists $trustStaffPersonal $schoolStaffPersonal "Trust Staff Personal -> School Staff Personal"
    AddMemberIfNotExists $trustStudentPersonal $schoolStudentPersonal "Trust Student Personal -> School Student Personal"
}

if ($SetupType -in @("Trust","School")) {
    AddMemberIfNotExists $schoolAllShared $schoolStaffShared "School Shared -> School Staff Shared"
    AddMemberIfNotExists $schoolAllShared $schoolStudentShared "School Shared -> School Student Shared"

    AddMemberIfNotExists $schoolAllPersonal $schoolStaffPersonal "School Personal -> School Staff Personal"
    AddMemberIfNotExists $schoolAllPersonal $schoolStudentPersonal "School Personal -> School Student Personal"

    AddMemberIfNotExists $schoolAllCorp $schoolAllShared "School All -> School Shared"
    AddMemberIfNotExists $schoolAllCorp $schoolAllPersonal "School All -> School Personal"
}

"🎉 Completed Mac Groups setup."

Disconnect-MgGraph | Out-Null