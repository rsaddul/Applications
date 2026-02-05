<#
Developed by: Rhys Saddul
GUI compatible
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

"🔹 Running Windows Devices Groups setup..."
"   SchoolCode : $SchoolCode"
if ($TrustCode) { "   TrustCode  : $TrustCode" }
"   SetupType  : $SetupType"

# --- Connect to Microsoft Graph ---
Connect-MgGraph -Scopes "Group.ReadWrite.All","Directory.ReadWrite.All" | Out-Null

# --- Helper Functions ---
function Get-GroupByName($displayName) {
    try { Get-MgGroup -Filter "displayName eq '$displayName'" -ConsistencyLevel eventual -ErrorAction Stop } catch { $null }
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
        "[EDU v2 Security] - $SchoolCode All Desktop Devices",
        "[EDU v2 Security] - $SchoolCode All Laptop Devices",
        "[EDU v2 Security] - $SchoolCode All Office Devices",
        "[EDU v2 Security] - $SchoolCode All Teacher Devices",
        "[EDU v2 Security] - $SchoolCode All Student Devices",
        "[EDU v2 Security] - $SchoolCode All Corporate Windows Devices",
        "[EDU v2 Security] - $SchoolCode All Assigned or Shared Windows Intune Devices",
        "[EDU v2 Security] - $SchoolCode Global Office Software Windows Intune Devices",
        "[EDU v2 Security] - $SchoolCode Global Teacher Software Windows Intune Devices",
        "[EDU v2 Security] - $SchoolCode Global Student Software Windows Intune Devices",
        "[EDU v2 Security] - $SchoolCode Global Software Windows Intune Devices"
    )
}
if ($SetupType -eq "Trust" -and $TrustCode) {
    $staticGroups += @(
        "[EDU v2 Security] - $TrustCode All Desktop Devices",
        "[EDU v2 Security] - $TrustCode All Laptop Devices",
        "[EDU v2 Security] - $TrustCode All Office Devices",
        "[EDU v2 Security] - $TrustCode All Office Desktops",
        "[EDU v2 Security] - $TrustCode All Office Laptops",
        "[EDU v2 Security] - $TrustCode All Teacher Devices",
        "[EDU v2 Security] - $TrustCode All Teacher Desktops",
        "[EDU v2 Security] - $TrustCode All Teacher Laptops",
        "[EDU v2 Security] - $TrustCode All Student Devices",
        "[EDU v2 Security] - $TrustCode All Student Laptops",
        "[EDU v2 Security] - $TrustCode All Student Desktops",
        "[EDU v2 Security] - $TrustCode All Corporate Windows Devices",
        "[EDU v2 Security] - $TrustCode All Assigned or Shared Windows Intune Devices",
        "[EDU v2 Security] - $TrustCode Global Office Software Windows Intune Devices",
        "[EDU v2 Security] - $TrustCode Global Teacher Software Windows Intune Devices",
        "[EDU v2 Security] - $TrustCode Global Student Software Windows Intune Devices",
        "[EDU v2 Security] - $TrustCode Global Software Windows Intune Devices"
    )
}

foreach ($groupName in $staticGroups) {
    if (GroupExists $groupName) {
        "⚠️ Group '$groupName' already exists. Skipping."
        continue
    }
    New-MgGroup -DisplayName $groupName -SecurityEnabled:$true -MailEnabled:$false `
        -MailNickname ("grp" + (Get-Random)) -GroupTypes @()
    "✅ Static Group '$groupName' created."
}

# --- Dynamic Groups ---
$dynamicGroups = @()
if ($SetupType -in @("Trust","School")) {
    $dynamicGroups += @(
        "[EDU v2 Security] - $SchoolCode All Office Desktops",
        "[EDU v2 Security] - $SchoolCode All Office Laptops",
        "[EDU v2 Security] - $SchoolCode All Teacher Desktops",
        "[EDU v2 Security] - $SchoolCode All Teacher Laptops",
        "[EDU v2 Security] - $SchoolCode All Student Desktops",
        "[EDU v2 Security] - $SchoolCode All Student Laptops",
        "All Personal Owned Non Corporate Devices"
    )
}

foreach ($dynamicGroupName in $dynamicGroups) {
    if (GroupExists $dynamicGroupName) {
        "⚠️ Group '$dynamicGroupName' already exists. Skipping."
        continue
    }
    $membershipRule = switch ($dynamicGroupName) {
        "[EDU v2 Security] - $SchoolCode All Office Desktops"   { "(device.displayName -startsWith ""$SchoolCode-OD-"")" }
        "[EDU v2 Security] - $SchoolCode All Office Laptops"    { "(device.displayName -startsWith ""$SchoolCode-OL-"")" }
        "[EDU v2 Security] - $SchoolCode All Teacher Desktops"  { "(device.displayName -startsWith ""$SchoolCode-TD-"")" }
        "[EDU v2 Security] - $SchoolCode All Teacher Laptops"   { "(device.displayName -startsWith ""$SchoolCode-TL-"")" }
        "[EDU v2 Security] - $SchoolCode All Student Desktops"  { "(device.displayName -startsWith ""$SchoolCode-SD-"")" }
        "[EDU v2 Security] - $SchoolCode All Student Laptops"   { "(device.displayName -startsWith ""$SchoolCode-SL-"")" }
        "All Personal Owned Non Corporate Devices"              { '(device.deviceOwnership -eq "Personal") -or (device.deviceOwnership -eq "Unknown")' }
    }
    if ($membershipRule) {
        New-MgGroup -DisplayName $dynamicGroupName -SecurityEnabled:$true -MailEnabled:$false `
            -MailNickname ("dyn" + (Get-Random)) -GroupTypes "DynamicMembership" `
            -MembershipRule $membershipRule -MembershipRuleProcessingState "On"
        "✅ Dynamic Group '$dynamicGroupName' created with rule: $membershipRule"
    }
}

# --- Wait for groups to propagate ---
"⏳ Waiting for newly created groups to be available in Graph..."
Start-Sleep -Seconds 20

# --- Resolve Group IDs ---
if ($SetupType -eq "Trust" -and $TrustCode) {
    # Trust core groups
    $trustAllCorp     = (Get-GroupByName "[EDU v2 Security] - $TrustCode All Corporate Windows Devices").Id
    $trustAllDesktop  = (Get-GroupByName "[EDU v2 Security] - $TrustCode All Desktop Devices").Id
    $trustAllLaptop   = (Get-GroupByName "[EDU v2 Security] - $TrustCode All Laptop Devices").Id
    $trustAllOffice   = (Get-GroupByName "[EDU v2 Security] - $TrustCode All Office Devices").Id
    $trustAllTeacher  = (Get-GroupByName "[EDU v2 Security] - $TrustCode All Teacher Devices").Id
    $trustAllStudent  = (Get-GroupByName "[EDU v2 Security] - $TrustCode All Student Devices").Id

    # Trust dynamic subgroups
    $trustOfficeDesktops = (Get-GroupByName "[EDU v2 Security] - $TrustCode All Office Desktops").Id
    $trustOfficeLaptops  = (Get-GroupByName "[EDU v2 Security] - $TrustCode All Office Laptops").Id
    $trustTeacherDesktops= (Get-GroupByName "[EDU v2 Security] - $TrustCode All Teacher Desktops").Id
    $trustTeacherLaptops = (Get-GroupByName "[EDU v2 Security] - $TrustCode All Teacher Laptops").Id
    $trustStudentDesktops= (Get-GroupByName "[EDU v2 Security] - $TrustCode All Student Desktops").Id
    $trustStudentLaptops = (Get-GroupByName "[EDU v2 Security] - $TrustCode All Student Laptops").Id

    # Trust software
    $trustGlobalSoftware        = (Get-GroupByName "[EDU v2 Security] - $TrustCode Global Software Windows Intune Devices").Id
    $trustGlobalOfficeSoftware  = (Get-GroupByName "[EDU v2 Security] - $TrustCode Global Office Software Windows Intune Devices").Id
    $trustGlobalTeacherSoftware = (Get-GroupByName "[EDU v2 Security] - $TrustCode Global Teacher Software Windows Intune Devices").Id
    $trustGlobalStudentSoftware = (Get-GroupByName "[EDU v2 Security] - $TrustCode Global Student Software Windows Intune Devices").Id

    # Trust autopilot
    $trustAllAssigned = (Get-GroupByName "[EDU v2 Security] - $TrustCode All Assigned or Shared Windows Intune Devices").Id
}

if ($SetupType -in @("Trust","School")) {
    # School core groups
    $schoolAllCorp    = (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Corporate Windows Devices").Id
    $schoolAllDesktop = (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Desktop Devices").Id
    $schoolAllLaptop  = (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Laptop Devices").Id
    $schoolAllOffice  = (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Office Devices").Id
    $schoolAllTeacher = (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Teacher Devices").Id
    $schoolAllStudent = (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Student Devices").Id

    # School dynamic subgroups
    $schoolOfficeDesktops = (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Office Desktops").Id
    $schoolOfficeLaptops  = (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Office Laptops").Id
    $schoolTeacherDesktops= (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Teacher Desktops").Id
    $schoolTeacherLaptops = (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Teacher Laptops").Id
    $schoolStudentDesktops= (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Student Desktops").Id
    $schoolStudentLaptops = (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Student Laptops").Id

    # School software
    $schoolGlobalSoftware        = (Get-GroupByName "[EDU v2 Security] - $SchoolCode Global Software Windows Intune Devices").Id
    $schoolGlobalOfficeSoftware  = (Get-GroupByName "[EDU v2 Security] - $SchoolCode Global Office Software Windows Intune Devices").Id
    $schoolGlobalTeacherSoftware = (Get-GroupByName "[EDU v2 Security] - $SchoolCode Global Teacher Software Windows Intune Devices").Id
    $schoolGlobalStudentSoftware = (Get-GroupByName "[EDU v2 Security] - $SchoolCode Global Student Software Windows Intune Devices").Id

    # School autopilot
    $schoolAllAssigned = (Get-GroupByName "[EDU v2 Security] - $SchoolCode All Assigned or Shared Windows Intune Devices").Id
}

# --- Nesting Logic ---
"🔄 Applying nesting rules..."

if ($SetupType -eq "Trust") {
    # Trust -> School All Device Nesting
    AddMemberIfNotExists $trustAllOffice $schoolAllOffice "Trust Office -> School Office"
    AddMemberIfNotExists $trustAllTeacher $schoolAllTeacher "Trust Teacher -> School Teacher"
    AddMemberIfNotExists $trustAllStudent $schoolAllStudent "Trust Student -> School Student"

    # Trust All Desktop -> Subgroups
    AddMemberIfNotExists $trustAllDesktop $trustOfficeDesktops "Trust Desktop -> Office Desktops"
    AddMemberIfNotExists $trustAllDesktop $trustTeacherDesktops "Trust Desktop -> Teacher Desktops"
    AddMemberIfNotExists $trustAllDesktop $trustStudentDesktops "Trust Desktop -> Student Desktops"

    # Trust -> School Desktop
    AddMemberIfNotExists $trustOfficeDesktops $schoolOfficeDesktops "Trust Office Desktops -> School Office Desktops"
    AddMemberIfNotExists $trustTeacherDesktops $schoolTeacherDesktops "Trust Teacher Desktops -> School Teacher Desktops"
    AddMemberIfNotExists $trustStudentDesktops $schoolStudentDesktops "Trust Student Desktops -> School Student Desktops"

    # Trust All Laptop -> Subgroups
    AddMemberIfNotExists $trustAllLaptop $trustOfficeLaptops "Trust Laptop -> Office Laptops"
    AddMemberIfNotExists $trustAllLaptop $trustTeacherLaptops "Trust Laptop -> Teacher Laptops"
    AddMemberIfNotExists $trustAllLaptop $trustStudentLaptops "Trust Laptop -> Student Laptops"

    # Trust -> School Laptop
    AddMemberIfNotExists $trustOfficeLaptops $schoolOfficeLaptops "Trust Office Laptops -> School Office Laptops"
    AddMemberIfNotExists $trustTeacherLaptops $schoolTeacherLaptops "Trust Teacher Laptops -> School Teacher Laptops"
    AddMemberIfNotExists $trustStudentLaptops $schoolStudentLaptops "Trust Student Laptops -> School Student Laptops"

    # Trust Global Software -> Subgroups
    AddMemberIfNotExists $trustGlobalSoftware $trustGlobalOfficeSoftware "Trust Software -> Office Software"
    AddMemberIfNotExists $trustGlobalSoftware $trustGlobalTeacherSoftware "Trust Software -> Teacher Software"
    AddMemberIfNotExists $trustGlobalSoftware $trustGlobalStudentSoftware "Trust Software -> Student Software"

    # Trust -> School Software
    AddMemberIfNotExists $trustGlobalOfficeSoftware $schoolGlobalOfficeSoftware "Trust Office SW -> School Office SW"
    AddMemberIfNotExists $trustGlobalTeacherSoftware $schoolGlobalTeacherSoftware "Trust Teacher SW -> School Teacher SW"
    AddMemberIfNotExists $trustGlobalStudentSoftware $schoolGlobalStudentSoftware "Trust Student SW -> School Student SW"

    # Trust All Corporate -> Core
    AddMemberIfNotExists $trustAllCorp $trustAllDesktop "Trust Corp -> Trust Desktop"
    AddMemberIfNotExists $trustAllCorp $trustAllLaptop "Trust Corp -> Trust Laptop"

    # Trust -> School Assigned (Autopilot)
    AddMemberIfNotExists $trustAllAssigned $schoolAllAssigned "Trust Assigned -> School Assigned"
}

if ($SetupType -in @("Trust","School")) {
    # School Devices -> Subgroups
    AddMemberIfNotExists $schoolAllOffice $schoolOfficeDesktops "School Office -> School Office Desktops"
    AddMemberIfNotExists $schoolAllOffice $schoolOfficeLaptops "School Office -> School Office Laptops"
    AddMemberIfNotExists $schoolAllTeacher $schoolTeacherDesktops "School Teacher -> School Teacher Desktops"
    AddMemberIfNotExists $schoolAllTeacher $schoolTeacherLaptops "School Teacher -> School Teacher Laptops"
    AddMemberIfNotExists $schoolAllStudent $schoolStudentDesktops "School Student -> School Student Desktops"
    AddMemberIfNotExists $schoolAllStudent $schoolStudentLaptops "School Student -> School Student Laptops"

    # School Desktop -> Subgroups
    AddMemberIfNotExists $schoolAllDesktop $schoolOfficeDesktops "School Desktop -> School Office Desktops"
    AddMemberIfNotExists $schoolAllDesktop $schoolTeacherDesktops "School Desktop -> School Teacher Desktops"
    AddMemberIfNotExists $schoolAllDesktop $schoolStudentDesktops "School Desktop -> School Student Desktops"

    # School Laptop -> Subgroups
    AddMemberIfNotExists $schoolAllLaptop $schoolOfficeLaptops "School Laptop -> School Office Laptops"
    AddMemberIfNotExists $schoolAllLaptop $schoolTeacherLaptops "School Laptop -> School Teacher Laptops"
    AddMemberIfNotExists $schoolAllLaptop $schoolStudentLaptops "School Laptop -> School Student Laptops"

    # School Software -> Subgroups
    AddMemberIfNotExists $schoolGlobalSoftware $schoolGlobalOfficeSoftware "School SW -> School Office SW"
    AddMemberIfNotExists $schoolGlobalSoftware $schoolGlobalTeacherSoftware "School SW -> School Teacher SW"
    AddMemberIfNotExists $schoolGlobalSoftware $schoolGlobalStudentSoftware "School SW -> School Student SW"

    # School Corp -> Core
    AddMemberIfNotExists $schoolAllCorp $schoolAllDesktop "School Corp -> School Desktop"
    AddMemberIfNotExists $schoolAllCorp $schoolAllLaptop "School Corp -> School Laptop"
}

"🎉 Completed Windows Devices Groups setup."

Disconnect-MgGraph | Out-Null