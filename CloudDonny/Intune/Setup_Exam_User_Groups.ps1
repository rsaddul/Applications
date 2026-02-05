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

"🔹 Running Exam Groups setup..."
"   SchoolCode : $SchoolCode"
if ($TrustCode) { "   TrustCode  : $TrustCode" }
"   SetupType  : $SetupType"

# --- Connect to Microsoft Graph ---
Connect-MgGraph -Scopes "Group.ReadWrite.All","Directory.ReadWrite.All" | Out-Null

# --- Helper Functions ---
function Get-GroupByName($displayName) {
    try { 
        Get-MgGroup -Filter "displayName eq '$displayName'" -ConsistencyLevel eventual -ErrorAction Stop 
    } catch { 
        $null 
    }
}
function GroupExists($displayName) { return (Get-GroupByName $displayName) -ne $null }
function Create-SecurityGroup($name) {
    if (GroupExists $name) {
        "⚠️ Group '$name' already exists."
        return (Get-GroupByName $name).Id
    } else {
        $newGroup = New-MgGroup -DisplayName $name `
            -SecurityEnabled:$true -MailEnabled:$false `
            -MailNickname ("grp_" + [guid]::NewGuid().ToString())
        "✅ Created Security Group: $name"
        return $newGroup.Id
    }
}
function MemberExists($groupId, $memberId) {
    try { (Get-MgGroupMember -GroupId $groupId -All).Id -contains $memberId } catch { $false }
}
function Add-GroupMemberIfNotExists($parentId, $childId, $relation) {
    if (-not $parentId -or -not $childId) { return }
    if (MemberExists $parentId $childId) {
        "⚠️ Nesting already exists: $relation"
    } else {
        New-MgGroupMember -GroupId $parentId -DirectoryObjectId $childId
        "✅ Added nesting: $relation"
    }
}

# --- Section 1: Create Exam Security Groups ---
$examGroups = @()
if ($SetupType -eq "Trust") {
    $examGroups += @(
        "$TrustCode Exam Disable Spellcheck",
        "$TrustCode Exam All Students",
        "$TrustCode Exam Enable Spellcheck"
    )
}
if ($SetupType -eq "School") {
    $examGroups += @(
        "$SchoolCode Exam Lockdown Internet",
        "$SchoolCode Exam All Students",
        "$SchoolCode Exam Enable Spellcheck",
        "$SchoolCode Exam Disable Spellcheck"
    )
}

$examGroupIds = @{}
foreach ($eg in $examGroups) {
    $examGroupIds[$eg] = Create-SecurityGroup $eg
}

# --- Section 2: Exam Group Nesting (Trust ↔ School) ---
if ($SetupType -eq "Trust" -and $TrustCode) {
    "🔄 Applying nesting rules..."

    $tasEnable  = (Get-GroupByName "$SchoolCode Exam Enable Spellcheck").Id
    $tasDisable = (Get-GroupByName "$SchoolCode Exam Disable Spellcheck").Id
    $tasAll     = (Get-GroupByName "$SchoolCode Exam All Students").Id

    $trustEnable  = (Get-GroupByName "$TrustCode Exam Enable Spellcheck").Id
    $trustDisable = (Get-GroupByName "$TrustCode Exam Disable Spellcheck").Id
    $trustAll     = (Get-GroupByName "$TrustCode Exam All Students").Id

    Add-GroupMemberIfNotExists $trustEnable  $tasEnable  "$TrustCode Enable Spellcheck -> $SchoolCode Enable Spellcheck"
    Add-GroupMemberIfNotExists $trustDisable $tasDisable "$TrustCode Disable Spellcheck -> $SchoolCode Disable Spellcheck"
    Add-GroupMemberIfNotExists $trustAll     $tasAll     "$TrustCode All Students -> $SchoolCode All Students"
}

"🎉 Completed Exam Groups setup."

Disconnect-MgGraph | Out-Null