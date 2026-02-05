<#
Developed by: Rhys Saddul

SharePoint Overview:

This script automates the setup and configuration of SharePoint Online sites, including the creation of hub and child sites, permission settings, and custom libraries. 
Key functions include site creation, hub association, library setup, and configuring custom security groups and permissions.

1.)  Setup and Confirmation Checks:
     The script starts by prompting the user to confirm that necessary security groups and variables have been configured. 
     It ensures these prerequisites are in place to prevent errors during execution.

2.)  Variable Setup:
     Key variables, such as SharePoint URLs, site codes, email suffixes, and site titles, are defined for later use in the script. 
     These include information for hub and child sites, permissions, and any special configuration needs, such as CloudMapper or Cloud Drive Mapper.

3.)  Child Sites and Hub Site Creation:
     The script checks if each child site (like Staff or Operations sites) and the designated hub site already exists. 
     If not, it creates them. It also verifies that the hub site is registered as such, and if not, registers it as a hub site.

4.)  Site Association with Hub Site:
     Once the hub site is set up, the script associates each child site with the hub site to enable structured content and navigation. 
     This association ensures all sites are part of a cohesive hub.

5.)  Library Creation:
     Libraries (document repositories) are created within each site if they do not already exist.
     These libraries might include folders like Archive, Risk Assessments, Pupil Progress, and others specific to each site’s function.

6.)  Removing Default SharePoint Groups:
     Default groups like the "Visitors" group and the "Hub Visitors" group are removed from each document library. 
     This is done to replace default permissions with custom configurations defined in the script.

7.)  Security Group Configuration:
     Custom security groups are applied to the sites, granting specific permissions. 
     For example, certain groups may receive "Edit" permissions on a library, while others may have "Read" access only. 
     This section configures granular permissions across different libraries, depending on the group's needs.

8.)  Inheritance Management:
     Role inheritance (inheriting permissions from the parent site) is disabled for each library. 
     This allows each library to have its unique permissions, independent of the overall site’s default permissions.

9:)  Regional and Storage:
     Regional settings for all SharePoint sites set as GMT and Storage set at 5TB.

10:) Access Request Settings:
     Disables Access Request Settings and Clears Members Shares options for all SharePoint sites.
     
 IMPORTANT --------- If you would like to intergrate a Student setup, un-hash these Varibles: 1, 2, , 4, 6 and 7 --------- IMPORTANT 

Office 365 Group Setup Overview:

This script automates the setup and configuration of Office 365 group email accounts for a Multi-Academy Trust (MAT) or individual school setup. 
It creates and organises email groups for different users and departments across the organisation, streamlining communication and access management.

1.) Define Key Variables: 
    Core variables, such as the MAT code alias, school site codes, email type prefix, domain suffix, and global admin account, are defined for consistent use throughout the script.

2.) Define MAT Groups: 
    For MAT-level setups, the script configures email groups for all users, all students, and all staff across the trust. 
    Each group is assigned a structured email and alias based on the MAT code alias, allowing for efficient, centralised communication and management.

3.) Define Site Groups: 
    For each specified school site, the script creates nested groups based on setup type (Trust or School). 
    These groups include All Users, Staff Site, and Operation Site, as well as department-specific groups.
    Such as Safeguarding, Finance, and SLT (Senior Leadership Team) for site-specific organisation.

4.) Group Naming and Email Structure: 
    The script dynamically generates email addresses and aliases for each group by combining site codes, group types, and domain suffixes. 
    This provides a standardised naming convention across all groups for consistency and clarity.

5.) Setup Type Conditional Logic: 
    Depending on whether the setup is for an entire MAT ("Trust") or an individual school ("School"), the script adjusts the group creation process.
    Allowing flexible deployment across various organisational structures.

6.) Nested Groups: 
    Specific groups are added into nested groups to allow easy management of SharePoint
                           
#>

param([System.Windows.Forms.TextBox]$outputBox)

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

# ==============================
#   INITIALISE LOGGING + OUTPUT
# ==============================

# Create a unified logging function (works with or without GUI)
function Append-Log {
    param([string]$msg)
    $time = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
    $entry = "[$time] $msg"

    if ($outputBox) {
        # Append the log text
        $outputBox.AppendText("$entry`r`n")

        # 🪄 Auto-scroll to the latest line
        $outputBox.SelectionStart = $outputBox.Text.Length
        $outputBox.ScrollToCaret()

        # Optional: keep focus if you want typing in the box
        $outputBox.Refresh()
    } else {
        Write-Host $entry
    }
}

Append-Log "Script starting..."

# Create or reuse desktop log folder
$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs\SharePoint_Creation"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}

$timestamp = (Get-Date).ToString("dd/MM/yyyy - HH:mm:ss")
$successLog = Join-Path $logDir "Setup_SharePoint_Success_$timestamp.log"
$failLog    = Join-Path $logDir "Setup_SharePoint_Failed_$timestamp.log"

# Unified file + screen logger
function Add-ToLog {
    param(
        [string]$Message,
        [switch]$Failed
    )
    $time = (Get-Date).ToString("HH:mm:ss")
    $entry = "[$time] $Message"

    try {
        if ($Failed) {
            Add-Content $failLog $entry -ErrorAction SilentlyContinue
        } else {
            Add-Content $successLog $entry -ErrorAction SilentlyContinue
        }
    } catch {}
}


# --------- Prompt to choose 365 Group setup type --------- #

$setupType = [System.Windows.Forms.MessageBox]::Show(
    "Select 365 Group setup type (Yes for Trust, No for School)", 
    "Setup Type", 
    [System.Windows.Forms.MessageBoxButtons]::YesNo, 
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if ($setupType -eq [System.Windows.Forms.DialogResult]::Yes) {
    $setupType = "Trust"
    Append-Log "Trust setup selected."
    Add-ToLog "Trust setup selected."
} else {
    $setupType = "School"
    Append-Log "School setup selected."
    Add-ToLog "School setup selected."
}

# --------- Confirm the setup type --------- #
$confirmationResult = [System.Windows.Forms.MessageBox]::Show(
    "Are you sure you want to proceed with the 365 $setupType setup?", 
    "Confirmation", 
    [System.Windows.Forms.MessageBoxButtons]::YesNo, 
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if ($confirmationResult -eq [System.Windows.Forms.DialogResult]::Yes) {
    Append-Log "$setupType setup confirmed. Proceeding with the script."
    Add-ToLog "$setupType setup confirmed. Proceeding with the script."
} else {
    Append-Log "$setupType setup not confirmed. Exiting the script."
    Add-ToLog "$setupType setup not confirmed. Exiting the script." -Failed
    exit 1
}

# ========= Variable Prompts =========

# MAT Code
if (-not $matCodeAlias) {
    $matCodeAlias = [Microsoft.VisualBasic.Interaction]::InputBox("Enter MAT Code Alias:", "MAT Code", "EDU")
    if ([string]::IsNullOrWhiteSpace($matCodeAlias)) { "❌ MAT Code is required."; exit 1 }
}

# Site Codes
if (-not $siteCodes) {
    $siteInput = [Microsoft.VisualBasic.Interaction]::InputBox("Enter School Codes (comma separated):", "Site Codes", "TCT")
    $siteCodes = $siteInput -split "," | ForEach-Object { $_.Trim() }
    if ($siteCodes.Count -eq 0) { "❌ Site Codes are required."; exit 1 }
}

# Type (fixed)
$Type = "MESG"

# Domain
if (-not $Domain) {
    $Domain = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Office 365 Domain:", "Domain", "thecloudschool.co.uk")
    if ([string]::IsNullOrWhiteSpace($Domain)) { "❌ Domain is required."; exit 1 }
}

# Admin Account
if (-not $AdminAccount) {
    $AdminAccount = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Global Admin Account (without domain):", "Admin Account", "rsaddul")
    if ([string]::IsNullOrWhiteSpace($AdminAccount)) { "❌ Admin Account is required."; exit 1 }
}

# SharePoint Base URL
if (-not $baseSharePointURL) {
    $baseSharePointURL = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Base SharePoint URL:", "SharePoint URL", "https://eduthingazurelab.sharepoint.com/sites/")
    if ([string]::IsNullOrWhiteSpace($baseSharePointURL)) { "❌ SharePoint URL is required."; exit 1 }
}

# Site Owner
if (-not $SiteOwner) {
    $SiteOwner = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Site Owner (UPN):", "Site Owner", "rsaddul@thecloudschool.co.uk")
    if ([string]::IsNullOrWhiteSpace($SiteOwner)) { "❌ Site Owner is required."; exit 1 }
}

# Hub Site Title
if (-not $hubSiteTitle) {
    $hubSiteTitle = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Hub Site Title:", "Hub Site", "TCT")
    if ([string]::IsNullOrWhiteSpace($hubSiteTitle)) { "❌ Hub Site Title is required."; exit 1 }
}

# Site Code
if (-not $siteCode) {
    $siteCode = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Site Code:", "Site Code", "TCT")
    if ([string]::IsNullOrWhiteSpace($siteCode)) { "❌ Site Code is required."; exit 1 }
}

# Suffix
if (-not $Suffix) {
    $Suffix = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Domain Suffix:", "Suffix", "@thecloudschool.co.uk")
    if ([string]::IsNullOrWhiteSpace($Suffix)) { "❌ Suffix is required."; exit 1 }
}

# Client ID
if (-not $clientId) {
    $clientId = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Azure App Client ID: Run the Setup PnP Automation script if you dont have these details ", "Client ID", "0a2e4081-bb3a-45a2-b63d-a5abe8657965")
    if ([string]::IsNullOrWhiteSpace($clientId)) { "❌ Client ID is required."; exit 1 }
}


Append-Log "✔ Variables captured:"
Append-Log "   SetupType: $setupType"
Append-Log "   MAT: $matCodeAlias"
Append-Log "   SiteCodes: $($siteCodes -join ', ')"
Append-Log "   Domain: $Domain"
Append-Log "   Admin: $AdminAccount"
Append-Log "   ClientID: $clientId"
Add-ToLog "Variables successfully captured for $setupType setup."


# --------- Office 365 Varible Section --------- #

<#
Varible 1: Define MAT groups as a hash table if setupType is Trust
#>

$matGroups = @{}
if ($setupType -eq "Trust") {
    $matGroups = @{
        "MAT_All_Users" = @{ Name = "${matCodeAlias} All Users"; Email = "${Type}_${matCodeAlias}_All_Users@$Domain"; Alias = "${Type}_${matCodeAlias}_All_Users" }
        "MAT_All_Students" = @{ Name = "${matCodeAlias} All Students"; Email = "${Type}_${matCodeAlias}_All_Students@$Domain"; Alias = "${Type}_${matCodeAlias}_All_Students" }
        "MAT_All_Staff" = @{ Name = "${matCodeAlias} All Staff"; Email = "${Type}_${matCodeAlias}_All_Staff@$Domain"; Alias = "${Type}_${matCodeAlias}_All_Staff" }
        # --------- Add more as needed --------- #     
    }
}

<#
Varible 3: Define site groups as a nested hash table for both Trust and School setup types
#>

$siteGroups = @{}
if ($setupType -eq "Trust" -or $setupType -eq "School") {
    foreach ($siteCode in $siteCodes) {
        $siteGroups[$siteCode] = @{

            # --------- Below Group Normally used for Hub Site --------- #
            "All_Users" = @{ Name = "${siteCode} All Users"; Email = "${Type}_${siteCode}_All_Users@$Domain"; Alias = "${Type}_${siteCode}_All_Users" }
            
            # --------- Below Groups Normally used for Staff Site and Libraries --------- #
            "Staff_Site" = @{ Name = "${siteCode} Staff Site Users"; Email = "${Type}_${siteCode}_Staff_Site@$Domain"; Alias = "${Type}_${siteCode}_Staff_Site" }
            "All_Staff" = @{ Name = "${siteCode} All Staff"; Email = "${Type}_${siteCode}_All_Staff@$Domain"; Alias = "${Type}_${siteCode}_All_Staff" }
            "Safeguarding_Staff" = @{ Name = "${siteCode} Safeguarding Staff"; Email = "${Type}_${siteCode}_Safeguarding_Staff@$Domain"; Alias = "${Type}_${siteCode}_Safeguarding_Staff" }

            # --------- Below Groups Normally used for Operation Site and Libraries --------- #
            "Operation_Site" = @{ Name = "${siteCode} Operation Site Users"; Email = "${Type}_${siteCode}_Operation_Site@$Domain"; Alias = "${Type}_${siteCode}_Operation_Site" }
            "Office_Staff" = @{ Name = "${siteCode} Office Staff"; Email = "${Type}_${siteCode}_Office_Staff@$Domain"; Alias = "${Type}_${siteCode}_Office_Staff" }
            "Finance_Staff" = @{ Name = "${siteCode} Finance Staff"; Email = "${Type}_${siteCode}_Finance_Staff@$Domain"; Alias = "${Type}_${siteCode}_Finance_Staff" }
            "Personnel_Staff" = @{ Name = "${siteCode} Personnel Staff"; Email = "${Type}_${siteCode}_Personnel_Staff@$Domain"; Alias = "${Type}_${siteCode}_Personnel_Staff" }
            "SLT_Staff" = @{ Name = "${siteCode} SLT Staff"; Email = "${Type}_${siteCode}_SLT_Staff@$Domain"; Alias = "${Type}_${siteCode}_SLT_Staff" }
            
            "All_Students" = @{ Name = "${siteCode} All Students"; Email = "${Type}_${siteCode}_All_Students@$Domain"; Alias = "${Type}_${siteCode}_All_Students" }
            "Year1" = @{ Name = "${siteCode} Year 1 Students"; Email = "${Type}_${siteCode}_Year_1_Students@$Domain"; Alias = "${Type}_${siteCode}_Year_1_Students" }
            "Year2" = @{ Name = "${siteCode} Year 2 Students"; Email = "${Type}_${siteCode}_Year_2_Students@$Domain"; Alias = "${Type}_${siteCode}_Year_2_Students" }
            "Year3" = @{ Name = "${siteCode} Year 3 Students"; Email = "${Type}_${siteCode}_Year_3_Students@$Domain"; Alias = "${Type}_${siteCode}_Year_3_Students" }
            "Year4" = @{ Name = "${siteCode} Year 4 Students"; Email = "${Type}_${siteCode}_Year_4_Students@$Domain"; Alias = "${Type}_${siteCode}_Year_4_Students" }
            "Year5" = @{ Name = "${siteCode} Year 5 Students"; Email = "${Type}_${siteCode}_Year_5_Students@$Domain"; Alias = "${Type}_${siteCode}_Year_5_Students" }
            # --------- Add more as needed --------- #
        }
    }
}


# --------- SharePoint Varible Section --------- #

$siteTemplate = "SITEPAGEPUBLISHING#0" # --------- Communication site template --------- #
$hubSiteDescription = "This is the hub site." 
$staffSite = "Staff" # ---------  Used throughout the script for SharePoint site targeting --------- #
$operationSite = "Operations" # --------- Used throughout the script for SharePoint site targeting --------- #
$studentSite = "Student" # --------- Used throughout the script for SharePoint site targeting --------- #

<#
Varible 1: Define sites that will be created and associated with the hub site. 
#>

$childSites = @(  
    @{
        "Title" = "$siteCode $staffSite"
        "RelativeURL" = "$siteCode-$staffSite"
    },
    @{
        "Title" = "$siteCode $operationSite"
        "RelativeURL" = "$siteCode-$operationSite"
    },
    @{
        "Title" = "$siteCode $studentSite"
        "RelativeURL" = "$siteCode-$studentSite"
    }
    # --------- DO NOT ADD THE HUB SITE --------- Add more sites as needed --------- #
)

<#
Varible 2: Define mail-enabled security groups which will be used for site and library permissions. 
#>

# ---------  Hub Site Groups --------- #
$allUsers = "MESG_${siteCode}_All_Users" + $Suffix

# --------- Staff Site Groups --------- #
$allStaff = "MESG_${siteCode}_All_Staff" + $Suffix
$staffStaffSite = "MESG_${siteCode}_Staff_Site" + $Suffix
$operationStaffSite = "MESG_${siteCode}_Operation_Site" + $Suffix
$office = "MESG_${siteCode}_Office_Staff" + $Suffix
$finance = "MESG_${siteCode}_Finance_Staff" + $Suffix
$personnel = "MESG_${siteCode}_Personnel_Staff" + $Suffix
$sLT = "MESG_${siteCode}_SLT_Staff" + $Suffix
$safeguarding = "MESG_${siteCode}_Safeguarding_Staff" + $Suffix
# --------- IMPORTANT --------- Add more groups as needed --------- #

# ---------  Student Site Groups --------- #
$allStudents = "MESG_${siteCode}_All_Students" + $Suffix 
$year1 = "MESG_${siteCode}_Year_1_Students" + $Suffix 
$year2 = "MESG_${siteCode}_Year_2_Students" + $Suffix
$year3 = "MESG_${siteCode}_Year_3_Students" + $Suffix
$year4 = "MESG_${siteCode}_Year_4_Students" + $Suffix
$year5 = "MESG_${siteCode}_Year_5_Students" + $Suffix 
# --------- IMPORTANT --------- Add more groups as needed --------- #

<#
Varible 3: Define the SharePoint sites and the Libaries you would like created.
#>

$sitesWithLibraries = @(  
   @{
        SiteUrl   = "$baseSharePointURL$($siteCode)-$($staffSite)"
        Libraries = @("General", "Media", "Planning", "Safeguarding")
        Title = "$($siteCode) $($staffSite)"
    },
    @{
        SiteUrl   = "$baseSharePointURL$($siteCode)-$($operationSite)"
        Libraries = @("Office", "Finance", "Personnel", "SLT")
        Title = "$($siteCode) $($operationSite)"
    },
    @{
        SiteUrl   = "$baseSharePointURL$($siteCode)-$($studentSite)"
        Libraries = @("General", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5")
        Title = "$($siteCode) $($studentSite)"
    }
    # --------- DO NOT ADD THE HUB SITE --------- Add more sites as needed --------- #
)

<#
Varible 4: Define variable for the site/s you dont want to create libraies on
#>

$siteWithoutLibraries = @{
    SiteUrl   = "$baseSharePointURL${siteCode}" 
    Groups    = @($allUsers)
    Title = "${siteCode}"
    # --------- TARGET THE HUB SITE ONLY --------- # 
}

<#
Varible 5: Define sites and security groups that will be used for setting the site SharePoint Group Visitors permissions.
#>

$groupsForPermissionsAndAssociation = @(  
    @{
        SiteUrl   = "$baseSharePointURL$($siteCode)-$($staffSite)"
        Groups    = @($staffStaffSite) # (Referencing Varible 3)
    },
    @{
        SiteUrl   = "$baseSharePointURL$($siteCode)-$($operationSite)"
        Groups    = @($operationStaffSite) # (Referencing Varible 3)
    },
    @{
        SiteUrl   = "$baseSharePointURL${siteCode}"
        Groups    = @($allUsers) # (Referencing Varible 3)
    },
    @{
        SiteUrl   = "$baseSharePointURL$($siteCode)-$($studentSite)"
        Groups    = @($allStaff,$allStudents) # (Referencing Varible 3)
    }
    # Add more groups as needed
)

<#
Varible 6: Define site and specify which security groups will be applied to each library with edit or read permissions. 
#>

$sitesLibrariesSecurityGroupsMapping = @{
    "$baseSharePointURL$($siteCode)-$($staffSite)" = @{
        "General" =  @{
            "Edit" = @("$allStaff")      # --------- Group with edit permissions
          # "Read" = @("")               # --------- No "Read" key specified for this library, indicating no read permissions
        }
        "Media" =  @{
            "Edit" = @("$allStaff")      # --------- Group with edit permissions
           # "Read" = @("")               # ---------No "Read" key specified for this library, indicating no read permissions
        }
        "Planning" =  @{
            "Edit" = @("$allStaff")      # --------- Group with edit permissions
          # "Read" = @("")               # --------- No "Read" key specified for this library, indicating no read permissions
        }
        "Safeguarding" =  @{
            "Edit" = @("$safeguarding")  # --------- Group with edit permissions
          # "Read" = @("")               # --------- No "Read" key specified for this library, indicating no read permissions
        }
        # Add more libraries, and security groups as needed
    } 
    "$baseSharePointURL$($siteCode)-$($operationSite)" = @{
        "Office" =  @{
            "Edit" = @("$office")        # --------- Group with edit permissions
          # "Read" = @("")               # --------- No "Read" key specified for this library, indicating no read permissions
        }
        "Finance" =  @{
            "Edit" = @("$finance")       # --------- Group with edit permissions
          # "Read" = @("")               # --------- No "Read" key specified for this library, indicating no read permissions
        }
        "Personnel" =  @{
            "Edit" = @("$personnel")     # --------- Group with edit permissions
          # "Read" = @("")               # --------- No "Read" key specified for this library, indicating no read permissions
        }
        "SLT" =  @{
            "Edit" = @("$sLT")  # Group with edit permissions
          # "Read" = @("")       # No "Read" key specified for this library, indicating no read permissions
        }
        # Add more libraries, and security groups as needed
    } 
    "$baseSharePointURL$($siteCode)-$($studentSite)" = @{
        "General" = @{
            "Edit" = @("$allStaff")      # --------- Group with edit permissions
            "Read" = @("$allStudents")   # --------- Group with read permissions
        }
        "Year 1" = @{
            "Edit" = @("$allStaff")      # --------- Group with edit permissions
            "Read" = @("$year1")          # --------- Group with read permissions
        }
        "Year 2" =  @{
            "Edit" = @("$allStaff")      # --------- Group with edit permissions
            "Read" = @("$year2")          # --------- Group with read permissions
        }
        "Year 3" =  @{
            "Edit" = @("$allStaff")      # --------- Group with edit permissions
            "Read" = @("$year3")          # --------- Group with read permissions
        }
        "Year 4" =  @{
            "Edit" = @("$allStaff")      # --------- Group with edit permissions
            "Read" = @("$year4")         # --------- Group with read permissions
        }
        "Year 5" =  @{
            "Edit" = @("$allStaff")      # --------- Group with edit permissions
            "Read" = @("$year5")         # --------- Group with read permissions
        }
        # --------- Add more libraries, and security groups as needed --------- #
    }
        # --------- Add mappings for other SharePoint sites as needed --------- #
}



# --------- Office 365 Script Execution Section --------- #

Append-Log "Starting Office 365 Script Execution"
Add-ToLog "Starting Office 365 Script Execution"

# ==============================================================
#   Switch to delegated user connection for group creation
# ==============================================================
Append-Log "🔑 Switching to delegated Exchange Online connection for group creation..."
Add-ToLog  "Switched to delegated Exchange Online connection for group creation."

Connect-ExchangeOnline

# Section 1: Loop through matGroups to create groups for matCodes staff (Trust setup only)

Append-Log "Skipping Section 1: Creating MAT groups"
Add-ToLog "Skipping Section 1: Creating MAT groups"

if ($setupType -eq "Trust") {
    
Append-Log "Starting Section 1: Creating MAT groups"
Add-ToLog "Starting Section 1: Creating MAT groups"

    foreach ($matGroup in $matGroups.Keys) {
        $matGroupName  = $matGroups[$matGroup].Name
        $matGroupEmail = $matGroups[$matGroup].Email
        $matGroupAlias = $matGroups[$matGroup].Alias
        $matGroupDescription = "This group is used for Intune and SharePoint permissions for $matGroupName"

        if (-not (Get-DistributionGroup -Identity $matGroupEmail -ErrorAction SilentlyContinue)) {	
	        Append-Log "Creating mail-enabled security group: $matGroupName with email $matGroupEmail"
		Add-ToLog "Creating mail-enabled security group: $matGroupName with email $matGroupEmail"

            New-DistributionGroup -Name $matGroupName -Alias $matGroupAlias -PrimarySmtpAddress $matGroupEmail -Description $matGroupDescription -Type "Security" | Out-Null
            Set-DistributionGroup -Identity $matGroupEmail -HiddenFromAddressListsEnabled $true -ErrorAction SilentlyContinue
        } else {
	        Append-Log "Group already exists: $matGroupName"
		Add-ToLog "Group already exists: $matGroupName"

        }
    }
}

# Section 2: Loop through siteGroups to create groups for site staff
foreach ($siteCode in $siteGroups.Keys) {

Append-Log "Starting Section 2: Create School groups"
Add-ToLog "Starting Section 2: Create School groups"
 
    $siteGroupData = $siteGroups[$siteCode]
    foreach ($staffCategory in $siteGroupData.Keys) {
        $siteGroupName  = $siteGroupData[$staffCategory].Name
        $siteGroupEmail = $siteGroupData[$staffCategory].Email
        $siteGroupAlias = $siteGroupData[$staffCategory].Alias
        $siteGroupDescription = "This group is used for Intune and SharePoint permissions for $siteGroupName"

        if (-not (Get-DistributionGroup -Identity $siteGroupEmail -ErrorAction SilentlyContinue)) {            		
		Append-Log "Creating mail-enabled security group: $siteGroupName with email $siteGroupEmail"
		Add-ToLog "Creating mail-enabled security group: $siteGroupName with email $siteGroupEmail"

            New-DistributionGroup -Name $siteGroupName -Alias $siteGroupAlias -PrimarySmtpAddress $siteGroupEmail -Description $siteGroupDescription -Type "Security" | Out-Null
            Set-DistributionGroup -Identity $siteGroupEmail -HiddenFromAddressListsEnabled $true -ErrorAction SilentlyContinue
        } else {
		Append-Log "Group already exists: $siteGroupName"
		Add-ToLog "Group already exists: $siteGroupName"

        }
    }
	Append-Log "Waiting for 5 seconds before proceeding..."
	Add-ToLog "Waiting for 5 seconds before proceeding..."
	
    Start-Sleep -Seconds 5
}

# Section 3: Loop through siteGroups to add site staff groups to corresponding MAT groups (Trust setup only)
if ($setupType -eq "Trust") {    
Append-Log "Starting Section 3: Adding School groups to MAT groups"
Add-ToLog "Starting Section 3: Adding School groups to MAT groups"

    $siteCategories = @("All_Staff", "All_Students")

    foreach ($siteCode in $siteGroups.Keys) {
        $siteGroupData = $siteGroups[$siteCode]
        foreach ($staffCategory in $siteCategories) {
            if (-not $siteGroupData.ContainsKey($staffCategory)) {
		Append-Log "⚠️ Warning: $staffCategory not defined for $siteCode"
		Add-ToLog "Warning: $staffCategory not defined for $siteCode" -Failed
                continue
            }

            $siteGroupEmail = $siteGroupData[$staffCategory].Email
            $matGroupKey    = "MAT_$staffCategory"
            $matGroupEmail  = if ($matGroups.ContainsKey($matGroupKey)) { $matGroups[$matGroupKey].Email } else { $null }

            if (-not $matGroupEmail) {
                Append-Log "⚠️ Warning: MAT group for $staffCategory not found."
		Add-ToLog "Warning: MAT group for $staffCategory not found." -Failed
                continue
            }
            if (-not (Get-DistributionGroup -Identity $matGroupEmail -ErrorAction SilentlyContinue)) {
                Append-Log "⚠️ Warning: MAT group $matGroupEmail does not exist."
		Add-ToLog "Warning: MAT group $matGroupEmail does not exist." -Failed
	        continue
            }

            if (-not (Get-DistributionGroupMember -Identity $matGroupEmail -ErrorAction SilentlyContinue | Where-Object { $_.PrimarySmtpAddress -eq $siteGroupEmail })) {
                Append-Log "Adding $siteGroupEmail to $matGroupEmail"
		Add-ToLog "Adding $siteGroupEmail to $matGroupEmail"
	                Add-DistributionGroupMember -Identity $matGroupEmail -Member $siteGroupEmail | Out-Null
            } else {
                Append-Log "$siteGroupEmail is already a member of $matGroupEmail"
		Add-ToLog "$siteGroupEmail is already a member of $matGroupEmail"

            }
        }
    }

    $matAllUsers    = $matGroups["MAT_All_Users"].Email
    $matAllStaff    = $matGroups["MAT_All_Staff"].Email
    $matAllStudents = $matGroups["MAT_All_Students"].Email

    if (Get-DistributionGroup -Identity $matAllUsers -ErrorAction SilentlyContinue) {
        foreach ($member in @($matAllStaff, $matAllStudents)) {
            if (-not (Get-DistributionGroupMember -Identity $matAllUsers -ErrorAction SilentlyContinue | Where-Object { $_.PrimarySmtpAddress -eq $member })) {
                Append-Log "Adding $member to $matAllUsers"
		Add-ToLog "Adding $member to $matAllUsers"
                Add-DistributionGroupMember -Identity $matAllUsers -Member $member | Out-Null
            } else {
                Append-Log "$member is already a member of $matAllUsers"
		Add-ToLog "$member is already a member of $matAllUsers"
            }
        }
    } else {
        Append-Log "⚠️ Warning: MAT All Users group $matAllUsers does not exist."
	Add-ToLog "Warning: MAT All Users group $matAllUsers does not exist." -Failed
    }
}

# Section 4: Add specific groups as members of the Operation Site group
foreach ($siteCode in $siteGroups.Keys) {
	Append-Log "Starting Section 4: Adding Office, Finance, Personnel, and SLT groups to Operation Site group"
	Add-ToLog "Starting Section 4: Adding Office, Finance, Personnel, and SLT groups to Operation Site group"
    $operationStaffSite = "MESG_${siteCode}_Operation_Site" + $Suffix
    $office    = "MESG_${siteCode}_Office_Staff" + $Suffix
    $finance   = "MESG_${siteCode}_Finance_Staff" + $Suffix
    $personnel = "MESG_${siteCode}_Personnel_Staff" + $Suffix
    $sLT       = "MESG_${siteCode}_SLT_Staff" + $Suffix

    $groupsToAdd = @($office, $finance, $personnel, $sLT)

    if (Get-DistributionGroup -Identity $operationStaffSite -ErrorAction SilentlyContinue) {
        foreach ($group in $groupsToAdd) {
            if (Get-DistributionGroup -Identity $group -ErrorAction SilentlyContinue) {
                if (-not (Get-DistributionGroupMember -Identity $operationStaffSite -ErrorAction SilentlyContinue | Where-Object { $_.PrimarySmtpAddress -eq $group })) {
			Append-Log "Adding $group to $operationStaffSite"
			Add-ToLog "Adding $group to $operationStaffSite"
	Add-DistributionGroupMember -Identity $operationStaffSite -Member $group | Out-Null
                } else {
			Append-Log "$group is already a member of $operationStaffSite"
			Add-ToLog "$group is already a member of $operationStaffSite"
                }
            } else {
			Append-Log "⚠️ Warning: Group $group does not exist."
			Add-ToLog "Warning: Group $group does not exist." -Failed
            }
        }
    } else {
		Append-Log "⚠️ Warning: Operation Site group $operationStaffSite does not exist."
		Add-ToLog "Warning: Operation Site group $operationStaffSite does not exist." -Failed

    }
	Append-Log "Waiting for 5 seconds before proceeding to the next site code..."
	Add-ToLog "Waiting for 5 seconds before proceeding to the next site code..."

    Start-Sleep -Seconds 5
}

# Section 5: Add All Staff group to Staff Site group
foreach ($siteCode in $siteGroups.Keys) {
	Append-Log "Starting Section 5: Adding All Staff to Staff Site"
	Add-ToLog "Starting Section 5: Adding All Staff to Staff Site"
    $allStaff       = "MESG_${siteCode}_All_Staff" + $Suffix
    $staffStaffSite = "MESG_${siteCode}_Staff_Site" + $Suffix

    if (Get-DistributionGroup -Identity $staffStaffSite -ErrorAction SilentlyContinue) {
        if (-not (Get-DistributionGroupMember -Identity $staffStaffSite -ErrorAction SilentlyContinue | Where-Object { $_.PrimarySmtpAddress -eq $allStaff })) {
		Append-Log "Adding $allStaff to $staffStaffSite"
		Add-ToLog "Adding $allStaff to $staffStaffSite"
            Add-DistributionGroupMember -Identity $staffStaffSite -Member $allStaff | Out-Null
        } else {
		Append-Log "$allStaff is already a member of $staffStaffSite"
		Add-ToLog "$allStaff is already a member of $staffStaffSite"
        }
    } else {
        Append-Log "⚠️ Warning: Staff Site group $staffStaffSite does not exist."
	Add-ToLog "Warning: Staff Site group $staffStaffSite does not exist." -Failed
    }
	Append-Log "Waiting for 5 seconds before proceeding to the next site code..."
	Add-ToLog "Waiting for 5 seconds before proceeding to the next site code..."
    Start-Sleep -Seconds 5
}

# Section 6: Add All Staff and All Students group to All Users group
foreach ($siteCode in $siteGroups.Keys) {
	Append-Log "Starting Section 6: Adding All Staff and All Students to All Users"
	Add-ToLog "Starting Section 6: Adding All Staff and All Students to All Users"
    $allUsers    = "MESG_${siteCode}_All_Users" + $Suffix
    $allStaff    = "MESG_${siteCode}_All_Staff" + $Suffix
    $allStudents = "MESG_${siteCode}_All_Students" + $Suffix

    $groupsToAddToAllUsers = @($allStaff, $allStudents)

    if (Get-DistributionGroup -Identity $allUsers -ErrorAction SilentlyContinue) {
        foreach ($group in $groupsToAddToAllUsers) {
            if (Get-DistributionGroup -Identity $group -ErrorAction SilentlyContinue) {
                if (-not (Get-DistributionGroupMember -Identity $allUsers -ErrorAction SilentlyContinue | Where-Object { $_.PrimarySmtpAddress -eq $group })) {
	                Append-Log "Adding $group to $allUsers"
			Add-ToLog "Adding $group to $allUsers"
                    Add-DistributionGroupMember -Identity $allUsers -Member $group | Out-Null
                } else {
	                Append-Log "$group is already a member of $allUsers"
			Add-ToLog "$group is already a member of $allUsers"
                }
            } else {
                Append-Log "⚠️ Warning: Group $group does not exist."
		Add-ToLog "Warning: Group $group does not exist." -Failed
            }
        }
    } else {
        Append-Log "⚠️ Warning: All Users group $allUsers does not exist."
	Add-ToLog "Warning: All Users group $allUsers does not exist." -Failed
    }
	Append-Log "Waiting for 5 seconds before proceeding to the next site code..."
	Add-ToLog "Waiting for 5 seconds before proceeding to the next site code..."
    Start-Sleep -Seconds 5
}

Append-Log "Office 365 Script Execution Finished"
Add-ToLog "Office 365 Script Execution Finished"

Disconnect-ExchangeOnline -Confirm:$false | Out-Null  # --------- Disconnect from Exchange Online --------- #

# --------- SharePoint Script Execution Section --------- #

Append-Log "Starting SharePoint Execution Script"
Add-ToLog "Starting SharePoint Execution Script"

# Section 1
Append-Log "Starting Section 1: Creating SharePoint sites"
Add-ToLog "Starting Section 1: Creating SharePoint sites"
Append-Log "Checking if the child sites exist before creation"
Add-ToLog "Checking if the child sites exist before creation"

foreach ($childSite in $childSites) {
    $siteUrl = "$baseSharePointURL$($childSite["RelativeURL"])"
    Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $clientId
    $existingSite = Get-PnPTenantSite -Url $siteUrl -ErrorAction SilentlyContinue
    if (-not $existingSite) {
        Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $clientId
        New-PnPTenantSite -Title $childSite["Title"] -Url $siteUrl -Owner $SiteOwner -Template $siteTemplate -TimeZone 4 -RemoveDeletedSite
		Append-Log "Creating $($childSite.Title)"
		Add-ToLog "Creating $($childSite.Title)"
		Append-Log "$($childSite.Title) Created"
		Add-ToLog "$($childSite.Title) Created"
    } else {
        Append-Log "Site $($childSite.Title) already exists."
	Add-ToLog "Site $($childSite.Title) already exists."
    }
}

# Section 2
Append-Log "Starting Section 2: Creating Hub Site"
Add-ToLog "Starting Section 2: Creating Hub Site"	
Append-Log "Checking if the hub site exists before creation"
Add-ToLog "Checking if the hub site exists before creation"

$hubSiteUrl = "$baseSharePointURL$siteCode"
Connect-PnPOnline -Url $hubSiteUrl -Interactive -ClientId $clientId
$existingHubSite = Get-PnPTenantSite -Url $hubSiteUrl -ErrorAction SilentlyContinue
if (-not $existingHubSite) {
    $hubSite = New-PnPTenantSite -Title $hubSiteTitle -Url $hubSiteUrl -Owner $SiteOwner -Template $siteTemplate -TimeZone 4 -RemoveDeletedSite
	Append-Log "Creating Hub site $hubSiteTitle"
	Add-ToLog "Creating Hub site $hubSiteTitle"
	Append-Log "Hub site created: $hubSiteTitle"
	Add-ToLog "Hub site created: $hubSiteTitle"
	Append-Log "Be patient while the sites are syncing"
	Add-ToLog "Be patient while the sites are syncing"
    for ($i = 90; $i -ge 0; $i--) {
    Append-Log "$i"
    Start-Sleep -Seconds 1
}
	Append-Log "Site creation Countdown complete! Thank you for being patient."
	Add-ToLog "Site creation Countdown complete! Thank you for being patient."
} else {
	Append-Log "Hub site $($existingHubSite.Title) already exists."
	Add-ToLog "Hub site $($existingHubSite.Title) already exists."
}

# Section 3
Append-Log "Starting Section 3: Register the Hub site"
Add-ToLog "Starting Section 3: Register the Hub site"
Connect-PnPOnline -Url $hubSiteUrl -Interactive -ClientId $clientId
$existingHub = Get-PnPHubSite | Where-Object { $_.SiteUrl -eq $hubSiteUrl }
if ($existingHub) {
	Append-Log "Hub site $hubSiteUrl is already registered as a HubSite."
	Add-ToLog "Hub site $hubSiteUrl is already registered as a HubSite."
} else {
	Append-Log "Registering hub site $hubSiteUrl"
	Add-ToLog "Registering hub site $hubSiteUrl"
    try {
        Register-PnPHubSite -Site $hubSiteUrl
for ($i = 20; $i -ge 0; $i--) {
    Append-Log "$i"
    Start-Sleep -Seconds 1
}
	Append-Log "Hubsite registration countdown complete!"
	Add-ToLog "Hubsite registration countdown complete!"
    } catch {
		Append-Log "$hubSiteUrl already exists as a HubSite"
		Add-ToLog "$hubSiteUrl already exists as a HubSite"
    }
}

# Section 4
Append-Log "Starting Section 4: Associate SharePoint sites with the Hub site"
Add-ToLog "Starting Section 4: Associate SharePoint sites with the Hub site"
foreach ($childSite in $childSites) {
    $siteUrl = "$baseSharePointURL$($childSite["RelativeURL"])"
    Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $clientId
    $siteProperties = Get-PnPTenantSite -Identity $siteUrl | Select -ExpandProperty HubSiteId
    if ($siteProperties -ne "00000000-0000-0000-0000-000000000000") {
        Append-Log "$($childSite.Title) is already associated with a hub site."
	Add-ToLog "$($childSite.Title) is already associated with a hub site."
    } else {
        Add-PnPHubSiteAssociation -Site $siteUrl -HubSite $hubSiteUrl
		Append-Log "Associating $($childSite.Title) with $hubSiteUrl"
		Add-ToLog "Associating $($childSite.Title) with $hubSiteUrl"
	for ($i = 15; $i -ge 0; $i--) {
	    Append-Log "$i"
	    Start-Sleep -Seconds 1
}
	Append-Log "$($childSite.Title) association countdown complete!"
	Add-ToLog "$($childSite.Title) association countdown complete!"
    }
}

# Section 5
Append-Log "Starting Section 5: Create the Libraries on the SharePoint sites"
Add-ToLog "Starting Section 5: Create the Libraries on the SharePoint sites"
$sitesWithLibraries += $siteWithoutLibraries
foreach ($site in $sitesWithLibraries) {
    if ($site.SiteUrl -ne $siteWithoutLibraries.SiteUrl) {
        Connect-PnPOnline -Url $site.SiteUrl -Interactive -ClientId $clientId
        foreach ($library in $site.Libraries) {
            $existingLibrary = Get-PnPList -Identity $library -ErrorAction SilentlyContinue
            if (-not $existingLibrary) {
                New-PnPList -Title $library -Template DocumentLibrary -ErrorAction SilentlyContinue | Out-Null
                	Append-Log "Library '$library' created on $($site.SiteUrl)"
			Add-ToLog "Library '$library' created on $($site.SiteUrl)"
            } else {
                Append-Log "Library '$library' already exists on $($site.SiteUrl)"
		Add-ToLog "Library '$library' already exists on $($site.SiteUrl)"
            }
        }
    }
}
$sitesWithLibraries = $sitesWithLibraries | Where-Object { $_ -ne $siteWithoutLibraries }

# Section 6
Append-Log "Starting Section 6: Remove the Visitors group from the Default Library"
Add-ToLog "Starting Section 6: Remove the Visitors group from the Default Library"
$sitesWithLibraries += $siteWithoutLibraries
foreach ($site in $sitesWithLibraries) {
    try {
        Connect-PnPOnline -Url $site.SiteUrl -Interactive -ClientId $clientId
        $defaultDocumentLibrary = Get-PnPList -Identity "Documents"
        if ($defaultDocumentLibrary) {
            Set-PnPList -Identity "Documents" -BreakRoleInheritance -CopyRoleAssignments | Out-Null
            $allGroups = Get-PnPGroup
            $specificVisitorsGroupTitle = "$($site.Title) Visitors"
            $specificVisitorsGroup = $allGroups | Where-Object { $_.Title -eq $specificVisitorsGroupTitle }
            if ($specificVisitorsGroup) {
                try {
                    $defaultDocumentLibrary.RoleAssignments.GetByPrincipal($specificVisitorsGroup).DeleteObject()
                    Invoke-PnPQuery
	                Append-Log "Removed $specificVisitorsGroupTitle from $($site.SiteUrl)"
			Add-ToLog "Removed $specificVisitorsGroupTitle from $($site.SiteUrl)"
                } catch {
	                Append-Log "$specificVisitorsGroupTitle already removed from $($site.SiteUrl)"
			Add-ToLog "$specificVisitorsGroupTitle already removed from $($site.SiteUrl)"
                }
            }
            $hubVisitorsGroup = $allGroups | Where-Object { $_.Title -eq "Hub Visitors" }
            if ($hubVisitorsGroup) {
                try {
                    $defaultDocumentLibrary.RoleAssignments.GetByPrincipal($hubVisitorsGroup).DeleteObject()
                    Invoke-PnPQuery
	                Append-Log "Removed Hub Visitors from $($site.SiteUrl)"
			Add-ToLog "Removed Hub Visitors from $($site.SiteUrl)"
                } catch {
			Append-Log "Failed to remove Hub Visitors from $($site.SiteUrl)"
			Add-ToLog "Failed to remove Hub Visitors from $($site.SiteUrl)" -Failed
                }
            }
        }
    } catch {
        Append-Log "Error processing $($site.SiteUrl): $_"
	Add-ToLog "Error processing $($site.SiteUrl): $_" -Failed
    }
}
$sitesWithLibraries = $sitesWithLibraries | Where-Object { $_ -ne $siteWithoutLibraries }

# Section 7
Append-Log "Starting Section 7: Remove Hub Visitors and add Visitors group members"
Add-ToLog "Starting Section 7: Remove Hub Visitors and add Visitors group members"
$visitorGroupId = 4
foreach ($site in $groupsForPermissionsAndAssociation) {
    $siteUrl = $site["SiteUrl"]
	Append-Log "Configuring security groups on $siteUrl"
	Add-ToLog "Configuring security groups on $siteUrl"
    Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $clientId
    $allGroups = Get-PnPGroup
    $hubVisitorsGroup = $allGroups | Where-Object { $_.Title -eq "Hub Visitors" }
    if ($hubVisitorsGroup) { Remove-PnPGroup -Identity $hubVisitorsGroup.Id -Force }
    $visitorsGroup = $allGroups | Where-Object { $_.Id -eq $visitorGroupId }
    if ($visitorsGroup) {
        foreach ($groupEmail in $site["Groups"]) {
            Add-PnPGroupMember -Group $visitorsGroup.Title -EmailAddress $groupEmail -ErrorAction SilentlyContinue
	        Append-Log "Added $groupEmail to Visitors on $siteUrl"
		Add-ToLog "Added $groupEmail to Visitors on $siteUrl"
        }
    }
}


# Section 8
Append-Log "Starting Section 8: Disable inheritance, clean Visitors from libraries"
Add-ToLog "Starting Section 8: Disable inheritance, clean Visitors from libraries"

foreach ($siteUrl in $sitesLibrariesSecurityGroupsMapping.Keys) {
    Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $clientId

    foreach ($library in $sitesLibrariesSecurityGroupsMapping[$siteUrl].Keys) {
        try {
            # --- Get and break inheritance ---
            $libraryObject = Get-PnPList -Identity $library
            $libraryObject.BreakRoleInheritance($true, $false)
            Invoke-PnPQuery
            Append-Log "🔓 Inheritance broken for $library"
            Add-ToLog "Inheritance broken for $library"

            # --- Remove Visitors groups if they exist ---
            $context = Get-PnPContext
            $context.Load($libraryObject.RoleAssignments)
            $context.ExecuteQuery()

            foreach ($assignment in $libraryObject.RoleAssignments) {
                $context.Load($assignment.Member)
            }
            $context.ExecuteQuery()

            $removedCount = 0
            foreach ($assignment in $libraryObject.RoleAssignments) {
                if ($assignment.Member.Title -like "*Visitors") {
                    Append-Log "🗑 Removing $($assignment.Member.Title) from $library"
                    Add-ToLog "Removing $($assignment.Member.Title) from $library"
                    $assignment.DeleteObject()
                    $removedCount++
                }
            }
            if ($removedCount -gt 0) { $context.ExecuteQuery() }

            # --- Get existing role assignments again (after cleanup) ---
            $context.Load($libraryObject.RoleAssignments)
            $context.ExecuteQuery()
            foreach ($assignment in $libraryObject.RoleAssignments) {
                $context.Load($assignment.Member)
                $context.Load($assignment.RoleDefinitionBindings)
            }
            $context.ExecuteQuery()

            # --- Apply mapped permissions only if not already present ---
            $permissions = $sitesLibrariesSecurityGroupsMapping[$siteUrl][$library]
            foreach ($permissionLevel in $permissions.Keys) {
                foreach ($group in $permissions[$permissionLevel]) {

                    $alreadyAssigned = $false
                    foreach ($assignment in $libraryObject.RoleAssignments) {
                        if ($assignment.Member.LoginName -like "*$group*" -and
                            ($assignment.RoleDefinitionBindings.Name -contains $permissionLevel)) {
                            $alreadyAssigned = $true
                            break
                        }
                    }

                    if (-not $alreadyAssigned) {
                        Set-PnPListPermission -Identity $library -User $group -AddRole $permissionLevel
                        Append-Log "✅ Added $group to $library with $permissionLevel"
                        Add-ToLog "Added $group to $library with $permissionLevel"
                    } else {
                        Append-Log "ℹ️ $group already has $permissionLevel on $library — skipping"
                        Add-ToLog "$group already has $permissionLevel on $library — skipping"
                    }
                }
            }

        } catch {
            Append-Log "❌ Error configuring ${library} on ${siteUrl}: $_"
            Add-ToLog "Error configuring ${library} on ${siteUrl}: $_" -Failed
        }
    }
}
<#
Section 9: Set the SharePoint sites Regional Settings and Storage
#>
Append-Log "Starting Section 9: Set the Regional and Storage Settings on all SharePoint sites"
Add-ToLog "Starting Section 9: Set the Regional and Storage Settings on all SharePoint sites"

foreach ($site in $childSites) {
    $SiteURL = "$baseSharePointURL$($site.RelativeURL)"
    Append-Log "Connecting to site: $($site.Title) at $SiteURL"
    Add-ToLog "Connecting to site: $($site.Title) at $SiteURL"

    $TimeZoneId = 2      # GMT Standard Time
    $LocaleId   = 2057   # English - United Kingdom

    Connect-PnPOnline -Url $SiteURL -Interactive -ClientId $clientId

    if (Get-PnPContext) {
        try {
            $Web = Get-PnPWeb -Includes RegionalSettings,RegionalSettings.TimeZones
            $TimeZone = $Web.RegionalSettings.TimeZones | Where-Object { $_.Id -eq $TimeZoneId }
            $Web.RegionalSettings.TimeZone = $TimeZone
            $Web.RegionalSettings.LocaleId = $LocaleId
            $Web.Update()
            Invoke-PnPQuery

            Set-PnPTenantSite -Url $SiteURL -StorageMaximumLevel 5242880

            Append-Log "✅ Regional settings and storage updated for $($site.Title)"
            Add-ToLog "Regional settings and storage updated for $($site.Title)"
        } catch {
            Append-Log "❌ Error updating $($site.Title): $_"
            Add-ToLog "Error updating $($site.Title): $_" -Failed
        }
    } else {
        Append-Log "❌ Failed to connect to $SiteURL"
        Add-ToLog "Failed to connect to $SiteURL" -Failed
    }
}


<#
Section 10: Disable Access Requests and Sharing Permissions
#>
Append-Log "Starting Section 10: Disable Access Requests and Member Sharing"
Add-ToLog "Starting Section 10: Disable Access Requests and Member Sharing"

foreach ($site in $childSites) {
    $RelativeURL = $site["RelativeURL"]
    $Title = $site["Title"]
    $SiteURL = "$baseSharePointURL$RelativeURL"

    try {
        Append-Log "Connecting to site: $Title ($SiteURL)"
        Add-ToLog "Connecting to site: $Title ($SiteURL)"

        Connect-PnPOnline -Url $SiteURL -Interactive -ClientId $clientId
        $Web = Get-PnPWeb

        # Disable access requests
        $Web.RequestAccessEmail = [string]::Empty
        $Web.SetUseAccessRequestDefaultAndUpdate($False)
        $Web.Update()
        Invoke-PnPQuery

        # Disable member sharing
        $Web.MembersCanShare = $false
        $Web.Update()
        Invoke-PnPQuery

        # Prevent members editing group membership
        $MembersGroup = Get-PnPGroup -AssociatedMemberGroup
        Set-PnPGroup -Identity $MembersGroup -AllowMembersEditMembership:$false

        Append-Log "🔒 Disabled access requests + member sharing on $Title"
        Add-ToLog "Disabled access requests + member sharing on $Title"
    } catch {
        Append-Log "❌ Failed to apply access/sharing settings to ${Title}: $_"
        Add-ToLog "Failed to apply access/sharing settings to ${Title}: $_" -Failed
    }
}

# Section 11
Append-Log "Starting Section 11: Disable Sync Button"
Add-ToLog "Starting Section 11: Disable Sync Button"
$sitesWithLibraries += $siteWithoutLibraries
foreach ($site in $sitesWithLibraries) {
    try {
        Connect-PnPOnline -Url $site.SiteUrl -Interactive -ClientId $clientId
        $Web = Get-PnPWeb -Includes ExcludeFromOfflineClient
        if ($Web.ExcludeFromOfflineClient -eq $True) {
		Append-Log "Offline Client already disabled on $($site.SiteUrl)"
		Add-ToLog "Offline Client already disabled on $($site.SiteUrl)"
        } else {
            $Web.ExcludeFromOfflineClient = $True; $Web.Update(); Invoke-PnPQuery
	        Append-Log "Disabled Offline Client on $($site.SiteUrl)"
		Add-ToLog "Disabled Offline Client on $($site.SiteUrl)"
        }
    } catch {
        Append-Log "Failed disabling Offline Client on $($site.SiteUrl)"
	Add-ToLog "Failed disabling Offline Client on $($site.SiteUrl)" -Failed
    }
}
$sitesWithLibraries = $sitesWithLibraries | Where-Object { $_ -ne $siteWithoutLibraries }

Append-Log "SharePoint Script Execution Finished"
Add-ToLog "SharePoint Script Execution Finished"
Disconnect-PnPOnline | Out-Null
Append-Log "-------------------------------- SharePoint design completed, proceed to Visual Design --------------------------------"
Add-ToLog "-------------------------------- SharePoint design completed, proceed to Visual Design --------------------------------"