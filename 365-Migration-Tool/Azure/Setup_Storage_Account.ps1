<#
Developed by: Rhys Saddul
#>

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

# ------------------------------------------------------------
# LOGGING
# ------------------------------------------------------------
$logDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "CloudDonny_Logs\Setup_Storage_Account"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$successLog = Join-Path $logDir "SUCCESS_$timestamp.log"
$failLog    = Join-Path $logDir "FAILED_$timestamp.log"

function Add-ToLog {
    param([string]$Message, [switch]$Failed)

    $time  = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
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

Add-ToLog "🚀 Starting CloudSchool Storage Account Setup"
Add-ToLog "------------------------------------------------------------"

# ------------------------------------------------------------
# INPUT
# ------------------------------------------------------------
$schoolCode = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter the School Code (e.g. ILT):",
    "School Code",
    "ILT"
)

if ([string]::IsNullOrWhiteSpace($schoolCode)) {
    Add-ToLog "❌ No school code entered" -Failed
    exit
}

Add-ToLog "🏫 School Code: $schoolCode"

# ------------------------------------------------------------
# VARIABLES (UNCHANGED FROM WORKING SCRIPT)
# ------------------------------------------------------------
$eduthing = "EDU"
$location = "uksouth"

$resourceGroupName = "${eduthing}_${schoolCode}_Cloudschool"
$resourceGroup     = $resourceGroupName

$storageAccountName = ("${schoolCode}cloudstorage").ToLower()
$accountKind = "StorageV2"
$skuName = "Standard_LRS"
$minimumTlsVersion = "TLS1_2"

$containerName = "edu-$($schoolCode.ToLower())-blobstorage"

$storeAccessPolicyName = "${schoolCode}IntuneBlobPolicy_ReadList"
$permissions = "rl"
$expiryTime = Get-Date -Year 2099 -Month 12 -Day 31

# ------------------------------------------------------------
# CONNECT
# ------------------------------------------------------------
Update-AzConfig -EnableLoginByWam $false | Out-Null
Connect-AzAccount | Out-Null

$SubscriptionID = (Get-AzSubscription).Id
Set-AzContext -SubscriptionId $SubscriptionID | Out-Null

Add-ToLog "📌 Using subscription $SubscriptionID"

# ------------------------------------------------------------
# RESOURCE GROUP
# ------------------------------------------------------------
$existingRG = Get-AzResourceGroup -Name $resourceGroupName -ErrorAction SilentlyContinue

if (-not $existingRG) {
    Add-ToLog "📁 Creating Resource Group $resourceGroupName"
    New-AzResourceGroup -Name $resourceGroupName -Location $location | Out-Null

    for ($i = 20; $i -ge 0; $i--) {
        Add-ToLog "⏳ RG countdown $i"
        Start-Sleep 1
    }
} else {
    Add-ToLog "⚠️ Resource Group already exists"
}

# ------------------------------------------------------------
# PROVIDER REGISTRATION (COPIED AS-IS)
# ------------------------------------------------------------
$provider = Get-AzResourceProvider -ProviderNamespace "Microsoft.Storage"

if ($provider.RegistrationState -eq "Registered") {
    Add-ToLog "✅ Microsoft.Storage already registered"
} else {
    Add-ToLog "📦 Registering Microsoft.Storage"
    Register-AzResourceProvider -ProviderNamespace "Microsoft.Storage"

    for ($i = 30; $i -ge 0; $i--) {
        Add-ToLog "⏳ Provider countdown $i"
        Start-Sleep 1
    }
}

# ------------------------------------------------------------
# STORAGE ACCOUNT
# ------------------------------------------------------------
$existingSA = Get-AzStorageAccount `
    -ResourceGroupName $resourceGroup `
    -Name $storageAccountName `
    -ErrorAction SilentlyContinue

if (-not $existingSA) {
    Add-ToLog "📦 Creating Storage Account $storageAccountName"

    New-AzStorageAccount `
        -ResourceGroupName $resourceGroup `
        -Name $storageAccountName `
        -Location $location `
        -SkuName $skuName `
        -Kind $accountKind `
        -AllowBlobPublicAccess $false `
        -MinimumTlsVersion $minimumTlsVersion | Out-Null

    for ($i = 20; $i -ge 0; $i--) {
        Add-ToLog "⏳ Storage countdown $i"
        Start-Sleep 1
    }
} else {
    Add-ToLog "⚠️ Storage Account already exists"
}

# ------------------------------------------------------------
# RETENTION POLICIES (COPIED)
# ------------------------------------------------------------
$properties = Get-AzStorageBlobServiceProperty `
    -ResourceGroupName $resourceGroup `
    -StorageAccountName $storageAccountName

if (-not $properties.DeleteRetentionPolicy.Enabled) {
    Add-ToLog "🛡 Enabling retention policies"

    Enable-AzStorageBlobDeleteRetentionPolicy `
        -ResourceGroupName $resourceGroup `
        -StorageAccountName $storageAccountName `
        -RetentionDays 7 | Out-Null

    Enable-AzStorageContainerDeleteRetentionPolicy `
        -ResourceGroupName $resourceGroup `
        -StorageAccountName $storageAccountName `
        -RetentionDays 7 | Out-Null

    for ($i = 20; $i -ge 0; $i--) {
        Add-ToLog "⏳ Retention countdown $i"
        Start-Sleep 1
    }
} else {
    Add-ToLog "⚠️ Retention already enabled"
}

# ------------------------------------------------------------
# STORAGE CONTEXT (KEY-BASED – COPIED)
# ------------------------------------------------------------
$storageKeys = Get-AzStorageAccountKey `
    -ResourceGroupName $resourceGroup `
    -Name $storageAccountName

$storageContext = New-AzStorageContext `
    -StorageAccountName $storageAccountName `
    -StorageAccountKey $storageKeys[0].Value

# ------------------------------------------------------------
# CONTAINER
# ------------------------------------------------------------
$existingContainer = Get-AzStorageContainer `
    -Name $containerName `
    -Context $storageContext `
    -ErrorAction SilentlyContinue

if (-not $existingContainer) {
    Add-ToLog "📁 Creating container $containerName"
    New-AzStorageContainer -Name $containerName -Context $storageContext | Out-Null

    for ($i = 20; $i -ge 0; $i--) {
        Add-ToLog "⏳ Container countdown $i"
        Start-Sleep 1
    }
} else {
    Add-ToLog "⚠️ Container already exists"
}

# ------------------------------------------------------------
# STORED ACCESS POLICY
# ------------------------------------------------------------
$policyExists = Get-AzStorageContainerStoredAccessPolicy `
    -Container $containerName `
    -Policy $storeAccessPolicyName `
    -Context $storageContext `
    -ErrorAction SilentlyContinue

if (-not $policyExists) {
    Add-ToLog "📝 Creating Stored Access Policy $storeAccessPolicyName"

    New-AzStorageContainerStoredAccessPolicy `
        -Container $containerName `
        -Policy $storeAccessPolicyName `
        -Permission $permissions `
        -ExpiryTime $expiryTime `
        -Context $storageContext | Out-Null

    for ($i = 20; $i -ge 0; $i--) {
        Add-ToLog "⏳ Policy countdown $i"
        Start-Sleep 1
    }
} else {
    Add-ToLog "⚠️ Stored Access Policy already exists"
}

# ------------------------------------------------------------
# CLEANUP
# ------------------------------------------------------------
Disconnect-AzAccount | Out-Null
Clear-AzContext -Force | Out-Null

Add-ToLog "🔒 Disconnected from Azure"
Add-ToLog "🎉 Storage Account Setup Complete"
Add-ToLog "------------------------------------------------------------"
