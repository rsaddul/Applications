<# 
Connected Cache Setup
#>

# Disable WAM-based login (quiet)
Update-AzConfig -EnableLoginByWam $false | Out-Null

# Prompt GUI for required values
Add-Type -AssemblyName Microsoft.VisualBasic

$schoolCode = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter the School Code (e.g., BWS):", "School Code Input", "BWS"
)
if ([string]::IsNullOrWhiteSpace($schoolCode)) { "❌ No School Code entered. Exiting..."; exit }

$trustCode = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter the Trust Code (e.g., ILT):", "Trust Code Input", "ILT"
)
if ([string]::IsNullOrWhiteSpace($trustCode)) { "❌ No Trust Code entered. Exiting..."; exit }

$connectedcacheVM = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter the Connected Cache VM FQDN:", "Connected Cache VM", "ILT-BWS-MCC.bishopwand.surrey.sch.uk"
)
if ([string]::IsNullOrWhiteSpace($connectedcacheVM)) { "❌ No VM entered. Exiting..."; exit }

$AdminEmail = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter the Global Admin Email:", "Admin Email", "admin@InstanterLearningTrust.onmicrosoft.com"
)
if ([string]::IsNullOrWhiteSpace($AdminEmail)) { "❌ No Admin Email entered. Exiting..."; exit }

# -----------------------
# Configurable Variables
# -----------------------
$autoUpdateDay    = "6"
$autoUpdateRing   = "fast"
$autoUpdateTime   = "19:00"
$autoUpdateWeek   = "2"
$cacheNodeName    = "${schoolCode}_ConnectedCacheNode"
$mccResourceName  = "ConnectedCache"
$proxy            = "Disabled"

# System Variables
$connectedCacheLocation = "northeurope"
$eduthing              = "EDU"
$hostOS                = "Windows"
$location              = "ukwest"
$physicalPathLetter    = "Y"
$physicalPathLocation  = "/var/mcc"
$resourceGroupName     = "${eduthing}_${schoolCode}_ConnectedCache"
$resourceGroup         = $resourceGroupName
$sizeInGb              = "120"

"ℹ️ Using SchoolCode: $schoolCode, TrustCode: $trustCode, AdminEmail: $AdminEmail"

# -----------------------
# Pre-checks
# -----------------------
# Hyper-V
if ((Get-WindowsFeature -Name Hyper-V).Installed) {
    "✔️ Hyper-V already installed."
}
else {
    "⚠️ Hyper-V not installed. Installing..."
    try {
        $result = Install-WindowsFeature -Name Hyper-V -IncludeManagementTools -ErrorAction Stop
        if ($result.Success) { "✔️ Hyper-V installed (restart may be required)." }
        else { "❌ Hyper-V installation failed: $($result.ExitCode)" }
    }
    catch {
        "❌ Hyper-V install failed: $($_.Exception.Message)"
    }
}

# WSL
$wslFeature = Get-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-Subsystem-Linux
if ($wslFeature.State -eq "Enabled") {
    "✔️ WSL already installed."
}
else {
    "⚠️ WSL not installed. Installing..."
    wsl.exe --install --no-distribution
    "⚠️ WSL installation executed. Restart required."
}

# Disk space
$psd = Get-PSDrive -PSProvider FileSystem -Name $physicalPathLetter.TrimEnd(':')
$freeGB = [math]::Round($psd.Free / 1GB, 2)
if ($freeGB -ge $sizeInGb) {
    "✔️ $physicalPathLetter drive has $freeGB GB free (>= $sizeInGb GB)"
} else {
    "❌ $physicalPathLetter drive only has $freeGB GB (< $sizeInGb GB)"
    exit 1
}

# Azure CLI
if (Get-Command az -ErrorAction SilentlyContinue) {
    "✔️ Azure CLI installed."
}
else {
    "⚠️ Azure CLI not found. Installing..."
    $installer = "$env:TEMP\AzureCLI.msi"
    Invoke-WebRequest -Uri https://aka.ms/installazurecliwindows -OutFile $installer
    Start-Process msiexec.exe -Wait -ArgumentList "/I `"$installer`" /quiet /norestart"
    Remove-Item $installer -Force
    "✔️ Azure CLI installed. Run 'az login' after restart if required."
}

# -----------------------
# Azure Setup
# -----------------------

# Clear any cached logins
az account clear | Out-Null 2>&1

# Prompt user to sign in again
"🔐 Please sign in with your Azure credentials..."
az login --only-show-errors | Out-Null 2>&1


$SubscriptionID = az account show --query id -o tsv
$scope = "/subscriptions/$SubscriptionID"

"🔑 Checking if $AdminEmail is Owner of subscription $SubscriptionID..."
$assignments = az role assignment list `
  --assignee $AdminEmail `
  --scope $scope `
  --role Owner `
  --include-inherited `
  -o json --only-show-errors | ConvertFrom-Json

if ($assignments.Count -gt 0) {
    "✔️ $AdminEmail is Owner at $scope"
} else {
    "❌ $AdminEmail is NOT Owner at $scope"
    exit 1
}

# -----------------------
# Resource Group
# -----------------------
$exists = (az group exists --name $resourceGroupName -o tsv) -eq "true"
if (-not $exists) {
    "⚙️ Creating resource group $resourceGroupName"
    az group create --name $resourceGroupName --location $location 1>$null
    for ($i = 20; $i -ge 0; $i--) { "$i seconds remaining..."; Start-Sleep 1 }
    "✔️ $resourceGroupName created"
} else {
    "✔️ Resource group $resourceGroupName already exists."
}

# -----------------------
# Connected Cache Resource
# -----------------------
az config set extension.dynamic_install_allow_preview=true --only-show-errors
az provider register --namespace Microsoft.ConnectedCache

$exists = az mcc ent resource list `
    --resource-group $resourceGroupName `
    --query "[?name=='$mccResourceName'] | length(@)" `
    -o tsv

if ($exists -gt 0) {
    "✔️ MCC resource '$mccResourceName' already exists."
}
else {
    "⚙️ Creating MCC resource '$mccResourceName'..."
    az mcc ent resource create --mcc-resource-name $mccResourceName --resource-group $resourceGroupName --location $connectedCacheLocation
    for ($i = 20; $i -ge 0; $i--) { "$i seconds remaining..."; Start-Sleep 1 }
    "✔️ MCC resource '$mccResourceName' created."
}

# -----------------------
# Cache Node
# -----------------------
$nodeExists = az mcc ent node list `
    --mcc-resource-name $mccResourceName `
    --resource-group $resourceGroupName `
    --query "[?name=='$cacheNodeName'] | length(@)" `
    -o tsv

if ($nodeExists -gt 0) {
    "✔️ Cache node '$cacheNodeName' already exists."
}
else {
    "⚙️ Creating cache node '$cacheNodeName'..."
    az mcc ent node create --cache-node-name $cacheNodeName --mcc-resource-name $mccResourceName --resource-group $resourceGroupName --host-os $hostOS
    for ($i = 20; $i -ge 0; $i--) { "$i seconds remaining..."; Start-Sleep 1 }
    "✔️ Cache node '$cacheNodeName' created."
}

# -----------------------
# DHCP Info
# -----------------------
$cacheNodeId = az mcc ent node show `
    --cache-node-name $cacheNodeName `
    --mcc-resource-name $mccResourceName `
    --resource-group $resourceGroupName `
    --query "cacheNodeId" -o tsv

"ℹ️ Cache Node ID (Group ID): $cacheNodeId"
"ℹ️ Set DHCP OptionID 234 to $cacheNodeId"
"ℹ️ Set DHCP OptionID 235 to $connectedcacheVM"

"✅ Connected Cache setup complete."

az account clear | Out-Null 2>&1
"✔️ Signed out of all Azure CLI sessions."