<#
Developed by: Rhys Saddul
#>

Add-Type -AssemblyName System.Windows.Forms
$WarningPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'

function Pause-AndExit {
    Write-Host ""
    Write-Host "PnP Automation check complete." -ForegroundColor Green
    Write-Host "Press any key to close this window..." -ForegroundColor DarkGray
    [void][System.Console]::ReadKey($true)
}

# Required setup notice (suppress return value)
$null = [System.Windows.Forms.MessageBox]::Show(
    "Before continuing:`n`nThis tenant must have the PnP Automation app registered.`nIf it is missing, CloudDonny will create it now.`n`nClosing this window will exit CloudDonny.",
    "Required Setup",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Warning
)

Write-Host "Checking for PnP Automation app registration..." -ForegroundColor Cyan

# --- Graph sign-in (hard gate) ---
try {
    Connect-MgGraph `
        -Scopes "Application.ReadWrite.All","Directory.ReadWrite.All" `
        -NoWelcome
}
catch {
    Write-Host "⚠️ Sign-in was cancelled. Setup cannot continue." -ForegroundColor Red
    exit 1
}

# --- Validate Graph context ---
$context = Get-MgContext
if (-not $context -or -not $context.TenantId) {
    Write-Host "⚠️ No Graph session available. Exiting setup." -ForegroundColor Red
    exit 1
}

$tenantId = $context.TenantId

try {
    # Check if the app already exists
    $existingApp = Get-MgApplication `
        -Filter "displayName eq 'PnP Automation'" `
        -ConsistencyLevel eventual `
        -ErrorAction Stop

    if ($existingApp) {
        Write-Host "✅ PnP Automation app already exists." -ForegroundColor Green
        Write-Host "Application (Client) ID : $($existingApp.AppId)" -ForegroundColor Cyan

        Disconnect-MgGraph | Out-Null
        Pause-AndExit
        exit 0
    }

    # App does not exist → create it
    Write-Host "PnP Automation app not found. Creating..." -ForegroundColor Red

    $newApp = Register-PnPEntraIDApp `
        -ApplicationName "PnP Automation" `
        -Tenant $tenantId `
	-ErrorAction Stop

    Write-Host "✅ App registration completed successfully." -ForegroundColor Green

    Disconnect-MgGraph | Out-Null
    Pause-AndExit
    exit 0
}
catch {
    Write-Host "❌ Setup failed: $($_.Exception.Message)" -ForegroundColor Red
    try { Disconnect-MgGraph | Out-Null } catch {}
    Pause-AndExit
    exit 1
}