<#
Developed by: Rhys Saddul
#>

# Prompt for tenant (default shown)
$tenantDomain = Read-Host "Enter your tenant domain (default: thecloudschool.co.uk)"
if ([string]::IsNullOrWhiteSpace($tenantDomain)) {
    $tenantDomain = "thecloudschool.co.uk"
}

Write-Host "Registering PnP Entra ID App for tenant: $tenantDomain" -ForegroundColor Cyan

try {
    Register-PnPEntraIDApp `
        -ApplicationName "PnP Automation" `
        -Tenant $tenantDomain
        
    Write-Host "✅ App registration completed successfully for $tenantDomain" -ForegroundColor Green
    Write-Host "You can now connect using:" -ForegroundColor Yellow
    Write-Host "Connect-PnPOnline -Url https://$($tenantDomain.Split('.')[0]).sharepoint.com -Interactive" -ForegroundColor White
}
catch {
    Write-Host "❌ ERROR: $($_.Exception.Message)" -ForegroundColor Red
}
