<#
Developed by: Rhys Saddul
#>

# Prompt to ensure varibles have been setup correctly prior to running the script
$dialogResult = [System.Windows.Forms.MessageBox]::Show("Have you set the varible for the output file path?", "Variable Setup", "YesNo", "Question")
if ($dialogResult -eq "Yes") {
    Write-Host "User confirmed variables have been set. Proceeding with the script." -ForegroundColor Green
} else {
    Write-Host "User confirmed Variables have not been set. Exiting the script." -ForegroundColor Red
    return
}

# Define variables
$MailboxName = "Finance"
$Alias = "Finance2" # If creating multiple finance boxes for more than one domain be sure to increment this nummber
$SubDomain = "@batleyparish.enhanceacad.org.uk"
$Domain = "@enhanceacad.org.uk"
$OnMicrosoft = "@enhanceacademytrust.onmicrosoft.com"

# Varibles that dont need changing
$PrimarySMTPAddress = "$MailboxName$SubDomain"
$NewMicrosoftOnlineServicesID = "$MailboxName$SubDomain"
$RemoveEmailAddresses = @("$Alias$OnMicrosoft", "$Alias$Domain")

# Connect to Exchange Online
	# Get path to the ExchangeOnlineManagement module
	$msalPath = [System.IO.Path]::GetDirectoryName((Get-Module ExchangeOnlineManagement).Path)

	# Load MSAL libraries from the module folder
	Add-Type -Path "$msalPath\Microsoft.IdentityModel.Abstractions.dll"
	Add-Type -Path "$msalPath\Microsoft.Identity.Client.dll"
	
	# Create a Public Client Application using Exchange Online's registered app ID
	[Microsoft.Identity.Client.IPublicClientApplication] $application = `
	    [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create("fb78d390-0c51-40cd-8e17-fdbfab77341b") `
	    .WithDefaultRedirectUri() `
	    .Build()

	# Acquire an access token interactively for Exchange Online
	$result = $application.AcquireTokenInteractive([string[]]"https://outlook.office365.com/.default").ExecuteAsync().Result

	# Connect to Exchange Online using the acquired token
	Connect-ExchangeOnline `
	    -AccessToken $result.AccessToken `
	    -UserPrincipalName $result.Account.Username `
	    -ShowBanner:$False `
	    -ErrorAction Stop

# Create new shared mailbox
New-Mailbox -Name $MailboxName -Alias $Alias -Shared -PrimarySMTPAddress $PrimarySMTPAddress

# Display mailbox info before modification
Get-Mailbox $Alias | Select-Object Name, Alias, UserPrincipalName

# Set new Microsoft Online Services ID
Set-Mailbox $Alias -MicrosoftOnlineServicesID $NewMicrosoftOnlineServicesID

# Display mailbox info after setting Microsoft Online Services ID
Get-Mailbox $Alias | Select-Object Name, Alias, UserPrincipalName

# Display mailbox info including email addresses before modification
Get-Mailbox $Alias | Select-Object Name, Alias, UserPrincipalName, EmailAddresses

# Remove specified email addresses
Set-Mailbox $Alias -EmailAddresses @{remove=$RemoveEmailAddresses}

# Display mailbox info including email addresses after modification
Get-Mailbox $Alias | Select-Object Name, Alias, UserPrincipalName, EmailAddresses
