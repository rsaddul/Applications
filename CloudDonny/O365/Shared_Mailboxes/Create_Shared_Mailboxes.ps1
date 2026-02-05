<#
Developed by: Rhys Saddul
#>

# Path to your CSV file
$csvPath   = "C:\Users\RhysSaddul\OneDrive - eduthing\Documents\Migrations\365\Scripts\Shared_Mailboxes\SharedMailboxes.csv"

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
	    -ShowBanner:$False | Out-Null
$mailboxes = Import-Csv -Path $csvPath
$required = 'Name','DisplayName','Alias','PrimarySmtpAddress','FullAccessMembers','SendAsMembers'

foreach ($col in $required) {
    if (-not ($mailboxes | Get-Member -Name $col)) {
        Throw "Missing required CSV column: $col"
    }
}

foreach ($mb in $mailboxes) {
    Write-Host "`n=== Processing: $($mb.Name) ===" -ForegroundColor Cyan

    # 1) Normalise the requested alias (strip domain if present)
    $alias = if ($mb.Alias -match '@') { $mb.Alias.Split('@')[0].Trim() } else { $mb.Alias.Trim() }

    # 2) Look up any existing mailbox object
    $mbx = Get-Mailbox -Identity $alias -ErrorAction SilentlyContinue

    if (-not $mbx) {
        # — doesn’t exist at all → create Shared as per CSV
        Write-Host "➜ Creating Shared mailbox '$($mb.Name)' (Alias: $alias)" -ForegroundColor Green
        try {
            New-Mailbox -Shared `
                -Name               $mb.Name `
                -DisplayName        $mb.DisplayName `
                -Alias              $alias `
                -PrimarySmtpAddress $mb.PrimarySmtpAddress
        }
        catch {
            Write-Warning "    → Creation failed: $($_.Exception.Message)"
            continue
        }
    }
    elseif ($mbx.RecipientTypeDetails -eq 'UserMailbox') {
        # — exists as a User mailbox → warn + append “1” and create that instead
        Write-Warning "User mailbox '$alias' already exists. Appending '1' to create shared mailbox."
        $newAlias          = $alias + '1'
        $newName           = $mb.Name + '1'
        $newDisplayName    = $mb.DisplayName + '1'
        $domain            = $mb.PrimarySmtpAddress.Split('@')[1]
        $newPrimaryAddress = "$newAlias@$domain"

        Write-Host "   → Creating Shared mailbox '$newName' (Alias: $newAlias)" -ForegroundColor Yellow
        try {
            New-Mailbox -Shared `
                -Name               $newName `
                -DisplayName        $newDisplayName `
                -Alias              $newAlias `
                -PrimarySmtpAddress $newPrimaryAddress | Out-Null
        }
        catch {
            Write-Warning "    → Failed to create '$newAlias': $($_.Exception.Message)"
            continue
        }

        # make sure we point the rest of the script at the new alias
        $alias = $newAlias
    }
    elseif ($mbx.RecipientTypeDetails -ne 'SharedMailbox') {
        # — exists but isn’t Shared → convert it
        Write-Host "➜ Converting existing mailbox '$alias' to Shared" -ForegroundColor Yellow
        try {
            Set-Mailbox -Identity $alias -Type Shared | Out-Null
        }
        catch {
            Write-Warning "    → Conversion failed: $($_.Exception.Message)"
            continue
        }
    }
    else {
        # — it’s already a Shared mailbox
        Write-Host "➜ Already a Shared mailbox, skipping creation/conversion." -ForegroundColor DarkGray
    }

    # 3) Grant FullAccess (idempotent)
    if ($mb.FullAccessMembers) {
        foreach ($user in $mb.FullAccessMembers -split ';') {
            $u = $user.Trim()
            $hasFA = Get-MailboxPermission -Identity $alias -User $u -ErrorAction SilentlyContinue |
                     Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited }
            if (-not $hasFA) {
                Write-Host "  • Granting FullAccess to $u"
                Add-MailboxPermission -Identity $alias `
                    -User             $u `
                    -AccessRights     FullAccess `
                    -InheritanceType  All `
                    -Confirm:$false | Out-Null
            }
            else {
                Write-Host "  • FullAccess already granted to $u"
            }
        }
    }

    # 4) Grant SendAs (idempotent)
    if ($mb.SendAsMembers) {
        foreach ($user in $mb.SendAsMembers -split ';') {
            $u = $user.Trim()
            $hasSA = Get-RecipientPermission -Identity $alias -Trustee $u -ErrorAction SilentlyContinue |
                     Where-Object { $_.AccessRights -contains 'SendAs' }
            if (-not $hasSA) {
                Write-Host "  • Granting SendAs to $u"
                Add-RecipientPermission -Identity    $alias `
                                       -Trustee     $u `
                                       -AccessRights SendAs `
                                       -Confirm:$false | Out-Null
            }
            else {
                Write-Host "  • SendAs already granted to $u"
            }
        }
    }
}

Write-Host "`nAll done." -ForegroundColor Magenta