<#
Developed by: Rhys Saddul

Overview:
This script sets the timezone, date format, and language for mailboxes in the "bhcs.sfet.org.uk" domain based on a list of UPNs provided in a CSV file.

Issue Context:
Mailboxes in Exchange Online have two time zones; one in regional settings and the other in working hours.
These are both set to Pacific Standard Time by default, regardless of where the mailbox is created or what tenant or Exchange regional settings are in place.

#>

Connect-ExchangeOnline -ShowBanner:$false # --------- Connect to Exchange Online --------- #

# Define your time zone, date format, and language
$Region = "GMT Standard Time"
$DateFormat = "dd/MM/yyyy"
$Language = "en-GB"

# Path to your CSV file containing UPNs (make sure to update the path)
$csvPath = "C:\Users\RhysSaddul\Downloads\BHCS_Staff.csv"

# Import the CSV file, assuming it has a column "UPN" with the users' UPNs
$users = Import-Csv -Path $csvPath

foreach ($user in $users) {
    $UPN = $user.UPN
    Write-Host "Setting time zones and regional settings for mailbox: $UPN" -ForegroundColor Green
    
    # Set Calendar and Regional configurations
    try {
        Set-MailboxCalendarConfiguration -Identity $UPN -WorkingHoursTimeZone $Region
        Set-MailboxRegionalConfiguration -Identity $UPN -TimeZone $Region -Language $Language -DateFormat $DateFormat -Confirm:$False
        Write-Host "Successfully updated settings for $UPN" -ForegroundColor Cyan
    } catch {
        Write-Host "Failed to update settings for $UPN. Error: $_" -ForegroundColor Red
    }
}

Disconnect-ExchangeOnline -Confirm:$false  # --------- Disconnect from Exchange Online --------- #
