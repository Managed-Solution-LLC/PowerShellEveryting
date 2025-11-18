# Connect to Exchange Online
Connect-ExchangeOnline

# Array of offboarded user emails
$offboardedUsers = @(
    "user1@ontoso.com",
    "user2@ontoso.com",
    "user3@contoso.com"
)

# Loop through each user and cancel their meetings
foreach ($user in $offboardedUsers) {
    Write-Host "Processing meetings for: $user" -ForegroundColor Cyan
    
    try {
        # Remove all calendar events where this user is the organizer
        Remove-CalendarEvents -Identity $user -CancelOrganizedMeetings -Confirm:$false
        
        Write-Host "Successfully cancelled meetings for: $user" -ForegroundColor Green
    }
    catch {
        Write-Host "Error processing $user : $_" -ForegroundColor Red
    }
}

Write-Host "`nAll done!" -ForegroundColor Green