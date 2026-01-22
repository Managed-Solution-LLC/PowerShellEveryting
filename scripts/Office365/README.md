# Office 365 Management Scripts

PowerShell scripts for Office 365 user management, automation, and administrative tasks.

## 📁 Scripts in this Directory

### New-Office365Accounts.ps1
Creates new user accounts in Microsoft 365 (or Active Directory) from CSV file, array, or individual parameters with automatic password generation and optional OneDrive initialization.

**Quick Start:**
```powershell
# From CSV file
.\New-Office365Accounts.ps1 -CsvPath "C:\Users\NewHires.csv"

# From array
$users = @(
    @{FirstName='John'; LastName='Doe'; EmailAddress='john.doe@contoso.com'; UsageLocation='US'},
    @{FirstName='Jane'; LastName='Smith'; EmailAddress='jane.smith@contoso.com'; UsageLocation='US'}
)
.\New-Office365Accounts.ps1 -UserArray $users

# Single user with OneDrive initialization
.\New-Office365Accounts.ps1 `
    -FirstName "John" `
    -LastName "Doe" `
    -EmailAddress "john.doe@contoso.com" `
    -UsageLocation "US" `
    -InitializeOneDrive
```

**Use Cases:**
- Bulk user account creation for new hires
- Automated onboarding workflows
- Testing environments with multiple accounts
- Migration preparation with pre-created accounts

**Features:**
- CSV batch import or array input
- Secure password generation (12-128 characters)
- Password export to timestamped CSV
- Automatic OneDrive provisioning
- Microsoft 365 or Active Directory support
- Comprehensive validation and error tracking

**Parameters:**
- `CsvPath` - Path to CSV file with user data
- `UserArray` - Array of user objects (hashtables or PSCustomObjects)
- `FirstName`, `LastName`, `EmailAddress` - Individual user parameters (required)
- `DisplayName`, `Password`, `UsageLocation`, `Department`, `JobTitle` - Optional fields
- `AccountType` - "Microsoft365" (default) or "ActiveDirectory"
- `GeneratePasswords` - Force password generation even if provided
- `PasswordLength` - Length of generated passwords (12-128, default: 16)
- `InitializeOneDrive` - Automatically provision OneDrive for new accounts
- `BlockSignIn` - Create accounts but block sign-in initially
- `ForceChangePassword` - Require password change on first sign-in (default: true)

**Output:**
- CSV file: `C:\Reports\AccountCreation\AccountCreation_Results_YYYYMMDD_HHmmss.csv`
- Contains all account details including generated passwords
- OneDrive provisioning status if enabled

---

### Remove-OrganizedMeetings.ps1
Cancels all organized meetings for offboarded users to clean up calendars and send cancellation notices.

**Quick Start:**
```powershell
# Connect to Exchange Online first
Connect-ExchangeOnline

# Array of offboarded users
$offboardedUsers = @(
    "user1@contoso.com",
    "user2@contoso.com"
)

# Cancel their meetings
foreach ($user in $offboardedUsers) {
    Remove-CalendarEvents -Identity $user -CancelOrganizedMeetings -Confirm:$false
}
```

**Use Cases:**
- Offboarding cleanup
- Canceling recurring meetings for departed employees
- Calendar hygiene during transitions

---

## 🚀 Prerequisites

### PowerShell Requirements
- **PowerShell 5.1 or later** (Windows PowerShell or PowerShell 7+)
- **Execution Policy**: RemoteSigned or Unrestricted
  ```powershell
  Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```

### Required Modules

**For New-Office365Accounts.ps1:**
- **Microsoft.Graph.Users** (required)
- **Microsoft.Graph.Files** (for OneDrive initialization)
- **Microsoft.Graph.Sites** (for OneDrive initialization)
- **ActiveDirectory** (optional, for AD account creation)

```powershell
# Install Microsoft Graph modules
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module Microsoft.Graph.Files -Scope CurrentUser
Install-Module Microsoft.Graph.Sites -Scope CurrentUser
```

**For Remove-OrganizedMeetings.ps1:**
- **ExchangeOnlineManagement**

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

### Permissions Required

**New-Office365Accounts.ps1:**
- Microsoft Graph: `User.ReadWrite.All`
- For OneDrive: `Files.ReadWrite.All`, `Sites.ReadWrite.All`
- For AD: Account creation rights in target OU

**Remove-OrganizedMeetings.ps1:**
- Exchange Online: Exchange Administrator or Global Administrator

---

## 📊 Related Resources

### Template Generator
Use the **New-AccountCreationTemplate.ps1** script in `scripts/Assessment/Microsoft365/` to generate a properly formatted CSV template for bulk account creation:

```powershell
..\Assessment\Microsoft365\New-AccountCreationTemplate.ps1
```

This generates `AccountCreation_Template.csv` with all available fields and sample data.

### Sample CSV Format
```csv
FirstName,LastName,EmailAddress,DisplayName,Password,UsageLocation,Department,JobTitle
John,Doe,john.doe@contoso.com,John Doe,,US,IT,Systems Administrator
Jane,Smith,jane.smith@contoso.com,Jane Smith,,US,Sales,Account Manager
```

**Note:** Leave `Password` column empty to auto-generate secure passwords.

---

## 🔧 Common Workflows

### Bulk User Creation with OneDrive
```powershell
# 1. Generate template
..\Assessment\Microsoft365\New-AccountCreationTemplate.ps1 -OutputPath "C:\NewHires.csv"

# 2. Fill in user data (edit CSV in Excel)

# 3. Create accounts with OneDrive
.\New-Office365Accounts.ps1 -CsvPath "C:\NewHires.csv" -InitializeOneDrive

# 4. Secure the password file immediately!
# Output: C:\Reports\AccountCreation\AccountCreation_Results_YYYYMMDD_HHmmss.csv
```

### Array-Based Account Creation
```powershell
# Create from array (useful for automation/API integration)
$newHires = @(
    @{
        FirstName = 'John'
        LastName = 'Doe'
        EmailAddress = 'john.doe@contoso.com'
        UsageLocation = 'US'
        Department = 'IT'
        JobTitle = 'Systems Administrator'
    },
    @{
        FirstName = 'Jane'
        LastName = 'Smith'
        EmailAddress = 'jane.smith@contoso.com'
        UsageLocation = 'US'
        Department = 'Sales'
    }
)

.\New-Office365Accounts.ps1 -UserArray $newHires -GeneratePasswords -PasswordLength 20
```

### Offboarding Cleanup
```powershell
# Cancel meetings for departed employees
Connect-ExchangeOnline

$offboarded = @("departed1@contoso.com", "departed2@contoso.com")

foreach ($user in $offboarded) {
    Remove-CalendarEvents -Identity $user -CancelOrganizedMeetings -Confirm:$false
}
```

---

## 🔒 Security Notes

### Password Security
- Generated passwords meet complexity requirements (uppercase, lowercase, numbers, special characters)
- **CRITICAL**: Secure the output CSV file immediately after creation - it contains plaintext passwords
- Consider encrypting the password file or using secure delivery methods (e.g., secure email, password manager)
- Delete or securely archive password files after distribution

### Best Practices
- Use `-BlockSignIn` to create accounts in disabled state until ready
- Always set `UsageLocation` for accounts that will receive licenses
- Review the output CSV for any creation errors before distributing credentials
- Test in a development tenant first when using array input from automation

---

## 📝 Additional Documentation

For detailed assessment and reporting scripts, see:
- [Assessment Scripts](../Assessment/Microsoft365/README.md)
- [Microsoft 365 Assessment Documentation](../../docs/Office365-Assessment-Guide.md)
