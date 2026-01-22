<#
.SYNOPSIS
    Creates new user accounts in Microsoft 365 or Active Directory from CSV file or parameter input.

.DESCRIPTION
    This script automates the creation of new user accounts with support for:
    - Batch creation from CSV file
    - Individual account creation via parameters
    - Automatic password generation with secure storage
    - Microsoft 365 (Microsoft Graph) or local Active Directory
    - Comprehensive validation and error handling
    - Detailed logging and reporting
    
    The script can generate secure random passwords or use provided passwords.
    All created accounts and passwords are exported to a timestamped CSV file
    for record keeping and distribution to users.

.PARAMETER CsvPath
    Path to CSV file containing user account information.
    Use New-AccountCreationTemplate.ps1 to generate a properly formatted template.
    CSV must contain: FirstName, LastName, EmailAddress
    Optional columns: DisplayName, Password, UsageLocation, Department, JobTitle, etc.

.PARAMETER UserArray
    Array of user objects (hashtables or PSCustomObjects) containing user information.
    Each object must contain: FirstName, LastName, EmailAddress
    Optional properties: DisplayName, Password, UsageLocation, Department, JobTitle, etc.

.PARAMETER FirstName
    First name of the user. Required when not using CSV input.

.PARAMETER LastName
    Last name of the user. Required when not using CSV input.

.PARAMETER EmailAddress
    Email address/User Principal Name for the account. Required when not using CSV input.

.PARAMETER DisplayName
    Display name for the user. If not provided, defaults to "FirstName LastName".

.PARAMETER Password
    Password for the account. If not provided, a secure random password will be generated.

.PARAMETER UsageLocation
    Two-letter country code for license assignment (e.g., "US", "GB", "CA").
    Required for Microsoft 365 accounts that will receive licenses.

.PARAMETER Department
    Department name for the user.

.PARAMETER JobTitle
    Job title for the user.

.PARAMETER AccountType
    Type of account to create: "Microsoft365" or "ActiveDirectory".
    Default: Microsoft365

.PARAMETER OutputDirectory
    Directory where output files will be saved.
    Default: C:\Reports\AccountCreation

.PARAMETER GeneratePasswords
    Force password generation even if passwords are provided in CSV.

.PARAMETER PasswordLength
    Length of generated passwords. Must be between 12 and 128 characters.
    Default: 16

.PARAMETER ForceChangePassword
    Require users to change password on first sign-in.
    Default: $true

.PARAMETER BlockSignIn
    Create accounts but block sign-in until ready to activate.
    Default: $false

.PARAMETER InitializeOneDrive
    Automatically provision and initialize OneDrive for each created account.
    Only applicable for Microsoft 365 accounts.
    Default: $false

.EXAMPLE
    .\New-Office365Accounts.ps1 -CsvPath "C:\Users\NewHires.csv"
    
    Creates accounts from CSV file using Microsoft Graph. Generates passwords
    for any accounts without passwords specified.

.EXAMPLE
    .\New-Office365Accounts.ps1 -FirstName "John" -LastName "Doe" -EmailAddress "john.doe@contoso.com" -UsageLocation "US"
    
    Creates a single Microsoft 365 account with generated password.

.EXAMPLE
    $users = @(
        @{FirstName='John'; LastName='Doe'; EmailAddress='john.doe@contoso.com'; UsageLocation='US'},
        @{FirstName='Jane'; LastName='Smith'; EmailAddress='jane.smith@contoso.com'; UsageLocation='US'}
    )
    .\New-Office365Accounts.ps1 -UserArray $users
    
    Creates accounts from an array of hashtables.

.EXAMPLE
    .\New-Office365Accounts.ps1 -CsvPath "C:\Users\NewHires.csv" -AccountType "ActiveDirectory"
    
    Creates Active Directory accounts from CSV file.

.EXAMPLE
    .\New-Office365Accounts.ps1 -CsvPath "C:\Users\NewHires.csv" -GeneratePasswords -PasswordLength 20
    
    Creates accounts and generates 20-character passwords for all users,
    ignoring any passwords in the CSV file.

.EXAMPLE
    .\New-Office365Accounts.ps1 -CsvPath "C:\Users\NewHires.csv" -InitializeOneDrive
    
    Creates accounts and automatically provisions OneDrive for each user.

.NOTES
    Author: W. Ford
    Date: 2026-01-22
    Version: 1.0
    
    Requirements:
    - For Microsoft 365: Microsoft.Graph.Users module
    - For OneDrive initialization: Microsoft.Graph.Files and Microsoft.Graph.Sites modules
    - For Active Directory: ActiveDirectory module
    - PowerShell 5.1 or later
    - Appropriate permissions:
      * M365: User.ReadWrite.All permission in Microsoft Graph
      * AD: Account creation rights in target OU
    
    Output:
    - CSV file with created accounts and passwords (timestamp in filename)
    - Detailed console logging with color-coded status messages
    
    Security:
    - Generated passwords meet complexity requirements
    - Password export file should be secured immediately after creation
    - Consider encrypting password file or using secure delivery method

.LINK
    https://learn.microsoft.com/en-us/graph/api/user-post-users
    https://learn.microsoft.com/en-us/powershell/module/activedirectory/new-aduser
#>

[CmdletBinding(DefaultParameterSetName = 'CSV')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='CSV', HelpMessage="Path to CSV file with user information")]
    [ValidateScript({
        if (Test-Path $_) { $true }
        else { throw "CSV file not found: $_" }
    })]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$true, ParameterSetName='Array', HelpMessage="Array of user objects")]
    [ValidateNotNullOrEmpty()]
    [array]$UserArray,
    
    [Parameter(Mandatory=$true, ParameterSetName='Single', HelpMessage="First name of the user")]
    [ValidateNotNullOrEmpty()]
    [string]$FirstName,
    
    [Parameter(Mandatory=$true, ParameterSetName='Single', HelpMessage="Last name of the user")]
    [ValidateNotNullOrEmpty()]
    [string]$LastName,
    
    [Parameter(Mandatory=$true, ParameterSetName='Single', HelpMessage="Email address/UPN for the account")]
    [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
    [string]$EmailAddress,
    
    [Parameter(Mandatory=$false, ParameterSetName='Single')]
    [string]$DisplayName,
    
    [Parameter(Mandatory=$false, ParameterSetName='Single')]
    [ValidateLength(8, 256)]
    [string]$Password,
    
    [Parameter(Mandatory=$false, ParameterSetName='Single')]
    [ValidateLength(2, 2)]
    [string]$UsageLocation,
    
    [Parameter(Mandatory=$false, ParameterSetName='Single')]
    [string]$Department,
    
    [Parameter(Mandatory=$false, ParameterSetName='Single')]
    [string]$JobTitle,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Microsoft365", "ActiveDirectory")]
    [string]$AccountType = "Microsoft365",
    
    [Parameter(Mandatory=$false)]
    [string]$OutputDirectory = "C:\Reports\AccountCreation",
    
    [Parameter(Mandatory=$false)]
    [switch]$GeneratePasswords,
    
    [Parameter(Mandatory=$false)]
    [ValidateRange(12, 128)]
    [int]$PasswordLength = 16,
    
    [Parameter(Mandatory=$false)]
    [bool]$ForceChangePassword = $true,
    
    [Parameter(Mandatory=$false)]
    [switch]$BlockSignIn,
    
    [Parameter(Mandatory=$false)]
    [switch]$InitializeOneDrive
)

#region Initialization

$ErrorActionPreference = 'Stop'
$StartTime = Get-Date
$Separator = "=" * 80
$SubSeparator = "-" * 60

# Statistics tracking
$SuccessCount = 0
$ErrorCount = 0
$WarningCount = 0
$CreatedAccounts = @()

Write-Host "`n$Separator" -ForegroundColor Cyan
Write-Host "ACCOUNT CREATION UTILITY" -ForegroundColor Cyan
Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host $Separator -ForegroundColor Cyan

#endregion

#region Functions

function Write-StatusMessage {
    param(
        [string]$Message,
        [ValidateSet("Info", "Success", "Warning", "Error")]
        [string]$Type = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    switch ($Type) {
        "Error" { 
            Write-Host "[$timestamp] ❌ ERROR: $Message" -ForegroundColor Red
            $script:ErrorCount++
        }
        "Warning" { 
            Write-Host "[$timestamp] ⚠️  WARNING: $Message" -ForegroundColor Yellow
            $script:WarningCount++
        }
        "Success" { 
            Write-Host "[$timestamp] ✅ SUCCESS: $Message" -ForegroundColor Green
            $script:SuccessCount++
        }
        default { 
            Write-Host "[$timestamp] INFO: $Message" -ForegroundColor Cyan
        }
    }
}

function New-SecurePassword {
    param(
        [int]$Length = 16
    )
    
    # Character sets for password generation
    $uppercase = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $lowercase = 'abcdefghijklmnopqrstuvwxyz'
    $numbers = '0123456789'
    $special = '!@#$%^&*-_=+[]{}|:,.<>?'
    
    # Ensure at least one character from each set
    $password = @()
    $password += $uppercase[(Get-Random -Maximum $uppercase.Length)]
    $password += $lowercase[(Get-Random -Maximum $lowercase.Length)]
    $password += $numbers[(Get-Random -Maximum $numbers.Length)]
    $password += $special[(Get-Random -Maximum $special.Length)]
    
    # Fill remaining length with random characters from all sets
    $allChars = $uppercase + $lowercase + $numbers + $special
    for ($i = $password.Count; $i -lt $Length; $i++) {
        $password += $allChars[(Get-Random -Maximum $allChars.Length)]
    }
    
    # Shuffle the password
    $shuffled = $password | Get-Random -Count $password.Count
    return -join $shuffled
}

function Test-ModuleAvailability {
    param([string]$ModuleName)
    
    if (-not (Get-Module -Name $ModuleName -ListAvailable)) {
        Write-StatusMessage "Required module '$ModuleName' is not installed" -Type Error
        Write-Host "   Install with: Install-Module $ModuleName -Scope CurrentUser" -ForegroundColor Yellow
        return $false
    }
    return $true
}

function Connect-ToMicrosoft365 {
    Write-StatusMessage "Connecting to Microsoft Graph..." -Type Info
    
    try {
        # Import required modules
        Import-Module Microsoft.Graph.Users -ErrorAction Stop
        
        if ($InitializeOneDrive.IsPresent) {
            Import-Module Microsoft.Graph.Files -ErrorAction Stop
            Import-Module Microsoft.Graph.Sites -ErrorAction Stop
        }
        
        # Build scopes list
        $scopes = @("User.ReadWrite.All")
        if ($InitializeOneDrive.IsPresent) {
            $scopes += "Files.ReadWrite.All"
            $scopes += "Sites.ReadWrite.All"
        }
        
        # Connect with required permissions
        Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
        
        # Verify connection
        $context = Get-MgContext
        if ($context) {
            Write-StatusMessage "Connected to Microsoft Graph as $($context.Account)" -Type Success
            return $true
        }
        else {
            Write-StatusMessage "Failed to verify Microsoft Graph connection" -Type Error
            return $false
        }
    }
    catch {
        Write-StatusMessage "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Type Error
        return $false
    }
}

function Connect-ToActiveDirectory {
    Write-StatusMessage "Verifying Active Directory module..." -Type Info
    
    try {
        # Import module
        Import-Module ActiveDirectory -ErrorAction Stop
        
        # Test connection by querying domain
        $domain = Get-ADDomain -ErrorAction Stop
        Write-StatusMessage "Connected to Active Directory domain: $($domain.DNSRoot)" -Type Success
        return $true
    }
    catch {
        Write-StatusMessage "Failed to connect to Active Directory: $($_.Exception.Message)" -Type Error
        return $false
    }
}

function New-Microsoft365User {
    param(
        [hashtable]$UserData
    )
    
    try {
        # Build user object
        $userParams = @{
            DisplayName = $UserData.DisplayName
            GivenName = $UserData.FirstName
            Surname = $UserData.LastName
            UserPrincipalName = $UserData.EmailAddress
            MailNickname = $UserData.EmailAddress.Split('@')[0]
            AccountEnabled = -not $BlockSignIn.IsPresent
            PasswordProfile = @{
                Password = $UserData.Password
                ForceChangePasswordNextSignIn = $ForceChangePassword
            }
        }
        
        # Add optional fields
        if ($UserData.UsageLocation) { $userParams.UsageLocation = $UserData.UsageLocation }
        if ($UserData.Department) { $userParams.Department = $UserData.Department }
        if ($UserData.JobTitle) { $userParams.JobTitle = $UserData.JobTitle }
        if ($UserData.MobilePhone) { $userParams.MobilePhone = $UserData.MobilePhone }
        if ($UserData.OfficeLocation) { $userParams.OfficeLocation = $UserData.OfficeLocation }
        if ($UserData.StreetAddress) { $userParams.StreetAddress = $UserData.StreetAddress }
        if ($UserData.City) { $userParams.City = $UserData.City }
        if ($UserData.State) { $userParams.State = $UserData.State }
        if ($UserData.PostalCode) { $userParams.PostalCode = $UserData.PostalCode }
        if ($UserData.Country) { $userParams.Country = $UserData.Country }
        
        # Create user
        $newUser = New-MgUser @userParams -ErrorAction Stop
        
        Write-StatusMessage "Created Microsoft 365 user: $($UserData.EmailAddress)" -Type Success
        
        return @{
            Success = $true
            UserId = $newUser.Id
            UserPrincipalName = $newUser.UserPrincipalName
        }
    }
    catch {
        Write-StatusMessage "Failed to create user $($UserData.EmailAddress): $($_.Exception.Message)" -Type Error
        return @{
            Success = $false
            Error = $_.Exception.Message
        }
    }
}

function New-ADUserAccount {
    param(
        [hashtable]$UserData
    )
    
    try {
        # Build user parameters
        $userParams = @{
            GivenName = $UserData.FirstName
            Surname = $UserData.LastName
            Name = $UserData.DisplayName
            DisplayName = $UserData.DisplayName
            SamAccountName = $UserData.EmailAddress.Split('@')[0]
            UserPrincipalName = $UserData.EmailAddress
            EmailAddress = $UserData.EmailAddress
            AccountPassword = (ConvertTo-SecureString $UserData.Password -AsPlainText -Force)
            Enabled = -not $BlockSignIn.IsPresent
            ChangePasswordAtLogon = $ForceChangePassword
        }
        
        # Add optional fields
        if ($UserData.Department) { $userParams.Department = $UserData.Department }
        if ($UserData.JobTitle) { $userParams.Title = $UserData.JobTitle }
        if ($UserData.MobilePhone) { $userParams.MobilePhone = $UserData.MobilePhone }
        if ($UserData.OfficeLocation) { $userParams.Office = $UserData.OfficeLocation }
        if ($UserData.StreetAddress) { $userParams.StreetAddress = $UserData.StreetAddress }
        if ($UserData.City) { $userParams.City = $UserData.City }
        if ($UserData.State) { $userParams.State = $UserData.State }
        if ($UserData.PostalCode) { $userParams.PostalCode = $UserData.PostalCode }
        if ($UserData.Country) { $userParams.Country = $UserData.Country }
        
        # Create user
        $newUser = New-ADUser @userParams -PassThru -ErrorAction Stop
        
        Write-StatusMessage "Created Active Directory user: $($UserData.EmailAddress)" -Type Success
        
        return @{
            Success = $true
            UserId = $newUser.ObjectGUID
            UserPrincipalName = $newUser.UserPrincipalName
        }
    }
    catch {
        Write-StatusMessage "Failed to create AD user $($UserData.EmailAddress): $($_.Exception.Message)" -Type Error
        return @{
            Success = $false
            Error = $_.Exception.Message
        }
    }
}

function Initialize-UserOneDrive {
    param(
        [string]$UserId,
        [string]$UserPrincipalName
    )
    
    try {
        Write-Host "   Initializing OneDrive..." -ForegroundColor Cyan -NoNewline
        
        # Provision OneDrive by accessing the drive endpoint
        # This triggers OneDrive provisioning if it doesn't exist
        $drive = Get-MgUserDrive -UserId $UserId -ErrorAction Stop
        
        if ($drive) {
            Write-Host " ✅" -ForegroundColor Green
            Write-StatusMessage "OneDrive initialized for $UserPrincipalName" -Type Success
            return @{
                Success = $true
                DriveId = $drive.Id
                WebUrl = $drive.WebUrl
            }
        }
        else {
            Write-Host " ⚠️" -ForegroundColor Yellow
            Write-StatusMessage "OneDrive provisioning may be pending for $UserPrincipalName" -Type Warning
            return @{
                Success = $false
                Pending = $true
            }
        }
    }
    catch {
        Write-Host " ❌" -ForegroundColor Red
        
        # OneDrive may not be immediately available - this is often expected
        if ($_.Exception.Message -like "*Request_ResourceNotFound*" -or 
            $_.Exception.Message -like "*does not exist*") {
            Write-StatusMessage "OneDrive not yet available for $UserPrincipalName - provisioning initiated" -Type Warning
            return @{
                Success = $false
                Pending = $true
            }
        }
        else {
            Write-StatusMessage "Failed to initialize OneDrive for ${UserPrincipalName}: $($_.Exception.Message)" -Type Error
            return @{
                Success = $false
                Error = $_.Exception.Message
            }
        }
    }
}

function Export-AccountResults {
    param(
        [array]$Accounts
    )
    
    if ($Accounts.Count -eq 0) {
        Write-StatusMessage "No accounts to export" -Type Warning
        return
    }
    
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outputFile = Join-Path $OutputDirectory "AccountCreation_Results_$timestamp.csv"
    
    try {
        $Accounts | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
        Write-Host "`n✅ Exported $($Accounts.Count) account(s) to: $outputFile" -ForegroundColor Green
        Write-Host "   ⚠️  IMPORTANT: Secure this file - it contains passwords!" -ForegroundColor Yellow
    }
    catch {
        Write-StatusMessage "Failed to export results: $($_.Exception.Message)" -Type Error
    }
}

#endregion

#region Pre-Execution Checks

# Validate PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-StatusMessage "This script requires PowerShell 5.1 or later" -Type Error
    exit 1
}

# Create output directory
if (!(Test-Path $OutputDirectory)) {
    try {
        New-Item -ItemType Directory -Path $OutputDirectory -Force -ErrorAction Stop | Out-Null
        Write-StatusMessage "Created output directory: $OutputDirectory" -Type Success
    }
    catch {
        Write-StatusMessage "Cannot create output directory: $($_.Exception.Message)" -Type Error
        exit 1
    }
}

# Verify module availability based on account type
if ($AccountType -eq "Microsoft365") {
    if (-not (Test-ModuleAvailability -ModuleName "Microsoft.Graph.Users")) {
        exit 1
    }
}
elseif ($AccountType -eq "ActiveDirectory") {
    if (-not (Test-ModuleAvailability -ModuleName "ActiveDirectory")) {
        exit 1
    }
}

#endregion

#region Main Processing

try {
    # Connect to appropriate service
    if ($AccountType -eq "Microsoft365") {
        if (-not (Connect-ToMicrosoft365)) {
            exit 1
        }
    }
    else {
        if (-not (Connect-ToActiveDirectory)) {
            exit 1
        }
    }
    
    # Build user list
    $usersToCreate = @()
    
    if ($PSCmdlet.ParameterSetName -eq 'CSV') {
        Write-StatusMessage "Loading users from CSV: $CsvPath" -Type Info
        
        try {
            $csvData = Import-Csv -Path $CsvPath -ErrorAction Stop
            
            # Validate required columns
            $requiredColumns = @('FirstName', 'LastName', 'EmailAddress')
            $csvColumns = $csvData[0].PSObject.Properties.Name
            $missingColumns = $requiredColumns | Where-Object { $_ -notin $csvColumns }
            
            if ($missingColumns) {
                Write-StatusMessage "CSV missing required columns: $($missingColumns -join ', ')" -Type Error
                exit 1
            }
            
            Write-StatusMessage "Loaded $($csvData.Count) user(s) from CSV" -Type Success
            $usersToCreate = $csvData
        }
        catch {
            Write-StatusMessage "Failed to load CSV: $($_.Exception.Message)" -Type Error
            exit 1
        }
    }
    elseif ($PSCmdlet.ParameterSetName -eq 'Array') {
        Write-StatusMessage "Processing user array with $($UserArray.Count) user(s)" -Type Info
        
        # Validate array has required properties
        $requiredProps = @('FirstName', 'LastName', 'EmailAddress')
        
        foreach ($user in $UserArray) {
            # Convert hashtables to PSCustomObjects for consistent handling
            if ($user -is [hashtable]) {
                $user = [PSCustomObject]$user
            }
            
            # Validate required properties
            $userProps = $user.PSObject.Properties.Name
            $missingProps = $requiredProps | Where-Object { $_ -notin $userProps }
            
            if ($missingProps) {
                Write-StatusMessage "User object missing required properties: $($missingProps -join ', ')" -Type Error
                Write-StatusMessage "User: $($user | ConvertTo-Json -Compress)" -Type Error
                exit 1
            }
        }
        
        Write-StatusMessage "Validated $($UserArray.Count) user object(s)" -Type Success
        $usersToCreate = $UserArray
    }
    else {
        # Single user from parameters
        $usersToCreate = @([PSCustomObject]@{
            FirstName = $FirstName
            LastName = $LastName
            EmailAddress = $EmailAddress
            DisplayName = $DisplayName
            Password = $Password
            UsageLocation = $UsageLocation
            Department = $Department
            JobTitle = $JobTitle
        })
    }
    
    Write-Host "`n$SubSeparator" -ForegroundColor Cyan
    Write-StatusMessage "Processing $($usersToCreate.Count) user account(s)..." -Type Info
    Write-Host $SubSeparator -ForegroundColor Cyan
    
    # Process each user
    foreach ($user in $usersToCreate) {
        # Build user data hashtable
        $userData = @{
            FirstName = $user.FirstName.Trim()
            LastName = $user.LastName.Trim()
            EmailAddress = $user.EmailAddress.Trim().ToLower()
        }
        
        # Set display name
        if ([string]::IsNullOrWhiteSpace($user.DisplayName)) {
            $userData.DisplayName = "$($userData.FirstName) $($userData.LastName)"
        }
        else {
            $userData.DisplayName = $user.DisplayName.Trim()
        }
        
        # Handle password
        if ($GeneratePasswords.IsPresent -or [string]::IsNullOrWhiteSpace($user.Password)) {
            $userData.Password = New-SecurePassword -Length $PasswordLength
            $passwordGenerated = $true
        }
        else {
            $userData.Password = $user.Password
            $passwordGenerated = $false
        }
        
        # Add optional fields if present
        $optionalFields = @('UsageLocation', 'Department', 'JobTitle', 'MobilePhone', 
                           'OfficeLocation', 'StreetAddress', 'City', 'State', 
                           'PostalCode', 'Country')
        
        foreach ($field in $optionalFields) {
            if (![string]::IsNullOrWhiteSpace($user.$field)) {
                $userData.$field = $user.$field.Trim()
            }
        }
        
        # Create account
        Write-Host "`nProcessing: $($userData.DisplayName) ($($userData.EmailAddress))" -ForegroundColor Yellow
        
        if ($AccountType -eq "Microsoft365") {
            $result = New-Microsoft365User -UserData $userData            
            # Initialize OneDrive if requested and account was created successfully
            if ($result.Success -and $InitializeOneDrive.IsPresent) {
                Start-Sleep -Seconds 2  # Brief delay for account provisioning
                $oneDriveResult = Initialize-UserOneDrive -UserId $result.UserId -UserPrincipalName $result.UserPrincipalName
                $result.OneDriveInitialized = $oneDriveResult.Success
                $result.OneDrivePending = $oneDriveResult.Pending
                $result.OneDriveUrl = $oneDriveResult.WebUrl
            }        }
        else {
            $result = New-ADUserAccount -UserData $userData
        }
        
        # Build result object
        $accountResult = [PSCustomObject]@{
            DisplayName = $userData.DisplayName
            FirstName = $userData.FirstName
            LastName = $userData.LastName
            EmailAddress = $userData.EmailAddress
            Password = $userData.Password
            PasswordGenerated = $passwordGenerated
            AccountType = $AccountType
            Created = $result.Success
            UserId = if ($result.Success) { $result.UserId } else { $null }
            OneDriveInitialized = if ($InitializeOneDrive.IsPresent -and $result.Success) { $result.OneDriveInitialized } else { $null }
            OneDrivePending = if ($InitializeOneDrive.IsPresent -and $result.Success) { $result.OneDrivePending } else { $null }
            OneDriveUrl = if ($InitializeOneDrive.IsPresent -and $result.Success) { $result.OneDriveUrl } else { $null }
            Error = if ($result.Success) { $null } else { $result.Error }
            Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
        
        # Add optional fields to result
        foreach ($field in $optionalFields) {
            if ($userData.ContainsKey($field)) {
                $accountResult | Add-Member -MemberType NoteProperty -Name $field -Value $userData.$field
            }
        }
        
        $CreatedAccounts += $accountResult
    }
    
    # Export results
    Write-Host "`n$SubSeparator" -ForegroundColor Cyan
    Export-AccountResults -Accounts $CreatedAccounts
    
    # Summary
    Write-Host "`n$Separator" -ForegroundColor Cyan
    Write-Host "ACCOUNT CREATION SUMMARY" -ForegroundColor Cyan
    Write-Host $Separator -ForegroundColor Cyan
    
    $successfulCreations = ($CreatedAccounts | Where-Object { $_.Created }).Count
    $failedCreations = ($CreatedAccounts | Where-Object { -not $_.Created }).Count
    
    Write-Host "Total Processed: $($CreatedAccounts.Count)" -ForegroundColor White
    Write-Host "Successfully Created: $successfulCreations" -ForegroundColor Green
    Write-Host "Failed: $failedCreations" -ForegroundColor $(if($failedCreations -gt 0){'Red'}else{'Green'})
    Write-Host "Warnings: $WarningCount" -ForegroundColor $(if($WarningCount -gt 0){'Yellow'}else{'Green'})
    
    $endTime = Get-Date
    $duration = $endTime - $StartTime
    Write-Host "`nCompleted in $($duration.ToString('mm\:ss'))" -ForegroundColor Cyan
    Write-Host $Separator -ForegroundColor Cyan
}
catch {
    Write-StatusMessage "Critical error: $($_.Exception.Message)" -Type Error
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    exit 1
}
finally {
    # Cleanup connections
    if ($AccountType -eq "Microsoft365") {
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            Write-Host "`nDisconnected from Microsoft Graph" -ForegroundColor Gray
        }
        catch {
            # Ignore disconnect errors
        }
    }
}

#endregion
