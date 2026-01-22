<#
.SYNOPSIS
    Generates a CSV template file for bulk account creation.

.DESCRIPTION
    Creates a properly formatted CSV template file with all available fields
    for use with New-Office365Accounts.ps1. The template includes:
    - Required fields (FirstName, LastName, EmailAddress)
    - Optional fields (DisplayName, Password, UsageLocation, etc.)
    - Sample data row for reference
    - Comments explaining each field
    
    The generated template can be filled out in Excel or any CSV editor
    and then used with New-Office365Accounts.ps1 for batch account creation.

.PARAMETER OutputPath
    Path where the template CSV file will be created.
    Default: Current directory\AccountCreation_Template.csv

.PARAMETER IncludeSampleData
    Include a sample data row in the template for reference.
    Default: $true

.PARAMETER AccountType
    Type of accounts the template is for: "Microsoft365" or "ActiveDirectory".
    This adjusts field relevance hints in the template.
    Default: Microsoft365

.EXAMPLE
    .\New-AccountCreationTemplate.ps1
    
    Creates a template file in the current directory with sample data.

.EXAMPLE
    .\New-AccountCreationTemplate.ps1 -OutputPath "C:\Templates\NewUsers.csv" -IncludeSampleData:$false
    
    Creates a blank template without sample data at the specified path.

.EXAMPLE
    .\New-AccountCreationTemplate.ps1 -AccountType "ActiveDirectory"
    
    Creates a template optimized for Active Directory account creation.

.NOTES
    Author: W. Ford
    Date: 2026-01-22
    Version: 1.0
    
    Requirements:
    - PowerShell 5.1 or later
    
    Template Fields:
    REQUIRED:
    - FirstName: User's first name
    - LastName: User's last name
    - EmailAddress: Email address/User Principal Name
    
    OPTIONAL:
    - DisplayName: Display name (defaults to "FirstName LastName")
    - Password: Account password (generated if blank)
    - UsageLocation: 2-letter country code for M365 licensing (e.g., US, GB)
    - Department: Department name
    - JobTitle: Job title
    - MobilePhone: Mobile phone number
    - OfficeLocation: Office location/building
    - StreetAddress: Street address
    - City: City
    - State: State/Province
    - PostalCode: Postal/ZIP code
    - Country: Country name

.LINK
    https://learn.microsoft.com/en-us/graph/api/user-post-users
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".\AccountCreation_Template.csv",
    
    [Parameter(Mandatory=$false)]
    [bool]$IncludeSampleData = $true,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Microsoft365", "ActiveDirectory")]
    [string]$AccountType = "Microsoft365"
)

$ErrorActionPreference = 'Stop'

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "ACCOUNT CREATION TEMPLATE GENERATOR" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

try {
    # Define template structure
    $templateData = @()
    
    if ($IncludeSampleData) {
        # Add sample data row
        $templateData += [PSCustomObject]@{
            FirstName = "John"
            LastName = "Doe"
            EmailAddress = "john.doe@contoso.com"
            DisplayName = "John Doe"
            Password = ""
            UsageLocation = "US"
            Department = "IT"
            JobTitle = "Systems Administrator"
            MobilePhone = "+1-555-0100"
            OfficeLocation = "Building A"
            StreetAddress = "123 Main Street"
            City = "Seattle"
            State = "WA"
            PostalCode = "98101"
            Country = "United States"
        }
    }
    
    # Add blank row for data entry
    $templateData += [PSCustomObject]@{
        FirstName = ""
        LastName = ""
        EmailAddress = ""
        DisplayName = ""
        Password = ""
        UsageLocation = ""
        Department = ""
        JobTitle = ""
        MobilePhone = ""
        OfficeLocation = ""
        StreetAddress = ""
        City = ""
        State = ""
        PostalCode = ""
        Country = ""
    }
    
    # Export template
    $templateData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "`n✅ Template created successfully!" -ForegroundColor Green
    Write-Host "   Location: $OutputPath" -ForegroundColor White
    
    # Display field information
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "TEMPLATE FIELD REFERENCE" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    Write-Host "`nREQUIRED FIELDS:" -ForegroundColor Yellow
    Write-Host "  • FirstName      - User's first name" -ForegroundColor White
    Write-Host "  • LastName       - User's last name" -ForegroundColor White
    Write-Host "  • EmailAddress   - Email/UPN (must be unique)" -ForegroundColor White
    
    Write-Host "`nOPTIONAL FIELDS:" -ForegroundColor Yellow
    Write-Host "  • DisplayName    - Display name (auto-generated if blank)" -ForegroundColor White
    Write-Host "  • Password       - Account password (auto-generated if blank)" -ForegroundColor White
    
    if ($AccountType -eq "Microsoft365") {
        Write-Host "  • UsageLocation  - 2-letter country code (REQUIRED for licensing)" -ForegroundColor White
    }
    
    Write-Host "  • Department     - Department name" -ForegroundColor White
    Write-Host "  • JobTitle       - Job title" -ForegroundColor White
    Write-Host "  • MobilePhone    - Mobile phone number" -ForegroundColor White
    Write-Host "  • OfficeLocation - Office/building location" -ForegroundColor White
    Write-Host "  • StreetAddress  - Street address" -ForegroundColor White
    Write-Host "  • City           - City name" -ForegroundColor White
    Write-Host "  • State          - State/Province" -ForegroundColor White
    Write-Host "  • PostalCode     - Postal/ZIP code" -ForegroundColor White
    Write-Host "  • Country        - Country name" -ForegroundColor White
    
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "USAGE INSTRUCTIONS" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "`n1. Open the template in Excel or a text editor" -ForegroundColor White
    Write-Host "2. Fill in user information (one user per row)" -ForegroundColor White
    Write-Host "3. Save the file" -ForegroundColor White
    Write-Host "4. Run: .\New-Office365Accounts.ps1 -CsvPath '$OutputPath'" -ForegroundColor White
    
    if ($IncludeSampleData) {
        Write-Host "`n⚠️  Remove or modify the sample data row before use" -ForegroundColor Yellow
    }
    
    Write-Host "`n✅ Template generation complete!`n" -ForegroundColor Green
}
catch {
    Write-Host "`n❌ Error creating template: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
