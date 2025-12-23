# PowerShellEveryting Wiki

Welcome to the PowerShellEveryting documentation wiki. This enterprise PowerShell toolkit provides production-ready scripts for Microsoft 365, Azure AD, Teams, Lync/Skype for Business, and Intune management.

## 📚 Documentation Categories

### 🎯 Lync/Skype for Business Assessments
**[Lync Assessment Scripts Overview →](Assessments/Lync/README)**

Complete suite of Lync/Skype for Business assessment and migration tools:
- **[Start-LyncCsvExporter](Assessments/Lync/Start-LyncCsvExporter)** - Interactive menu-driven CSV export tool
- **[Get-ComprehensiveLyncReport](Assessments/Lync/Get-ComprehensiveLyncReport)** - Complete environment assessment with recommendations
- **[Get-LyncHealthReport](Assessments/Lync/Get-LyncHealthReport)** - Health monitoring and diagnostics
- **[Get-LyncInfrastructureReport](Assessments/Lync/Get-LyncInfrastructureReport)** - Infrastructure configuration analysis
- **[Get-LyncServiceStatus](Assessments/Lync/Get-LyncServiceStatus)** - Service status and performance monitoring
- **[Get-LyncUserRegistrationReport](Assessments/Lync/Get-LyncUserRegistrationReport)** - User registration and activity tracking
- **[Export-ADLyncTeamsMigrationData](Assessments/Lync/Export-ADLyncTeamsMigrationData)** - AD export for Teams migration

### 📊 Microsoft 365 Assessments
Office 365 tenant assessment and reporting tools:
- **[Get-QuickO365Report](Assessments/Microsoft365/Get-QuickO365Report)** - Complete O365 assessment with Excel output
- **[Get-MailboxPermissionsReport](Assessments/Microsoft365/Get-MailboxPermissionsReport)** - Mailbox delegation and permissions audit

### 🔧 Development Resources
- **Code Standards** - PowerShell coding standards and best practices _(coming soon)_
- **Graph Commands** - Microsoft Graph API helpers and utilities _(documentation pending)_

## 🎯 Featured Scripts

### Get-QuickO365Report.ps1
Complete Office 365 tenant assessment collecting mailboxes, licenses, OneDrive, SharePoint, Groups, Teams, and permissions. Generates professional Excel workbook with formatted tables.

**Quick Start:**
```powershell
.\Get-QuickO365Report.ps1 -TenantDomain "contoso"
```

[View Documentation →](Assessments/Microsoft365/Get-QuickO365Report)

### Get-MailboxPermissionsReport.ps1
Comprehensive mailbox delegation audit for Full Access, Send As, Send on Behalf, and folder-level permissions.

**Quick Start:**
```powershell
.\Get-MailboxPermissionsReport.ps1 -MailboxFilter SharedMailboxes
```

[View Documentation →](Assessments/Microsoft365/Get-MailboxPermissionsReport)

## 🚀 Quick Start

1. Clone the repository
2. Review script requirements in comment-based help
3. Install required PowerShell modules
4. Run scripts with appropriate permissions

## 📋 Requirements

- PowerShell 5.1 or later (PowerShell 7+ recommended)
- Appropriate Microsoft 365 admin roles
- Required PowerShell modules (installed automatically by most scripts)

## 🔗 Key Links

- [GitHub Repository](https://github.com/Managed-Solution-LLC/PowerShellEveryting)
- [Contributing Guidelines](https://github.com/Managed-Solution-LLC/PowerShellEveryting/blob/main/CONTRIBUTING.md)
- [License](https://github.com/Managed-Solution-LLC/PowerShellEveryting/blob/main/LICENSE)

## ⚠️ Important Notes

**Client-Agnostic Development**: All public scripts are designed to work with any customer environment. Customer-specific scripts belong in `.prep/` directories only.

**Production Ready**: These scripts are actively used in enterprise IT environments for assessments, migrations, and automation.
