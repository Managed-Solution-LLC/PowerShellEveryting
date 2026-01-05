# PowerShellEveryting Wiki

Welcome to the PowerShellEveryting documentation wiki. This enterprise PowerShell toolkit provides production-ready scripts for Microsoft 365, Azure AD, Teams, Lync/Skype for Business, and Intune management.

## 📚 Documentation Categories

### 🎯 Lync/Skype for Business Assessments
**[Lync Assessment Scripts Overview →](README)**

Complete suite of Lync/Skype for Business assessment and migration tools:
- **[Start-LyncCsvExporter](Start-LyncCsvExporter)** - Interactive menu-driven CSV export tool
- **[Get-ComprehensiveLyncReport](Get-ComprehensiveLyncReport)** - Complete environment assessment with recommendations
- **[Get-LyncHealthReport](Get-LyncHealthReport)** - Health monitoring and diagnostics
- **[Get-LyncInfrastructureReport](Get-LyncInfrastructureReport)** - Infrastructure configuration analysis
- **[Get-LyncServiceStatus](Get-LyncServiceStatus)** - Service status and performance monitoring
- **[Get-LyncUserRegistrationReport](Get-LyncUserRegistrationReport)** - User registration and activity tracking
- **[Export-ADLyncTeamsMigrationData](Export-ADLyncTeamsMigrationData)** - AD export for Teams migration

### 📊 Microsoft 365 Assessments

Office 365 tenant assessment and reporting tools:
- **[Get-QuickO365Report](Get-QuickO365Report)** - Complete O365 assessment with Excel output
- **[Get-MailboxPermissionsReport](Get-MailboxPermissionsReport)** - Mailbox delegation and permissions audit
- **[Get-MailboxRules](Get-MailboxRules)** - Export mailbox rules (forwarding, redirects, auto-replies)
- **[Get-MigrationWizLicensing](Get-MigrationWizLicensing)** - BitTitan MigrationWiz license calculator

### 🖥️ On-Premise Infrastructure Assessments

Windows Server and file share assessment tools:
- **[Start-FileShareAssessment](Start-FileShareAssessment)** - Comprehensive file share assessment with Excel reporting
  - Automatic SMB share discovery
  - Storage analysis and NTFS permission mapping
  - SharePoint/OneDrive compatibility checking
  - Professional Excel report generation

### 🔐 PKI Assessments

Public Key Infrastructure assessment and reporting:
- **[Get-ComprehensivePKIReport](Get-ComprehensivePKIReport)** - Complete PKI environment assessment
- **[Get-PKIHealthReport](Get-PKIHealthReport)** - PKI health monitoring and diagnostics
- **[Merge-PKIAssessmentReports](Merge-PKIAssessmentReports)** - Combine multiple PKI assessment reports

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

[View Documentation →](Get-QuickO365Report)

### Get-MailboxPermissionsReport.ps1
Comprehensive mailbox delegation audit for Full Access, Send As, Send on Behalf, and folder-level permissions.

**Quick Start:**
```powershell
.\Get-MailboxPermissionsReport.ps1 -MailboxFilter SharedMailboxes
```

[View Documentation →](Get-MailboxPermissionsReport)

### Get-MailboxRules.ps1
Export and audit mailbox rules (inbox rules) to identify forwarding rules, auto-replies, folder moves, and automated actions. Essential for security audits and compliance.

**Quick Start:**
```powershell
# All users
.\Get-MailboxRules.ps1

# Specific user
.\Get-MailboxRules.ps1 -UserPrincipalName "user@contoso.com"
```

[View Documentation →](Get-MailboxRules)

### Start-FileShareAssessment.ps1
All-in-one file share assessment tool that discovers SMB shares, analyzes storage and permissions, checks SharePoint/OneDrive compatibility, and generates professional Excel reports.

**Quick Start:**
```powershell
.\Start-FileShareAssessment.ps1 -Domain "Contoso"
```

[View Documentation →](Start-FileShareAssessment)

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
