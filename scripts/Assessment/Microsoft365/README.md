# Microsoft 365 Assessment Scripts

Comprehensive PowerShell scripts for assessing and auditing Microsoft 365 environments, including Exchange Online, mailbox configurations, and user settings.

## 📁 Scripts in this Directory

### [Get-MailboxPermissionsReport.ps1](../../docs/wiki/Assessments-Microsoft365-Get-MailboxPermissionsReport.md)
Exports mailbox delegation permissions including SendAs, SendOnBehalf, and FullAccess rights across all mailboxes.

**Quick Start:**
```powershell
.\Get-MailboxPermissionsReport.ps1
```

**Use Cases:**
- Security audits for delegated access
- Compliance reporting
- Pre-migration permission documentation
- Unauthorized delegation detection

---

### [Get-MailboxRules.ps1](../../docs/wiki/Assessments-Microsoft365-Get-MailboxRules.md)
Retrieves and exports mailbox rules (inbox rules) for all users or specific users, including forwarding rules, auto-replies, and folder moves.

**Quick Start:**
```powershell
# All users
.\Get-MailboxRules.ps1

# Specific user
.\Get-MailboxRules.ps1 -UserPrincipalName "user@contoso.com"
```

**Use Cases:**
- Security audit for auto-forwarding rules
- Compliance documentation
- Troubleshooting missing emails
- Pre-migration rule inventory

---

### [Get-QuickO365Report.ps1](../../docs/wiki/Assessments-Microsoft365-Get-QuickO365Report.md)
Generates a comprehensive snapshot report of Office 365/Microsoft 365 tenant configuration and usage.

**Quick Start:**
```powershell
.\Get-QuickO365Report.ps1
```

**Use Cases:**
- Quick tenant health check
- Executive summary reporting
- Pre-sales assessments
- Regular compliance snapshots

---

## 🚀 Prerequisites

### PowerShell Requirements
- **PowerShell 5.1 or later** (Windows PowerShell or PowerShell 7+)
- **Execution Policy**: RemoteSigned or Unrestricted
  ```powershell
  Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```

### Required Modules
All scripts automatically check for required modules and provide installation instructions if missing.

**Common Modules:**
- **ExchangeOnlineManagement** - Exchange Online cmdlets
- **Microsoft.Graph** - Microsoft Graph API access
- **MSOnline** - Azure AD management (legacy, being phased out)

**Install Modules:**
```powershell
# Exchange Online
Install-Module ExchangeOnlineManagement -Scope CurrentUser

# Microsoft Graph
Install-Module Microsoft.Graph -Scope CurrentUser

# MSOnline (if needed for legacy scripts)
Install-Module MSOnline -Scope CurrentUser
```

### Permissions Required
- **Global Administrator** or **Exchange Administrator** role
- Alternatively, specific roles:
  - Exchange Administrator
  - Security Administrator (for security-related queries)
  - Global Reader (for read-only access)

### System Requirements
- **Operating System**: Windows 10/11, Windows Server 2016+, Linux, macOS
- **Internet Connectivity**: Access to Microsoft 365 services
- **Disk Space**: 100MB - 1GB depending on tenant size
- **Memory**: 4GB+ recommended for large tenants

## 📊 Common Usage Patterns

### Security Audit Workflow
```powershell
# 1. Check mailbox permissions
.\Get-MailboxPermissionsReport.ps1 -OutputDirectory "C:\Audit\Permissions"

# 2. Review mailbox rules (forwarding, redirects)
.\Get-MailboxRules.ps1 -OutputDirectory "C:\Audit\Rules"

# 3. Generate overall tenant report
.\Get-QuickO365Report.ps1 -OutputDirectory "C:\Audit\Overview"
```

### Pre-Migration Assessment
```powershell
# Document current state before migration
$auditDate = Get-Date -Format 'yyyyMMdd'
$outputPath = "C:\Migration\PreMigration_$auditDate"

.\Get-MailboxPermissionsReport.ps1 -OutputDirectory $outputPath
.\Get-MailboxRules.ps1 -OutputDirectory $outputPath
.\Get-QuickO365Report.ps1 -OutputDirectory $outputPath
```

### Monthly Compliance Reporting
```powershell
# Scheduled monthly audit
$month = Get-Date -Format 'yyyy-MM'
$reportPath = "C:\Compliance\Reports\$month"

# Run all assessments
Get-ChildItem "*.ps1" | ForEach-Object {
    & $_.FullName -OutputDirectory $reportPath
}
```

### Troubleshooting Specific User
```powershell
$user = "problematic.user@contoso.com"

# Check their mailbox rules
.\Get-MailboxRules.ps1 -UserPrincipalName $user

# Check permissions on their mailbox
.\Get-MailboxPermissionsReport.ps1 | Where-Object { $_.Mailbox -eq $user }
```

## 🔧 Troubleshooting

### Common Issues Across All Scripts

#### Issue: "Module not found"
**Solution:**
```powershell
# List installed modules
Get-Module -ListAvailable

# Install missing module
Install-Module <ModuleName> -Scope CurrentUser -Force
```

#### Issue: "Access Denied" or "Insufficient Permissions"
**Solutions:**
1. Verify your admin role: `Get-MgUserMemberOf -UserId your.email@contoso.com`
2. Request appropriate permissions from Global Administrator
3. Use delegated admin account if available

#### Issue: "Connection failed" or "Unable to connect"
**Solutions:**
1. Check internet connectivity
2. Verify MFA settings (use app password if needed)
3. Try interactive connection first:
   ```powershell
   Connect-ExchangeOnline
   Connect-MgGraph -Scopes "User.Read.All","Directory.Read.All"
   ```
4. Check for proxy settings: `Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'`

#### Issue: Script hangs or runs slowly
**Solutions:**
1. Check network latency to Microsoft 365
2. Run during off-hours for large tenants
3. Use specific user parameters when available
4. Consider batching for very large tenants

#### Issue: CSV files not created
**Solutions:**
1. Verify output directory permissions
2. Check disk space
3. Review antivirus exclusions
4. Try different output directory:
   ```powershell
   -OutputDirectory "$env:USERPROFILE\Documents\M365Reports"
   ```

### Module Compatibility Notes

**PowerShell 7+ Users:**
- Most modules work natively
- Some legacy modules may require Windows PowerShell compatibility layer
- Use `-UseWindowsPowerShell` parameter when importing if needed

**Windows PowerShell 5.1 Users:**
- All modules fully supported
- Recommended for maximum compatibility
- Default on Windows 10/11 and Server 2016+

## 📖 Best Practices

### Before Running Scripts
1. ✅ **Test in non-production** tenant first (if available)
2. ✅ **Verify permissions** - confirm you have required roles
3. ✅ **Check disk space** - large tenants generate large reports
4. ✅ **Review firewall rules** - ensure access to *.microsoftonline.com
5. ✅ **Plan for MFA** - have authentication device ready

### During Script Execution
1. ✅ **Monitor progress** - watch console output for errors
2. ✅ **Don't interrupt** - let scripts complete (use Ctrl+C only if necessary)
3. ✅ **Check timestamps** - verify current data is being retrieved
4. ✅ **Watch for throttling** - scripts handle this automatically but may slow down

### After Script Completion
1. ✅ **Review output files** - verify completeness
2. ✅ **Check log files** - review any warnings or errors
3. ✅ **Secure reports** - contains sensitive organizational data
4. ✅ **Archive appropriately** - follow data retention policies
5. ✅ **Delete when done** - remove sensitive data after analysis

### Security Best Practices
- 🔒 **Store reports securely** - use encrypted drives or secure storage
- 🔒 **Limit access** - share only with authorized personnel
- 🔒 **Use least privilege** - assign minimum required permissions
- 🔒 **Audit regularly** - schedule periodic assessments
- 🔒 **Clean up** - delete old reports per retention policy

## 📁 Output Structure

All scripts follow consistent output patterns:

```
<OutputDirectory>\
├── MailboxRules_<User>_<Timestamp>.csv
├── MailboxPermissions_<Timestamp>.csv
└── QuickO365Report_<Timestamp>.txt
```

**Filename Patterns:**
- Single user: `<ScriptName>_<Username>_YYYYMMDD_HHmmss.csv`
- All users: `<ScriptName>_AllUsers_YYYYMMDD_HHmmss.csv`
- Reports: `<ScriptName>_YYYYMMDD_HHmmss.txt`

**Default Locations:**
- Most scripts: `C:\Temp\<ScriptSpecificFolder>\`
- Customizable via `-OutputDirectory` parameter

## 🔗 Related Resources

### Internal Documentation
- [Assessment Scripts Overview](../../docs/wiki/Assessments-README.md)
- [Office365 Quick Start Guide](../../docs/Office365-Quick-Start.md)
- [Office365 Project Summary](../../docs/Office365-Project-Summary.md)

### Microsoft Documentation
- [Exchange Online PowerShell](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell)
- [Microsoft Graph PowerShell SDK](https://docs.microsoft.com/en-us/powershell/microsoftgraph/overview)
- [Azure AD PowerShell](https://docs.microsoft.com/en-us/powershell/azure/active-directory/overview)

### Training Resources
- [Exchange Online Administration](https://docs.microsoft.com/en-us/training/modules/manage-exchange-online/)
- [Microsoft 365 Security](https://docs.microsoft.com/en-us/training/modules/m365-security-management/)

## 📝 Contributing

These scripts follow project coding standards outlined in `.github/copilot-instructions.md`:
- **Client-agnostic**: No hardcoded organization names
- **Comprehensive help**: Comment-based help in all scripts
- **Error handling**: Try-catch blocks with detailed messages
- **Progress tracking**: Visual feedback for long operations
- **UTF8 encoding**: CSV exports use UTF8 without BOM

## 📄 Version Information

See individual script documentation for specific version history.

**Last Updated**: 2025-12-23

---

## Quick Reference Card

| Task | Command |
|------|---------|
| Check all mailbox rules | `.\Get-MailboxRules.ps1` |
| Check specific user rules | `.\Get-MailboxRules.ps1 -UserPrincipalName user@domain.com` |
| Export mailbox permissions | `.\Get-MailboxPermissionsReport.ps1` |
| Generate quick tenant report | `.\Get-QuickO365Report.ps1` |
| Custom output directory | Add `-OutputDirectory "C:\Path"` to any script |
| Install required modules | `Install-Module ExchangeOnlineManagement,Microsoft.Graph -Scope CurrentUser` |
| Connect to Exchange | `Connect-ExchangeOnline` |
| Connect to Graph | `Connect-MgGraph -Scopes "User.Read.All"` |

---

**Note**: Scripts in this folder are production-ready and validated for public release. Scripts in `.prep\` subdirectories are work-in-progress and not yet validated.
