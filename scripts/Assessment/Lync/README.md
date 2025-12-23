# Lync/Skype for Business Assessment Scripts

Comprehensive assessment and reporting tools for Lync/Skype for Business environments. These scripts help IT professionals with infrastructure documentation, health monitoring, migration planning, and operational reporting.

## Available Scripts

### 📊 Interactive Export Tools
- **[Start-LyncCsvExporter.ps1](../../../docs/wiki/Assessments/Lync/Start-LyncCsvExporter.md)** - Menu-driven CSV export tool
  - User data exports (Summary, Voice, SBA, Complete)
  - Phone/device inventory (Common area phones, Analog devices, USB devices)
  - Infrastructure exports (Pools, Policies)
  - Bulk export operations

### 📋 Comprehensive Reports
- **[Get-ComprehensiveLyncReport.ps1](../../../docs/wiki/Assessments/Lync/Get-ComprehensiveLyncReport.md)** - Complete environment assessment
  - Executive summary with key metrics
  - Pool architecture analysis
  - Certificate health monitoring
  - User distribution and database mirror state
  - Infrastructure recommendations

### 🏥 Health & Diagnostics
- **[Get-LyncHealthReport.ps1](../../../docs/wiki/Assessments/Lync/Get-LyncHealthReport.md)** - Health monitoring and diagnostics
  - Certificate expiration tracking
  - Database mirror state analysis
  - Event log error analysis
  - System performance metrics

### 🏗️ Infrastructure Analysis
- **[Get-LyncInfrastructureReport.ps1](../../../docs/wiki/Assessments/Lync/Get-LyncInfrastructureReport.md)** - Infrastructure configuration report
  - Pool categorization (Standard, SBA, IVR, Edge)
  - Computer deployment analysis
  - Service configuration and topology documentation

- **[Get-LyncServiceStatus.ps1](../../../docs/wiki/Assessments/Lync/Get-LyncServiceStatus.md)** - Service status monitoring
  - Windows service analysis
  - Process performance metrics
  - Service dependency mapping

### 👥 User Analysis
- **[Get-LyncUserRegistrationReport.ps1](../../../docs/wiki/Assessments/Lync/Get-LyncUserRegistrationReport.md)** - User registration and activity
  - Registration status and activity tracking
  - Voice enablement statistics
  - SBA user identification

### 🔄 Migration Tools
- **[Export-ADLyncTeamsMigrationData.ps1](../../../docs/wiki/Assessments/Lync/Export-ADLyncTeamsMigrationData.md)** - AD export for Teams migration
  - Active Directory user attributes
  - Lync-specific configuration
  - Voice routing and SIP address export

## Quick Start

### Prerequisites
All scripts require:
- **PowerShell**: 3.0 or higher
- **Lync Management Shell**: For Lync-specific cmdlets
- **Permissions**: CsAdministrator or CsUserAdministrator role
- **Network Access**: Connectivity to Lync Front End servers

### Installation
1. Open Lync Server Management Shell (Start Menu → Lync Server Management Shell)
2. Navigate to this directory
3. Run scripts with appropriate parameters

### Common Usage Patterns

**Daily Health Check**:
```powershell
.\Get-LyncHealthReport.ps1 -PoolFQDN "lyncpool.contoso.com" -OrganizationName "Contoso"
```

**Export User Data**:
```powershell
.\Start-LyncCsvExporter.ps1 -OrganizationName "Contoso"
# Select menu option for desired export
```

**Complete Assessment**:
```powershell
.\Get-ComprehensiveLyncReport.ps1 -PoolFQDN "lyncpool.contoso.com" -OrganizationName "Contoso"
```

**Migration Planning**:
```powershell
# Export AD data
.\Export-ADLyncTeamsMigrationData.ps1 -OrganizationName "Contoso" -ExportToCsv

# Export all Lync data
.\Start-LyncCsvExporter.ps1 -OrganizationName "Contoso"
# Select option 12 (Export Everything)
```

## Output Locations

Default output directories:
- **Text Reports**: `C:\Reports\`
- **CSV Exports**: `C:\Reports\CSV_Exports\`

All scripts support custom output paths via parameters.

## Common Parameters

Most scripts support:
- `-OrganizationName` - Organization name for reports (default: "Organization")
- `-ReportPath` or `-OutputDirectory` - Custom output location
- `-SBAPattern` - Pattern to identify SBA pools (default: "*MSSBA*")

## Troubleshooting

### "Cmdlet not found" errors
Run from Lync Management Shell or import the module:
```powershell
Import-Module "C:\Program Files\Common Files\Skype for Business Server 2015\Modules\SkypeForBusiness\SkypeForBusiness.psd1"
```

### "Access Denied" errors
- Verify CsAdministrator role assignment
- Run PowerShell as Administrator
- Check network connectivity to Front End servers

### Slow performance
- Run during off-peak hours
- Reduce sample sizes: `-SampleUserCount 5`
- Limit time ranges: `-EventLogHours 24`

## Documentation

For detailed documentation on each script, see:
- **[Wiki Documentation](../../../docs/wiki/Assessments/Lync/)** - Comprehensive guides for all scripts
- **Built-in Help** - Use `Get-Help .\ScriptName.ps1 -Full`

## Related Resources
- [Office 365 Assessment Guide](../../../docs/Office365-Assessment-Guide.md)
- [Office 365 Quick Start](../../../docs/Office365-Quick-Start.md)
- [Microsoft Lync Server Documentation](https://docs.microsoft.com/en-us/skypeforbusiness/)
