# On-Premise Assessment Scripts

This folder contains scripts for assessing on-premise Windows Server infrastructure, focusing on file share analysis, Active Directory, and server configurations.

## Available Scripts

### File Share Assessment

#### [Start-FileShareAssessment.ps1](Start-FileShareAssessment.ps1)
**Purpose**: Comprehensive file share assessment with Excel reporting

**Features**:
- Automatic SMB share discovery
- Storage analysis (sizes, file counts)
- NTFS permission mapping
- SharePoint/OneDrive compatibility checking
- Excel report generation

**Quick Start**:
```powershell
.\Start-FileShareAssessment.ps1 -Domain "YourDomain"
```

**Documentation**: [Full Documentation](../../../docs/wiki/Assessments/OnPremise/Start-FileShareAssessment.md)

**Typical Use Cases**:
- File server migration planning
- Storage capacity management
- Security audit of file permissions
- SharePoint migration preparation

---

## Prerequisites

### All Scripts
- **PowerShell 5.1 or later**
- **Administrator privileges**
- **Execution Policy**: RemoteSigned or Unrestricted

### File Share Scripts
- **ImportExcel module** (auto-installed)
- **Local access** to file server
- **Long paths enabled** (auto-configured by script)

## Common Parameters

### -Domain
The domain or organization name for the assessment. Used in report naming.

**Example**: `"Contoso"`, `"Lawson"`

### -OutputDirectory
Directory where reports are saved.

**Default**: Current directory  
**Example**: `"C:\Reports"`, `".\Assessments"`

## Quick Start Guide

### 1. File Share Assessment
```powershell
# Basic assessment
.\Start-FileShareAssessment.ps1 -Domain "Contoso"

# Fast assessment (skip permissions)
.\Start-FileShareAssessment.ps1 -Domain "Contoso" -SkipPermissions

# High-performance assessment
.\Start-FileShareAssessment.ps1 -Domain "Contoso" -Workers 200 -OutputDirectory "D:\Reports"
```

## Output Structure

All scripts generate structured output for easy analysis:

```
OutputDirectory/
├── {Domain}_File_Share_Assessment.xlsx  # Main Excel report
├── fileaudit_*.csv                      # Size analysis files
├── unsupported_filenames_*.csv          # Compatibility reports
└── RawData/                             # Detailed permission data
    └── permissions_*.csv
```

## Common Troubleshooting

### "Access to path is denied"
**Solution**: Run as Domain Admin or with appropriate delegated permissions

### "Long paths are not enabled"
**Solution**: Script will prompt to enable. Alternatively:
```powershell
Set-ItemProperty 'HKLM:\System\CurrentControlSet\Control\FileSystem' -Name 'LongPathsEnabled' -Value 1
```

### "Failed to install ImportExcel module"
**Solution**: Manual installation:
```powershell
Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber
```

### Slow Performance
**Solution**: 
- Increase workers: `-Workers 200`
- Skip permissions: `-SkipPermissions`
- Run during off-hours

## Best Practices

### Before Running Assessments
1. ✅ Verify administrator privileges
2. ✅ Check available disk space for reports
3. ✅ Close Excel files in output directory
4. ✅ Test on single share first (use `-ExcludeShares`)
5. ✅ Schedule during off-hours for production servers

### During Execution
1. ✅ Monitor console output for errors
2. ✅ Check progress indicators
3. ✅ Note any access denied warnings

### After Completion
1. ✅ Review Excel report for completeness
2. ✅ Check error and warning counts
3. ✅ Secure reports (contain sensitive data)
4. ✅ Delete temporary CSV files if needed

## Security Considerations

⚠️ **Assessment reports contain sensitive information**:
- User and group permissions
- Folder structures and file names
- Security group memberships

**Recommendations**:
- Store reports in secure locations
- Restrict access to assessment files
- Delete temporary files after review
- Encrypt reports if transmitting

## Performance Guidelines

### Small Environments (< 100 GB)
- **Duration**: 5-15 minutes
- **Settings**: Default parameters

### Medium Environments (100 GB - 1 TB)
- **Duration**: 15-60 minutes
- **Settings**: `-Workers 100`

### Large Environments (> 1 TB)
- **Duration**: 1-4 hours
- **Settings**: `-Workers 200 -SkipPermissions` (initially)

## Related Documentation

- [Start-FileShareAssessment.ps1 Full Documentation](../../../docs/wiki/Assessments/OnPremise/Start-FileShareAssessment.md)
- [File Share Migration Planning Guide](../../../docs/guides/FileShareMigration.md)
- [Assessment Best Practices](../../../docs/guides/AssessmentBestPractices.md)

## Support

For issues, questions, or contributions:
- GitHub Issues: [PowerShellEveryting Issues](https://github.com/Managed-Solution-LLC/PowerShellEveryting/issues)
- Wiki: [Project Wiki](https://github.com/Managed-Solution-LLC/PowerShellEveryting/wiki)

## Version History

- **2026-01-05**: Added Start-FileShareAssessment.ps1 - All-in-one file share assessment with Excel reporting
