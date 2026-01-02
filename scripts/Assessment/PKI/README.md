# PKI Assessment Scripts

This folder contains PowerShell scripts for comprehensive Public Key Infrastructure (PKI) management: deep assessment, operational health monitoring, and enterprise-wide consolidation across multiple Certificate Authority (CA) servers.

## Overview

The PKI script suite provides end-to-end visibility into Certificate Authority infrastructure:
- **Assessment**: Deep analysis of certificates, templates, and permissions
- **Health Monitoring**: Operational validation with automated scoring
- **Consolidation**: Multi-CA data aggregation into single Excel workbook

Perfect for enterprise PKI environments with multiple CA servers requiring regular assessment, compliance reporting, and operational monitoring.

## Available Scripts

### Get-ComprehensivePKIReport.ps1
**Purpose**: Deep PKI infrastructure assessment  
**Best For**: Quarterly audits, migration planning, compliance reporting

**Features:**
- Export all issued certificates from CA database
- Analyze certificate expiration and validity status
- Export all certificate templates with full properties
- Extract and analyze template permissions (ACLs)
- Flag certificates expiring within threshold
- Generate detailed CSV exports and comprehensive text reports

**Quick Start:**
```powershell
# Basic assessment with auto-discovery
.\Get-ComprehensivePKIReport.ps1

# Specify CA server and include revoked certificates
.\Get-ComprehensivePKIReport.ps1 -CAServerName "CA01.contoso.com" -IncludeRevokedCertificates

# Custom output directory and expiration threshold
.\Get-ComprehensivePKIReport.ps1 -OutputDirectory "D:\Reports" -DaysToExpiration 30
```

[Full Documentation](../../../docs/wiki/Assessments/PKI/Get-ComprehensivePKIReport.md)

---

### Get-PKIHealthReport.ps1
**Purpose**: Operational health monitoring with automated scoring  
**Best For**: Weekly health checks, troubleshooting, SLA monitoring

**Features:**
- Certificate Services status and responsiveness validation
- CA certificate expiration tracking and alerting
- Database health metrics (size, records, statistics)
- CRL publication validation and distribution point testing
- AIA distribution point accessibility verification
- Certificate template availability in Active Directory
- Event log analysis (errors and warnings)
- Automated health scoring (0-100) with weighted issue detection
- Actionable recommendations based on findings

**Quick Start:**
```powershell
# Full health assessment
.\Get-PKIHealthReport.ps1

# Extended monitoring with custom thresholds
.\Get-PKIHealthReport.ps1 -DaysToExpiration 15 -EventLogHours 48

# Health check with custom output
.\Get-PKIHealthReport.ps1 -OutputDirectory "D:\PKI_Health" -OrganizationName "Contoso"
```

[Full Documentation](../../../docs/wiki/Assessments/PKI/Get-PKIHealthReport.md)

---

### Merge-PKIAssessmentReports.ps1
**Purpose**: Multi-CA data consolidation into single Excel workbook  
**Best For**: Enterprise reporting, executive summaries, cross-CA analysis

**Features:**
- Auto-discovery of latest assessments from multiple CAs
- Combines certificates, templates, permissions, and health data
- Generates professional Excel workbook with 10 worksheets
- Adds CAServer column for filtering and pivot analysis
- Parses health scores and operational metrics
- Supports UNC paths for centralized reporting

**Quick Start:**
```powershell
# Consolidate from default locations
.\Merge-PKIAssessmentReports.ps1

# Enterprise consolidation with UNC paths
.\Merge-PKIAssessmentReports.ps1 `
    -AssessmentPath "\\FileServer\PKI\Assessment" `
    -HealthPath "\\FileServer\PKI\Health" `
    -OutputDirectory "\\FileServer\PKI\Reports" `
    -OrganizationName "Contoso"
```

[Full Documentation](../../../docs/wiki/Assessments/PKI/Merge-PKIAssessmentReports.md)

---

## Prerequisites

### Assessment & Health Scripts
- **Windows Server** with Certificate Authority role installed
- **Administrator privileges** on the CA server
- **PowerShell 5.1 or later**
- **Certificate Services** must be running
- **Network access** to distribution points (for health checks)

Optional but recommended:
- **ActiveDirectory PowerShell module** for enhanced template permissions analysis

### Merge Script
- **PowerShell 5.1 or later**
- **ImportExcel module** (auto-installs if missing)
- **Read access** to assessment and health folders

## Common Prerequisites Installation

```powershell
# Install RSAT-ADCS feature (on CA servers)
Install-WindowsFeature RSAT-ADCS

# Import the module
Import-Module ADCS-Administration

# Install ImportExcel module (for Merge script)
Install-Module ImportExcel -Scope CurrentUser -Force

# Verify CA connectivity
certutil -ping
```

## Typical Workflows

### Workflow 1: Single CA Assessment
```powershell
# Step 1: Deep assessment
.\Get-ComprehensivePKIReport.ps1 -OrganizationName "Contoso"

# Step 2: Health check
.\Get-PKIHealthReport.ps1 -OrganizationName "Contoso"

# Review outputs in C:\Reports\PKI_Assessment and C:\Reports\PKI_Health
```

### Workflow 2: Multi-CA Enterprise Assessment
```powershell
# Step 1: Run assessments on each CA (manual or scripted)
# On CA1:
.\Get-ComprehensivePKIReport.ps1 -OutputDirectory "\\FileServer\PKI\Assessment"
.\Get-PKIHealthReport.ps1 -OutputDirectory "\\FileServer\PKI\Health"

# On CA2:
.\Get-ComprehensivePKIReport.ps1 -OutputDirectory "\\FileServer\PKI\Assessment"
.\Get-PKIHealthReport.ps1 -OutputDirectory "\\FileServer\PKI\Health"

# ... repeat for all CAs

# Step 2: Consolidate from management workstation
.\Merge-PKIAssessmentReports.ps1 `
    -AssessmentPath "\\FileServer\PKI\Assessment" `
    -HealthPath "\\FileServer\PKI\Health" `
    -OutputDirectory "\\FileServer\PKI\Reports" `
    -OrganizationName "Contoso"

# Open consolidated Excel: \\FileServer\PKI\Reports\PKI_Consolidated_*.xlsx
```

### Workflow 3: Weekly Health Monitoring
```powershell
# Create scheduled task on each CA
$action = New-ScheduledTaskAction -Execute 'PowerShell.exe' `
    -Argument '-File "C:\Scripts\Get-PKIHealthReport.ps1" -OutputDirectory "\\FileServer\PKI\Health"'

$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 6am

Register-ScheduledTask -TaskName "PKI Health Check" `
    -Action $action `
    -Trigger $trigger `
    -RunLevel Highest

# Review health scores weekly from consolidated reports
```

### Workflow 4: Migration Assessment
```powershell
# Before migration: Baseline assessment
.\Get-ComprehensivePKIReport.ps1 -OutputDirectory "C:\Migration\Baseline"

# After migration: Validation assessment
.\Get-ComprehensivePKIReport.ps1 -OutputDirectory "C:\Migration\Validation"

# Compare certificate counts, templates, permissions
```

## Output Files

All PKI assessment scripts generate timestamped output files:

### Assessment Script Outputs
**Folder Pattern**: `HOSTNAME_PKI_Assessment_YYYYMMDD_HHMMSS\`
- **Certificates CSV**: `HOSTNAME_PKI_IssuedCertificates_YYYYMMDD_HHMMSS.csv`
- **Templates CSV**: `HOSTNAME_PKI_CertificateTemplates_YYYYMMDD_HHMMSS.csv`
- **Permissions CSV**: `HOSTNAME_PKI_TemplatePermissions_YYYYMMDD_HHMMSS.csv`
- **Text Report**: `HOSTNAME_PKI_Assessment_Report_YYYYMMDD_HHMMSS.txt`

### Health Script Outputs
**Folder Pattern**: `HOSTNAME_PKI_Health_YYYYMMDD_HHMMSS\`
- **CRL Health CSV**: `HOSTNAME_PKI_CRLHealth_YYYYMMDD_HHMMSS.csv`
- **AIA Health CSV**: `HOSTNAME_PKI_AIAHealth_YYYYMMDD_HHMMSS.csv`
- **Template Health CSV**: `HOSTNAME_PKI_TemplateHealth_YYYYMMDD_HHMMSS.csv`
- **Event Log CSV**: `HOSTNAME_PKI_EventLogIssues_YYYYMMDD_HHMMSS.csv`
- **Health Report TXT**: `HOSTNAME_PKI_Health_Report_YYYYMMDD_HHMMSS.txt`

### Merge Script Output
**File Pattern**: `PKI_Consolidated_Assessment_YYYYMMDD_HHMMSS.xlsx`
- **10 Excel Worksheets**: Overview, CA Summary, Certificates, Templates, Permissions, Health Summary, CRL Health, AIA Health, Template Health, Event Log Issues

### Default Output Locations
- **Assessment**: `C:\Reports\PKI_Assessment\`
- **Health**: `C:\Reports\PKI_Health\`
- **Consolidated**: `C:\Reports\PKI_Consolidated\`

## Common Use Cases

### Use Case 1: Certificate Lifecycle Management
```powershell
# Find certificates expiring within 30 days
.\Get-ComprehensivePKIReport.ps1 -DaysToExpiration 30

# Review the CSV: HOSTNAME_PKI_IssuedCertificates_*.csv
# Filter by ExpirationStatus = "Expiring Soon"

# Send notification to certificate owners before expiration
```

### Use Case 2: Template Security Audit
```powershell
# Export all templates with permissions
.\Get-ComprehensivePKIReport.ps1

# Review HOSTNAME_PKI_TemplatePermissions_*.csv
# Look for unexpected Enroll or AutoEnroll permissions
# Verify only authorized groups have access
```

### Use Case 3: Pre-Migration Assessment
```powershell
# Full assessment including revoked certificates
.\Get-ComprehensivePKIReport.ps1 -IncludeRevokedCertificates -OrganizationName "Contoso"

# Review all CSV exports before CA migration
# Document certificate counts, templates, dependencies
```

### Use Case 4: Monthly Health Reporting
```powershell
# Run health checks on all CAs
.\Get-PKIHealthReport.ps1 -OrganizationName "Contoso"

# Review health score trends over time
# Address any scores below 90 (EXCELLENT threshold)
# Track error/warning counts
```

### Use Case 5: Executive PKI Status Report
```powershell
# Step 1: Collect data from all CAs (save to network share)
# Step 2: Consolidate into single Excel file
.\Merge-PKIAssessmentReports.ps1 `
    -AssessmentPath "\\FileServer\PKI\Assessment" `
    -HealthPath "\\FileServer\PKI\Health" `
    -OutputDirectory "\\FileServer\PKI\Reports" `
    -OrganizationName "Contoso"

# Step 3: Send Excel file to stakeholders
# Contains: 193K certificates, 192 templates, health scores for 4 CAs
```

### Use Case 6: Troubleshooting Certificate Issuance
```powershell
# Run health check to identify issues
.\Get-PKIHealthReport.ps1 -EventLogHours 48

# Review Event Log Issues CSV for errors
# Check CRL/AIA health for distribution problems
# Verify template availability in AD
```

## Troubleshooting

### Issue: "ADCS-Administration module not found"
**Solution:**
```powershell
Install-WindowsFeature RSAT-ADCS
Import-Module ADCS-Administration
```

### Issue: "Cannot connect to CA server"
**Solution:**
- Verify the CA server is online: `Test-Connection CAServerName`
- Check firewall rules allow RPC traffic
- Ensure you have admin rights on the CA server
- Try specifying the CA explicitly: `-CAServerName "CA01.domain.com"`

### Issue: "Access Denied when retrieving templates"
**Solution:**
- Run PowerShell as Administrator
- Ensure your account has permissions to read Certificate Templates in AD
- Verify connectivity to Domain Controller: `nltest /dclist:domain.com`

### Issue: "No certificates found"
**Solution:**
- Verify the CA has issued certificates: `certutil -view`
- Check CA service is running: `Get-Service CertSvc`
- Ensure you're querying the correct CA: `certutil -dump` (check Config line)

### Issue: "ImportExcel module not found" (Merge script)
**Solution:**
```powershell
# Module should auto-install, but can install manually
Install-Module ImportExcel -Scope CurrentUser -Force
```

### Issue: "CRL distribution points not accessible"
**Health Impact:** -5 per failed URL  
**Solution:**
- Verify web server hosting CRL files is running
- Check firewall allows HTTP/HTTPS from client networks
- Test URL from different network segment: `Invoke-WebRequest http://...`
- Confirm IIS/Apache virtual directory configuration

### Issue: "Health score dropped below 70"
**Solution:**
1. Review health report recommendations section
2. Check event log issues - address errors first
3. Verify Certificate Services running
4. Test CRL/AIA distribution points
5. Confirm CA certificate not expiring soon
6. Check database size (<10 GB recommended)

## Best Practices

1. **Run Locally**: Always execute assessment/health scripts ON the CA server
2. **Schedule Regular Checks**: Weekly health monitoring, monthly assessments
3. **Archive Reports**: Keep historical data for trend analysis and compliance
4. **Use Network Shares**: For multi-CA environments, centralize output
5. **Monitor Health Scores**: Investigate any drop below 90 (EXCELLENT)
6. **Test Distribution Points**: Ensure CRL/AIA accessible from all client networks
7. **Review Event Logs**: Certificate Services errors indicate operational issues
8. **Plan CA Renewals**: Track CA certificate expiration 90+ days in advance
9. **Audit Template Permissions**: Review who can enroll for each template
10. **Consolidate Reporting**: Use Merge script for executive summaries

## Performance Notes

### Assessment Script
- **10,000 certificates**: ~30 seconds
- **100,000 certificates**: ~5 minutes
- **Large environments**: Consider filtering by date range if needed

### Health Script
- **Typical runtime**: 30-60 seconds
- **Distribution testing**: Adds ~5 seconds per URL
- **Event log analysis**: Scales with log size

### Merge Script
- **193,000 certificates**: ~90 seconds
- **Excel file size**: ~50 MB for 200K certificates
- **Memory usage**: ~500 MB during Excel generation

## Related Documentation

- [Get-ComprehensivePKIReport.ps1 Wiki](../../../docs/wiki/Assessments/PKI/Get-ComprehensivePKIReport.md)
- [Get-PKIHealthReport.ps1 Wiki](../../../docs/wiki/Assessments/PKI/Get-PKIHealthReport.md)
- [Merge-PKIAssessmentReports.ps1 Wiki](../../../docs/wiki/Assessments/PKI/Merge-PKIAssessmentReports.md)
- [Microsoft PKI Documentation](https://docs.microsoft.com/en-us/windows-server/networking/core-network-guide/cng/server-certs/install-the-certification-authority)
- [Certificate Templates](https://docs.microsoft.com/en-us/windows-server/networking/core-network-guide/cng/server-certs/configure-server-certificate-autoenrollment)
- [ADCS PowerShell Cmdlets](https://docs.microsoft.com/en-us/powershell/module/adcsadministration/)

## Contributing

When adding new PKI assessment scripts:
1. Follow the project's PowerShell coding standards (see `.github/copilot-instructions.md`)
2. Include comprehensive comment-based help
3. Use client-agnostic parameters (avoid hardcoded org names)
4. Generate timestamped CSV and text reports
5. Include error handling and color-coded status messages
6. Update this README with the new script information
7. Create wiki documentation in `docs/wiki/Assessments/PKI/`
8. Test with production-scale data before release
