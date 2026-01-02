# Office 365 Cloud Shell Quick Start

**Get your Office 365 assessment done in 3 steps!**

## ⚡ Quick Start (5 minutes)

### Step 1: Open Cloud Shell
1. Go to [https://shell.azure.com](https://shell.azure.com)
2. Select **PowerShell** mode
3. Wait for shell to initialize

### Step 2: Upload & Run
```powershell
# If you have the script locally, use Upload button
# Or clone the repo:
git clone <your-repo-url>
cd PowerShellEveryting/scripts/Assessment/Office365

# Run the quick assessment
.\Get-QuickO365Report.ps1
```

### Step 3: Download Results
```powershell
# After script completes, download the ZIP file
download cloudshell:\O365Report_*.zip
```

**Done!** Open the ZIP file and review your CSV reports.

---

## 📊 What You'll Get

### Files in ZIP Archive:
- **Mailboxes.csv** - All mailbox sizes, quotas, and usage
- **OneDrive.csv** - Personal OneDrive storage for all users  
- **SharePoint.csv** - Team site storage and ownership
- **Summary.txt** - Executive summary with totals

### Data Included:
✅ Mailbox sizes in GB  
✅ Storage quotas  
✅ Last access times  
✅ OneDrive usage per user  
✅ SharePoint site storage  
✅ Total tenant storage  

---

## 🎯 Common Use Cases

### Capacity Planning
**Goal**: Understand current storage usage and forecast needs

**Use**: Quick script (default)
```powershell
.\Get-QuickO365Report.ps1
```
**Review**: Summary.txt for totals, Mailboxes.csv for growth trends

### Migration Assessment  
**Goal**: Inventory all mailboxes and sites before migration

**Use**: Comprehensive script with archives
```powershell
.\Get-ComprehensiveO365Report.ps1 -IncludeArchives
```
**Review**: All CSV files for complete inventory

### Security Audit
**Goal**: Find inbox rules that forward email externally

**Use**: Comprehensive script with rules (WARNING: slow for large tenants)
```powershell
.\Get-ComprehensiveO365Report.ps1 -IncludeMailboxRules
```
**Review**: MailboxRules.csv, filter for ForwardTo/RedirectTo

### Storage Cleanup
**Goal**: Identify largest mailboxes and sites

**Use**: Quick script (default)
```powershell
.\Get-QuickO365Report.ps1
```
**Review**: Sort Mailboxes.csv by SizeGB descending

---

## ⏱️ Time Estimates

| Tenant Size | Quick Script | With Archives | With Rules |
|-------------|-------------|---------------|------------|
| Small (<100 mbx) | 2-5 min | 5-10 min | 15-30 min |
| Medium (100-500) | 5-10 min | 15-30 min | 1-2 hours |
| Large (500-2000) | 10-20 min | 30-60 min | 2-5 hours |
| XL (2000+) | 20-40 min | 1-2 hours | 5+ hours |

---

## 🔧 Troubleshooting

### Script won't start
**Error**: "Cannot find path"
```powershell
# Verify you're in the right directory
pwd
# Should show: .../Assessment/Office365

# If not, navigate there
cd PowerShellEveryting/scripts/Assessment/Office365
```

### SharePoint connection fails
**Error**: "Failed to connect to SharePoint Online"
```powershell
# Find your tenant name
Get-OrganizationConfig | Select-Object Name

# Run with explicit tenant (replace "contoso")
.\Get-QuickO365Report.ps1 -TenantDomain "contoso"
```

### Permission denied
**Error**: "Access denied" or "Unauthorized"

**Fix**: You need one of these roles:
- Global Administrator
- Exchange Administrator + SharePoint Administrator  
- Global Reader

Contact your IT admin to assign appropriate role.

### Cloud Shell timeout
**Error**: Shell becomes unresponsive after 20 minutes

**Fix**: Cloud Shell has 20-minute idle timeout. Keep browser tab active during script execution.

---

## 💡 Pro Tips

### Fastest Assessment
```powershell
# Just mailboxes and sites, no extras
.\Get-QuickO365Report.ps1
```

### Most Complete Assessment  
```powershell
# Everything except rules (rules are VERY slow)
.\Get-ComprehensiveO365Report.ps1 -IncludeArchives
```

### Storage Analysis Only
```powershell
# Quick script is perfect for storage analysis
.\Get-QuickO365Report.ps1
```

### Run During Off-Peak Hours
- API calls are faster during off-peak times
- For large tenants (1000+ mailboxes), run overnight or on weekends

### Keep Browser Active
- Cloud Shell requires active browser tab
- Don't minimize or switch tabs for long periods
- Script will resume if reconnected quickly

---

## 📖 Need More Help?

### Full Documentation
- [Office 365 Assessment Guide](../../docs/Office365-Assessment-Guide.md) - Complete documentation
- [Assessment Folder README](README.md) - Script comparison and details

### Script Selection Guide

**Use Quick Script if**:
- ✅ You want fast results
- ✅ Storage analysis is your primary goal
- ✅ You have <500 mailboxes
- ✅ You don't need archive or rules data

**Use Comprehensive Script if**:
- ✅ You need archive mailbox data
- ✅ You need inbox rules audit
- ✅ You want filtering options
- ✅ Migration planning requires complete data
- ✅ Compliance audit requires forwarding rules

---

## 🔒 Security Notes

### Scripts are Read-Only
- ✅ No mailbox modifications
- ✅ No data deletion  
- ✅ No permission changes
- ✅ Safe to run in production

### Data Protection
After downloading reports:
1. Delete ZIP from Cloud Shell: `rm cloudshell:\O365Report_*.zip`
2. Store reports securely on encrypted drive
3. Limit access to IT administrators
4. Follow your organization's data retention policies

---

## 🎓 Training Your Team

### For Help Desk
Quick script provides:
- Mailbox sizes for quota questions
- OneDrive usage for storage requests
- Quick capacity reports

### For IT Managers  
Summary.txt provides:
- Total tenant storage
- Average mailbox sizes
- Largest consumers
- Storage trends

### For Auditors
Comprehensive script with rules provides:
- Complete mailbox inventory
- External forwarding rules
- Storage compliance data
- Audit trails

---

## ✅ Success Checklist

Before running:
- [ ] Azure Cloud Shell is open in PowerShell mode
- [ ] You have required permissions (Global Reader minimum)
- [ ] You've uploaded or cloned the script
- [ ] You understand expected execution time

After running:
- [ ] ZIP file downloaded successfully
- [ ] All expected CSV files are in ZIP
- [ ] Summary.txt shows expected counts
- [ ] Reports deleted from Cloud Shell
- [ ] Reports stored securely

---

**Questions?** See [README.md](README.md) or [Office365-Assessment-Guide.md](../../docs/Office365-Assessment-Guide.md)

**Ready to run?** Execute `.\Get-QuickO365Report.ps1` now!
