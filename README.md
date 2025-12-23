# Managed Solution PowerShell Everything

This repository contains a collection of PowerShell scripts and helpers for managing Microsoft 365, Azure AD, Intune, and related cloud environments. It is designed to provide practical automation, reporting, and migration tools for IT professionals.

## Features
- Scripts for exporting and managing cloud-only users, groups, and distribution groups
- AzCopy automation for archiving files to Azure Blob Storage
- BitLocker recovery key backup from Microsoft Graph
- Modular folder structure for easy navigation
- Example templates and documentation for extending functionality

## Folder Structure
```
scripts/
├── Assessment/      # Comprehensive environment assessments
│   ├── Lync/       # Lync/Skype for Business assessment tools
│   ├── Microsoft365/  # Microsoft 365 assessment tools
│   ├── Office365/  # Office 365 tenant assessments (legacy location)
│   ├── Security/   # Security posture assessments
│   └── Teams/      # Teams infrastructure assessments
├── Azure/           # Azure and Microsoft 365 automation scripts
├── Defender/        # Microsoft Defender scripts
├── Graph Commands/  # Microsoft Graph API scripts
├── Intune/          # Intune management scripts
│   └── Assessment/  # Intune assessment scripts
├── Office365/       # Office 365 user/mailbox management
├── Data Processing/ # Data analysis and reporting tools
└── Security/        # Security-related scripts and CVE fixes
build/               # Build and helper scripts
docs/                # Project documentation and guides
├── wiki/            # Detailed script documentation
│   └── Assessments/ # Assessment script documentation
│       ├── Lync/    # Lync/Skype documentation
│       └── Microsoft365/ # M365 assessment documentation
└── *.md             # General guides and project docs
```

## Getting Started
1. **Clone this repository** to your local machine.
2. **Review the scripts** in the `scripts/` directory. Each script includes documentation and parameter help.
3. **Install required PowerShell modules** as noted in each script (e.g., Microsoft.Graph, ExchangeOnlineManagement).
4. **Run scripts** in PowerShell 7+ or Windows PowerShell 5.1, as appropriate.
5. **Customize and extend** scripts as needed for your environment.

## Featured Scripts

### Office 365 Assessments (Cloud Shell Ready) 🚀
- `scripts/Assessment/Office365/Get-QuickO365Report.ps1` – Fast Office 365 tenant assessment with automatic ZIP download
- `scripts/Assessment/Office365/Get-ComprehensiveO365Report.ps1` – Advanced assessment with archives, rules, and full analytics
- See [Office 365 Assessment Guide](docs/Office365-Assessment-Guide.md) for complete documentation

### Azure & Graph API
- `scripts/Azure/Get-CloudOnlyUsers.ps1` – Export all cloud-only users, groups, and distribution groups
- `scripts/Azure/AzCopyCommand.ps1` – Archive files to Azure Blob Storage using AzCopy
- `scripts/Azure/Backup-MgGraphBitLockerKeys.ps1` – Backup BitLocker recovery keys from Microsoft Graph

### Lync/Skype for Business
- `scripts/Assessment/Lync/Start-LyncCsvExporter.ps1` – Interactive menu-based Lync data exporter ([docs](docs/wiki/Assessments/Lync/Start-LyncCsvExporter.md))
- `scripts/Assessment/Lync/Get-ComprehensiveLyncReport.ps1` – Complete Lync environment assessment ([docs](docs/wiki/Assessments/Lync/Get-ComprehensiveLyncReport.md))
- `scripts/Assessment/Lync/Get-LyncHealthReport.ps1` – Health monitoring and diagnostics ([docs](docs/wiki/Assessments/Lync/Get-LyncHealthReport.md))
- `scripts/Assessment/Lync/Export-ADLyncTeamsMigrationData.ps1` – AD export for Teams migration ([docs](docs/wiki/Assessments/Lync/Export-ADLyncTeamsMigrationData.md))
- See [Lync Assessment Scripts Overview](docs/wiki/Assessments/Lync/README.md) for complete documentation

### Microsoft Teams
- `scripts/Assessment/Teams/Get-ComprehensiveTeamsReport.ps1` – Full Teams infrastructure assessment

## Documentation

### Wiki Documentation
Detailed documentation for scripts is available in the `docs/wiki/` directory:
- **[Lync Assessment Scripts](docs/wiki/Assessments/Lync/README.md)** - Complete Lync/Skype for Business assessment suite
- **[Microsoft 365 Assessment Scripts](docs/wiki/Assessments/Microsoft365/)** - M365 tenant assessment tools
- **[Office 365 Quick Start Guide](docs/Office365-Quick-Start.md)** - Getting started with O365 assessments

### Script Documentation
Each script includes:
- Comprehensive comment-based help (`.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, `.EXAMPLE`, `.NOTES`)
- Usage examples and parameter descriptions
- Prerequisites and required modules
- Output format and file naming conventions

For detailed script documentation, see the wiki articles linked above or use PowerShell's built-in help:
```powershell
Get-Help .\ScriptName.ps1 -Full
```

## Build and Testing
- The `build/` folder contains helper scripts for automation and validation.
- Scripts are validated for public release and include comment-based help for usage.

## License
This repository is provided under the GNU General Public License v3.0. See [LICENSE](LICENSE) for details.

---

> **Tip:** Use and adapt these scripts to accelerate your Microsoft 365 and Azure automation projects. Contributions and improvements are welcome!
