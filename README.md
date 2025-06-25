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
├── Azure/           # Azure and Microsoft 365 automation scripts
├── Defender/        # Microsoft Defender scripts
├── Graph Commands/  # Microsoft Graph API scripts
├── Intune/          # Intune management scripts
│   └── Assessment/  # Intune assessment scripts
├── Office365/       # Office 365 scripts
build/               # Build and helper scripts
```

## Getting Started
1. **Clone this repository** to your local machine.
2. **Review the scripts** in the `scripts/` directory. Each script includes documentation and parameter help.
3. **Install required PowerShell modules** as noted in each script (e.g., Microsoft.Graph, ExchangeOnlineManagement).
4. **Run scripts** in PowerShell 7+ or Windows PowerShell 5.1, as appropriate.
5. **Customize and extend** scripts as needed for your environment.

## Example Scripts
- `scripts/Azure/Get-CloudOnlyUsers.ps1` – Export all cloud-only users, groups, and distribution groups
- `scripts/Azure/AzCopyCommand.ps1` – Archive files to Azure Blob Storage using AzCopy
- `scripts/Azure/Backup-MgGraphBitLockerKeys.ps1` – Backup BitLocker recovery keys from Microsoft Graph

## Build and Testing
- The `build/` folder contains helper scripts for automation and validation.
- Scripts are validated for public release and include comment-based help for usage.

## License
This repository is provided under the GNU General Public License v3.0. See [LICENSE](LICENSE) for details.

---

> **Tip:** Use and adapt these scripts to accelerate your Microsoft 365 and Azure automation projects. Contributions and improvements are welcome!
