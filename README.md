# PowerShell Module Template

This repository provides a standardized template for building robust, maintainable PowerShell modules.

## Features
- Standardized folder structure for source, tests, resources, and documentation
- Example templates for classes, enums, public/private functions, and integration tests
- Automated build and versioning scripts (see `build/build.ps1`)
- Documentation and testing helpers

## Folder Structure
```
module-template/
├── build/                # Build scripts and helpers (keep `build.ps1` for automation)
├── docs/                 # Documentation (markdown)
├── resources/            # Data and templates for your module
├── src/                  # Source code (classes, enums, functions)
├── tests/                # Unit and integration tests (keep all unit tests for quality)
├── Module.psd1           # Module manifest
├── Module.psm1           # Main module file
├── README.md             # This file
└── LICENSE               # License file
```

## Getting Started
1. **Clone this repository** and rename the root folder to your module name.
2. **Update `Module.psd1`** with your module's metadata and exported functions.
3. **Add your code** to the `src/` directory, using the provided templates as a starting point.
4. **Write tests** in the `tests/` directory (keep and expand unit tests for reliability).
5. **Document your functions** in the `docs/` folder using markdown.
6. **Use the build scripts** in the `build/` folder to automate versioning and validation.

## Templates Included
- `src/classes/class-template.ps1` – PowerShell class template
- `src/enum/enum-template.ps1` – Enum template
- `src/public/public-function-template.ps1` – Public function template
- `src/private/private-function-template.ps1` – Private function template
- `resources/templates/temp-template.ps1` – Script template resource
- `tests/integration/temp-integration-test.ps1` – Integration test template

## Build and Testing
- The build process is managed by `build/build.ps1`. Always keep this file to automate versioning, fingerprinting, and validation.
- Unit tests in `tests/unit/` are essential for maintaining code quality. Keep and expand these tests as your module grows.

## License
This template is provided under the GNU General Public License v3.0. See [LICENSE](LICENSE) for details.

---

> **Tip:** Use this template as a starting point for all new PowerShell modules to ensure consistency, maintainability, and best practices.
