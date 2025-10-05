# Repository Structure

This document describes the complete structure of the OfficeScrubC2R repository.

## Directory Layout

```
OfficeScrubC2R/
│
├── .github/
│   └── workflows/
│       └── ci.yml                           # GitHub Actions CI/CD workflow
│
├── docs/
│   ├── BUILD.md                             # Build instructions for native DLL
│   └── source/                              # Original VBS source reference
│       ├── OfficeScrubC2R.vbs
│       └── README
│
├── en-US/
│   └── about_OfficeScrubC2R.help.txt       # PowerShell help documentation
│
├── tests/
│   └── Test-OfficeScrubC2R.ps1             # Test suite
│
├── .gitignore                               # Git ignore patterns
├── build.ps1                                # Build script for C# DLL
├── CHANGELOG.md                             # Version history and changes
├── CONTRIBUTING.md                          # Contribution guidelines
├── LICENSE                                  # MIT License
├── OfficeScrubC2R-Native.cs                # C# performance library source
├── OfficeScrubC2R-Utilities.ps1            # Legacy utilities (kept for reference)
├── OfficeScrubC2R-Utilities.psm1           # PowerShell utilities module
├── OfficeScrubC2R.ps1                      # Main script (can be run standalone)
├── OfficeScrubC2R.psd1                     # Module manifest
├── OfficeScrubC2R.psm1                     # Main module file
├── OfficeScrubNative.dll                   # Pre-compiled C# library
├── PSScriptAnalyzerSettings.psd1           # Code analysis settings
├── PUBLISH.md                               # Publishing guide for PSGallery
├── README.md                                # Main documentation
└── REPOSITORY_STRUCTURE.md                 # This file
```

## Core Files

### Module Files

| File                            | Purpose                                                                   |
| ------------------------------- | ------------------------------------------------------------------------- |
| `OfficeScrubC2R.psd1`           | Module manifest with metadata for PowerShell Gallery                      |
| `OfficeScrubC2R.psm1`           | Main module that exports functions and imports utilities                  |
| `OfficeScrubC2R-Utilities.psm1` | Utilities module with helper functions                                    |
| `OfficeScrubC2R.ps1`            | Main script with all removal logic (can run standalone or be dot-sourced) |

### Native Components

| File                       | Purpose                                        |
| -------------------------- | ---------------------------------------------- |
| `OfficeScrubC2R-Native.cs` | C# source code for high-performance operations |
| `OfficeScrubNative.dll`    | Pre-compiled native library (faster loading)   |
| `build.ps1`                | Script to compile the DLL from source          |

### Documentation

| File                                  | Purpose                                           |
| ------------------------------------- | ------------------------------------------------- |
| `README.md`                           | Main documentation with usage examples            |
| `CHANGELOG.md`                        | Version history and release notes                 |
| `CONTRIBUTING.md`                     | Guidelines for contributors                       |
| `PUBLISH.md`                          | Instructions for publishing to PowerShell Gallery |
| `docs/BUILD.md`                       | Detailed build instructions                       |
| `en-US/about_OfficeScrubC2R.help.txt` | PowerShell help file                              |
| `REPOSITORY_STRUCTURE.md`             | This file                                         |

### Configuration

| File                            | Purpose                               |
| ------------------------------- | ------------------------------------- |
| `.gitignore`                    | Files and folders to exclude from git |
| `PSScriptAnalyzerSettings.psd1` | Code quality rules                    |
| `.github/workflows/ci.yml`      | Continuous integration configuration  |
| `LICENSE`                       | MIT License with attribution          |

## Module Architecture

### Loading Sequence

1. **User runs**: `Import-Module OfficeScrubC2R`
2. **PowerShell loads**: `OfficeScrubC2R.psm1` (root module)
3. **Root module imports**: `OfficeScrubC2R-Utilities.psm1` (nested module)
4. **Utilities module loads**: `OfficeScrubNative.dll` (native assembly)
5. **Root module dot-sources**: `OfficeScrubC2R.ps1` (main script functions)
6. **Module exports**: Public functions to user session

### Function Organization

#### Public Functions (Exported)

- `Invoke-OfficeScrubC2R` - Main removal function
- `Get-InstalledOfficeProducts` - Detection function
- `Test-IsC2R` - Validation function
- `Initialize-Environment` - Setup function
- `Stop-OfficeProcesses` - Process management

#### Aliases

- `Remove-OfficeC2R` → `Invoke-OfficeScrubC2R`
- `Uninstall-OfficeC2R` → `Invoke-OfficeScrubC2R`

#### Internal Functions (Not Exported)

All other functions in utilities and main script are internal.

### Native C# Components

The `OfficeScrubNative.dll` contains:

```
OfficeScrubNative
├── RegistryHiveType (Enum)
├── OfficeConstants (Static Class)
├── GuidHelper (Static Class)
├── RegistryHelper (Class)
├── FileHelper (Class)
├── ProcessHelper (Class)
├── ShellHelper (Class)
├── WindowsInstallerHelper (Class)
├── TypeLibHelper (Class)
├── LicenseHelper (Class)
├── ServiceHelper (Class)
└── OfficeScrubOrchestrator (Class)
```

## PowerShell Gallery Requirements

### Required Files

- ✅ Module manifest (`.psd1`)
- ✅ Module file (`.psm1`)
- ✅ License file (`LICENSE`)
- ✅ README with examples
- ✅ Native assembly (`.dll`)

### Manifest Metadata

- ✅ Unique GUID
- ✅ Semantic version (2.19.0)
- ✅ Author information
- ✅ Description
- ✅ Tags for discovery
- ✅ Project URI
- ✅ License URI
- ✅ Release notes
- ✅ Function exports
- ✅ Required assemblies
- ✅ PowerShell version (5.1+)
- ✅ Compatible editions (Desktop, Core)

### Quality Checks

- ✅ PSScriptAnalyzer compliant
- ✅ Help documentation
- ✅ Examples in README
- ✅ Error handling
- ✅ Logging
- ✅ Parameter validation

## GitHub Repository Setup

### Branch Strategy

- `main` - Stable releases
- `develop` - Development branch
- `feature/*` - Feature branches
- `hotfix/*` - Urgent fixes

### GitHub Actions

- CI on push/PR
- Module validation
- PSScriptAnalyzer checks
- Build verification
- Function tests

### Releases

1. Tag: `v2.19.0`
2. GitHub Release with notes
3. Attached assets (DLL)
4. Published to PowerShell Gallery

## Installation Methods

### From PowerShell Gallery

```powershell
Install-Module -Name OfficeScrubC2R
```

### From GitHub

```powershell
git clone https://github.com/Calvindd2f/OfficeScrubC2R.git
Import-Module .\OfficeScrubC2R\OfficeScrubC2R.psd1
```

### Manual

1. Download release
2. Extract to module path
3. Unblock files
4. Import module

## Development Workflow

### Setup

```powershell
git clone https://github.com/Calvindd2f/OfficeScrubC2R.git
cd OfficeScrubC2R
.\build.ps1
Import-Module .\OfficeScrubC2R.psd1 -Force
```

### Making Changes

1. Create feature branch
2. Make changes
3. Run PSScriptAnalyzer
4. Test thoroughly
5. Update CHANGELOG.md
6. Create pull request

### Building

```powershell
.\build.ps1           # Build DLL
.\build.ps1 -Clean    # Clean and rebuild
```

### Testing

```powershell
# Import module
Import-Module .\OfficeScrubC2R.psd1 -Force -Verbose

# Run detection
Invoke-OfficeScrubC2R -DetectOnly

# Check for issues
Invoke-ScriptAnalyzer -Path . -Recurse
```

## File Size Considerations

| File Type        | Size Range | Notes            |
| ---------------- | ---------- | ---------------- |
| `.ps1` scripts   | 30-50 KB   | Main logic files |
| `.psm1` modules  | 20-40 KB   | Module files     |
| `.psd1` manifest | 5-10 KB    | Metadata         |
| `.dll` assembly  | 35-40 KB   | Compiled C#      |
| `.cs` source     | 40-50 KB   | C# source        |
| `.md` docs       | 5-20 KB    | Documentation    |

**Total module size**: ~200-250 KB

## Version History

- **2.19.0**: Initial PowerShell Gallery release
- **0.1.0**: Development version

## Support

- **Issues**: https://github.com/Calvindd2f/OfficeScrubC2R/issues
- **Discussions**: GitHub Discussions
- **Email**: calvindd2f@gmail.com

## License

MIT License - See LICENSE file for details.

Derived from Microsoft's OffScrubC2R.vbs © Microsoft Corporation.
