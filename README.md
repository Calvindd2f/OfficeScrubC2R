# OfficeScrubC2R - PowerShell Edition

[![PowerShell Gallery](https://img.shields.io/powershellgallery/v/OfficeScrubC2R)](https://www.powershellgallery.com/packages/OfficeScrubC2R)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)](https://github.com/PowerShell/PowerShell)

Complete PowerShell/C# implementation of Microsoft's **OffScrubC2R.vbs** (v2.19) with **10-50x performance improvements**.

Provides comprehensive removal of Office 2013, 2016, 2019, and Office 365 Click-to-Run installations when standard uninstall methods fail.

## 🚀 Features

- ✅ **Full VBS Compatibility** - All 3,803 lines of VBS functionality ported
- ⚡ **High Performance** - Native C# library for 10-50x faster execution
- 🔒 **Type Safety** - Strongly-typed C# helpers
- 📊 **Comprehensive Logging** - Detailed operation logs
- 🔄 **Parallel Processing** - Multi-threaded operations
- 💾 **Smart Fallback** - Pre-compiled DLL with source compilation fallback
- 🎯 **PowerShell 7+ Support** - Works on both Windows PowerShell and PowerShell Core

## 📋 Requirements

- **OS**: Windows 7 SP1 or later (Windows 10/11 recommended)
- **PowerShell**: 5.1 or later
- **.NET**: Framework 4.5 or later
- **Privileges**: Administrator rights required

## 📦 Installation

### From PowerShell Gallery (Recommended)

```powershell
# Install for current user
Install-Module -Name OfficeScrubC2R -Scope CurrentUser

# Install system-wide (requires admin)
Install-Module -Name OfficeScrubC2R -Scope AllUsers
```

### From GitHub

```powershell
# Clone the repository
git clone https://github.com/Calvindd2f/OfficeScrubC2R.git
cd OfficeScrubC2R

# Import the module
Import-Module .\OfficeScrubC2R.psd1
```

### Manual Installation

1. Download the latest release
2. Extract to a PowerShell module directory:
   - User: `$HOME\Documents\PowerShell\Modules\OfficeScrubC2R`
   - System: `C:\Program Files\PowerShell\Modules\OfficeScrubC2R`
3. Unblock files: `Get-ChildItem -Recurse | Unblock-File`
4. Import: `Import-Module OfficeScrubC2R`

## 🎮 Quick Start

### Basic Usage

```powershell
# Interactive mode (prompts for confirmation)
Invoke-OfficeScrubC2R

# Silent mode (no prompts)
Invoke-OfficeScrubC2R -Quiet -Force

# Detection only (no removal)
Invoke-OfficeScrubC2R -DetectOnly

# Keep Office licenses
Invoke-OfficeScrubC2R -KeepLicense

# Custom log location
Invoke-OfficeScrubC2R -LogPath "C:\Logs"
```

### Advanced Examples

```powershell
# Check what would be removed
Invoke-OfficeScrubC2R -DetectOnly -Verbose

# Silent removal with license preservation
Invoke-OfficeScrubC2R -Quiet -Force -KeepLicense

# Unpin Office from taskbar only
Invoke-OfficeScrubC2R -UnpinMode -SkipSD

# Offline mode (no ODT download)
Invoke-OfficeScrubC2R -Offline -Force
```

### Using Aliases

```powershell
# Short alias
Remove-OfficeC2R -Quiet -Force

# Alternative alias
Uninstall-OfficeC2R -DetectOnly
```

## 📖 Available Functions

| Function | Description |
|----------|-------------|
| `Invoke-OfficeScrubC2R` | Main removal function |
| `Get-InstalledOfficeProducts` | Detect installed Office products |
| `Test-IsC2R` | Check if path/value is C2R-related |
| `Initialize-Environment` | Initialize module environment |
| `Stop-OfficeProcesses` | Stop running Office processes |

## 🔧 Parameters

| Parameter | Description |
|-----------|-------------|
| `-Quiet` | Run in quiet mode with minimal output |
| `-DetectOnly` | Detect products without removing them |
| `-Force` | Skip confirmation prompts |
| `-RemoveAll` | Remove all Office products |
| `-KeepLicense` | Preserve Office licensing information |
| `-Offline` | Skip ODT download |
| `-ForceArpUninstall` | Force ARP-based uninstall |
| `-ClearTaskBand` | Clear taskband shortcuts |
| `-UnpinMode` | Unpin from taskbar/start menu |
| `-SkipSD` | Skip scheduled deletion |
| `-NoElevate` | Don't attempt elevation |
| `-LogPath` | Custom log file location |

## 📊 Performance Comparison

| Operation | VBScript | PowerShell/C# | Improvement |
|-----------|----------|---------------|-------------|
| Registry enumeration | ~10s | ~0.3s | **30x faster** |
| File deletion | ~15s | ~1s | **15x faster** |
| Overall execution | 5-10 min | 30-60s | **10-50x faster** |

*Times measured on Windows 10 with Office 365 installed*

## 🏗️ Architecture

```
OfficeScrubC2R/
├── OfficeScrubC2R.psd1              # Module manifest
├── OfficeScrubC2R.psm1              # Main module
├── OfficeScrubC2R-Utilities.psm1    # Utilities module
├── OfficeScrubC2R-Native.cs         # C# source
├── OfficeScrubNative.dll            # Pre-compiled library
├── build.ps1                        # Build script
├── en-US/
│   └── about_OfficeScrubC2R.help.txt
├── LICENSE
├── README.md
└── CHANGELOG.md
```

### Native C# Components

The `OfficeScrubNative.dll` provides high-performance operations:

- **RegistryHelper**: Win32 registry operations with WOW64 support
- **FileHelper**: Optimized file/folder deletion with reboot scheduling
- **ProcessHelper**: Process termination and monitoring
- **WindowsInstallerHelper**: MSI metadata cleanup
- **TypeLibHelper**: COM type library cleanup
- **LicenseHelper**: Office license and SPP operations
- **ServiceHelper**: Windows service management
- **GuidHelper**: GUID encoding/decoding utilities
- **ShellHelper**: Shell integration (taskbar, start menu)

## 🔨 Building from Source

```powershell
# Compile the native DLL
.\build.ps1

# Clean and rebuild
.\build.ps1 -Clean

# Test the module
Import-Module .\OfficeScrubC2R.psd1 -Force
Invoke-OfficeScrubC2R -DetectOnly
```

## 📝 Error Codes

| Code | Description |
|------|-------------|
| 0 | Success |
| 1 | General failure |
| 2 | Reboot required |
| 4 | User cancelled |
| 8 | MSI uninstall failed |
| 16 | Cleanup failed |
| 32 | Incomplete removal |
| 64 | Second attempt still incomplete |
| 128 | User declined elevation |
| 256 | Elevation failed |
| 512 | Initialization error |
| 1024 | Relaunch error |
| 2048 | Unknown error |

## 🛠️ Troubleshooting

### DLL Load Failure

```powershell
# Unblock the DLL
Unblock-File .\OfficeScrubNative.dll

# Verify .NET version
[System.Environment]::Version
```

### Incomplete Removal

1. Reboot the system
2. Run the tool again: `Invoke-OfficeScrubC2R -Force`
3. Check logs in `$env:TEMP\OfficeScrubC2R\`

### Permission Issues

```powershell
# Check if running as admin
Test-IsElevated

# Re-run PowerShell as Administrator
Start-Process powershell -Verb RunAs
```

## 📜 Logs

Default log location:
```
$env:TEMP\OfficeScrubC2R\<ComputerName>_<Timestamp>_ScrubLog.txt
```

View recent log:
```powershell
Get-Content "$env:TEMP\OfficeScrubC2R\*.txt" | Select-Object -Last 50
```

## 🤝 Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## 📄 License

MIT License - see [LICENSE](LICENSE) file for details.

This software is derived from Microsoft's Office Scrub C2R tool (OffScrubC2R.vbs).  
Original VBScript implementation © Microsoft Corporation.

## ⚠️ Disclaimer

This tool performs deep system changes. Always:
- ✅ Backup important data before use
- ✅ Close all Office applications
- ✅ Run in a test environment first
- ✅ Read the logs if issues occur

## 🔗 Links

- **GitHub**: https://github.com/Calvindd2f/OfficeScrubC2R
- **PowerShell Gallery**: https://www.powershellgallery.com/packages/OfficeScrubC2R
- **Issues**: https://github.com/Calvindd2f/OfficeScrubC2R/issues
- **Changelog**: [CHANGELOG.md](CHANGELOG.md)

## 👤 Author

**Calvin** ([@Calvindd2f](https://github.com/Calvindd2f))

PowerShell/C# port of Microsoft's OffScrubC2R.vbs v2.19

---

⭐ If this tool helped you, please star the repository!
