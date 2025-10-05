# OfficeScrubC2R - Product Overview

## Project Purpose
Complete PowerShell/C# implementation of Microsoft's OffScrubC2R.vbs (v2.19) that provides comprehensive removal of Office 2013, 2016, 2019, and Office 365 Click-to-Run installations when standard uninstall methods fail.

## Value Proposition
- **10-50x Performance Improvement**: Native C# library delivers dramatically faster execution compared to original VBScript
- **Full VBS Compatibility**: All 3,803 lines of VBScript functionality ported with complete feature parity
- **Modern PowerShell Support**: Works on both Windows PowerShell 5.1+ and PowerShell Core 7+
- **Enterprise Ready**: Comprehensive logging, error handling, and parallel processing capabilities

## Key Features

### Core Capabilities
- Complete Office Click-to-Run removal when standard uninstall fails
- Registry cleanup with WOW64 support for both 32-bit and 64-bit entries
- File system cleanup with reboot scheduling for locked files
- Process termination and service management
- MSI metadata cleanup and COM type library cleanup
- Office license and SPP (Software Protection Platform) operations
- Shell integration cleanup (taskbar, start menu)

### Performance Enhancements
- Native C# library for high-performance registry and file operations
- Parallel processing for multi-threaded operations
- Smart fallback system with pre-compiled DLL and source compilation
- Optimized GUID encoding/decoding utilities

### Operational Features
- Interactive and silent execution modes
- Detection-only mode for assessment without removal
- License preservation options
- Comprehensive logging with detailed operation logs
- Offline mode support (no ODT download required)
- Custom log location support

## Target Users

### Primary Users
- **IT Administrators**: Managing Office deployments in enterprise environments
- **System Administrators**: Troubleshooting failed Office installations
- **Technical Support**: Resolving Office installation conflicts

### Use Cases
- Failed Office uninstallations requiring deep cleanup
- Office version migration preparation
- Corrupted Office installation recovery
- Enterprise deployment troubleshooting
- System preparation for fresh Office installations

## System Requirements
- **OS**: Windows 7 SP1 or later (Windows 10/11 recommended)
- **PowerShell**: 5.1 or later
- **.NET**: Framework 4.5 or later
- **Privileges**: Administrator rights required
- **Architecture**: Supports both x86 and x64 systems

## Distribution
- PowerShell Gallery package for easy installation
- GitHub repository with source code and documentation
- Pre-compiled native DLL with automatic fallback compilation
- Comprehensive help documentation and examples