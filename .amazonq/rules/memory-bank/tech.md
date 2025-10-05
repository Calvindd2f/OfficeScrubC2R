# OfficeScrubC2R - Technology Stack

## Programming Languages

### PowerShell
- **Version**: 5.1+ (Windows PowerShell) and 7+ (PowerShell Core)
- **Usage**: Primary interface, workflow orchestration, and user interaction
- **Compatibility**: Cross-edition support (Desktop and Core)

### C#
- **Version**: .NET Framework 4.5+
- **Usage**: High-performance native operations and Win32 API interactions
- **Compilation**: Pre-compiled DLL with source fallback using csc.exe

## Framework Dependencies

### .NET Framework
- **Minimum**: 4.5
- **Target**: Framework 4.0.30319 for maximum compatibility
- **Assemblies**: System.Management.dll, Microsoft.CSharp.dll

### PowerShell Modules
- **Core Module**: OfficeScrubC2R.psm1
- **Nested Module**: OfficeScrubC2R-Utilities.psm1
- **Manifest**: OfficeScrubC2R.psd1

## Build System

### Compilation Process
```powershell
# Build the native DLL
.\build.ps1

# Clean and rebuild
.\build.ps1 -Clean
```

### Build Configuration
- **Compiler**: C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
- **Target**: Library (DLL)
- **Optimization**: Enabled (/optimize+)
- **Warning Level**: 4 (/warn:4)

### Build Dependencies
- System.Management.dll (WMI operations)
- Microsoft.CSharp.dll (dynamic compilation)
- Windows SDK (Win32 API declarations)

## Development Commands

### Module Operations
```powershell
# Import the module
Import-Module .\OfficeScrubC2R.psd1

# Test module functionality
Import-Module .\OfficeScrubC2R.psd1 -Force
Invoke-OfficeScrubC2R -DetectOnly

# Run tests
.\tests\Test-OfficeScrubC2R.ps1
```

### Code Analysis
```powershell
# PowerShell Script Analyzer
Invoke-ScriptAnalyzer -Path . -Settings .\PSScriptAnalyzerSettings.psd1

# Module validation
.\.github\scripts\Validate-Module.ps1
```

## Native API Integration

### Win32 APIs
- **Registry**: advapi32.dll (RegOpenKeyEx, RegDeleteKeyEx, RegEnumKeyEx)
- **File System**: kernel32.dll (MoveFileEx, SetFileAttributes, RemoveDirectory)
- **Shell**: shell32.dll (ILCreateFromPath, SHCreateShellItem)

### COM Interfaces
- **IShellItem**: Shell item manipulation
- **Windows Installer**: MSI operations
- **Type Libraries**: COM registration cleanup

## Performance Optimizations

### Native Code Benefits
- **Registry Operations**: 30x faster enumeration
- **File Operations**: 15x faster deletion
- **Overall Execution**: 10-50x performance improvement

### Parallel Processing
- Multi-threaded registry enumeration
- Concurrent file system operations
- Asynchronous process management

## Development Tools

### Code Quality
- **PSScriptAnalyzer**: PowerShell code analysis
- **Custom Rules**: Project-specific analysis settings
- **GitHub Actions**: Automated CI/CD pipeline

### Version Control
- **Git**: Source control with .gitignore and .gitattributes
- **GitHub**: Repository hosting and issue tracking
- **Semantic Versioning**: Version 2.19.0 following semver

### Documentation
- **PowerShell Help**: about_OfficeScrubC2R.help.txt
- **Markdown**: README.md, CHANGELOG.md, BUILD.md
- **Inline Comments**: Comprehensive code documentation

## Deployment

### PowerShell Gallery
- **Package Name**: OfficeScrubC2R
- **Installation**: Install-Module -Name OfficeScrubC2R
- **Scopes**: CurrentUser and AllUsers

### Manual Installation
- **User Path**: $HOME\Documents\PowerShell\Modules\OfficeScrubC2R
- **System Path**: C:\Program Files\PowerShell\Modules\OfficeScrubC2R
- **File Unblocking**: Get-ChildItem -Recurse | Unblock-File

## System Requirements

### Operating System
- Windows 7 SP1 or later
- Windows 10/11 recommended
- Both x86 and x64 architectures supported

### Runtime Requirements
- Administrator privileges required
- .NET Framework 4.5 or later
- PowerShell execution policy allowing module import