# OfficeScrubC2R - Project Structure

## Directory Organization

### Root Level Files
- **OfficeScrubC2R.psd1**: PowerShell module manifest with metadata and dependencies
- **OfficeScrubC2R.psm1**: Main PowerShell module containing core functions
- **OfficeScrubC2R-Utilities.psm1**: Nested utilities module with helper functions
- **OfficeScrubC2R-Native.cs**: C# source code for high-performance native operations
- **OfficeScrubNative.dll**: Pre-compiled native library for optimal performance
- **build.ps1**: Build script for compiling the native C# library
- **README.md**: Comprehensive project documentation and usage guide

### Configuration Files
- **PSScriptAnalyzerSettings.psd1**: PowerShell code analysis configuration
- **.gitignore**: Git version control exclusions
- **.gitattributes**: Git file handling attributes
- **LICENSE**: MIT license terms

### Documentation Structure
```
docs/
├── BUILD.md              # Build instructions and requirements
├── CHANGELOG.md          # Version history and changes
├── CHECKLIST.md          # Development and release checklist
├── CONTRIBUTING.md       # Contribution guidelines
├── REPOSITORY_STRUCTURE.md # Detailed project structure
├── SETUP_COMPLETE.md     # Setup completion guide
└── source/               # Original Microsoft source files
    ├── OfficeScrubC2R.vbs # Original VBScript implementation
    └── README            # Original documentation
```

### Development Infrastructure
```
.github/
├── workflows/
│   └── ci.yml           # GitHub Actions CI/CD pipeline
└── scripts/
    ├── utils.ps1        # Build utility functions
    └── Validate-Module.ps1 # Module validation script
```

### Testing and Help
```
tests/
└── Test-OfficeScrubC2R.ps1  # Module test suite

en-US/
└── about_OfficeScrubC2R.help.txt  # PowerShell help documentation

release/                     # Release artifacts directory
```

## Core Components

### PowerShell Module Layer
- **Main Module (OfficeScrubC2R.psm1)**: Primary interface with exported functions
- **Utilities Module**: Helper functions and common operations
- **Module Manifest**: Metadata, dependencies, and export definitions

### Native C# Library
- **RegistryHelper**: Win32 registry operations with WOW64 support
- **FileHelper**: Optimized file/folder deletion with reboot scheduling
- **ProcessHelper**: Process termination and monitoring
- **WindowsInstallerHelper**: MSI metadata cleanup
- **TypeLibHelper**: COM type library cleanup
- **LicenseHelper**: Office license and SPP operations
- **ServiceHelper**: Windows service management
- **GuidHelper**: GUID encoding/decoding utilities
- **ShellHelper**: Shell integration (taskbar, start menu)

## Architectural Patterns

### Hybrid Architecture
- PowerShell provides user interface and workflow orchestration
- C# native library handles performance-critical operations
- Smart fallback from pre-compiled DLL to source compilation

### Modular Design
- Separation of concerns between PowerShell and C# layers
- Nested module structure for utilities and helpers
- Clear interface boundaries between components

### Performance Optimization
- Native code for registry enumeration and file operations
- Parallel processing capabilities for multi-threaded operations
- Efficient GUID manipulation and pattern matching

### Error Handling Strategy
- Comprehensive error codes for different failure scenarios
- Detailed logging with operation tracing
- Graceful degradation with fallback mechanisms

## Component Relationships

### Module Dependencies
```
OfficeScrubC2R.psm1
├── Imports: OfficeScrubC2R-Utilities.psm1
├── Loads: OfficeScrubNative.dll
└── Exports: Core functions and aliases
```

### Build Dependencies
```
build.ps1
├── Compiles: OfficeScrubC2R-Native.cs
├── Outputs: OfficeScrubNative.dll
└── References: System.Management.dll, Microsoft.CSharp.dll
```

### Runtime Flow
1. PowerShell module initialization
2. Native DLL loading with fallback compilation
3. Function export and alias registration
4. User command execution through PowerShell interface
5. Performance-critical operations delegated to C# native code