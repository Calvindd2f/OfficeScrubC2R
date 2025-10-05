# Building OfficeScrubC2R Native Components

## Overview

The OfficeScrubC2R utilities use a native C# component (`OfficeScrubNative.dll`) for performance-critical operations like registry manipulation, file operations, and process management.

## Building the DLL

### Quick Build

Run the build script:

```powershell
.\build.ps1
```

To clean and rebuild:

```powershell
.\build.ps1 -Clean
```

### Manual Build

If you need to compile manually:

```cmd
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe ^
  /target:library ^
  /out:OfficeScrubNative.dll ^
  /reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Management.dll" ^
  /reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\Microsoft.CSharp.dll" ^
  OfficeScrubC2R-Native.cs
```

## How It Works

The PowerShell module (`OfficeScrubC2R-Utilities.psm1`) uses the following loading strategy:

1. **First**: Attempts to load the pre-compiled `OfficeScrubNative.dll` (fastest)
2. **Fallback**: If DLL not found, compiles `OfficeScrubC2R-Native.cs` inline (slower startup)

This approach provides:

- **Performance**: No compilation delay when DLL is present
- **Portability**: Can still run with just the .cs source file
- **Flexibility**: Easy development and testing

## Distribution

For distribution, you can include either:

- **Recommended**: Both `OfficeScrubNative.dll` and `OfficeScrubC2R-Native.cs`
- **Source-only**: Just `OfficeScrubC2R-Native.cs` (will compile on first use)
- **Binary-only**: Just `OfficeScrubNative.dll` (fastest, but no fallback)

## Requirements

- .NET Framework 4.0 or later
- Windows PowerShell 5.1 or PowerShell 7+
- Administrator privileges (for Office removal operations)

## Architecture

The native DLL contains:

- **RegistryHelper**: High-performance registry operations with WOW64 support
- **FileHelper**: Robust file/folder deletion with reboot scheduling
- **ProcessHelper**: Process termination and monitoring
- **WindowsInstallerHelper**: MSI metadata cleanup
- **TypeLibHelper**: COM type library cleanup
- **LicenseHelper**: Office license and SPP cleanup
- **ServiceHelper**: Windows service management
- **GuidHelper**: GUID encoding/decoding utilities
- **ShellHelper**: Shell integration (taskbar pins, start menu, etc.)

## Troubleshooting

### DLL Load Failure

If you see "Failed to load DLL", check:

1. DLL is in the same directory as the .psm1 file
2. DLL is not blocked (run `Unblock-File OfficeScrubNative.dll`)
3. .NET Framework 4.0+ is installed

### Compilation Errors

The fallback inline compilation requires:

- System.Management.dll
- Microsoft.CSharp.dll

These are included with Windows by default.
