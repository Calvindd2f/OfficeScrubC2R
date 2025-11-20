# OfficeScrubC2R - Native Wrapper

This repository now focuses on a **lightweight PowerShell wrapper** around the converted `OfficeScrubNative.dll`, which implements the original OffScrubC2R logic with direct native methods and Windows API calls. The wrapper keeps the original source materials alongside the module so you can clearly see the lineage from VBScript to C#.

## What's in the repository

- `OfficeScrubC2R.psd1` / `OfficeScrubC2R.psm1` – the minimal PowerShell module that loads the native DLL and exposes a handful of helper commands.
- `OfficeScrubNative.dll` – the latest converted binary that uses native methods/Windows API calls.
- Reference sources kept for contrast:
  - `docs/source/OfficeScrubC2R.vbs` – the original Microsoft OffScrubC2R VBScript.
  - `direct_conversion.cs` – a direct C# conversion of the VBScript logic.
  - `OfficeScrubC2R-Native.cs` – the native-method focused C# implementation that drives the DLL.
- Supporting files: `LICENSE`, `CHANGELOG.md`, and build metadata.

## Quick start

```powershell
# Import the wrapper module from the repository root
Import-Module .\OfficeScrubC2R.psd1 -Force

# Load the native DLL and create an orchestrator (defaults to your OS bitness)
$orchestrator = New-OfficeScrubOrchestrator

# Check whether a path looks like a Click-to-Run installation
Test-OfficeC2RPath -Path "C:\\Program Files\\Microsoft Office"

# Check whether a product code is in scope for scrubbing
Test-OfficeProductScope -ProductCode '{9AC08E99-230B-47E8-9721-4577B7F124EA}'

# Surface the helper classes from the native DLL for direct use
$helpers = Get-OfficeScrubHelpers -Orchestrator $orchestrator
$helpers.RegistryHelper
```

The module only concerns itself with loading the native binary and handing you the orchestrator/helpers. Any higher-level orchestration can be built on top using these primitives.

## Helper commands

| Command | Purpose |
| --- | --- |
| `Import-OfficeScrubNative` | Loads `OfficeScrubNative.dll` from the module folder (or a provided path). |
| `New-OfficeScrubOrchestrator` | Creates the orchestrator that wires up all helper classes with the correct bitness. |
| `Test-OfficeC2RPath` | Uses the native helper to determine whether a path matches Click-to-Run patterns. |
| `Test-OfficeProductScope` | Tests whether a product code is in scope for scrubbing. |
| `Get-OfficeScrubHelpers` | Returns the helper class instances (registry, files, processes, MSI, etc.). |

## Building the native DLL

The prebuilt `OfficeScrubNative.dll` is included. If you need to rebuild it on Windows, use `build.ps1` and the `OfficeScrubC2R-Native.cs` source as the input.

### Verifying the native build on both PowerShell editions

Use `tests/Verify-NativeBuild.ps1` to confirm the native source compiles cleanly for **.NET Framework (Windows PowerShell)** and **.NET SDK (PowerShell 7+)** and that the resulting assembly can be loaded:

```powershell
# From Windows PowerShell 5.1 (.NET Framework)
powershell.exe -NoLogo -NoProfile -File .\tests\Verify-NativeBuild.ps1 -Verbose

# From PowerShell 7+ (uses the .NET SDK, target defaults to net7.0-windows)
pwsh -NoLogo -NoProfile -File .\tests\Verify-NativeBuild.ps1 -Verbose

# Clean build outputs from prior runs
powershell.exe -NoLogo -NoProfile -File .\tests\Verify-NativeBuild.ps1 -Clean
```

The script emits warnings when the relevant compiler toolchain is not available and exercises `OfficeScrubNative.OfficeScrubOrchestrator` from the produced assembly in the matching PowerShell edition.

## Licensing

MIT License. See `LICENSE` for details.
