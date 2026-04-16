# OfficeScrubC2R

OfficeScrubC2R is a hardened binary PowerShell module for inspecting and scrubbing Microsoft Office Click-to-Run state. Version 3.0.0 now includes guarded destructive execution for a first VBS-inspired cleanup slice: Office process termination, Office scheduled task deletion, Click-to-Run service stop/delete, C2R registry cleanup, known/detected C2R folder deletion with reboot scheduling fallback, and optional current-user license cache preservation.

This repository includes Microsoft `OffScrubC2R.vbs` and earlier conversion sources as reference material. The shipped module is built from the SDK-style projects under `src/`.

## Supported Commands

| Command | Purpose | Destructive |
| --- | --- | --- |
| `Get-InstalledOfficeProducts` | Reads Office C2R configuration and uninstall registry evidence across explicit 32-bit and 64-bit registry views. | No |
| `Test-OfficeC2RState` | Reports elevation/SYSTEM status, C2R products, package paths, services, running Office processes, pending reboot-delete evidence, and preflight issues. | No |
| `Invoke-OfficeScrubC2R -PlanOnly` | Returns a structured plan describing cleanup actions. | No |
| `Invoke-OfficeScrubC2R -WhatIf` | Exercises PowerShell `ShouldProcess` behavior and returns the same planning-only operation list. | No |
| `Invoke-OfficeScrubC2R` | Runs the guarded cleanup executor from an elevated session and returns structured operation results. | Yes |

`Invoke-OfficeScrubC2R` uses the stable error id `OfficeScrubC2R.AdminRequired` when destructive execution is attempted from a non-elevated session.

## Requirements

- Windows 10 or Windows 11
- Windows PowerShell 5.1 or PowerShell 7+
- .NET Framework 4.7.2 or later for Windows PowerShell hosts
- .NET SDK 8.0+ to build and test from source

Windows 7 SP1 and .NET 4.5 are no longer advertised as the production support floor for the v3 binary baseline.

## Build

```powershell
.\build.ps1 -Clean
```

The build script compiles the binary cmdlets and writes package-ready output to:

```text
artifacts/module/
```

It also writes SHA256 checksums to:

```text
artifacts/checksums.sha256
```

Compiled DLLs and PDBs are build or release artifacts only. They should not be committed to the repository.

## Usage From Source

Build first, then import the root manifest:

```powershell
.\build.ps1
Import-Module .\OfficeScrubC2R.psd1 -Force

Get-InstalledOfficeProducts
Test-OfficeC2RState
Invoke-OfficeScrubC2R -PlanOnly
Invoke-OfficeScrubC2R -WhatIf
```

Real execution requires an elevated shell. Use `-Confirm:$false` only when you are ready for destructive cleanup:

```powershell
Invoke-OfficeScrubC2R -Confirm:$false
Invoke-OfficeScrubC2R -KeepLicense -Confirm:$false
```

## Structured Output

The core output types are:

- `OfficeScrubC2R.OfficeProductInfo`
- `OfficeScrubC2R.OfficeC2RState`
- `OfficeScrubC2R.ScrubPlan`
- `OfficeScrubC2R.OperationResult`

`OperationResult` includes the step, action, target kind, target, registry hive/view when applicable, status, message, exception type, HRESULT, Win32 error, reboot scheduling state, and stable error id.

After execution, inspect:

```powershell
$result = Invoke-OfficeScrubC2R -Confirm:$false
$result.ExecutedOperations | Format-Table Step,Action,Status,Target -AutoSize
```

## Test

```powershell
dotnet test .\tests\OfficeScrubC2R.Core.Tests\OfficeScrubC2R.Core.Tests.csproj
Invoke-Pester -Path .\tests\OfficeScrubC2R.Tests.ps1 -CI
Invoke-ScriptAnalyzer -Path . -Recurse -Settings .\PSScriptAnalyzerSettings.psd1
.\.github\scripts\Validate-Module.ps1
```

CI validates the module in both Windows PowerShell 5.1 and PowerShell 7.

## Production Readiness

Version 3.0.0 is a usable destructive milestone, not full OffScrubC2R parity yet. It does not yet run the Office Deployment Tool uninstall flow, every Windows Installer metadata branch, every COM TypeLib branch, all shell shortcut/taskband cleanup, or multi-profile license cleanup from the original VBScript. Before broad production use, the project still needs signed release artifacts, published checksums, a VM-based destructive test matrix, full scrub parity, and clean Office reinstall validation after cleanup.
