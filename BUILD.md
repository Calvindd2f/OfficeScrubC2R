# Building OfficeScrubC2R

OfficeScrubC2R v3 builds a binary PowerShell module from the SDK-style projects in `src/`.

## Quick Build

```powershell
.\build.ps1 -Clean
```

The build creates:

```text
artifacts/
  checksums.sha256
  module/
    OfficeScrubC2R.psd1
    OfficeScrubC2R.psm1
    lib/
      netstandard2.0/
        OfficeScrubC2R.dll
        OfficeScrubC2R.Core.dll
        dependency DLLs
```

The repository root does not receive compiled DLLs. Build outputs remain ignored artifacts.

## Projects

- `src/OfficeScrubC2R.Core`: detection, registry-view access, preflight state, scrub plans, guarded cleanup execution including Teams/Copilot companion app cleanup, and structured operation results.
- `src/OfficeScrubC2R.PowerShell`: binary cmdlets for `Get-InstalledOfficeProducts`, `Test-OfficeC2RState`, and `Invoke-OfficeScrubC2R`.
- `tests/OfficeScrubC2R.Core.Tests`: xUnit tests for core behavior.
- `tests/OfficeScrubC2R.Tests.ps1`: Pester tests for the PowerShell module contract.

## Requirements

- Windows 10 or Windows 11
- Windows PowerShell 5.1 and/or PowerShell 7+
- .NET SDK 8.0+
- Pester 5+ for PowerShell contract tests
- PSScriptAnalyzer for PowerShell linting

## Validation

```powershell
.\build.ps1 -Clean
dotnet test .\tests\OfficeScrubC2R.Core.Tests\OfficeScrubC2R.Core.Tests.csproj
Invoke-Pester -Path .\tests\OfficeScrubC2R.Tests.ps1 -CI
Invoke-ScriptAnalyzer -Path . -Recurse -Settings .\PSScriptAnalyzerSettings.psd1
.\.github\scripts\Validate-Module.ps1 -SkipBuild
```

`Validate-Module.ps1` verifies the manifest, exact command exports, generated checksums, and absence of tracked `.dll` or `.pdb` files.

If PowerShell has already imported a previous artifact, Windows can hold DLL locks under `artifacts/module`. `build.ps1 -Clean` falls back to a timestamped `artifacts/module-*` folder and records it in `artifacts/current-module.txt`; the root loader reads that file on the next import.

## Artifact Policy

Compiled binaries are not source files. Do not commit DLLs or PDBs. Release automation should publish signed artifacts and SHA256 checksums from `artifacts/module`.
