# OfficeScrubC2R - Technology Stack

## Languages

- C# for core detection, preflight, planning, and binary cmdlets.
- PowerShell for build, validation, module loading, CI scripts, and Pester tests.

## Projects

- `src/OfficeScrubC2R.Core`: netstandard2.0 core library.
- `src/OfficeScrubC2R.PowerShell`: netstandard2.0 binary cmdlet module using `PowerShellStandard.Library`.
- `tests/OfficeScrubC2R.Core.Tests`: xUnit tests.
- `tests/OfficeScrubC2R.Tests.ps1`: Pester module contract tests.

## Build

```powershell
.\build.ps1 -Clean
```

Build output goes to `artifacts/module`. SHA256 checksums go to `artifacts/checksums.sha256`.

## Runtime

- Windows PowerShell 5.1
- PowerShell 7+
- Windows 10/11
- .NET Framework 4.7.2+ for Windows PowerShell hosts

## Registry Strategy

Core code uses `Microsoft.Win32.RegistryKey.OpenBaseKey` with explicit `RegistryView.Registry64` and `RegistryView.Registry32`. New source code must not manually inject `Wow6432Node` into registry paths.

## Validation Commands

```powershell
.\build.ps1 -Clean
dotnet test .\tests\OfficeScrubC2R.Core.Tests\OfficeScrubC2R.Core.Tests.csproj
Invoke-Pester -Path .\tests\OfficeScrubC2R.Tests.ps1 -CI
Invoke-ScriptAnalyzer -Path . -Recurse -Settings .\PSScriptAnalyzerSettings.psd1
.\.github\scripts\Validate-Module.ps1 -SkipBuild
```

## Artifact Policy

DLLs, PDBs, release packages, and checksums are generated artifacts. They are not committed to source control.
