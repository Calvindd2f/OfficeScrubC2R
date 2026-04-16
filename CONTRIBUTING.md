# Contributing to OfficeScrubC2R

Thank you for helping harden OfficeScrubC2R.

## Current Scope

Version 3.0.0 is a guarded destructive binary milestone:

- `Get-InstalledOfficeProducts`
- `Test-OfficeC2RState`
- `Invoke-OfficeScrubC2R`

`Invoke-OfficeScrubC2R` must continue to require elevation, honor `-PlanOnly` and `-WhatIf`, honor preservation switches such as `-KeepLicense`, `-KeepTeams`, and `-KeepCopilot`, and return structured operation results. New destructive branches need tests or narrowly documented lab validation.

## Development Setup

```powershell
git clone https://github.com/Calvindd2f/OfficeScrubC2R.git
cd OfficeScrubC2R

Install-Module Pester -Scope CurrentUser -Force -SkipPublisherCheck
Install-Module PSScriptAnalyzer -Scope CurrentUser -Force

.\build.ps1 -Clean
Import-Module .\OfficeScrubC2R.psd1 -Force
```

## Validation

Run these before opening a pull request:

```powershell
.\build.ps1 -Clean
dotnet test .\tests\OfficeScrubC2R.Core.Tests\OfficeScrubC2R.Core.Tests.csproj
Invoke-Pester -Path .\tests\OfficeScrubC2R.Tests.ps1 -CI
Invoke-ScriptAnalyzer -Path . -Recurse -Settings .\PSScriptAnalyzerSettings.psd1
.\.github\scripts\Validate-Module.ps1 -SkipBuild
```

Also verify import in both hosts when available:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "Import-Module .\OfficeScrubC2R.psd1 -Force; Get-Command -Module OfficeScrubC2R"
pwsh -NoProfile -Command "Import-Module .\OfficeScrubC2R.psd1 -Force; Get-Command -Module OfficeScrubC2R"
```

## Code Standards

- Keep source code under `src/`.
- Keep compiled DLLs, PDBs, packages, and checksums out of git.
- Use explicit `RegistryView.Registry32` and `RegistryView.Registry64`; do not manually rewrite paths through `Wow6432Node`.
- Preserve structured `OperationResult` diagnostics for operational failures.
- Write xUnit tests for core behavior and Pester tests for public command behavior.
- Update `README.md`, `BUILD.md`, and `CHANGELOG.md` for public behavior changes.

## Pull Requests

Include:

- what changed
- why it changed
- validation commands and results
- whether behavior is detection, planning, or destructive
- screenshots or logs only when they add useful evidence

Do not include real endpoint secrets, tenant identifiers, product keys, or unredacted customer logs.

## Release Notes

Release notes must call out whether the release is detection-only, planning-only, or enables cleanup behavior. Any destructive release must document signing, checksums, VM coverage, and reinstall validation evidence.
