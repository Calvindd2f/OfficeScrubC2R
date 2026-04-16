# Changelog

All notable changes to this project are documented here.

The project follows semantic versioning.

## [3.0.0] - 2026-04-16

### Added

- Binary PowerShell cmdlet baseline with exactly three public commands:
  - `Get-InstalledOfficeProducts`
  - `Test-OfficeC2RState`
  - `Invoke-OfficeScrubC2R`
- Structured output types for products, preflight state, scrub plans, and operation results.
- Explicit 32-bit and 64-bit registry view access through `RegistryView`.
- Non-destructive scrub planning through `Invoke-OfficeScrubC2R -PlanOnly` and `-WhatIf`.
- Guarded destructive execution through `Invoke-OfficeScrubC2R` when run elevated.
- Cleanup executor steps for Office process termination, Office scheduled task deletion, Click-to-Run service stop/delete, C2R registry key/value deletion through explicit registry views, known/detected C2R folder deletion, reboot delete scheduling fallback, Teams/Copilot local app cleanup, and current-user licensing cache deletion unless `-KeepLicense` is used.
- `Invoke-OfficeScrubC2R -KeepTeams` and `-KeepCopilot` switches for preserving Office-adjacent companion apps during real cleanup.
- `Invoke-OfficeScrubC2R` accepts legacy v2 invocation switches `-Quiet`, `-Force`, and `-RemoveAll` so older automation does not fail at parameter binding before reaching v3 behavior.
- xUnit tests for core behavior and Pester tests for module import/command behavior.
- CI gates for build, tests, PowerShell 5.1 import, PowerShell 7 import, analyzer checks, source hygiene, and checksum generation.

### Changed

- The repository now treats compiled DLLs and PDBs as build/release artifacts only.
- The compatibility floor is Windows 10/11 with Windows PowerShell 5.1 or PowerShell 7+.
- Build output now goes to `artifacts/module` instead of the repository root.
- Build output now also includes a PowerShell Gallery-ready `artifacts/psgallery/OfficeScrubC2R` package folder.
- Teams and Copilot are treated as local companion app cleanup targets because modern Office installations can leave them behind after the older OffScrubC2R flow completes.

### Removed

- The helper-only public surface from the v3 module contract.
- Runtime source-fallback compilation documentation.
- The checked-in `OfficeScrubNative.dll` artifact.

### Security

- Destructive scrub execution requires elevation and blocks with `OfficeScrubC2R.AdminRequired` when not elevated.
- Teams and Copilot cleanup is limited to local packages/installers/profile remnants. It does not remove Microsoft 365 Copilot tenant licensing or cloud app state.
- Full OffScrubC2R parity remains incomplete; broad production use still requires signed artifacts, VM destructive-path coverage, and clean reinstall validation.

## Historical Notes

Older 2.x entries described a script-heavy/module-helper packaging line and PowerShell Gallery package attempts. They are no longer the authoritative description of the v3 module contract.
