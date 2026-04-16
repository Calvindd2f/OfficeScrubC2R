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
- Cleanup executor steps for Office process termination, Office scheduled task deletion, Click-to-Run service stop/delete, C2R registry key/value deletion through explicit registry views, known/detected C2R folder deletion, reboot delete scheduling fallback, and current-user licensing cache deletion unless `-KeepLicense` is used.
- xUnit tests for core behavior and Pester tests for module import/command behavior.
- CI gates for build, tests, PowerShell 5.1 import, PowerShell 7 import, analyzer checks, source hygiene, and checksum generation.

### Changed

- The repository now treats compiled DLLs and PDBs as build/release artifacts only.
- The compatibility floor is Windows 10/11 with Windows PowerShell 5.1 or PowerShell 7+.
- Build output now goes to `artifacts/module` instead of the repository root.

### Removed

- The helper-only public surface from the v3 module contract.
- Runtime source-fallback compilation documentation.
- The checked-in `OfficeScrubNative.dll` artifact.

### Security

- Destructive scrub execution requires elevation and blocks with `OfficeScrubC2R.AdminRequired` when not elevated.
- Full OffScrubC2R parity remains incomplete; broad production use still requires signed artifacts, VM destructive-path coverage, and clean reinstall validation.

## Historical Notes

Older 2.x entries described a script-heavy/module-helper packaging line and PowerShell Gallery package attempts. They are no longer the authoritative description of the v3 module contract.
