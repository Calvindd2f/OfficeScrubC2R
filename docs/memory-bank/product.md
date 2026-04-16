# OfficeScrubC2R - Product Overview

## Project Purpose

OfficeScrubC2R v3 is a hardened binary PowerShell module for inspecting Microsoft Office Click-to-Run installations and planning future scrub actions. The current baseline is intentionally non-destructive.

## Current Value

- Detects Office C2R products from configuration and uninstall registry evidence.
- Reports preflight state: elevation, SYSTEM context, Office processes, C2R services, package paths, pending reboot-delete evidence, and issues.
- Produces structured scrub plans without deleting anything.
- Uses explicit 32-bit and 64-bit registry views instead of string-based `Wow6432Node` path rewriting.
- Emits structured operation diagnostics through `OperationResult`.

## Public Commands

- `Get-InstalledOfficeProducts`
- `Test-OfficeC2RState`
- `Invoke-OfficeScrubC2R`

`Invoke-OfficeScrubC2R` supports `-PlanOnly` and `-WhatIf`. Real destructive execution is blocked with `OfficeScrubC2R.DestructiveExecutionNotSupported`.

## Target Users

- IT administrators investigating broken Office C2R installs.
- Support engineers preparing evidence for a cleanup or reinstall workflow.
- Maintainers hardening a future destructive scrub implementation.

## Requirements

- Windows 10 or Windows 11
- Windows PowerShell 5.1 or PowerShell 7+
- .NET Framework 4.7.2 or later for Windows PowerShell hosts

## Release Readiness Boundary

The v3 baseline is not a production scrubber. Destructive cleanup requires a later milestone with signing, checksums, VM destructive-path tests, full scrub parity, and clean Office reinstall validation.
