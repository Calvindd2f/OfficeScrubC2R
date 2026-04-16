# OfficeScrubC2R - Product Overview

## Project Purpose

OfficeScrubC2R v3 is a hardened binary PowerShell module for inspecting and scrubbing Microsoft Office Click-to-Run installations. The current milestone includes guarded destructive execution for a first VBS-inspired cleanup slice.

## Current Value

- Detects Office C2R products from configuration and uninstall registry evidence.
- Reports preflight state: elevation, SYSTEM context, Office processes, C2R services, package paths, pending reboot-delete evidence, and issues.
- Produces structured scrub plans and structured execution results.
- Uses explicit 32-bit and 64-bit registry views instead of string-based `Wow6432Node` path rewriting.
- Emits structured operation diagnostics through `OperationResult`.

## Public Commands

- `Get-InstalledOfficeProducts`
- `Test-OfficeC2RState`
- `Invoke-OfficeScrubC2R`

`Invoke-OfficeScrubC2R` supports `-PlanOnly` and `-WhatIf`. Real destructive execution requires elevation and blocks with `OfficeScrubC2R.AdminRequired` when not elevated.

## Target Users

- IT administrators investigating broken Office C2R installs.
- Support engineers preparing evidence for a cleanup or reinstall workflow.
- Maintainers hardening a future destructive scrub implementation.

## Requirements

- Windows 10 or Windows 11
- Windows PowerShell 5.1 or PowerShell 7+
- .NET Framework 4.7.2 or later for Windows PowerShell hosts

## Release Readiness Boundary

The v3 milestone is usable for lab/pilot cleanup, but it is not full OffScrubC2R parity. Broad production cleanup requires signing, checksums, VM destructive-path tests, full scrub parity, and clean Office reinstall validation.
