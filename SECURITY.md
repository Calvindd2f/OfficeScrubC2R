# Security Policy

OfficeScrubC2R is intended for administrators working on failed Microsoft Office Click-to-Run installations. The project can eventually affect registry keys, files, services, licensing data, scheduled tasks, and installer metadata, so release trust and diagnostics matter.

## Supported Baseline

Version 3.0.0 includes guarded destructive cleanup for a first OffScrubC2R-inspired slice. It requires elevation and reports every attempted operation through structured `OperationResult` records. Local Teams and Copilot companion app cleanup is included by default, with `-KeepTeams` and `-KeepCopilot` available when those apps should be preserved. Full parity, signed releases, destructive-path VM coverage, and reinstall validation are still required before broad production use.

Supported runtime baseline:

- Windows 10 or Windows 11
- Windows PowerShell 5.1 or PowerShell 7+
- .NET Framework 4.7.2 or later for Windows PowerShell hosts

## Reporting Security Issues

Do not open public issues for vulnerabilities that could put users or endpoints at risk. Contact the maintainer privately through the GitHub repository owner profile and include:

- affected version or commit
- exact command or workflow involved
- expected behavior
- observed behavior
- logs or structured `OperationResult` output with secrets removed
- whether the issue requires administrator or SYSTEM context

## Release Trust Requirements

A production destructive release must include:

- reproducible build from source
- signed module artifacts
- signed binary artifacts
- published SHA256 checksums
- CI evidence for Windows PowerShell 5.1 and PowerShell 7+
- VM test evidence for locked files, denied ACLs, broken C2R state, x86/x64 registry views, stale MSI metadata, and clean Office reinstall after scrub
- validation that Teams and Copilot local package cleanup behaves as expected without altering tenant-side licensing or policy state

Until those gates exist, treat this project as a lab/pilot cleanup utility rather than a broad production cleanup utility.
