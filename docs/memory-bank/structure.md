# OfficeScrubC2R - Project Structure

## Root Files

- `OfficeScrubC2R.psd1`: module manifest exporting the binary cmdlet contract.
- `OfficeScrubC2R.psm1`: loader that imports the built binary module from package or artifact layout.
- `build.ps1`: builds the SDK projects and creates `artifacts/module`.
- `README.md`, `BUILD.md`, `CHANGELOG.md`, `SECURITY.md`, `CONTRIBUTING.md`: current project documentation.

## Source

```text
src/
├── OfficeScrubC2R.Core/
│   ├── Models.cs
│   ├── RegistryAccess.cs
│   ├── OfficeDetectionService.cs
│   ├── PreflightService.cs
│   └── ScrubPlanner.cs
└── OfficeScrubC2R.PowerShell/
    ├── GetInstalledOfficeProductsCommand.cs
    ├── TestOfficeC2RStateCommand.cs
    └── InvokeOfficeScrubC2RCommand.cs
```

## Tests

```text
tests/
├── OfficeScrubC2R.Core.Tests/
│   └── CoreBehaviorTests.cs
├── OfficeScrubC2R.Tests.ps1
└── Verify-NativeBuild.ps1
```

## Reference Material

- `docs/source/OfficeScrubC2R.vbs`: original Microsoft VBScript reference.
- `direct_conversion.cs`: earlier broad conversion reference.
- `OfficeScrubC2R-Native.cs`: earlier helper-library reference.

Reference files are not the v3 runtime implementation.

## Generated Artifacts

```text
artifacts/
├── checksums.sha256
└── module/
    ├── OfficeScrubC2R.psd1
    ├── OfficeScrubC2R.psm1
    └── lib/netstandard2.0/*.dll
```

`artifacts/`, `bin/`, and `obj/` are ignored.
