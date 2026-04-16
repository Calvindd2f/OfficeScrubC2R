@{
    RootModule           = 'OfficeScrubC2R.psm1'
    ModuleVersion        = '3.0.0'
    CompatiblePSEditions = @('Desktop', 'Core')
    GUID                 = 'f8e7c4d1-5a3b-4e2d-9c8f-1b6a4d7e9c3a'
    Author               = 'Calvin'
    CompanyName          = '@Calvindd2f'
    Copyright            = '(c) 2025 Calvin. MIT License. Derived from Microsoft OffScrubC2R.vbs.'

    Description          = @'
Non-destructive v3 hardened baseline for Office Click-to-Run cleanup tooling. The module exposes binary PowerShell cmdlets for Office product detection, C2R state preflight, and scrub planning. Destructive cleanup execution is intentionally blocked until the full scrub parity milestone is implemented and validated.
'@

    PowerShellVersion    = '5.1'
    DotNetFrameworkVersion = '4.7.2'
    ProcessorArchitecture = 'None'

    RequiredModules      = @()
    RequiredAssemblies   = @()
    ScriptsToProcess     = @()
    TypesToProcess       = @()
    FormatsToProcess     = @()
    NestedModules        = @()

    FunctionsToExport    = @()
    CmdletsToExport      = @(
        'Get-InstalledOfficeProducts',
        'Test-OfficeC2RState',
        'Invoke-OfficeScrubC2R'
    )
    VariablesToExport    = @()
    AliasesToExport      = @()
    DscResourcesToExport = @()

    ModuleList           = @()
    FileList             = @(
        'OfficeScrubC2R.psd1',
        'OfficeScrubC2R.psm1',
        'LICENSE',
        'README.md',
        'CHANGELOG.md',
        'SECURITY.md'
    )

    PrivateData          = @{
        PSData = @{
            Tags = @(
                'Office',
                'Microsoft',
                'ClickToRun',
                'C2R',
                'Uninstall',
                'Removal',
                'Scrub',
                'Office365',
                'O365',
                'Administration',
                'Maintenance',
                'Windows',
                'PSEdition_Desktop',
                'PSEdition_Core'
            )
            LicenseUri               = 'https://github.com/Calvindd2f/OfficeScrubC2R/blob/main/LICENSE'
            ProjectUri               = 'https://github.com/Calvindd2f/OfficeScrubC2R'
            IconUri                  = ''
            ReleaseNotes             = @'
# Release Notes v3.0.0

## Hardened Binary Baseline
- Exposes the supported public contract as binary cmdlets: Get-InstalledOfficeProducts, Test-OfficeC2RState, and Invoke-OfficeScrubC2R.
- Adds structured detection, preflight, scrub-plan, and operation-result output.
- Blocks destructive scrub execution while preserving -PlanOnly and -WhatIf planning behavior.
- Treats compiled binaries as build/release artifacts instead of source files.
'@
            RequireLicenseAcceptance = $false
            ExternalModuleDependencies = @()
        }
    }

    HelpInfoURI          = 'https://github.com/Calvindd2f/OfficeScrubC2R/blob/main/README.md'
    DefaultCommandPrefix = ''
}
