@{
    RootModule           = 'OfficeScrubC2R.psm1'
    ModuleVersion        = '3.0.0'
    CompatiblePSEditions = @('Desktop', 'Core')
    GUID                 = 'f8e7c4d1-5a3b-4e2d-9c8f-1b6a4d7e9c3a'
    Author               = 'Calvin'
    CompanyName          = '@Calvindd2f'
    Copyright            = '(c) 2025 Calvin. MIT License. Derived from Microsoft OffScrubC2R.vbs.'

    Description          = @'
Hardened v3 binary Office Click-to-Run cleanup module. The module exposes binary PowerShell cmdlets for Office product detection, C2R state preflight, scrub planning, and guarded destructive cleanup with structured operation results.
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

## Hardened Binary Cleanup Baseline
- Exposes the supported public contract as binary cmdlets: Get-InstalledOfficeProducts, Test-OfficeC2RState, and Invoke-OfficeScrubC2R.
- Adds structured detection, preflight, scrub-plan, and operation-result output.
- Runs guarded destructive cleanup from elevated sessions while preserving -PlanOnly and -WhatIf planning behavior.
- Adds local Teams and Copilot companion app cleanup with -KeepTeams and -KeepCopilot preservation switches.
- Accepts legacy -Quiet, -Force, and -RemoveAll switches for v2 invocation compatibility.
- Treats compiled binaries as build/release artifacts instead of source files.
'@
            RequireLicenseAcceptance = $false
            ExternalModuleDependencies = @()
        }
    }

    HelpInfoURI          = 'https://github.com/Calvindd2f/OfficeScrubC2R/blob/main/README.md'
    DefaultCommandPrefix = ''
}
