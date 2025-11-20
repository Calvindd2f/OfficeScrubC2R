@{
    # Script module or binary module file associated with this manifest.
    RootModule             = 'OfficeScrubC2R.psm1'

    # Version number of this module.
    ModuleVersion          = '3.0.0'

    # Supported PSEditions
    CompatiblePSEditions   = @('Desktop', 'Core')

    # ID used to uniquely identify this module
    GUID                   = 'f8e7c4d1-5a3b-4e2d-9c8f-1b6a4d7e9c3a'

    # Author of this module
    Author                 = 'Calvin'

    # Company or vendor of this module
    CompanyName            = '@Calvindd2f'

    # Copyright statement for this module
    Copyright              = '(c) 2025 Calvin. All rights reserved. MIT License. Derived from Microsoft OffScrubC2R.vbs.'

    # Description of the functionality provided by this module
    Description            = @'
Lightweight PowerShell wrapper for the converted OfficeScrub native DLL. The module focuses on surfacing the native methods/Windows API code through a small set of helper functions while keeping the original VBScript and C# reference sources in the repository for contrast.
'@

    # Minimum version of the PowerShell engine required by this module
    PowerShellVersion      = '5.1'

    # Minimum version of the .NET Framework required by this module
    DotNetFrameworkVersion = '4.5'

    # Processor architecture (None, X86, Amd64) required by this module
    ProcessorArchitecture  = 'None'

    # Modules that must be imported into the global environment prior to importing this module
    RequiredModules        = @()

    # Assemblies that must be loaded prior to importing this module
    # Note: OfficeScrubNative.dll is loaded dynamically by Initialize-NativeTypes
    # to avoid conflicts when the assembly is already loaded
    RequiredAssemblies     = @()

    # Script files (.ps1) that are run in the caller's environment prior to importing this module.
    ScriptsToProcess       = @()

    # Type files (.ps1xml) to be loaded when importing this module
    TypesToProcess         = @()

    # Format files (.ps1xml) to be loaded when importing this module
    FormatsToProcess       = @()

    # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
    NestedModules          = @()

    # Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
    FunctionsToExport      = @(
        'Import-OfficeScrubNative',
        'New-OfficeScrubOrchestrator',
        'Test-OfficeC2RPath',
        'Test-OfficeProductScope',
        'Get-OfficeScrubHelpers'
    )

    # Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
    CmdletsToExport        = @()

    # Variables to export from this module
    VariablesToExport      = @()

    # Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
    AliasesToExport        = @()

    # DSC resources to export from this module
    DscResourcesToExport   = @()

    # List of all modules packaged with this module
    ModuleList             = @()

    # List of all files packaged with this module
    FileList               = @(
        'OfficeScrubC2R.psd1',
        'OfficeScrubC2R.psm1',
        'OfficeScrubNative.dll',
        'OfficeScrubC2R-Native.cs',
        'direct_conversion.cs',
        'docs\source\OfficeScrubC2R.vbs',
        'LICENSE',
        'README.md'
    )

    # Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
    PrivateData            = @{
        PSData = @{
            # Tags applied to this module. These help with module discovery in online galleries.
            Tags                       = @(
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

            # A URL to the license for this module.
            LicenseUri                 = 'https://github.com/Calvindd2f/OfficeScrubC2R/blob/main/LICENSE'

            # A URL to the main website for this project.
            ProjectUri                 = 'https://github.com/Calvindd2f/OfficeScrubC2R'

            # A URL to an icon representing this module.
            IconUri                    = ''

            # ReleaseNotes of this module
            ReleaseNotes               = @'
# Release Notes v3.0.0

## Overview
The module now wraps the native OfficeScrub DLL directly, exposing a handful of helper commands instead of the previous script-heavy implementation. Original VBScript and C# sources remain in the repo for reference.

## What's New
- Streamlined PowerShell module focused on loading the native binary
- Helper functions for creating the orchestrator and checking product scope
- Repository cleaned to highlight the native binary alongside the original sources

## Breaking Changes
- Legacy script-based functions have been removed in favor of the native wrapper
'@

            # Prerelease string of this module
            # Prerelease = 'preview'

            # Flag to indicate whether the module requires explicit user acceptance for install/update/save
            RequireLicenseAcceptance   = $false

            # External dependent modules of this module
            ExternalModuleDependencies = @()

        } # End of PSData hashtable

    } # End of PrivateData hashtable

    # HelpInfo URI of this module
    HelpInfoURI            = 'https://github.com/Calvindd2f/OfficeScrubC2R/blob/main/README.md'

    # Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
    DefaultCommandPrefix   = ''
}
