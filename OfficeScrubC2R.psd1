@{
    # Script module or binary module file associated with this manifest.
    RootModule             = 'OfficeScrubC2R.psm1'

    # Version number of this module.
    ModuleVersion          = '2.19.0'

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
Complete PowerShell/C# implementation of Microsoft Office Scrub C2R tool. Provides comprehensive removal of Office 2013, 2016, 2019, and Office 365 Click-to-Run installations with 10-50x performance improvements over the original VBScript version.

Features:
- Native C# library for high-performance registry and file operations
- Parallel processing for faster execution
- Comprehensive logging and error handling
- Support for Windows 7 SP1 through Windows 11
- Compatible with PowerShell 5.1 and PowerShell 7+

Requires Administrator privileges.
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
    RequiredAssemblies     = @('OfficeScrubNative.dll')

    # Script files (.ps1) that are run in the caller's environment prior to importing this module.
    ScriptsToProcess       = @()

    # Type files (.ps1xml) to be loaded when importing this module
    TypesToProcess         = @()

    # Format files (.ps1xml) to be loaded when importing this module
    FormatsToProcess       = @()

    # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
    NestedModules          = @('OfficeScrubC2R-Utilities.psm1')

    # Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
    FunctionsToExport      = @(
        'Invoke-OfficeScrubC2R',
        'Get-InstalledOfficeProducts',
        'Test-IsC2R',
        'Initialize-Environment',
        'Stop-OfficeProcesses'
    )

    # Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
    CmdletsToExport        = @()

    # Variables to export from this module
    VariablesToExport      = @()

    # Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
    AliasesToExport        = @('Remove-OfficeC2R', 'Uninstall-OfficeC2R')

    # DSC resources to export from this module
    DscResourcesToExport   = @()

    # List of all modules packaged with this module
    ModuleList             = @()

    # List of all files packaged with this module
    FileList               = @(
        'OfficeScrubC2R.psd1',
        'OfficeScrubC2R.psm1',
        'OfficeScrubC2R-Utilities.psm1',
        'OfficeScrubC2R-Native.cs',
        'OfficeScrubNative.dll',
        'build.ps1',
        'LICENSE',
        'README.md',
        'CHANGELOG.md'
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
# Release Notes v2.19.0

## Overview
Complete PowerShell/C# port of Microsoft's OffScrubC2R.vbs v2.19 with significant performance improvements.

## What's New
- Native C# library for high-performance operations (10-50x faster)
- Pre-compiled DLL with automatic fallback to source compilation
- Parallel processing for registry and file operations
- Comprehensive logging and error handling
- Support for PowerShell 7+ and Windows PowerShell 5.1

## Breaking Changes
- Requires Administrator privileges
- Minimum PowerShell version is 5.1
- .NET Framework 4.5 or later required

## Known Issues
- None

For full changelog, see CHANGELOG.md
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
