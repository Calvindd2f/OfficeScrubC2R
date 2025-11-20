# OfficeScrubC2R.psm1
# Lightweight wrapper that exposes the native OfficeScrub C# binary via PowerShell.
# The module intentionally focuses on the converted DLL that uses native Windows APIs
# while keeping the original reference sources in the repository for contrast.

$script:ModuleRoot = $PSScriptRoot
$script:NativeAssemblyPath = Join-Path $script:ModuleRoot 'OfficeScrubNative.dll'
$script:NativeAssembly = $null

function Import-OfficeScrubNative {
    [CmdletBinding()]
    param(
        # Path to the OfficeScrub native DLL. Defaults to the copy next to the module.
        [string]$Path = $script:NativeAssemblyPath
    )

    $resolvedPath = (Resolve-Path -Path $Path -ErrorAction Stop).ProviderPath

    if (-not (Test-Path -Path $resolvedPath -PathType Leaf)) {
        throw "Native assembly not found at '$resolvedPath'. Ensure OfficeScrubNative.dll is present alongside the module."
    }

    if ($script:NativeAssembly -and $script:NativeAssembly.Location -eq $resolvedPath) {
        return $script:NativeAssembly
    }

    $script:NativeAssembly = [AppDomain]::CurrentDomain.GetAssemblies() |
        Where-Object { $_.Location -eq $resolvedPath }

    if (-not $script:NativeAssembly) {
        $script:NativeAssembly = [System.Reflection.Assembly]::LoadFrom($resolvedPath)
    }

    return $script:NativeAssembly
}

function New-OfficeScrubOrchestrator {
    [CmdletBinding()]
    param(
        # Whether to use 64-bit registry/file views when interacting with the OS.
        [bool]$Is64Bit = [Environment]::Is64BitOperatingSystem
    )

    Import-OfficeScrubNative | Out-Null
    return [OfficeScrubNative.OfficeScrubOrchestrator]::new($Is64Bit)
}

function Test-OfficeC2RPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [OfficeScrubNative.OfficeScrubOrchestrator]$Orchestrator
    )

    if (-not $PSBoundParameters.ContainsKey('Orchestrator') -or -not $Orchestrator) {
        $Orchestrator = New-OfficeScrubOrchestrator
    }

    return $Orchestrator.IsC2RPath($Path)
}

function Test-OfficeProductScope {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ProductCode,
        [OfficeScrubNative.OfficeScrubOrchestrator]$Orchestrator
    )

    if (-not $PSBoundParameters.ContainsKey('Orchestrator') -or -not $Orchestrator) {
        $Orchestrator = New-OfficeScrubOrchestrator
    }

    return $Orchestrator.IsInScope($ProductCode)
}

function Get-OfficeScrubHelpers {
    [CmdletBinding()]
    param(
        [OfficeScrubNative.OfficeScrubOrchestrator]$Orchestrator
    )

    if (-not $PSBoundParameters.ContainsKey('Orchestrator') -or -not $Orchestrator) {
        $Orchestrator = New-OfficeScrubOrchestrator
    }

    [pscustomobject]@{
        RegistryHelper         = $Orchestrator.Registry
        FileHelper             = $Orchestrator.Files
        ProcessHelper          = $Orchestrator.Processes
        ShellHelper            = $Orchestrator.Shell
        WindowsInstallerHelper = $Orchestrator.WindowsInstaller
        TypeLibHelper          = $Orchestrator.TypeLib
        LicenseHelper          = $Orchestrator.License
        ServiceHelper          = $Orchestrator.Services
    }
}

Export-ModuleMember -Function Import-OfficeScrubNative, New-OfficeScrubOrchestrator, Test-OfficeC2RPath, Test-OfficeProductScope, Get-OfficeScrubHelpers
