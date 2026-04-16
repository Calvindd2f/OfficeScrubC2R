[CmdletBinding()]
param(
    [switch]$SkipBuild
)

$ErrorActionPreference = 'Stop'
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..')
$expectedCommands = @('Get-InstalledOfficeProducts', 'Invoke-OfficeScrubC2R', 'Test-OfficeC2RState')

Push-Location $repoRoot
try {
    if (-not $SkipBuild) {
        .\build.ps1 -Clean
    }

    $manifest = Test-ModuleManifest -Path .\OfficeScrubC2R.psd1
    if ($manifest.Version.ToString() -ne '3.0.0') {
        throw "Unexpected module version: $($manifest.Version)"
    }

    Import-Module .\OfficeScrubC2R.psd1 -Force
    $actualCommands = Get-Command -Module OfficeScrubC2R | Select-Object -ExpandProperty Name | Sort-Object
    if (Compare-Object -ReferenceObject $expectedCommands -DifferenceObject $actualCommands) {
        throw "Unexpected command exports. Expected: $($expectedCommands -join ', '); Actual: $($actualCommands -join ', ')"
    }

    $trackedBinaries = @(git ls-files '*.dll' '*.pdb' | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf })
    if ($trackedBinaries.Count -gt 0) {
        throw "Compiled binaries must not be tracked: $($trackedBinaries -join ', ')"
    }

    if (-not (Test-Path .\artifacts\checksums.sha256 -PathType Leaf)) {
        throw 'Checksum file was not generated.'
    }

    if (-not (Test-Path .\artifacts\current-psgallery-module.txt -PathType Leaf)) {
        throw 'PSGallery package pointer was not generated.'
    }

    $galleryRelativePath = Get-Content -LiteralPath .\artifacts\current-psgallery-module.txt -TotalCount 1
    $galleryModulePath = Join-Path (Resolve-Path .\artifacts) $galleryRelativePath
    if ((Split-Path -Leaf $galleryModulePath) -ne 'OfficeScrubC2R') {
        throw "PSGallery package folder must be named OfficeScrubC2R: $galleryModulePath"
    }

    Test-ModuleManifest -Path (Join-Path $galleryModulePath 'OfficeScrubC2R.psd1') | Out-Null

    Write-Host 'OfficeScrubC2R module validation passed.' -ForegroundColor Green
}
finally {
    Pop-Location
}
