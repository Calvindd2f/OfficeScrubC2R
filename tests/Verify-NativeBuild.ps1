[CmdletBinding()]
param(
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$buildScript = Join-Path $repoRoot 'build.ps1'
$moduleRoot = Join-Path $repoRoot 'artifacts\module'
$binaryModule = Join-Path $moduleRoot 'lib\netstandard2.0\OfficeScrubC2R.dll'
$checksumFile = Join-Path $repoRoot 'artifacts\checksums.sha256'

if (-not (Test-Path -LiteralPath $buildScript -PathType Leaf)) {
    throw "Build script not found: $buildScript"
}

& $buildScript -Clean:$Clean

foreach ($path in @($binaryModule, $checksumFile)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Expected build artifact not found: $path"
    }
}

Import-Module (Join-Path $repoRoot 'OfficeScrubC2R.psd1') -Force
$expected = @('Get-InstalledOfficeProducts', 'Invoke-OfficeScrubC2R', 'Test-OfficeC2RState')
$actual = Get-Command -Module OfficeScrubC2R | Select-Object -ExpandProperty Name | Sort-Object

if (Compare-Object -ReferenceObject $expected -DifferenceObject $actual) {
    throw "Unexpected exported command set. Expected: $($expected -join ', '); Actual: $($actual -join ', ')"
}

Write-Host 'Verification complete.' -ForegroundColor Green
