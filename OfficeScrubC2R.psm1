# OfficeScrubC2R.psm1
# Host-aware loader for the binary cmdlet module built from src/.

$script:ModuleRoot = $PSScriptRoot
$script:ArtifactRoot = Join-Path $script:ModuleRoot 'artifacts'
$script:CurrentModulePathFile = Join-Path $script:ArtifactRoot 'current-module.txt'

$script:CurrentArtifactBinary = $null
if (Test-Path -LiteralPath $script:CurrentModulePathFile -PathType Leaf) {
    $script:CurrentArtifactRoot = Get-Content -LiteralPath $script:CurrentModulePathFile -TotalCount 1
    if (-not [string]::IsNullOrWhiteSpace($script:CurrentArtifactRoot)) {
        $script:CurrentArtifactBinary = Join-Path $script:ArtifactRoot (Join-Path $script:CurrentArtifactRoot 'lib\netstandard2.0\OfficeScrubC2R.dll')
    }
}

$script:BinaryCandidates = @(
    (Join-Path $script:ModuleRoot 'lib\netstandard2.0\OfficeScrubC2R.dll'),
    $script:CurrentArtifactBinary,
    (Join-Path $script:ModuleRoot 'artifacts\module\lib\netstandard2.0\OfficeScrubC2R.dll')
) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

$script:BinaryModulePath = $script:BinaryCandidates |
    Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
    Select-Object -First 1

if (-not $script:BinaryModulePath) {
    $searched = $script:BinaryCandidates -join [Environment]::NewLine
    throw "OfficeScrubC2R binary module was not found. Run '.\build.ps1' first. Searched:$([Environment]::NewLine)$searched"
}

Import-Module -Name $script:BinaryModulePath -Force -Global
