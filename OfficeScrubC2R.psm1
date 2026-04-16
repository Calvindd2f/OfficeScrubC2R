# OfficeScrubC2R.psm1
# Host-aware loader for the binary cmdlet module built from src/.

$script:ModuleRoot = $PSScriptRoot

$script:BinaryCandidates = @(
    (Join-Path $script:ModuleRoot 'lib\netstandard2.0\OfficeScrubC2R.dll'),
    (Join-Path $script:ModuleRoot 'artifacts\module\lib\netstandard2.0\OfficeScrubC2R.dll')
)

$script:BinaryModulePath = $script:BinaryCandidates |
    Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
    Select-Object -First 1

if (-not $script:BinaryModulePath) {
    $searched = $script:BinaryCandidates -join [Environment]::NewLine
    throw "OfficeScrubC2R binary module was not found. Run '.\build.ps1' first. Searched:$([Environment]::NewLine)$searched"
}

Import-Module -Name $script:BinaryModulePath -Force -Global
