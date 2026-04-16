[CmdletBinding()]
param(
    [switch]$Clean,
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$repoRoot = $PSScriptRoot
$artifactRoot = Join-Path $repoRoot 'artifacts'
$moduleRoot = Join-Path $artifactRoot 'module'
$moduleLib = Join-Path $moduleRoot 'lib\netstandard2.0'
$currentModulePathFile = Join-Path $artifactRoot 'current-module.txt'
$projectPath = Join-Path $repoRoot 'src\OfficeScrubC2R.PowerShell\OfficeScrubC2R.PowerShell.csproj'
$buildOutput = Join-Path $repoRoot "src\OfficeScrubC2R.PowerShell\bin\$Configuration\netstandard2.0"

if ($Clean -and (Test-Path -LiteralPath $artifactRoot)) {
    try {
        Remove-Item -LiteralPath $artifactRoot -Recurse -Force
    }
    catch {
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $moduleRoot = Join-Path $artifactRoot "module-$stamp"
        $moduleLib = Join-Path $moduleRoot 'lib\netstandard2.0'
        Write-Warning "Could not remove '$artifactRoot' because files may be loaded by PowerShell. Writing fresh artifact to '$moduleRoot'."
    }
}

if (-not (Test-Path -LiteralPath $projectPath -PathType Leaf)) {
    throw "Project file not found: $projectPath"
}

Write-Host "Building OfficeScrubC2R binary module ($Configuration)" -ForegroundColor Cyan
dotnet build $projectPath -c $Configuration -nologo

if (-not (Test-Path -LiteralPath (Join-Path $buildOutput 'OfficeScrubC2R.dll') -PathType Leaf)) {
    throw "Expected binary module was not created: $(Join-Path $buildOutput 'OfficeScrubC2R.dll')"
}

New-Item -ItemType Directory -Force -Path $moduleLib | Out-Null

Get-ChildItem -LiteralPath $buildOutput -File |
    Where-Object { $_.Extension -in '.dll', '.pdb', '.deps.json' } |
    Copy-Item -Destination $moduleLib -Force

foreach ($file in @('OfficeScrubC2R.psd1', 'OfficeScrubC2R.psm1', 'LICENSE', 'README.md', 'CHANGELOG.md', 'SECURITY.md')) {
    $source = Join-Path $repoRoot $file
    if (Test-Path -LiteralPath $source -PathType Leaf) {
        Copy-Item -LiteralPath $source -Destination $moduleRoot -Force
    }
}

$checksumPath = Join-Path $artifactRoot 'checksums.sha256'
Get-ChildItem -LiteralPath $moduleRoot -Recurse -File |
    Where-Object { $_.Extension -in '.dll', '.psd1', '.psm1' } |
    Sort-Object FullName |
    ForEach-Object {
        $hash = Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256
        '{0}  {1}' -f $hash.Hash.ToLowerInvariant(), $_.FullName.Substring($moduleRoot.Length + 1).Replace('\', '/')
    } |
    Set-Content -LiteralPath $checksumPath -Encoding ASCII

$relativeModuleRoot = $moduleRoot.Substring($artifactRoot.Length + 1)
Set-Content -LiteralPath $currentModulePathFile -Value $relativeModuleRoot -Encoding ASCII

Write-Host "Module artifact: $moduleRoot" -ForegroundColor Green
Write-Host "Checksums: $checksumPath" -ForegroundColor Green
