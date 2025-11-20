[CmdletBinding()]
param(
    [switch]$Clean,
    [string]$NetCoreFramework = 'net7.0-windows'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceFile = Join-Path $root 'OfficeScrubC2R-Native.cs'
$artifactRoot = Join-Path $root 'artifacts'

if ($Clean -and (Test-Path $artifactRoot)) {
    Write-Verbose 'Cleaning previous build artifacts.'
    Remove-Item $artifactRoot -Recurse -Force
}

if (-not (Test-Path $sourceFile)) {
    throw "Source file not found: $sourceFile"
}

function Invoke-NetFrameworkBuild {
    param(
        [string]$Source,
        [string]$OutputFolder
    )

    $frameworkCandidates = @(
        Join-Path $env:WINDIR 'Microsoft.NET\\Framework64\\v4.0.30319\\csc.exe'),
        Join-Path $env:WINDIR 'Microsoft.NET\\Framework\\v4.0.30319\\csc.exe'
    ) | Where-Object { $_ -and (Test-Path $_) }

    if (-not $frameworkCandidates) {
        Write-Warning '.NET Framework compiler (csc.exe) not found. Skipping Desktop build check.'
        return $null
    }

    $compiler = $frameworkCandidates | Select-Object -First 1
    $compilerDir = Split-Path $compiler -Parent
    $references = @(
        Join-Path $compilerDir 'System.Management.dll',
        Join-Path $compilerDir 'Microsoft.CSharp.dll'
    ) | Where-Object { Test-Path $_ }

    if (-not $references) {
        Write-Warning 'Required reference assemblies missing for .NET Framework build.'
        return $null
    }

    New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
    $outputPath = Join-Path $OutputFolder 'OfficeScrubNative.dll'

    $arguments = @(
        '/target:library',
        '/out:' + $outputPath,
        '/optimize+',
        '/unsafe',
        '/warn:4'
    )

    $arguments += $references | ForEach-Object { '/reference:' + $_ }
    $arguments += $Source

    Write-Host "Building .NET Framework assembly with $compiler" -ForegroundColor Cyan
    $process = Start-Process -FilePath $compiler -ArgumentList $arguments -NoNewWindow -Wait -PassThru

    if ($process.ExitCode -ne 0) {
        throw ".NET Framework build failed with exit code $($process.ExitCode)."
    }

    return $outputPath
}

function Invoke-NetCoreBuild {
    param(
        [string]$Source,
        [string]$OutputFolder,
        [string]$Framework
    )

    $dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
    if (-not $dotnet) {
        Write-Warning '.NET SDK (dotnet) not found. Skipping Core build check.'
        return $null
    }

    New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
    $projectPath = Join-Path $OutputFolder 'OfficeScrubNative.Core.csproj'

    @"
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>$Framework</TargetFramework>
    <OutputType>Library</OutputType>
    <ImplicitUsings>disable</ImplicitUsings>
    <Nullable>disable</Nullable>
    <AllowUnsafeBlocks>true</AllowUnsafeBlocks>
    <RootNamespace>OfficeScrubNative</RootNamespace>
    <AssemblyName>OfficeScrubNative.Core</AssemblyName>
    <PlatformTarget>AnyCPU</PlatformTarget>
  </PropertyGroup>
  <ItemGroup>
    <Compile Include="$Source" />
  </ItemGroup>
</Project>
"@ | Set-Content -Path $projectPath -Encoding UTF8

    Write-Host "Building .NET SDK assembly targeting $Framework" -ForegroundColor Cyan
    dotnet build $projectPath -nologo -clp:NoSummary -p:GenerateAssemblyInfo=false | Write-Host

    $outputPath = Join-Path $OutputFolder "bin\\Debug\\$Framework\\OfficeScrubNative.Core.dll"
    if (-not (Test-Path $outputPath)) {
        throw "Expected build output not found: $outputPath"
    }

    return $outputPath
}

function Test-AssemblyLoad {
    param(
        [string]$AssemblyPath
    )

    if (-not $AssemblyPath) {
        return
    }

    Write-Host "Validating load: $AssemblyPath" -ForegroundColor Green
    [System.Reflection.Assembly]::LoadFrom($AssemblyPath) | Out-Null
    [OfficeScrubNative.OfficeScrubOrchestrator]::new($true) | Out-Null
}

$netfxOutput = Invoke-NetFrameworkBuild -Source $sourceFile -OutputFolder (Join-Path $artifactRoot 'netfx')
$netcoreOutput = Invoke-NetCoreBuild -Source $sourceFile -OutputFolder (Join-Path $artifactRoot 'netcore') -Framework $NetCoreFramework

if ($PSEdition -eq 'Desktop') {
    Test-AssemblyLoad -AssemblyPath $netfxOutput
}
else {
    Test-AssemblyLoad -AssemblyPath $netcoreOutput
}

Write-Host 'Verification complete.' -ForegroundColor Green
