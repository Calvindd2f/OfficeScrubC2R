#
# Build script for OfficeScrubC2R Native DLL
# Compiles OfficeScrubC2R-Native.cs into OfficeScrubNative.dll
#

[CmdletBinding()]
param(
    [switch]$Clean
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "OfficeScrubC2R Native Build Script" -ForegroundColor Cyan
Write-Host "==================================`n" -ForegroundColor Cyan

# Clean previous build
if ($Clean -and (Test-Path "$scriptDir\OfficeScrubNative.dll")) {
    Write-Host "Cleaning previous build..." -ForegroundColor Yellow
    Remove-Item "$scriptDir\OfficeScrubNative.dll" -Force
    if (Test-Path "$scriptDir\OfficeScrubNative.pdb") {
        Remove-Item "$scriptDir\OfficeScrubNative.pdb" -Force
    }
}

# Locate csc.exe
$cscPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $cscPath)) {
    Write-Error "C# compiler not found at: $cscPath"
    exit 1
}

# Check source file
$sourceFile = Join-Path $scriptDir "OfficeScrubC2R-Native.cs"
if (-not (Test-Path $sourceFile)) {
    Write-Error "Source file not found: $sourceFile"
    exit 1
}

Write-Host "Source file: $sourceFile" -ForegroundColor Gray
Write-Host "Compiler: $cscPath" -ForegroundColor Gray
Write-Host ""

# Compile
Write-Host "Compiling..." -ForegroundColor Yellow

$outputDll = Join-Path $scriptDir "OfficeScrubNative.dll"
$systemManagementDll = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Management.dll"
$microsoftCSharpDll = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\Microsoft.CSharp.dll"

$compileArgs = @(
    "/target:library",
    "/out:$outputDll",
    "/reference:$systemManagementDll",
    "/reference:$microsoftCSharpDll",
    "/optimize+",
    "/warn:4",
    $sourceFile
)

try {
    $process = Start-Process -FilePath $cscPath -ArgumentList $compileArgs -NoNewWindow -Wait -PassThru

    if ($process.ExitCode -eq 0) {
        Write-Host "`nBuild succeeded!" -ForegroundColor Green

        if (Test-Path $outputDll) {
            $fileInfo = Get-Item $outputDll
            Write-Host "Output: $outputDll" -ForegroundColor Green
            Write-Host "Size: $($fileInfo.Length) bytes" -ForegroundColor Gray
            Write-Host "Modified: $($fileInfo.LastWriteTime)" -ForegroundColor Gray
        }
    }
    else {
        Write-Error "Build failed with exit code: $($process.ExitCode)"
        exit $process.ExitCode
    }
}
catch {
    Write-Error "Build failed: $_"
    exit 1
}
