# Rebuild script that handles locked DLL
# This script should be run in a NEW PowerShell session

Write-Host "OfficeScrubC2R Rebuild Helper" -ForegroundColor Cyan
Write-Host "============================`n" -ForegroundColor Cyan

# Check if DLL is locked
$dllPath = ".\OfficeScrubNative.dll"
if (Test-Path $dllPath) {
    try {
        $file = [System.IO.File]::Open($dllPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $file.Close()
        Write-Host "DLL is not locked - proceeding with rebuild" -ForegroundColor Green
    }
    catch {
        Write-Host "DLL is currently locked (module is loaded)" -ForegroundColor Yellow
        Write-Host "Please close all PowerShell sessions and run this script in a new session." -ForegroundColor Yellow
        Write-Host "`nAlternatively, you can:" -ForegroundColor Cyan
        Write-Host "1. Close this PowerShell window" -ForegroundColor Gray
        Write-Host "2. Open a new PowerShell window" -ForegroundColor Gray
        Write-Host "3. Navigate to: $PSScriptRoot" -ForegroundColor Gray
        Write-Host "4. Run: .\build.ps1 -Clean" -ForegroundColor Gray
        exit 1
    }
}

# Build
Write-Host "`nStarting build..." -ForegroundColor Yellow
& .\build.ps1 -Clean

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nBuild completed successfully!" -ForegroundColor Green
    Write-Host "You can now import the module: Import-Module .\OfficeScrubC2R.psd1" -ForegroundColor Cyan
}
else {
    Write-Host "`nBuild failed. Check errors above." -ForegroundColor Red
    exit $LASTEXITCODE
}

