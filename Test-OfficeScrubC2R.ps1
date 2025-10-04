# Test-OfficeScrubC2R.ps1
# Test script for Office C2R removal functionality

[CmdletBinding()]
param(
    [switch]$TestDetection,
    [switch]$TestRegistry,
    [switch]$TestFileOperations,
    [switch]$TestAll
)

# Import the utility module
Import-Module -Name (Join-Path $PSScriptRoot "OfficeScrubC2R-Utilities.psm1") -Force

function Test-OfficeDetection {
    Write-Host "Testing Office Detection Functions..." -ForegroundColor Green

    # Test C2R detection
    $testValues = @(
        "officeclicktorun",
        "o365proplus",
        "o365business",
        "microsoft office 2016",
        "notoffice"
    )

    foreach ($value in $testValues) {
        $isC2R = Test-IsC2R $value
        Write-Host "  $value -> IsC2R: $isC2R" -ForegroundColor $(if ($isC2R) { "Yellow" } else { "Gray" })
    }

    # Test registry reading
    Write-Host "`nTesting Registry Operations..." -ForegroundColor Green

    $testKeys = @(
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    )

    foreach ($key in $testKeys) {
        $exists = Test-RegistryKeyExists -Hive LocalMachine -SubKey $key
        Write-Host "  $key -> Exists: $exists" -ForegroundColor $(if ($exists) { "Green" } else { "Red" })

        if ($exists) {
            $subKeys = Get-RegistrySubKeys -Hive LocalMachine -SubKey $key
            Write-Host "    SubKeys: $($subKeys.Count)" -ForegroundColor Cyan
        }
    }
}

function Test-RegistryOperations {
    Write-Host "`nTesting Registry Operations..." -ForegroundColor Green

    # Test registry value operations
    $testKey = "SOFTWARE\Microsoft\Office\Test"
    $testValue = "TestValue"
    $testData = "TestData"

    try {
        # Test setting a value
        $result = Set-RegistryValue -Hive LocalMachine -SubKey $testKey -ValueName $testValue -Value $testData
        Write-Host "  Set registry value: $result" -ForegroundColor $(if ($result) { "Green" } else { "Red" })

        # Test reading the value
        $readValue = Get-RegistryValue -Hive LocalMachine -SubKey $testKey -ValueName $testValue
        Write-Host "  Read registry value: $readValue" -ForegroundColor $(if ($readValue -eq $testData) { "Green" } else { "Red" })

        # Test removing the value
        $removeResult = Remove-RegistryValue -Hive LocalMachine -SubKey $testKey -ValueName $testValue
        Write-Host "  Remove registry value: $removeResult" -ForegroundColor $(if ($removeResult) { "Green" } else { "Red" })

        # Clean up test key
        Remove-RegistryKey -Hive LocalMachine -SubKey $testKey
        Write-Host "  Cleaned up test key" -ForegroundColor Green

    }
    catch {
        Write-Host "  Registry test failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Test-FileOperations {
    Write-Host "`nTesting File Operations..." -ForegroundColor Green

    $testDir = Join-Path $env:TEMP "OfficeScrubTest"
    $testFile = Join-Path $testDir "testfile.txt"

    try {
        # Create test directory and file
        New-Item -ItemType Directory -Path $testDir -Force | Out-Null
        Set-Content -Path $testFile -Value "Test content"
        Write-Host "  Created test file: $testFile" -ForegroundColor Green

        # Test file deletion
        $deleteResult = Remove-FileFast $testFile
        Write-Host "  Delete file: $deleteResult" -ForegroundColor $(if ($deleteResult) { "Green" } else { "Red" })

        # Test directory deletion
        $dirDeleteResult = Remove-FolderFast $testDir
        Write-Host "  Delete directory: $dirDeleteResult" -ForegroundColor $(if ($dirDeleteResult) { "Green" } else { "Red" })

    }
    catch {
        Write-Host "  File operations test failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Test-ProcessManagement {
    Write-Host "`nTesting Process Management..." -ForegroundColor Green

    # Test process detection
    $testProcesses = @("notepad", "calc")
    foreach ($processName in $testProcesses) {
        $isRunning = Test-ProcessRunning $processName
        Write-Host "  $processName running: $isRunning" -ForegroundColor $(if ($isRunning) { "Yellow" } else { "Gray" })
    }
}

function Test-SystemInfo {
    Write-Host "`nTesting System Information..." -ForegroundColor Green

    # Test elevation check
    $isElevated = Test-IsElevated
    Write-Host "  Is Elevated: $isElevated" -ForegroundColor $(if ($isElevated) { "Green" } else { "Yellow" })

    # Test system info
    Get-SystemInfo
    Write-Host "  OS Version: $script:OSVersion" -ForegroundColor Cyan
    Write-Host "  Is 64-bit: $script:Is64Bit" -ForegroundColor Cyan
    Write-Host "  OS Info: $script:OSInfo" -ForegroundColor Cyan
}

function Test-OfficeProductDetection {
    Write-Host "`nTesting Office Product Detection..." -ForegroundColor Green

    try {
        $products = Get-InstalledOfficeProducts
        Write-Host "  Found $($products.Count) Office products" -ForegroundColor Cyan

        foreach ($product in $products.GetEnumerator()) {
            Write-Host "    $($product.Key): $($product.Value)" -ForegroundColor Yellow
        }

        Write-Host "  Found $($script:C2RSuite.Count) C2R suite entries" -ForegroundColor Cyan
        foreach ($suite in $script:C2RSuite.GetEnumerator()) {
            Write-Host "    $($suite.Key): $($suite.Value)" -ForegroundColor Yellow
        }

    }
    catch {
        Write-Host "  Office product detection failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Test-GUIDOperations {
    Write-Host "`nTesting GUID Operations..." -ForegroundColor Green

    $testGuids = @(
        "12345678901234567890123456789012",
        "{12345678-1234-1234-1234-123456789012}",
        "12345678-1234-1234-1234-123456789012"
    )

    foreach ($guid in $testGuids) {
        $expanded = Get-ExpandedGuid $guid
        $compressed = Get-CompressedGuid $guid
        Write-Host "  Original: $guid" -ForegroundColor Gray
        Write-Host "    Expanded: $expanded" -ForegroundColor Cyan
        Write-Host "    Compressed: $compressed" -ForegroundColor Cyan
    }
}

function Test-ErrorHandling {
    Write-Host "`nTesting Error Handling..." -ForegroundColor Green

    # Test error code setting
    Set-ErrorCode $script:ERROR_REBOOT_REQUIRED
    Write-Host "  Set ERROR_REBOOT_REQUIRED: $script:ErrorCode" -ForegroundColor Yellow

    Clear-ErrorCode $script:ERROR_REBOOT_REQUIRED
    Write-Host "  Cleared ERROR_REBOOT_REQUIRED: $script:ErrorCode" -ForegroundColor Green

    # Test return value setting
    Set-ReturnValue $script:ERROR_SUCCESS
    Write-Host "  Set return value to SUCCESS" -ForegroundColor Green
}

function Test-Logging {
    Write-Host "`nTesting Logging Functions..." -ForegroundColor Green

    $testLogPath = Join-Path $env:TEMP "OfficeScrubTest.log"
    Initialize-Log $testLogPath

    Write-LogHeader "Test Header"
    Write-LogSubHeader "Test Sub Header"
    Write-Log "Test log message"
    Write-LogOnly "Test log only message"

    Write-Host "  Log file created at: $testLogPath" -ForegroundColor Cyan

    if (Test-Path $testLogPath) {
        $logContent = Get-Content $testLogPath
        Write-Host "  Log entries: $($logContent.Count)" -ForegroundColor Green
        Remove-Item $testLogPath -Force
    }
}

function Show-TestSummary {
    Write-Host "`n" + "="*50 -ForegroundColor Green
    Write-Host "TEST SUMMARY" -ForegroundColor Green
    Write-Host "="*50 -ForegroundColor Green

    Write-Host "`nAvailable Test Functions:" -ForegroundColor Yellow
    Write-Host "  -TestDetection    : Test Office detection functions" -ForegroundColor White
    Write-Host "  -TestRegistry     : Test registry operations" -ForegroundColor White
    Write-Host "  -TestFileOperations: Test file operations" -ForegroundColor White
    Write-Host "  -TestAll          : Run all tests" -ForegroundColor White

    Write-Host "`nUsage Examples:" -ForegroundColor Yellow
    Write-Host "  .\Test-OfficeScrubC2R.ps1 -TestDetection" -ForegroundColor White
    Write-Host "  .\Test-OfficeScrubC2R.ps1 -TestAll" -ForegroundColor White
    Write-Host "  .\Test-OfficeScrubC2R.ps1 -TestRegistry -Verbose" -ForegroundColor White
}

# Main test execution
if ($TestAll -or $TestDetection) {
    Test-OfficeDetection
    Test-OfficeProductDetection
    Test-SystemInfo
    Test-GUIDOperations
    Test-ErrorHandling
    Test-Logging
}

if ($TestAll -or $TestRegistry) {
    Test-RegistryOperations
}

if ($TestAll -or $TestFileOperations) {
    Test-FileOperations
    Test-ProcessManagement
}

if (-not ($TestAll -or $TestDetection -or $TestRegistry -or $TestFileOperations)) {
    Show-TestSummary
}

Write-Host "`nTest completed!" -ForegroundColor Green
