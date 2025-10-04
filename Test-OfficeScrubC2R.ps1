# Test-OfficeScrubC2R.ps1
# Comprehensive test script for Office C2R removal functionality

[CmdletBinding()]
param(
    [switch]$TestDetection,
    [switch]$TestRegistry,
    [switch]$TestFileOperations,
    [switch]$TestNativeCode,
    [switch]$TestIntegration,
    [switch]$TestAll,
    [switch]$InstallOfficeViaWinget,
    [switch]$FullCycle
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

function Test-OfficeInstallation {
    Write-Host "`nChecking for Office Installation..." -ForegroundColor Green
    
    # Check via Winget
    try {
        $wingetOffice = winget list --id Microsoft.Office --exact 2>$null
        if ($LASTEXITCODE -eq 0 -and $wingetOffice) {
            Write-Host "  Office detected via Winget" -ForegroundColor Green
            return $true
        }
    }
    catch {
        Write-Host "  Winget check failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    # Check via registry
    $products = Get-InstalledOfficeProducts
    if ($products.Count -gt 0) {
        Write-Host "  Office detected via registry ($($products.Count) products)" -ForegroundColor Green
        return $true
    }
    
    # Check common installation paths
    $commonPaths = @(
        "$env:ProgramFiles\Microsoft Office",
        "${env:ProgramFiles(x86)}\Microsoft Office",
        "$env:ProgramFiles\Microsoft Office\root\Office16",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16"
    )
    
    foreach ($path in $commonPaths) {
        if (Test-Path $path) {
            Write-Host "  Office detected at: $path" -ForegroundColor Green
            return $true
        }
    }
    
    Write-Host "  Office not detected on system" -ForegroundColor Yellow
    return $false
}

function Install-OfficeViaWinget {
    Write-Host "`nInstalling Office via Winget..." -ForegroundColor Green
    
    # Check if Winget is available
    try {
        $wingetVersion = winget --version
        Write-Host "  Winget version: $wingetVersion" -ForegroundColor Cyan
    }
    catch {
        Write-Host "  ERROR: Winget is not available on this system" -ForegroundColor Red
        Write-Host "  Please install Windows App Installer from the Microsoft Store" -ForegroundColor Yellow
        return $false
    }
    
    # Search for Office
    Write-Host "`n  Searching for Office packages..." -ForegroundColor Cyan
    try {
        $searchResult = winget search "Microsoft 365" --source winget 2>&1
        Write-Host "  Available Office packages:" -ForegroundColor Yellow
        Write-Host $searchResult
    }
    catch {
        Write-Host "  Search failed: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`n  NOTE: Office installation via Winget requires:" -ForegroundColor Yellow
    Write-Host "    1. A valid Microsoft 365 subscription" -ForegroundColor White
    Write-Host "    2. Administrative privileges" -ForegroundColor White
    Write-Host "    3. Internet connectivity" -ForegroundColor White
    Write-Host "    4. User interaction for sign-in" -ForegroundColor White
    
    $confirm = Read-Host "`n  Do you want to install Microsoft 365 Apps? (Y/N)"
    if ($confirm -notmatch "^[Yy]") {
        Write-Host "  Installation cancelled by user" -ForegroundColor Yellow
        return $false
    }
    
    # Install Office (Microsoft 365 Apps)
    Write-Host "`n  Installing Microsoft 365 Apps..." -ForegroundColor Green
    Write-Host "  This may take 10-30 minutes depending on your connection..." -ForegroundColor Cyan
    
    try {
        $installResult = winget install --id Microsoft.Office --exact --accept-source-agreements --accept-package-agreements 2>&1
        Write-Host $installResult
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "`n  Office installation completed successfully!" -ForegroundColor Green
            Write-Host "  You may need to restart your computer" -ForegroundColor Yellow
            return $true
        }
        else {
            Write-Host "`n  Installation returned exit code: $LASTEXITCODE" -ForegroundColor Yellow
            return $false
        }
    }
    catch {
        Write-Host "`n  Installation failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Test-NativeCode {
    Write-Host "`nTesting Native C# Code Integration..." -ForegroundColor Green
    
    try {
        # Initialize environment to load C# types
        Initialize-Environment
        
        # Test Orchestrator initialization
        Write-Host "  Testing Orchestrator initialization..." -ForegroundColor Cyan
        $is64Bit = [Environment]::Is64BitOperatingSystem
        $orchestrator = New-Object OfficeScrubNative.OfficeScrubOrchestrator($is64Bit)
        
        if ($orchestrator) {
            Write-Host "    Orchestrator created successfully" -ForegroundColor Green
        }
        else {
            Write-Host "    Failed to create Orchestrator" -ForegroundColor Red
            return $false
        }
        
        # Test GUID Helper
        Write-Host "`n  Testing GuidHelper..." -ForegroundColor Cyan
        $testCompressed = "00004159110000000000000000F01FEC"
        $expanded = [OfficeScrubNative.GuidHelper]::GetExpandedGuid($testCompressed)
        Write-Host "    Compressed: $testCompressed" -ForegroundColor Gray
        Write-Host "    Expanded: $expanded" -ForegroundColor Yellow
        
        if ($expanded -and $expanded.Length -eq 38) {
            Write-Host "    GUID expansion: PASS" -ForegroundColor Green
        }
        else {
            Write-Host "    GUID expansion: FAIL" -ForegroundColor Red
        }
        
        # Test reverse operation
        $recompressed = [OfficeScrubNative.GuidHelper]::GetCompressedGuid($expanded)
        Write-Host "    Recompressed: $recompressed" -ForegroundColor Yellow
        
        if ($recompressed -eq $testCompressed) {
            Write-Host "    GUID round-trip: PASS" -ForegroundColor Green
        }
        else {
            Write-Host "    GUID round-trip: FAIL" -ForegroundColor Red
        }
        
        # Test C2R Path Detection
        Write-Host "`n  Testing C2R Path Detection..." -ForegroundColor Cyan
        $testPaths = @(
            "C:\Program Files\Microsoft Office\root\Office16",
            "C:\Program Files\Microsoft shared\ClickToRun",
            "C:\Windows\System32"
        )
        
        foreach ($path in $testPaths) {
            $isC2R = $orchestrator.IsC2RPath($path)
            $status = if ($isC2R) { "C2R" } else { "Not C2R" }
            $color = if ($isC2R) { "Yellow" } else { "Gray" }
            Write-Host "    $path -> $status" -ForegroundColor $color
        }
        
        # Test Product Scope
        Write-Host "`n  Testing Product Scope..." -ForegroundColor Cyan
        $testProducts = @(
            "{90150000-008F-0000-1000-0000000FF1CE}",  # O365 ProPlus
            "{90160000-007E-0000-0000-0000000FF1CE}",  # O365 Business
            "{91150000-0011-0000-0000-0000000FF1CE}"   # Out of scope
        )
        
        foreach ($product in $testProducts) {
            $inScope = $orchestrator.IsInScope($product)
            $status = if ($inScope) { "In Scope" } else { "Out of Scope" }
            $color = if ($inScope) { "Yellow" } else { "Gray" }
            Write-Host "    $product -> $status" -ForegroundColor $color
        }
        
        # Test ProcessHelper
        Write-Host "`n  Testing ProcessHelper..." -ForegroundColor Cyan
        $isNotepadRunning = $orchestrator.Processes.IsProcessRunning("notepad")
        Write-Host "    Notepad running: $isNotepadRunning" -ForegroundColor $(if ($isNotepadRunning) { "Yellow" } else { "Gray" })
        
        return $true
    }
    catch {
        Write-Host "  Native code test failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Stack trace: $($_.Exception.StackTrace)" -ForegroundColor DarkGray
        return $false
    }
}

function Test-Integration {
    Write-Host "`nTesting Integration Scenarios..." -ForegroundColor Green
    
    # Test 1: End-to-End Detection Flow
    Write-Host "`n  Test 1: End-to-End Detection Flow" -ForegroundColor Cyan
    try {
        Initialize-Environment
        Get-SystemInfo
        
        Write-Host "    Environment initialized" -ForegroundColor Green
        Write-Host "    64-bit system: $script:Is64Bit" -ForegroundColor Yellow
        
        $products = Get-InstalledOfficeProducts
        Write-Host "    Detected $($products.Count) Office products" -ForegroundColor Yellow
        
        if ($products.Count -gt 0) {
            Write-Host "    Detection flow: PASS" -ForegroundColor Green
        }
        else {
            Write-Host "    Detection flow: No products found (expected if Office not installed)" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "    Detection flow: FAIL - $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Test 2: Registry Operations Integration
    Write-Host "`n  Test 2: Registry Operations Integration" -ForegroundColor Cyan
    try {
        $testKey = "SOFTWARE\OfficeScrubbTest_$(Get-Random)"
        $testValue = "TestValue"
        $testData = "TestData_$(Get-Date -Format 'yyyyMMddHHmmss')"
        
        # Create
        $setResult = Set-RegistryValue -Hive LocalMachine -SubKey $testKey -ValueName $testValue -Value $testData
        Write-Host "    Set value: $setResult" -ForegroundColor $(if ($setResult) { "Green" } else { "Red" })
        
        # Read
        $getValue = Get-RegistryValue -Hive LocalMachine -SubKey $testKey -ValueName $testValue
        $readMatch = ($getValue -eq $testData)
        Write-Host "    Read value: $readMatch" -ForegroundColor $(if ($readMatch) { "Green" } else { "Red" })
        
        # Check existence
        $exists = Test-RegistryKeyExists -Hive LocalMachine -SubKey $testKey
        Write-Host "    Key exists: $exists" -ForegroundColor $(if ($exists) { "Green" } else { "Red" })
        
        # Clean up
        Remove-RegistryKey -Hive LocalMachine -SubKey $testKey
        $stillExists = Test-RegistryKeyExists -Hive LocalMachine -SubKey $testKey
        Write-Host "    Cleanup: $(!$stillExists)" -ForegroundColor $(if (!$stillExists) { "Green" } else { "Red" })
        
        if ($setResult -and $readMatch -and $exists -and !$stillExists) {
            Write-Host "    Registry integration: PASS" -ForegroundColor Green
        }
        else {
            Write-Host "    Registry integration: FAIL" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "    Registry integration: FAIL - $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Test 3: File Operations Integration
    Write-Host "`n  Test 3: File Operations Integration" -ForegroundColor Cyan
    try {
        $testDir = Join-Path $env:TEMP "OfficeScrubbTest_$(Get-Random)"
        $testFile = Join-Path $testDir "test.txt"
        
        # Create test environment
        New-Item -ItemType Directory -Path $testDir -Force | Out-Null
        Set-Content -Path $testFile -Value "Test content"
        
        # Test file deletion via C# helper
        $fileExists = Test-Path $testFile
        Write-Host "    Test file created: $fileExists" -ForegroundColor $(if ($fileExists) { "Green" } else { "Red" })
        
        $deleteResult = Remove-FileForced -Path $testFile
        $fileDeleted = -not (Test-Path $testFile)
        Write-Host "    File deleted: $fileDeleted" -ForegroundColor $(if ($fileDeleted) { "Green" } else { "Red" })
        
        # Test folder deletion
        $folderDeleteResult = Remove-FolderRecursive -Path $testDir
        $folderDeleted = -not (Test-Path $testDir)
        Write-Host "    Folder deleted: $folderDeleted" -ForegroundColor $(if ($folderDeleted) { "Green" } else { "Red" })
        
        if ($fileExists -and $fileDeleted -and $folderDeleted) {
            Write-Host "    File operations integration: PASS" -ForegroundColor Green
        }
        else {
            Write-Host "    File operations integration: FAIL" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "    File operations integration: FAIL - $($_.Exception.Message)" -ForegroundColor Red
        # Clean up on error
        if (Test-Path $testDir) {
            Remove-Item $testDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Test-FullCycle {
    Write-Host "`n" + "="*70 -ForegroundColor Magenta
    Write-Host "FULL CYCLE TEST: Install -> Detect -> Remove" -ForegroundColor Magenta
    Write-Host "="*70 -ForegroundColor Magenta
    
    Write-Host "`nWARNING: This test will:" -ForegroundColor Yellow
    Write-Host "  1. Check if Office is installed" -ForegroundColor White
    Write-Host "  2. Optionally install Office via Winget (if not found)" -ForegroundColor White
    Write-Host "  3. Detect Office installation" -ForegroundColor White
    Write-Host "  4. Run Office scrubber in DETECT ONLY mode" -ForegroundColor White
    Write-Host "  5. Optionally run full removal (DESTRUCTIVE)" -ForegroundColor White
    
    $confirm = Read-Host "`nDo you want to proceed? (Y/N)"
    if ($confirm -notmatch "^[Yy]") {
        Write-Host "Full cycle test cancelled" -ForegroundColor Yellow
        return
    }
    
    # Phase 1: Check Installation
    Write-Host "`n--- Phase 1: Check Installation ---" -ForegroundColor Cyan
    $officeInstalled = Test-OfficeInstallation
    
    if (-not $officeInstalled) {
        Write-Host "`nOffice is not installed. Would you like to install it?" -ForegroundColor Yellow
        $installConfirm = Read-Host "Install Office via Winget? (Y/N)"
        
        if ($installConfirm -match "^[Yy]") {
            $installed = Install-OfficeViaWinget
            if (-not $installed) {
                Write-Host "`nOffice installation failed. Cannot continue full cycle test." -ForegroundColor Red
                return
            }
        }
        else {
            Write-Host "`nSkipping full cycle test - Office not installed" -ForegroundColor Yellow
            return
        }
    }
    
    # Phase 2: Detection
    Write-Host "`n--- Phase 2: Run Detection ---" -ForegroundColor Cyan
    try {
        & (Join-Path $PSScriptRoot "OfficeScrubC2R.ps1") -DetectOnly -LogPath $env:TEMP
        Write-Host "Detection completed successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "Detection failed: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Phase 3: Optional Removal
    Write-Host "`n--- Phase 3: Optional Removal ---" -ForegroundColor Cyan
    Write-Host "`nWARNING: The next step will REMOVE Office from your system!" -ForegroundColor Red
    Write-Host "This is a DESTRUCTIVE operation and cannot be undone!" -ForegroundColor Red
    
    $removeConfirm = Read-Host "`nDo you want to proceed with Office removal? (Type 'YES' to confirm)"
    
    if ($removeConfirm -eq "YES") {
        try {
            Write-Host "`nRunning Office Scrubber..." -ForegroundColor Yellow
            & (Join-Path $PSScriptRoot "OfficeScrubC2R.ps1") -Quiet -Force -LogPath $env:TEMP
            Write-Host "Office removal completed" -ForegroundColor Green
        }
        catch {
            Write-Host "Office removal failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Removal cancelled - Office remains installed" -ForegroundColor Yellow
    }
}

function Show-TestSummary {
    Write-Host "`n" + "="*70 -ForegroundColor Green
    Write-Host "TEST SUITE FOR OFFICE SCRUB C2R" -ForegroundColor Green
    Write-Host "="*70 -ForegroundColor Green

    Write-Host "`nAvailable Test Functions:" -ForegroundColor Yellow
    Write-Host "  -TestDetection      : Test Office detection functions" -ForegroundColor White
    Write-Host "  -TestRegistry       : Test registry operations" -ForegroundColor White
    Write-Host "  -TestFileOperations : Test file operations" -ForegroundColor White
    Write-Host "  -TestNativeCode     : Test C# native code integration" -ForegroundColor White
    Write-Host "  -TestIntegration    : Test integration scenarios" -ForegroundColor White
    Write-Host "  -TestAll            : Run all tests" -ForegroundColor White
    Write-Host "  -InstallOfficeViaWinget : Install Office using Winget" -ForegroundColor Cyan
    Write-Host "  -FullCycle          : Full cycle test (Install -> Detect -> Remove)" -ForegroundColor Magenta

    Write-Host "`nUsage Examples:" -ForegroundColor Yellow
    Write-Host "  .\Test-OfficeScrubC2R.ps1 -TestDetection" -ForegroundColor White
    Write-Host "  .\Test-OfficeScrubC2R.ps1 -TestAll" -ForegroundColor White
    Write-Host "  .\Test-OfficeScrubC2R.ps1 -TestNativeCode -Verbose" -ForegroundColor White
    Write-Host "  .\Test-OfficeScrubC2R.ps1 -InstallOfficeViaWinget" -ForegroundColor Cyan
    Write-Host "  .\Test-OfficeScrubC2R.ps1 -FullCycle" -ForegroundColor Magenta
    
    Write-Host "`nNotes:" -ForegroundColor Yellow
    Write-Host "  - Most tests require Administrator privileges" -ForegroundColor White
    Write-Host "  - Office installation via Winget requires internet connectivity" -ForegroundColor White
    Write-Host "  - Full cycle test is destructive and should be used with caution" -ForegroundColor White
}

# Main test execution
Write-Host "`n" + "="*70 -ForegroundColor Green
Write-Host "OFFICE SCRUB C2R - TEST SUITE" -ForegroundColor Green
Write-Host "="*70 -ForegroundColor Green

# Check for admin privileges
$isAdmin = Test-IsElevated
if (-not $isAdmin) {
    Write-Host "`nWARNING: Not running as Administrator!" -ForegroundColor Yellow
    Write-Host "Some tests may fail without elevated privileges.`n" -ForegroundColor Yellow
}

# Handle special modes first
if ($InstallOfficeViaWinget) {
    Install-OfficeViaWinget
    exit
}

if ($FullCycle) {
    Test-FullCycle
    exit
}

# Run standard tests
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

if ($TestAll -or $TestNativeCode) {
    Test-NativeCode
}

if ($TestAll -or $TestIntegration) {
    Test-Integration
}

if (-not ($TestAll -or $TestDetection -or $TestRegistry -or $TestFileOperations -or $TestNativeCode -or $TestIntegration)) {
    Show-TestSummary
}
else {
    Write-Host "`n" + "="*70 -ForegroundColor Green
    Write-Host "All tests completed!" -ForegroundColor Green
    Write-Host "="*70 -ForegroundColor Green
}
