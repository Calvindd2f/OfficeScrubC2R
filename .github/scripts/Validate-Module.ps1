# Quick module validation script
Write-Host "`n=== OfficeScrubC2R Module Validation ===" -ForegroundColor Cyan

$results = @{}

# Test manifest
try {
    $manifest = Test-ModuleManifest -Path .\OfficeScrubC2R.psd1 -ErrorAction Stop
    $results['Manifest'] = "PASS (v$($manifest.Version))"
}
catch {
    $results['Manifest'] = "FAIL: $_"
}

# Test DLL
if (Test-Path .\OfficeScrubNative.dll) {
    $results['DLL'] = "PASS"
}
else {
    $results['DLL'] = "FAIL: Not found"
}

# Test module import
try {
    Import-Module .\OfficeScrubC2R.psd1 -Force -ErrorAction Stop
    $results['Import'] = "PASS"
    
    # Test commands
    $commands = Get-Command -Module OfficeScrubC2R
    $results['Commands'] = "PASS ($($commands.Count) exported)"
    
    # Test aliases
    $aliases = Get-Alias | Where-Object { $_.ModuleName -eq 'OfficeScrubC2R' }
    $results['Aliases'] = "PASS ($($aliases.Count) aliases)"
}
catch {
    $results['Import'] = "FAIL: $_"
    $results['Commands'] = "SKIP"
    $results['Aliases'] = "SKIP"
}

# Test required files
$requiredFiles = @('LICENSE', 'README.md', 'CHANGELOG.md', 'CONTRIBUTING.md')
$missingFiles = @()
foreach ($file in $requiredFiles) {
    if (-not (Test-Path $file)) {
        $missingFiles += $file
    }
}
if ($missingFiles.Count -eq 0) {
    $results['Required Files'] = "PASS"
}
else {
    $results['Required Files'] = "FAIL: Missing $($missingFiles -join ', ')"
}

# Test help
try {
    $help = Get-Help Invoke-OfficeScrubC2R -ErrorAction Stop
    if ($help.Synopsis) {
        $results['Help'] = "PASS"
    }
    else {
        $results['Help'] = "FAIL: No synopsis"
    }
}
catch {
    $results['Help'] = "FAIL: $_"
}

# Display results
Write-Host ""
$maxKeyLength = ($results.Keys | Measure-Object -Property Length -Maximum).Maximum
foreach ($key in $results.Keys | Sort-Object) {
    $value = $results[$key]
    $color = if ($value -like "PASS*") { 'Green' } elseif ($value -like "FAIL*") { 'Red' } else { 'Yellow' }
    Write-Host ("{0,-$($maxKeyLength + 2)}: {1}" -f $key, $value) -ForegroundColor $color
}

# Summary
$passed = ($results.Values | Where-Object { $_ -like "PASS*" }).Count
$total = $results.Count
Write-Host "`n=== Summary: $passed/$total checks passed ===" -ForegroundColor $(if ($passed -eq $total) { 'Green' } else { 'Yellow' })

if ($passed -eq $total) {
    Write-Host "`nModule is ready for publication!" -ForegroundColor Green
}
else {
    Write-Host "`nSome checks failed. Review above." -ForegroundColor Yellow
}

