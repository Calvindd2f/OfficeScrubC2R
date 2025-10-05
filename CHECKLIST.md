# Pre-Publish Checklist

Use this checklist before publishing to PowerShell Gallery or pushing to GitHub.

## ✅ Code Quality

- [ ] All PowerShell files pass PSScriptAnalyzer
  ```powershell
  Invoke-ScriptAnalyzer -Path . -Recurse -Settings .\PSScriptAnalyzerSettings.psd1
  ```

- [ ] Native DLL is built and functional
  ```powershell
  .\build.ps1 -Clean
  Test-Path .\OfficeScrubNative.dll
  ```

- [ ] Module manifest validates
  ```powershell
  Test-ModuleManifest -Path .\OfficeScrubC2R.psd1
  ```

- [ ] Module imports without errors
  ```powershell
  Import-Module .\OfficeScrubC2R.psd1 -Force -Verbose
  ```

- [ ] All exported functions are accessible
  ```powershell
  Get-Command -Module OfficeScrubC2R
  ```

## ✅ Documentation

- [ ] README.md is up to date
- [ ] CHANGELOG.md has new version entry
- [ ] Version number matches in all files:
  - [ ] OfficeScrubC2R.psd1 (ModuleVersion)
  - [ ] CHANGELOG.md (Version header)
  - [ ] README.md (Badges)

- [ ] All examples in README work
- [ ] Help documentation is complete
  ```powershell
  Get-Help Invoke-OfficeScrubC2R -Full
  Get-Help about_OfficeScrubC2R
  ```

- [ ] LICENSE file is present and correct
- [ ] CONTRIBUTING.md guidelines are clear

## ✅ Required Files

- [ ] OfficeScrubC2R.psd1 (manifest)
- [ ] OfficeScrubC2R.psm1 (main module)
- [ ] OfficeScrubC2R-Utilities.psm1 (utilities)
- [ ] OfficeScrubC2R.ps1 (script)
- [ ] OfficeScrubNative.dll (compiled)
- [ ] OfficeScrubC2R-Native.cs (source)
- [ ] build.ps1 (build script)
- [ ] LICENSE
- [ ] README.md
- [ ] CHANGELOG.md
- [ ] .gitignore
- [ ] en-US/about_OfficeScrubC2R.help.txt

## ✅ Module Manifest

Verify all fields in `OfficeScrubC2R.psd1`:

- [ ] ModuleVersion is correct
- [ ] GUID is present and unique
- [ ] Author information is set
- [ ] Description is comprehensive
- [ ] PowerShellVersion is 5.1
- [ ] RequiredAssemblies includes DLL
- [ ] FunctionsToExport lists all public functions
- [ ] NestedModules includes utilities
- [ ] FileList includes all distributed files
- [ ] Tags are relevant
- [ ] ProjectUri is correct
- [ ] LicenseUri is correct
- [ ] ReleaseNotes are updated

## ✅ Testing

- [ ] Module loads on Windows PowerShell 5.1
  ```powershell
  powershell.exe -NoProfile -Command "Import-Module .\OfficeScrubC2R.psd1"
  ```

- [ ] Module loads on PowerShell 7+
  ```powershell
  pwsh -NoProfile -Command "Import-Module .\OfficeScrubC2R.psd1"
  ```

- [ ] Detection mode works
  ```powershell
  Invoke-OfficeScrubC2R -DetectOnly
  ```

- [ ] Help works for all functions
  ```powershell
  Get-Command -Module OfficeScrubC2R | ForEach-Object { Get-Help $_.Name }
  ```

- [ ] Aliases work
  ```powershell
  Remove-OfficeC2R -DetectOnly
  Uninstall-OfficeC2R -DetectOnly
  ```

## ✅ GitHub

- [ ] All changes committed
  ```powershell
  git status
  ```

- [ ] .gitignore excludes build artifacts
- [ ] No sensitive information in commits
- [ ] Branch is up to date with main
  ```powershell
  git pull origin main
  ```

- [ ] CI/CD workflow is configured
  - [ ] .github/workflows/ci.yml exists
  - [ ] Workflow runs successfully

- [ ] README badges are correct
- [ ] Repository description matches module

## ✅ Git Tag and Release

- [ ] Version tag created
  ```powershell
  git tag -a v2.19.0 -m "Release v2.19.0"
  git push origin v2.19.0
  ```

- [ ] GitHub Release created
  - [ ] Release notes from CHANGELOG
  - [ ] DLL attached as asset
  - [ ] Installation instructions included

## ✅ PowerShell Gallery

- [ ] API key is available
- [ ] Test publish to test environment (optional)
- [ ] Ready to publish
  ```powershell
  Publish-Module -Path . -NuGetApiKey $apiKey -Verbose -WhatIf
  ```

## ✅ Post-Publication

- [ ] Module appears in PowerShell Gallery
  ```powershell
  Find-Module -Name OfficeScrubC2R
  ```

- [ ] Test installation from Gallery
  ```powershell
  Install-Module -Name OfficeScrubC2R -Scope CurrentUser
  Import-Module OfficeScrubC2R
  Get-Command -Module OfficeScrubC2R
  ```

- [ ] Repository README updated with Gallery badge
- [ ] Social media announcement (optional)
- [ ] Community notification (optional)

## ✅ Final Verification

Run this complete verification script:

```powershell
# Complete Pre-Publish Verification
Write-Host "=== OfficeScrubC2R Pre-Publish Verification ===" -ForegroundColor Cyan

# 1. Build
Write-Host "`n1. Building DLL..." -ForegroundColor Yellow
.\build.ps1 -Clean
if (-not (Test-Path .\OfficeScrubNative.dll)) {
    throw "DLL build failed"
}
Write-Host "   ✓ DLL built successfully" -ForegroundColor Green

# 2. Validate Manifest
Write-Host "`n2. Validating manifest..." -ForegroundColor Yellow
$manifest = Test-ModuleManifest -Path .\OfficeScrubC2R.psd1
Write-Host "   ✓ Manifest valid. Version: $($manifest.Version)" -ForegroundColor Green

# 3. Check Required Files
Write-Host "`n3. Checking required files..." -ForegroundColor Yellow
$requiredFiles = @(
    'OfficeScrubC2R.psd1', 'OfficeScrubC2R.psm1', 'OfficeScrubC2R-Utilities.psm1',
    'OfficeScrubNative.dll', 'LICENSE', 'README.md', 'CHANGELOG.md'
)
$missing = @()
foreach ($file in $requiredFiles) {
    if (-not (Test-Path $file)) {
        $missing += $file
        Write-Host "   ✗ Missing: $file" -ForegroundColor Red
    }
}
if ($missing.Count -eq 0) {
    Write-Host "   ✓ All required files present" -ForegroundColor Green
} else {
    throw "Missing files: $($missing -join ', ')"
}

# 4. Import Module
Write-Host "`n4. Importing module..." -ForegroundColor Yellow
Import-Module .\OfficeScrubC2R.psd1 -Force
Write-Host "   ✓ Module imported successfully" -ForegroundColor Green

# 5. Check Exported Functions
Write-Host "`n5. Checking exported functions..." -ForegroundColor Yellow
$functions = Get-Command -Module OfficeScrubC2R
Write-Host "   ✓ Exported $($functions.Count) commands" -ForegroundColor Green
$functions.Name | ForEach-Object { Write-Host "     - $_" -ForegroundColor Gray }

# 6. PSScriptAnalyzer
Write-Host "`n6. Running PSScriptAnalyzer..." -ForegroundColor Yellow
$issues = Invoke-ScriptAnalyzer -Path . -Recurse -Settings .\PSScriptAnalyzerSettings.psd1
if ($issues) {
    Write-Host "   ✗ Found $($issues.Count) issues" -ForegroundColor Red
    $issues | Format-Table -AutoSize
    throw "PSScriptAnalyzer found issues"
} else {
    Write-Host "   ✓ No issues found" -ForegroundColor Green
}

# 7. Help Check
Write-Host "`n7. Checking help documentation..." -ForegroundColor Yellow
$help = Get-Help Invoke-OfficeScrubC2R -Full
if ($help.Synopsis) {
    Write-Host "   ✓ Help documentation present" -ForegroundColor Green
} else {
    Write-Host "   ✗ Help documentation missing" -ForegroundColor Red
}

Write-Host "`n=== Verification Complete ===" -ForegroundColor Cyan
Write-Host "Module is ready for publication!" -ForegroundColor Green
```

## Quick Publish

Once all checks pass:

```powershell
# 1. Final test
Test-ModuleManifest .\OfficeScrubC2R.psd1

# 2. Commit and tag
git add .
git commit -m "Release v2.19.0"
git tag -a v2.19.0 -m "Release v2.19.0"
git push origin main --tags

# 3. Publish to PowerShell Gallery
$apiKey = Read-Host "Enter PowerShell Gallery API key" -AsSecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($apiKey)
$plainKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
Publish-Module -Path . -NuGetApiKey $plainKey -Verbose

# 4. Verify
Find-Module -Name OfficeScrubC2R
```

## 🎉 Ready to Publish!

When all checkboxes are marked, your module is ready for:
- ✅ GitHub publication
- ✅ PowerShell Gallery submission
- ✅ Production use

Good luck! 🚀

