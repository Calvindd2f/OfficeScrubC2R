# OfficeScrubC2R-Utilities.psm1
# PowerShell utilities with C# inline code for performance

#region Script Variables

$script:SCRIPT_VERSION = "2.19"
$script:SCRIPT_NAME = "OfficeScrubC2R"

# Error codes (matching VBS)
$script:ERROR_SUCCESS = 0
$script:ERROR_FAIL = 1
$script:ERROR_REBOOT_REQUIRED = 2
$script:ERROR_USERCANCEL = 4
$script:ERROR_STAGE1 = 8
$script:ERROR_STAGE2 = 16
$script:ERROR_INCOMPLETE = 32
$script:ERROR_DCAF_FAILURE = 64
$script:ERROR_ELEVATION_USERDECLINED = 128
$script:ERROR_ELEVATION = 256
$script:ERROR_SCRIPTINIT = 512
$script:ERROR_RELAUNCH = 1024
$script:ERROR_UNKNOWN = 2048

# Registry constants
$script:REG_ARP = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"

# Global state
$script:ErrorCode = 0
$script:RebootRequired = $false
$script:IsElevated = $false
$script:Is64Bit = $false
$script:OSInfo = ""
$script:LogStream = $null
$script:LogDir = ""
$script:ScrubDir = ""

# Dictionaries
$script:InstalledSku = @{}
$script:C2RSuite = @{}
$script:KeepSku = @{}
$script:KeepFolder = @{}
$script:DelInUse = @{}

# Environment paths
$script:ProgramFiles = ""
$script:ProgramFilesX86 = ""
$script:CommonProgramFiles = ""
$script:CommonProgramFilesX86 = ""
$script:ProgramData = ""
$script:AppData = ""
$script:LocalAppData = ""
$script:AllUsersProfile = ""
$script:WinDir = ""
$script:Temp = ""
$script:WICacheDir = ""

# Native orchestrator
$script:Orchestrator = $null

#endregion

#region C# Type Loading

function Initialize-NativeTypes {
    $csharpPath = Join-Path $PSScriptRoot "OfficeScrubC2R-Native.cs"
    if (-not (Test-Path $csharpPath)) {
        throw "C# helper file not found: $csharpPath"
    }

    $csharpCode = Get-Content $csharpPath -Raw

    try {
        # For Windows PowerShell 5.1 (.NET Framework), use simpler assembly references
        # For PowerShell 7+ (.NET Core), we would need more specific assemblies
        if ($PSVersionTable.PSVersion.Major -ge 6) {
            # PowerShell 7+ (.NET Core)
            $assemblies = @(
                "System",
                "System.Core",
                "System.Collections",
                "System.Linq",
                "System.Management",
                "System.Threading.Thread",
                "System.ComponentModel.Primitives",
                "System.Diagnostics.Process",
                "System.IO.FileSystem",
                "System.Runtime",
                "Microsoft.CSharp",
                "Microsoft.Win32.Registry",
                "mscorlib",
                "netstandard"
            )
        }
        else {
            # Windows PowerShell 5.1 (.NET Framework) - much simpler!
            $assemblies = @(
                "System",
                "System.Core",
                "System.Management",
                "Microsoft.CSharp"
            )
        }
        
        Add-Type -TypeDefinition $csharpCode -Language CSharp `
            -ReferencedAssemblies $assemblies -ErrorAction Stop

        Write-Verbose "Native C# types loaded successfully"
    }
    catch {
        if ($_.Exception.Message -notlike "*already exists*") {
            throw "Failed to load C# types: $_"
        }
    }
}

#endregion

#region Environment Initialization

function Initialize-Environment {
    [CmdletBinding()]
    param()

    # Load C# types
    Initialize-NativeTypes

    # Initialize orchestrator
    $script:Is64Bit = [Environment]::Is64BitOperatingSystem
    $script:Orchestrator = New-Object OfficeScrub.Native.OfficeScrubOrchestrator($script:Is64Bit)

    # Set environment paths
    $script:ProgramFiles = [Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles)
    $script:CommonProgramFiles = [Environment]::GetFolderPath([Environment+SpecialFolder]::CommonProgramFiles)
    $script:ProgramData = [Environment]::GetFolderPath([Environment+SpecialFolder]::CommonApplicationData)
    $script:AppData = [Environment]::GetFolderPath([Environment+SpecialFolder]::ApplicationData)
    $script:LocalAppData = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
    $script:AllUsersProfile = [Environment]::GetFolderPath([Environment+SpecialFolder]::CommonDesktopDirectory) | Split-Path -Parent
    $script:WinDir = $env:windir
    $script:Temp = [System.IO.Path]::GetTempPath()

    if ($script:Is64Bit) {
        $script:ProgramFilesX86 = ${env:ProgramFiles(x86)}
        $script:CommonProgramFilesX86 = ${env:CommonProgramFiles(x86)}
    }
    else {
        $script:ProgramFilesX86 = $script:ProgramFiles
        $script:CommonProgramFilesX86 = $script:CommonProgramFiles
    }

    $script:WICacheDir = Join-Path $script:WinDir "Installer"
    $script:ScrubDir = Join-Path $script:Temp $script:SCRIPT_NAME

    if (-not (Test-Path $script:ScrubDir)) {
        New-Item -Path $script:ScrubDir -ItemType Directory -Force | Out-Null
    }

    $script:LogDir = $script:ScrubDir
}

function Get-SystemInfo {
    [CmdletBinding()]
    param()

    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $cs = Get-CimInstance -ClassName Win32_ComputerSystem

    $script:OSInfo = "{0}, Version: {1}, Architecture: {2}" -f `
        $os.Caption, $os.Version, $cs.SystemType

    $script:Is64Bit = $cs.SystemType -like "*64*"
}

function Test-IsElevated {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

#endregion

#region Logging Functions

function Initialize-Log {
    [CmdletBinding()]
    param(
        [string]$LogPath = $script:LogDir
    )

    $computerName = $env:COMPUTERNAME
    $timestamp = Get-Date -Format "yyyyMMddHHmmss"
    $logFile = Join-Path $LogPath ("{0}_{1}_ScrubLog.txt" -f $computerName, $timestamp)

    try {
        $script:LogStream = [System.IO.StreamWriter]::new($logFile, $false, [System.Text.Encoding]::UTF8)
        $script:LogStream.AutoFlush = $true

        Write-LogHeader "Microsoft Customer Support Services - Office C2R Removal Utility"
        Write-Log ("Version: {0}" -f $script:SCRIPT_VERSION)
        Write-Log ("64-bit OS: {0}" -f $script:Is64Bit)
        Write-Log ("Removal start: {0}" -f (Get-Date))
        Write-Log ("OS Details: {0}" -f $script:OSInfo)
        Write-Log ""
    }
    catch {
        Write-Warning "Failed to initialize log: $_"
    }
}

function Write-LogHeader {
    [CmdletBinding()]
    param([string]$Message)

    $separator = "=" * $Message.Length
    $output = "`r`n{0}`r`n{1}" -f $Message, $separator

    if (-not $script:Quiet) {
        Write-Host $output -ForegroundColor Cyan
    }

    if ($script:LogStream) {
        $script:LogStream.WriteLine("")
        $script:LogStream.WriteLine($output)
    }
}

function Write-LogSubHeader {
    [CmdletBinding()]
    param([string]$Message)

    $separator = "-" * $Message.Length
    $output = "`r`n{0}`r`n{1}" -f $Message, $separator

    if (-not $script:Quiet) {
        Write-Host $output -ForegroundColor Yellow
    }

    if ($script:LogStream) {
        $script:LogStream.WriteLine("")
        $script:LogStream.WriteLine($output)
    }
}

function Write-Log {
    [CmdletBinding()]
    param([string]$Message)

    $timestamp = Get-Date -Format "HH:mm:ss"
    $output = "   {0}: {1}" -f $timestamp, $Message

    if (-not $script:Quiet) {
        Write-Host $output
    }

    if ($script:LogStream) {
        $script:LogStream.WriteLine($output)
    }
}

function Write-LogOnly {
    [CmdletBinding()]
    param([string]$Message)

    if ($script:LogStream) {
        $timestamp = Get-Date -Format "HH:mm:ss"
        $script:LogStream.WriteLine("   {0}: {1}" -f $timestamp, $Message)
    }
}

function Close-Log {
    if ($script:LogStream) {
        $script:LogStream.Flush()
        $script:LogStream.Close()
        $script:LogStream.Dispose()
        $script:LogStream = $null
    }
}

#endregion

#region Error Handling

function Set-ErrorCode {
    [CmdletBinding()]
    param([int]$ErrorBit)

    $script:ErrorCode = $script:ErrorCode -bor $ErrorBit

    # Cascade critical errors to FAIL bit
    $criticalErrors = $script:ERROR_DCAF_FAILURE -bor $script:ERROR_STAGE2 -bor
    $script:ERROR_ELEVATION_USERDECLINED -bor $script:ERROR_ELEVATION -bor
    $script:ERROR_SCRIPTINIT

    if ($script:ErrorCode -band $criticalErrors) {
        $script:ErrorCode = $script:ErrorCode -bor $script:ERROR_FAIL
    }
}

function Clear-ErrorCode {
    [CmdletBinding()]
    param([int]$ErrorBit)

    $script:ErrorCode = $script:ErrorCode -band (-bnot $ErrorBit)

    # Clear FAIL bit if clearing critical errors
    $clearableErrors = $script:ERROR_ELEVATION_USERDECLINED -bor $script:ERROR_ELEVATION -bor
    $script:ERROR_SCRIPTINIT

    if ($ErrorBit -band $clearableErrors) {
        $script:ErrorCode = $script:ErrorCode -band (-bnot $script:ERROR_FAIL)
    }
}

function Set-ReturnValue {
    [CmdletBinding()]
    param([int]$Value)

    $retValFile = Join-Path $script:ScrubDir "ScrubRetValFile.txt"

    try {
        [System.IO.File]::WriteAllText($retValFile, $Value.ToString())
    }
    catch {
        Write-LogOnly "Failed to write return value file: $_"
    }
}

#endregion

#region Registry Operations (using C# helpers)

function Get-RegistryValue {
    [CmdletBinding()]
    param(
        [OfficeScrub.Native.RegistryHiveType]$Hive,
        [string]$SubKey,
        [string]$ValueName,
        [object]$DefaultValue = $null
    )

    return $script:Orchestrator.Registry.GetValue($Hive, $SubKey, $ValueName, $DefaultValue)
}

function Set-RegistryValue {
    [CmdletBinding()]
    param(
        [OfficeScrub.Native.RegistryHiveType]$Hive,
        [string]$SubKey,
        [string]$ValueName,
        [object]$Value,
        [Microsoft.Win32.RegistryValueKind]$Kind = [Microsoft.Win32.RegistryValueKind]::String
    )

    return $script:Orchestrator.Registry.SetValue($Hive, $SubKey, $ValueName, $Value, $Kind)
}

function Remove-RegistryKey {
    [CmdletBinding()]
    param(
        [OfficeScrub.Native.RegistryHiveType]$Hive,
        [string]$SubKey,
        [switch]$Recursive = $true
    )

    if (-not $script:DetectOnly) {
        Write-LogOnly "Delete registry key: $($Hive)\$SubKey"
        return $script:Orchestrator.Registry.DeleteKey($Hive, $SubKey, $Recursive)
    }
    else {
        Write-LogOnly "Preview mode. Would delete registry key: $($Hive)\$SubKey"
        return $false
    }
}

function Remove-RegistryValue {
    [CmdletBinding()]
    param(
        [OfficeScrub.Native.RegistryHiveType]$Hive,
        [string]$SubKey,
        [string]$ValueName
    )

    if (-not $script:DetectOnly) {
        Write-LogOnly "Delete registry value: $($Hive)\$SubKey -> $ValueName"
        return $script:Orchestrator.Registry.DeleteValue($Hive, $SubKey, $ValueName)
    }
    else {
        Write-LogOnly "Preview mode. Would delete registry value: $($Hive)\$SubKey -> $ValueName"
        return $false
    }
}

function Get-RegistryKeys {
    [CmdletBinding()]
    param(
        [OfficeScrub.Native.RegistryHiveType]$Hive,
        [string]$SubKey
    )

    return $script:Orchestrator.Registry.EnumerateKeys($Hive, $SubKey)
}

function Get-RegistryValues {
    [CmdletBinding()]
    param(
        [OfficeScrub.Native.RegistryHiveType]$Hive,
        [string]$SubKey
    )

    return $script:Orchestrator.Registry.EnumerateValues($Hive, $SubKey)
}

function Test-RegistryKeyExists {
    [CmdletBinding()]
    param(
        [OfficeScrub.Native.RegistryHiveType]$Hive,
        [string]$SubKey
    )

    return $script:Orchestrator.Registry.KeyExists($Hive, $SubKey)
}

#endregion

#region File Operations (using C# helpers)

function Remove-FolderRecursive {
    [CmdletBinding()]
    param(
        [string]$Path,
        [switch]$Force
    )

    if (-not (Test-Path $Path)) {
        return $true
    }

    if (-not $script:DetectOnly) {
        Write-LogOnly "Delete folder: $Path"
        $result = $script:Orchestrator.Files.DeleteDirectory($Path, $true, $true)

        if (-not $result) {
            Write-Log "Failed to delete folder, scheduled for reboot: $Path"
            $script:RebootRequired = $true
            Set-ErrorCode $script:ERROR_REBOOT_REQUIRED
        }

        return $result
    }
    else {
        Write-LogOnly "Preview mode. Would delete folder: $Path"
        return $false
    }
}

function Remove-FileForced {
    [CmdletBinding()]
    param(
        [string]$Path,
        [switch]$ScheduleOnFail
    )

    if (-not (Test-Path $Path)) {
        return $true
    }

    if (-not $script:DetectOnly) {
        Write-LogOnly "Delete file: $Path"
        $result = $script:Orchestrator.Files.DeleteFile($Path, $ScheduleOnFail)

        if (-not $result -and $ScheduleOnFail) {
            Write-Log "Failed to delete file, scheduled for reboot: $Path"
            $script:RebootRequired = $true
            Set-ErrorCode $script:ERROR_REBOOT_REQUIRED
        }

        return $result
    }
    else {
        Write-LogOnly "Preview mode. Would delete file: $Path"
        return $false
    }
}

function Add-PendingFileDelete {
    [CmdletBinding()]
    param([string]$Path)

    $script:Orchestrator.Registry.AddPendingFileRenameOperation($Path)
    $script:DelInUse[$Path] = $Path
    $script:RebootRequired = $true
    Set-ErrorCode $script:ERROR_REBOOT_REQUIRED
}

#endregion

#region Process Operations

function Stop-OfficeProcesses {
    [CmdletBinding()]
    param([switch]$Force)

    Write-LogSubHeader "Stopping Office processes"

    $processes = [OfficeScrub.Native.OfficeConstants]::OFFICE_PROCESSES
    $terminated = $script:Orchestrator.Processes.TerminateProcesses($processes, 10000)

    if ($terminated.Count -gt 0) {
        Write-Log ("Terminated {0} Office process(es)" -f $terminated.Count)
        Start-Sleep -Seconds 2
    }
}

function Test-ProcessRunning {
    [CmdletBinding()]
    param([string]$ProcessName)

    return $script:Orchestrator.Processes.IsProcessRunning($ProcessName)
}

#endregion

#region Product Detection

function Get-InstalledOfficeProducts {
    [CmdletBinding()]
    param()

    Write-LogSubHeader "Detect installed products"

    $products = @{}
    $script:C2RSuite = @{}

    # O15 Configuration
    Write-LogOnly "Check for O15 C2R products"
    $o15Products = Get-RegistryValue -Hive LocalMachine `
        -SubKey "SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration" `
        -ValueName "ProductReleaseIds"

    if ($o15Products) {
        foreach ($prod in ($o15Products -split ',')) {
            Write-LogOnly "Found O15 C2R product in Configuration: $prod"
            $version = Get-RegistryValue -Hive LocalMachine `
                -SubKey "SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs\Active\culture" `
                -ValueName "x-none"

            $products[$prod.ToLower()] = $version
            $script:C2RSuite[$prod] = "$prod - $version"
        }
    }

    # O15 PropertyBag
    $o15PropBag = Get-RegistryValue -Hive LocalMachine `
        -SubKey "SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag" `
        -ValueName "productreleaseid"

    if ($o15PropBag) {
        foreach ($prod in ($o15PropBag -split ',')) {
            Write-LogOnly "Found O15 C2R product in PropertyBag: $prod"
            if (-not $products.ContainsKey($prod.ToLower())) {
                $products[$prod.ToLower()] = "15.0"
                $script:C2RSuite[$prod] = "$prod - 15.0"
            }
        }
    }

    # Office C2R (QR8+)
    Write-LogOnly "Check for Office C2R products (>=QR8)"
    $activeConfig = Get-RegistryValue -Hive LocalMachine `
        -SubKey "SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs" `
        -ValueName "ActiveConfiguration"

    if ($activeConfig) {
        $configKeys = Get-RegistryKeys -Hive LocalMachine `
            -SubKey "SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\$activeConfig"

        foreach ($key in $configKeys) {
            if ($key -notin @("culture", "stream")) {
                $prod = $key
                if ($prod -like "*.*") {
                    $prod = $prod.Substring(0, $prod.IndexOf('.'))
                }

                Write-LogOnly "Found Office C2R product in Configuration: $prod"
                $version = Get-RegistryValue -Hive LocalMachine `
                    -SubKey "SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\$activeConfig\$key\x-none" `
                    -ValueName "Version"

                if (-not $products.ContainsKey($prod.ToLower())) {
                    $products[$prod.ToLower()] = $version
                    $script:C2RSuite[$prod] = "$prod - $version"
                }
            }
        }
    }

    # Office C2R (QR7)
    Write-LogOnly "Check for Office C2R products (QR7)"
    $qr7Products = Get-RegistryValue -Hive LocalMachine `
        -SubKey "SOFTWARE\Microsoft\Office\ClickToRun\Configuration" `
        -ValueName "ProductReleaseIds"

    if ($qr7Products) {
        foreach ($prod in ($qr7Products -split ',')) {
            Write-LogOnly "Found Office C2R product in Configuration: $prod"
            if (-not $products.ContainsKey($prod.ToLower())) {
                $version = Get-RegistryValue -Hive LocalMachine `
                    -SubKey "SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\Active\culture" `
                    -ValueName "x-none"
                $products[$prod.ToLower()] = $version
                $script:C2RSuite[$prod] = "$prod - $version"
            }
        }
    }

    # O16 Configuration (QR6)
    Write-LogOnly "Check for O16 C2R products (QR6)"
    $o16Products = Get-RegistryValue -Hive LocalMachine `
        -SubKey "SOFTWARE\Microsoft\Office\16.0\ClickToRun\Configuration" `
        -ValueName "ProductReleaseIds"

    if ($o16Products) {
        foreach ($prod in ($o16Products -split ',')) {
            Write-LogOnly "Found O16 (QR6) C2R product in Configuration: $prod"
            if (-not $products.ContainsKey($prod.ToLower())) {
                $version = Get-RegistryValue -Hive LocalMachine `
                    -SubKey "SOFTWARE\Microsoft\Office\16.0\ClickToRun\ProductReleaseIDs\culture" `
                    -ValueName "x-none"
                $products[$prod.ToLower()] = $version
                $script:C2RSuite[$prod] = "$prod - $version"
            }
        }
    }

    # ARP Check
    Write-LogOnly "Check ARP for Office C2R products"
    $arpKeys = Get-RegistryKeys -Hive LocalMachine -SubKey $script:REG_ARP

    foreach ($arpKey in $arpKeys) {
        $uninstallString = Get-RegistryValue -Hive LocalMachine `
            -SubKey "$($script:REG_ARP)$arpKey" `
            -ValueName "UninstallString"

        if ($uninstallString -and
            (($uninstallString -like "*Microsoft Office 1*") -or
            ($uninstallString -like "*OfficeClickToRun.exe*"))) {

            $displayVersion = Get-RegistryValue -Hive LocalMachine `
                -SubKey "$($script:REG_ARP)$arpKey" `
                -ValueName "DisplayVersion"

            # Extract product ID from uninstall string
            if ($uninstallString -match "productstoremove=([^\s]+)") {
                $prod = $matches[1] -replace "_.*", "" -replace "\.1.*", ""

                Write-LogOnly "Found C2R product in ARP: $prod"
                if (-not $products.ContainsKey($prod.ToLower())) {
                    $products[$prod.ToLower()] = $displayVersion
                    $script:C2RSuite[$arpKey] = "$prod - $displayVersion"
                }
            }
        }
    }

    $script:InstalledSku = $products
    return $products
}

function Test-IsC2R {
    [CmdletBinding()]
    param([string]$Path)

    return $script:Orchestrator.IsC2RPath($Path)
}

function Test-ProductInScope {
    [CmdletBinding()]
    param([string]$ProductCode)

    return $script:Orchestrator.IsInScope($ProductCode)
}

#endregion

#region Service Operations

function Remove-Service {
    [CmdletBinding()]
    param([string]$ServiceName)

    Write-Log "Attempting to delete service: $ServiceName"

    if (-not $script:DetectOnly) {
        $result = $script:Orchestrator.Services.DeleteService($ServiceName)
        if ($result) {
            Write-Log "Successfully deleted service: $ServiceName"
        }
        else {
            Write-Log "Failed to delete service: $ServiceName"
        }
        return $result
    }
    else {
        Write-Log "Preview mode. Would delete service: $ServiceName"
        return $false
    }
}

#endregion

#region License/SPP Operations

function Clear-OfficeLicenses {
    [CmdletBinding()]
    param()

    if ($script:KeepLicense) {
        Write-Log "Skipping license cleanup (KeepLicense flag set)"
        return
    }

    Write-LogSubHeader "Cleaning Office licenses"

    # Clean OSPP
    Write-Log "Cleaning OSPP licenses..."
    $osVersion = [Environment]::OSVersion.Version
    $versionNT = $osVersion.Major * 100 + $osVersion.Minor

    if (-not $script:DetectOnly) {
        $script:Orchestrator.License.CleanOSPP($versionNT)
    }

    # Clean VNext license cache
    Write-Log "Cleaning VNext license cache..."
    if (-not $script:DetectOnly) {
        $script:Orchestrator.License.ClearVNextLicenseCache($script:LocalAppData)
    }
}

#endregion

#region Windows Installer Cleanup

function Clear-WindowsInstallerMetadata {
    [CmdletBinding()]
    param()

    Write-LogSubHeader "Cleaning Windows Installer metadata"

    $shouldDelete = {
        param([string]$guid)
        return Test-ProductInScope $guid
    }

    if (-not $script:DetectOnly) {
        Write-Log "Cleaning UpgradeCodes..."
        $script:Orchestrator.WindowsInstaller.CleanupUpgradeCodes($shouldDelete)

        Write-Log "Cleaning Products..."
        $script:Orchestrator.WindowsInstaller.CleanupProducts($shouldDelete)

        Write-Log "Cleaning Components..."
        $script:Orchestrator.WindowsInstaller.CleanupComponents($shouldDelete)

        Write-Log "Cleaning Published Components..."
        $script:Orchestrator.WindowsInstaller.CleanupPublishedComponents($shouldDelete)
    }
    else {
        Write-Log "Preview mode. Would clean Windows Installer metadata"
    }
}

#endregion

#region TypeLib Cleanup

function Clear-TypeLibRegistrations {
    [CmdletBinding()]
    param()

    Write-LogSubHeader "Cleaning TypeLib registrations"

    if (-not $script:DetectOnly) {
        $script:Orchestrator.TypeLib.CleanupKnownTypeLibs()
    }
    else {
        Write-Log "Preview mode. Would clean TypeLib registrations"
    }
}

#endregion

#region Shell Integration

function Clear-ShellIntegration {
    [CmdletBinding()]
    param()

    Write-LogSubHeader "Cleaning shell integration"

    # Protocol Handlers
    Remove-RegistryKey -Hive LocalMachine -SubKey "SOFTWARE\Classes\Protocols\Handler\osf"

    # Context Menu Handlers
    $contextMenuHandlers = @(
        "SOFTWARE\Classes\CLSID\{573FFD05-2805-47C2-BCE0-5F19512BEB8D}",
        "SOFTWARE\Classes\CLSID\{8BA85C75-763B-4103-94EB-9470F12FE0F7}",
        "SOFTWARE\Classes\CLSID\{CD55129A-B1A1-438E-A9AA-ABA463DBD3BF}",
        "SOFTWARE\Classes\CLSID\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}",
        "SOFTWARE\Classes\CLSID\{E768CD3B-BDDC-436D-9C13-E1B39CA257B1}"
    )

    foreach ($key in $contextMenuHandlers) {
        Remove-RegistryKey -Hive LocalMachine -SubKey $key
    }

    # Groove ShellIconOverlayIdentifiers
    $overlayIdentifiers = @(
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 1 (ErrorConflict)",
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 2 (SyncInProgress)",
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 3 (InSync)",
        "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 1 (ErrorConflict)",
        "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 2 (SyncInProgress)",
        "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 3 (InSync)"
    )

    foreach ($key in $overlayIdentifiers) {
        Remove-RegistryKey -Hive LocalMachine -SubKey $key
    }

    # Shell Extensions
    $shellExtensions = @(
        "{B28AA736-876B-46DA-B3A8-84C5E30BA492}",
        "{8B02D659-EBBB-43D7-9BBA-52CF22C5B025}",
        "{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}",
        "{42042206-2D85-11D3-8CFF-005004838597}",
        "{993BE281-6695-4BA5-8A2A-7AACBFAAB69E}",
        "{C41662BB-1FA0-4CE0-8DC5-9B7F8279FF97}",
        "{506F4668-F13E-4AA1-BB04-B43203AB3CC0}",
        "{D66DC78C-4F61-447F-942B-3FB6980118CF}",
        "{46137B78-0EC3-426D-8B89-FF7C3A458B5E}",
        "{8BA85C75-763B-4103-94EB-9470F12FE0F7}",
        "{CD55129A-B1A1-438E-A9AA-ABA463DBD3BF}",
        "{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}",
        "{E768CD3B-BDDC-436D-9C13-E1B39CA257B1}"
    )

    foreach ($guid in $shellExtensions) {
        Remove-RegistryValue -Hive LocalMachine `
            -SubKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved" `
            -ValueName $guid
    }

    # BHO (Browser Helper Objects)
    $bhoKeys = @(
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}",
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}",
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}",
        "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}",
        "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}",
        "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}"
    )

    foreach ($key in $bhoKeys) {
        Remove-RegistryKey -Hive LocalMachine -SubKey $key
    }

    # OneNote Namespace Extension
    Remove-RegistryKey -Hive LocalMachine `
        -SubKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}"

    # Web Sites
    Remove-RegistryKey -Hive LocalMachine `
        -SubKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\Namespace\{B28AA736-876B-46DA-B3A8-84C5E30BA492}"
    Remove-RegistryKey -Hive LocalMachine `
        -SubKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\NetworkNeighborhood\Namespace\{46137B78-0EC3-426D-8B89-FF7C3A458B5E}"

    # VolumeCaches
    Remove-RegistryKey -Hive LocalMachine `
        -SubKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Microsoft Office Temp Files"

    # Restart Explorer to release locks
    if (-not $script:DetectOnly) {
        Write-Log "Restarting Explorer..."
        $script:Orchestrator.Shell.RestartExplorer()
    }
}

function Clear-Shortcuts {
    [CmdletBinding()]
    param(
        [string]$RootPath,
        [switch]$Delete,
        [switch]$Unpin
    )

    if ($script:SkipSD) {
        return
    }

    Write-LogSubHeader "Cleaning shortcuts in: $RootPath"

    $shortcuts = Get-ChildItem -Path $RootPath -Filter "*.lnk" -Recurse -ErrorAction SilentlyContinue

    foreach ($shortcut in $shortcuts) {
        try {
            $shell = New-Object -ComObject WScript.Shell
            $link = $shell.CreateShortcut($shortcut.FullName)

            $shouldDelete = $false

            # Check if target is C2R-related
            if (Test-IsC2R $link.TargetPath) {
                $shouldDelete = $true
            }
            # Check if target contains a product GUID
            elseif ($link.TargetPath -match '\{[A-F0-9-]{36}\}') {
                $guid = $matches[0]
                if (Test-ProductInScope $guid) {
                    $shouldDelete = $true
                }
            }

            if ($shouldDelete) {
                if ($Unpin) {
                    Write-LogOnly "Unpinning shortcut: $($shortcut.FullName)"
                    $script:Orchestrator.Shell.UnpinFromTaskbar($shortcut.FullName)
                    $script:Orchestrator.Shell.UnpinFromStartMenu($shortcut.FullName)
                }

                if ($Delete) {
                    Write-LogOnly "Deleting shortcut: $($shortcut.FullName)"
                    if (-not $script:DetectOnly) {
                        Remove-Item -Path $shortcut.FullName -Force -ErrorAction SilentlyContinue
                    }
                }
            }

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
        }
        catch {
            Write-LogOnly "Error processing shortcut $($shortcut.FullName): $_"
        }
    }
}

#endregion

#region Export Module Members

Export-ModuleMember -Function @(
    'Initialize-Environment',
    'Get-SystemInfo',
    'Test-IsElevated',
    'Initialize-Log',
    'Write-LogHeader',
    'Write-LogSubHeader',
    'Write-Log',
    'Write-LogOnly',
    'Close-Log',
    'Set-ErrorCode',
    'Clear-ErrorCode',
    'Set-ReturnValue',
    'Get-RegistryValue',
    'Set-RegistryValue',
    'Remove-RegistryKey',
    'Remove-RegistryValue',
    'Get-RegistryKeys',
    'Get-RegistryValues',
    'Test-RegistryKeyExists',
    'Remove-FolderRecursive',
    'Remove-FileForced',
    'Add-PendingFileDelete',
    'Stop-OfficeProcesses',
    'Test-ProcessRunning',
    'Get-InstalledOfficeProducts',
    'Test-IsC2R',
    'Test-ProductInScope',
    'Remove-Service',
    'Clear-OfficeLicenses',
    'Clear-WindowsInstallerMetadata',
    'Clear-TypeLibRegistrations',
    'Clear-ShellIntegration',
    'Clear-Shortcuts'
)

Export-ModuleMember -Variable @(
    'SCRIPT_VERSION',
    'SCRIPT_NAME',
    'ERROR_SUCCESS',
    'ERROR_FAIL',
    'ERROR_REBOOT_REQUIRED',
    'ERROR_USERCANCEL',
    'ERROR_STAGE1',
    'ERROR_STAGE2',
    'ERROR_INCOMPLETE',
    'ERROR_DCAF_FAILURE',
    'ERROR_ELEVATION_USERDECLINED',
    'ERROR_ELEVATION',
    'ERROR_SCRIPTINIT',
    'ERROR_RELAUNCH',
    'ERROR_UNKNOWN',
    'REG_ARP',
    'ErrorCode',
    'RebootRequired',
    'IsElevated',
    'Is64Bit',
    'OSInfo',
    'LogDir',
    'ScrubDir',
    'InstalledSku',
    'C2RSuite',
    'KeepSku',
    'KeepFolder',
    'DelInUse',
    'ProgramFiles',
    'ProgramFilesX86',
    'CommonProgramFiles',
    'CommonProgramFilesX86',
    'ProgramData',
    'AppData',
    'LocalAppData',
    'AllUsersProfile',
    'WinDir',
    'Temp',
    'WICacheDir',
    'Orchestrator'
)

#endregion
