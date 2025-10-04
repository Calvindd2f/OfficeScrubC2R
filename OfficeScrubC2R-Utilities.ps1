# OfficeScrubC2R-Utilities.psm1
# High-performance utility functions for Office C2R removal
# Uses C# for performance-critical operations

using namespace System
using namespace System.Collections.Generic
using namespace System.IO
using namespace System.Management
using namespace System.Security.Principal
using namespace Microsoft.Win32

#region C# Performance Optimizations

# C# class for high-performance registry operations
Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Win32;

public class FastRegistryOperations
{
    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern int RegOpenKeyEx(IntPtr hKey, string subKey, int options, int samDesired, out IntPtr phkResult);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern int RegCloseKey(IntPtr hKey);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern int RegQueryValueEx(IntPtr hKey, string lpValueName, IntPtr lpReserved, out uint lpType, IntPtr lpData, ref uint lpcbData);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern int RegEnumKeyEx(IntPtr hKey, uint dwIndex, System.Text.StringBuilder lpName, ref uint lpcbName, IntPtr lpReserved, IntPtr lpClass, IntPtr lpcbClass, IntPtr lpftLastWriteTime);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern int RegEnumValue(IntPtr hKey, uint dwIndex, System.Text.StringBuilder lpValueName, ref uint lpcbValueName, IntPtr lpReserved, out uint lpType, IntPtr lpData, IntPtr lpcbData);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern int RegDeleteKey(IntPtr hKey, string lpSubKey);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern int RegDeleteValue(IntPtr hKey, string lpValueName);

    private const int KEY_READ = 0x20019;
    private const int KEY_WRITE = 0x20006;
    private const int KEY_ALL_ACCESS = 0xF003F;

    public static Dictionary<string, string> GetRegistrySubKeys(RegistryHive hive, string subKey)
    {
        var result = new Dictionary<string, string>();
        IntPtr hKey = IntPtr.Zero;

        try
        {
            IntPtr hHive = GetHiveHandle(hive);
            if (RegOpenKeyEx(hHive, subKey, 0, KEY_READ, out hKey) == 0)
            {
                uint index = 0;
                var nameBuffer = new System.Text.StringBuilder(256);
                uint nameSize = 256;

                while (RegEnumKeyEx(hKey, index, nameBuffer, ref nameSize, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero) == 0)
                {
                    result[nameBuffer.ToString()] = "";
                    nameSize = 256;
                    nameBuffer.Clear();
                    index++;
                }
            }
        }
        finally
        {
            if (hKey != IntPtr.Zero)
                RegCloseKey(hKey);
        }

        return result;
    }

    public static Dictionary<string, object> GetRegistryValues(RegistryHive hive, string subKey)
    {
        var result = new Dictionary<string, object>();
        IntPtr hKey = IntPtr.Zero;

        try
        {
            IntPtr hHive = GetHiveHandle(hive);
            if (RegOpenKeyEx(hHive, subKey, 0, KEY_READ, out hKey) == 0)
            {
                uint index = 0;
                var nameBuffer = new System.Text.StringBuilder(256);
                uint nameSize = 256;
                uint valueType;

                while (RegEnumValue(hKey, index, nameBuffer, ref nameSize, IntPtr.Zero, out valueType, IntPtr.Zero, IntPtr.Zero) == 0)
                {
                    result[nameBuffer.ToString()] = valueType;
                    nameSize = 256;
                    nameBuffer.Clear();
                    index++;
                }
            }
        }
        finally
        {
            if (hKey != IntPtr.Zero)
                RegCloseKey(hKey);
        }

        return result;
    }

    public static bool DeleteRegistryKey(RegistryHive hive, string subKey)
    {
        IntPtr hHive = GetHiveHandle(hive);
        return RegDeleteKey(hHive, subKey) == 0;
    }

    public static bool DeleteRegistryValue(RegistryHive hive, string subKey, string valueName)
    {
        IntPtr hKey = IntPtr.Zero;

        try
        {
            IntPtr hHive = GetHiveHandle(hive);
            if (RegOpenKeyEx(hHive, subKey, 0, KEY_WRITE, out hKey) == 0)
            {
                return RegDeleteValue(hKey, valueName) == 0;
            }
        }
        finally
        {
            if (hKey != IntPtr.Zero)
                RegCloseKey(hKey);
        }

        return false;
    }

    private static IntPtr GetHiveHandle(RegistryHive hive)
    {
        switch (hive)
        {
            case RegistryHive.ClassesRoot: return new IntPtr(0x80000000);
            case RegistryHive.CurrentUser: return new IntPtr(0x80000001);
            case RegistryHive.LocalMachine: return new IntPtr(0x80000002);
            case RegistryHive.Users: return new IntPtr(0x80000003);
            default: return IntPtr.Zero;
        }
    }
}

public class FastFileOperations
{
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool DeleteFile(string lpFileName);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool RemoveDirectory(string lpPathName);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool MoveFileEx(string lpExistingFileName, string lpNewFileName, int dwFlags);

    private const int MOVEFILE_DELAY_UNTIL_REBOOT = 0x4;

    public static bool DeleteFileFast(string filePath)
    {
        return DeleteFile(filePath);
    }

    public static bool RemoveDirectoryFast(string directoryPath)
    {
        return RemoveDirectory(directoryPath);
    }

    public static bool ScheduleDeleteOnReboot(string filePath)
    {
        return MoveFileEx(filePath, null, MOVEFILE_DELAY_UNTIL_REBOOT);
    }
}

public class ProcessManager
{
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, int dwProcessId);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool TerminateProcess(IntPtr hProcess, uint uExitCode);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr hObject);

    private const uint PROCESS_TERMINATE = 0x0001;
    private const uint PROCESS_QUERY_INFORMATION = 0x0400;

    public static bool TerminateProcessById(int processId)
    {
        IntPtr hProcess = OpenProcess(PROCESS_TERMINATE | PROCESS_QUERY_INFORMATION, false, processId);
        if (hProcess == IntPtr.Zero)
            return false;

        bool result = TerminateProcess(hProcess, 0);
        CloseHandle(hProcess);
        return result;
    }
}

public class GuidDecoder
{
    // Base85 decoding table (ASCII character to value mapping)
    private static readonly byte[] DecodeTable = new byte[256];

    static GuidDecoder()
    {
        // Initialize decode table with 0xff (invalid) for all characters
        for (int i = 0; i < 256; i++)
        {
            DecodeTable[i] = 0xff;
        }

        // Set valid Base85 character mappings
        byte[] validChars = { 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124 };
        byte[] values = { 0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0a, 0x0b, 0xff, 0x0c, 0x0d, 0x0e, 0x0f, 0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0xff, 0xff, 0xff, 0x16, 0xff, 0x17, 0x18, 0x19, 0x1a, 0x1b, 0x1c, 0x1d, 0x1e, 0x1f, 0x20, 0x21, 0x22, 0x23, 0x24, 0x25, 0x26, 0x27, 0x28, 0x29, 0x2a, 0x2b, 0x2c, 0x2d, 0x2e, 0x2f, 0x30, 0x31, 0x32, 0x33, 0xff, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3a, 0x3b, 0x3c, 0x3d, 0x3e, 0x3f, 0x40, 0x41, 0x42, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48, 0x49, 0x4a, 0x4b, 0x4c, 0x4d, 0x4e, 0x4f, 0x50, 0x51, 0x52, 0xff, 0x53, 0x54, 0xff };

        for (int i = 0; i < validChars.Length && i < values.Length; i++)
        {
            DecodeTable[validChars[i]] = values[i];
        }
    }

    public static string DecodeGuid(string encodedGuid)
    {
        if (string.IsNullOrEmpty(encodedGuid))
            return null;

        try
        {
            bool failed = false;
            uint total = 0;
            uint pow85 = 1;
            var decode = new System.Text.StringBuilder();

            int maxLength = Math.Min(encodedGuid.Length, 20);

            for (int i = 0; i < maxLength; i++)
            {
                failed = true;

                if (i % 5 == 0)
                {
                    total = 0;
                    pow85 = 1;
                }

                int charCode = (int)encodedGuid[i];

                if (charCode >= 128 || DecodeTable[charCode] == 0xff)
                {
                    break;
                }

                uint hexValue = DecodeTable[charCode];
                total += hexValue * pow85;

                if (i % 5 == 4)
                {
                    decode.Append(total.ToString("X8"));
                }

                pow85 *= 85;
                failed = false;
            }

            if (!failed && decode.Length >= 32)
            {
                string decodedString = decode.ToString();

                // Reconstruct GUID in the specific order from the VBScript
                return "{" +
                       decodedString.Substring(0, 8) + "-" +
                       decodedString.Substring(12, 4) + "-" +
                       decodedString.Substring(8, 4) + "-" +
                       decodedString.Substring(22, 2) + decodedString.Substring(20, 2) + "-" +
                       decodedString.Substring(18, 2) + decodedString.Substring(16, 2) +
                       decodedString.Substring(30, 2) + decodedString.Substring(28, 2) +
                       decodedString.Substring(26, 2) + decodedString.Substring(24, 2) +
                       "}";
            }

            return null;
        }
        catch
        {
            return null;
        }
    }
}
"@

#endregion

#region Constants
$script:SCRIPT_VERSION = "2.19"
$script:SCRIPT_NAME = "OfficeScrubC2R"
$script:RET_VAL_FILE = "ScrubRetValFile.txt"
$script:OFFICE_NAME = "Office C2R / O365"

# Registry Hive Constants
$script:HKCR = 0x80000000
$script:HKCU = 0x80000001
$script:HKLM = 0x80000002
$script:HKU = 0x80000003

# Error Constants
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
$script:ERROR_ALL = 4095
$script:ERROR_USER_ABORT = 0xC000013A
$script:ERROR_SUCCESS_CONFIG_COMPLETE = 1728
$script:ERROR_SUCCESS_REBOOT_REQUIRED = 3010

# Registry Paths
$script:REG_ARP = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
$script:REG_O15RPROPERTYBAG = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag\"
$script:REG_O15C2RCONFIGURATION = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration\"
$script:REG_O15C2RPRODUCTIDS = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs\Active\"
$script:REG_O16C2RCONFIGURATION = "SOFTWARE\Microsoft\Office\16.0\ClickToRun\Configuration\"
$script:REG_O16C2RPRODUCTIDS = "SOFTWARE\Microsoft\Office\16.0\ClickToRun\ProductReleaseIDs\Active\"
$script:REG_C2RCONFIGURATION = "SOFTWARE\Microsoft\Office\ClickToRun\Configuration\"
$script:REG_C2RPRODUCTIDS = "SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\"
#endregion

#region Global Variables
$script:ErrorCode = $script:ERROR_SUCCESS
$script:LogInitialized = $false
$script:IsElevated = $false
$script:Is64Bit = $false
$script:RebootRequired = $false
$script:Quiet = $true
$script:DetectOnly = $false
$script:Force = $false
$script:RemoveAll = $false
$script:KeepLicense = $false
$script:Offline = $false
$script:ForceArpUninstall = $false
$script:ClearTaskBand = $false
$script:UnpinMode = $false
$script:SkipSD = $false
$script:NoElevate = $false
$script:UserConsent = $false
$script:ReturnErrorOrSuccess = $false
$script:TestRerun = $false
$script:SetRunOnce = $false
$script:Rerun = $false
$script:RemoveOse = $false
$script:NoCancel = $false
$script:C2R = $true

# Path Variables
$script:AppData = $null
$script:LocalAppData = $null
$script:Temp = $null
$script:AllUsersProfile = $null
$script:ProgramFiles = $null
$script:ProgramFilesX86 = $null
$script:CommonProgramFiles = $null
$script:CommonProgramFilesX86 = $null
$script:ProgramData = $null
$script:WinDir = $null
$script:WICacheDir = $null
$script:ScrubDir = $null
$script:ScriptDir = $null
$script:LogDir = $null
$script:ProfilesDirectory = $null
$script:PackageFolder = $null
$script:PackageGuid = $null
$script:OSVersion = $null
$script:OSInfo = $null
$script:VersionNT = 0

# Collections
$script:InstalledSku = @{}
$script:RemoveSku = @{}
$script:KeepSku = @{}
$script:KeepLis = @{}
$script:KeepFolder = @{}
$script:Apps = @{}
$script:DelRegKey = @{}
$script:KeepReg = @{}
$script:C2RSuite = @{}
$script:DelInUse = @{}
$script:DelFolder = @{}
$script:SC = @{}

# Office Process Names
$script:OfficeProcesses = @(
    "appvshnotify.exe",
    "integratedoffice.exe",
    "integrator.exe",
    "firstrun.exe",
    "communicator.exe",
    "msosync.exe",
    "OneNoteM.exe",
    "iexplore.exe",
    "mavinject32.exe",
    "werfault.exe",
    "perfboost.exe",
    "roamingoffice.exe",
    "officeclicktorun.exe",
    "officeondemand.exe",
    "OfficeC2RClient.exe"
)
#endregion

#region Logging Functions
function Write-LogHeader {
    param([string]$Message)
    if ($script:LogInitialized) {
        $timestamp = Get-Date -Format "HH:mm:ss"
        Write-Output "================================================"
        Write-Output "$timestamp - $Message"
        Write-Output "================================================"
    }
}

function Write-LogSubHeader {
    param([string]$Message)
    if ($script:LogInitialized) {
        $timestamp = Get-Date -Format "HH:mm:ss"
        Write-Output "----------------------------------------"
        Write-Output "$timestamp - $Message"
        Write-Output "----------------------------------------"
    }
}

function Write-Log {
    param([string]$Message)
    if ($script:LogInitialized) {
        $timestamp = Get-Date -Format "HH:mm:ss"
        Write-Output "$timestamp - $Message"
    }
}

function Write-LogOnly {
    param([string]$Message)
    if ($script:LogInitialized -and -not $script:Quiet) {
        $timestamp = Get-Date -Format "HH:mm:ss"
        Write-Output "$timestamp - $Message"
    }
}

function Initialize-Log {
    param([string]$LogPath)

    if (-not $script:LogInitialized) {
        $script:LogDir = $LogPath
        if (-not (Test-Path $script:LogDir)) {
            [void](New-Item -ItemType Directory -Path $script:LogDir -Force)
        }
        $script:LogInitialized = $true
        Write-LogHeader "Office C2R Scrubber v$script:SCRIPT_VERSION - Log Initialized"
    }
}
#endregion

#region Registry Functions
function Test-RegistryKeyExists {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [string]$SubKey
    )

    try {
        $regKey = [Microsoft.Win32.Registry]::$($Hive.ToString()).OpenSubKey($SubKey, $false)
        if ($regKey) {
            $regKey.Close()
            return $true
        }
        return $false
    }
    catch {
        return $false
    }
}

function Get-RegistryValue {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [string]$SubKey,
        [string]$ValueName,
        [string]$ValueType = "REG_SZ"
    )

    try {
        $regKey = [Microsoft.Win32.Registry]::$($Hive.ToString()).OpenSubKey($SubKey, $false)
        if ($regKey) {
            $value = $regKey.GetValue($ValueName)
            $regKey.Close()
            return $value
        }
        return $null
    }
    catch {
        return $null
    }
}

function Set-RegistryValue {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [string]$SubKey,
        [string]$ValueName,
        [object]$Value,
        [Microsoft.Win32.RegistryValueKind]$ValueType = [Microsoft.Win32.RegistryValueKind]::String
    )

    try {
        $regKey = [Microsoft.Win32.Registry]::$($Hive.ToString()).OpenSubKey($SubKey, $true)
        if ($regKey) {
            $regKey.SetValue($ValueName, $Value, $ValueType)
            $regKey.Close()
            return $true
        }
        return $false
    }
    catch {
        return $false
    }
}

function Remove-RegistryKey {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [string]$SubKey
    )

    try {
        $parentKey = Split-Path $SubKey -Parent
        $keyName = Split-Path $SubKey -Leaf

        if ($parentKey -eq "") {
            $parentKey = $SubKey
            $keyName = ""
        }

        $regKey = [Microsoft.Win32.Registry]::$($Hive.ToString()).OpenSubKey($parentKey, $true)
        if ($regKey) {
            if ($keyName -eq "") {
                $regKey.Close()
                [Microsoft.Win32.Registry]::$($Hive.ToString()).DeleteSubKey($parentKey, $true)
            }
            else {
                $regKey.DeleteSubKey($keyName, $true)
                $regKey.Close()
            }
            return $true
        }
        return $false
    }
    catch {
        return $false
    }
}

function Remove-RegistryValue {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [string]$SubKey,
        [string]$ValueName
    )

    try {
        $regKey = [Microsoft.Win32.Registry]::$($Hive.ToString()).OpenSubKey($SubKey, $true)
        if ($regKey) {
            $regKey.DeleteValue($ValueName, $true)
            $regKey.Close()
            return $true
        }
        return $false
    }
    catch {
        return $false
    }
}

function Get-RegistrySubKeys {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [string]$SubKey
    )

    try {
        $regKey = [Microsoft.Win32.Registry]::$($Hive.ToString()).OpenSubKey($SubKey, $false)
        if ($regKey) {
            $subKeys = $regKey.GetSubKeyNames()
            $regKey.Close()
            return $subKeys
        }
        return @()
    }
    catch {
        return @()
    }
}

function Get-RegistryValues {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [string]$SubKey
    )

    try {
        $regKey = [Microsoft.Win32.Registry]::$($Hive.ToString()).OpenSubKey($SubKey, $false)
        if ($regKey) {
            $valueNames = $regKey.GetValueNames()
            $regKey.Close()
            return $valueNames
        }
        return @()
    }
    catch {
        return @()
    }
}
#endregion

#region File Operations
function Remove-FileFast {
    param([string]$FilePath)

    try {
        if (Test-Path $FilePath) {
            return [FastFileOperations]::DeleteFileFast($FilePath)
        }
        return $true
    }
    catch {
        return $false
    }
}

function Remove-FolderFast {
    param([string]$FolderPath)

    try {
        if (Test-Path $FolderPath) {
            return [FastFileOperations]::RemoveDirectoryFast($FolderPath)
        }
        return $true
    }
    catch {
        return $false
    }
}

function Schedule-DeleteOnReboot {
    param([string]$Path)

    try {
        return [FastFileOperations]::ScheduleDeleteOnReboot($Path)
    }
    catch {
        return $false
    }
}

function Remove-FolderRecursive {
    param(
        [string]$Path,
        [switch]$Force
    )

    try {
        if (Test-Path $Path) {
            if ($Force) {
                Get-ChildItem -Path $Path -Recurse -Force | Remove-Item -Force -Recurse
            }
            Remove-Item -Path $Path -Force -Recurse
            return $true
        }
        return $true
    }
    catch {
        return $false
    }
}
#endregion

#region Process Management
function Stop-OfficeProcesses {
    param([switch]$Force)

    Write-LogSubHeader "Stopping Office processes"

    foreach ($processName in $script:OfficeProcesses) {
        $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
        foreach ($process in $processes) {
            try {
                Write-Log "Stopping process: $($process.ProcessName) (PID: $($process.Id))"
                if ($Force) {
                    $process.Kill()
                }
                else {
                    $process.CloseMainWindow()
                    if (-not $process.WaitForExit(5000)) {
                        $process.Kill()
                    }
                }
            }
            catch {
                Write-Log "Failed to stop process: $($process.ProcessName) - $($_.Exception.Message)"
            }
        }
    }
}

function Test-ProcessRunning {
    param([string]$ProcessName)

    $processes = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    return $processes.Count -gt 0
}
#endregion

#region Office Detection
function Test-IsC2R {
    param([string]$Value)

    $c2rPatterns = @(
        "officeclicktorun",
        "o365proplus",
        "o365business",
        "o365homepremium",
        "o365home",
        "o365smallbusiness",
        "o365midsizebusiness",
        "o365enterprise"
    )

    $valueLower = $Value.ToLower()
    foreach ($pattern in $c2rPatterns) {
        if ($valueLower -like "*$pattern*") {
            return $true
        }
    }
    return $false
}

function Get-InstalledOfficeProducts {
    Write-LogSubHeader "Detecting installed Office products"

    $products = @{}

    # Check O15 C2R Configuration
    Write-LogOnly "Check for O15 C2R products"
    $o15ProductReleaseIds = Get-RegistryValue -Hive LocalMachine -SubKey $script:REG_O15C2RCONFIGURATION -ValueName "ProductReleaseIds"
    if ($o15ProductReleaseIds) {
        $o15Products = $o15ProductReleaseIds -split ","
        $o15Version = Get-RegistryValue -Hive LocalMachine -SubKey "$script:REG_O15C2RPRODUCTIDS\culture" -ValueName "x-none"
        foreach ($prod in $o15Products) {
            Write-LogOnly "Found O15 C2R product in Configuration: $prod"
            if (-not $products.ContainsKey($prod.ToLower())) {
                Write-LogOnly "Add new product to dictionary: $($prod.ToLower())"
                $products[$prod.ToLower()] = $o15Version
            }
        }
    }

    # Check O15 PropertyBag
    $o15PropertyBag = Get-RegistryValue -Hive LocalMachine -SubKey $script:REG_O15RPROPERTYBAG -ValueName "productreleaseid"
    if ($o15PropertyBag) {
        $o15Products = $o15PropertyBag -split ","
        $o15Version = Get-RegistryValue -Hive LocalMachine -SubKey "$script:REG_O15C2RPRODUCTIDS\culture" -ValueName "x-none"
        foreach ($prod in $o15Products) {
            Write-LogOnly "Found O15 C2R product in PropertyBag: $prod"
            if (-not $products.ContainsKey($prod.ToLower())) {
                Write-LogOnly "Add new product to dictionary: $($prod.ToLower())"
                $products[$prod.ToLower()] = $o15Version
            }
        }
    }

    # Check Office C2R products (>=QR8)
    Write-LogOnly "Check for Office C2R products (>=QR8)"
    $activeConfiguration = Get-RegistryValue -Hive LocalMachine -SubKey $script:REG_C2RPRODUCTIDS -ValueName "ActiveConfiguration"
    if ($activeConfiguration) {
        $versionFallback = Get-RegistryValue -Hive LocalMachine -SubKey "$script:REG_C2RPRODUCTIDS\$activeConfiguration\culture" -ValueName "x-none"

        $cultures = Get-RegistrySubKeys -Hive LocalMachine -SubKey "$script:REG_C2RPRODUCTIDS\$activeConfiguration\culture"
        if ($cultures) {
            foreach ($cult in $cultures) {
                if ($cult.ToLower() -like "*x-none*") {
                    $versionFallback = Get-RegistryValue -Hive LocalMachine -SubKey "$script:REG_C2RPRODUCTIDS\$activeConfiguration\culture\$cult" -ValueName "Version"
                }
            }
        }

        $officeProducts = Get-RegistrySubKeys -Hive LocalMachine -SubKey "$script:REG_C2RPRODUCTIDS\$activeConfiguration"
        if ($officeProducts) {
            foreach ($prod in $officeProducts) {
                $sProd = $prod.ToLower()
                if ($sProd -like "*.*") {
                    $sProd = $sProd.Substring(0, $sProd.IndexOf("."))
                }

                if ($sProd -notin @("culture", "stream")) {
                    Write-LogOnly "Found Office C2R product in Configuration: $prod"
                    if (-not $products.ContainsKey($sProd)) {
                        Write-LogOnly "Add new product to dictionary: $sProd"
                        $displayVersion = Get-RegistryValue -Hive LocalMachine -SubKey "$script:REG_C2RPRODUCTIDS\$activeConfiguration\$prod\x-none" -ValueName "Version"
                        if ($displayVersion) {
                            $products[$sProd] = $displayVersion
                        }
                        else {
                            $products[$sProd] = $versionFallback
                        }
                    }
                }
            }
        }
    }

    # Check ARP entries
    Write-LogOnly "Check ARP entries for C2R products"
    $arpKeys = Get-RegistrySubKeys -Hive LocalMachine -SubKey $script:REG_ARP
    foreach ($key in $arpKeys) {
        $displayName = Get-RegistryValue -Hive LocalMachine -SubKey "$script:REG_ARP$key" -ValueName "DisplayName"
        if ($displayName -and (Test-IsC2R $displayName)) {
            $uninstallString = Get-RegistryValue -Hive LocalMachine -SubKey "$script:REG_ARP$key" -ValueName "UninstallString"
            if ($uninstallString) {
                Write-LogOnly "Found C2R product in ARP: $displayName"
                $script:C2RSuite[$key] = $displayName
            }
        }
    }

    return $products
}
#endregion

#region System Information
function Test-IsElevated {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-SystemInfo {
    try {
        $computerSystem = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop
        $script:Is64Bit = $computerSystem.SystemType -like "*64*"

        $os = Get-WmiObject -Class Win32_OperatingSystem -ErrorAction Stop
        $script:OSVersion = $os.Version
        $script:OSInfo = "$($os.Caption) $($os.OtherTypeDescription), SP $($os.ServicePackMajorVersion), Version: $($os.Version), Codepage: $($os.CodeSet), Country Code: $($os.CountryCode), Language: $($os.OSLanguage)"

        if ($script:OSVersion) {
            $versionParts = $script:OSVersion -split "\."
            if ($versionParts.Count -ge 2) {
                $script:VersionNT = [int]$versionParts[0] * 100 + [int]$versionParts[1]
            }
        }
    }
    catch {
        Write-Warning "Failed to get system information: $($_.Exception.Message)"
        $script:Is64Bit = [Environment]::Is64BitOperatingSystem
        $script:OSVersion = [Environment]::OSVersion.Version.ToString()
        $script:OSInfo = "Unknown OS"
        $script:VersionNT = 0
    }
}

function Initialize-Environment {
    # Set environment variables
    $script:AppData = $env:APPDATA
    $script:LocalAppData = $env:LOCALAPPDATA
    $script:Temp = $env:TEMP
    $script:AllUsersProfile = $env:ALLUSERSPROFILE
    $script:ProgramFiles = $env:ProgramFiles
    $script:CommonProgramFiles = $env:CommonProgramFiles
    $script:ProgramData = $env:ProgramData
    $script:WinDir = $env:WINDIR
    $script:WICacheDir = "$script:WinDir\Installer"
    $script:ScrubDir = "$script:Temp\$script:SCRIPT_NAME"
    $script:ScriptDir = Split-Path -Parent $MyInvocation.PSCommandPath
    $script:Notepad = "$script:WinDir\notepad.exe"

    # Set 64-bit specific paths
    if ($script:Is64Bit) {
        $script:ProgramFilesX86 = ${env:ProgramFiles(x86)}
        $script:CommonProgramFilesX86 = ${env:CommonProgramFiles(x86)}
    }

    # Get profiles directory
    $script:ProfilesDirectory = Get-RegistryValue -Hive LocalMachine -SubKey "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" -ValueName "ProfilesDirectory"
    if (-not $script:ProfilesDirectory -or -not (Test-Path $script:ProfilesDirectory)) {
        $script:ProfilesDirectory = Split-Path -Parent $env:USERPROFILE
    }

    # Create temp directory
    if (-not (Test-Path $script:ScrubDir)) {
        [void](New-Item -ItemType Directory -Path $script:ScrubDir -Force)
    }

    # Set default log directory
    $script:LogDir = $script:ScrubDir
}
#endregion

#region Error Handling
function Set-ErrorCode {
    param([int]$ErrorBit)
    $script:ErrorCode = $script:ErrorCode -bor $ErrorBit
}

function Clear-ErrorCode {
    param([int]$ErrorBit)
    $script:ErrorCode = $script:ErrorCode -band (-bnot $ErrorBit)
}

function Set-ReturnValue {
    param([int]$Err)
    $script:ErrorCode = $Err
    if ($script:ScrubDir) {
        $retValPath = Join-Path $script:ScrubDir $script:RET_VAL_FILE
        Set-Content -Path $retValPath -Value $Err -Force
    }
}

function Get-ReturnValueFromFile {
    if ($script:ScrubDir) {
        $retValPath = Join-Path $script:ScrubDir $script:RET_VAL_FILE
        if (Test-Path $retValPath) {
            $reader = [System.IO.StreamReader]::new($retValPath)
            try {
                $content = $reader.ReadToEnd()
            }
            finally {
                $reader.Close()
            }
            return [int]$content
        }
    }
    return $script:ERROR_UNKNOWN
}
#endregion

#region GUID Operations
function Get-ExpandedGuid {
    param([string]$Guid)

    if ($Guid.Length -eq 32) {
        # Convert compressed GUID to expanded format
        $expanded = $Guid.Substring(0, 8) + "-" + $Guid.Substring(8, 4) + "-" + $Guid.Substring(12, 4) + "-" + $Guid.Substring(16, 4) + "-" + $Guid.Substring(20, 12)
        return $expanded.ToUpper()
    }
    return $Guid.ToUpper()
}

function Get-CompressedGuid {
    param([string]$Guid)

    if ($Guid.Length -eq 36 -and $Guid.Contains("-")) {
        # Convert expanded GUID to compressed format
        return $Guid.Replace("-", "").ToUpper()
    }
    return $Guid.ToUpper()
}

function Get-DecodedGuid {
    param(
        [string]$EncGuid,
        [string]$Guid
    )

    try {
        # Use C# for GUID decoding
        $decodedGuid = [GuidDecoder]::DecodeGuid($EncGuid)

        if ($decodedGuid) {
            return $decodedGuid
        }

        # Fallback to expanded GUID format if decoding fails
        return Get-ExpandedGuid $Guid
    }
    catch {
        Write-Log "Failed to decode GUID: $($_.Exception.Message)"
        return Get-ExpandedGuid $Guid
    }
}
#endregion

#region Service Management
function Remove-Service {
    param([string]$ServiceName)

    try {
        $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($service) {
            Write-Log "Removing service: $ServiceName"
            if ($service.Status -eq "Running") {
                Stop-Service -Name $ServiceName -Force
            }
            [void](sc.exe delete $ServiceName)
            return $true
        }
        return $true
    }
    catch {
        Write-Log "Failed to remove service: $ServiceName - $($_.Exception.Message)"
        return $false
    }
}
#endregion

Export-ModuleMember -Function *
