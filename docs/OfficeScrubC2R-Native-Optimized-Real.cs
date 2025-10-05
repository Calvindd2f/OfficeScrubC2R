using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace OfficeScrubNative
{
    #region Enums and Constants

    public enum RegistryHiveType
    {
        ClassesRoot = unchecked((int)0x80000000),
        CurrentUser = unchecked((int)0x80000001),
        LocalMachine = unchecked((int)0x80000002),
        Users = unchecked((int)0x80000003)
    }

    public static class OfficeConstants
    {
        public const string OFFICE_ID = "0000000FF1CE}";
        public const int PROD_LEN = 13;
        public const int SQUISHED = 20;
        public const int COMPRESSED = 32;

        // Keep as array - only used with .Any() which doesn't benefit from HashSet
        public static readonly string[] C2R_PATTERNS = new[]
        {
            @"\ROOT\OFFICE1",
            @"Microsoft Office\Root\",
            @"\microsoft shared\ClickToRun",
            @"\Microsoft Office\PackageManifests",
            @"\Microsoft Office\PackageSunrisePolicies",
            @"Microsoft Office 15",
            @"Microsoft Office 16"
        };

        public static readonly string[] OFFICE_PROCESSES = new[]
        {
            "appvshnotify.exe", "integratedoffice.exe", "integrator.exe", "firstrun.exe",
            "communicator.exe", "msosync.exe", "OneNoteM.exe", "iexplore.exe",
            "mavinject32.exe", "werfault.exe", "perfboost.exe", "roamingoffice.exe",
            "officeclicktorun.exe", "officeondemand.exe", "OfficeC2RClient.exe",
            "winword.exe", "excel.exe", "powerpnt.exe", "outlook.exe", "onenote.exe",
            "mspub.exe", "msaccess.exe", "lync.exe", "skype.exe", "teams.exe"
        };
    }

    #endregion

    #region P/Invoke Declarations

    internal static class NativeMethods
    {
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern bool MoveFileEx(
            string lpExistingFileName,
            string lpNewFileName,
            int dwFlags);

        [DllImport("advapi32.dll", SetLastError = true)]
        internal static extern int RegOpenKeyEx(
            UIntPtr hKey,
            string lpSubKey,
            uint ulOptions,
            int samDesired,
            out UIntPtr phkResult);

        [DllImport("advapi32.dll", SetLastError = true)]
        internal static extern int RegCloseKey(UIntPtr hKey);

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern int RegDeleteKeyEx(
            UIntPtr hKey,
            string lpSubKey,
            int samDesired,
            uint Reserved);

        [DllImport("advapi32.dll", SetLastError = true)]
        internal static extern int RegDeleteValue(
            UIntPtr hKey,
            string lpValueName);

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern int RegEnumKeyEx(
            UIntPtr hKey,
            uint dwIndex,
            StringBuilder lpName,
            ref uint lpcchName,
            IntPtr lpReserved,
            IntPtr lpClass,
            IntPtr lpcchClass,
            IntPtr lpftLastWriteTime);

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern int RegEnumValue(
            UIntPtr hKey,
            uint dwIndex,
            StringBuilder lpValueName,
            ref uint lpcchValueName,
            IntPtr lpReserved,
            IntPtr lpType,
            IntPtr lpData,
            IntPtr lpcbData);

        [DllImport("advapi32.dll", SetLastError = true)]
        internal static extern int RegQueryValueEx(
            UIntPtr hKey,
            string lpValueName,
            IntPtr lpReserved,
            out uint lpType,
            IntPtr lpData,
            ref uint lpcbData);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        internal static extern IntPtr ILCreateFromPath(string pszPath);

        [DllImport("shell32.dll")]
        internal static extern void ILFree(IntPtr pidl);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        internal static extern int SHCreateShellItem(
            IntPtr pidlParent,
            IntPtr psfParent,
            IntPtr pidl,
            out IShellItem ppsi);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern bool SetFileAttributes(
            string lpFileName,
            uint dwFileAttributes);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern bool RemoveDirectory(string lpPathName);

        internal const int KEY_WOW64_64KEY = 0x0100;
        internal const int KEY_WOW64_32KEY = 0x0200;
        internal const int KEY_READ = 0x20019;
        internal const int KEY_WRITE = 0x20006;
        internal const int KEY_ALL_ACCESS = 0xF003F;
        internal const int MOVEFILE_DELAY_UNTIL_REBOOT = 0x4;
        internal const int ERROR_SUCCESS = 0;
        internal const int ERROR_NO_MORE_ITEMS = 259;
    }

    [ComImport]
    [Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IShellItem
    {
        void BindToHandler(IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid bhid, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(uint sigdnName, out IntPtr ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }

    #endregion

    #region GUID Utilities

    public static class GuidHelper
    {
        // OPTIMIZATION: Static decode table - allocate once instead of per call
        private static readonly byte[] DecodeTable = new byte[128]
        {
            0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,
            0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,
            0xff,0x00,0xff,0xff,0x01,0x02,0x03,0x04,0x05,0x06,0x07,0x08,0x09,0x0a,0x0b,0xff,
            0x0c,0x0d,0x0e,0x0f,0x10,0x11,0x12,0x13,0x14,0x15,0xff,0xff,0xff,0x16,0xff,0x17,
            0x18,0x19,0x1a,0x1b,0x1c,0x1d,0x1e,0x1f,0x20,0x21,0x22,0x23,0x24,0x25,0x26,0x27,
            0x28,0x29,0x2a,0x2b,0x2c,0x2d,0x2e,0x2f,0x30,0x31,0x32,0x33,0xff,0x34,0x35,0x36,
            0x37,0x38,0x39,0x3a,0x3b,0x3c,0x3d,0x3e,0x3f,0x40,0x41,0x42,0x43,0x44,0x45,0x46,
            0x47,0x48,0x49,0x4a,0x4b,0x4c,0x4d,0x4e,0x4f,0x50,0x51,0x52,0xff,0x53,0x54,0xff
        };

        // OPTIMIZATION: Use Span<char> to avoid StringBuilder allocations
        public static string GetExpandedGuid(string compressedGuid)
        {
            if (string.IsNullOrEmpty(compressedGuid) || compressedGuid.Length != 32)
                return null;

            try
            {
                Span<char> result = stackalloc char[38];
                result[0] = '{';
                result[9] = '-';
                result[14] = '-';
                result[19] = '-';
                result[24] = '-';
                result[37] = '}';

                ReverseToSpan(compressedGuid.AsSpan(0, 8), result.Slice(1, 8));
                ReverseToSpan(compressedGuid.AsSpan(8, 4), result.Slice(10, 4));
                ReverseToSpan(compressedGuid.AsSpan(12, 4), result.Slice(15, 4));

                for (int i = 0; i < 4; i += 2)
                {
                    result[20 + i] = compressedGuid[17 + i];
                    result[21 + i] = compressedGuid[16 + i];
                }

                for (int i = 0; i < 12; i += 2)
                {
                    result[25 + i] = compressedGuid[21 + i];
                    result[26 + i] = compressedGuid[20 + i];
                }

                return new string(result).ToUpperInvariant();
            }
            catch
            {
                return null;
            }
        }

        // OPTIMIZATION: Use Span<char> to avoid StringBuilder allocations
        public static string GetCompressedGuid(string expandedGuid)
        {
            if (string.IsNullOrEmpty(expandedGuid) || expandedGuid.Length != 38)
                return null;

            try
            {
                ReadOnlySpan<char> guid = expandedGuid.AsSpan(1, 36);
                Span<char> result = stackalloc char[32];

                ReverseToSpan(guid.Slice(0, 8), result.Slice(0, 8));
                ReverseToSpan(guid.Slice(9, 4), result.Slice(8, 4));
                ReverseToSpan(guid.Slice(14, 4), result.Slice(12, 4));

                for (int i = 0; i < 4; i += 2)
                {
                    result[16 + i] = guid[20 + i + 1];
                    result[17 + i] = guid[20 + i];
                }

                for (int i = 0; i < 12; i += 2)
                {
                    result[20 + i] = guid[25 + i + 1];
                    result[21 + i] = guid[25 + i];
                }

                return new string(result).ToUpperInvariant();
            }
            catch
            {
                return null;
            }
        }

        // OPTIMIZATION: Use Span<char> and static DecodeTable
        public static bool GetDecodedGuid(string encodedGuid, out string decodedGuid)
        {
            decodedGuid = null;

            if (string.IsNullOrEmpty(encodedGuid) || encodedGuid.Length != 20)
                return false;

            try
            {
                Span<char> decoded = stackalloc char[32];
                long total = 0;
                long pow85 = 1;
                int pos = 0;

                for (int i = 0; i < 20; i++)
                {
                    if (i % 5 == 0)
                    {
                        total = 0;
                        pow85 = 1;
                    }

                    int ascii = encodedGuid[i];
                    if (ascii >= 128 || DecodeTable[ascii] == 0xff)
                        return false;

                    total += DecodeTable[ascii] * pow85;

                    if (i % 5 == 4)
                    {
                        total.TryFormat(decoded.Slice(pos, 8), out _, "X8");
                        pos += 8;
                    }

                    pow85 *= 85;
                }

                decodedGuid = string.Format("{{{0}-{1}-{2}-{3}{4}-{5}{6}{7}{8}{9}{10}}}",
                    new string(decoded.Slice(0, 8)),
                    new string(decoded.Slice(12, 4)),
                    new string(decoded.Slice(8, 4)),
                    new string(decoded.Slice(22, 2)),
                    new string(decoded.Slice(20, 2)),
                    new string(decoded.Slice(18, 2)),
                    new string(decoded.Slice(16, 2)),
                    new string(decoded.Slice(30, 2)),
                    new string(decoded.Slice(28, 2)),
                    new string(decoded.Slice(26, 2)),
                    new string(decoded.Slice(24, 2)));
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void ReverseToSpan(ReadOnlySpan<char> source, Span<char> destination)
        {
            for (int i = 0; i < source.Length; i++)
                destination[i] = source[source.Length - 1 - i];
        }

        private static string Reverse(string s)
        {
            char[] arr = s.ToCharArray();
            Array.Reverse(arr);
            return new string(arr);
        }
    }

    #endregion

    #region Registry Operations

    public class RegistryHelper
    {
        private readonly bool _is64Bit;

        public RegistryHelper(bool is64Bit)
        {
            _is64Bit = is64Bit;
        }

        public bool KeyExists(RegistryHiveType hive, string subKey)
        {
            var hiveKey = GetHiveKey(hive);
            return KeyExistsInternal(hiveKey, subKey) ||
                   (_is64Bit && KeyExistsInternal(hiveKey, GetWow64Key(subKey)));
        }

        public bool DeleteKey(RegistryHiveType hive, string subKey, bool recursive = true)
        {
            var hiveKey = GetHiveKey(hive);
            bool result = false;

            if (KeyExists(hive, subKey))
            {
                result = DeleteKeyInternal(hiveKey, subKey, recursive);
            }

            if (_is64Bit && KeyExists(hive, GetWow64Key(subKey)))
            {
                result = DeleteKeyInternal(hiveKey, GetWow64Key(subKey), recursive) || result;
            }

            return result;
        }

        public bool DeleteValue(RegistryHiveType hive, string subKey, string valueName)
        {
            bool result = false;

            try
            {
                using (var key = OpenKey(GetHiveKey(hive), subKey, true))
                {
                    if (key != null && key.GetValue(valueName) != null)
                    {
                        key.DeleteValue(valueName, false);
                        result = true;
                    }
                }
            }
            catch { }

            if (_is64Bit)
            {
                try
                {
                    using (var key = OpenKey(GetHiveKey(hive), GetWow64Key(subKey), true))
                    {
                        if (key != null && key.GetValue(valueName) != null)
                        {
                            key.DeleteValue(valueName, false);
                            result = true;
                        }
                    }
                }
                catch { }
            }

            return result;
        }

        public string[] EnumerateKeys(RegistryHiveType hive, string subKey)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var hiveKey = GetHiveKey(hive);

            AddKeysFromPath(hiveKey, subKey, keys);
            if (_is64Bit)
                AddKeysFromPath(hiveKey, GetWow64Key(subKey), keys);

            return keys.ToArray();
        }

        private void AddKeysFromPath(RegistryKey hiveKey, string subKey, HashSet<string> keys)
        {
            try
            {
                using var key = hiveKey.OpenSubKey(subKey, false);
                if (key != null)
                {
                    foreach (var name in key.GetSubKeyNames())
                        keys.Add(name);
                }
            }
            catch { }
        }

        public string[] EnumerateValues(RegistryHiveType hive, string subKey)
        {
            var values = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var hiveKey = GetHiveKey(hive);

            AddValuesFromPath(hiveKey, subKey, values);
            if (_is64Bit)
                AddValuesFromPath(hiveKey, GetWow64Key(subKey), values);

            return values.ToArray();
        }

        private void AddValuesFromPath(RegistryKey hiveKey, string subKey, HashSet<string> values)
        {
            try
            {
                using var key = hiveKey.OpenSubKey(subKey, false);
                if (key != null)
                {
                    foreach (var name in key.GetValueNames())
                        values.Add(name);
                }
            }
            catch { }
        }

        public object GetValue(RegistryHiveType hive, string subKey, string valueName, object defaultValue = null)
        {
            var hiveKey = GetHiveKey(hive);
            object result = GetValueFromPath(hiveKey, subKey, valueName, defaultValue);

            if (result == null && _is64Bit)
                result = GetValueFromPath(hiveKey, GetWow64Key(subKey), valueName, defaultValue);

            return result ?? defaultValue;
        }

        private object GetValueFromPath(RegistryKey hiveKey, string subKey, string valueName, object defaultValue)
        {
            try
            {
                using var key = hiveKey.OpenSubKey(subKey, false);
                return key?.GetValue(valueName, defaultValue);
            }
            catch
            {
                return null;
            }
        }

        public bool SetValue(RegistryHiveType hive, string subKey, string valueName, object value, RegistryValueKind kind)
        {
            try
            {
                using (var key = OpenKey(GetHiveKey(hive), subKey, true) ??
                               GetHiveKey(hive).CreateSubKey(subKey, true))
                {
                    if (key != null)
                    {
                        key.SetValue(valueName, value, kind);
                        return true;
                    }
                }
            }
            catch { }

            return false;
        }

        private bool KeyExistsInternal(RegistryKey hive, string subKey)
        {
            try
            {
                using var key = hive.OpenSubKey(subKey, false);
                return key != null;
            }
            catch
            {
                return false;
            }
        }

        private bool DeleteKeyInternal(RegistryKey hive, string subKey, bool recursive)
        {
            try
            {
                if (recursive)
                    hive.DeleteSubKeyTree(subKey, false);
                else
                    hive.DeleteSubKey(subKey, false);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private RegistryKey OpenKey(RegistryKey hive, string subKey, bool writable)
        {
            try
            {
                return hive.OpenSubKey(subKey, writable);
            }
            catch
            {
                return null;
            }
        }

        // OPTIMIZATION: Use switch expression
        private RegistryKey GetHiveKey(RegistryHiveType hive) => hive switch
        {
            RegistryHiveType.ClassesRoot => Registry.ClassesRoot,
            RegistryHiveType.CurrentUser => Registry.CurrentUser,
            RegistryHiveType.LocalMachine => Registry.LocalMachine,
            RegistryHiveType.Users => Registry.Users,
            _ => throw new ArgumentException("Invalid hive type")
        };

        private string GetWow64Key(string subKey)
        {
            if (subKey.StartsWith("Software\\Classes\\", StringComparison.OrdinalIgnoreCase))
                return "Software\\Classes\\Wow6432Node\\" + subKey.Substring(18);

            int pos = subKey.IndexOf('\\');
            if (pos > 0)
                return subKey.Substring(0, pos) + "\\Wow6432Node\\" + subKey.Substring(pos + 1);

            return "Wow6432Node\\" + subKey;
        }

        public void AddPendingFileRenameOperation(string filePath)
        {
            const string keyPath = @"SYSTEM\CurrentControlSet\Control\Session Manager";
            const string valueName = "PendingFileRenameOperations";

            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(keyPath, true))
                {
                    if (key != null)
                    {
                        var existing = key.GetValue(valueName) as string[];
                        var list = new List<string>(existing ?? Array.Empty<string>());
                        list.Add("\\??\\" + filePath);
                        list.Add(string.Empty);
                        key.SetValue(valueName, list.ToArray(), RegistryValueKind.MultiString);
                    }
                }
            }
            catch { }
        }
    }

    #endregion

    #region File Operations

    public class FileHelper
    {
        private readonly List<string> _pendingDeletes = new List<string>();

        public bool DeleteFile(string filePath, bool scheduleOnFail = true)
        {
            if (!File.Exists(filePath))
                return true;

            try
            {
                var attrs = File.GetAttributes(filePath);
                if ((attrs & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                {
                    File.SetAttributes(filePath, attrs & ~FileAttributes.ReadOnly);
                }

                File.Delete(filePath);
                return true;
            }
            catch
            {
                if (scheduleOnFail)
                {
                    ScheduleDeleteOnReboot(filePath);
                }
                return false;
            }
        }

        public bool DeleteDirectory(string directoryPath, bool recursive = true, bool scheduleOnFail = true)
        {
            if (!Directory.Exists(directoryPath))
                return true;

            try
            {
                // Try using cmd.exe rd first for performance
                using (var process = Process.Start(new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/c rd /s /q \"{directoryPath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                }))
                {
                    process.WaitForExit(30000);
                }

                if (!Directory.Exists(directoryPath))
                    return true;

                // Fallback to .NET method
                Directory.Delete(directoryPath, recursive);
                return true;
            }
            catch
            {
                if (scheduleOnFail)
                {
                    ScheduleDirectoryDeleteOnReboot(directoryPath);
                }
                return false;
            }
        }

        public void ScheduleDeleteOnReboot(string filePath)
        {
            try
            {
                NativeMethods.MoveFileEx(filePath, null, NativeMethods.MOVEFILE_DELAY_UNTIL_REBOOT);
                _pendingDeletes.Add(filePath);
            }
            catch { }
        }

        public void ScheduleDirectoryDeleteOnReboot(string directoryPath)
        {
            try
            {
                if (Directory.Exists(directoryPath))
                {
                    // Schedule all files first
                    foreach (var file in Directory.GetFiles(directoryPath, "*", SearchOption.AllDirectories))
                    {
                        ScheduleDeleteOnReboot(file);
                    }

                    // Then schedule directories
                    var dirs = Directory.GetDirectories(directoryPath, "*", SearchOption.AllDirectories)
                                       .OrderByDescending(d => d.Length)
                                       .ToList();

                    foreach (var dir in dirs)
                    {
                        NativeMethods.MoveFileEx(dir, null, NativeMethods.MOVEFILE_DELAY_UNTIL_REBOOT);
                        _pendingDeletes.Add(dir);
                    }

                    NativeMethods.MoveFileEx(directoryPath, null, NativeMethods.MOVEFILE_DELAY_UNTIL_REBOOT);
                    _pendingDeletes.Add(directoryPath);
                }
            }
            catch { }
        }

        public List<string> GetPendingDeletes()
        {
            return new List<string>(_pendingDeletes);
        }

        public long GetDirectorySize(string directoryPath)
        {
            if (!Directory.Exists(directoryPath))
                return 0;

            try
            {
                return Directory.GetFiles(directoryPath, "*", SearchOption.AllDirectories)
                               .Sum(f => new FileInfo(f).Length);
            }
            catch
            {
                return 0;
            }
        }

        public bool IsFileLocked(string filePath)
        {
            if (!File.Exists(filePath))
                return false;

            try
            {
                using (File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    return false;
                }
            }
            catch
            {
                return true;
            }
        }
    }

    #endregion

    #region Process Operations

    public class ProcessHelper
    {
        public List<int> TerminateProcesses(string[] processNames, int timeoutMs = 5000)
        {
            var terminatedPids = new List<int>();
            var tasks = new List<Task>();

            foreach (var processName in processNames)
            {
                var nameWithoutExt = Path.GetFileNameWithoutExtension(processName);
                var processes = Process.GetProcessesByName(nameWithoutExt);

                foreach (var process in processes)
                {
                    var pid = process.Id;
                    tasks.Add(Task.Run(() =>
                    {
                        try
                        {
                            if (!process.HasExited)
                            {
                                process.Kill();
                                if (process.WaitForExit(timeoutMs))
                                {
                                    lock (terminatedPids)
                                    {
                                        terminatedPids.Add(pid);
                                    }
                                }
                            }
                        }
                        catch { }
                        finally
                        {
                            process.Dispose();
                        }
                    }));
                }
            }

            if (tasks.Count > 0)
                Task.WaitAll(tasks.ToArray(), timeoutMs * 2);

            return terminatedPids;
        }

        public List<string> GetProcessesUsingPath(string path)
        {
            var processes = new List<string>();

            try
            {
                foreach (var process in Process.GetProcesses())
                {
                    try
                    {
                        if (process.MainModule != null &&
                            process.MainModule.FileName.StartsWith(path, StringComparison.OrdinalIgnoreCase))
                        {
                            processes.Add(process.ProcessName);
                        }
                    }
                    catch { }
                    finally
                    {
                        process.Dispose();
                    }
                }
            }
            catch { }

            return processes;
        }

        // OPTIMIZATION: Proper disposal of Process objects
        public bool IsProcessRunning(string processName)
        {
            try
            {
                var processes = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(processName));
                var hasRunning = processes.Length > 0;
                foreach (var p in processes)
                    p.Dispose();
                return hasRunning;
            }
            catch
            {
                return false;
            }
        }
    }

    #endregion

    #region Shell Integration

    public class ShellHelper
    {
        public bool UnpinFromTaskbar(string shortcutPath)
        {
            try
            {
                if (!File.Exists(shortcutPath))
                    return false;

                var shell = Type.GetTypeFromProgID("Shell.Application");
                dynamic shellApp = Activator.CreateInstance(shell);

                var file = new FileInfo(shortcutPath);
                var folder = shellApp.NameSpace(file.DirectoryName);
                var item = folder.ParseName(file.Name);

                foreach (dynamic verb in item.Verbs())
                {
                    var verbName = verb.Name.Replace("&", "").ToLower();

                    // English and common localized versions
                    if (verbName.Contains("unpin from taskbar") ||
                        verbName.Contains("von taskleiste lösen") ||
                        verbName.Contains("détacher de la barre des tâches") ||
                        verbName.Contains("desanclar de la barra de tareas") ||
                        verbName.Contains("ta bort från aktivitetsfältet") ||
                        verbName.Contains("frigør fra proceslinje") ||
                        verbName.Contains("odepnout z hlavního panelu") ||
                        verbName.Contains("van de taakbalk losmaken") ||
                        verbName.Contains("poista kiinnitys tehtäväpalkista") ||
                        verbName.Contains("rimuovi dalla barra delle applicazioni"))
                    {
                        verb.DoIt();
                        Thread.Sleep(100);
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool UnpinFromStartMenu(string shortcutPath)
        {
            try
            {
                if (!File.Exists(shortcutPath))
                    return false;

                var shell = Type.GetTypeFromProgID("Shell.Application");
                dynamic shellApp = Activator.CreateInstance(shell);

                var file = new FileInfo(shortcutPath);
                var folder = shellApp.NameSpace(file.DirectoryName);
                var item = folder.ParseName(file.Name);

                foreach (dynamic verb in item.Verbs())
                {
                    var verbName = verb.Name.Replace("&", "").ToLower();

                    if (verbName.Contains("unpin from start") ||
                        verbName.Contains("vom startmenü lösen") ||
                        verbName.Contains("détacher du menu démarrer") ||
                        verbName.Contains("desanclar del menú inicio") ||
                        verbName.Contains("odepnout z nabídky start") ||
                        verbName.Contains("frigør fra menuen start") ||
                        verbName.Contains("van het menu start losmaken") ||
                        verbName.Contains("poista kiinnitys käynnistä-valikosta") ||
                        verbName.Contains("irrota aloitusvalikosta"))
                    {
                        verb.DoIt();
                        Thread.Sleep(100);
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        public void RestartExplorer()
        {
            try
            {
                var explorerProcesses = Process.GetProcessesByName("explorer");
                foreach (var process in explorerProcesses)
                {
                    try
                    {
                        process.Kill();
                        process.WaitForExit(5000);
                    }
                    catch { }
                    finally
                    {
                        process.Dispose();
                    }
                }

                Thread.Sleep(1000);
                Process.Start("explorer.exe");
            }
            catch { }
        }
    }

    #endregion

    #region Windows Installer Helper

    public class WindowsInstallerHelper
    {
        private readonly RegistryHelper _regHelper;

        public WindowsInstallerHelper(RegistryHelper regHelper)
        {
            _regHelper = regHelper;
        }

        public void CleanupUpgradeCodes(Func<string, bool> shouldDelete)
        {
            const string path = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UpgradeCodes";

            var keys = _regHelper.EnumerateKeys(RegistryHiveType.LocalMachine, path);
            foreach (var key in keys)
            {
                if (key.Length == 32)
                {
                    var guid = GuidHelper.GetExpandedGuid(key);
                    if (guid != null && shouldDelete(guid))
                    {
                        _regHelper.DeleteKey(RegistryHiveType.LocalMachine, $"{path}\\{key}");
                    }
                }
            }
        }

        public void CleanupProducts(Func<string, bool> shouldDelete)
        {
            var productPaths = new[]
            {
                ("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products", RegistryHiveType.LocalMachine),
                ("Installer\\Products", RegistryHiveType.ClassesRoot)
            };

            foreach (var (path, hive) in productPaths)
            {
                var keys = _regHelper.EnumerateKeys(hive, path);
                foreach (var key in keys)
                {
                    if (key.Length == 32)
                    {
                        var guid = GuidHelper.GetExpandedGuid(key);
                        if (guid != null && shouldDelete(guid))
                        {
                            _regHelper.DeleteKey(hive, $"{path}\\{key}");
                        }
                    }
                }
            }
        }

        public void CleanupComponents(Func<string, bool> shouldDelete)
        {
            const string componentPath = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Components";

            var components = _regHelper.EnumerateKeys(RegistryHiveType.LocalMachine, componentPath);
            foreach (var component in components)
            {
                if (component.Length == 32)
                {
                    var values = _regHelper.EnumerateValues(RegistryHiveType.LocalMachine,
                        $"{componentPath}\\{component}");

                    foreach (var value in values)
                    {
                        if (value.Length == 32)
                        {
                            var guid = GuidHelper.GetExpandedGuid(value);
                            if (guid != null && shouldDelete(guid))
                            {
                                _regHelper.DeleteValue(RegistryHiveType.LocalMachine,
                                    $"{componentPath}\\{component}", value);
                            }
                        }
                    }
                }
            }
        }

        public void CleanupPublishedComponents(Func<string, bool> shouldDelete)
        {
            const string componentPath = "Installer\\Components";

            var components = _regHelper.EnumerateKeys(RegistryHiveType.ClassesRoot, componentPath);
            foreach (var component in components)
            {
                if (component.Length == 32)
                {
                    var values = _regHelper.EnumerateValues(RegistryHiveType.ClassesRoot,
                        $"{componentPath}\\{component}");

                    foreach (var value in values)
                    {
                        var data = _regHelper.GetValue(RegistryHiveType.ClassesRoot,
                            $"{componentPath}\\{component}", value) as string[];

                        if (data != null)
                        {
                            bool modified = false;
                            var newData = new List<string>();

                            foreach (var item in data)
                            {
                                if (item.Length > 20)
                                {
                                    var encoded = item.Substring(0, 20);
                                    if (GuidHelper.GetDecodedGuid(encoded, out string guid))
                                    {
                                        if (shouldDelete(guid))
                                        {
                                            modified = true;
                                            continue;
                                        }
                                    }
                                }
                                newData.Add(item);
                            }

                            if (modified)
                            {
                                if (newData.Count == 0)
                                {
                                    _regHelper.DeleteValue(RegistryHiveType.ClassesRoot,
                                        $"{componentPath}\\{component}", value);
                                }
                                else
                                {
                                    _regHelper.SetValue(RegistryHiveType.ClassesRoot,
                                        $"{componentPath}\\{component}", value,
                                        newData.ToArray(), RegistryValueKind.MultiString);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    #endregion

    #region TypeLib Helper

    public class TypeLibHelper
    {
        private readonly RegistryHelper _regHelper;

        // Known Office TypeLibs
        private static readonly string[] KnownTypeLibs = new[]
        {
            "{000204EF-0000-0000-C000-000000000046}", "{00020802-0000-0000-C000-000000000046}",
            "{00020813-0000-0000-C000-000000000046}", "{00020905-0000-0000-C000-000000000046}",
            "{0002123C-0000-0000-C000-000000000046}", "{00024517-0000-0000-C000-000000000046}",
            "{0002E157-0000-0000-C000-000000000046}", "{00062FFF-0000-0000-C000-000000000046}",
            "{0006F062-0000-0000-C000-000000000046}", "{0006F080-0000-0000-C000-000000000046}",
            "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", "{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}",
            "{5B87B6F0-17C8-11D0-AD41-00A0C90DC8D9}", "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}",
            "{91493440-5A91-11CF-8700-00AA0060263B}", "{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}",
            "{BDEADE33-C265-11D0-BCED-00A0C90AB50F}", "{BDEADEF0-C265-11D0-BCED-00A0C90AB50F}",
            "{EDCD5812-6A06-43C3-AFAC-46EF5D14E22C}"
        };

        public TypeLibHelper(RegistryHelper regHelper)
        {
            _regHelper = regHelper;
        }

        public void CleanupKnownTypeLibs()
        {
            const string typeLibPath = "Software\\Classes\\TypeLib";

            foreach (var typeLib in KnownTypeLibs)
            {
                var tlKey = $"{typeLibPath}\\{typeLib}";
                if (!_regHelper.KeyExists(RegistryHiveType.LocalMachine, tlKey))
                    continue;

                var versions = _regHelper.EnumerateKeys(RegistryHiveType.LocalMachine, tlKey);

                foreach (var version in versions)
                {
                    bool canDelete = true;
                    var versionKey = $"{tlKey}\\{version}";

                    // Check Win32 and Win64 paths
                    foreach (var platform in new[] { "0\\Win32", "9\\Win32", "0\\Win64", "9\\Win64" })
                    {
                        var platformKey = $"{versionKey}\\{platform}";
                        var filePath = _regHelper.GetValue(RegistryHiveType.LocalMachine,
                            platformKey, "", null) as string;

                        if (!string.IsNullOrEmpty(filePath))
                        {
                            var safePath = filePath.Substring(0,
                                Math.Min(filePath.LastIndexOf('.') + 4, filePath.Length));

                            if (File.Exists(safePath))
                            {
                                canDelete = false;
                                break;
                            }
                        }
                    }

                    if (canDelete)
                    {
                        _regHelper.DeleteKey(RegistryHiveType.LocalMachine, versionKey);
                    }
                }

                // If all versions removed, remove the TypeLib key
                versions = _regHelper.EnumerateKeys(RegistryHiveType.LocalMachine, tlKey);
                if (versions.Length == 0)
                {
                    _regHelper.DeleteKey(RegistryHiveType.LocalMachine, tlKey);
                }
            }
        }
    }

    #endregion

    #region License/SPP Helper

    public class LicenseHelper
    {
        public void CleanOSPP(int versionNT)
        {
            const string officeAppId = "0ff1ce15-a989-479d-af46-f275c6370663";

            try
            {
                var scope = new ManagementScope("\\\\.\\root\\cimv2");
                scope.Connect();

                var className = versionNT > 601 ? "SoftwareLicensingProduct" : "OfficeSoftwareProtectionProduct";
                var queryString = $"SELECT ProductKeyID FROM {className} WHERE ApplicationId = '{officeAppId}' AND PartialProductKey <> NULL";

                using (var searcher = new ManagementObjectSearcher(scope, new SelectQuery(queryString)))
                {
                    var products = searcher.Get();

                    foreach (ManagementObject product in products)
                    {
                        try
                        {
                            var productKeyID = product["ProductKeyID"];
                            if (productKeyID != null)
                            {
                                product.InvokeMethod("UninstallProductKey", new object[] { productKeyID });
                            }
                        }
                        catch { }
                        finally
                        {
                            product.Dispose();
                        }
                    }
                }
            }
            catch { }
        }

        public void ClearVNextLicenseCache(string localAppData)
        {
            var licensePaths = new[]
            {
                Path.Combine(localAppData, "Microsoft", "Office", "Licenses"),
                Path.Combine(localAppData, "Microsoft", "Office", "15.0", "Licensing"),
                Path.Combine(localAppData, "Microsoft", "Office", "16.0", "Licensing")
            };

            var fileHelper = new FileHelper();
            foreach (var path in licensePaths)
            {
                if (Directory.Exists(path))
                {
                    fileHelper.DeleteDirectory(path, true, true);
                }
            }
        }
    }

    #endregion

    #region Service Helper

    public class ServiceHelper
    {
        public bool DeleteService(string serviceName)
        {
            try
            {
                var scope = new ManagementScope("\\\\.\\root\\cimv2");
                scope.Connect();

                var queryString = $"SELECT * FROM Win32_Service WHERE Name LIKE '{serviceName}%'";
                var query = new SelectQuery(queryString);

                using (var searcher = new ManagementObjectSearcher(scope, query))
                {
                    foreach (ManagementObject service in searcher.Get())
                    {
                        try
                        {
                            // Try to stop the service
                            var state = service["State"] as string;
                            if (state == "Running" || state == "Started")
                            {
                                service.InvokeMethod("StopService", null);
                                Thread.Sleep(1000);
                            }

                            // Delete the service
                            service.InvokeMethod("Delete", null);
                        }
                        catch { }
                        finally
                        {
                            service.Dispose();
                        }
                    }
                }

                return true;
            }
            catch
            {
                // Fallback to sc.exe
                try
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = "sc.exe",
                        Arguments = $"delete {serviceName}",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };

                    using (var process = Process.Start(psi))
                    {
                        process.WaitForExit(5000);
                        return process.ExitCode == 0;
                    }
                }
                catch
                {
                    return false;
                }
            }
        }
    }

    #endregion

    #region Orchestrator

    public class OfficeScrubOrchestrator
    {
        // OPTIMIZATION: Static collections for validation
        private static readonly HashSet<string> ValidSkus = new HashSet<string>
        {
            "007E", "008F", "008C", "24E1", "237A", "00DD"
        };

        private static readonly HashSet<string> SpecialProducts = new HashSet<string>
        {
            "{6C1ADE97-24E1-4AE4-AEDD-86D3A209CE60}",
            "{9520DDEB-237A-41DB-AA20-F2EF2360DCEB}",
            "{9AC08E99-230B-47E8-9721-4577B7F124EA}"
        };

        public RegistryHelper Registry { get; private set; }
        public FileHelper Files { get; private set; }
        public ProcessHelper Processes { get; private set; }
        public ShellHelper Shell { get; private set; }
        public WindowsInstallerHelper WindowsInstaller { get; private set; }
        public TypeLibHelper TypeLib { get; private set; }
        public LicenseHelper License { get; private set; }
        public ServiceHelper Services { get; private set; }

        public OfficeScrubOrchestrator(bool is64Bit)
        {
            Registry = new RegistryHelper(is64Bit);
            Files = new FileHelper();
            Processes = new ProcessHelper();
            Shell = new ShellHelper();
            WindowsInstaller = new WindowsInstallerHelper(Registry);
            TypeLib = new TypeLibHelper(Registry);
            License = new LicenseHelper();
            Services = new ServiceHelper();
        }

        // OPTIMIZATION: Use StringComparison.OrdinalIgnoreCase instead of .ToLower()
        public bool IsC2RPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;

            return OfficeConstants.C2R_PATTERNS.Any(pattern =>
                path.Contains(pattern, StringComparison.OrdinalIgnoreCase));
        }

        public bool IsInScope(string productCode)
        {
            if (string.IsNullOrEmpty(productCode) || productCode.Length != 38)
                return false;

            var upper = productCode.ToUpperInvariant();

            // OPTIMIZATION: Check special products first with HashSet O(1) lookup
            if (SpecialProducts.Contains(upper))
                return true;

            if (!upper.EndsWith(OfficeConstants.OFFICE_ID))
                return false;

            // Check version
            if (!int.TryParse(upper.Substring(3, 2), out int version) || version <= 14)
                return false;

            // Check SKU - OPTIMIZATION: HashSet O(1) lookup instead of array Contains O(n)
            var sku = upper.Substring(10, 4);
            return ValidSkus.Contains(sku);
        }
    }

    #endregion
}

