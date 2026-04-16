using System;
using System.Collections.Generic;
using Microsoft.Win32;

namespace OfficeScrubC2R
{
    public sealed class RegistryAccess
    {
        private readonly bool _is64BitOperatingSystem;

        public List<OperationResult> Diagnostics { get; } = new List<OperationResult>();

        public RegistryAccess(bool is64BitOperatingSystem)
        {
            _is64BitOperatingSystem = is64BitOperatingSystem;
        }

        public IReadOnlyList<RegistryView> GetCandidateViews()
        {
            if (_is64BitOperatingSystem)
            {
                return new[] { RegistryView.Registry64, RegistryView.Registry32 };
            }

            return new[] { RegistryView.Registry32 };
        }

        public bool KeyExists(RegistryHive hive, string subKey, RegistryView view)
        {
            try
            {
                using (var baseKey = RegistryKey.OpenBaseKey(hive, view))
                using (var key = baseKey.OpenSubKey(subKey, false))
                {
                    return key != null;
                }
            }
            catch (Exception exception)
            {
                AddDiagnostic("KeyExists", hive, subKey, view, exception);
                return false;
            }
        }

        public string[] GetSubKeyNames(RegistryHive hive, string subKey, RegistryView view)
        {
            try
            {
                using (var baseKey = RegistryKey.OpenBaseKey(hive, view))
                using (var key = baseKey.OpenSubKey(subKey, false))
                {
                    return key == null ? Array.Empty<string>() : key.GetSubKeyNames();
                }
            }
            catch (Exception exception)
            {
                AddDiagnostic("EnumerateSubKeys", hive, subKey, view, exception);
                return Array.Empty<string>();
            }
        }

        public object? GetValue(RegistryHive hive, string subKey, string valueName, RegistryView view)
        {
            try
            {
                using (var baseKey = RegistryKey.OpenBaseKey(hive, view))
                using (var key = baseKey.OpenSubKey(subKey, false))
                {
                    return key == null ? null : key.GetValue(valueName);
                }
            }
            catch (Exception exception)
            {
                AddDiagnostic("GetValue", hive, subKey, view, exception);
                return null;
            }
        }

        public string GetStringValue(RegistryHive hive, string subKey, string valueName, RegistryView view)
        {
            return GetValue(hive, subKey, valueName, view) as string ?? string.Empty;
        }

        private void AddDiagnostic(string action, RegistryHive hive, string subKey, RegistryView view, Exception exception)
        {
            Diagnostics.Add(OperationResult.FromException(
                "Registry",
                action,
                "RegistryKey",
                hive + "\\" + subKey,
                exception,
                hive,
                view));
        }
    }
}
