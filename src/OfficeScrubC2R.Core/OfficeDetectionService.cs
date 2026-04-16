using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Win32;

namespace OfficeScrubC2R
{
    public sealed class OfficeDetectionService
    {
        private static readonly string[] ConfigurationKeys =
        {
            @"SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration",
            @"SOFTWARE\Microsoft\Office\16.0\ClickToRun\Configuration",
            @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
        };

        private readonly RegistryAccess _registry;

        public OfficeDetectionService()
            : this(new RegistryAccess(Environment.Is64BitOperatingSystem))
        {
        }

        public OfficeDetectionService(RegistryAccess registry)
        {
            _registry = registry;
        }

        public IReadOnlyList<OfficeProductInfo> GetInstalledProducts()
        {
            var products = new List<OfficeProductInfo>();

            foreach (var view in _registry.GetCandidateViews())
            {
                ReadConfigurationProducts(products, view);
                ReadArpProducts(products, view);
            }

            return products
                .GroupBy(product => string.Join("|", product.ProductId, product.DisplayName, product.RegistryPath, product.RegistryView))
                .Select(group => group.First())
                .OrderBy(product => product.DisplayName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(product => product.ProductId, StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        public IReadOnlyList<string> GetPackagePaths()
        {
            var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var view in _registry.GetCandidateViews())
            {
                foreach (var key in ConfigurationKeys)
                {
                    AddIfPresent(paths, _registry.GetStringValue(RegistryHive.LocalMachine, key, "PackageFolder", view));
                    AddIfPresent(paths, _registry.GetStringValue(RegistryHive.LocalMachine, key, "ClientFolder", view));
                    AddIfPresent(paths, _registry.GetStringValue(RegistryHive.LocalMachine, key, "InstallationPath", view));
                }
            }

            return paths.OrderBy(path => path, StringComparer.OrdinalIgnoreCase).ToArray();
        }

        private void ReadConfigurationProducts(List<OfficeProductInfo> products, RegistryView view)
        {
            foreach (var key in ConfigurationKeys)
            {
                var releaseIds = _registry.GetStringValue(RegistryHive.LocalMachine, key, "ProductReleaseIds", view);
                if (string.IsNullOrWhiteSpace(releaseIds))
                {
                    releaseIds = _registry.GetStringValue(RegistryHive.LocalMachine, key, "ProductReleaseIDs", view);
                }

                if (string.IsNullOrWhiteSpace(releaseIds))
                {
                    continue;
                }

                var version = FirstNonEmpty(
                    _registry.GetStringValue(RegistryHive.LocalMachine, key, "ClientVersionToReport", view),
                    _registry.GetStringValue(RegistryHive.LocalMachine, key, "VersionToReport", view),
                    _registry.GetStringValue(RegistryHive.LocalMachine, key, "Version", view));

                foreach (var releaseId in SplitReleaseIds(releaseIds))
                {
                    products.Add(new OfficeProductInfo
                    {
                        ProductId = releaseId,
                        DisplayName = releaseId,
                        DisplayVersion = version,
                        Architecture = view == RegistryView.Registry32 ? "x86" : "x64",
                        Source = "ClickToRunConfiguration",
                        RegistryHive = RegistryHive.LocalMachine,
                        RegistryView = view,
                        RegistryPath = key,
                        IsClickToRun = true
                    });
                }
            }
        }

        private void ReadArpProducts(List<OfficeProductInfo> products, RegistryView view)
        {
            const string uninstallRoot = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            foreach (var subKeyName in _registry.GetSubKeyNames(RegistryHive.LocalMachine, uninstallRoot, view))
            {
                var path = uninstallRoot + "\\" + subKeyName;
                var displayName = _registry.GetStringValue(RegistryHive.LocalMachine, path, "DisplayName", view);
                var uninstallString = _registry.GetStringValue(RegistryHive.LocalMachine, path, "UninstallString", view);
                var displayVersion = _registry.GetStringValue(RegistryHive.LocalMachine, path, "DisplayVersion", view);

                if (!LooksLikeC2RProduct(displayName, uninstallString))
                {
                    continue;
                }

                products.Add(new OfficeProductInfo
                {
                    ProductId = subKeyName,
                    DisplayName = string.IsNullOrWhiteSpace(displayName) ? subKeyName : displayName,
                    DisplayVersion = displayVersion,
                    Architecture = view == RegistryView.Registry32 ? "x86" : "x64",
                    Source = "UninstallRegistry",
                    RegistryHive = RegistryHive.LocalMachine,
                    RegistryView = view,
                    RegistryPath = path,
                    UninstallString = uninstallString,
                    IsClickToRun = true
                });
            }
        }

        private static bool LooksLikeC2RProduct(string displayName, string uninstallString)
        {
            var combined = string.Concat(displayName ?? string.Empty, " ", uninstallString ?? string.Empty);
            return combined.IndexOf("OfficeClickToRun.exe", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   combined.IndexOf("Microsoft Office 15", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   combined.IndexOf("Microsoft Office 16", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   OfficeScope.IsC2RPath(combined);
        }

        private static IEnumerable<string> SplitReleaseIds(string value)
        {
            return value
                .Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(item => item.Trim())
                .Where(item => item.Length > 0);
        }

        private static string FirstNonEmpty(params string[] values)
        {
            return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? string.Empty;
        }

        private static void AddIfPresent(HashSet<string> paths, string value)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                paths.Add(value);
            }
        }
    }
}
