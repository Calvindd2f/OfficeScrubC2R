using System;
using System.Diagnostics;
using System.Linq;
using System.Security.Principal;
using Microsoft.Win32;

namespace OfficeScrubC2R
{
    public sealed class PreflightService
    {
        private readonly OfficeDetectionService _detectionService;
        private readonly RegistryAccess _registry;

        public PreflightService()
            : this(CreateDefaultRegistry())
        {
        }

        private PreflightService(RegistryAccess registry)
            : this(new OfficeDetectionService(registry), registry)
        {
        }

        public PreflightService(OfficeDetectionService detectionService, RegistryAccess registry)
        {
            _detectionService = detectionService;
            _registry = registry;
        }

        public OfficeC2RState GetState()
        {
            var state = new OfficeC2RState
            {
                Is64BitOperatingSystem = Environment.Is64BitOperatingSystem
            };

            AddIdentityState(state);
            AddProducts(state);
            AddPackagePaths(state);
            AddRunningProcesses(state);
            AddServiceEvidence(state);
            AddPendingRebootDeletes(state);
            AddRegistryDiagnostics(state);

            if (!state.IsElevated)
            {
                state.Issues.Add(OperationResult.Blocked(
                    "Preflight",
                    "CheckElevation",
                    "Privilege",
                    "Administrator",
                    "Administrator privileges are required before destructive Office cleanup can be enabled.",
                    "OfficeScrubC2R.AdminRequired"));
            }

            state.PreflightResults.Add(OperationResult.Completed(
                "Preflight",
                "CollectState",
                "OfficeC2RState",
                "LocalComputer",
                "Collected non-destructive Office Click-to-Run state."));

            return state;
        }

        private static void AddIdentityState(OfficeC2RState state)
        {
            try
            {
                var identity = WindowsIdentity.GetCurrent();
                var principal = new WindowsPrincipal(identity);
                state.IsElevated = principal.IsInRole(WindowsBuiltInRole.Administrator);
                state.IsSystem = identity.User != null && identity.User.IsWellKnown(WellKnownSidType.LocalSystemSid);
            }
            catch (Exception exception)
            {
                state.Issues.Add(OperationResult.FromException(
                    "Preflight",
                    "CheckIdentity",
                    "Privilege",
                    "CurrentUser",
                    exception));
            }
        }

        private void AddProducts(OfficeC2RState state)
        {
            foreach (var product in _detectionService.GetInstalledProducts())
            {
                state.InstalledProducts.Add(product);
            }
        }

        private void AddPackagePaths(OfficeC2RState state)
        {
            foreach (var path in _detectionService.GetPackagePaths())
            {
                state.PackagePaths.Add(path);
            }
        }

        private static void AddRunningProcesses(OfficeC2RState state)
        {
            try
            {
                foreach (var process in Process.GetProcesses())
                {
                    try
                    {
                        if (OfficeConstants.OfficeProcesses.Contains(process.ProcessName, StringComparer.OrdinalIgnoreCase))
                        {
                            state.RunningOfficeProcesses.Add(process.ProcessName);
                        }
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }
            }
            catch (Exception exception)
            {
                state.Issues.Add(OperationResult.FromException(
                    "Preflight",
                    "EnumerateProcesses",
                    "Process",
                    "OfficeProcesses",
                    exception));
            }
        }

        private void AddServiceEvidence(OfficeC2RState state)
        {
            foreach (var serviceName in OfficeConstants.ClickToRunServiceNames)
            {
                var path = @"SYSTEM\CurrentControlSet\Services\" + serviceName;
                if (_registry.KeyExists(RegistryHive.LocalMachine, path, RegistryView.Default))
                {
                    state.ClickToRunServices.Add(serviceName);
                }
            }
        }

        private void AddPendingRebootDeletes(OfficeC2RState state)
        {
            const string keyPath = @"SYSTEM\CurrentControlSet\Control\Session Manager";
            var value = _registry.GetValue(
                RegistryHive.LocalMachine,
                keyPath,
                "PendingFileRenameOperations",
                RegistryView.Default) as string[];

            if (value == null)
            {
                return;
            }

            foreach (var item in value.Where(item =>
                         !string.IsNullOrWhiteSpace(item) &&
                         (item.IndexOf("Office", StringComparison.OrdinalIgnoreCase) >= 0 ||
                          item.IndexOf("ClickToRun", StringComparison.OrdinalIgnoreCase) >= 0)))
            {
                state.PendingRebootDeletes.Add(item);
            }
        }

        private void AddRegistryDiagnostics(OfficeC2RState state)
        {
            foreach (var diagnostic in _registry.Diagnostics)
            {
                state.Issues.Add(diagnostic);
            }
        }

        private static RegistryAccess CreateDefaultRegistry()
        {
            return new RegistryAccess(Environment.Is64BitOperatingSystem);
        }
    }
}
