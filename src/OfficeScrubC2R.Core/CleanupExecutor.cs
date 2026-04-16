using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace OfficeScrubC2R
{
    public sealed class CleanupExecutor
    {
        private const int MoveFileDelayUntilReboot = 0x4;

        private static readonly string[] OfficeScheduledTasks =
        {
            "FF_INTEGRATEDstreamSchedule",
            "FF_INTEGRATEDUPDATEDETECTION",
            "C2RAppVLoggingStart",
            "Office 15 Subscription Heartbeat",
            "Microsoft Office 15 Sync Maintenance for {d068b555-9700-40b8-992c-f866287b06c1}",
            "OfficeInventoryAgentFallBack",
            @"\Microsoft\Office\OfficeInventoryAgentFallBack",
            @"\Microsoft\Office\OfficeTelemetryAgentFallBack",
            @"\Microsoft\Office\OfficeInventoryAgentLogOn",
            @"\Microsoft\Office\OfficeTelemetryAgentLogOn",
            "Office Background Streaming",
            @"\Microsoft\Office\Office Automatic Updates",
            @"\Microsoft\Office\Office ClickToRun Service Monitor",
            "Office Subscription Maintenance"
        };

        private static readonly RegistryTarget[] BuiltInRegistryTargets =
        {
            new RegistryTarget(RegistryHive.CurrentUser, @"SOFTWARE\Microsoft\Office\15.0\ClickToRun"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Office\15.0\ClickToRun"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Office\15.0\ClickToRunStore"),
            new RegistryTarget(RegistryHive.CurrentUser, @"SOFTWARE\Microsoft\Office\16.0\ClickToRun"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Office\16.0\ClickToRun"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Office\16.0\ClickToRunStore"),
            new RegistryTarget(RegistryHive.CurrentUser, @"SOFTWARE\Microsoft\Office\ClickToRun"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Office\ClickToRun"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Office\ClickToRunStore"),
            new RegistryTarget(RegistryHive.CurrentUser, @"Software\Microsoft\Office\15.0\Registration"),
            new RegistryTarget(RegistryHive.CurrentUser, @"Software\Microsoft\Office\16.0\Registration"),
            new RegistryTarget(RegistryHive.CurrentUser, @"Software\Microsoft\Office\Registration"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot\Virtual"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot\Virtual"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Office\Common\InstallRoot\Virtual"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Classes\Protocols\Handler\osf"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Classes\CLSID\{2027FC3B-CF9D-4ec7-A823-38BA308625CC}"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Classes\CLSID\{573FFD05-2805-47C2-BCE0-5F19512BEB8D}"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Classes\CLSID\{8BA85C75-763B-4103-94EB-9470F12FE0F7}"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Classes\CLSID\{CD55129A-B1A1-438E-A425-CEBC7DC684EE}"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Classes\CLSID\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Classes\CLSID\{E768CD3B-BDDC-436D-9C13-E1B39CA257B1}"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 1 (ErrorConflict)"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 2 (SyncInProgress)"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 3 (InSync)"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\Namespace\{B28AA736-876B-46DA-B3A8-84C5E30BA492}"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\NetworkNeighborhood\Namespace\{46137B78-0EC3-426D-8B89-FF7C3A458B5E}"),
            new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Microsoft Office Temp Files")
        };

        private static readonly RegistryValueTarget[] BuiltInRegistryValueTargets =
        {
            new RegistryValueTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Lync15"),
            new RegistryValueTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Lync16"),
            new RegistryValueTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", "{B28AA736-876B-46DA-B3A8-84C5E30BA492}"),
            new RegistryValueTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", "{8B02D659-EBBB-43D7-9BBA-52CF22C5B025}"),
            new RegistryValueTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", "{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}"),
            new RegistryValueTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", "{42042206-2D85-11D3-8CFF-005004838597}"),
            new RegistryValueTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", "{993BE281-6695-4BA5-8A2A-7AACBFAAB69E}"),
            new RegistryValueTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", "{C41662BB-1FA0-4CE0-8DC5-9B7F8279FF97}"),
            new RegistryValueTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", "{506F4668-F13E-4AA1-BB04-B43203AB3CC0}"),
            new RegistryValueTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", "{D66DC78C-4F61-447F-942B-3FB6980118CF}"),
            new RegistryValueTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", "{46137B78-0EC3-426D-8B89-FF7C3A458B5E}"),
            new RegistryValueTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", "{8BA85C75-763B-4103-94EB-9470F12FE0F7}"),
            new RegistryValueTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", "{CD55129A-B1A1-438E-A425-CEBC7DC684EE}"),
            new RegistryValueTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", "{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}"),
            new RegistryValueTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", "{E768CD3B-BDDC-436D-9C13-E1B39CA257B1}"),
            new RegistryValueTarget(RegistryHive.CurrentUser, @"SOFTWARE\Microsoft\Office\15.0\CleanC2R", "Rerun")
        };

        public ScrubPlan Execute(ScrubExecutionRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            var state = request.State ?? new OfficeC2RState();
            var plan = ScrubPlanner.CreatePlan(state, request.KeepLicense, planOnly: false);
            plan.ExecutionStatus = "Running";
            plan.Message = "Executing Office Click-to-Run cleanup.";

            if (!state.IsElevated)
            {
                plan.ExecutionStatus = "Blocked";
                plan.Message = "Administrator privileges are required for destructive cleanup.";
                plan.ExecutedOperations.Add(OperationResult.Blocked(
                    "Preflight",
                    "RequireElevation",
                    "Privilege",
                    "Administrator",
                    plan.Message,
                    "OfficeScrubC2R.AdminRequired"));
                return plan;
            }

            foreach (var operation in ExecuteOperations(request))
            {
                plan.ExecutedOperations.Add(operation);
            }

            plan.ExecutionStatus = plan.ExecutedOperations.Any(item => item.Status == OperationStatus.Failed || item.Status == OperationStatus.Blocked)
                ? "CompletedWithFailures"
                : "Completed";
            plan.Message = "Office Click-to-Run cleanup execution finished. Review ExecutedOperations for details.";
            return plan;
        }

        private IEnumerable<OperationResult> ExecuteOperations(ScrubExecutionRequest request)
        {
            if (!request.SkipBuiltInTargets)
            {
                foreach (var result in TerminateOfficeProcesses())
                {
                    yield return result;
                }

                foreach (var task in OfficeScheduledTasks)
                {
                    yield return DeleteScheduledTask(task);
                }

                foreach (var serviceName in OfficeConstants.ClickToRunServiceNames.Concat(new[] { "OfficeSvc" }))
                {
                    yield return StopService(serviceName);
                    yield return DeleteService(serviceName);
                }
            }

            var registryTargets = DiscoverRegistryCleanupTargets(request);
            foreach (var diagnostic in registryTargets.Diagnostics)
            {
                yield return diagnostic;
            }

            foreach (var target in registryTargets.KeyTargets)
            {
                foreach (var result in DeleteRegistryKey(target, request.State.Is64BitOperatingSystem))
                {
                    yield return result;
                }
            }

            foreach (var target in registryTargets.ValueTargets)
            {
                foreach (var result in DeleteRegistryValue(target, request.State.Is64BitOperatingSystem))
                {
                    yield return result;
                }
            }

            if (!request.SkipBuiltInTargets)
            {
                foreach (var result in DeleteScopedRunValues(request.State.Is64BitOperatingSystem))
                {
                    yield return result;
                }
            }

            foreach (var path in BuildFileSystemTargets(request))
            {
                yield return DeleteFileSystemPath(path);
            }

            if (!request.KeepLicense)
            {
                foreach (var path in BuildLicenseTargets())
                {
                    yield return DeleteFileSystemPath(path, "Licensing", "DeleteLicenseCache");
                }
            }
            else
            {
                yield return OperationResult.Skipped(
                    "Licensing",
                    "DeleteLicenseCache",
                    "DirectorySet",
                    "Office licensing data",
                    "License cleanup skipped because KeepLicense was requested.");
            }
        }

        private static RegistryCleanupTargets DiscoverRegistryCleanupTargets(ScrubExecutionRequest request)
        {
            var targets = new RegistryCleanupTargets();

            if (!request.SkipBuiltInTargets)
            {
                foreach (var target in BuiltInRegistryTargets)
                {
                    targets.KeyTargets.Add(target);
                }

                foreach (var product in request.State.InstalledProducts.Where(product => product.RegistryHive.HasValue && !string.IsNullOrWhiteSpace(product.RegistryPath)))
                {
                    var hive = product.RegistryHive.GetValueOrDefault();
                    targets.KeyTargets.Add(new RegistryTarget(hive, product.RegistryPath));
                }

                AddScopedInstallerRegistryTargets(targets, request.State.Is64BitOperatingSystem);

                foreach (var target in BuiltInRegistryValueTargets)
                {
                    targets.ValueTargets.Add(target);
                }

                AddScopedInstallerRegistryValueTargets(targets, request.State.Is64BitOperatingSystem);
            }

            foreach (var target in request.ExtraRegistryTargets)
            {
                targets.KeyTargets.Add(target);
            }

            return targets;
        }

        private static IEnumerable<string> BuildFileSystemTargets(ScrubExecutionRequest request)
        {
            var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (!request.SkipBuiltInTargets)
            {
                foreach (var path in request.State.PackagePaths)
                {
                    AddPath(paths, path);
                }

                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\Microsoft Office 15"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\Microsoft Office 16"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\Microsoft Office\PackageManifests"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\Microsoft Office\PackageSunrisePolicies"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\Microsoft Office\root"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\Microsoft Office\Office16"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\Microsoft Office\Office15"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\Microsoft Office\AppXManifest.xml"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\Microsoft Office\FileSystemMetadata.xml"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramFiles(x86)%\Microsoft Office 15"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramFiles(x86)%\Microsoft Office 16"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramFiles(x86)%\Microsoft Office\PackageManifests"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramFiles(x86)%\Microsoft Office\PackageSunrisePolicies"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramFiles(x86)%\Microsoft Office\root"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramFiles(x86)%\Microsoft Office\Office16"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramFiles(x86)%\Microsoft Office\Office15"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramFiles(x86)%\Microsoft Office\AppXManifest.xml"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramFiles(x86)%\Microsoft Office\FileSystemMetadata.xml"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%CommonProgramFiles(x86)%\Microsoft Office 15"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%CommonProgramFiles(x86)%\Microsoft Office 16"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramData%\Microsoft\ClickToRun"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%CommonProgramFiles%\microsoft shared\ClickToRun"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramData%\Microsoft\office\FFPackageLocker"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramData%\Microsoft\office\ClickToRunPackageLocker"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramData%\Microsoft\office\FFStatePBLocker"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%ProgramData%\Microsoft\office\Heartbeat"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Microsoft Office"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Microsoft Office 15"));
                AddPath(paths, Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Microsoft Office 16"));
            }

            foreach (var path in request.ExtraFileSystemTargets)
            {
                AddPath(paths, path);
            }

            return paths;
        }

        private static IEnumerable<string> BuildLicenseTargets()
        {
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (string.IsNullOrWhiteSpace(localAppData))
            {
                yield break;
            }

            yield return Path.Combine(localAppData, "Microsoft", "Office", "Licenses");
            yield return Path.Combine(localAppData, "Microsoft", "Office", "15.0", "Licensing");
            yield return Path.Combine(localAppData, "Microsoft", "Office", "16.0", "Licensing");
        }

        private static IReadOnlyList<OperationResult> TerminateOfficeProcesses()
        {
            var results = new List<OperationResult>();

            foreach (var processName in OfficeConstants.OfficeProcesses)
            {
                Process[] processes;
                try
                {
                    processes = Process.GetProcessesByName(processName);
                }
                catch (Exception exception)
                {
                    results.Add(OperationResult.FromException("Processes", "EnumerateProcess", "Process", processName, exception));
                    continue;
                }

                if (processes.Length == 0)
                {
                    results.Add(OperationResult.Skipped("Processes", "TerminateProcess", "Process", processName, "Process was not running."));
                    continue;
                }

                foreach (var process in processes)
                {
                    var target = process.ProcessName + "#" + process.Id;
                    try
                    {
                        process.Kill();
                        process.WaitForExit(5000);
                        results.Add(OperationResult.Completed("Processes", "TerminateProcess", "Process", target, "Process terminated."));
                    }
                    catch (Exception exception)
                    {
                        results.Add(OperationResult.FromException("Processes", "TerminateProcess", "Process", target, exception));
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }
            }

            return results;
        }

        private static OperationResult DeleteScheduledTask(string taskName)
        {
            var result = RunProcess("schtasks.exe", "/Delete /TN \"" + taskName + "\" /F", 15000);
            if (result.ExitCode == 0)
            {
                return OperationResult.Completed("ScheduledTasks", "DeleteTask", "ScheduledTask", taskName, "Scheduled task deleted.");
            }

            if (result.Output.IndexOf("cannot find", StringComparison.OrdinalIgnoreCase) >= 0 ||
                result.Output.IndexOf("does not exist", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return OperationResult.Skipped("ScheduledTasks", "DeleteTask", "ScheduledTask", taskName, "Scheduled task was not present.");
            }

            return OperationResult.Blocked(
                "ScheduledTasks",
                "DeleteTask",
                "ScheduledTask",
                taskName,
                result.Output,
                "OfficeScrubC2R.ScheduledTaskDeleteFailed");
        }

        private static OperationResult StopService(string serviceName)
        {
            var result = RunProcess("sc.exe", "stop \"" + serviceName + "\"", 15000);
            if (result.ExitCode == 0)
            {
                return OperationResult.Completed("Services", "StopService", "Service", serviceName, "Service stop requested.");
            }

            if (result.Output.IndexOf("does not exist", StringComparison.OrdinalIgnoreCase) >= 0 ||
                result.Output.IndexOf("not been started", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return OperationResult.Skipped("Services", "StopService", "Service", serviceName, "Service was not running or not present.");
            }

            return OperationResult.Blocked("Services", "StopService", "Service", serviceName, result.Output, "OfficeScrubC2R.ServiceStopFailed");
        }

        private static OperationResult DeleteService(string serviceName)
        {
            var result = RunProcess("sc.exe", "delete \"" + serviceName + "\"", 15000);
            if (result.ExitCode == 0)
            {
                return OperationResult.Completed("Services", "DeleteService", "Service", serviceName, "Service delete requested.");
            }

            if (result.Output.IndexOf("does not exist", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return OperationResult.Skipped("Services", "DeleteService", "Service", serviceName, "Service was not present.");
            }

            return OperationResult.Blocked("Services", "DeleteService", "Service", serviceName, result.Output, "OfficeScrubC2R.ServiceDeleteFailed");
        }

        private static IReadOnlyList<OperationResult> DeleteRegistryKey(RegistryTarget target, bool is64BitOperatingSystem)
        {
            var results = new List<OperationResult>();

            foreach (var view in GetRegistryViews(is64BitOperatingSystem))
            {
                var displayTarget = target.Hive + "\\" + target.SubKey;
                try
                {
                    using (var baseKey = RegistryKey.OpenBaseKey(target.Hive, view))
                    using (var existing = baseKey.OpenSubKey(target.SubKey, false))
                    {
                        if (existing == null)
                        {
                            results.Add(OperationResult.Skipped("Registry", "DeleteKey", "RegistryKey", displayTarget, "Registry key was not present.", target.Hive, view));
                            continue;
                        }
                    }

                    using (var baseKey = RegistryKey.OpenBaseKey(target.Hive, view))
                    {
                        baseKey.DeleteSubKeyTree(target.SubKey, false);
                    }

                    results.Add(OperationResult.Completed("Registry", "DeleteKey", "RegistryKey", displayTarget, "Registry key deleted.", target.Hive, view));
                }
                catch (Exception exception)
                {
                    results.Add(OperationResult.FromException("Registry", "DeleteKey", "RegistryKey", displayTarget, exception, target.Hive, view));
                }
            }

            return results;
        }

        private static IReadOnlyList<OperationResult> DeleteRegistryValue(RegistryValueTarget target, bool is64BitOperatingSystem)
        {
            var results = new List<OperationResult>();

            foreach (var view in GetRegistryViews(is64BitOperatingSystem))
            {
                var displayTarget = target.Hive + "\\" + target.SubKey + "\\" + target.ValueName;
                try
                {
                    using (var baseKey = RegistryKey.OpenBaseKey(target.Hive, view))
                    using (var key = baseKey.OpenSubKey(target.SubKey, true))
                    {
                        if (key == null || key.GetValue(target.ValueName) == null)
                        {
                            results.Add(OperationResult.Skipped("Registry", "DeleteValue", "RegistryValue", displayTarget, "Registry value was not present.", target.Hive, view));
                            continue;
                        }

                        key.DeleteValue(target.ValueName, false);
                    }

                    results.Add(OperationResult.Completed("Registry", "DeleteValue", "RegistryValue", displayTarget, "Registry value deleted.", target.Hive, view));
                }
                catch (Exception exception)
                {
                    results.Add(OperationResult.FromException("Registry", "DeleteValue", "RegistryValue", displayTarget, exception, target.Hive, view));
                }
            }

            return results;
        }

        private static IReadOnlyList<OperationResult> DeleteScopedRunValues(bool is64BitOperatingSystem)
        {
            const string runKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
            var results = new List<OperationResult>();

            foreach (var view in GetRegistryViews(is64BitOperatingSystem))
            {
                try
                {
                    using (var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                    using (var key = baseKey.OpenSubKey(runKey, true))
                    {
                        if (key == null)
                        {
                            results.Add(OperationResult.Skipped("Registry", "DeleteScopedRunValue", "RegistryKey", RegistryHive.LocalMachine + "\\" + runKey, "Run key was not present.", RegistryHive.LocalMachine, view));
                            continue;
                        }

                        foreach (var valueName in key.GetValueNames())
                        {
                            var value = key.GetValue(valueName) as string;
                            if (!OfficeScope.IsC2RPath(value ?? string.Empty))
                            {
                                continue;
                            }

                            var displayTarget = RegistryHive.LocalMachine + "\\" + runKey + "\\" + valueName;
                            try
                            {
                                key.DeleteValue(valueName, false);
                                results.Add(OperationResult.Completed("Registry", "DeleteScopedRunValue", "RegistryValue", displayTarget, "Run value referencing Office Click-to-Run was deleted.", RegistryHive.LocalMachine, view));
                            }
                            catch (Exception exception)
                            {
                                results.Add(OperationResult.FromException("Registry", "DeleteScopedRunValue", "RegistryValue", displayTarget, exception, RegistryHive.LocalMachine, view));
                            }
                        }
                    }
                }
                catch (Exception exception)
                {
                    results.Add(OperationResult.FromException("Registry", "DeleteScopedRunValue", "RegistryKey", RegistryHive.LocalMachine + "\\" + runKey, exception, RegistryHive.LocalMachine, view));
                }
            }

            return results;
        }

        private static void AddScopedInstallerRegistryTargets(RegistryCleanupTargets targets, bool is64BitOperatingSystem)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var roots = new[]
            {
                new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
                new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products"),
                new RegistryTarget(RegistryHive.ClassesRoot, @"Installer\Features"),
                new RegistryTarget(RegistryHive.ClassesRoot, @"Installer\Products")
            };

            foreach (var view in GetRegistryViews(is64BitOperatingSystem))
            {
                foreach (var root in roots)
                {
                    try
                    {
                        using (var baseKey = RegistryKey.OpenBaseKey(root.Hive, view))
                        using (var key = baseKey.OpenSubKey(root.SubKey, false))
                        {
                            if (key == null)
                            {
                                continue;
                            }

                            foreach (var child in key.GetSubKeyNames())
                            {
                                var expanded = ExpandPotentialProductCode(child);
                                if (!OfficeScope.IsInScope(expanded))
                                {
                                    continue;
                                }

                                var target = new RegistryTarget(root.Hive, root.SubKey + "\\" + child);
                                if (seen.Add(target.Hive + "\\" + target.SubKey))
                                {
                                    targets.KeyTargets.Add(target);
                                }
                            }
                        }
                    }
                    catch (Exception exception)
                    {
                        targets.Diagnostics.Add(OperationResult.FromException(
                            "Registry",
                            "DiscoverInstallerKeyTargets",
                            "RegistryKey",
                            root.Hive + "\\" + root.SubKey,
                            exception,
                            root.Hive,
                            view));
                    }
                }
            }
        }

        private static void AddScopedInstallerRegistryValueTargets(RegistryCleanupTargets targets, bool is64BitOperatingSystem)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var roots = new[]
            {
                new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes"),
                new RegistryTarget(RegistryHive.ClassesRoot, @"Installer\UpgradeCodes"),
                new RegistryTarget(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components")
            };

            foreach (var view in GetRegistryViews(is64BitOperatingSystem))
            {
                foreach (var root in roots)
                {
                    try
                    {
                        using (var baseKey = RegistryKey.OpenBaseKey(root.Hive, view))
                        using (var key = baseKey.OpenSubKey(root.SubKey, false))
                        {
                            if (key == null)
                            {
                                continue;
                            }

                            foreach (var child in key.GetSubKeyNames())
                            {
                                using (var childKey = key.OpenSubKey(child, false))
                                {
                                    if (childKey == null)
                                    {
                                        continue;
                                    }

                                    foreach (var valueName in childKey.GetValueNames())
                                    {
                                        var expanded = ExpandPotentialProductCode(valueName);
                                        if (!OfficeScope.IsInScope(expanded))
                                        {
                                            continue;
                                        }

                                        var target = new RegistryValueTarget(root.Hive, root.SubKey + "\\" + child, valueName);
                                        if (seen.Add(target.Hive + "\\" + target.SubKey + "\\" + target.ValueName))
                                        {
                                            targets.ValueTargets.Add(target);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception exception)
                    {
                        targets.Diagnostics.Add(OperationResult.FromException(
                            "Registry",
                            "DiscoverInstallerValueTargets",
                            "RegistryKey",
                            root.Hive + "\\" + root.SubKey,
                            exception,
                            root.Hive,
                            view));
                    }
                }
            }
        }

        private static string ExpandPotentialProductCode(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var trimmed = value.Trim();
            if (trimmed.Length == 38 && trimmed[0] == '{' && trimmed[37] == '}')
            {
                return trimmed.ToUpperInvariant();
            }

            return trimmed.Length == 32 ? GuidHelper.GetExpandedGuid(trimmed) : string.Empty;
        }

        private static OperationResult DeleteFileSystemPath(string path, string step = "Files", string action = "DeleteDirectory")
        {
            try
            {
                if (File.Exists(path))
                {
                    ClearFileAttributes(path);
                    File.Delete(path);
                    return OperationResult.Completed(step, "DeleteFile", "File", path, "File deleted.");
                }

                if (!Directory.Exists(path))
                {
                    return OperationResult.Skipped(step, action, "Directory", path, "Directory was not present.");
                }

                ClearDirectoryAttributes(path);
                Directory.Delete(path, true);
                return OperationResult.Completed(step, action, "Directory", path, "Directory deleted.");
            }
            catch (Exception exception)
            {
                if (ScheduleDeleteOnReboot(path))
                {
                    return new OperationResult
                    {
                        Step = step,
                        Action = action,
                        TargetKind = "FileSystemPath",
                        Target = path,
                        Status = OperationStatus.ScheduledForReboot,
                        Message = "Immediate deletion failed; path was scheduled for deletion on reboot.",
                        ExceptionType = exception.GetType().Name,
                        HResult = exception.HResult,
                        RebootScheduled = true
                    };
                }

                return OperationResult.FromException(step, action, "FileSystemPath", path, exception);
            }
        }

        private static void ClearDirectoryAttributes(string path)
        {
            foreach (var directory in Directory.GetDirectories(path, "*", SearchOption.AllDirectories))
            {
                ClearFileAttributes(directory);
            }

            foreach (var file in Directory.GetFiles(path, "*", SearchOption.AllDirectories))
            {
                ClearFileAttributes(file);
            }

            ClearFileAttributes(path);
        }

        private static void ClearFileAttributes(string path)
        {
            File.SetAttributes(path, FileAttributes.Normal);
        }

        private static bool ScheduleDeleteOnReboot(string path)
        {
            var scheduled = false;

            if (Directory.Exists(path))
            {
                foreach (var file in Directory.GetFiles(path, "*", SearchOption.AllDirectories))
                {
                    scheduled = MoveFileEx(file, null!, MoveFileDelayUntilReboot) || scheduled;
                }

                foreach (var directory in Directory.GetDirectories(path, "*", SearchOption.AllDirectories).OrderByDescending(item => item.Length))
                {
                    scheduled = MoveFileEx(directory, null!, MoveFileDelayUntilReboot) || scheduled;
                }
            }

            return MoveFileEx(path, null!, MoveFileDelayUntilReboot) || scheduled;
        }

        private static IEnumerable<RegistryView> GetRegistryViews(bool is64BitOperatingSystem)
        {
            if (is64BitOperatingSystem)
            {
                yield return RegistryView.Registry64;
                yield return RegistryView.Registry32;
                yield break;
            }

            yield return RegistryView.Registry32;
        }

        private static void AddPath(HashSet<string> paths, string path)
        {
            if (!string.IsNullOrWhiteSpace(path) && path.IndexOf('%') < 0)
            {
                paths.Add(path);
            }
        }

        private static ProcessRunResult RunProcess(string fileName, string arguments, int timeoutMilliseconds)
        {
            try
            {
                using (var process = new Process())
                {
                    process.StartInfo = new ProcessStartInfo
                    {
                        FileName = fileName,
                        Arguments = arguments,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true
                    };

                    process.Start();
                    if (!process.WaitForExit(timeoutMilliseconds))
                    {
                        try
                        {
                            process.Kill();
                        }
                        catch (Exception exception)
                        {
                            return new ProcessRunResult(-1, fileName + " timed out. Process kill also failed: " + exception.GetType().Name + ": " + exception.Message);
                        }

                        return new ProcessRunResult(-1, fileName + " timed out.");
                    }

                    var output = (process.StandardOutput.ReadToEnd() + Environment.NewLine + process.StandardError.ReadToEnd()).Trim();
                    return new ProcessRunResult(process.ExitCode, output);
                }
            }
            catch (Exception exception)
            {
                return new ProcessRunResult(-1, exception.GetType().Name + ": " + exception.Message);
            }
        }

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool MoveFileEx(string lpExistingFileName, string lpNewFileName, int dwFlags);

        private sealed class ProcessRunResult
        {
            public ProcessRunResult(int exitCode, string output)
            {
                ExitCode = exitCode;
                Output = output ?? string.Empty;
            }

            public int ExitCode { get; }
            public string Output { get; }
        }

        private sealed class RegistryCleanupTargets
        {
            public List<RegistryTarget> KeyTargets { get; } = new List<RegistryTarget>();
            public List<RegistryValueTarget> ValueTargets { get; } = new List<RegistryValueTarget>();
            public List<OperationResult> Diagnostics { get; } = new List<OperationResult>();
        }
    }
}
