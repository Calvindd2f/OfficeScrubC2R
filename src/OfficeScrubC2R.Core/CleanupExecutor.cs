using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;

namespace OfficeScrubC2R
{
    public sealed class CleanupExecutor
    {
        private const int MoveFileDelayUntilReboot = 0x4;
        private const string TeamsMachineWideProductCode = "{731F6BAA-A986-45A4-8936-7C3AAAAA760B}";
        private readonly ICommandRunner _commandRunner;

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

        public CleanupExecutor()
            : this(new ProcessCommandRunner())
        {
        }

        public CleanupExecutor(ICommandRunner commandRunner)
        {
            _commandRunner = commandRunner ?? throw new ArgumentNullException(nameof(commandRunner));
        }

        public ScrubPlan Execute(ScrubExecutionRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            var state = request.State ?? new OfficeC2RState();
            var plan = ScrubPlanner.CreatePlan(
                state,
                request.KeepLicense,
                planOnly: false,
                request.KeepTeams,
                request.KeepCopilot);
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
                foreach (var result in TerminateOfficeProcesses(GetProcessNamesForCleanup(request)))
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

            if (!request.SkipCompanionAppTargets)
            {
                foreach (var result in RemoveCompanionApps(request))
                {
                    yield return result;
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

        private static IEnumerable<string> GetProcessNamesForCleanup(ScrubExecutionRequest request)
        {
            foreach (var processName in OfficeConstants.OfficeProcesses)
            {
                if (request.KeepTeams && IsTeamsProcessName(processName))
                {
                    continue;
                }

                if (request.KeepCopilot && IsCopilotProcessName(processName))
                {
                    continue;
                }

                yield return processName;
            }
        }

        private static bool IsTeamsProcessName(string processName)
        {
            return string.Equals(processName, "teams", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(processName, "ms-teams", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(processName, "msteams", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsCopilotProcessName(string processName)
        {
            return string.Equals(processName, "copilot", StringComparison.OrdinalIgnoreCase);
        }

        private static IReadOnlyList<OperationResult> TerminateOfficeProcesses(IEnumerable<string> processNames)
        {
            var results = new List<OperationResult>();

            foreach (var processName in processNames)
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

        private IEnumerable<OperationResult> RemoveCompanionApps(ScrubExecutionRequest request)
        {
            if (request.KeepTeams)
            {
                yield return OperationResult.Skipped(
                    "CompanionApps",
                    "RemoveTeams",
                    "Application",
                    "Microsoft Teams",
                    "Teams cleanup skipped because KeepTeams was requested.");
            }
            else
            {
                yield return RunPowerShellCleanup(
                    "RemoveTeamsAppxPackage",
                    "AppxPackage",
                    "*MSTeams*",
                    CreateRemoveAppxPackagesScript("*MSTeams*", "Teams AppX/MSIX packages"));

                yield return RunPowerShellCleanup(
                    "RemoveTeamsProvisionedPackage",
                    "ProvisionedAppxPackage",
                    "*MSTeams*",
                    CreateRemoveProvisionedAppxPackagesScript("*MSTeams*", "Teams provisioned AppX/MSIX packages"));

                yield return RemoveTeamsMachineWideInstaller();

                foreach (var result in RemoveTeamsBootstrapper())
                {
                    yield return result;
                }

                if (!request.SkipCompanionProfileTargets)
                {
                    foreach (var result in DeleteTeamsProfileRemnants())
                    {
                        yield return result;
                    }

                    foreach (var result in DeleteTeamsRegistryRemnants(request.State.Is64BitOperatingSystem))
                    {
                        yield return result;
                    }
                }
            }

            if (request.KeepCopilot)
            {
                yield return OperationResult.Skipped(
                    "CompanionApps",
                    "RemoveCopilot",
                    "Application",
                    "Microsoft Copilot",
                    "Copilot cleanup skipped because KeepCopilot was requested.");
            }
            else
            {
                yield return RunPowerShellCleanup(
                    "RemoveCopilotAppxPackage",
                    "AppxPackage",
                    "Microsoft.Copilot",
                    CreateRemoveAppxPackagesScript("Microsoft.Copilot", "Microsoft Copilot AppX packages"));

                yield return RunPowerShellCleanup(
                    "RemoveCopilotProvisionedPackage",
                    "ProvisionedAppxPackage",
                    "Microsoft.Copilot",
                    CreateRemoveProvisionedAppxPackagesScript("Microsoft.Copilot", "Microsoft Copilot provisioned AppX packages"));
            }
        }

        private OperationResult DeleteScheduledTask(string taskName)
        {
            var result = _commandRunner.Run("schtasks.exe", "/Delete /TN \"" + taskName + "\" /F", 15000);
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

        private OperationResult StopService(string serviceName)
        {
            var result = _commandRunner.Run("sc.exe", "stop \"" + serviceName + "\"", 15000);
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

        private OperationResult DeleteService(string serviceName)
        {
            var result = _commandRunner.Run("sc.exe", "delete \"" + serviceName + "\"", 15000);
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

        private OperationResult RunPowerShellCleanup(string action, string targetKind, string target, string script)
        {
            var result = _commandRunner.Run(
                "powershell.exe",
                "-NoProfile -ExecutionPolicy Bypass -EncodedCommand " + EncodePowerShellCommand(script),
                120000);

            return ConvertCommandResult(
                "CompanionApps",
                action,
                targetKind,
                target,
                result,
                "OfficeScrubC2R." + action + "Failed");
        }

        private OperationResult RemoveTeamsMachineWideInstaller()
        {
            var result = _commandRunner.Run(
                "msiexec.exe",
                "/x " + TeamsMachineWideProductCode + " /qn /norestart",
                120000);

            return ConvertCommandResult(
                "CompanionApps",
                "RemoveTeamsMachineWideInstaller",
                "MsiProduct",
                TeamsMachineWideProductCode,
                result,
                "OfficeScrubC2R.TeamsMachineWideUninstallFailed");
        }

        private IEnumerable<OperationResult> RemoveTeamsBootstrapper()
        {
            var bootstrapperPath = FindExecutable("teamsbootstrapper.exe");
            if (string.IsNullOrWhiteSpace(bootstrapperPath))
            {
                yield return OperationResult.Skipped(
                    "CompanionApps",
                    "RemoveTeamsBootstrapperMachineWide",
                    "Executable",
                    "teamsbootstrapper.exe",
                    "teamsbootstrapper.exe was not found on PATH or in common install locations.");
                yield break;
            }

            var result = _commandRunner.Run(
                bootstrapperPath,
                "-x -m",
                120000);

            yield return ConvertCommandResult(
                "CompanionApps",
                "RemoveTeamsBootstrapperMachineWide",
                "Executable",
                bootstrapperPath,
                result,
                "OfficeScrubC2R.TeamsBootstrapperUninstallFailed");
        }

        private IEnumerable<OperationResult> DeleteTeamsProfileRemnants()
        {
            var targets = GetTeamsProfileTargets().ToArray();
            if (targets.Length == 0)
            {
                yield return OperationResult.Skipped(
                    "CompanionApps",
                    "DeleteTeamsProfileRemnants",
                    "DirectorySet",
                    "Local user profiles",
                    "No classic Teams profile folders were found.");
                yield break;
            }

            foreach (var target in targets)
            {
                yield return DeleteFileSystemPath(target, "CompanionApps", "DeleteTeamsProfileRemnants");
            }
        }

        private IEnumerable<OperationResult> DeleteTeamsRegistryRemnants(bool is64BitOperatingSystem)
        {
            var targets = new List<RegistryTarget>
            {
                new RegistryTarget(RegistryHive.CurrentUser, @"Software\Microsoft\Teams"),
                new RegistryTarget(RegistryHive.CurrentUser, @"Software\Microsoft\Office\Teams")
            };

            foreach (var sid in GetLoadedUserSids())
            {
                targets.Add(new RegistryTarget(RegistryHive.Users, sid + @"\Software\Microsoft\Teams"));
                targets.Add(new RegistryTarget(RegistryHive.Users, sid + @"\Software\Microsoft\Office\Teams"));
            }

            foreach (var target in targets)
            {
                foreach (var result in DeleteRegistryKey(target, is64BitOperatingSystem))
                {
                    result.Step = "CompanionApps";
                    result.Action = "DeleteTeamsRegistryRemnants";
                    yield return result;
                }
            }
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

        private static OperationResult ConvertCommandResult(
            string step,
            string action,
            string targetKind,
            string target,
            CommandRunResult result,
            string errorId)
        {
            if (result.ExitCode == 0)
            {
                if (result.Output.IndexOf("OfficeScrubC2R:NOOP:", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return OperationResult.Skipped(step, action, targetKind, target, result.Output);
                }

                return OperationResult.Completed(step, action, targetKind, target, result.Output);
            }

            if (result.ExitCode == 1605 || result.ExitCode == 1614)
            {
                return OperationResult.Skipped(step, action, targetKind, target, "Installer product was not present.");
            }

            if (result.ExitCode == 3010)
            {
                return new OperationResult
                {
                    Step = step,
                    Action = action,
                    TargetKind = targetKind,
                    Target = target,
                    Status = OperationStatus.ScheduledForReboot,
                    Message = string.IsNullOrWhiteSpace(result.Output) ? "Operation completed and requested a reboot." : result.Output,
                    RebootScheduled = true
                };
            }

            return OperationResult.Blocked(step, action, targetKind, target, result.Output, errorId);
        }

        private static string CreateRemoveAppxPackagesScript(string packageName, string friendlyName)
        {
            return string.Join(
                Environment.NewLine,
                "$ErrorActionPreference = 'Stop'",
                "$packages = @(Get-AppxPackage -Name '" + EscapePowerShellSingleQuoted(packageName) + "' -AllUsers)",
                "if ($packages.Count -eq 0) { Write-Output 'OfficeScrubC2R:NOOP:No " + EscapePowerShellSingleQuoted(friendlyName) + " found.'; exit 0 }",
                "foreach ($package in $packages) {",
                "    Remove-AppxPackage -Package $package.PackageFullName -AllUsers -ErrorAction Stop",
                "    Write-Output ('Removed ' + $package.PackageFullName)",
                "}");
        }

        private static string CreateRemoveProvisionedAppxPackagesScript(string displayNamePattern, string friendlyName)
        {
            return string.Join(
                Environment.NewLine,
                "$ErrorActionPreference = 'Stop'",
                "$packages = @(Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like '" + EscapePowerShellSingleQuoted(displayNamePattern) + "' })",
                "if ($packages.Count -eq 0) { Write-Output 'OfficeScrubC2R:NOOP:No " + EscapePowerShellSingleQuoted(friendlyName) + " found.'; exit 0 }",
                "foreach ($package in $packages) {",
                "    Remove-AppxProvisionedPackage -Online -PackageName $package.PackageName -ErrorAction Stop | Out-Null",
                "    Write-Output ('Removed provisioned ' + $package.PackageName)",
                "}");
        }

        private static string EncodePowerShellCommand(string script)
        {
            return Convert.ToBase64String(Encoding.Unicode.GetBytes(script));
        }

        private static string EscapePowerShellSingleQuoted(string value)
        {
            return (value ?? string.Empty).Replace("'", "''");
        }

        private static string FindExecutable(string fileName)
        {
            var candidates = new List<string>
            {
                Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\Teams Installer\" + fileName),
                Environment.ExpandEnvironmentVariables(@"%ProgramFiles(x86)%\Teams Installer\" + fileName),
                Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\Microsoft Teams\" + fileName),
                Environment.ExpandEnvironmentVariables(@"%ProgramFiles(x86)%\Microsoft Teams\" + fileName)
            };

            var pathValue = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            foreach (var pathPart in pathValue.Split(new[] { Path.PathSeparator }, StringSplitOptions.RemoveEmptyEntries))
            {
                candidates.Add(Path.Combine(pathPart.Trim(), fileName));
            }

            return candidates.FirstOrDefault(candidate =>
                !string.IsNullOrWhiteSpace(candidate) &&
                candidate.IndexOf('%') < 0 &&
                File.Exists(candidate)) ?? string.Empty;
        }

        private static IEnumerable<string> GetTeamsProfileTargets()
        {
            var usersRoot = Path.Combine(Path.GetPathRoot(Environment.SystemDirectory) ?? @"C:\", "Users");
            if (!Directory.Exists(usersRoot))
            {
                yield break;
            }

            foreach (var profilePath in Directory.GetDirectories(usersRoot))
            {
                var profileName = Path.GetFileName(profilePath);
                if (string.Equals(profileName, "Public", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(profileName, "Default", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(profileName, "Default User", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(profileName, "All Users", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var targets = new[]
                {
                    Path.Combine(profilePath, "AppData", "Local", "Microsoft", "Teams"),
                    Path.Combine(profilePath, "AppData", "Roaming", "Microsoft", "Teams"),
                    Path.Combine(profilePath, "AppData", "Local", "SquirrelTemp"),
                    Path.Combine(profilePath, "AppData", "Local", "Microsoft", "TeamsMeetingAddin")
                };

                foreach (var target in targets.Where(target => Directory.Exists(target) || File.Exists(target)))
                {
                    yield return target;
                }
            }
        }

        private static IReadOnlyList<string> GetLoadedUserSids()
        {
            var sids = new List<string>();

            try
            {
                using (var users = RegistryKey.OpenBaseKey(RegistryHive.Users, RegistryView.Default))
                {
                    foreach (var sid in users.GetSubKeyNames())
                    {
                        if (sid.StartsWith("S-1-5-21-", StringComparison.OrdinalIgnoreCase) &&
                            sid.IndexOf("_Classes", StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            sids.Add(sid);
                        }
                    }
                }
            }
            catch (Exception)
            {
                return sids;
            }

            return sids;
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

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool MoveFileEx(string lpExistingFileName, string lpNewFileName, int dwFlags);

        private sealed class RegistryCleanupTargets
        {
            public List<RegistryTarget> KeyTargets { get; } = new List<RegistryTarget>();
            public List<RegistryValueTarget> ValueTargets { get; } = new List<RegistryValueTarget>();
            public List<OperationResult> Diagnostics { get; } = new List<OperationResult>();
        }
    }
}
