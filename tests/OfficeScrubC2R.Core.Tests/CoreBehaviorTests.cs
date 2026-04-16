using System.ComponentModel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using OfficeScrubC2R;
using Xunit;

namespace OfficeScrubC2R.Core.Tests;

public sealed class CoreBehaviorTests
{
    [Fact]
    public void OperationResult_FromExceptionPreservesDiagnosticMetadata()
    {
        var exception = new Win32Exception(5, "Access is denied.");

        var result = OperationResult.FromException(
            step: "Registry",
            action: "DeleteKey",
            targetKind: "RegistryKey",
            target: @"HKLM\Software\Microsoft\Office",
            exception: exception,
            hive: RegistryHive.LocalMachine,
            view: RegistryView.Registry64);

        Assert.Equal(OperationStatus.Failed, result.Status);
        Assert.Equal("Registry", result.Step);
        Assert.Equal("DeleteKey", result.Action);
        Assert.Equal("RegistryKey", result.TargetKind);
        Assert.Equal(@"HKLM\Software\Microsoft\Office", result.Target);
        Assert.Equal(RegistryHive.LocalMachine, result.RegistryHive);
        Assert.Equal(RegistryView.Registry64, result.RegistryView);
        Assert.Equal(nameof(Win32Exception), result.ExceptionType);
        Assert.Equal(exception.HResult, result.HResult);
        Assert.Equal(5, result.Win32Error);
        Assert.False(result.RebootScheduled);
    }

    [Fact]
    public void RegistryAccess_UsesExplicitViewsForSixtyFourBitOperatingSystems()
    {
        var access = new RegistryAccess(is64BitOperatingSystem: true);

        var views = access.GetCandidateViews().ToArray();

        Assert.Equal(new[] { RegistryView.Registry64, RegistryView.Registry32 }, views);
    }

    [Fact]
    public void RegistryAccess_UsesThirtyTwoBitViewForThirtyTwoBitOperatingSystems()
    {
        var access = new RegistryAccess(is64BitOperatingSystem: false);

        var view = Assert.Single(access.GetCandidateViews());

        Assert.Equal(RegistryView.Registry32, view);
    }

    [Fact]
    public void RegistryAccess_RecordsDiagnosticsForRegistryFailures()
    {
        var access = new RegistryAccess(is64BitOperatingSystem: true);

        var exists = access.KeyExists((RegistryHive)12345, "Software", RegistryView.Registry64);

        Assert.False(exists);
        var diagnostic = Assert.Single(access.Diagnostics);
        Assert.Equal(OperationStatus.Failed, diagnostic.Status);
        Assert.Equal("Registry", diagnostic.Step);
        Assert.Equal("KeyExists", diagnostic.Action);
        Assert.Equal(RegistryView.Registry64, diagnostic.RegistryView);
    }

    [Fact]
    public void OfficeScope_DetectsKnownC2RPathsAndProductCodes()
    {
        Assert.True(OfficeScope.IsC2RPath(@"C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE"));
        Assert.True(OfficeScope.IsInScope("{90160000-008F-0000-1000-0000000FF1CE}"));
        Assert.False(OfficeScope.IsInScope("{90140000-008F-0000-1000-0000000FF1CE}"));
    }

    [Fact]
    public void GuidHelper_RoundTripsCompressedProductCodes()
    {
        const string expanded = "{90160000-008F-0000-1000-0000000FF1CE}";

        var compressed = GuidHelper.GetCompressedGuid(expanded);
        var roundTripped = GuidHelper.GetExpandedGuid(compressed);

        Assert.Equal(expanded, roundTripped);
    }

    [Fact]
    public void ScrubPlanner_ReturnsPlanOnlyOperations()
    {
        var state = new OfficeC2RState
        {
            IsElevated = true,
            IsSystem = false,
            Is64BitOperatingSystem = true
        };
        state.InstalledProducts.Add(new OfficeProductInfo
        {
            ProductId = "O365ProPlusRetail",
            DisplayName = "Microsoft 365 Apps",
            Source = "Test"
        });

        var plan = ScrubPlanner.CreatePlan(state, keepLicense: true, planOnly: true);

        Assert.True(plan.PlanOnly);
        Assert.True(plan.KeepLicense);
        Assert.Contains(plan.PlannedOperations, item => item.Status == OperationStatus.WouldRun);
        Assert.Contains(plan.PlannedOperations, item =>
            item.Step == "CompanionApps" &&
            item.Action == "RemoveTeamsAndCopilot" &&
            item.Status == OperationStatus.WouldRun);
    }

    [Fact]
    public void ScrubPlanner_HonorsCompanionAppKeepFlags()
    {
        var state = new OfficeC2RState
        {
            IsElevated = true,
            Is64BitOperatingSystem = true
        };

        var plan = ScrubPlanner.CreatePlan(
            state,
            keepLicense: false,
            planOnly: true,
            keepTeams: true,
            keepCopilot: true);

        Assert.True(plan.KeepTeams);
        Assert.True(plan.KeepCopilot);
        Assert.Contains(plan.PlannedOperations, item =>
            item.Step == "CompanionApps" &&
            item.Action == "RemoveTeamsAndCopilot" &&
            item.Status == OperationStatus.Skipped);
    }

    [Fact]
    public void CleanupExecutor_DeletesRequestedDirectoryAndReportsResult()
    {
        var root = Path.Combine(Path.GetTempPath(), "OfficeScrubC2R.Tests", Guid.NewGuid().ToString("N"));
        var child = Path.Combine(root, "child");
        Directory.CreateDirectory(child);
        File.WriteAllText(Path.Combine(child, "sample.txt"), "delete me");

        var request = new ScrubExecutionRequest
        {
            State = new OfficeC2RState { IsElevated = true, Is64BitOperatingSystem = true },
            SkipBuiltInTargets = true,
            SkipCompanionAppTargets = true
        };
        request.ExtraFileSystemTargets.Add(root);

        var result = new CleanupExecutor().Execute(request);

        Assert.False(Directory.Exists(root));
        Assert.Contains(result.ExecutedOperations, item =>
            item.Step == "Files" &&
            item.Action == "DeleteDirectory" &&
            item.Target == root &&
            item.Status == OperationStatus.Completed);
    }

    [Fact]
    public void CleanupExecutor_DeletesRequestedRegistryKeyAndReportsExplicitView()
    {
        var subKey = @"Software\OfficeScrubC2R\Tests\" + Guid.NewGuid().ToString("N");
        using (var key = Registry.CurrentUser.CreateSubKey(subKey))
        {
            key!.SetValue("Value", "delete me");
        }

        var request = new ScrubExecutionRequest
        {
            State = new OfficeC2RState { IsElevated = true, Is64BitOperatingSystem = true },
            SkipBuiltInTargets = true,
            SkipCompanionAppTargets = true
        };
        request.ExtraRegistryTargets.Add(new RegistryTarget(RegistryHive.CurrentUser, subKey));

        var result = new CleanupExecutor().Execute(request);

        using var remaining = Registry.CurrentUser.OpenSubKey(subKey);
        Assert.Null(remaining);
        Assert.Contains(result.ExecutedOperations, item =>
            item.Step == "Registry" &&
            item.Action == "DeleteKey" &&
            item.Target.EndsWith(subKey) &&
            item.Status == OperationStatus.Completed &&
            item.RegistryView.HasValue);
    }

    [Fact]
    public void CleanupExecutor_BlocksWhenNotElevated()
    {
        var request = new ScrubExecutionRequest
        {
            State = new OfficeC2RState { IsElevated = false, Is64BitOperatingSystem = true },
            SkipBuiltInTargets = true,
            SkipCompanionAppTargets = true
        };

        var result = new CleanupExecutor().Execute(request);

        Assert.Equal("Blocked", result.ExecutionStatus);
        Assert.Contains(result.ExecutedOperations, item =>
            item.Status == OperationStatus.Blocked &&
            item.ErrorId == "OfficeScrubC2R.AdminRequired");
    }

    [Fact]
    public void CleanupExecutor_RemovesTeamsAndCopilotCompanionAppsThroughCommandRunner()
    {
        var runner = new RecordingCommandRunner();
        var request = new ScrubExecutionRequest
        {
            State = new OfficeC2RState { IsElevated = true, Is64BitOperatingSystem = true },
            SkipBuiltInTargets = true,
            SkipCompanionProfileTargets = true
        };

        var result = new CleanupExecutor(runner).Execute(request);

        Assert.Contains(result.ExecutedOperations, item =>
            item.Step == "CompanionApps" &&
            item.Action == "RemoveTeamsAppxPackage" &&
            item.Status == OperationStatus.Completed);
        Assert.Contains(result.ExecutedOperations, item =>
            item.Step == "CompanionApps" &&
            item.Action == "RemoveCopilotAppxPackage" &&
            item.Status == OperationStatus.Completed);
        Assert.Contains(runner.PowerShellScripts, script => script.Contains("MSTeams"));
        Assert.Contains(runner.PowerShellScripts, script => script.Contains("Microsoft.Copilot"));
        Assert.Contains(runner.Invocations, invocation =>
            invocation.FileName == "msiexec.exe" &&
            invocation.Arguments.Contains("{731F6BAA-A986-45A4-8936-7C3AAAAA760B}"));
    }

    [Fact]
    public void CleanupExecutor_SkipsCompanionAppsWhenKeepFlagsAreSet()
    {
        var runner = new RecordingCommandRunner();
        var request = new ScrubExecutionRequest
        {
            State = new OfficeC2RState { IsElevated = true, Is64BitOperatingSystem = true },
            KeepTeams = true,
            KeepCopilot = true,
            SkipBuiltInTargets = true,
            SkipCompanionProfileTargets = true
        };

        var result = new CleanupExecutor(runner).Execute(request);

        Assert.Contains(result.ExecutedOperations, item =>
            item.Step == "CompanionApps" &&
            item.Action == "RemoveTeams" &&
            item.Status == OperationStatus.Skipped);
        Assert.Contains(result.ExecutedOperations, item =>
            item.Step == "CompanionApps" &&
            item.Action == "RemoveCopilot" &&
            item.Status == OperationStatus.Skipped);
        Assert.Empty(runner.Invocations);
    }

    [Fact]
    public void CleanupExecutor_DoesNotUseFragileAllDirectoriesEnumeration()
    {
        var sourcePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "src", "OfficeScrubC2R.Core", "CleanupExecutor.cs"));
        var source = File.ReadAllText(sourcePath);

        Assert.DoesNotContain("SearchOption.AllDirectories", source);
    }

    [Fact]
    public void ProcessCommandRunner_CapturesStdoutAndStderrWithoutSequentialReadDeadlock()
    {
        var script = "$stdout = 'o' * 100000; $stderr = 'e' * 100000; [Console]::Out.Write($stdout); [Console]::Error.Write($stderr)";
        var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(script));

        var result = new ProcessCommandRunner().Run(
            "powershell.exe",
            "-NoProfile -ExecutionPolicy Bypass -EncodedCommand " + encoded,
            30000);

        Assert.Equal(0, result.ExitCode);
        Assert.Contains(new string('o', 100), result.Output);
        Assert.Contains(new string('e', 100), result.Output);
    }

    private sealed class RecordingCommandRunner : ICommandRunner
    {
        public List<CommandInvocation> Invocations { get; } = new List<CommandInvocation>();

        public IEnumerable<string> PowerShellScripts =>
            Invocations
                .Where(invocation => invocation.FileName == "powershell.exe")
                .Select(invocation => DecodePowerShellScript(invocation.Arguments));

        public CommandRunResult Run(string fileName, string arguments, int timeoutMilliseconds)
        {
            Invocations.Add(new CommandInvocation(fileName, arguments, timeoutMilliseconds));
            return new CommandRunResult(0, "removed");
        }

        private static string DecodePowerShellScript(string arguments)
        {
            const string marker = "-EncodedCommand ";
            var markerIndex = arguments.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            Assert.True(markerIndex >= 0, "PowerShell command should use -EncodedCommand.");
            var encoded = arguments.Substring(markerIndex + marker.Length).Trim().Trim('"');
            return Encoding.Unicode.GetString(Convert.FromBase64String(encoded));
        }
    }

    private sealed class CommandInvocation
    {
        public CommandInvocation(string fileName, string arguments, int timeoutMilliseconds)
        {
            FileName = fileName;
            Arguments = arguments;
            TimeoutMilliseconds = timeoutMilliseconds;
        }

        public string FileName { get; }
        public string Arguments { get; }
        public int TimeoutMilliseconds { get; }
    }
}
