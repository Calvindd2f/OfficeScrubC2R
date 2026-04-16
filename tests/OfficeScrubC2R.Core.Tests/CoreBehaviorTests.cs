using System.ComponentModel;
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
    public void ScrubPlanner_BlocksExecutionAndReturnsPlanOnlyOperations()
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
        var block = ScrubPlanner.CreateBlockedExecutionResult();

        Assert.True(plan.PlanOnly);
        Assert.True(plan.KeepLicense);
        Assert.Contains(plan.PlannedOperations, item => item.Status == OperationStatus.WouldRun);
        Assert.Equal(OperationStatus.Blocked, block.Status);
        Assert.Equal("OfficeScrubC2R.DestructiveExecutionNotSupported", block.ErrorId);
    }
}
