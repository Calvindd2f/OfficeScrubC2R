using System;
using System.Collections.Generic;
using System.ComponentModel;
using Microsoft.Win32;

namespace OfficeScrubC2R
{
    public enum OperationStatus
    {
        Completed,
        Skipped,
        WouldRun,
        Blocked,
        Failed,
        ScheduledForReboot
    }

    public sealed class OperationResult
    {
        public string Step { get; set; } = string.Empty;
        public string Action { get; set; } = string.Empty;
        public string TargetKind { get; set; } = string.Empty;
        public string Target { get; set; } = string.Empty;
        public RegistryHive? RegistryHive { get; set; }
        public RegistryView? RegistryView { get; set; }
        public OperationStatus Status { get; set; }
        public string Message { get; set; } = string.Empty;
        public string ExceptionType { get; set; } = string.Empty;
        public int? HResult { get; set; }
        public int? Win32Error { get; set; }
        public bool RebootScheduled { get; set; }
        public string ErrorId { get; set; } = string.Empty;

        public static OperationResult Completed(string step, string action, string targetKind, string target, string message, RegistryHive? hive = null, RegistryView? view = null)
        {
            return Create(step, action, targetKind, target, OperationStatus.Completed, message, hive, view);
        }

        public static OperationResult Skipped(string step, string action, string targetKind, string target, string message, RegistryHive? hive = null, RegistryView? view = null)
        {
            return Create(step, action, targetKind, target, OperationStatus.Skipped, message, hive, view);
        }

        public static OperationResult WouldRun(string step, string action, string targetKind, string target, string message, RegistryHive? hive = null, RegistryView? view = null)
        {
            return Create(step, action, targetKind, target, OperationStatus.WouldRun, message, hive, view);
        }

        public static OperationResult Blocked(string step, string action, string targetKind, string target, string message, string errorId, RegistryHive? hive = null, RegistryView? view = null)
        {
            var result = Create(step, action, targetKind, target, OperationStatus.Blocked, message, hive, view);
            result.ErrorId = errorId;
            return result;
        }

        public static OperationResult FromException(string step, string action, string targetKind, string target, Exception exception, RegistryHive? hive = null, RegistryView? view = null, bool rebootScheduled = false)
        {
            var result = Create(step, action, targetKind, target, OperationStatus.Failed, exception.Message, hive, view);
            result.ExceptionType = exception.GetType().Name;
            result.HResult = exception.HResult;
            result.RebootScheduled = rebootScheduled;

            if (exception is Win32Exception win32Exception)
            {
                result.Win32Error = win32Exception.NativeErrorCode;
            }

            return result;
        }

        private static OperationResult Create(string step, string action, string targetKind, string target, OperationStatus status, string message, RegistryHive? hive, RegistryView? view)
        {
            return new OperationResult
            {
                Step = step,
                Action = action,
                TargetKind = targetKind,
                Target = target,
                RegistryHive = hive,
                RegistryView = view,
                Status = status,
                Message = message
            };
        }
    }

    public sealed class OfficeProductInfo
    {
        public string ProductId { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string DisplayVersion { get; set; } = string.Empty;
        public string Architecture { get; set; } = string.Empty;
        public string Source { get; set; } = string.Empty;
        public RegistryHive? RegistryHive { get; set; }
        public RegistryView? RegistryView { get; set; }
        public string RegistryPath { get; set; } = string.Empty;
        public string UninstallString { get; set; } = string.Empty;
        public bool IsClickToRun { get; set; }
    }

    public sealed class OfficeC2RState
    {
        public bool IsElevated { get; set; }
        public bool IsSystem { get; set; }
        public bool Is64BitOperatingSystem { get; set; }
        public List<OfficeProductInfo> InstalledProducts { get; } = new List<OfficeProductInfo>();
        public List<string> RunningOfficeProcesses { get; } = new List<string>();
        public List<string> ClickToRunServices { get; } = new List<string>();
        public List<string> PackagePaths { get; } = new List<string>();
        public List<string> PendingRebootDeletes { get; } = new List<string>();
        public List<OperationResult> PreflightResults { get; } = new List<OperationResult>();
        public List<OperationResult> Issues { get; } = new List<OperationResult>();
    }

    public sealed class ScrubPlan
    {
        public DateTimeOffset CreatedAt { get; set; } = DateTimeOffset.UtcNow;
        public bool KeepLicense { get; set; }
        public bool PlanOnly { get; set; }
        public string ExecutionStatus { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public OfficeC2RState State { get; set; } = new OfficeC2RState();
        public List<OperationResult> PlannedOperations { get; } = new List<OperationResult>();
        public List<OperationResult> ExecutedOperations { get; } = new List<OperationResult>();
    }

    public sealed class ScrubExecutionRequest
    {
        public OfficeC2RState State { get; set; } = new OfficeC2RState();
        public bool KeepLicense { get; set; }
        public bool SkipBuiltInTargets { get; set; }
        public List<string> ExtraFileSystemTargets { get; } = new List<string>();
        public List<RegistryTarget> ExtraRegistryTargets { get; } = new List<RegistryTarget>();
    }

    public sealed class RegistryTarget
    {
        public RegistryTarget(RegistryHive hive, string subKey)
        {
            Hive = hive;
            SubKey = subKey;
        }

        public RegistryHive Hive { get; }
        public string SubKey { get; }
    }

    public sealed class RegistryValueTarget
    {
        public RegistryValueTarget(RegistryHive hive, string subKey, string valueName)
        {
            Hive = hive;
            SubKey = subKey;
            ValueName = valueName;
        }

        public RegistryHive Hive { get; }
        public string SubKey { get; }
        public string ValueName { get; }
    }
}
