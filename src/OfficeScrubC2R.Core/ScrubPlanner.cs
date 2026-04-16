namespace OfficeScrubC2R
{
    public static class ScrubPlanner
    {
        public const string DestructiveExecutionNotSupportedErrorId = "OfficeScrubC2R.DestructiveExecutionNotSupported";

        public static ScrubPlan CreatePlan(OfficeC2RState state, bool keepLicense, bool planOnly)
        {
            var plan = new ScrubPlan
            {
                KeepLicense = keepLicense,
                PlanOnly = planOnly,
                State = state,
                ExecutionStatus = "PlanOnly",
                Message = "This milestone is non-destructive. The plan describes actions that a future full scrub implementation may perform."
            };

            plan.PlannedOperations.Add(OperationResult.WouldRun(
                "Preflight",
                "ValidatePrivileges",
                "Privilege",
                "Administrator",
                "Validate elevation and SYSTEM readiness before destructive cleanup."));

            plan.PlannedOperations.Add(OperationResult.WouldRun(
                "Processes",
                "TerminateOfficeProcesses",
                "ProcessSet",
                "OfficeProcesses",
                "Terminate Office-related processes after explicit destructive execution is implemented."));

            plan.PlannedOperations.Add(OperationResult.WouldRun(
                "Services",
                "StopAndDeleteClickToRunServices",
                "ServiceSet",
                "ClickToRunSvc,OfficeClickToRun",
                "Stop and delete Click-to-Run services after explicit destructive execution is implemented."));

            plan.PlannedOperations.Add(OperationResult.WouldRun(
                "Registry",
                "RemoveClickToRunRegistry",
                "RegistryKeySet",
                @"HKLM/HKCU Office ClickToRun keys",
                "Remove Office Click-to-Run registry keys through explicit 32-bit and 64-bit registry views."));

            plan.PlannedOperations.Add(OperationResult.WouldRun(
                "Files",
                "RemoveClickToRunFiles",
                "DirectorySet",
                "Detected Office package paths",
                "Remove Office package files and schedule locked files for reboot deletion."));

            plan.PlannedOperations.Add(keepLicense
                ? OperationResult.Skipped(
                    "Licensing",
                    "RemoveOfficeLicensing",
                    "LicenseStore",
                    "Office licensing data",
                    "License cleanup skipped because KeepLicense was requested.")
                : OperationResult.WouldRun(
                    "Licensing",
                    "RemoveOfficeLicensing",
                    "LicenseStore",
                    "Office licensing data",
                    "Remove Office licensing data after explicit destructive execution is implemented."));

            return plan;
        }

        public static OperationResult CreateBlockedExecutionResult()
        {
            return OperationResult.Blocked(
                "Invoke",
                "ExecuteScrub",
                "OfficeInstallation",
                "Office Click-to-Run",
                "Destructive Office cleanup is intentionally blocked in the hardened baseline. Use -PlanOnly or -WhatIf; full scrub execution belongs to a later parity milestone.",
                DestructiveExecutionNotSupportedErrorId);
        }
    }
}
