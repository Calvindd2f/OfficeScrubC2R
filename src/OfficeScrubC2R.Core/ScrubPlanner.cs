namespace OfficeScrubC2R
{
    public static class ScrubPlanner
    {
        public static ScrubPlan CreatePlan(OfficeC2RState state, bool keepLicense, bool planOnly)
        {
            var plan = new ScrubPlan
            {
                KeepLicense = keepLicense,
                PlanOnly = planOnly,
                State = state,
                ExecutionStatus = planOnly ? "PlanOnly" : "Ready",
                Message = planOnly
                    ? "The plan describes the cleanup actions Invoke-OfficeScrubC2R will run outside -PlanOnly/-WhatIf."
                    : "The cleanup plan is ready to execute."
            };

            plan.PlannedOperations.Add(OperationResult.WouldRun(
                "Preflight",
                "ValidatePrivileges",
                "Privilege",
                "Administrator",
                "Validate elevation before destructive cleanup."));

            plan.PlannedOperations.Add(OperationResult.WouldRun(
                "Processes",
                "TerminateOfficeProcesses",
                "ProcessSet",
                "OfficeProcesses",
                "Terminate Office-related processes."));

            plan.PlannedOperations.Add(OperationResult.WouldRun(
                "Services",
                "StopAndDeleteClickToRunServices",
                "ServiceSet",
                "ClickToRunSvc,OfficeClickToRun",
                "Stop and delete Click-to-Run services."));

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
                    "Remove Office licensing data."));

            return plan;
        }
    }
}
