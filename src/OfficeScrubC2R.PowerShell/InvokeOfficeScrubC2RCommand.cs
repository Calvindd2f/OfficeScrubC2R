using System;
using OfficeScrubC2R;
using System.Management.Automation;

namespace OfficeScrubC2R.PowerShell
{
    [Cmdlet(VerbsLifecycle.Invoke, "OfficeScrubC2R", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.High)]
    [OutputType(typeof(ScrubPlan))]
    public sealed class InvokeOfficeScrubC2RCommand : PSCmdlet
    {
        [Parameter]
        public SwitchParameter PlanOnly { get; set; }

        [Parameter]
        public SwitchParameter KeepLicense { get; set; }

        [Parameter]
        public SwitchParameter KeepTeams { get; set; }

        [Parameter]
        public SwitchParameter KeepCopilot { get; set; }

        [Parameter]
        public SwitchParameter PassThru { get; set; }

        protected override void ProcessRecord()
        {
            var state = new PreflightService().GetState();
            var plan = ScrubPlanner.CreatePlan(
                state,
                KeepLicense.IsPresent,
                PlanOnly.IsPresent,
                KeepTeams.IsPresent,
                KeepCopilot.IsPresent);

            if (PlanOnly.IsPresent)
            {
                WriteObject(plan);
                return;
            }

            if (MyInvocation.BoundParameters.ContainsKey("WhatIf"))
            {
                ShouldProcess("Office Click-to-Run installation", "Scrub Office C2R");
                plan.ExecutionStatus = "WhatIf";
                plan.Message = "WhatIf requested; no cleanup operations were executed.";
                WriteObject(plan);
                return;
            }

            if (!state.IsElevated)
            {
                ThrowTerminatingError(new ErrorRecord(
                    new UnauthorizedAccessException("Administrator privileges are required for destructive Office cleanup. Start an elevated PowerShell session and rerun Invoke-OfficeScrubC2R."),
                    "OfficeScrubC2R.AdminRequired",
                    ErrorCategory.PermissionDenied,
                    "Administrator"));
                return;
            }

            if (!ShouldProcess("Office Click-to-Run installation", "Scrub Office C2R"))
            {
                WriteObject(plan);
                return;
            }

            var request = new ScrubExecutionRequest
            {
                State = state,
                KeepLicense = KeepLicense.IsPresent,
                KeepTeams = KeepTeams.IsPresent,
                KeepCopilot = KeepCopilot.IsPresent
            };

            var result = new CleanupExecutor().Execute(request);
            WriteObject(result);
        }
    }
}
