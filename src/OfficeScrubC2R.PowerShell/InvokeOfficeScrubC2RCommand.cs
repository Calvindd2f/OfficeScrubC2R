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
        public SwitchParameter PassThru { get; set; }

        protected override void ProcessRecord()
        {
            var state = new PreflightService().GetState();
            var plan = ScrubPlanner.CreatePlan(state, KeepLicense.IsPresent, PlanOnly.IsPresent);

            if (PlanOnly.IsPresent)
            {
                WriteObject(plan);
                return;
            }

            if (MyInvocation.BoundParameters.ContainsKey("WhatIf"))
            {
                ShouldProcess("Office Click-to-Run installation", "Scrub Office C2R");
                WriteObject(plan);
                return;
            }

            var blocked = ScrubPlanner.CreateBlockedExecutionResult();
            var exception = new InvalidOperationException(blocked.Message);
            var error = new ErrorRecord(
                exception,
                blocked.ErrorId,
                ErrorCategory.NotImplemented,
                "Office Click-to-Run");

            ThrowTerminatingError(error);
        }
    }
}
