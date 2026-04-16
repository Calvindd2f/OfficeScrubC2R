using OfficeScrubC2R;
using System.Management.Automation;

namespace OfficeScrubC2R.PowerShell
{
    [Cmdlet(VerbsDiagnostic.Test, "OfficeC2RState")]
    [OutputType(typeof(OfficeC2RState))]
    public sealed class TestOfficeC2RStateCommand : PSCmdlet
    {
        protected override void ProcessRecord()
        {
            var service = new PreflightService();
            WriteObject(service.GetState());
        }
    }
}
