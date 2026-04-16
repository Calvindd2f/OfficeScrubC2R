using OfficeScrubC2R;
using System.Management.Automation;

namespace OfficeScrubC2R.PowerShell
{
    [Cmdlet(VerbsCommon.Get, "InstalledOfficeProducts")]
    [OutputType(typeof(OfficeProductInfo))]
    public sealed class GetInstalledOfficeProductsCommand : PSCmdlet
    {
        protected override void ProcessRecord()
        {
            var service = new OfficeDetectionService();
            foreach (var product in service.GetInstalledProducts())
            {
                WriteObject(product);
            }
        }
    }
}
