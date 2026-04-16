using System;
using System.Collections.Generic;

namespace OfficeScrubC2R
{
    public static class OfficeConstants
    {
        public const string OfficeProductCodeSuffix = "0000000FF1CE}";

        public static readonly string[] C2RPatterns = new[]
        {
            @"\ROOT\OFFICE1",
            @"Microsoft Office\Root\",
            @"\microsoft shared\ClickToRun",
            @"\Microsoft Office\PackageManifests",
            @"\Microsoft Office\PackageSunrisePolicies",
            @"Microsoft Office 15",
            @"Microsoft Office 16",
            "OfficeClickToRun.exe"
        };

        public static readonly string[] OfficeProcesses = new[]
        {
            "appvshnotify", "integratedoffice", "integrator", "firstrun",
            "communicator", "msosync", "onenotem", "iexplore",
            "mavinject32", "werfault", "perfboost", "roamingoffice",
            "officeclicktorun", "officeondemand", "officec2rclient",
            "winword", "excel", "powerpnt", "outlook", "onenote",
            "mspub", "msaccess", "lync", "skype", "teams",
            "ms-teams", "msteams", "copilot"
        };

        public static readonly string[] ClickToRunServiceNames = new[]
        {
            "ClickToRunSvc",
            "OfficeClickToRun"
        };

        public static readonly ISet<string> ValidOfficeSkuCodes = new HashSet<string>(
            new[] { "007E", "008F", "008C", "24E1", "237A", "00DD" },
            StringComparer.OrdinalIgnoreCase);
    }
}
