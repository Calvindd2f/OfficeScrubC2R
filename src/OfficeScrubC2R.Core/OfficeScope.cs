using System;

namespace OfficeScrubC2R
{
    public static class OfficeScope
    {
        public static bool IsC2RPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            foreach (var pattern in OfficeConstants.C2RPatterns)
            {
                if (path.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        public static bool IsInScope(string productCode)
        {
            if (string.IsNullOrWhiteSpace(productCode) || productCode.Length != 38)
            {
                return false;
            }

            var upper = productCode.ToUpperInvariant();
            if (!upper.EndsWith(OfficeConstants.OfficeProductCodeSuffix, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (!int.TryParse(upper.Substring(3, 2), out var version) || version <= 14)
            {
                return false;
            }

            var sku = upper.Substring(10, 4);
            return OfficeConstants.ValidOfficeSkuCodes.Contains(sku) ||
                   upper == "{6C1ADE97-24E1-4AE4-AEDD-86D3A209CE60}" ||
                   upper == "{9520DDEB-237A-41DB-AA20-F2EF2360DCEB}" ||
                   upper == "{9AC08E99-230B-47E8-9721-4577B7F124EA}";
        }
    }
}
