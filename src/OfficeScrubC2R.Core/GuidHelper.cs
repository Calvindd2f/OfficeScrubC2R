using System;
using System.Text;

namespace OfficeScrubC2R
{
    public static class GuidHelper
    {
        public static string GetExpandedGuid(string compressedGuid)
        {
            if (string.IsNullOrWhiteSpace(compressedGuid) || compressedGuid.Length != 32)
            {
                return string.Empty;
            }

            var normalized = compressedGuid.Trim().ToUpperInvariant();
            var builder = new StringBuilder(38);
            builder.Append('{');
            builder.Append(Reverse(normalized.Substring(0, 8)));
            builder.Append('-');
            builder.Append(Reverse(normalized.Substring(8, 4)));
            builder.Append('-');
            builder.Append(Reverse(normalized.Substring(12, 4)));
            builder.Append('-');
            builder.Append(SwapPairs(normalized.Substring(16, 4)));
            builder.Append('-');
            builder.Append(SwapPairs(normalized.Substring(20, 12)));
            builder.Append('}');
            return builder.ToString();
        }

        public static string GetCompressedGuid(string expandedGuid)
        {
            if (string.IsNullOrWhiteSpace(expandedGuid))
            {
                return string.Empty;
            }

            var normalized = expandedGuid.Trim().Trim('{', '}').ToUpperInvariant();
            var parts = normalized.Split('-');
            if (parts.Length != 5 ||
                parts[0].Length != 8 ||
                parts[1].Length != 4 ||
                parts[2].Length != 4 ||
                parts[3].Length != 4 ||
                parts[4].Length != 12)
            {
                return string.Empty;
            }

            return string.Concat(
                Reverse(parts[0]),
                Reverse(parts[1]),
                Reverse(parts[2]),
                SwapPairs(parts[3]),
                SwapPairs(parts[4]));
        }

        public static bool GetDecodedGuid(string encodedGuid, out string decodedGuid)
        {
            decodedGuid = string.Empty;
            if (string.IsNullOrWhiteSpace(encodedGuid) || encodedGuid.Length != 20)
            {
                return false;
            }

            var expanded = GetExpandedGuid(encodedGuid.PadRight(32, '0'));
            if (string.IsNullOrEmpty(expanded))
            {
                return false;
            }

            decodedGuid = expanded;
            return true;
        }

        private static string Reverse(string value)
        {
            var chars = value.ToCharArray();
            Array.Reverse(chars);
            return new string(chars);
        }

        private static string SwapPairs(string value)
        {
            var chars = value.ToCharArray();
            for (var index = 0; index < chars.Length - 1; index += 2)
            {
                var current = chars[index];
                chars[index] = chars[index + 1];
                chars[index + 1] = current;
            }

            return new string(chars);
        }
    }
}
