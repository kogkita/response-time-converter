using OfficeOpenXml;
using System;
using System.Linq;

namespace TestApp
{
    /// <summary>
    /// Shared Excel naming helpers used by JTLFileProcessing,
    /// ResponseTimeConverter, and RunComparisonProcessor.
    /// Ensures sheet and table names are unique and within Excel's limits.
    /// </summary>
    public static class ExcelNameHelper
    {
        /// <summary>
        /// Returns a worksheet name that is unique within <paramref name="pkg"/>
        /// and within Excel's 31-character limit. Appends a numeric suffix if
        /// a collision is detected.
        /// </summary>
        public static string UniqueSheetName(ExcelPackage pkg, string name)
        {
            if (name.Length > 31) name = name[..31];

            string candidate = name;
            int n = 2;
            while (pkg.Workbook.Worksheets.Any(
                ws => ws.Name.Equals(candidate, StringComparison.OrdinalIgnoreCase)))
            {
                candidate = $"{name[..Math.Min(name.Length, 28)]} {n++}";
            }

            return candidate;
        }

        /// <summary>
        /// Returns a table name that is unique across all worksheets in
        /// <paramref name="pkg"/> (Excel requires workbook-wide uniqueness).
        /// </summary>
        public static string UniqueTableName(ExcelPackage pkg, string name)
        {
            var existing = pkg.Workbook.Worksheets
                .SelectMany(ws => ws.Tables.Select(t => t.Name))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            string candidate = name;
            int n = 2;
            while (existing.Contains(candidate))
                candidate = $"{name}{n++}";

            return candidate;
        }
    }
}
