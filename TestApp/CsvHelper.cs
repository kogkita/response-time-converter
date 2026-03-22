using System.Collections.Generic;
using System.Text;

namespace TestApp
{
    /// <summary>
    /// Shared CSV parsing utilities used across multiple processors
    /// (JTL, ResponseTime, RunComparison, Nmon, BLG).
    /// Consolidates the quote-aware CSV line splitter that was previously
    /// duplicated in five separate classes.
    /// </summary>
    public static class CsvHelper
    {
        /// <summary>
        /// Splits a CSV line respecting quoted fields and escaped quotes ("").
        /// Returns fields as a string array.
        /// </summary>
        public static string[] SplitCsvLine(string line)
        {
            var fields = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; }
                        else inQuotes = false;
                    }
                    else sb.Append(c);
                }
                else
                {
                    if (c == '"') inQuotes = true;
                    else if (c == ',') { fields.Add(sb.ToString()); sb.Clear(); }
                    else sb.Append(c);
                }
            }
            fields.Add(sb.ToString());
            return fields.ToArray();
        }

        /// <summary>
        /// Splits a CSV line into a List&lt;string&gt; (convenience overload).
        /// </summary>
        public static List<string> SplitCsvLineToList(string line)
        {
            return new List<string>(SplitCsvLine(line));
        }
    }
}
