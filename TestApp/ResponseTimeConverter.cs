using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System.IO;

namespace TestApp
{
    public class ResponseTimeRecord
    {
        public string TransactionName { get; set; }
        public int Samples { get; set; }
        public double Average { get; set; }
        public double Median { get; set; }
        public Dictionary<string, double> Percentiles { get; set; } = new();
        public double Min { get; set; }
        public double Max { get; set; }
        public double ErrorPercent { get; set; }
    }

    public static class ResponseTimeConverter
    {
        // ─────────────────────────────────────────────────────────────────────
        // Public entry points
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>
        /// Converts a single CSV file to an Excel workbook saved at
        /// <paramref name="excelPath"/>.
        /// </summary>
        public static void Convert(string csvPath, string excelPath, bool includeCharts = true)
        {
            ExcelPackage.License.SetNonCommercialPersonal("Response Time Converter");

            var (records, percentileHeaders) = ReadCsv(csvPath);

            // Sheet 1: A-Z by transaction name
            var recordsAZ = records
                .OrderBy(r => r.TransactionName, StringComparer.Ordinal)
                .ToList();

            using var package = new ExcelPackage();

            string dataName = UniqueSheetName(package, "Response Times");
            string chartName = UniqueSheetName(package, "Latency Charts");

            WriteResponseSheet(package, recordsAZ, percentileHeaders, dataName);

            if (includeCharts && records.Count > 0)
                // AddMiniChartsAndSave sorts by avg desc internally and handles SaveAs
                ResponseTimeConverterExcelCharts.AddMiniChartsAndSave(
                    package, records, chartName, excelPath);
            else
                package.SaveAs(new FileInfo(excelPath));
        }

        public static void AppendToPackage(
            ExcelPackage package,
            string csvPath,
            string? prefix,
            bool includeCharts = true)
        {
            var (records, percentileHeaders) = ReadCsv(csvPath);

            // Sheet A-Z
            var recordsAZ = records
                .OrderBy(r => r.TransactionName, StringComparer.Ordinal)
                .ToList();

            string dataName = prefix != null ? $"{prefix} \u2013 Response Times" : "Response Times";
            string chartName = prefix != null ? $"{prefix} \u2013 Latency Charts" : "Latency Charts";

            dataName = UniqueSheetName(package, dataName);
            chartName = UniqueSheetName(package, chartName);

            WriteResponseSheet(package, recordsAZ, percentileHeaders, dataName);

            if (includeCharts && records.Count > 0)
            {
                var byAvg = records.OrderByDescending(r => r.Average).ToList();
                _pendingCharts[chartName] = (byAvg, package);

                // Build the full chart sheet (same setup as AddMiniChartsAndSave)
                // Scale shell is registered FIRST so it becomes chart1 in the ZIP
                ResponseTimeConverterExcelCharts.BuildChartSheetShells(
                    package, chartName, byAvg);
            }
        }

        private static readonly Dictionary<string, (List<ResponseTimeRecord> records, ExcelPackage pkg)>
            _pendingCharts = new();

        /// <summary>
        /// Disposes any pending ExcelPackage objects and clears the chart cache.
        /// Call at the start of each run and in any error/cancel path.
        /// </summary>
        public static void ClearPendingCharts()
        {
            foreach (var entry in _pendingCharts.Values)
            {
                try { entry.pkg?.Dispose(); } catch { }
            }
            _pendingCharts.Clear();
        }

        /// <summary>Called by MainWindow after SaveAs in clubbed mode.</summary>
        public static void InjectPendingCharts(string xlsxPath)
        {
            try
            {
                foreach (var kvp in _pendingCharts)
                    ResponseTimeConverterExcelCharts.InjectChartForSheet(xlsxPath, kvp.Key, kvp.Value.records);
            }
            finally
            {
                _pendingCharts.Clear();
            }
        }

        // ── Sheet-name / table-name helpers (delegate to shared ExcelNameHelper) ──

        /// <summary>Delegates to <see cref="ExcelNameHelper.UniqueSheetName"/>.</summary>
        internal static string UniqueSheetName(ExcelPackage pkg, string name)
            => ExcelNameHelper.UniqueSheetName(pkg, name);

        /// <summary>Delegates to <see cref="ExcelNameHelper.UniqueTableName"/>.</summary>
        internal static string UniqueTableName(ExcelPackage pkg, string name)
            => ExcelNameHelper.UniqueTableName(pkg, name);

        // ─────────────────────────────────────────────────────────────────────
        // CSV parsing
        // ─────────────────────────────────────────────────────────────────────

        private static (List<ResponseTimeRecord> records, List<string> percentileHeaders)
            ReadCsv(string csvPath)
        {
            var records = new List<ResponseTimeRecord>();
            var percentileHeaders = new List<string>();

            if (!File.Exists(csvPath))
                throw new FileNotFoundException("CSV file not found", csvPath);

            var lines = File.ReadAllLines(csvPath);
            var headers = SplitCsvLine(lines[0]);

            // ── Column indices ────────────────────────────────────────────────
            int labelIndex = Array.IndexOf(headers, "Label");
            int sampleIndex = Array.IndexOf(headers, "# Samples");
            int avgIndex = Array.IndexOf(headers, "Average");
            int medianIndex = Array.IndexOf(headers, "Median");
            int minIndex = Array.IndexOf(headers, "Min");
            int maxIndex = Array.IndexOf(headers, "Max");
            int errIndex = Array.IndexOf(headers, "Error %");

            if (labelIndex < 0 || sampleIndex < 0 || avgIndex < 0 ||
                medianIndex < 0 || minIndex < 0 || maxIndex < 0 || errIndex < 0)
                throw new InvalidDataException(
                    "CSV file is missing one or more required columns (Label, # Samples, Average, Median, Min, Max, Error %).");

            var percentileIndexes = new List<int>();
            for (int i = 0; i < headers.Length; i++)
            {
                if (headers[i].Contains("% Line"))
                {
                    percentileIndexes.Add(i);
                    percentileHeaders.Add(headers[i]);
                }
            }

            // ── Data rows ─────────────────────────────────────────────────────
            for (int i = 1; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;
                var values = SplitCsvLine(lines[i]);
                if (values.Length <= labelIndex) continue;

                // Skip the TOTAL summary row and web request rows (URLs start with /)
                string label = values[labelIndex].Trim();
                if (label.Equals("TOTAL", StringComparison.OrdinalIgnoreCase))
                    continue;
                if (label.StartsWith("/") || label.StartsWith("http://") || label.StartsWith("https://"))
                    continue;

                var record = new ResponseTimeRecord
                {
                    TransactionName = label,
                    Samples = ParseInt(values[sampleIndex]),
                    Average = ToSeconds(values[avgIndex]),
                    Median = ToSeconds(values[medianIndex]),
                    Min = ToSeconds(values[minIndex]),
                    Max = ToSeconds(values[maxIndex]),
                    ErrorPercent = ParsePercent(values[errIndex])
                };

                foreach (var idx in percentileIndexes)
                    record.Percentiles[headers[idx]] = ToSeconds(values[idx]);

                records.Add(record);
            }

            return (records, percentileHeaders);
        }

        // ─────────────────────────────────────────────────────────────────────
        // Parsing helpers
        // ─────────────────────────────────────────────────────────────────────

        private static int ParseInt(string value) =>
            int.TryParse(value, out int result) ? result : 0;

        private static double ParsePercent(string value)
        {
            value = value.Replace("%", "");
            return double.TryParse(
                value,
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out double result) ? result : 0;
        }

        private static double ToSeconds(string ms) =>
            double.TryParse(
                ms,
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out double value) ? value / 1000 : 0;

        /// <summary>Delegates to shared <see cref="CsvHelper.SplitCsvLine"/>.</summary>
        private static string[] SplitCsvLine(string line)
            => CsvHelper.SplitCsvLine(line);

        // ─────────────────────────────────────────────────────────────────────
        // Excel sheet writer
        // ─────────────────────────────────────────────────────────────────────

        private static void WriteResponseSheet(
                ExcelPackage package,
                List<ResponseTimeRecord> records,
                List<string> percentileHeaders,
                string sheetName = "Response Times")
        {
            var sheet = package.Workbook.Worksheets.Add(sheetName);
            int col = 1;

            // ── Header row ────────────────────────────────────────────────────
            sheet.Cells[1, col++].Value = "Transaction Name";
            sheet.Cells[1, col++].Value = "# Samples";
            sheet.Cells[1, col++].Value = "Average (Seconds)";
            sheet.Cells[1, col++].Value = "Median (Seconds)";

            int percentileStartCol = col; // capture BEFORE writing percentile headers

            foreach (var p in percentileHeaders)
                sheet.Cells[1, col++].Value = p.Replace("% Line", " Percentile (Seconds)");

            sheet.Cells[1, col++].Value = "Min (Seconds)";
            sheet.Cells[1, col++].Value = "Max (Seconds)";
            sheet.Cells[1, col++].Value = "Error %";

            // Style the header row
            using (var range = sheet.Cells[1, 1, 1, col - 1])
            {
                range.Style.Font.Bold = true;
                range.Style.Font.Color.SetColor(Color.White);
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x1E, 0x40, 0xAF)); // dark blue
            }

            // ── Data rows ─────────────────────────────────────────────────────
            int rowWidth = 4 + percentileHeaders.Count + 3; // Name + Samples + Avg + Med + percentiles + Min + Max + Err
            int row = 2;
            foreach (var r in records)
            {
                col = 1;

                // Alternating row colour: white / very light grey
                if (row % 2 == 0)
                {
                    sheet.Cells[row, 1, row, rowWidth].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[row, 1, row, rowWidth].Style.Fill.BackgroundColor
                        .SetColor(Color.FromArgb(0xF3, 0xF4, 0xF6));
                }

                sheet.Cells[row, col++].Value = r.TransactionName;
                sheet.Cells[row, col++].Value = r.Samples;
                sheet.Cells[row, col++].Value = r.Average;
                sheet.Cells[row, col++].Value = r.Median;

                foreach (var p in percentileHeaders)
                    sheet.Cells[row, col++].Value = r.Percentiles[p];

                sheet.Cells[row, col++].Value = r.Min;
                sheet.Cells[row, col++].Value = r.Max;

                var errorCell = sheet.Cells[row, col++];
                errorCell.Value = r.ErrorPercent / 100.0;
                errorCell.Style.Numberformat.Format = "0.00%";

                row++;
            }

            sheet.Cells.AutoFitColumns();

            // ── Wrap in an Excel Table ────────────────────────────────────────
            if (records.Count > 0)
            {
                int totalRows = records.Count + 1;
                int totalCols = col - 1;
                var tableRange = sheet.Cells[1, 1, totalRows, totalCols];
                var table = sheet.Tables.Add(tableRange, UniqueTableName(package, "ResponseTimes"));
                table.ShowHeader = true;
                table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
            }
        }
    }
}