using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System.IO;

namespace TestApp
{
    // ── Shared metric bag used for both CSV and JTL runs ─────────────────────

    public class ComparisonRecord
    {
        public string TransactionName { get; set; } = string.Empty;

        // Baseline
        public double BaseAvg    { get; set; }
        public double BaseMedian { get; set; }
        public double BaseP90    { get; set; }
        public double BaseMin    { get; set; }
        public double BaseMax    { get; set; }
        public double BaseErrors { get; set; }   // 0-100
        public int    BaseSamples{ get; set; }

        // Current
        public double CurAvg    { get; set; }
        public double CurMedian { get; set; }
        public double CurP90    { get; set; }
        public double CurMin    { get; set; }
        public double CurMax    { get; set; }
        public double CurErrors { get; set; }
        public int    CurSamples{ get; set; }

        // Delta helpers (positive = current is SLOWER / worse)
        public double DeltaAvgMs    => CurAvg    - BaseAvg;
        public double DeltaMedianMs => CurMedian - BaseMedian;
        public double DeltaP90Ms    => CurP90    - BaseP90;
        public double DeltaErrorPct => CurErrors - BaseErrors;

        public double DeltaAvgPct    => BaseAvg    > 0 ? (CurAvg    - BaseAvg)    / BaseAvg    * 100 : 0;
        public double DeltaMedianPct => BaseMedian > 0 ? (CurMedian - BaseMedian) / BaseMedian * 100 : 0;
        public double DeltaP90Pct    => BaseP90    > 0 ? (CurP90    - BaseP90)    / BaseP90    * 100 : 0;

        public bool OnlyInBaseline   { get; set; }
        public bool OnlyInCurrent    { get; set; }
    }

    // ── Input type ────────────────────────────────────────────────────────────

    public enum ComparisonFileType { Csv, Jtl }

    // ── Main processing class ─────────────────────────────────────────────────

    public static class RunComparisonProcessor
    {
        // Regression threshold: flag if avg worsens by more than this %
        private const double RegressionThresholdPct = 10.0;

        public static void Compare(
            string baselinePath,
            string currentPath,
            string outputPath,
            ComparisonFileType fileType,
            double slaThresholdMs = 0)
        {
            ExcelPackage.License.SetNonCommercialPersonal("Run Comparison");

            var baseRecords = LoadRecords(baselinePath, fileType);
            var curRecords  = LoadRecords(currentPath,  fileType);

            var comparison = BuildComparison(baseRecords, curRecords);

            using var package = new ExcelPackage();

            WriteSummarySheet(package, comparison, baselinePath, currentPath, slaThresholdMs);
            WriteDetailSheet(package, comparison, slaThresholdMs);
            WriteRawSheets(package, baseRecords, curRecords, fileType);

            package.SaveAs(new FileInfo(outputPath));
        }

        // ── Record loading ────────────────────────────────────────────────────

        private static List<FlatRecord> LoadRecords(string path, ComparisonFileType fileType)
        {
            if (fileType == ComparisonFileType.Jtl)
            {
                return JTLFileProcessing.ParseJtl(path)
                    .Select(r => new FlatRecord
                    {
                        Name    = r.TransactionName,
                        Average = r.Average,
                        Median  = r.Median,
                        P90     = r.P90,
                        Min     = r.Min,
                        Max     = r.Max,
                        Errors  = r.ErrorPercent,
                        Samples = r.Samples
                    }).ToList();
            }
            else
            {
                return LoadCsvRecords(path);
            }
        }

        private static List<FlatRecord> LoadCsvRecords(string csvPath)
        {
            if (!File.Exists(csvPath))
                throw new FileNotFoundException("CSV file not found", csvPath);

            var lines   = File.ReadAllLines(csvPath);
            var headers = lines[0].Split(',');

            int labelIdx  = Array.IndexOf(headers, "Label");
            int sampIdx   = Array.IndexOf(headers, "# Samples");
            int avgIdx    = Array.IndexOf(headers, "Average");
            int medIdx    = Array.IndexOf(headers, "Median");
            int minIdx    = Array.IndexOf(headers, "Min");
            int maxIdx    = Array.IndexOf(headers, "Max");
            int errIdx    = Array.IndexOf(headers, "Error %");
            int p90Idx    = -1;
            for (int i = 0; i < headers.Length; i++)
                if (headers[i].Contains("90%") || headers[i].Contains("90 %"))
                { p90Idx = i; break; }

            var records = new List<FlatRecord>();
            for (int i = 1; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;
                var v = lines[i].Split(',');
                string label = v[labelIdx].Trim();
                if (label.Equals("TOTAL", StringComparison.OrdinalIgnoreCase)) continue;
                if (label.StartsWith("/") || label.StartsWith("http")) continue;

                records.Add(new FlatRecord
                {
                    Name    = label,
                    Samples = ParseInt(v[sampIdx]),
                    Average = ParseMs(v[avgIdx]),
                    Median  = ParseMs(v[medIdx]),
                    P90     = p90Idx >= 0 ? ParseMs(v[p90Idx]) : 0,
                    Min     = ParseMs(v[minIdx]),
                    Max     = ParseMs(v[maxIdx]),
                    Errors  = ParsePct(v[errIdx])
                });
            }
            return records;
        }

        // ── Delta computation ─────────────────────────────────────────────────

        private static List<ComparisonRecord> BuildComparison(
            List<FlatRecord> baseline, List<FlatRecord> current)
        {
            var baseDict = baseline.ToDictionary(
                r => r.Name, StringComparer.OrdinalIgnoreCase);
            var curDict  = current.ToDictionary(
                r => r.Name, StringComparer.OrdinalIgnoreCase);

            var allNames = baseDict.Keys.Union(curDict.Keys, StringComparer.OrdinalIgnoreCase)
                           .OrderBy(n => n, StringComparer.Ordinal)
                           .ToList();

            var result = new List<ComparisonRecord>();
            foreach (var name in allNames)
            {
                bool inBase = baseDict.TryGetValue(name, out var b);
                bool inCur  = curDict.TryGetValue(name, out var c);

                var rec = new ComparisonRecord
                {
                    TransactionName = name,
                    OnlyInBaseline  = inBase && !inCur,
                    OnlyInCurrent   = !inBase && inCur,

                    BaseAvg    = inBase ? b!.Average : 0,
                    BaseMedian = inBase ? b!.Median  : 0,
                    BaseP90    = inBase ? b!.P90     : 0,
                    BaseMin    = inBase ? b!.Min     : 0,
                    BaseMax    = inBase ? b!.Max     : 0,
                    BaseErrors = inBase ? b!.Errors  : 0,
                    BaseSamples= inBase ? b!.Samples : 0,

                    CurAvg    = inCur ? c!.Average : 0,
                    CurMedian = inCur ? c!.Median  : 0,
                    CurP90    = inCur ? c!.P90     : 0,
                    CurMin    = inCur ? c!.Min     : 0,
                    CurMax    = inCur ? c!.Max     : 0,
                    CurErrors = inCur ? c!.Errors  : 0,
                    CurSamples= inCur ? c!.Samples : 0,
                };
                result.Add(rec);
            }
            return result;
        }

        // ── Summary sheet ─────────────────────────────────────────────────────

        private static void WriteSummarySheet(
            ExcelPackage pkg,
            List<ComparisonRecord> rows,
            string baselinePath,
            string currentPath,
            double slaMs)
        {
            var ws = pkg.Workbook.Worksheets.Add("Summary");

            // ── Header block ──────────────────────────────────────────────────
            ws.Cells[1, 1].Value = "Run Comparison Report";
            ws.Cells[1, 1].Style.Font.Bold = true;
            ws.Cells[1, 1].Style.Font.Size = 16;

            ws.Cells[2, 1].Value = "Baseline";
            ws.Cells[2, 2].Value = Path.GetFileName(baselinePath);
            ws.Cells[3, 1].Value = "Current";
            ws.Cells[3, 2].Value = Path.GetFileName(currentPath);
            ws.Cells[4, 1].Value = "Generated";
            ws.Cells[4, 2].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            if (slaMs > 0)
            {
                ws.Cells[5, 1].Value = "SLA Threshold";
                ws.Cells[5, 2].Value = $"{slaMs:0} ms";
            }

            using (var r = ws.Cells[2, 1, slaMs > 0 ? 5 : 4, 1])
            {
                r.Style.Font.Bold = true;
                r.Style.Font.Color.SetColor(Color.FromArgb(0x8B, 0x93, 0xA5));
            }

            // ── KPI band ──────────────────────────────────────────────────────
            int kpiRow = 7;
            var matched = rows.Where(r => !r.OnlyInBaseline && !r.OnlyInCurrent).ToList();

            var regressions = matched
                .Where(r => r.DeltaAvgPct > RegressionThresholdPct)
                .OrderByDescending(r => r.DeltaAvgPct)
                .ToList();
            var improvements = matched
                .Where(r => r.DeltaAvgPct < -RegressionThresholdPct)
                .OrderBy(r => r.DeltaAvgPct)
                .ToList();

            WriteKpi(ws, kpiRow, 1, "Transactions Compared", matched.Count.ToString(),
                Color.FromArgb(0x1E, 0x29, 0x4A));
            WriteKpi(ws, kpiRow, 3, "Regressions (>10%)",    regressions.Count.ToString(),
                regressions.Count > 0 ? Color.FromArgb(0x7F, 0x1D, 0x1D) : Color.FromArgb(0x14, 0x2A, 0x1E));
            WriteKpi(ws, kpiRow, 5, "Improvements (>10%)",   improvements.Count.ToString(),
                Color.FromArgb(0x14, 0x2A, 0x1E));
            WriteKpi(ws, kpiRow, 7, "Only in Baseline",
                rows.Count(r => r.OnlyInBaseline).ToString(), Color.FromArgb(0x2A, 0x20, 0x10));
            WriteKpi(ws, kpiRow, 9, "Only in Current",
                rows.Count(r => r.OnlyInCurrent).ToString(),  Color.FromArgb(0x2A, 0x20, 0x10));

            // ── Top regressions table ─────────────────────────────────────────
            int tableRow = kpiRow + 5;
            if (regressions.Count > 0)
            {
                ws.Cells[tableRow, 1].Value = "Top Regressions";
                ws.Cells[tableRow, 1].Style.Font.Bold = true;
                ws.Cells[tableRow, 1].Style.Font.Size = 13;
                tableRow++;

                ws.Cells[tableRow, 1].Value = "Transaction";
                ws.Cells[tableRow, 2].Value = "Baseline Avg (s)";
                ws.Cells[tableRow, 3].Value = "Current Avg (s)";
                ws.Cells[tableRow, 4].Value = "Delta (ms)";
                ws.Cells[tableRow, 5].Value = "Delta %";
                StyleHeader(ws.Cells[tableRow, 1, tableRow, 5], Color.FromArgb(0x7F, 0x1D, 0x1D));
                tableRow++;

                foreach (var rec in regressions.Take(10))
                {
                    ws.Cells[tableRow, 1].Value = rec.TransactionName;
                    ws.Cells[tableRow, 2].Value = Math.Round(rec.BaseAvg / 1000.0, 3);
                    ws.Cells[tableRow, 3].Value = Math.Round(rec.CurAvg  / 1000.0, 3);
                    ws.Cells[tableRow, 4].Value = Math.Round(rec.DeltaAvgMs, 0);
                    ws.Cells[tableRow, 5].Value = rec.DeltaAvgPct / 100.0;
                    ws.Cells[tableRow, 5].Style.Numberformat.Format = "+0.00%;-0.00%";
                    ws.Cells[tableRow, 1, tableRow, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[tableRow, 1, tableRow, 5].Style.Fill.BackgroundColor
                        .SetColor(Color.FromArgb(0x3B, 0x0D, 0x0D));
                    tableRow++;
                }
                tableRow++;
            }

            // ── SLA breaches ──────────────────────────────────────────────────
            if (slaMs > 0)
            {
                var slaBreaches = matched.Where(r => r.CurP90 > slaMs)
                    .OrderByDescending(r => r.CurP90).ToList();
                if (slaBreaches.Count > 0)
                {
                    ws.Cells[tableRow, 1].Value = $"SLA Breaches — P90 > {slaMs:0} ms";
                    ws.Cells[tableRow, 1].Style.Font.Bold = true;
                    ws.Cells[tableRow, 1].Style.Font.Size = 13;
                    tableRow++;

                    ws.Cells[tableRow, 1].Value = "Transaction";
                    ws.Cells[tableRow, 2].Value = "Baseline P90 (s)";
                    ws.Cells[tableRow, 3].Value = "Current P90 (s)";
                    ws.Cells[tableRow, 4].Value = "SLA (ms)";
                    StyleHeader(ws.Cells[tableRow, 1, tableRow, 4], Color.FromArgb(0x78, 0x35, 0x0F));
                    tableRow++;

                    foreach (var rec in slaBreaches)
                    {
                        ws.Cells[tableRow, 1].Value = rec.TransactionName;
                        ws.Cells[tableRow, 2].Value = Math.Round(rec.BaseP90 / 1000.0, 3);
                        ws.Cells[tableRow, 3].Value = Math.Round(rec.CurP90  / 1000.0, 3);
                        ws.Cells[tableRow, 4].Value = slaMs;
                        ws.Cells[tableRow, 1, tableRow, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[tableRow, 1, tableRow, 4].Style.Fill.BackgroundColor
                            .SetColor(Color.FromArgb(0x3D, 0x1A, 0x08));
                        tableRow++;
                    }
                }
            }

            ws.Cells.AutoFitColumns(12, 60);
        }

        private static void WriteKpi(
            ExcelWorksheet ws, int row, int col,
            string label, string value, Color bg)
        {
            var labelCell = ws.Cells[row, col, row, col + 1];
            labelCell.Merge = true;
            labelCell.Value = label;
            labelCell.Style.Font.Size = 10;
            labelCell.Style.Font.Color.SetColor(Color.FromArgb(0x9C, 0xA3, 0xAF));
            labelCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            labelCell.Style.Fill.BackgroundColor.SetColor(bg);
            labelCell.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

            var valCell = ws.Cells[row + 1, col, row + 1, col + 1];
            valCell.Merge = true;
            valCell.Value = value;
            valCell.Style.Font.Size = 20;
            valCell.Style.Font.Bold = true;
            valCell.Style.Font.Color.SetColor(Color.FromArgb(0xE2, 0xE8, 0xF0));
            valCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            valCell.Style.Fill.BackgroundColor.SetColor(bg);
            valCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            valCell.Style.VerticalAlignment   = ExcelVerticalAlignment.Top;

            // Light border around the KPI box
            var box = ws.Cells[row, col, row + 1, col + 1];
            box.Style.Border.BorderAround(ExcelBorderStyle.Thin,
                Color.FromArgb(0x2A, 0x2F, 0x3E));
        }

        private static void StyleHeader(ExcelRange range, Color bg)
        {
            range.Style.Font.Bold = true;
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(bg);
            range.Style.Font.Color.SetColor(Color.FromArgb(0xE2, 0xE8, 0xF0));
        }

        // ── Detail delta sheet ────────────────────────────────────────────────

        private static void WriteDetailSheet(
            ExcelPackage pkg, List<ComparisonRecord> rows, double slaMs)
        {
            var ws = pkg.Workbook.Worksheets.Add("Delta Detail");

            // Column headers
            var headers = new[]
            {
                "Transaction",
                "Status",
                "Base Avg (s)", "Cur Avg (s)",  "Δ Avg (ms)",  "Δ Avg %",
                "Base P90 (s)", "Cur P90 (s)",   "Δ P90 (ms)",  "Δ P90 %",
                "Base Med (s)", "Cur Med (s)",   "Δ Med (ms)",
                "Base Err %",  "Cur Err %",     "Δ Err pts",
                "Base Samples","Cur Samples"
            };
            if (slaMs > 0) headers = headers.Append("SLA Breach").ToArray();

            for (int c = 0; c < headers.Length; c++)
                ws.Cells[1, c + 1].Value = headers[c];

            using (var hdr = ws.Cells[1, 1, 1, headers.Length])
            {
                hdr.Style.Font.Bold = true;
                hdr.Style.Fill.PatternType = ExcelFillStyle.Solid;
                hdr.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x1A, 0x1F, 0x2E));
                hdr.Style.Font.Color.SetColor(Color.FromArgb(0xE2, 0xE8, 0xF0));
            }

            int row = 2;
            foreach (var r in rows)
            {
                string status =
                    r.OnlyInBaseline ? "Removed" :
                    r.OnlyInCurrent  ? "New" :
                    r.DeltaAvgPct > RegressionThresholdPct  ? "Regression" :
                    r.DeltaAvgPct < -RegressionThresholdPct ? "Improvement" :
                    "Stable";

                Color rowColor = status switch
                {
                    "Regression"  => Color.FromArgb(0x3B, 0x0D, 0x0D),
                    "Improvement" => Color.FromArgb(0x0D, 0x2B, 0x17),
                    "New"         => Color.FromArgb(0x12, 0x20, 0x3A),
                    "Removed"     => Color.FromArgb(0x2A, 0x1A, 0x08),
                    _             => Color.FromArgb(0x0F, 0x11, 0x17)
                };

                int c = 1;
                ws.Cells[row, c++].Value = r.TransactionName;
                ws.Cells[row, c++].Value = status;

                ws.Cells[row, c++].Value = r.OnlyInCurrent   ? (object)"—" : Math.Round(r.BaseAvg    / 1000.0, 3);
                ws.Cells[row, c++].Value = r.OnlyInBaseline  ? (object)"—" : Math.Round(r.CurAvg     / 1000.0, 3);
                ws.Cells[row, c++].Value = (r.OnlyInBaseline || r.OnlyInCurrent) ? (object)"—" : Math.Round(r.DeltaAvgMs,    0);

                var pctCell = ws.Cells[row, c++];
                if (r.OnlyInBaseline || r.OnlyInCurrent) pctCell.Value = "—";
                else { pctCell.Value = r.DeltaAvgPct / 100.0; pctCell.Style.Numberformat.Format = "+0.00%;-0.00%"; }

                ws.Cells[row, c++].Value = r.OnlyInCurrent   ? (object)"—" : Math.Round(r.BaseP90    / 1000.0, 3);
                ws.Cells[row, c++].Value = r.OnlyInBaseline  ? (object)"—" : Math.Round(r.CurP90     / 1000.0, 3);
                ws.Cells[row, c++].Value = (r.OnlyInBaseline || r.OnlyInCurrent) ? (object)"—" : Math.Round(r.DeltaP90Ms,    0);

                var p90PctCell = ws.Cells[row, c++];
                if (r.OnlyInBaseline || r.OnlyInCurrent) p90PctCell.Value = "—";
                else { p90PctCell.Value = r.DeltaP90Pct / 100.0; p90PctCell.Style.Numberformat.Format = "+0.00%;-0.00%"; }

                ws.Cells[row, c++].Value = r.OnlyInCurrent   ? (object)"—" : Math.Round(r.BaseMedian / 1000.0, 3);
                ws.Cells[row, c++].Value = r.OnlyInBaseline  ? (object)"—" : Math.Round(r.CurMedian  / 1000.0, 3);
                ws.Cells[row, c++].Value = (r.OnlyInBaseline || r.OnlyInCurrent) ? (object)"—" : Math.Round(r.DeltaMedianMs, 0);

                ws.Cells[row, c++].Value = r.OnlyInCurrent   ? (object)"—" : r.BaseErrors / 100.0;
                ws.Cells[row, c++].Value = r.OnlyInBaseline  ? (object)"—" : r.CurErrors  / 100.0;
                ws.Cells[row, c++].Value = (r.OnlyInBaseline || r.OnlyInCurrent) ? (object)"—" : Math.Round(r.DeltaErrorPct, 2);

                ws.Cells[row, c++].Value = r.OnlyInCurrent  ? (object)"—" : r.BaseSamples;
                ws.Cells[row, c++].Value = r.OnlyInBaseline ? (object)"—" : r.CurSamples;

                if (slaMs > 0)
                    ws.Cells[row, c++].Value =
                        !r.OnlyInBaseline && r.CurP90 > slaMs ? "YES" : "";

                // Format error % columns
                int errBaseCol = Array.IndexOf(headers, "Base Err %") + 1;
                int errCurCol  = errBaseCol + 1;
                if (ws.Cells[row, errBaseCol].Value is double)
                    ws.Cells[row, errBaseCol].Style.Numberformat.Format = "0.00%";
                if (ws.Cells[row, errCurCol].Value is double)
                    ws.Cells[row, errCurCol].Style.Numberformat.Format = "0.00%";

                // Row background
                ws.Cells[row, 1, row, headers.Length].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 1, row, headers.Length].Style.Fill.BackgroundColor.SetColor(rowColor);

                // Status cell colour accent
                var statusColor = status switch
                {
                    "Regression"  => Color.FromArgb(0xF8, 0x71, 0x71),
                    "Improvement" => Color.FromArgb(0x6E, 0xE7, 0xB7),
                    "New"         => Color.FromArgb(0x93, 0xC5, 0xFD),
                    "Removed"     => Color.FromArgb(0xFB, 0xBF, 0x24),
                    _             => Color.FromArgb(0x6B, 0x72, 0x80)
                };
                ws.Cells[row, 2].Style.Font.Color.SetColor(statusColor);
                ws.Cells[row, 2].Style.Font.Bold = true;

                row++;
            }

            ws.Cells.AutoFitColumns(10, 50);

            // Freeze top row
            ws.View.FreezePanes(2, 1);

            // Excel table
            var tableRange = ws.Cells[1, 1, row - 1, headers.Length];
            var table = ws.Tables.Add(tableRange,
                JTLFileProcessing.UniqueTableName(pkg, "DeltaDetail"));
            table.ShowHeader = true;
            table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
        }

        // ── Raw baseline / current sheets ─────────────────────────────────────

        private static void WriteRawSheets(
            ExcelPackage pkg,
            List<FlatRecord> baseline,
            List<FlatRecord> current,
            ComparisonFileType fileType)
        {
            WriteRawSheet(pkg, baseline, "Baseline", fileType);
            WriteRawSheet(pkg, current,  "Current",  fileType);
        }

        private static void WriteRawSheet(
            ExcelPackage pkg,
            List<FlatRecord> records,
            string sheetName,
            ComparisonFileType fileType)
        {
            var ws = pkg.Workbook.Worksheets.Add(sheetName);

            var cols = new[] { "Transaction", "Samples", "Avg (s)", "Median (s)",
                               "P90 (s)", "Min (s)", "Max (s)", "Error %" };

            for (int c = 0; c < cols.Length; c++)
                ws.Cells[1, c + 1].Value = cols[c];

            using (var hdr = ws.Cells[1, 1, 1, cols.Length])
            {
                hdr.Style.Font.Bold = true;
                hdr.Style.Fill.PatternType = ExcelFillStyle.Solid;
                hdr.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x1A, 0x1F, 0x2E));
                hdr.Style.Font.Color.SetColor(Color.FromArgb(0xE2, 0xE8, 0xF0));
            }

            int row = 2;
            foreach (var r in records.OrderBy(x => x.Name, StringComparer.Ordinal))
            {
                ws.Cells[row, 1].Value = r.Name;
                ws.Cells[row, 2].Value = r.Samples;
                ws.Cells[row, 3].Value = Math.Round(r.Average / 1000.0, 3);
                ws.Cells[row, 4].Value = Math.Round(r.Median  / 1000.0, 3);
                ws.Cells[row, 5].Value = Math.Round(r.P90     / 1000.0, 3);
                ws.Cells[row, 6].Value = Math.Round(r.Min     / 1000.0, 3);
                ws.Cells[row, 7].Value = Math.Round(r.Max     / 1000.0, 3);
                ws.Cells[row, 8].Value = r.Errors / 100.0;
                ws.Cells[row, 8].Style.Numberformat.Format = "0.00%";
                row++;
            }

            ws.Cells.AutoFitColumns();
            var table = ws.Tables.Add(ws.Cells[1, 1, row - 1, cols.Length],
                JTLFileProcessing.UniqueTableName(pkg, sheetName + "Raw"));
            table.ShowHeader = true;
            table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
        }

        // ── Parsing helpers ───────────────────────────────────────────────────

        private static int    ParseInt(string v) =>
            int.TryParse(v.Trim(), out var i) ? i : 0;

        private static double ParseMs(string v) =>
            double.TryParse(v.Trim(), NumberStyles.Any,
                CultureInfo.InvariantCulture, out var d) ? d : 0;

        private static double ParsePct(string v) =>
            double.TryParse(v.Replace("%", "").Trim(), NumberStyles.Any,
                CultureInfo.InvariantCulture, out var d) ? d : 0;

        // ── Internal flat record ──────────────────────────────────────────────

        private class FlatRecord
        {
            public string Name    { get; set; } = string.Empty;
            public int    Samples { get; set; }
            public double Average { get; set; }
            public double Median  { get; set; }
            public double P90     { get; set; }
            public double Min     { get; set; }
            public double Max     { get; set; }
            public double Errors  { get; set; }
        }
    }
}
