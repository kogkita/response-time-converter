using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestApp
{
    // ─────────────────────────────────────────────────────────────────────────
    // Options passed from the UI to the analyzer engine
    // ─────────────────────────────────────────────────────────────────────────

    public class NmonAnalyzerOptions
    {
        /// <summary>Full path to nmon_analyser_v69_2.xlsm</summary>
        public string XlsmPath { get; set; } = string.Empty;

        /// <summary>List of .nmon / .csv files to process</summary>
        public List<string> NmonFiles { get; set; } = new();

        // ── Analyser sheet options ───────────────────────────────────────────

        /// <summary>ALL or LIST</summary>
        public string GraphsScope  { get; set; } = "ALL";

        /// <summary>CHARTS, PICTURES, PRINT or WEB</summary>
        public string GraphsOutput { get; set; } = "CHARTS";

        /// <summary>NO / YES / TOP / KEEP / ONLY</summary>
        public string Merge        { get; set; } = "NO";

        /// <summary>First interval number (1-based). Empty = 1.</summary>
        public string IntervalFirst { get; set; } = "1";

        /// <summary>Last interval number. Empty = 999999.</summary>
        public string IntervalLast  { get; set; } = "999999";

        public bool Ess              { get; set; } = true;
        public bool Scatter          { get; set; } = true;
        public bool BigData          { get; set; } = true;
        public bool ShowLinuxCpuUtil { get; set; } = false;

        // ── Settings sheet options ───────────────────────────────────────────

        public bool   Reorder     { get; set; } = true;
        public bool   SortDefault { get; set; } = true;

        /// <summary>Comma-separated sheet names for LIST mode</summary>
        public string List   { get; set; } = "CPU_ALL,DISKBUSY,ESS*,EMC*,FAST*,LPAR,MEM*,NET,PAGE,PROC,TOP";

        /// <summary>Output directory for Excel files. Empty = same as input.</summary>
        public string OutDir { get; set; } = string.Empty;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // NmonAnalyzer
    //
    // Automates nmon_analyser_v69_2.xlsm by:
    //   1. Writing a FILELIST temp file listing the .nmon input files
    //   2. Writing a VBScript that opens the XLSM, sets all options on the
    //      Analyser and Settings sheets via cell writes, then calls the "Main"
    //      macro (exactly as documented in the nmon Analyser user guide)
    //   3. Shelling out to cscript.exe to execute the VBScript
    //
    // No COM interop or Office SDK dependency is required.
    // ─────────────────────────────────────────────────────────────────────────

    public static class NmonAnalyzer
    {
        public static void Run(NmonAnalyzerOptions opts)
        {
            if (!File.Exists(opts.XlsmPath))
                throw new FileNotFoundException("nmon_analyser_v69_2.xlsm not found.", opts.XlsmPath);

            if (opts.NmonFiles.Count == 0)
                throw new InvalidOperationException("No .nmon files specified.");

            string tmpDir = Path.Combine(Path.GetTempPath(), "NmonAnalyzer_" + Guid.NewGuid().ToString("N")[..8]);
            Directory.CreateDirectory(tmpDir);

            try
            {
                // ── Step 1: write FILELIST ────────────────────────────────────
                string fileListPath = Path.Combine(tmpDir, "filelist.txt");
                File.WriteAllLines(fileListPath, opts.NmonFiles);

                // ── Step 2: write VBScript ────────────────────────────────────
                string vbsPath = Path.Combine(tmpDir, "run_nmon.vbs");
                string vbsContent = BuildVbScript(opts, fileListPath);
                File.WriteAllText(vbsPath, vbsContent, Encoding.UTF8);

                // ── Step 3: execute via cscript ───────────────────────────────
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName  = "cscript.exe",
                    Arguments = $"//NoLogo \"{vbsPath}\"",
                    UseShellExecute        = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError  = true,
                    CreateNoWindow         = true
                };

                using var proc = System.Diagnostics.Process.Start(psi)
                    ?? throw new InvalidOperationException("Failed to start cscript.exe");

                string stdout = proc.StandardOutput.ReadToEnd();
                string stderr = proc.StandardError.ReadToEnd();
                proc.WaitForExit();

                if (proc.ExitCode != 0)
                {
                    string detail = string.IsNullOrWhiteSpace(stderr) ? stdout : stderr;
                    throw new InvalidOperationException(
                        $"nmon Analyser exited with code {proc.ExitCode}.\n\n{detail}");
                }
            }
            finally
            {
                // Clean up temp files (leave for a short while in case of error review)
                try { Directory.Delete(tmpDir, recursive: true); } catch { }
            }
        }

        // ── VBScript builder ──────────────────────────────────────────────────

        private static string BuildVbScript(NmonAnalyzerOptions opts, string fileListPath)
        {
            string yesNo(bool b) => b ? "YES" : "NO";

            // Escape backslashes for VBS string literals
            string xlsm = opts.XlsmPath.Replace("\\", "\\\\");
            string fl   = fileListPath.Replace("\\", "\\\\");
            string od   = opts.OutDir.Replace("\\", "\\\\");

            var sb = new StringBuilder();
            sb.AppendLine("Option Explicit");
            sb.AppendLine("On Error Resume Next");
            sb.AppendLine();
            sb.AppendLine("Dim xlApp, xlBook, ws1, ws2");
            sb.AppendLine("Set xlApp  = CreateObject(\"Excel.Application\")");
            sb.AppendLine($"Set xlBook = xlApp.Workbooks.Open(\"{xlsm}\", 0, False)");
            sb.AppendLine("xlApp.Visible = False");
            sb.AppendLine("xlApp.DisplayAlerts = False");
            sb.AppendLine();
            sb.AppendLine("If Err.Number <> 0 Then");
            sb.AppendLine("  WScript.Echo \"ERROR opening XLSM: \" & Err.Description");
            sb.AppendLine("  WScript.Quit 1");
            sb.AppendLine("End If");
            sb.AppendLine();

            // ── Write Analyser sheet (sheet 1) options ────────────────────────
            sb.AppendLine("Set ws1 = xlBook.Worksheets(\"Analyser\")");
            sb.AppendLine();
            sb.AppendLine($"ws1.Range(\"B10\").Value = \"{opts.GraphsScope}\"");
            sb.AppendLine($"ws1.Range(\"C10\").Value = \"{opts.GraphsOutput}\"");
            sb.AppendLine($"ws1.Range(\"B11\").Value = \"{opts.IntervalFirst}\"");
            sb.AppendLine($"ws1.Range(\"C11\").Value = \"{opts.IntervalLast}\"");
            sb.AppendLine($"ws1.Range(\"B13\").Value = \"{opts.Merge}\"");
            sb.AppendLine($"ws1.Range(\"B15\").Value = \"{yesNo(opts.Scatter)}\"");
            sb.AppendLine($"ws1.Range(\"B16\").Value = \"{yesNo(opts.BigData)}\"");
            sb.AppendLine($"ws1.Range(\"B17\").Value = \"{yesNo(opts.Ess)}\"");
            sb.AppendLine($"ws1.Range(\"B20\").Value = \"{yesNo(opts.ShowLinuxCpuUtil)}\"");
            sb.AppendLine($"ws1.Range(\"B21\").Value = \"{fl}\"");  // FILELIST
            sb.AppendLine();

            // ── Write Settings sheet (sheet 2) options ────────────────────────
            sb.AppendLine("Set ws2 = xlBook.Worksheets(\"Settings\")");
            sb.AppendLine();
            sb.AppendLine($"ws2.Range(\"B8\").Value  = \"{EscapeVbs(opts.List)}\"");        // LIST
            sb.AppendLine($"ws2.Range(\"B11\").Value = \"{yesNo(opts.Reorder)}\"");         // REORDER
            sb.AppendLine($"ws2.Range(\"B14\").Value = \"{yesNo(opts.SortDefault)}\"");     // SORTDEFAULT
            if (!string.IsNullOrEmpty(opts.OutDir))
                sb.AppendLine($"ws2.Range(\"B3\").Value  = \"{od}\\\\\"");                  // OUTDIR
            sb.AppendLine();

            // ── Run the Main macro ────────────────────────────────────────────
            sb.AppendLine("Err.Clear");
            sb.AppendLine("xlApp.Run \"Main\"");
            sb.AppendLine();
            sb.AppendLine("If Err.Number <> 0 Then");
            sb.AppendLine("  WScript.Echo \"ERROR running Main macro: \" & Err.Description");
            sb.AppendLine("  xlApp.Quit");
            sb.AppendLine("  WScript.Quit 1");
            sb.AppendLine("End If");
            sb.AppendLine();
            sb.AppendLine("xlApp.Quit");
            sb.AppendLine("Set xlBook = Nothing");
            sb.AppendLine("Set xlApp  = Nothing");
            sb.AppendLine("WScript.Quit 0");

            return sb.ToString();
        }

        private static string EscapeVbs(string s)
            => s.Replace("\"", "\"\"").Replace("\\", "\\\\");
    }
}
