using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace TestApp
{
    // ─────────────────────────────────────────────────────────────────────────
    //  ScriptParam  — one declared parameter coming from the metadata block
    //                 or from argparse auto-detection
    // ─────────────────────────────────────────────────────────────────────────
    public class ScriptParam
    {
        /// <summary>CLI argument name, e.g. "input" (positional) or "--out" (optional)</summary>
        public string ArgName    { get; set; } = "";

        /// <summary>Human-readable label shown in the UI</summary>
        public string Label      { get; set; } = "";

        /// <summary>
        /// Param type:
        ///   file-in   → open-file picker
        ///   file-out  → save-file picker
        ///   string    → plain textbox
        ///   float     → numeric textbox
        ///   int       → numeric textbox
        ///   flag      → checkbox  (maps to presence/absence of the flag)
        /// </summary>
        public string Type       { get; set; } = "string";

        /// <summary>File dialog filter string, e.g. "HTML files|*.html;*.htm"</summary>
        public string? Filter    { get; set; }

        /// <summary>Default value shown in the field</summary>
        public string? Default   { get; set; }

        /// <summary>Whether the field can be left blank</summary>
        public bool   Optional   { get; set; } = false;

        /// <summary>Current value — set by the dynamic UI and read at run time</summary>
        public string  Value     { get; set; } = "";
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  ScriptParamDetector  — parses a script file and returns ScriptParam list
    // ─────────────────────────────────────────────────────────────────────────
    public static class ScriptParamDetector
    {
        // ── Public entry point ────────────────────────────────────────────────

        /// <summary>
        /// Attempts to detect parameters from <paramref name="scriptPath"/>.
        /// Returns an empty list if nothing can be found.
        /// </summary>
        public static List<ScriptParam> Detect(string scriptPath)
        {
            if (!File.Exists(scriptPath)) return [];

            string ext = Path.GetExtension(scriptPath).ToLowerInvariant();

            // Only Python and PowerShell supported for now
            if (ext != ".py" && ext != ".ps1") return [];

            try
            {
                var lines = File.ReadAllLines(scriptPath);

                // Priority 1: explicit metadata block
                var fromBlock = ParseMetadataBlock(lines, ext);
                if (fromBlock.Count > 0) return fromBlock;

                // Priority 2: argparse / param block auto-detection
                if (ext == ".py")  return ScanArgparse(lines);
                if (ext == ".ps1") return ScanPowerShellParams(lines);
            }
            catch { /* never crash the host app */ }

            return [];
        }

        // ── Metadata block parser ─────────────────────────────────────────────
        //
        //  Python  convention (comment lines at the top of the file):
        //
        //    ## SCRIPT_RUNNER
        //    ## param: input,         label=Input HTML File,  type=file-in,  filter=HTML|*.html;*.htm
        //    ## param: --out,         label=Output File,      type=file-out, filter=HTML|*.html,   optional=true
        //    ## param: --min-query-time, label=Min Query Time (s), type=float, default=0.05, optional=true
        //    ## param: --keep-rtime-graph, label=Keep rtime Graph, type=flag,  optional=true
        //
        //  PowerShell convention (same, but with #):
        //
        //    # SCRIPT_RUNNER
        //    # param: -InputFile, label=Input File, type=file-in, ...

        private static List<ScriptParam> ParseMetadataBlock(string[] lines, string ext)
        {
            string prefix  = ext == ".ps1" ? "# " : "## ";
            string header  = (prefix + "SCRIPT_RUNNER").TrimEnd();
            string paramKw = prefix + "param:";

            var result = new List<ScriptParam>();
            bool inBlock = false;

            foreach (var raw in lines)
            {
                var line = raw.Trim();

                if (!inBlock)
                {
                    if (line == header) { inBlock = true; }
                    continue;
                }

                // End of block — first non-comment line after the header
                if (!line.StartsWith(prefix.TrimEnd()))
                    break;

                if (!line.StartsWith(paramKw, StringComparison.OrdinalIgnoreCase))
                    continue;

                var body  = line[paramKw.Length..].Trim();
                var param = ParseParamDeclaration(body);
                if (param != null) result.Add(param);
            }

            return result;
        }

        // Parses a single param declaration:
        //   input, label=Input HTML File, type=file-in, filter=HTML files|*.html
        private static ScriptParam? ParseParamDeclaration(string decl)
        {
            // Split on commas that are NOT inside a filter value  
            // (filter values look like "HTML|*.html;*.htm" — they may contain commas)
            var parts = SplitDeclaration(decl);
            if (parts.Count == 0) return null;

            var param = new ScriptParam { ArgName = parts[0].Trim() };

            foreach (var part in parts.Skip(1))
            {
                var eq  = part.IndexOf('=');
                if (eq < 0) continue;
                var key = part[..eq].Trim().ToLowerInvariant();
                var val = part[(eq + 1)..].Trim();

                switch (key)
                {
                    case "label":    param.Label    = val; break;
                    case "type":     param.Type     = val.ToLowerInvariant(); break;
                    case "filter":   param.Filter   = val; break;
                    case "default":  param.Default  = val;
                                     param.Value    = val; break;
                    case "optional": param.Optional = val.Equals("true", StringComparison.OrdinalIgnoreCase); break;
                }
            }

            // Auto-fill label if not given
            if (string.IsNullOrEmpty(param.Label))
                param.Label = param.ArgName.TrimStart('-').Replace('-', ' ').ToTitleCase();

            return param;
        }

        // Split "a, b=x, c=HTML|*.html;*.htm, d=y" into ["a","b=x","c=HTML|*.html;*.htm","d=y"]
        // by treating the filter value (contains |) as opaque
        private static List<string> SplitDeclaration(string s)
        {
            var parts  = new List<string>();
            int start  = 0;
            bool inVal = false;   // once we've seen = in this segment, the rest is the value

            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == '=') inVal = true;
                if (s[i] == ',' && !inVal)
                {
                    parts.Add(s[start..i]);
                    start  = i + 1;
                    inVal  = false;
                }
            }
            if (start < s.Length) parts.Add(s[start..]);
            return parts;
        }

        // ── argparse auto-detection (Python) ─────────────────────────────────
        //
        //  Scans for:
        //    parser.add_argument('input', ...)
        //    parser.add_argument('--out', ...)
        //    parser.add_argument('--flag', action='store_true', ...)
        //    parser.add_argument('--value', type=float, ...)
        //
        //  Works for the common single-file argparse pattern.

        private static readonly Regex RxAddArg = new(
            @"add_argument\(\s*(?:'([^']+)'|""([^""]+)"")",
            RegexOptions.Compiled);

        private static readonly Regex RxKwarg = new(
            @"(\w+)\s*=\s*(?:'([^']*)'|""([^""]*)""|(\S+?)(?:[,\)]|$))",
            RegexOptions.Compiled);

        private static List<ScriptParam> ScanArgparse(string[] lines)
        {
            var result = new List<ScriptParam>();

            // Join continuation lines (argparse calls often span multiple lines)
            var joined = string.Join(" ", lines);

            foreach (Match m in RxAddArg.Matches(joined))
            {
                var argName = (m.Groups[1].Value + m.Groups[2].Value).Trim();
                if (string.IsNullOrEmpty(argName)) continue;

                // Grab the kwargs that follow this add_argument call
                int start   = m.Index + m.Length;
                int parenD  = 0;
                int end     = start;
                for (; end < joined.Length; end++)
                {
                    if (joined[end] == '(') parenD++;
                    if (joined[end] == ')') { if (parenD-- <= 0) break; }
                }
                var kwargs = joined[start..end];

                var param = new ScriptParam { ArgName = argName };
                param.Label   = argName.TrimStart('-').Replace('-', ' ').Replace('_', ' ').ToTitleCase();
                param.Optional = argName.StartsWith('-');

                // Inspect kwargs
                foreach (Match kw in RxKwarg.Matches(kwargs))
                {
                    var k = kw.Groups[1].Value.ToLowerInvariant();
                    var v = (kw.Groups[2].Value + kw.Groups[3].Value + kw.Groups[4].Value).Trim();

                    switch (k)
                    {
                        case "help":
                            if (!string.IsNullOrEmpty(v)) param.Label = v.ToTitleCase();
                            break;
                        case "default":
                            if (v != "None") { param.Default = v; param.Value = v; }
                            break;
                        case "type":
                            param.Type = v switch
                            {
                                "float"            => "float",
                                "int"              => "int",
                                "argparse.FileType" or "FileType" => "file-in",
                                _ => "string"
                            };
                            break;
                        case "action":
                            if (v == "store_true" || v == "store_false") param.Type = "flag";
                            break;
                    }
                }

                // Heuristic: classify by argument name when type is still "string"
                if (param.Type == "string")
                {
                    var lower = argName.ToLowerInvariant();
                    if (IsInputHint(lower))  param.Type = "file-in";
                    if (IsOutputHint(lower)) param.Type = "file-out";
                }

                result.Add(param);
            }

            return result;
        }

        // ── PowerShell param block auto-detection ─────────────────────────────
        //
        //  Scans for:
        //    [Parameter()] [string] $InputFile
        //    [Parameter(Mandatory)] [string] $OutputPath

        private static readonly Regex RxPsParam = new(
            @"\[\s*Parameter[^\]]*\]\s*(?:\[[^\]]+\]\s*)?\$(\w+)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static List<ScriptParam> ScanPowerShellParams(string[] lines)
        {
            var result = new List<ScriptParam>();
            bool inParamBlock = false;

            foreach (var raw in lines)
            {
                var line = raw.Trim();
                if (line.Equals("param(", StringComparison.OrdinalIgnoreCase) ||
                    line.Equals("param (", StringComparison.OrdinalIgnoreCase))
                {
                    inParamBlock = true; continue;
                }
                if (!inParamBlock) continue;
                if (line == ")") break;

                var m = RxPsParam.Match(line);
                if (!m.Success) continue;

                var name  = "-" + m.Groups[1].Value;
                var lower = name.ToLowerInvariant();
                var param = new ScriptParam
                {
                    ArgName  = name,
                    Label    = m.Groups[1].Value.ToTitleCase(),
                    Type     = IsInputHint(lower) ? "file-in"
                             : IsOutputHint(lower) ? "file-out"
                             : "string",
                    Optional = !line.Contains("Mandatory", StringComparison.OrdinalIgnoreCase)
                };
                result.Add(param);
            }

            return result;
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private static bool IsInputHint(string name) =>
            name.Contains("input")  || name.Contains("infile") ||
            name.Contains("source") || name.Contains("src")    ||
            name.Contains("file")   || (name == "path" || name == "filepath");

        private static bool IsOutputHint(string name) =>
            name.Contains("output") || name.Contains("outfile") ||
            name.Contains("out")    || name.Contains("dest")    ||
            name.Contains("result") || name.Contains("report");
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  String extension
    // ─────────────────────────────────────────────────────────────────────────
    internal static class StringExtensions
    {
        public static string ToTitleCase(this string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            return char.ToUpperInvariant(s[0]) + s[1..];
        }
    }
}
