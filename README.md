# Performance Test Utilities

A Windows desktop application for processing and analysing performance test data. Built with WPF (.NET 10) and EPPlus, it provides a dark-themed GUI with eight tools covering the most common post-test analysis workflows — from raw file conversion to multi-run trend tracking with automated watch.

---

## Tools

### 1. Convert Response Times
Imports one or more JMeter CSV (`.csv`) report files and exports formatted Excel workbooks with summary statistics.

- Multiple files supported — each gets its own `.xlsx` output
- Optional **Include Charts** — adds a latency chart sheet (Avg vs P90 mini bars per transaction, aligned on a shared 0–60 s scale)
- Optional **Club into one file** — merges all inputs into a single Excel workbook with per-file sheets

### 2. JTL File Processing
Imports one or more JMeter JTL (`.jtl`) result files and exports formatted Excel workbooks with summary stats and charts.

- Supports `.jtl` JMeter result files (CSV-formatted JTL output)
- Detects transaction-level samplers automatically (filters out individual HTTP requests)
- Optional **Include Charts** and **Club into one file** options, same as above

### 3. BLG File Conversion
Converts Windows Performance Monitor binary log (`.blg`) files to CSV using `relog.exe`, applying a performance counter filter template.

- **App Server template** — OS and process counters:
  - `\Memory\Available MBytes`, `Pages Input/sec`, `Pages Output/sec`
  - `\Network Interface(*)\Bytes Total/sec`
  - `\PhysicalDisk(_Total)\% Idle Time`, `Avg. Disk sec/Transfer`, `Current Disk Queue Length`
  - `\Paging File(_Total)\% Usage`
  - `\Process(*)\% Processor Time`, `Working Set`
  - `\Processor(_Total)\% Processor Time`
  - `\PhysicalDisk(*)\*`

- **DB Server template** — all App Server counters plus SQL Server counters:
  - `\SQLServer:Buffer Manager\Buffer cache hit ratio`, `Page life expectancy`
  - `\SQLServer:General Statistics\User Connections`
  - `\SQLServer:Latches\Average Latch Wait Time (ms)`
  - `\SQLServer:Locks(_Total)\Average Wait Time (ms)`, `Number of Deadlocks/sec`

- **Custom counter file** — supply your own `.txt` counter list to override the built-in template. One counter path per line; trailing whitespace and carriage returns are sanitised automatically.
- **Produce Graphs** — optionally generate an Excel workbook with line charts for Available Memory, CPU, Pages Input/sec, Disk Queue, Network (Mbps), and Disk Busy % across one or more servers.
- **Command preview** — shows the exact `relog.exe` command before execution.
- Requires `relog.exe` (ships with Windows, normally at `C:\Windows\System32\relog.exe`).

### 4. nmon Analysis
Two processing paths for NMON files (`.nmon`):

- **nmon Analyzer (macro-based)** — Automates `nmon_analyser_v69_2.xlsm` via VBScript/cscript.exe. Configurable options: GRAPHS scope (ALL/LIST), output format (CHARTS/PICTURES/PRINT/WEB), MERGE, INTERVALS, ESS, SCATTER, BIGDATA, SHOWLINUXCPUUTIL, REORDER, SORTDEFAULT, LIST, and custom output directory. Requires Microsoft Excel.
- **nmon Excel Producer (built-in)** — Parses `.nmon` files directly (no Excel dependency) and produces a standalone Excel workbook with per-section data sheets (CPU_ALL, MEM, DISKBUSY, DISKREAD, DISKWRITE, DISKXFER, NET, NETPACKET, PAGE, PROC, LPAR, VM, JFSFILE, TOP) plus line charts. Supports multiple files with a summary sheet.

### 5. Run Comparison
Compares two or more performance test runs (`.csv` or `.jtl`) and produces a detailed Excel report highlighting regressions, improvements, and SLA breaches.

- Accepts **2 to N run files** — auto-detects CSV vs JTL per file from extension
- **All vs Baseline** mode — compares every run against Run 1
- **Sequential** mode — compares each run against the previous one (Run 2 vs 1, Run 3 vs 2, …)
- Optional **SLA threshold** — flags transactions whose P90 exceeds the threshold
- Output includes:
  - **Summary** sheet with KPI cards (regression count, improvements, SLA breaches)
  - **Avg Comparison** and **P90 Comparison** sheets per pair with colour-coded status (Regression / Improvement / Stable / New / Removed)
  - **Raw data** sheets for each unique run file

### 6. Script Runner
A built-in script library and execution environment for Python, PowerShell, batch, and other scripts used in performance testing workflows.

- **Script Library** — save, organise, and recall scripts with metadata (name, description, path, runtime, arguments, working directory, environment variables)
- **Parameter auto-detection** — scans Python `argparse` and PowerShell `param()` blocks to build a dynamic parameter form (file pickers, text fields, checkboxes, numeric inputs)
- **Metadata block support** — scripts can declare parameters explicitly via a `## SCRIPT_RUNNER` / `# SCRIPT_RUNNER` comment block for full control over UI rendering
- **Scheduling** — register scripts as Windows Task Scheduler tasks (Once, Daily, Weekly) with configurable time-of-day; generates a launcher `.bat` that handles working directory, environment variables, and log redirection
- **Export / Import** — share script libraries between machines as JSON files

### 7. Test Run Trends
Tracks test execution results across multiple runs for one or more customers, producing trend reports with failure streak detection.

- **Customer library** — manage multiple customers, each with a Runs folder and Reports folder
- **Trend generation** — reads all `.xlsx` run files from a customer's runs folder and produces a multi-sheet Excel report:
  - **Summary** — per-run pass/fail/pass% with total and average runtime
  - **Test Case Trends** — matrix of all test cases × all runs with PASS/FAIL status, runtime, and colour-coded cells. Includes a fail-count column for the configurable fail window and a consecutive fail streak column
  - **By Test Plan** — same data grouped by Test Plan Name with collapsible row groups; highlights test cases that moved between plans
  - **Flags** — auto-detected issues: STREAK (consecutive failures), FAIL, MISSING, NEW, SLOWER (>25%), FASTER (>25%)
  - **Charts** — Pass % trend, Total Runtime trend, and Failed count bar chart
- **Configurable fail window** — global or per-customer setting for how many recent runs to count failures in
- **Auto-watch** — monitors each customer's runs folder for new, modified, or removed `.xlsx` files (using size + timestamp fingerprinting via a `_runs_manifest.json`). Regenerates trends automatically when changes are detected
- **Per-customer watch interval** — override the global polling interval for individual customers
- **Bulk Import** — discover new customer folders from a root directory, with options for report output placement (same as runs, one shared folder, or per-customer subfolder)
- **System tray** — minimises to tray when auto-watch is active; shows balloon notifications on trend updates; double-click to restore
- **Export / Import** — share customer libraries between machines as JSON; global backup/restore of all data (scripts + customers + settings)

### 8. Settings
Persists UI preferences across sessions — all toggle states, folder paths, and feature options are saved automatically and restored on launch.

- Per-tool defaults: chart inclusion, club output, server type, comparison mode, SLA threshold, fail window
- **Application logging** — optional file logging with rolling daily files (5 MB segments), 30-day retention, and background-threaded writes that never block the UI
- **Global backup / restore** — export or import all application data (script library + trends library + settings) as a single JSON file

---

## Requirements

- Windows 10 / 11
- .NET 10 Runtime (Windows)
- Microsoft Excel — required for nmon Analyzer (macro-based) only; all other tools are standalone
- `relog.exe` — present by default on all Windows installations; required for BLG conversion only
- `nmon_analyser_v69_2.xlsm` — required for nmon Analyzer (macro-based) only; place alongside the executable or browse to locate it

---

## Setup

### Running from the StandAloneExecutable folder
1. Navigate to the `StandAloneExecutable` folder.
2. Run `TestApp.exe`.
3. No installation required — all dependencies are bundled.

### Building from source
1. Open `TestApp.slnx` in Visual Studio 2022 (v18+).
2. Restore NuGet packages (EPPlus and Microsoft.Extensions).
3. Build and run (`F5`).

---

## Usage

### Convert Response Times / JTL File Processing
1. Navigate to the relevant tool in the sidebar.
2. Browse or drag and drop your CSV / JTL files.
3. Toggle **Include Charts** and/or **Club into one file** as needed.
4. Click **Run Processing** — Excel files are saved alongside each input file (or to the chosen combined path).

### BLG File Conversion
1. Navigate to **BLG File Conversion** in the sidebar.
2. Select **App Server** or **DB Server** to use the built-in counter template, or browse for a **Custom Counter File** to override.
3. Optionally enable **Produce Graphs** to generate server performance charts.
4. Browse for `.blg` files or drag and drop them into the drop zone.
5. Review the **Command Preview** to confirm the relog command.
6. Click **Run Conversion** — CSV files are saved alongside each input `.blg` file; graph workbook (if enabled) is saved to the same directory.

### nmon Analysis
1. Navigate to **nmon** in the sidebar.
2. Browse or drag and drop `.nmon` files.
3. **For macro-based analysis:** configure analyser options (GRAPHS, MERGE, etc.), set the path to `nmon_analyser_v69_2.xlsm`, and click **Run nmon Analyzer**.
4. **For built-in analysis:** click **Produce Excel** to generate a standalone workbook with data sheets and charts — no Excel installation required.

### Run Comparison
1. Navigate to **Run Comparison** in the sidebar.
2. Add two or more run files (CSV or JTL) — the first file is the baseline.
3. Select comparison mode: **All vs Baseline** or **Sequential**.
4. Optionally set an **SLA threshold** (in ms) to flag P90 breaches.
5. Click **Run Comparison** — the report is saved to the chosen output path.

### Script Runner
1. Navigate to **Script Runner** in the sidebar.
2. Click **Add Script** to register a new script (Python, PowerShell, batch, etc.).
3. Fill in the script path, runtime, arguments, and working directory.
4. If the script uses argparse or PowerShell params, the parameter panel auto-populates.
5. Click **Run** to execute, or **Schedule** to register a Windows Task Scheduler task.

### Test Run Trends
1. Navigate to **Test Run Trends** in the sidebar.
2. Click **Add Customer** to register a new customer with a Runs folder and Reports folder.
3. Drop `.xlsx` run files into the customer's Runs folder.
4. Click **Generate** to produce the trends report, or enable **Auto-Watch** to monitor for changes and regenerate automatically.
5. Use **Bulk Import** to discover multiple customer folders from a root directory at once.

---

## Project Structure

```
PerformanceAnalysisUtilities/
├── TestApp/
│   ├── MainWindow.xaml(.cs)                  # Main UI layout and all page logic
│   ├── MainWindow_patch.cs                   # Supplementary UI patches
│   │
│   ├── # ── Shared Utilities ──
│   ├── CsvHelper.cs                          # Shared quote-aware CSV line splitter
│   ├── ExcelNameHelper.cs                    # Shared unique sheet/table name helpers
│   ├── AppDataManager.cs                     # JSON persistence (libraries, settings, backup)
│   ├── AppLogger.cs                          # Thread-safe rolling file logger
│   │
│   ├── # ── Convert Response Times ──
│   ├── ResponseTimeConverter.cs              # JMeter CSV → Excel conversion
│   ├── ResponseTimeConverterExcelCharts.cs   # Mini bar-chart sheet generation
│   │
│   ├── # ── JTL File Processing ──
│   ├── JTLFileProcessing.cs                  # JTL → Excel conversion
│   ├── JTLFileProcessingExcelCharts.cs       # Mini bar-chart sheet generation
│   │
│   ├── # ── BLG File Conversion ──
│   ├── BLGConverter.cs                       # relog.exe wrapper, counter templates
│   ├── BLGGraphProducer.cs                   # Server performance line charts from CSV
│   │
│   ├── # ── nmon Analysis ──
│   ├── NmonParser.cs                         # .nmon file parser
│   ├── NmonExcelProducer.cs                  # Built-in nmon → Excel (no macro needed)
│   ├── NmonAnalyzer.cs                       # nmon_analyser Excel macro automation
│   │
│   ├── # ── Run Comparison ──
│   ├── RunComparisonProcessor.cs             # Multi-run delta computation and reporting
│   │
│   ├── # ── Script Runner ──
│   ├── ScriptParamDetector.cs                # argparse / PowerShell param auto-detection
│   ├── ScriptParamPanel.cs                   # Dynamic parameter UI builder
│   ├── WindowsTaskScheduler.cs               # schtasks.exe wrapper for scheduling
│   │
│   ├── # ── Test Run Trends ──
│   ├── TestRunTrendsProcessor.cs             # Trend report generation engine
│   ├── TrendsManifest.cs                     # File change detection for auto-watch
│   ├── BulkImportOptionsDialog.cs            # Bulk import options dialog
│   │
│   ├── # ── System Tray ──
│   ├── TrayManager.cs                        # System tray icon (reflection-based)
│   │
│   └── App.xaml(.cs), AssemblyInfo.cs        # Application entry point
│
├── StandAloneExecutable/                     # Pre-built binaries + dependencies
├── README.md
├── CONTRIBUTING.md
├── CHANGES.md                                # Refactoring change log
└── LICENSE
```

---

## License

See [LICENSE](LICENSE) for details.
