# Performance Test Utilities

A Windows desktop application for processing and analysing performance test data. Built with WPF (.NET 10) and EPPlus, it provides a dark-themed GUI with four tools covering the most common post-test analysis workflows.

---

## Tools

### 1. Convert Response Times
Imports one or more JMeter CSV (`.csv`) report files and exports formatted Excel workbooks with summary statistics.

- Multiple files supported — each gets its own `.xlsx` output
- Optional **Include Charts** — adds a latency chart sheet to the output
- Optional **Club into one file** — merges all inputs into a single Excel workbook

### 2. JTL File Processing
Imports one or more JMeter JTL (`.jtl`) result files and exports formatted Excel workbooks with summary stats and charts.

- Supports `.jtl` JMeter result files (CSV-formatted JTL output)
- Optional **Include Charts** and **Club into one file** options, same as above

### 3. BLG File Conversion
Converts Windows Performance Monitor binary log (`.blg`) files to CSV using `relog.exe`, applying a performance counter filter template.

- **App Server template** — captures OS and process counters:
  - `\Memory\Available MBytes`
  - `\Memory\Pages Input/sec` / `Pages Output/sec`
  - `\Network Interface(*)\Bytes Total/sec`
  - `\PhysicalDisk(_Total)\% Idle Time`, `Avg. Disk sec/Transfer`, `Current Disk Queue Length`
  - `\Paging File(_Total)\% Usage`
  - `\Process(*)\% Processor Time`, `Working Set`
  - `\Processor(_Total)\% Processor Time`
  - `\PhysicalDisk(*)\*`

- **DB Server template** — all App Server counters plus SQL Server-specific counters:
  - `\SQLServer:Buffer Manager\Buffer cache hit ratio`
  - `\SQLServer:Buffer Manager\Page life expectancy`
  - `\SQLServer:General Statistics\User Connections`
  - `\SQLServer:Latches\Average Latch Wait Time (ms)`
  - `\SQLServer:Locks(_Total)\Average Wait Time (ms)`
  - `\SQLServer:Locks(_Total)\Number of Deadlocks/sec`

- **Custom counter file** — optionally supply your own `.txt` counter list to override the built-in template. One counter path per line; trailing whitespace and carriage returns are sanitised automatically.
- **Command preview** — shows the exact `relog.exe` command that will be executed before you run it.
- Requires `relog.exe` (ships with Windows, normally at `C:\Windows\System32\relog.exe`).

### 4. nmon Analyzer
Analyses NMON files (`.nmon`) using the `nmon_analyser_v69_2.xlsm` Excel macro engine.

- Configurable options: GRAPHS, MERGE, INTERVALS, ESS, SCATTER, BIGDATA, SHOWLINUXCPUUTIL, REORDER, SORTDEFAULT, LIST
- Custom output directory support
- Runs Excel macro automation in the background; UI remains responsive

---

## Requirements

- Windows 10 / 11
- .NET 10 Runtime (Windows)
- Microsoft Excel — required for nmon Analyzer only
- `relog.exe` — present by default on all Windows installations; required for BLG conversion only
- `nmon_analyser_v69_2.xlsm` — place alongside the executable or browse to locate it; required for nmon Analyzer only

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

### BLG File Conversion
1. Navigate to **BLG File Conversion** in the sidebar.
2. Select **App Server** or **DB Server** to use the built-in counter template, or browse for a **Custom Counter File** to override.
3. Browse for `.blg` files or drag and drop them into the drop zone.
4. Review the **Command Preview** to confirm the relog command.
5. Click **Run Conversion** — CSV files are saved alongside each input `.blg` file.

### Convert Response Times / JTL File Processing
1. Navigate to the relevant tool in the sidebar.
2. Browse or drag and drop your CSV / JTL files.
3. Toggle **Include Charts** and/or **Club into one file** as needed.
4. Click **Run Processing** — Excel files are saved alongside each input file (or to the chosen combined path).

### nmon Analyzer
1. Navigate to **nmon Analyzer** in the sidebar.
2. Browse or drag and drop `.nmon` files.
3. Configure analyser options (GRAPHS, MERGE, etc.).
4. Set the path to `nmon_analyser_v69_2.xlsm`.
5. Click **Run nmon Analyzer**.

---

## Project Structure

```
PerformanceAnalysisUtilities/
├── TestApp/
│   ├── MainWindow.xaml(.cs)               # Main UI and all page logic
│   ├── BLGConverter.cs                    # relog.exe wrapper, counter templates
│   ├── ResponseTimeConverter.cs           # JMeter CSV → Excel
│   ├── ResponseTimeConverterExcelCharts.cs
│   ├── JTLFileProcessing.cs               # JTL → Excel
│   ├── JTLFileProcessingExcelCharts.cs
│   └── NmonAnalyzer.cs                    # nmon_analyser Excel automation
├── StandAloneExecutable/                  # Pre-built binaries + dependencies
├── README.md
└── CONTRIBUTING.md
```

---

## License

See [LICENSE](LICENSE) for details.
