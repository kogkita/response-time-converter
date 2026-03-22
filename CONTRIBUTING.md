# Contributing to Performance Test Utilities

Thank you for your interest in contributing! This document covers how to report bugs, suggest enhancements, and submit code changes.

---

## Getting Started

1. **Fork the repository** and clone your fork locally.
2. Open `TestApp.slnx` in Visual Studio 2022 (v18+).
3. Restore NuGet packages — the project uses EPPlus and Microsoft.Extensions.
4. Build and run (`F5`) to verify everything works before making changes.

---

## Reporting Bugs

Before filing a bug, check existing issues to avoid duplicates. When creating a bug report please include:

- A clear, descriptive title
- Steps to reproduce the problem
- What you expected to happen vs what actually happened
- Screenshots if relevant
- Your Windows version and .NET runtime version

---

## Suggesting Enhancements

Enhancement suggestions are welcome as GitHub issues. Please include:

- A clear title and description of the proposed change
- The use case / problem it solves
- Any examples or mockups if applicable

---

## Pull Requests

1. Create a branch from `main`:
   ```bash
   git checkout -b feature/your-feature-name
   ```
2. Make your changes — see the coding standards below.
3. Commit with a clear message:
   ```bash
   git commit -m "Add your message here"
   ```
4. Push and open a pull request against `main`.
5. Describe what the PR does and why.

---

## Project Structure

| File | Responsibility |
|---|---|
| `MainWindow.xaml` | All UI layout and styles |
| `MainWindow.xaml.cs` | UI event handlers and page logic |
| `MainWindow_patch.cs` | Supplementary UI patches |
| **Shared Utilities** | |
| `CsvHelper.cs` | Shared quote-aware CSV line splitter |
| `ExcelNameHelper.cs` | Shared unique sheet/table name helpers |
| `AppDataManager.cs` | JSON persistence — libraries, settings, global backup |
| `AppLogger.cs` | Thread-safe rolling file logger |
| **Convert Response Times** | |
| `ResponseTimeConverter.cs` | JMeter CSV → Excel conversion |
| `ResponseTimeConverterExcelCharts.cs` | Mini bar-chart sheet generation |
| **JTL File Processing** | |
| `JTLFileProcessing.cs` | JTL → Excel conversion |
| `JTLFileProcessingExcelCharts.cs` | Mini bar-chart sheet generation |
| **BLG File Conversion** | |
| `BLGConverter.cs` | relog.exe wrapper, counter filter templates, BLG → CSV |
| `BLGGraphProducer.cs` | Server performance line charts from relog CSV output |
| **nmon Analysis** | |
| `NmonParser.cs` | .nmon file parser |
| `NmonExcelProducer.cs` | Built-in nmon → Excel (no Excel installation needed) |
| `NmonAnalyzer.cs` | nmon_analyser Excel macro automation via VBScript |
| **Run Comparison** | |
| `RunComparisonProcessor.cs` | Multi-run delta computation and Excel reporting |
| **Script Runner** | |
| `ScriptParamDetector.cs` | argparse / PowerShell param auto-detection |
| `ScriptParamPanel.cs` | Dynamic parameter UI builder |
| `WindowsTaskScheduler.cs` | schtasks.exe wrapper for scheduling |
| **Test Run Trends** | |
| `TestRunTrendsProcessor.cs` | Trend report generation engine |
| `TrendsManifest.cs` | File change detection for auto-watch |
| `BulkImportOptionsDialog.cs` | Bulk import options dialog |
| **System Tray** | |
| `TrayManager.cs` | System tray icon (reflection-based, no WinForms reference) |

---

## Coding Standards

- **Language:** C# — follow standard C# conventions throughout
- **Naming:** PascalCase for classes and methods; camelCase for local variables and fields
- **Async:** Long-running operations (file I/O, relog, Excel automation) must run on a background thread via `Task.Run` to keep the UI responsive; use `Dispatcher.Invoke` to update UI from background threads
- **Error handling:** Wrap all processing in try/catch and surface errors via `MessageBox` — never silently swallow exceptions
- **New tools:** Add a new page in `MainWindow.xaml` following the existing pattern (heading, browse row, file list, drop zone, run button); add the corresponding nav button and `SetActivePage` call in `MainWindow.xaml.cs`; implement processing logic in a dedicated `.cs` file
- **Counter files for BLG:** If adding or modifying counter templates in `BLGConverter.cs`, ensure counter strings have no trailing whitespace and no BOM — `relog.exe` is sensitive to both

---

## Commit Message Style

- Use the present tense: `Add feature` not `Added feature`
- Use the imperative mood: `Fix crash on empty file list` not `Fixes crash`
- Keep the first line under 72 characters
- Reference issues where relevant: `Fix #42 — BLG conversion missing counter`

---

Thank you for contributing!
