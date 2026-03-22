# Redundancy Scan & Comment Cleanup — Change Summary

**Project:** PerformanceAnalysisUtilities  
**Files scanned:** 22 C# source files (12,663 lines)  
**Files modified:** 10 | **Files added:** 2 | **Net lines removed:** 86  

---

## New Files Created

### `CsvHelper.cs`
Shared quote-aware CSV line splitter. Replaces **5 duplicate implementations** that existed independently in JTLFileProcessing, ResponseTimeConverter, RunComparisonProcessor, NmonParser, and BLGGraphProducer. Provides both `string[]` and `List<string>` return variants.

### `ExcelNameHelper.cs`
Shared `UniqueSheetName` and `UniqueTableName` helpers. Replaces **2 identical implementations** in JTLFileProcessing and ResponseTimeConverter. RunComparisonProcessor was already calling JTLFileProcessing's copy externally.

---

## Files Modified

### `JTLFileProcessing.cs`
- **Removed** duplicate `<summary>` block on `ClearPendingCharts()` (stale first copy kept describing the old name/purpose).
- **Replaced** `UniqueSheetName()` and `UniqueTableName()` bodies with one-line delegates to `ExcelNameHelper`.
- **Replaced** `SplitCsvLine()` body with one-line delegate to `CsvHelper.SplitCsvLineToList()`.

### `ResponseTimeConverter.cs`
- **Removed** duplicate `<summary>` block on `ClearPendingCharts()` (same issue as JTL).
- **Replaced** `UniqueSheetName()` and `UniqueTableName()` bodies with delegates to `ExcelNameHelper`.
- **Replaced** `SplitCsvLine()` body with delegate to `CsvHelper.SplitCsvLine()`.

### `RunComparisonProcessor.cs`
- **Removed** duplicate section comment `// ── Entry point ──` (appeared twice consecutively with blank lines).
- **Removed** duplicate section comment `// ── Summary sheet ──` (appeared twice consecutively).
- **Replaced** `SplitCsvLine()` body with delegate to `CsvHelper.SplitCsvLine()`.

### `NmonParser.cs`
- **Replaced** `SplitCsvLine()` body (simpler variant without escaped-quote handling) with delegate to `CsvHelper.SplitCsvLine()`. The shared version handles escaped quotes, making NmonParser more robust than before.

### `BLGGraphProducer.cs`
- **Replaced** `SplitCsvLine()` body (simpler variant) with delegate to `CsvHelper.SplitCsvLine()`. Same robustness improvement as NmonParser.

### `BLGConverter.cs`
- **Removed** redundant `bool usingTemp = true;` variable. It was always true and the accompanying comment already explained why. The conditional `if (usingTemp && File.Exists(...))` was simplified to `if (File.Exists(...))`.

### `JTLFileProcessingExcelCharts.cs`
- **Added** class-level `<summary>` documentation.
- **Added** `// TODO` comment flagging the ~400-line duplication with `ResponseTimeConverterExcelCharts` (BuildScaleChartXml, BuildDrawingXml, EscapeXml, InjectChartsForSheet are near-identical).
- **Removed** unused `ExtractNum()` method (dead code — never called anywhere in the codebase).

### `ResponseTimeConverterExcelCharts.cs`
- **Added** class-level `<summary>` documentation.
- **Added** matching `// TODO` comment flagging the duplication with JTL charts.
- **Removed** unused `ExtractNum()` method (dead code — same as JTL charts copy).

### `ScriptParamPanel.cs`
- **Removed** empty no-op method `SyncValueFromControl()` and its call site in `BuildArgumentString()`. The method body contained only a comment explaining that event handlers already sync values directly — the method served no purpose.

### `TrendsManifest.cs`
- **Collapsed** two redundant early-return checks into one. The original had:
  ```csharp
  if (!contentChanged && renames.Count == 0) return (false, "");
  if (!contentChanged && renames.Count > 0)  return (false, "");
  ```
  Both branches returned the same result. Simplified to:
  ```csharp
  if (!contentChanged) return (false, "");
  ```

---

## Identified but Not Refactored (Future Work)

### Chart XML Builder Duplication (~600 lines)
`JTLFileProcessingExcelCharts.cs` and `ResponseTimeConverterExcelCharts.cs` contain near-identical implementations of:

| Method | Status |
|--------|--------|
| `BuildScaleChartXml()` | Verbatim identical |
| `BuildDrawingXml()` | Verbatim identical |
| `EscapeXml()` | Verbatim identical |
| `InjectChartsForSheet()` | Near-identical (~100 lines each) |
| `BuildMiniChartXml()` | Structurally identical, differs in value source |

**Recommendation:** Extract a shared `ChartXmlBuilder` static class. The mini-chart builder could accept a simple record interface or delegate for value extraction. This would eliminate approximately 400 redundant lines. Marked with `// TODO` comments in both files.
