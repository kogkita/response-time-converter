// ═══════════════════════════════════════════════════════════════════════════
//  MainWindow_patch.cs  —  drop into the TestApp project as-is.
//
//  NO XAML CHANGES NEEDED.  All new UI elements are built in code and
//  injected into the existing Script Runner grid at startup.
//
//  Two small edits still required in MainWindow.xaml.cs:
//    ①  Add the call  InitScriptParamPanel();  inside the constructor,
//        after InitializeComponent().
//    ②  Delete the three methods that this file replaces:
//           SetScriptFile(string)
//           ScriptRun_Click(object, RoutedEventArgs)
//           ScriptFileClear_Click(object, RoutedEventArgs)
// ═══════════════════════════════════════════════════════════════════════════

using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace TestApp
{
    public partial class MainWindow
    {
        // ── New fields ────────────────────────────────────────────────────
        private ScriptParamPanel? _paramPanel;
        private bool              _useManualArgs = false;

        // These are created in code — no x:Name in XAML needed
        private Border?      _scriptDynParamsContainer;
        private StackPanel?  _scriptDynParamsHost;
        private Border?      _scriptManualArgsContainer;   // the existing Runtime+Args border
        private TextBlock?   _paramDetectionLabel;
        private Border?      _paramDetectionBadge;
        private Button?      _paramManualToggleBtn;


        // ── ① Call this from the constructor after InitializeComponent() ──

        private void InitScriptParamPanel()
        {
            // Walk up from ScriptLogPanel to find the outer runner Grid
            // (ScriptLogPanel is inside StackPanel > ScrollViewer > Grid)
            var p = ScriptLogPanel.Parent;
            while (p != null && p is not Grid) p = (p as System.Windows.FrameworkElement)?.Parent;
            if (p is not Grid runnerGrid) return;

            // Find the ScrollViewer and the StackPanel it wraps
            var scrollViewer = runnerGrid.Children.OfType<ScrollViewer>().FirstOrDefault();
            var stackPanel   = scrollViewer?.Content as StackPanel;
            if (stackPanel == null) return;

            // Find the Runtime+Args border by checking which border contains ScriptRuntimeBox
            _scriptManualArgsContainer = stackPanel.Children.OfType<Border>()
                .FirstOrDefault(b => IsAncestorOf(b, ScriptRuntimeBox));

            // Build the smart params container and insert after the script file border (index 1)
            _scriptDynParamsContainer = BuildDynParamsContainer();
            stackPanel.Children.Insert(1, _scriptDynParamsContainer);

            _paramPanel = new ScriptParamPanel(
                host:         _scriptDynParamsHost!,
                container:    _scriptDynParamsContainer,
                noParamsHint: new TextBlock());
        }

        // Builds the entire smart params panel in code
        private Border BuildDynParamsContainer()
        {
            // Host for the individual param rows (injected by ScriptParamPanel.Build)
            _scriptDynParamsHost = new StackPanel();

            // Detection badge
            _paramDetectionLabel = new TextBlock
            {
                Text       = "auto-detected",
                Foreground = Brush("#4ADE80"),
                FontSize   = 10,
                FontFamily = FF("Segoe UI Variable, Segoe UI")
            };
            _paramDetectionBadge = new Border
            {
                Background      = Brush("#1A2C1A"),
                BorderBrush     = Brush("#2D4A2D"),
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(4),
                Padding         = new Thickness(8, 2, 8, 2),
                Margin          = new Thickness(10, 0, 0, 0),
                Child           = _paramDetectionLabel
            };

            // "Edit manually instead" toggle link
            _paramManualToggleBtn = new Button
            {
                Content         = "Edit manually instead",
                Background      = Brushes.Transparent,
                Foreground      = Brush("#4A5F88"),
                BorderThickness = new Thickness(0),
                FontSize        = 10.5,
                Cursor          = System.Windows.Input.Cursors.Hand,
            };
            // Underline style via template
            var toggleTemplate = new ControlTemplate(typeof(Button));
            var tbFactory = new FrameworkElementFactory(typeof(TextBlock));
            tbFactory.SetBinding(TextBlock.TextProperty,
                new System.Windows.Data.Binding("Content")
                { RelativeSource = new System.Windows.Data.RelativeSource(
                    System.Windows.Data.RelativeSourceMode.TemplatedParent) });
            tbFactory.SetBinding(TextBlock.ForegroundProperty,
                new System.Windows.Data.Binding("Foreground")
                { RelativeSource = new System.Windows.Data.RelativeSource(
                    System.Windows.Data.RelativeSourceMode.TemplatedParent) });
            tbFactory.SetBinding(TextBlock.FontSizeProperty,
                new System.Windows.Data.Binding("FontSize")
                { RelativeSource = new System.Windows.Data.RelativeSource(
                    System.Windows.Data.RelativeSourceMode.TemplatedParent) });
            tbFactory.SetValue(TextBlock.TextDecorationsProperty, TextDecorations.Underline);
            tbFactory.SetValue(TextBlock.CursorProperty, System.Windows.Input.Cursors.Hand);
            toggleTemplate.VisualTree = tbFactory;
            _paramManualToggleBtn.Template = toggleTemplate;
            _paramManualToggleBtn.Click += ParamManualToggle_Click;

            // Header row: "SCRIPT PARAMETERS" label + badge + toggle link
            var headerLeft = new StackPanel { Orientation = Orientation.Horizontal };
            headerLeft.Children.Add(new TextBlock
            {
                Text            = "SCRIPT PARAMETERS",
                FontSize        = 10,
                FontWeight      = FontWeights.Bold,
                Foreground      = Brush("#6B7A99"),
                FontFamily      = FF("Segoe UI Variable, Segoe UI"),
                VerticalAlignment = VerticalAlignment.Center
            });
            headerLeft.Children.Add(_paramDetectionBadge);

            var headerGrid = new Grid { Margin = new Thickness(0, 0, 0, 10) };
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            Grid.SetColumn(headerLeft, 0);
            Grid.SetColumn(_paramManualToggleBtn, 1);
            headerGrid.Children.Add(headerLeft);
            headerGrid.Children.Add(_paramManualToggleBtn);

            // Required fields note
            var requiredNote = new TextBlock
            {
                Text       = "  * required",
                Foreground = Brush("#4A5F88"),
                FontSize   = 10.5,
                FontFamily = FF("Segoe UI Variable, Segoe UI"),
                Margin     = new Thickness(0, 8, 0, 0)
            };

            var inner = new StackPanel();
            inner.Children.Add(headerGrid);
            inner.Children.Add(_scriptDynParamsHost);
            inner.Children.Add(requiredNote);

            return new Border
            {
                Background      = Brush("#0D1020"),
                BorderBrush     = Brush("#1A2030"),
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(8),
                Padding         = new Thickness(14, 12, 14, 12),
                Margin          = new Thickness(0, 0, 0, 10),
                Visibility      = Visibility.Collapsed,
                Child           = inner
            };
        }


        // ── ② Replace SetScriptFile ───────────────────────────────────────

        private void SetScriptFile(string path)
        {
            _scriptFilePath = path;
            ScriptFileLabel.Text = path;
            ScriptFileLabel.Foreground = Brush("#CBD5E1");
            ScriptFileClearBtn.Visibility = Visibility.Visible;

            if (SaveLogCheckbox.IsChecked == true)
            {
                _saveLogPath = System.IO.Path.Combine(
                    System.IO.Path.GetDirectoryName(path) ?? "",
                    System.IO.Path.GetFileNameWithoutExtension(path) + "_run.log");
                SaveLogPathLabel.Text = _saveLogPath;
            }

            // Detect parameters and switch panel mode
            var detected = ScriptParamDetector.Detect(path);
            _useManualArgs = false;

            if (detected.Count > 0 && _scriptDynParamsContainer != null)
            {
                _paramPanel?.Build(detected);
                _scriptDynParamsContainer.Visibility  = Visibility.Visible;
                if (_scriptManualArgsContainer != null)
                    _scriptManualArgsContainer.Visibility = Visibility.Collapsed;

                bool fromBlock = IsFromMetadataBlock(path);
                if (_paramDetectionLabel != null)
                {
                    _paramDetectionLabel.Text       = fromBlock ? "declared" : "auto-detected";
                    _paramDetectionLabel.Foreground = Brush(fromBlock ? "#60A5FA" : "#4ADE80");
                }
                if (_paramDetectionBadge != null)
                {
                    _paramDetectionBadge.Background   = Brush(fromBlock ? "#1E2640" : "#1A2C1A");
                    _paramDetectionBadge.BorderBrush  = Brush(fromBlock ? "#3D4D70" : "#2D4A2D");
                }
                if (_paramManualToggleBtn != null)
                    _paramManualToggleBtn.Content = "Edit manually instead";
            }
            else
            {
                _paramPanel?.Clear();
                if (_scriptDynParamsContainer != null)
                    _scriptDynParamsContainer.Visibility = Visibility.Collapsed;
                if (_scriptManualArgsContainer != null)
                    _scriptManualArgsContainer.Visibility = Visibility.Visible;
            }

            // Type badge (unchanged)
            string ext = System.IO.Path.GetExtension(path).ToLowerInvariant();
            if (ScriptTypes.TryGetValue(ext, out var info))
            {
                ScriptTypeLabel.Text = info.Label;
                ScriptTypeBadge.Background = Brush("#1E2640");
                ScriptTypeLabel.Foreground = Brush(info.Color);
            }
            else
            {
                ScriptTypeLabel.Text = "Unknown";
                ScriptTypeLabel.Foreground = Brush("#F87171");
            }

            MarkScriptEntryDirty();
        }


        // ── ② Replace ScriptFileClear_Click ──────────────────────────────

        private void ScriptFileClear_Click(object sender, RoutedEventArgs e)
        {
            _scriptFilePath = null;
            ScriptFileLabel.Text = "No script selected — browse or drag & drop here";
            ScriptFileLabel.Foreground = Brush("#6B7FA8");
            ScriptFileClearBtn.Visibility = Visibility.Collapsed;
            ScriptTypeLabel.Text = "None";
            ScriptTypeBadge.Background = Brush("#1E2640");

            _paramPanel?.Clear();
            if (_scriptDynParamsContainer != null)
                _scriptDynParamsContainer.Visibility = Visibility.Collapsed;
            if (_scriptManualArgsContainer != null)
                _scriptManualArgsContainer.Visibility = Visibility.Visible;
            _useManualArgs = false;
        }


        // ── New: "Edit manually instead" toggle ───────────────────────────

        private void ParamManualToggle_Click(object sender, RoutedEventArgs e)
        {
            _useManualArgs = !_useManualArgs;

            if (_useManualArgs)
            {
                if (_scriptManualArgsContainer != null)
                    _scriptManualArgsContainer.Visibility = Visibility.Visible;
                if (_scriptDynParamsContainer != null)
                    _scriptDynParamsContainer.Visibility = Visibility.Collapsed;
                if (_paramManualToggleBtn != null)
                    _paramManualToggleBtn.Content = "← Use smart parameters";
                if (_paramPanel != null)
                    ScriptArgsBox.Text = _paramPanel.BuildArgumentString();
            }
            else
            {
                if (_scriptManualArgsContainer != null)
                    _scriptManualArgsContainer.Visibility = Visibility.Collapsed;
                if (_scriptDynParamsContainer != null)
                    _scriptDynParamsContainer.Visibility = Visibility.Visible;
                if (_paramManualToggleBtn != null)
                    _paramManualToggleBtn.Content = "Edit manually instead";
            }
        }


        // ── ② Replace ScriptRun_Click ─────────────────────────────────────

        private void ScriptRun_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_scriptFilePath) || !System.IO.File.Exists(_scriptFilePath))
            {
                DarkMessageBox.Show("Please select a script file first.", "No Script");
                return;
            }

            string ext = System.IO.Path.GetExtension(_scriptFilePath).ToLowerInvariant();

            string runtime, argsPrefix;
            string runtimeOverride = ScriptRuntimeBox.Text.Trim();
            if (!string.IsNullOrEmpty(runtimeOverride))
            {
                var parts = runtimeOverride.Split(' ', 2);
                runtime    = parts[0];
                argsPrefix = parts.Length > 1 ? parts[1] : "";
            }
            else if (ScriptTypes.TryGetValue(ext, out var info))
            {
                runtime    = info.Runtime;
                argsPrefix = info.ArgsPrefix;
            }
            else
            {
                DarkMessageBox.Show(
                    $"Unknown script type '{ext}'.\nEnter a runtime in the Runtime Override field.",
                    "Unknown Type");
                return;
            }

            string userArgs;
            bool smartMode = !_useManualArgs
                          && _paramPanel != null
                          && _scriptDynParamsContainer?.Visibility == Visibility.Visible;

            if (smartMode)
            {
                if (!_paramPanel!.Validate(out string missing))
                {
                    DarkMessageBox.Show(
                        $"The required field \"{missing}\" is empty.\nPlease provide a value before running.",
                        "Missing Required Field");
                    return;
                }
                userArgs = _paramPanel.BuildArgumentString();
            }
            else
            {
                userArgs = ScriptArgsBox.Text.Trim();
            }

            string scriptPath = _scriptFilePath;
            string workDir    = ScriptWorkDirBox.Text.Trim();
            if (string.IsNullOrEmpty(workDir))
                workDir = System.IO.Path.GetDirectoryName(scriptPath) ?? "";

            string fullArgs = string.IsNullOrEmpty(argsPrefix)
                ? $"\"{scriptPath}\" {userArgs}".Trim()
                : $"{argsPrefix} \"{scriptPath}\" {userArgs}".Trim();

            var envVars = new Dictionary<string, string>();
            foreach (var child in ScriptEnvVarPanel.Children)
            {
                if (child is Grid row && row.Children.Count >= 2)
                {
                    var k = (row.Children[0] as TextBox)?.Text?.Trim() ?? "";
                    var v = (row.Children[1] as TextBox)?.Text?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(k)) envVars[k] = v;
                }
            }

            ScriptLogPanel.Visibility     = Visibility.Visible;
            ScriptLog.Text                = "";
            ScriptProgress.Visibility     = Visibility.Visible;
            SaveLogAfterRunBtn.Visibility = Visibility.Collapsed;
            ScriptExitCodeLabel.Text      = "";
            ScriptRunBtn.IsEnabled        = false;
            ScriptStopBtn.Visibility      = Visibility.Visible;
            ScriptStatusLabel.Text        = $"Running {System.IO.Path.GetFileName(scriptPath)}…";
            ScriptStatusLabel.Foreground  = new SolidColorBrush(Color.FromRgb(0x60, 0xA5, 0xFA));

            AppendScriptLog($"▶ {runtime} {fullArgs}", "#60A5FA");
            WriteLogHeader();
            AppendScriptLog($"  Working dir: {workDir}\n", "#6B7A99");

            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    var psi = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName               = runtime,
                        Arguments              = fullArgs,
                        WorkingDirectory       = workDir,
                        UseShellExecute        = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError  = true,
                        CreateNoWindow         = true,
                    };
                    foreach (var kv in envVars)
                        psi.EnvironmentVariables[kv.Key] = kv.Value;
                    if (ext == ".py" && !psi.EnvironmentVariables.ContainsKey("PYTHONIOENCODING"))
                        psi.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";

                    _scriptProcess = new System.Diagnostics.Process
                    {
                        StartInfo           = psi,
                        EnableRaisingEvents = true
                    };
                    _scriptProcess.OutputDataReceived += (s, ev) =>
                    {
                        if (ev.Data != null)
                            Dispatcher.Invoke(() => AppendScriptLog(ev.Data, "#A8B3C8"));
                    };
                    _scriptProcess.ErrorDataReceived += (s, ev) =>
                    {
                        if (ev.Data != null)
                            Dispatcher.Invoke(() => AppendScriptLog(ev.Data, "#F87171"));
                    };
                    _scriptProcess.Start();
                    _scriptProcess.BeginOutputReadLine();
                    _scriptProcess.BeginErrorReadLine();
                    _scriptProcess.WaitForExit();

                    int code = _scriptProcess.ExitCode;
                    _scriptProcess = null;

                    Dispatcher.Invoke(() =>
                    {
                        ScriptProgress.Visibility = Visibility.Collapsed;
                        ScriptRunBtn.IsEnabled    = true;
                        ScriptStopBtn.Visibility  = Visibility.Collapsed;
                        bool ok = code == 0;
                        ScriptExitCodeLabel.Text = $"Exit code: {code}";
                        ScriptExitCodeLabel.Foreground = new SolidColorBrush(ok
                            ? Color.FromRgb(0x4A, 0xDE, 0x80) : Color.FromRgb(0xF8, 0x71, 0x71));
                        ScriptStatusLabel.Text = ok
                            ? "Completed successfully." : $"Finished with exit code {code}.";
                        ScriptStatusLabel.Foreground = new SolidColorBrush(ok
                            ? Color.FromRgb(0x4A, 0xDE, 0x80) : Color.FromRgb(0xF8, 0x71, 0x71));
                        AppendScriptLog($"\n■ Process exited with code {code}",
                            ok ? "#4ADE80" : "#F87171");
                        SaveLogAfterRunBtn.Visibility = SaveLogCheckbox.IsChecked == true
                            ? Visibility.Collapsed : Visibility.Visible;
                    });
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        ScriptProgress.Visibility = Visibility.Collapsed;
                        ScriptRunBtn.IsEnabled    = true;
                        ScriptStopBtn.Visibility  = Visibility.Collapsed;
                        ScriptStatusLabel.Text    = $"Failed to start: {ex.Message}";
                        ScriptStatusLabel.Foreground = new SolidColorBrush(
                            Color.FromRgb(0xF8, 0x71, 0x71));
                        AppendScriptLog($"\n✗ Failed to start process: {ex.Message}", "#F87171");
                        _scriptProcess = null;
                    });
                }
            });
        }


        // ── Helpers ───────────────────────────────────────────────────────

        private static bool IsFromMetadataBlock(string path)
        {
            try
            {
                foreach (var line in System.IO.File.ReadLines(path))
                {
                    var t = line.Trim();
                    if (t == "## SCRIPT_RUNNER" || t == "# SCRIPT_RUNNER") return true;
                    if (!t.StartsWith('#') && !string.IsNullOrWhiteSpace(t)) break;
                }
            }
            catch { }
            return false;
        }

        private static SolidColorBrush Brush(string hex)
            => new((Color)ColorConverter.ConvertFromString(hex));

        private static FontFamily FF(string name)
            => new(name);
        // Helper: check if 'ancestor' contains 'element' anywhere in its visual subtree
        private static bool IsAncestorOf(System.Windows.DependencyObject ancestor, System.Windows.DependencyObject element)
        {
            var current = element;
            while (current != null)
            {
                if (current == ancestor) return true;
                current = System.Windows.Media.VisualTreeHelper.GetParent(current);
            }
            return false;
        }

    }
}
