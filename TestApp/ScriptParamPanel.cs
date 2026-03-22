using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;

namespace TestApp
{
    // ─────────────────────────────────────────────────────────────────────────
    //  ScriptParamPanel
    //
    //  Builds and manages the dynamic "Script Parameters" panel in the
    //  Script Runner page.  Call Build() when a script is loaded; call
    //  BuildArgumentString() at run-time to get the full CLI args.
    // ─────────────────────────────────────────────────────────────────────────
    public class ScriptParamPanel
    {
        private readonly StackPanel _host;               // ScriptDynParamsPanel in XAML
        private readonly Border     _container;          // ScriptDynParamsContainer in XAML
        private readonly TextBlock  _noParamsHint;       // ScriptNoParamsHint in XAML
        private List<ScriptParam>   _params = new();

        // Map param → its value control (TextBox or CheckBox)
        private readonly Dictionary<ScriptParam, FrameworkElement> _controls = new();

        public ScriptParamPanel(StackPanel host, Border container, TextBlock noParamsHint)
        {
            _host        = host;
            _container   = container;
            _noParamsHint = noParamsHint;
        }

        // ── Public API ────────────────────────────────────────────────────────

        /// <summary>Rebuild the panel for a newly loaded script.</summary>
        public void Build(List<ScriptParam> detected)
        {
            _host.Children.Clear();
            _controls.Clear();
            _params = detected;

            bool any = detected.Count > 0;
            _container.Visibility   = any ? Visibility.Visible : Visibility.Collapsed;
            _noParamsHint.Visibility = any ? Visibility.Collapsed : Visibility.Visible;

            foreach (var param in detected)
                _host.Children.Add(BuildRow(param));
        }

        /// <summary>Clear the panel when no script is loaded.</summary>
        public void Clear()
        {
            _host.Children.Clear();
            _controls.Clear();
            _params = new();
            _container.Visibility    = Visibility.Collapsed;
            _noParamsHint.Visibility = Visibility.Collapsed;
        }

        /// <summary>
        /// Returns true if all required params have values.
        /// <paramref name="missing"/> is set to the label of the first missing one.
        /// </summary>
        public bool Validate(out string missing)
        {
            foreach (var param in _params)
            {
                if (param.Optional) continue;
                if (string.IsNullOrWhiteSpace(param.Value))
                {
                    missing = param.Label;
                    return false;
                }
            }
            missing = "";
            return true;
        }

        /// <summary>
        /// Builds the CLI argument string from the current param values.
        /// Positional args come first; named args follow.
        /// </summary>
        public string BuildArgumentString()
        {
            var positional = new System.Text.StringBuilder();
            var named      = new System.Text.StringBuilder();

            foreach (var p in _params)
            {
                if (string.IsNullOrEmpty(p.Value)) continue;

                bool isFlag      = p.Type == "flag";
                bool isPositional = !p.ArgName.StartsWith('-') && !p.ArgName.StartsWith('/');

                if (isFlag)
                {
                    if (p.Value == "true")
                        named.Append($" {p.ArgName}");
                }
                else if (isPositional)
                {
                    positional.Append($" \"{p.Value}\"");
                }
                else
                {
                    named.Append($" {p.ArgName} \"{p.Value}\"");
                }
            }

            return (positional.ToString() + named.ToString()).Trim();
        }

        // ── Row builder ───────────────────────────────────────────────────────

        private FrameworkElement BuildRow(ScriptParam param)
        {
            // Outer container
            var row = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(180) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // Label
            var labelPanel = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            var labelText  = new TextBlock
            {
                Text       = param.Label + (param.Optional ? "" : "  *"),
                Foreground = new SolidColorBrush(HexColor(param.Optional ? "#A8B3C8" : "#E2E8F0")),
                FontSize   = 12,
                FontFamily = new System.Windows.Media.FontFamily("Segoe UI Variable, Segoe UI"),
                TextWrapping = TextWrapping.Wrap,
            };
            var typeHint = new TextBlock
            {
                Text       = TypeHintText(param),
                Foreground = new SolidColorBrush(HexColor(TypeHintColor(param))),
                FontSize   = 10,
                FontFamily = new System.Windows.Media.FontFamily("Segoe UI Variable, Segoe UI"),
            };
            labelPanel.Children.Add(labelText);
            labelPanel.Children.Add(typeHint);
            Grid.SetColumn(labelPanel, 0);
            row.Children.Add(labelPanel);

            // Value control
            var control = BuildValueControl(param);
            Grid.SetColumn(control, 1);
            row.Children.Add(control);

            return row;
        }

        private FrameworkElement BuildValueControl(ScriptParam param)
        {
            switch (param.Type)
            {
                case "flag":
                    return BuildCheckBox(param);

                case "file-in":
                case "file-out":
                    return BuildFilePicker(param);

                default: // string, float, int
                    return BuildTextBox(param);
            }
        }

        // ── Individual controls ───────────────────────────────────────────────

        private FrameworkElement BuildFilePicker(ScriptParam param)
        {
            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            // Path display — TextBlock inside a clipping Border (TextBox.TextTrimming doesn't exist)
            var pathLabel = new TextBlock
            {
                Text         = param.Value,
                Foreground   = new SolidColorBrush(HexColor(string.IsNullOrEmpty(param.Value) ? "#6B7FA8" : "#CBD5E1")),
                FontSize     = 11.5,
                FontFamily   = new System.Windows.Media.FontFamily("Consolas, Segoe UI Mono, Segoe UI"),
                TextTrimming = TextTrimming.CharacterEllipsis,
                VerticalAlignment = VerticalAlignment.Center,
                Padding      = new Thickness(10, 0, 10, 0),
                ToolTip      = string.IsNullOrEmpty(param.Value) ? "No file selected" : param.Value
            };
            var pathBox = new Border
            {
                Background      = new SolidColorBrush(HexColor("#161B2A")),
                BorderBrush     = new SolidColorBrush(HexColor("#252D42")),
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(4),
                Height          = 32,
                Child           = pathLabel,
                ClipToBounds    = true,
            };

            // Declare clearBtn first so browseBtn lambda can reference it
            var clearBtn = new Button
            {
                Content         = "✕",
                Background      = new SolidColorBrush(HexColor("#3D1F1F")),
                Foreground      = new SolidColorBrush(HexColor("#F87171")),
                FontSize        = 12,
                Width           = 32,
                Height          = 32,
                Margin          = new Thickness(4, 0, 0, 0),
                BorderThickness = new Thickness(0),
                Cursor          = System.Windows.Input.Cursors.Hand,
                Visibility      = string.IsNullOrEmpty(param.Value) ? Visibility.Collapsed : Visibility.Visible,
            };

            // Browse button
            var browseBtn = new Button
            {
                Content         = "Browse…",
                Background      = new SolidColorBrush(HexColor("#374151")),
                Foreground      = new SolidColorBrush(HexColor("#B0BAC8")),
                FontSize        = 11.5,
                FontWeight      = FontWeights.SemiBold,
                Width           = 80,
                Height          = 32,
                Margin          = new Thickness(6, 0, 0, 0),
                BorderThickness = new Thickness(0),
                Cursor          = System.Windows.Input.Cursors.Hand,
            };
            ApplyActionButtonTemplate(browseBtn);
            browseBtn.Click += (_, __) =>
            {
                string? picked = param.Type == "file-out"
                    ? ShowSaveDialog(param)
                    : ShowOpenDialog(param);
                if (picked != null)
                {
                    param.Value          = picked;
                    pathLabel.Text       = picked;
                    pathLabel.Foreground = new SolidColorBrush(HexColor("#CBD5E1"));
                    pathLabel.ToolTip    = picked;
                    pathBox.ToolTip      = picked;
                    clearBtn.Visibility  = Visibility.Visible;
                }
            };
            ApplyActionButtonTemplate(clearBtn);
            clearBtn.Click += (_, __) =>
            {
                param.Value          = "";
                pathLabel.Text       = "";
                pathLabel.Foreground = new SolidColorBrush(HexColor("#6B7FA8"));
                pathLabel.ToolTip    = "No file selected";
                clearBtn.Visibility  = Visibility.Collapsed;
            };

            Grid.SetColumn(pathBox,   0);
            Grid.SetColumn(browseBtn, 1);
            Grid.SetColumn(clearBtn,  2);
            grid.Children.Add(pathBox);
            grid.Children.Add(browseBtn);
            grid.Children.Add(clearBtn);

            _controls[param] = pathLabel;
            return grid;
        }

        private FrameworkElement BuildTextBox(ScriptParam param)
        {
            var tb = new TextBox
            {
                Background            = new SolidColorBrush(HexColor("#161B2A")),
                Foreground            = new SolidColorBrush(HexColor("#CBD5E1")),
                BorderBrush           = new SolidColorBrush(HexColor("#252D42")),
                BorderThickness       = new Thickness(1),
                FontSize              = 12,
                FontFamily            = new System.Windows.Media.FontFamily("Consolas, Segoe UI Mono, Segoe UI"),
                Height                = 32,
                Padding               = new Thickness(10, 0, 10, 0),
                VerticalContentAlignment = VerticalAlignment.Center,
                Text                  = param.Value,
            };
            tb.TextChanged += (_, __) => param.Value = tb.Text;
            _controls[param] = tb;
            return tb;
        }

        private FrameworkElement BuildCheckBox(ScriptParam param)
        {
            var cb = new CheckBox
            {
                IsChecked  = param.Value == "true" || param.Default == "true",
                Foreground = new SolidColorBrush(HexColor("#CBD5E1")),
                FontSize   = 12,
                FontFamily = new System.Windows.Media.FontFamily("Segoe UI Variable, Segoe UI"),
                Content    = new TextBlock
                {
                    Text       = "Enable",
                    Foreground = new SolidColorBrush(HexColor("#7A8AA0")),
                    FontSize   = 11.5
                },
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(2, 4, 0, 0)
            };
            cb.Checked   += (_, __) => param.Value = "true";
            cb.Unchecked += (_, __) => param.Value = "";
            param.Value    = cb.IsChecked == true ? "true" : "";
            _controls[param] = cb;
            return cb;
        }

        // ── File dialogs ──────────────────────────────────────────────────────

        private static string? ShowOpenDialog(ScriptParam param)
        {
            var dlg = new OpenFileDialog
            {
                Title  = $"Select {param.Label}",
                Filter = BuildFilter(param.Filter, param.Label)
            };
            return dlg.ShowDialog() == true ? dlg.FileName : null;
        }

        private static string? ShowSaveDialog(ScriptParam param)
        {
            var dlg = new SaveFileDialog
            {
                Title            = $"Save {param.Label} as…",
                Filter           = BuildFilter(param.Filter, param.Label),
                OverwritePrompt  = true,
                AddExtension     = true,
            };
            return dlg.ShowDialog() == true ? dlg.FileName : null;
        }

        private static string BuildFilter(string? hint, string label)
        {
            if (!string.IsNullOrEmpty(hint)) return hint + "|All Files (*.*)|*.*";

            // Infer from label
            var lower = label.ToLowerInvariant();
            if (lower.Contains("html") || lower.Contains("htm"))
                return "HTML Files (*.html;*.htm)|*.html;*.htm|All Files (*.*)|*.*";
            if (lower.Contains("csv"))
                return "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";
            if (lower.Contains("excel") || lower.Contains("xlsx"))
                return "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            if (lower.Contains("json"))
                return "JSON Files (*.json)|*.json|All Files (*.*)|*.*";
            if (lower.Contains("log"))
                return "Log Files (*.log;*.txt)|*.log;*.txt|All Files (*.*)|*.*";

            return "All Files (*.*)|*.*";
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private static string TypeHintText(ScriptParam p) => p.Type switch
        {
            "file-in"  => "input file",
            "file-out" => "output file",
            "flag"     => "on/off flag",
            "float"    => "number",
            "int"      => "integer",
            _          => "text"
        };

        private static string TypeHintColor(ScriptParam p) => p.Type switch
        {
            "file-in"  => "#60A5FA",
            "file-out" => "#34D399",
            "flag"     => "#A78BFA",
            "float"    => "#FBBF24",
            "int"      => "#FBBF24",
            _          => "#6B7A99"
        };

        private static Color HexColor(string hex)
            => (Color)System.Windows.Media.ColorConverter.ConvertFromString(hex);

        // Apply a minimal hover/press template similar to ActionButtonStyle
        private static void ApplyActionButtonTemplate(Button btn)
        {
            var tpl = new ControlTemplate(typeof(Button));
            var bd  = new FrameworkElementFactory(typeof(Border));
            bd.Name = "Bd";
            bd.SetBinding(Border.BackgroundProperty,
                new System.Windows.Data.Binding("Background")
                { RelativeSource = new System.Windows.Data.RelativeSource(
                    System.Windows.Data.RelativeSourceMode.TemplatedParent) });
            bd.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
            var cp = new FrameworkElementFactory(typeof(ContentPresenter));
            cp.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            cp.SetValue(ContentPresenter.VerticalAlignmentProperty,   VerticalAlignment.Center);
            bd.AppendChild(cp);
            tpl.VisualTree = bd;
            btn.Template   = tpl;
        }
    }
}
